"""
pptx_to_jpeg.py
---------------
Converts a PowerPoint file to JPEGs, expanding slides with entrance animations
into multiple images — one per click state.

Usage:
    python pptx_to_jpeg.py input.pptx [--output-dir ./output] [--dpi 150]

For a slide with 3 entrance animation steps, it produces 4 images:
    slide_05_state_0.jpg  <- initial state (before any click)
    slide_05_state_1.jpg  <- after click 1
    slide_05_state_2.jpg  <- after click 2
    slide_05_state_3.jpg  <- after click 3 (fully revealed)

Slides with no animations produce a single image:
    slide_01.jpg
"""

import argparse
import copy
import io
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_R_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

def ptag(local): return f"{{{P_NS}}}{local}"
def atag(local): return f"{{{A_NS}}}{local}"


def parse_animation_steps(slide_xml_bytes):
    root = etree.fromstring(slide_xml_bytes)
    timing = root.find(f".//{ptag('timing')}")
    if timing is None:
        return []

    steps = []
    for seq in timing.findall(f".//{ptag('seq')}"):
        childTnLst = seq.find(f".//{ptag('childTnLst')}")
        if childTnLst is None:
            continue
        for par in childTnLst.findall(f"{ptag('par')}"):
            cTn = par.find(f".//{ptag('cTn')}")
            if cTn is None:
                continue
            stCondLst = cTn.find(f".//{ptag('stCondLst')}")
            is_click = False
            if stCondLst is not None:
                for cond in stCondLst:
                    if cond.get("evt") == "onClick" or cond.get("delay") == "indefinite":
                        is_click = True
                        break
            if not is_click:
                continue

            step = {}
            for set_el in par.iter(ptag("set")):
                attrNameLst = set_el.find(f".//{ptag('attrNameLst')}")
                if attrNameLst is None:
                    continue
                names = [a.text for a in attrNameLst.findall(ptag("attrName"))]
                if "style.visibility" not in names:
                    continue
                to_el = set_el.find(ptag("to"))
                if to_el is None:
                    continue
                strVal = to_el.find(f".//{ptag('strVal')}")
                if strVal is None or strVal.get("val") != "visible":
                    continue
                tgtEl = set_el.find(f".//{ptag('tgtEl')}")
                if tgtEl is None:
                    continue
                spTgt = tgtEl.find(ptag("spTgt"))
                if spTgt is None:
                    continue
                spid = spTgt.get("spid")
                txEl = spTgt.find(ptag("txEl"))
                if txEl is not None:
                    pRg = txEl.find(ptag("pRg"))
                    if pRg is not None:
                        st, end = int(pRg.get("st")), int(pRg.get("end"))
                        if spid not in step:
                            step[spid] = set()
                        if step[spid] != "shape":
                            step[spid].add((st, end))
                else:
                    step[spid] = "shape"
            if step:
                steps.append(step)
    return steps


def all_animated_elements(steps):
    all_anim = {}
    for step in steps:
        for spid, info in step.items():
            if info == "shape":
                all_anim[spid] = "shape"
            else:
                if spid not in all_anim:
                    all_anim[spid] = set()
                if all_anim[spid] != "shape":
                    all_anim[spid] |= info
    return all_anim


def visible_at_state(steps, state_idx):
    visible = {}
    for i in range(state_idx):
        for spid, info in steps[i].items():
            if info == "shape":
                visible[spid] = "shape"
            else:
                if spid not in visible:
                    visible[spid] = set()
                if visible[spid] != "shape":
                    visible[spid] |= info
    return visible


def make_para_transparent(para_el):
    for r_el in para_el.findall(f"{{{A_NS}}}r"):
        rPr = r_el.find(f"{{{A_NS}}}rPr")
        if rPr is None:
            rPr = etree.Element(f"{{{A_NS}}}rPr")
            r_el.insert(0, rPr)
        for tag in ["solidFill", "gradFill", "noFill", "pattFill", "blipFill", "grpFill"]:
            for old in rPr.findall(f"{{{A_NS}}}{tag}"):
                rPr.remove(old)
        solidFill = etree.SubElement(rPr, f"{{{A_NS}}}solidFill")
        srgbClr = etree.SubElement(solidFill, f"{{{A_NS}}}srgbClr")
        srgbClr.set("val", "FFFFFF")
        alpha = etree.SubElement(srgbClr, f"{{{A_NS}}}alpha")
        alpha.set("val", "0")


def apply_state_to_slide_xml(slide_xml_bytes, all_anim, visible):
    root = etree.fromstring(slide_xml_bytes)

    for tag in [ptag("timing"), ptag("transition")]:
        el = root.find(f".//{tag}")
        if el is not None:
            parent = el.getparent()
            if parent is not None:
                parent.remove(el)

    SHAPE_TAGS = {ptag('sp'), ptag('pic'), ptag('graphicFrame'), ptag('grpSp'), ptag('cxnSp')}
    spTree = root.find(f".//{ptag('spTree')}")
    if spTree is not None:
        for sp in spTree:
            if sp.tag not in SHAPE_TAGS:
                continue
            # cNvPr lives at depth 2 for all shape types (nvSpPr/nvPicPr/etc.)
            cNvPr = sp.find(f".//{ptag('cNvPr')}")
            if cNvPr is None:
                continue
            spid = cNvPr.get("id")

            if spid not in all_anim:
                cNvPr.attrib.pop("hidden", None)
                continue

            anim_info = all_anim[spid]
            vis_info = visible.get(spid)

            if anim_info == "shape":
                if vis_info == "shape":
                    cNvPr.attrib.pop("hidden", None)
                else:
                    cNvPr.set("hidden", "1")
            else:
                cNvPr.attrib.pop("hidden", None)
                txBody = sp.find(f"{ptag('txBody')}")
                if txBody is None:
                    continue
                paras = txBody.findall(f"{{{A_NS}}}p")

                visible_para_idx = set()
                if vis_info and vis_info != "shape":
                    for (st, end) in vis_info:
                        for idx in range(st, end + 1):
                            visible_para_idx.add(idx)

                animated_para_idx = set()
                for (st, end) in anim_info:
                    for idx in range(st, end + 1):
                        animated_para_idx.add(idx)

                for i, para_el in enumerate(paras):
                    if i in animated_para_idx and i not in visible_para_idx:
                        make_para_transparent(para_el)

    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def build_expanded_pptx(input_path, output_path):
    with zipfile.ZipFile(input_path, 'r') as zin:
        all_data = {name: zin.read(name) for name in zin.namelist()}

    prs_xml = etree.fromstring(all_data["ppt/presentation.xml"])
    sldIdLst = prs_xml.find(f".//{ptag('sldIdLst')}")
    slide_entries = sldIdLst.findall(ptag("sldId"))

    prs_rels_xml = etree.fromstring(all_data["ppt/_rels/presentation.xml.rels"])
    rid_to_target = {}
    for rel in prs_rels_xml.findall(f"{{{PKG_R_NS}}}Relationship"):
        rid_to_target[rel.get("Id")] = rel.get("Target")

    slide_paths = []
    for entry in slide_entries:
        rid = entry.get(f"{{{R_NS}}}id")
        target = rid_to_target.get(rid, "")
        path = f"ppt/{target}" if not target.startswith("ppt/") else target
        slide_paths.append(path)

    manifest = []
    new_slides = []  # (new_path, new_rels_path, xml_bytes, rels_bytes)
    slide_counter = 0

    for orig_idx, orig_path in enumerate(slide_paths):
        base = orig_path.split("/")[-1]
        orig_rels_path = f"ppt/slides/_rels/{base}.rels"
        slide_xml = all_data.get(orig_path, b"")
        slide_rels = all_data.get(orig_rels_path, b"")
        steps = parse_animation_steps(slide_xml)

        if not steps:
            slide_counter += 1
            new_path = f"ppt/slides/slide{slide_counter}.xml"
            new_rels_path = f"ppt/slides/_rels/slide{slide_counter}.xml.rels"
            clean_xml = apply_state_to_slide_xml(slide_xml, {}, {})
            new_slides.append((new_path, new_rels_path, clean_xml, slide_rels))
            manifest.append((orig_idx, 0, f"slide_{orig_idx+1:02d}"))
        else:
            all_anim = all_animated_elements(steps)
            n_states = len(steps) + 1
            for state_idx in range(n_states):
                slide_counter += 1
                new_path = f"ppt/slides/slide{slide_counter}.xml"
                new_rels_path = f"ppt/slides/_rels/slide{slide_counter}.xml.rels"
                vis = visible_at_state(steps, state_idx)
                mod_xml = apply_state_to_slide_xml(slide_xml, all_anim, vis)
                new_slides.append((new_path, new_rels_path, mod_xml, slide_rels))
                manifest.append((orig_idx, state_idx, f"slide_{orig_idx+1:02d}_state_{state_idx}"))

    # Rebuild presentation.xml.rels — remove old slide rels, add new
    SLIDE_RTYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
    for rel in list(prs_rels_xml.findall(f"{{{PKG_R_NS}}}Relationship")):
        if rel.get("Type") == SLIDE_RTYPE:
            prs_rels_xml.remove(rel)

    max_rid = 0
    for rel in prs_rels_xml.findall(f"{{{PKG_R_NS}}}Relationship"):
        m = re.match(r"rId(\d+)", rel.get("Id", ""))
        if m:
            max_rid = max(max_rid, int(m.group(1)))

    # Rebuild sldIdLst
    for old in list(sldIdLst.findall(ptag("sldId"))):
        sldIdLst.remove(old)

    max_id = 256
    for entry in slide_entries:
        try:
            max_id = max(max_id, int(entry.get("id", 256)))
        except:
            pass

    for i, (new_path, _, _, _) in enumerate(new_slides):
        max_rid += 1
        rid = f"rId{max_rid}"
        target = new_path.replace("ppt/", "")
        rel_el = etree.SubElement(prs_rels_xml, f"{{{PKG_R_NS}}}Relationship")
        rel_el.set("Id", rid)
        rel_el.set("Type", SLIDE_RTYPE)
        rel_el.set("Target", target)
        max_id += 1
        sld_id_el = etree.SubElement(sldIdLst, ptag("sldId"))
        sld_id_el.set("id", str(max_id))
        sld_id_el.set(f"{{{R_NS}}}id", rid)

    # Write new ZIP
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        skip_exact = {"ppt/presentation.xml", "ppt/_rels/presentation.xml.rels", "[Content_Types].xml"}

        for name, data in all_data.items():
            if name in skip_exact:
                continue
            if name.startswith("ppt/slides/"):
                continue
            zout.writestr(name, data)

        zout.writestr("ppt/presentation.xml",
                      etree.tostring(prs_xml, xml_declaration=True, encoding="UTF-8", standalone=True))
        zout.writestr("ppt/_rels/presentation.xml.rels",
                      etree.tostring(prs_rels_xml))

        for new_path, new_rels_path, xml_bytes, rels_bytes in new_slides:
            zout.writestr(new_path, xml_bytes)
            if rels_bytes:
                zout.writestr(new_rels_path, rels_bytes)

        # Update [Content_Types].xml
        ct_xml = etree.fromstring(all_data["[Content_Types].xml"])
        CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
        SLIDE_CT = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
        for override in list(ct_xml.findall(f"{{{CT_NS}}}Override")):
            pt = override.get("PartName", "")
            if "/slides/slide" in pt and "_rels" not in pt:
                ct_xml.remove(override)
        for new_path, _, _, _ in new_slides:
            override = etree.SubElement(ct_xml, f"{{{CT_NS}}}Override")
            override.set("PartName", f"/{new_path}")
            override.set("ContentType", SLIDE_CT)
        zout.writestr("[Content_Types].xml",
                      etree.tostring(ct_xml, xml_declaration=True, encoding="UTF-8", standalone=True))

    return manifest


def pptx_to_jpegs(pptx_path, tmp_dir, dpi=150):
    tmp_dir = Path(tmp_dir)
    soffice_script = Path("/mnt/skills/public/pptx/scripts/office/soffice.py")
    if soffice_script.exists():
        cmd = [sys.executable, str(soffice_script), "--headless",
               "--convert-to", "pdf", str(pptx_path), "--outdir", str(tmp_dir)]
    else:
        cmd = ["soffice", "--headless", "--convert-to", "pdf",
               "--outdir", str(tmp_dir), str(pptx_path)]

    result = subprocess.run(cmd, capture_output=True, text=True)
    pdf_files = list(tmp_dir.glob("*.pdf"))
    if not pdf_files:
        raise RuntimeError(f"LibreOffice failed.\nSTDOUT: {result.stdout}\nSTDERR: {result.stderr}")

    jpeg_prefix = str(tmp_dir / "page")
    result2 = subprocess.run(
        ["pdftoppm", "-jpeg", "-r", str(dpi), str(pdf_files[0]), jpeg_prefix],
        capture_output=True, text=True
    )
    if result2.returncode != 0:
        raise RuntimeError(f"pdftoppm failed:\n{result2.stderr}")

    return sorted(tmp_dir.glob("page*.jpg"))


def main():
    parser = argparse.ArgumentParser(
        description="Convert PPTX to JPEGs, expanding animated slides."
    )
    parser.add_argument("input", help="Input .pptx file")
    parser.add_argument("--output-dir", default="./pptx_output")
    parser.add_argument("--dpi", type=int, default=150)
    args = parser.parse_args()

    input_path = Path(args.input)
    if not input_path.exists():
        print(f"Error: File not found: {input_path}")
        sys.exit(1)

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    print(f"Reading: {input_path.name}")
    print(f"Output:  {output_dir}/\n")

    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        expanded_pptx = tmp / "expanded.pptx"

        print("Parsing animations and building states...")
        manifest = build_expanded_pptx(str(input_path), str(expanded_pptx))

        from collections import Counter
        orig_counts = Counter(orig for orig, _, _ in manifest)
        print("\nSlide breakdown:")
        for orig_idx, state_idx, label in manifest:
            if state_idx == 0:
                count = orig_counts[orig_idx]
                if count > 1:
                    print(f"  Slide {orig_idx+1:2d}  -> {count} states (animated)")
                else:
                    print(f"  Slide {orig_idx+1:2d}  -> 1 image")

        print(f"\nTotal images: {len(manifest)}")
        print(f"PPTX size: {expanded_pptx.stat().st_size / 1024:.0f} KB\n")

        print("Converting to JPEG (may take 30-60s)...")
        raw_jpegs = pptx_to_jpegs(expanded_pptx, tmp, args.dpi)

        if len(raw_jpegs) != len(manifest):
            print(f"[Warning] Expected {len(manifest)}, got {len(raw_jpegs)} images.")

        renamed = []
        for i, (orig_idx, state_idx, label) in enumerate(manifest):
            if i < len(raw_jpegs):
                dst = output_dir / f"{label}.jpg"
                shutil.copy2(raw_jpegs[i], dst)
                renamed.append(dst)
                print(f"  {dst.name}")

    print(f"\nDone! {len(renamed)} JPEGs saved to: {output_dir}/")


if __name__ == "__main__":
    main()
