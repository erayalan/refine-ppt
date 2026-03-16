#!/usr/bin/env python3
"""
Add modern PowerPoint comments anchored to specific text ranges.

Usage:
    python add_comments.py <pptx_path> <comments_json_path>

The comments JSON should have this structure:
{
    "author_name": "Claude",
    "author_initials": "CL",
    "comments": [
        {
            "slide_number": 1,
            "highlight_text": "exact substring to highlight",
            "comment_text": "Suggested replacement text"
        }
    ]
}

The script automatically:
- Finds which shape contains the highlight_text
- Extracts or injects shape/slide creation IDs
- Computes character positions and text hashes
- Builds the modern comment XML with proper text-range anchoring
"""

import json
import os
import random
import re
import shutil
import sys
import tempfile
import uuid
import zipfile
from datetime import datetime, timezone
from xml.dom import minidom


# ---------------------------------------------------------------------------
# Namespaces
# ---------------------------------------------------------------------------

NS = {
    "p188": "http://schemas.microsoft.com/office/powerpoint/2018/8/main",
    "pc": "http://schemas.microsoft.com/office/powerpoint/2013/main/command",
    "ac": "http://schemas.microsoft.com/office/drawing/2013/main/command",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "a16": "http://schemas.microsoft.com/office/drawing/2014/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "p14": "http://schemas.microsoft.com/office/powerpoint/2010/main",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}

CONTENT_TYPE_AUTHORS = "application/vnd.ms-powerpoint.authors+xml"
CONTENT_TYPE_COMMENTS = "application/vnd.ms-powerpoint.comments+xml"
REL_TYPE_AUTHORS = "http://schemas.microsoft.com/office/2018/10/relationships/authors"
REL_TYPE_COMMENTS = "http://schemas.microsoft.com/office/2018/10/relationships/comments"
COMMENT_REL_EXT_URI = "{6950BFC3-D8DA-4A85-94F7-54DA5524770B}"
SHAPE_CREATION_ID_URI = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"
SLIDE_CREATION_ID_URI = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}"


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def new_guid():
    return "{" + str(uuid.uuid4()).upper() + "}"


def random_cid():
    """Generate a random uint32 creation ID for slides."""
    return random.randint(1, 0xFFFFFFFF)


def compute_text_hash(text: str) -> int:
    """MS-ODRAWXML section 1.3.9.1 hash algorithm."""
    n_hash = 0
    for ch in text:
        n_hash = ((n_hash << 5) + n_hash + ord(ch)) & 0xFFFFFFFF
    return n_hash


def now_iso():
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def parse_xml(path):
    """Parse XML file, return DOM document."""
    with open(path, "r", encoding="utf-8") as f:
        return minidom.parseString(f.read())


def write_xml(doc, path):
    """Write DOM document to file."""
    xml_str = doc.toxml(encoding="UTF-8").decode("utf-8")
    xml_str = xml_str.replace(
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    )
    with open(path, "w", encoding="utf-8") as f:
        f.write(xml_str)


def find_next_rid(rels_path):
    """Find the next available rId in a .rels file."""
    if not os.path.exists(rels_path):
        return "rId1", None
    doc = parse_xml(rels_path)
    max_id = 0
    for rel in doc.getElementsByTagName("Relationship"):
        rid = rel.getAttribute("Id")
        m = re.match(r"rId(\d+)", rid)
        if m:
            max_id = max(max_id, int(m.group(1)))
    return f"rId{max_id + 1}", doc


def ensure_rels_file(rels_path):
    """Create a .rels file if it doesn't exist."""
    if not os.path.exists(rels_path):
        os.makedirs(os.path.dirname(rels_path), exist_ok=True)
        doc = minidom.getDOMImplementation().createDocument(None, None, None)
        root = doc.createElementNS(NS["rel"], "Relationships")
        root.setAttribute("xmlns", NS["rel"])
        doc.appendChild(root)
        write_xml(doc, rels_path)


def add_relationship(rels_path, rid, rel_type, target):
    """Add a relationship entry to a .rels file."""
    ensure_rels_file(rels_path)
    doc = parse_xml(rels_path)
    root = doc.documentElement
    rel = doc.createElement("Relationship")
    rel.setAttribute("Id", rid)
    rel.setAttribute("Type", rel_type)
    rel.setAttribute("Target", target)
    root.appendChild(rel)
    write_xml(doc, rels_path)
    return rid


# ---------------------------------------------------------------------------
# Extract text from a shape element
# ---------------------------------------------------------------------------

def get_shape_plain_text(sp_element):
    """Extract the concatenated plain text from a shape's txBody.

    Paragraphs are separated by newline characters, matching how PowerPoint
    computes text ranges internally.
    """
    # txBody can be in p: or a: namespace depending on context, use wildcard
    tx_bodies = sp_element.getElementsByTagNameNS("*", "txBody")
    if not tx_bodies:
        return ""
    tx_body = tx_bodies[0]
    paragraphs = tx_body.getElementsByTagNameNS("*", "p")
    para_texts = []
    for para in paragraphs:
        runs = para.getElementsByTagNameNS("*", "r")
        run_text = ""
        for run in runs:
            t_elements = run.getElementsByTagNameNS("*", "t")
            for t in t_elements:
                for child in t.childNodes:
                    if child.nodeType == child.TEXT_NODE:
                        run_text += child.data
        para_texts.append(run_text)
    # PowerPoint uses \r as paragraph separator and adds a trailing \r
    return "\r".join(para_texts) + "\r"


# ---------------------------------------------------------------------------
# Extract or inject creation IDs
# ---------------------------------------------------------------------------

def get_shape_creation_id(sp_element):
    """Get the a16:creationId from a shape, or None if not present."""
    # Look in cNvPr > extLst > ext[uri=SHAPE_CREATION_ID_URI] > a16:creationId
    cnv_prs = sp_element.getElementsByTagNameNS("*", "cNvPr")
    if not cnv_prs:
        return None

    cnv_pr = cnv_prs[0]
    for ext_lst in cnv_pr.childNodes:
        if ext_lst.nodeType == ext_lst.ELEMENT_NODE and ext_lst.localName == "extLst":
            for ext in ext_lst.childNodes:
                if (ext.nodeType == ext.ELEMENT_NODE
                        and ext.localName == "ext"
                        and ext.getAttribute("uri") == SHAPE_CREATION_ID_URI):
                    for child in ext.childNodes:
                        if child.nodeType == child.ELEMENT_NODE and child.localName == "creationId":
                            return child.getAttribute("id")
    return None


def inject_shape_creation_id(sp_element, doc):
    """Inject an a16:creationId into a shape's cNvPr and return the GUID."""
    guid = new_guid()

    # Find cNvPr
    cnv_pr = None
    for nvSpPr in sp_element.childNodes:
        if nvSpPr.nodeType == nvSpPr.ELEMENT_NODE and nvSpPr.localName == "nvSpPr":
            for child in nvSpPr.childNodes:
                if child.nodeType == child.ELEMENT_NODE and child.localName == "cNvPr":
                    cnv_pr = child
                    break
    if not cnv_pr:
        return guid  # Can't inject, return guid anyway for comment

    # Find or create extLst
    ext_lst = None
    for child in cnv_pr.childNodes:
        if child.nodeType == child.ELEMENT_NODE and child.localName == "extLst":
            ext_lst = child
            break
    if not ext_lst:
        ext_lst = doc.createElementNS(NS["a"], "a:extLst")
        cnv_pr.appendChild(ext_lst)

    # Create ext with creationId
    ext = doc.createElementNS(NS["a"], "a:ext")
    ext.setAttribute("uri", SHAPE_CREATION_ID_URI)
    creation_id = doc.createElementNS(NS["a16"], "a16:creationId")
    creation_id.setAttribute("xmlns:a16", NS["a16"])
    creation_id.setAttribute("id", guid)
    ext.appendChild(creation_id)
    ext_lst.appendChild(ext)

    return guid


def get_slide_creation_id(slide_xml_path):
    """Get the p14:creationId from a slide's extLst, or None."""
    doc = parse_xml(slide_xml_path)
    root = doc.documentElement

    for child in root.childNodes:
        if child.nodeType == child.ELEMENT_NODE and child.localName == "extLst":
            for ext in child.childNodes:
                if (ext.nodeType == ext.ELEMENT_NODE
                        and ext.localName == "ext"
                        and ext.getAttribute("uri") == SLIDE_CREATION_ID_URI):
                    for sub in ext.childNodes:
                        if sub.nodeType == sub.ELEMENT_NODE and sub.localName == "creationId":
                            val = sub.getAttribute("val")
                            if val:
                                return int(val)
    return None


def inject_slide_creation_id(slide_xml_path):
    """Inject a p14:creationId into a slide and return the cId value."""
    cid = random_cid()
    doc = parse_xml(slide_xml_path)
    root = doc.documentElement

    # Find or create extLst on the root <p:sld> element
    ext_lst = None
    for child in root.childNodes:
        if child.nodeType == child.ELEMENT_NODE and child.localName == "extLst":
            ext_lst = child
            break
    if not ext_lst:
        ext_lst = doc.createElementNS(NS["p"], "p:extLst")
        root.appendChild(ext_lst)

    # Create ext with creationId
    ext = doc.createElementNS(NS["p"], "p:ext")
    ext.setAttribute("uri", SLIDE_CREATION_ID_URI)
    creation_id = doc.createElementNS(NS["p14"], "p14:creationId")
    creation_id.setAttribute("xmlns:p14", NS["p14"])
    creation_id.setAttribute("val", str(cid))
    ext.appendChild(creation_id)
    ext_lst.appendChild(ext)

    write_xml(doc, slide_xml_path)
    return cid


# ---------------------------------------------------------------------------
# Scan a slide for shapes and match highlight text
# ---------------------------------------------------------------------------

def get_slide_id_from_presentation(unpack_dir, slide_number):
    """Get the slide ID from presentation.xml for a given slide number (1-indexed)."""
    pres_path = os.path.join(unpack_dir, "ppt", "presentation.xml")
    doc = parse_xml(pres_path)

    sld_id_lst = doc.getElementsByTagNameNS(NS["p"], "sldIdLst")
    if not sld_id_lst:
        return None

    sld_ids = sld_id_lst[0].getElementsByTagNameNS(NS["p"], "sldId")
    if slide_number - 1 < len(sld_ids):
        return int(sld_ids[slide_number - 1].getAttribute("id"))
    return None


def find_shape_for_text(slide_xml_path, highlight_text):
    """Find the shape containing the highlight_text and return shape info.

    Returns dict with: shape_id, shape_creation_id, shape_full_text,
    or None if not found.
    """
    doc = parse_xml(slide_xml_path)

    # Find all sp elements (shapes) using wildcard namespace
    shapes = doc.getElementsByTagNameNS("*", "sp")

    for sp in shapes:
        full_text = get_shape_plain_text(sp)
        if highlight_text in full_text:
            # Get shape ID from cNvPr
            shape_id = None
            for nvSpPr in sp.childNodes:
                if nvSpPr.nodeType == nvSpPr.ELEMENT_NODE and nvSpPr.localName == "nvSpPr":
                    for child in nvSpPr.childNodes:
                        if child.nodeType == child.ELEMENT_NODE and child.localName == "cNvPr":
                            shape_id = int(child.getAttribute("id"))
                            break

            # Get or inject creation ID
            creation_id = get_shape_creation_id(sp)
            if not creation_id:
                creation_id = inject_shape_creation_id(sp, doc)
                # Write back the modified slide XML
                write_xml(doc, slide_xml_path)

            return {
                "shape_id": shape_id,
                "shape_creation_id": creation_id,
                "shape_full_text": full_text,
            }

    return None


# ---------------------------------------------------------------------------
# Authors
# ---------------------------------------------------------------------------

def create_authors_xml(unpack_dir, author_name, author_initials):
    """Create ppt/authors.xml and return the author GUID."""
    author_id = new_guid()
    authors_path = os.path.join(unpack_dir, "ppt", "authors.xml")

    doc = minidom.getDOMImplementation().createDocument(NS["p188"], "p188:authorLst", None)
    root = doc.documentElement
    root.setAttribute("xmlns:p188", NS["p188"])

    author = doc.createElementNS(NS["p188"], "p188:author")
    author.setAttribute("id", author_id)
    author.setAttribute("name", author_name)
    author.setAttribute("initials", author_initials)
    author.setAttribute("userId", f"claude-review-{uuid.uuid4().hex[:8]}")
    author.setAttribute("providerId", "None")
    root.appendChild(author)

    write_xml(doc, authors_path)

    # Add relationship in presentation .rels
    pres_rels = os.path.join(unpack_dir, "ppt", "_rels", "presentation.xml.rels")
    rid, _ = find_next_rid(pres_rels)
    add_relationship(pres_rels, rid, REL_TYPE_AUTHORS, "authors.xml")

    # Add content type
    add_content_type_override(unpack_dir, "/ppt/authors.xml", CONTENT_TYPE_AUTHORS)

    return author_id


# ---------------------------------------------------------------------------
# Content Types
# ---------------------------------------------------------------------------

def add_content_type_override(unpack_dir, part_name, content_type):
    """Add an Override entry to [Content_Types].xml."""
    ct_path = os.path.join(unpack_dir, "[Content_Types].xml")
    doc = parse_xml(ct_path)
    root = doc.documentElement

    for override in doc.getElementsByTagName("Override"):
        if override.getAttribute("PartName") == part_name:
            return

    override = doc.createElement("Override")
    override.setAttribute("PartName", part_name)
    override.setAttribute("ContentType", content_type)
    root.appendChild(override)
    write_xml(doc, ct_path)


# ---------------------------------------------------------------------------
# Comment XML building
# ---------------------------------------------------------------------------

def build_comment_xml(author_id, resolved_comments):
    """Build the XML DOM for a single slide's comments file.

    resolved_comments is a list of dicts with all IDs already resolved:
    slide_id, slide_creation_id, shape_id, shape_creation_id,
    shape_full_text, highlight_text, comment_text
    """
    doc = minidom.getDOMImplementation().createDocument(NS["p188"], "p188:cmLst", None)
    root = doc.documentElement
    root.setAttribute("xmlns:p188", NS["p188"])
    root.setAttribute("xmlns:pc", NS["pc"])
    root.setAttribute("xmlns:ac", NS["ac"])
    root.setAttribute("xmlns:a", NS["a"])
    root.setAttribute("xmlns:r", NS["r"])

    timestamp = now_iso()

    for c in resolved_comments:
        comment_id = new_guid()
        full_text = c["shape_full_text"]
        highlight = c["highlight_text"]

        cp = full_text.find(highlight)
        if cp == -1:
            print(f"  WARNING: Could not find highlight text in shape. Skipping.")
            print(f"    Shape text: {full_text[:80]}...")
            print(f"    Highlight:  {highlight[:80]}...")
            continue
        length = len(highlight)

        text_hash = compute_text_hash(full_text)
        text_len = len(full_text)

        cm = doc.createElementNS(NS["p188"], "p188:cm")
        cm.setAttribute("id", comment_id)
        cm.setAttribute("authorId", author_id)
        cm.setAttribute("status", "active")
        cm.setAttribute("created", timestamp)

        # --- Text range moniker (the only anchor for text-anchored comments) ---
        tx_mk_lst = doc.createElementNS(NS["ac"], "ac:txMkLst")
        doc_mk = doc.createElementNS(NS["pc"], "pc:docMk")
        tx_mk_lst.appendChild(doc_mk)
        sld_mk = doc.createElementNS(NS["pc"], "pc:sldMk")
        sld_mk.setAttribute("cId", str(c["slide_creation_id"]))
        sld_mk.setAttribute("sldId", str(c["slide_id"]))
        tx_mk_lst.appendChild(sld_mk)
        sp_mk = doc.createElementNS(NS["ac"], "ac:spMk")
        sp_mk.setAttribute("id", str(c["shape_id"]))
        sp_mk.setAttribute("creationId", c["shape_creation_id"])
        tx_mk_lst.appendChild(sp_mk)
        tx_mk = doc.createElementNS(NS["ac"], "ac:txMk")
        tx_mk.setAttribute("cp", str(cp))
        tx_mk.setAttribute("len", str(length))
        ctx = doc.createElementNS(NS["ac"], "ac:context")
        ctx.setAttribute("hash", str(text_hash))
        ctx.setAttribute("len", str(text_len))
        tx_mk.appendChild(ctx)
        tx_mk_lst.appendChild(tx_mk)
        cm.appendChild(tx_mk_lst)

        # --- Position ---
        pos = doc.createElementNS(NS["p188"], "p188:pos")
        pos.setAttribute("x", "0")
        pos.setAttribute("y", "0")
        cm.appendChild(pos)

        # --- Comment text body ---
        tx_body = doc.createElementNS(NS["p188"], "p188:txBody")
        body_pr = doc.createElementNS(NS["a"], "a:bodyPr")
        tx_body.appendChild(body_pr)
        lst_style = doc.createElementNS(NS["a"], "a:lstStyle")
        tx_body.appendChild(lst_style)
        para = doc.createElementNS(NS["a"], "a:p")
        run = doc.createElementNS(NS["a"], "a:r")
        r_pr = doc.createElementNS(NS["a"], "a:rPr")
        r_pr.setAttribute("lang", "en-US")
        r_pr.setAttribute("dirty", "0")
        run.appendChild(r_pr)
        t = doc.createElementNS(NS["a"], "a:t")
        t.appendChild(doc.createTextNode(c["comment_text"]))
        run.appendChild(t)
        para.appendChild(run)
        tx_body.appendChild(para)
        cm.appendChild(tx_body)

        root.appendChild(cm)

    return doc


# ---------------------------------------------------------------------------
# Slide extension for comment relationship
# ---------------------------------------------------------------------------

def add_comment_ext_to_slide(slide_xml_path, comment_rid):
    """Add the commentRel extension to a slide's extLst."""
    doc = parse_xml(slide_xml_path)
    root = doc.documentElement

    ext_lsts = [
        n for n in root.childNodes
        if n.nodeType == n.ELEMENT_NODE and n.localName == "extLst"
    ]

    if ext_lsts:
        ext_lst = ext_lsts[0]
    else:
        ext_lst = doc.createElementNS(NS["p"], "p:extLst")
        root.appendChild(ext_lst)

    for ext in ext_lst.childNodes:
        if ext.nodeType == ext.ELEMENT_NODE and ext.getAttribute("uri") == COMMENT_REL_EXT_URI:
            return

    ext = doc.createElementNS(NS["p"], "p:ext")
    ext.setAttribute("uri", COMMENT_REL_EXT_URI)

    comment_rel = doc.createElementNS(NS["p188"], "p188:commentRel")
    comment_rel.setAttribute("xmlns:p188", NS["p188"])
    comment_rel.setAttribute("r:id", comment_rid)
    comment_rel.setAttribute("xmlns:r", NS["r"])
    ext.appendChild(comment_rel)

    ext_lst.appendChild(ext)
    write_xml(doc, slide_xml_path)


# ---------------------------------------------------------------------------
# Unpack / Repack
# ---------------------------------------------------------------------------

def unpack_pptx(pptx_path, unpack_dir):
    """Extract a .pptx ZIP archive."""
    with zipfile.ZipFile(pptx_path, "r") as zf:
        zf.extractall(unpack_dir)


def repack_pptx(unpack_dir, pptx_path):
    """Re-zip the unpacked directory into a .pptx file."""
    if os.path.exists(pptx_path):
        os.remove(pptx_path)

    with zipfile.ZipFile(pptx_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for dirpath, dirnames, filenames in os.walk(unpack_dir):
            for filename in filenames:
                file_path = os.path.join(dirpath, filename)
                arcname = os.path.relpath(file_path, unpack_dir)
                zf.write(file_path, arcname)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    if len(sys.argv) != 3:
        print(f"Usage: {sys.argv[0]} <pptx_path> <comments_json_path>")
        sys.exit(1)

    pptx_path = os.path.abspath(sys.argv[1])
    comments_json_path = os.path.abspath(sys.argv[2])

    if not os.path.exists(pptx_path):
        print(f"Error: {pptx_path} not found")
        sys.exit(1)
    if not os.path.exists(comments_json_path):
        print(f"Error: {comments_json_path} not found")
        sys.exit(1)

    with open(comments_json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    author_name = data.get("author_name", "Claude")
    author_initials = data.get("author_initials", "CL")
    comments = data.get("comments", [])

    if not comments:
        print("No comments to add.")
        return

    # Group comments by slide number
    by_slide = {}
    for c in comments:
        sn = c["slide_number"]
        by_slide.setdefault(sn, []).append(c)

    # Unpack
    unpack_dir = tempfile.mkdtemp(prefix="refine_ppt_")
    print(f"Unpacking {pptx_path}...")
    unpack_pptx(pptx_path, unpack_dir)

    # Create author
    author_id = create_authors_xml(unpack_dir, author_name, author_initials)
    print(f"Created author: {author_name} ({author_id})")

    # Ensure comments directory exists
    comments_dir = os.path.join(unpack_dir, "ppt", "comments")
    os.makedirs(comments_dir, exist_ok=True)

    # Cache for slide creation IDs (so we only inject once per slide)
    slide_cid_cache = {}

    total_added = 0

    # Process each slide
    for slide_num, slide_comments in sorted(by_slide.items()):
        print(f"\nProcessing slide {slide_num} ({len(slide_comments)} comments)...")

        slide_xml_path = os.path.join(unpack_dir, "ppt", "slides", f"slide{slide_num}.xml")
        if not os.path.exists(slide_xml_path):
            print(f"  ERROR: slide{slide_num}.xml not found, skipping")
            continue

        # Get slide ID from presentation.xml
        slide_id = get_slide_id_from_presentation(unpack_dir, slide_num)
        if slide_id is None:
            print(f"  ERROR: Could not find slide ID for slide {slide_num}, skipping")
            continue

        # Get or inject slide creation ID
        if slide_num not in slide_cid_cache:
            cid = get_slide_creation_id(slide_xml_path)
            if cid is None:
                cid = inject_slide_creation_id(slide_xml_path)
                print(f"  Injected slide creation ID: {cid}")
            else:
                print(f"  Found existing slide creation ID: {cid}")
            slide_cid_cache[slide_num] = cid
        slide_cid = slide_cid_cache[slide_num]

        # Resolve each comment: find shapes, get/inject creation IDs
        resolved = []
        for c in slide_comments:
            highlight = c["highlight_text"]
            shape_info = find_shape_for_text(slide_xml_path, highlight)
            if shape_info is None:
                print(f"  WARNING: Could not find text '{highlight[:50]}...' in any shape on slide {slide_num}")
                continue

            resolved.append({
                "slide_id": slide_id,
                "slide_creation_id": slide_cid,
                "shape_id": shape_info["shape_id"],
                "shape_creation_id": shape_info["shape_creation_id"],
                "shape_full_text": shape_info["shape_full_text"],
                "highlight_text": highlight,
                "comment_text": c["comment_text"],
            })
            print(f"  Matched: '{highlight[:40]}...' -> shape {shape_info['shape_id']}")

        if not resolved:
            print(f"  No comments resolved for slide {slide_num}")
            continue

        # Build comment XML
        comment_doc = build_comment_xml(author_id, resolved)
        comment_filename = f"comment{slide_num}.xml"
        comment_path = os.path.join(comments_dir, comment_filename)
        write_xml(comment_doc, comment_path)

        # Add content type
        add_content_type_override(
            unpack_dir,
            f"/ppt/comments/{comment_filename}",
            CONTENT_TYPE_COMMENTS,
        )

        # Add relationship to slide
        slide_rels_dir = os.path.join(unpack_dir, "ppt", "slides", "_rels")
        os.makedirs(slide_rels_dir, exist_ok=True)
        slide_rels_path = os.path.join(slide_rels_dir, f"slide{slide_num}.xml.rels")
        rid, _ = find_next_rid(slide_rels_path)
        add_relationship(
            slide_rels_path,
            rid,
            REL_TYPE_COMMENTS,
            f"../comments/{comment_filename}",
        )

        # Add extension to slide XML
        add_comment_ext_to_slide(slide_xml_path, rid)

        total_added += len(resolved)

    # Repack
    print(f"\nRepacking to {pptx_path}...")
    repack_pptx(unpack_dir, pptx_path)

    # Cleanup
    shutil.rmtree(unpack_dir, ignore_errors=True)

    print(f"Done! Added {total_added} comments to {pptx_path}")


if __name__ == "__main__":
    main()
