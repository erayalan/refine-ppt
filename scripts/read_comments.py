#!/usr/bin/env python3
"""
Read all modern comments and their reply threads from a .pptx file.

Usage:
    python read_comments.py <pptx_path>

Outputs JSON to stdout with all comment threads grouped by slide.
"""

import json
import os
import sys
import tempfile
import shutil
import zipfile
from xml.dom import minidom


def extract_text_from_txBody(txBody):
    """Extract plain text from a p188:txBody element."""
    paragraphs = []
    for node in txBody.childNodes:
        if node.nodeType == node.ELEMENT_NODE and node.localName == "p":
            runs_text = ""
            for child in node.childNodes:
                if child.nodeType == child.ELEMENT_NODE and child.localName == "r":
                    for t in child.childNodes:
                        if t.nodeType == t.ELEMENT_NODE and t.localName == "t":
                            for text_node in t.childNodes:
                                if text_node.nodeType == text_node.TEXT_NODE:
                                    runs_text += text_node.data
            paragraphs.append(runs_text)
    return "\n".join(paragraphs).strip()


def get_authors(unpack_dir):
    """Read authors.xml and return a dict of authorId -> author info."""
    authors = {}
    authors_path = os.path.join(unpack_dir, "ppt", "authors.xml")
    if not os.path.exists(authors_path):
        return authors

    with open(authors_path, "r", encoding="utf-8") as f:
        doc = minidom.parseString(f.read())

    for author in doc.getElementsByTagNameNS("*", "author"):
        aid = author.getAttribute("id")
        authors[aid] = {
            "id": aid,
            "name": author.getAttribute("name"),
            "initials": author.getAttribute("initials"),
        }
    return authors


def get_slide_number_for_comment_file(unpack_dir, comment_filename):
    """Figure out which slide number a comment file belongs to by checking rels."""
    slides_dir = os.path.join(unpack_dir, "ppt", "slides")
    rels_dir = os.path.join(slides_dir, "_rels")
    if not os.path.exists(rels_dir):
        return None

    for rels_file in os.listdir(rels_dir):
        if not rels_file.endswith(".xml.rels"):
            continue
        rels_path = os.path.join(rels_dir, rels_file)
        with open(rels_path, "r", encoding="utf-8") as f:
            doc = minidom.parseString(f.read())
        for rel in doc.getElementsByTagName("Relationship"):
            target = rel.getAttribute("Target")
            if comment_filename in target:
                # Extract slide number from rels filename: slide3.xml.rels -> 3
                import re
                m = re.match(r"slide(\d+)\.xml\.rels", rels_file)
                if m:
                    return int(m.group(1))
    return None


def parse_comment_file(comment_path, authors, slide_number):
    """Parse a single modern comment XML file and return comment threads."""
    with open(comment_path, "r", encoding="utf-8") as f:
        doc = minidom.parseString(f.read())

    threads = []

    for cm in doc.getElementsByTagNameNS("*", "cm"):
        comment_id = cm.getAttribute("id")
        author_id = cm.getAttribute("authorId")
        created = cm.getAttribute("created")
        status = cm.getAttribute("status") or "active"

        # Get comment text
        comment_text = ""
        for child in cm.childNodes:
            if child.nodeType == child.ELEMENT_NODE and child.localName == "txBody":
                comment_text = extract_text_from_txBody(child)

        # Get text anchor info
        anchor_info = {}
        for txMkLst in cm.getElementsByTagNameNS("*", "txMkLst"):
            for txMk in txMkLst.getElementsByTagNameNS("*", "txMk"):
                anchor_info["cp"] = int(txMk.getAttribute("cp") or 0)
                anchor_info["len"] = int(txMk.getAttribute("len") or 0)

        # Get replies
        replies = []
        for replyLst in cm.getElementsByTagNameNS("*", "replyLst"):
            for reply in replyLst.getElementsByTagNameNS("*", "reply"):
                reply_id = reply.getAttribute("id")
                reply_author_id = reply.getAttribute("authorId")
                reply_created = reply.getAttribute("created")

                reply_text = ""
                for child in reply.childNodes:
                    if child.nodeType == child.ELEMENT_NODE and child.localName == "txBody":
                        reply_text = extract_text_from_txBody(child)

                author_info = authors.get(reply_author_id, {})
                replies.append({
                    "reply_id": reply_id,
                    "author_name": author_info.get("name", "Unknown"),
                    "author_id": reply_author_id,
                    "created": reply_created,
                    "text": reply_text,
                })

        author_info = authors.get(author_id, {})
        threads.append({
            "comment_id": comment_id,
            "slide_number": slide_number,
            "author_name": author_info.get("name", "Unknown"),
            "author_id": author_id,
            "created": created,
            "status": status,
            "text": comment_text,
            "anchor": anchor_info,
            "replies": replies,
        })

    return threads


def main():
    if len(sys.argv) != 2:
        print(f"Usage: {sys.argv[0]} <pptx_path>", file=sys.stderr)
        sys.exit(1)

    pptx_path = os.path.abspath(sys.argv[1])
    if not os.path.exists(pptx_path):
        print(f"Error: {pptx_path} not found", file=sys.stderr)
        sys.exit(1)

    # Unpack
    unpack_dir = tempfile.mkdtemp(prefix="read_comments_")
    with zipfile.ZipFile(pptx_path, "r") as zf:
        zf.extractall(unpack_dir)

    # Read authors
    authors = get_authors(unpack_dir)

    # Find and parse all comment files
    comments_dir = os.path.join(unpack_dir, "ppt", "comments")
    all_threads = []

    if os.path.exists(comments_dir):
        for filename in sorted(os.listdir(comments_dir)):
            if not filename.endswith(".xml"):
                continue
            comment_path = os.path.join(comments_dir, filename)
            slide_num = get_slide_number_for_comment_file(unpack_dir, filename)
            threads = parse_comment_file(comment_path, authors, slide_num)
            all_threads.extend(threads)

    # Cleanup
    shutil.rmtree(unpack_dir, ignore_errors=True)

    # Output
    result = {
        "file": pptx_path,
        "authors": list(authors.values()),
        "comments": all_threads,
    }
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()
