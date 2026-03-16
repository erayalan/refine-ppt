#!/usr/bin/env python3
"""
Add replies to existing modern comments in a .pptx file.

Usage:
    python add_replies.py <pptx_path> <replies_json_path>

The replies JSON should have this structure:
{
    "author_name": "Claude",
    "author_initials": "CL",
    "replies": [
        {
            "comment_id": "{GUID-of-comment-to-reply-to}",
            "reply_text": "The revised suggestion text"
        }
    ]
}

The script finds the existing comment by its ID across all comment files,
then appends a reply to its replyLst (creating one if needed).
"""

import json
import os
import re
import shutil
import sys
import tempfile
import uuid
import zipfile
from datetime import datetime, timezone
from xml.dom import minidom


NS = {
    "p188": "http://schemas.microsoft.com/office/powerpoint/2018/8/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def new_guid():
    return "{" + str(uuid.uuid4()).upper() + "}"


def now_iso():
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def parse_xml(path):
    with open(path, "r", encoding="utf-8") as f:
        return minidom.parseString(f.read())


def write_xml(doc, path):
    xml_str = doc.toxml(encoding="UTF-8").decode("utf-8")
    xml_str = xml_str.replace(
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    )
    with open(path, "w", encoding="utf-8") as f:
        f.write(xml_str)


def find_or_create_claude_author(unpack_dir, author_name, author_initials):
    """Find existing Claude author or create one. Return author ID."""
    authors_path = os.path.join(unpack_dir, "ppt", "authors.xml")
    if not os.path.exists(authors_path):
        return None  # No authors file means no comments exist

    doc = parse_xml(authors_path)
    root = doc.documentElement

    # Check if Claude author already exists
    for author in root.getElementsByTagNameNS("*", "author"):
        if author.getAttribute("name") == author_name:
            return author.getAttribute("id")

    # Create new author
    author_id = new_guid()
    author = doc.createElementNS(NS["p188"], "p188:author")
    author.setAttribute("id", author_id)
    author.setAttribute("name", author_name)
    author.setAttribute("initials", author_initials)
    author.setAttribute("userId", f"claude-review-{uuid.uuid4().hex[:8]}")
    author.setAttribute("providerId", "None")
    root.appendChild(author)
    write_xml(doc, authors_path)

    return author_id


def add_reply_to_comment(comment_file_path, comment_id, author_id, reply_text):
    """Add a reply to an existing comment in a comment XML file.

    Returns True if the comment was found and reply added, False otherwise.
    """
    doc = parse_xml(comment_file_path)

    # Find the comment with the matching ID
    for cm in doc.getElementsByTagNameNS("*", "cm"):
        if cm.getAttribute("id") == comment_id:
            # Find or create replyLst
            reply_lst = None
            for child in cm.childNodes:
                if child.nodeType == child.ELEMENT_NODE and child.localName == "replyLst":
                    reply_lst = child
                    break

            if reply_lst is None:
                # Create replyLst — insert before txBody
                reply_lst = doc.createElementNS(NS["p188"], "p188:replyLst")
                # Find txBody to insert before it
                tx_body = None
                for child in cm.childNodes:
                    if child.nodeType == child.ELEMENT_NODE and child.localName == "txBody":
                        tx_body = child
                        break
                if tx_body:
                    cm.insertBefore(reply_lst, tx_body)
                else:
                    cm.appendChild(reply_lst)

            # Create the reply element
            reply = doc.createElementNS(NS["p188"], "p188:reply")
            reply.setAttribute("id", new_guid())
            reply.setAttribute("authorId", author_id)
            reply.setAttribute("created", now_iso())

            # Create reply text body
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
            t.appendChild(doc.createTextNode(reply_text))
            run.appendChild(t)
            para.appendChild(run)
            tx_body.appendChild(para)
            reply.appendChild(tx_body)

            reply_lst.appendChild(reply)
            write_xml(doc, comment_file_path)
            return True

    return False


def main():
    if len(sys.argv) != 3:
        print(f"Usage: {sys.argv[0]} <pptx_path> <replies_json_path>")
        sys.exit(1)

    pptx_path = os.path.abspath(sys.argv[1])
    replies_json_path = os.path.abspath(sys.argv[2])

    if not os.path.exists(pptx_path):
        print(f"Error: {pptx_path} not found")
        sys.exit(1)
    if not os.path.exists(replies_json_path):
        print(f"Error: {replies_json_path} not found")
        sys.exit(1)

    with open(replies_json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    author_name = data.get("author_name", "Claude")
    author_initials = data.get("author_initials", "CL")
    replies = data.get("replies", [])

    if not replies:
        print("No replies to add.")
        return

    # Unpack
    unpack_dir = tempfile.mkdtemp(prefix="add_replies_")
    print(f"Unpacking {pptx_path}...")
    with zipfile.ZipFile(pptx_path, "r") as zf:
        zf.extractall(unpack_dir)

    # Find or create Claude author
    author_id = find_or_create_claude_author(unpack_dir, author_name, author_initials)
    if author_id is None:
        print("Error: No authors.xml found — no existing comments to reply to")
        shutil.rmtree(unpack_dir, ignore_errors=True)
        sys.exit(1)
    print(f"Author: {author_name} ({author_id})")

    # Find all comment files
    comments_dir = os.path.join(unpack_dir, "ppt", "comments")
    comment_files = []
    if os.path.exists(comments_dir):
        comment_files = [
            os.path.join(comments_dir, f)
            for f in sorted(os.listdir(comments_dir))
            if f.endswith(".xml")
        ]

    # Process each reply
    total_added = 0
    for r in replies:
        comment_id = r["comment_id"]
        reply_text = r["reply_text"]
        found = False

        for cf in comment_files:
            if add_reply_to_comment(cf, comment_id, author_id, reply_text):
                print(f"  Added reply to comment {comment_id[:20]}...")
                found = True
                total_added += 1
                break

        if not found:
            print(f"  WARNING: Comment {comment_id} not found in any file")

    # Repack
    print(f"Repacking to {pptx_path}...")
    if os.path.exists(pptx_path):
        os.remove(pptx_path)
    with zipfile.ZipFile(pptx_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for dirpath, dirnames, filenames in os.walk(unpack_dir):
            for filename in filenames:
                file_path = os.path.join(dirpath, filename)
                arcname = os.path.relpath(file_path, unpack_dir)
                zf.write(file_path, arcname)

    # Cleanup
    shutil.rmtree(unpack_dir, ignore_errors=True)

    print(f"Done! Added {total_added} replies to {pptx_path}")


if __name__ == "__main__":
    main()
