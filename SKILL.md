---
name: refine-ppt
description: >
  Add review comments to PowerPoint presentations with suggested text changes,
  and respond to user feedback on those comments. Adds native PowerPoint
  comments anchored to specific text ranges so the user can accept or ignore
  each suggestion. Also handles reading user replies and responding with
  revised suggestions.
disable-model-invocation: true
argument-hint: <pptx-path> [direction]
---

# Refine PPT

You add review suggestions to PowerPoint files as native comments anchored to specific text — the same kind a human would add via Review > New Comment. The original slide text is **never modified**. All suggestions live in the comment sidebar so the user stays in control.

You have three bundled scripts in `scripts/` (relative to this SKILL.md):

| Script | Purpose |
|--------|---------|
| `scripts/add_comments.py <pptx> <json>` | Add new text-anchored comments |
| `scripts/read_comments.py <pptx>` | Read all comment threads (outputs JSON to stdout) |
| `scripts/add_replies.py <pptx> <json>` | Add replies to existing comment threads |

## Adding comments (initial review)

The user invokes this skill with `/refine-ppt <pptx-path> [direction]`. Parse the arguments: `$ARGUMENTS`

1. Extract the text: `python -m markitdown "<pptx_path>"`
2. Decide which text needs changes based on the user's direction. If they specified slide numbers, only review those slides.
3. Write a JSON file and run `add_comments.py`:

```json
{
  "author_name": "Claude",
  "author_initials": "CL",
  "comments": [
    {
      "slide_number": 2,
      "highlight_text": "exact verbatim substring from the slide",
      "comment_text": "Your suggested replacement"
    }
  ]
}
```

`highlight_text` must be a verbatim substring of text on that slide. The script handles everything else (finding shapes, computing hashes, injecting XML). It overwrites the original file.

## Responding to replies

1. Run `read_comments.py` to get all threads as JSON
2. Find comments where the user has replied
3. Write a JSON file and run `add_replies.py`:

```json
{
  "author_name": "Claude",
  "author_initials": "CL",
  "replies": [
    {
      "comment_id": "{GUID-from-read-output}",
      "reply_text": "Revised suggestion based on feedback"
    }
  ]
}
```

## Key rules

- **Never edit slide text directly.** Every suggestion goes in a comment.
- Comments require PowerPoint for Microsoft 365 or PowerPoint Online (not 2019 or earlier).
- Tell the user which slides were affected when done.
