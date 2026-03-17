---
name: refine-ppt
description: >
  Add review comments to PowerPoint presentations with suggested text changes,
  and respond to user feedback on those comments. Adds native PowerPoint
  comments anchored to specific text ranges so the user can accept or ignore
  each suggestion. Reads existing user comments (instructions like "make this
  more concise") and replies with suggested rewrites. Also reviews uncommented
  text independently. Use whenever a user wants to refine, review, or improve
  a .pptx file.
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

## Workflow

The user invokes this skill with `/refine-ppt <pptx-path> [direction]`. Parse the arguments: `$ARGUMENTS`

Every run follows the same two-phase approach. User comments are instructions — they take priority over your own judgment about what to change.

### Phase 1: Respond to existing user comments

User comments are directives. If a user highlighted a paragraph and wrote "make this shorter" or "rewrite for executives", that is an instruction you must follow. Your response goes as a **reply** to their comment — not as a new comment.

1. Run `read_comments.py` to get all comment threads as JSON.
2. Identify comments left by users (i.e., not authored by "Claude"). Look at root-level comments that have no reply from Claude yet.
3. For each user comment, read the anchored text on the slide and follow the user's instruction to write a suggested replacement.
4. Write a JSON file and run `add_replies.py`:

```json
{
  "author_name": "Claude",
  "author_initials": "CL",
  "replies": [
    {
      "comment_id": "{GUID-from-read-output}",
      "reply_text": "Suggested replacement text based on the user's instruction"
    }
  ]
}
```

If there are no user comments, skip this phase.

### Phase 2: Independent review of uncommented text

After handling user comments, review the remaining text that has no comments attached.

1. Extract the full text: `python -m markitdown "<pptx_path>"`
2. Cross-reference with the comment data from Phase 1 to identify which text already has comments (either user comments you just replied to, or previous Claude comments). Skip that text.
3. For the remaining uncommented text, decide what needs changes based on the user's direction. If they specified slide numbers, only review those slides.
4. Write a JSON file and run `add_comments.py`:

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

If the user didn't ask for a general review (they only left specific comments and gave no additional direction), skip this phase.

## Key rules

- **Never edit slide text directly.** Every suggestion goes in a comment or reply.
- **User comments are priority.** Always handle them before doing independent review.
- Comments require PowerPoint for Microsoft 365 or PowerPoint Online (not 2019 or earlier).
- Tell the user which slides were affected when done, distinguishing between replies to their comments and new suggestions.
