# refine-ppt

A [Claude Code skill](https://code.claude.com/docs/en/skills) that reviews PowerPoint presentations and adds suggestions as **native comments** — the same kind you'd add via Review > New Comment. Your original slide text is never modified.

## What it does

- Reads your `.pptx` file and adds comments anchored to specific text ranges
- Each comment contains a suggested replacement for the highlighted text
- You review suggestions in PowerPoint's comment sidebar and accept or ignore each one
- Supports a reply loop: respond to comments in PowerPoint, then ask Claude to revise

## Requirements

- **Claude Code** (the CLI tool)
- **Python 3.8+** with `markitdown` installed: `pip install markitdown[pptx]`
- **PowerPoint for Microsoft 365** or **PowerPoint Online** to view comments (older versions like 2019 won't display modern comments)

## Install

Clone this repo into your Claude Code skills directory:

```bash
git clone https://github.com/YOUR_USERNAME/refine-ppt.git ~/.claude/skills/refine-ppt
```

Or if you want it available only in a specific project:

```bash
git clone https://github.com/YOUR_USERNAME/refine-ppt.git .claude/skills/refine-ppt
```

## Usage

### Review a presentation

```
/refine-ppt ~/Documents/my_deck.pptx make it more concise and professional
```

### Review specific slides

```
/refine-ppt ~/Documents/my_deck.pptx review slides 2 and 3 for an executive audience
```

### Respond to your feedback

After you reply to comments in PowerPoint and save:

```
/refine-ppt ~/Documents/my_deck.pptx read my replies and respond with revised suggestions
```

## How it works

The skill bundles three Python scripts that manipulate the PowerPoint XML directly:

| Script | Purpose |
|--------|---------|
| `scripts/add_comments.py` | Adds new text-anchored modern comments |
| `scripts/read_comments.py` | Reads all comment threads as JSON |
| `scripts/add_replies.py` | Adds replies to existing comment threads |

Comments use the [modern comment format](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-pptx/161bc2c9-98fc-46b7-852b-ba7ee77e2e54) (Office 365) with text-range anchoring via `txMkLst` monikers. The scripts handle all the XML complexity — finding shapes, computing text hashes, injecting creation IDs, and managing relationships.

## License

MIT
