# hp-pptx: markdown to HP-template PowerPoint (agent skill)

Turns any markdown into a deck on the official HP template. Built for agent tools such as
GitHub Copilot CLI with a Claude Sonnet model; the agent restructures your markdown into
slide-sized content, runs the converter, and checks the result.

## Install (once)

1. **Python packages:** `pip install -r requirements.txt`
2. **For visual QA (recommended):** LibreOffice and Poppler
   - macOS: `brew install --cask libreoffice && brew install poppler`
   - Ubuntu/Debian: `sudo apt install libreoffice poppler-utils`
   - Windows: install LibreOffice, and Poppler for Windows; add both `bin` folders to PATH
     (or run inside WSL)
3. **Put the skill where Copilot CLI finds skills:**
   - for one repo: copy this folder to `.github/skills/hp-pptx/` in that repo
   - for all your work: copy it to your personal skills folder (commonly `~/.copilot/skills/hp-pptx/`)

   Skill folder locations can differ between Copilot CLI versions. Check GitHub's Copilot CLI
   documentation for your version. If your version doesn't load skills, add this line to the
   repo's `AGENTS.md` (or `.github/copilot-instructions.md`):
   `For any request to make slides or a deck, read and follow .github/skills/hp-pptx/SKILL.md.`
4. **Choose the model:** start `copilot`, run `/model` and pick a Claude Sonnet model
   (availability depends on your Copilot plan and organisation policy).

## Use

In Copilot CLI, from the folder with your markdown:

```
Use the hp-pptx skill to turn notes/q3-update.md into an HP deck for the steering
committee. Show me the slide plan first.
```

Or run the converter yourself on a file already in deck-markdown format:

```
python scripts/md2hp.py examples/status-update-example.md -o status.pptx
bash scripts/render.sh status.pptx      # optional: slide previews as JPGs
```

## Test it

```
python scripts/md2hp.py examples/architecture-example.md -o arch.pptx
python scripts/md2hp.py examples/status-update-example.md -o status.pptx
```

Both should finish with no `WARN` lines.

## What's inside

| Path | Purpose |
| --- | --- |
| `SKILL.md` | Instructions the agent follows |
| `scripts/md2hp.py` | Converter: layouts, auto-fit, auto-split, warnings |
| `scripts/hp_helpers.py` | Drawing helpers (tables, boxes, arrows, titles) |
| `scripts/render.sh` | Deck to per-slide JPGs for checking |
| `assets/hp_template.pptx` | HP template: masters and layouts, no slides |
| `references/markdown-spec.md` | Input format, with an example of every layout |
| `references/design-system.md` | HP colours, type and component rules |
| `examples/` | Two complete sample decks |

## Note on fonts

The template uses HP Simplified. Machines without it render previews with a wider fallback
font; text that fits in the preview fits in PowerPoint.
