---
name: hp-pptx
description: Convert any markdown (notes, docs, READMEs, meeting summaries, architecture write-ups, status reports) into a PowerPoint deck in the official HP corporate template. Use this skill whenever the user wants slides, a deck, a presentation or a .pptx made from markdown or text, mentions the HP template, HP branding or "HP-style" slides, or asks to turn a document into a presentation - even if they don't say "HP" or "template" explicitly.
---

# HP markdown-to-PowerPoint

This skill turns markdown into a deck built on the real HP template (`assets/hp_template.pptx`):
HP fonts, colours, logo, footer, title slide, section dividers and closing slide. The layout work
is done by a tested script, `scripts/md2hp.py`. Your job is the part a script can't do well:
**turning the user's markdown into slide-sized content** written in the deck-markdown format,
then checking the result.

## Workflow

1. **Read the source markdown fully.** Note the audience and purpose if stated. If the user gives
   several files, read them all.
2. **Write `deck.md`** (next to the source, or where the user asks) in the format described in
   `references/markdown-spec.md`. Read that file before writing your first deck. Follow the
   content rules below. Never edit the user's original file.
3. **Show the user the slide plan** (one line per slide: title and layout) and wait for approval if
   the deck has more than about 8 slides or the source needed heavy restructuring. For small,
   simple conversions, go straight to the build.
4. **Build:**
   ```bash
   python <skill-dir>/scripts/md2hp.py deck.md -o deck.pptx
   ```
   The script prints the layout chosen for each slide and a `WARN` line for anything it could not
   fit. Exit code 1 means there were warnings.
5. **Fix every WARN by editing `deck.md`** (shorten text, split the slide, or force a layout with
   `<!-- layout: ... -->`) and rebuild. Do not edit the script to fix content problems.
6. **Visual check:**
   ```bash
   bash <skill-dir>/scripts/render.sh deck.pptx
   ```
   This writes one JPG per slide next to the deck. If you can view images, look at every slide
   and check for cut-off text, overlaps, near-empty slides and wrong layouts. If you cannot view
   images, run `python -m markitdown deck.pptx` and read the full text instead, and tell the user
   to page through the deck once in PowerPoint.
7. **Report** the deck path, the slide list, anything you shortened or restructured, and every
   assumption.

## Content rules (why they matter)

Slides are read in seconds, so the source usually needs restructuring, not copying.

- **One message per slide.** The `##` title should say what the slide is about. Split a source
  section that makes two points into two slides.
- **Bullets: at most 6 per slide, at most ~15 words each.** Longer bullets get shrunk or split
  by the script, which reads worse than writing them short in the first place.
- **Turn prose into structure.** Paragraphs become bullets; comparisons become a table; sequences
  become a numbered list (process layout); before/after becomes the compare layout; 2-4 parallel
  themes become `###` groups (columns layout); headline metrics become the stats layout.
- **Keep one lead sentence** (a plain paragraph right under the title) when it adds context, and
  use a `>` blockquote for the single most important takeaway or decision on the slide.
- **Never invent facts, numbers, owners or dates.** If the source is missing something the deck
  needs, write `TBC` and list it in your report.
- **Keep the user's terminology**, product names and numbers exactly as written.
- Put detail the presenter needs (but the audience doesn't) in `<!-- notes: ... -->`.

## Choosing a layout

The script picks a layout automatically from the content. Force one only when the automatic
choice is wrong. Full syntax and examples are in `references/markdown-spec.md`.

| Content in the slide | Automatic layout |
| --- | --- |
| Bullets and/or paragraphs | `bullets` (two columns when there are many short bullets) |
| A markdown table | `table` (splits across slides if long) |
| One numbered list of 3-6 steps | `process` (numbered circles with arrows); 7+ steps become stacked cards |
| Two `###` groups named like Current / Proposed (or Before / After) | `compare` (paired cards, optional `[Tag]` per row) |
| 2-4 `###` groups | `columns` (2 = split columns, 3-4 = cards) |
| Bullets that each start with a bold number (`**42%** label`) | `stats` (big-number cards) |
| Only a `>` quote | `quote` |
| An image `![caption](file.png)` | `image` |
| A ```` ```flow ```` block | `flow` (boxes and arrows) |
| A plain code block | `code` |

## Front matter

```yaml
---
title: Deck title            # title slide
subtitle: Architecture Point of View
author: Name                 # shown as "Author: Name"
date: 06 Oct 2026            # shown as "Published: ..."
title_prefix: Systems & Infrastructure   # slide titles become "prefix | Title"; set to "" for plain titles
footer: HP Confidential
closing: true                # set false to omit the HP closing slide
---
```

## Brand rules the script already enforces

You don't need to style anything, but know the rules so your content fits them: HP Simplified
font only; HP blue `0096D6` for emphasis; purple `6B3A97` only for IT-owned systems in flow
diagrams; clean tables (blue header text, grey row rules, no fills); no decorative shapes. See
`references/design-system.md` if the user asks for something the layouts don't cover.

## Troubleshooting

- **Fonts look wider in rendered JPGs.** The build machine usually lacks HP Simplified, so the
  preview uses a wider fallback. If text fits in the preview, it fits in PowerPoint.
- **`soffice` or `pdftoppm` not found.** Rendering needs LibreOffice and Poppler. The deck itself
  builds without them; skip step 6 and tell the user.
- **A layout didn't match** (e.g. `process` chosen for a numbered list that isn't a sequence):
  add `<!-- layout: bullets -->` under the `##` title.
- **Flow diagram looks crowded:** keep flows to 8 boxes or fewer, labels under 25 characters;
  for bigger diagrams, use an image instead.

## Files

- `scripts/md2hp.py` - the converter (python-pptx). Requires `pip install python-pptx lxml pillow`.
- `scripts/hp_helpers.py` - drawing helpers used by the converter.
- `scripts/render.sh` - deck -> PDF -> one JPG per slide.
- `assets/hp_template.pptx` - the HP template (masters and layouts, no slides).
- `references/markdown-spec.md` - the deck-markdown format with an example of every layout.
- `references/design-system.md` - HP colours, type and component rules.
- `examples/` - two complete sample decks in deck-markdown.
