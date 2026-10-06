# HP design system (as used by md2hp.py)

Use this when the user asks for something the built-in layouts don't cover, so any custom
slide still matches the template. Slide size is 13.33 x 7.5 in (16:9).

## Colours

| Role | Hex | Use |
| --- | --- | --- |
| HP blue (theme accent 1) | 0096D6 | Title section name, table headers, emphasis, S&I-owned systems, new flows |
| Purple (theme accent 2) | 6B3A97 | IT-owned systems only |
| Near-black | 1A1A1A | Body text, current flows |
| Dark grey | 535455 | Secondary text, lead lines |
| Mid grey | 87898B | Muted labels, step connector arrows |
| Light grey fill | F3F4F4 | Cards, as-is items, stat tiles |
| Blue tint fill | E6F4FB | Callouts, to-be items |
| Rule grey | D9D9D9 | Table row lines |

Status colours (green 2E7D32, amber C77700, red C62828) only for status indicators.

## Type

- Font: HP Simplified Light (theme font). Never set another font, except Consolas for code.
- Slide titles: template title placeholder; "Prefix | " in black, section name in HP blue.
- Body: 13-18 pt for bullets, 9.5-14 pt for tables, 8-9 pt for diagram labels.

## Components

- Tables: blue capitalised header text (not bold), 1 pt grey line under each row, no fills.
- System boxes: white fill, 1.5 pt outline in the owner's colour; dashed outline = new.
- Arrows: solid near-black = current; dashed HP blue = new or changed.
- Cards: flat rounded rectangles, no outline, no shadow.
- Spacing: content from x 0.67 to 12.67 in, y 1.35 to 6.85 in; footer at 7.08 in.
- Never add decorative bars or lines under titles, gradients, clip-art or stock icons.

## Template layouts used

| Layout name in hp_template.pptx | Used for |
| --- | --- |
| Title Slide | Title slide |
| Blue Section Header | `#` section dividers |
| Title Only | All content slides |
| Closing Option 2 | Closing slide |

Other layouts in the template (Two Content, Quote, Picture with Caption, etc.) are available if
you add a custom slide with python-pptx: `prs.slide_layouts` by name.
