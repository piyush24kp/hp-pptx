# Deck-markdown format

A deck is one markdown file: optional front matter, then slides.

- `# Heading` starts a **section divider** slide (blue). An optional paragraph under it becomes
  the divider's subtitle.
- `## Heading` starts a **content slide**. Everything until the next `#`/`##` belongs to it.
- `<!-- layout: name -->` anywhere in a slide forces a layout (`bullets`, `table`, `process`,
  `compare`, `columns`, `stats`, `quote`, `image`, `flow`, `code`).
- `<!-- notes: text -->` adds speaker notes.
- Inline: `**bold**`, `*italic*`, `` `code` ``, `[link text](url)`.
- The title slide comes from front matter; an HP closing slide is added automatically.

Two elements work on most layouts:

- **Lead paragraph:** the first plain paragraph of a slide is shown as a grey lead line under the title.
- **Callout:** a `>` blockquote is shown as a blue-tinted takeaway box at the bottom.

On a slide that is only paragraphs, all paragraphs are shown as body text instead.

---

## bullets

```markdown
## Executive Summary

The migration is on track for the November cut-over.

- **Scope:** 14 of 18 services migrated
- **Cost:** run-rate down 22% versus Q2
  - rightsizing complete for compute
- **Risks:** vendor licence and payments freeze window

> **Decision needed:** approve the extended freeze window by 15 Oct.
```

One level of sub-bullets (indent 2 spaces). Eight or more short bullets switch to two columns.
Lists too long for one slide are split into "(1/2)", "(2/2)" slides automatically.

## table

```markdown
## Technical Dependencies

| Dependency | Why it matters | How to address it |
|---|---|---|
| DQ check needs sync first | In-flight items look missing | Start on a sync-complete signal |
| Batch and throttle limits | Large batches load the DB | Define batch size and retries |
```

Column widths follow content length. A short first column (labels) is bolded automatically;
columns of short values (status, sizes) are centred. Long tables split across slides.
Keep cells under ~20 words.

## process

```markdown
## Rollout Plan

1. **Approve freeze window**: steering decision by 15 Oct
2. **Migrate payments**: during the approved window
3. **Cut over**: final traffic switch in late November
4. **Decommission legacy**: by end of December
```

Each step is `**Name**: detail` (or `Name: detail`). 3-6 steps show as numbered circles joined by
arrows; 7 or more become numbered cards stacked vertically (two columns if needed).

## compare

```markdown
## Current vs. Proposed Flow

### Current (as-is)
1. **US daily recon**: S3 dump of the last 24h
2. **Email alert**: list emailed to the team
3. -

### Proposed (to-be)
1. [Unchanged] **US daily recon**: same job
2. [Replaced] **Report to S3**: machine-readable list
3. [New] **Auto-reprocess**: resend triggered by S3
```

Triggered automatically when there are exactly two `###` groups whose headings look like
current/before/today and proposed/after/target/future. Rows pair up by position. Use `-` for a
row with no equivalent. The optional `[Tag]` at the start of a right-hand item becomes a pill
between the cards: `Unchanged` (grey), `Enhanced` / `Changed` (light blue), anything else
(blue).

## columns

```markdown
## Benefits

### HP
- Protects revenue
- Cuts manual effort

### Customer / Partner
- More accurate invoices

### Employee
- Less repetitive work
```

Two groups: split columns with a dashed divider. Three or four: cards. Keep the bullet counts
similar across groups.

## stats

```markdown
## Key Numbers

- **14/18** services migrated
- **22%** lower run-rate
- **9%** change failure rate
- **3** P1 incidents, down from 7
```

Used when every bullet starts with a bold value containing a digit. Up to 4 per row.

## quote

```markdown
## What We Heard

> The new pipeline cut our release time from two days to forty minutes. — Payments squad lead
```

A slide containing only a quote. Text after a final ` — ` becomes the attribution.

## image

```markdown
## Revenue Trend

![Monthly platform revenue, Apr to Sep](chart.png)
```

Path is relative to the markdown file. The image is scaled to fit; the alt text becomes the caption.

## flow

````markdown
## Target Architecture

```flow
Usage Service -> BIRD : 1. SNS publish
Usage Service -> Convergent Mediation [it] : 1. SNS publish
Usage Service --> BIRD : 3. generateXML sync
BIRD -> DQ Check : 4. after sync
DQ Check --> Missed-Txn Report [new] : 5. CM-only gaps
```
````

- `->` current flow (solid), `-->` new flow (dashed blue). Chains like `A -> B -> C` work.
- `: label` after an edge labels it (keep under 25 characters).
- `[it]` marks an IT-owned system (purple outline); `[new]` marks a new component (dashed).
- Boxes are arranged left to right by flow order. Keep it to 8 boxes or fewer.
- A legend is added when any `[it]`, `[new]` or `-->` is used.

## code

````markdown
## Deployment Command

```
./deploy.sh --env prod --wave 4
```
````

Monospace on a light panel. Keep to about 15 lines.

---

## Converting free-form markdown: a worked example

Source (a typical doc section):

```markdown
## Background
Our reconciliation process today relies on a daily job in Usage Service that dumps
everything sent in the last 24 hours to S3. BIRD then compares this with what it
received over SNS. Separately, a DQ check compares BIRD with CM, but because the two
jobs run in parallel the DQ check never learns what BIRD found missing. When gaps are
found an email goes to the team and someone resends each transaction by hand.
```

Deck-markdown (one message, structure instead of prose):

```markdown
## How Reconciliation Works Today

Three jobs run independently, and recovery is manual.

1. **US daily recon (T+1)**: dumps every transaction sent in the last 24h to S3
2. **BIRD daily recon (T+2)**: compares the dump with what BIRD received via SNS
3. **DQ check (T+2)**: compares BIRD with CM, in parallel, blind to step 2
4. **Email + manual resend**: someone resends each missed transaction by hand

> Because steps 2 and 3 run in parallel, transactions BIRD never received are never reported.
```
