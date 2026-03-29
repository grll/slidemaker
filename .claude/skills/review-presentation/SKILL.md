---
name: review-presentation
description: Review an existing presentation by downloading thumbnails and analyzing each slide visually — use when the user asks to check, review, or improve a deck
allowed-tools: Bash(source *), Read, Glob
argument-hint: <presentation_id or url>
---

# Review Presentation

Download thumbnails for a presentation and provide a visual review with actionable feedback.

## Arguments

`$ARGUMENTS` — presentation ID or Google Slides URL.

Extract the presentation ID from a URL like `https://docs.google.com/presentation/d/PRESENTATION_ID/edit`.

## Steps

### 1. Get current state

```bash
source venv/bin/activate && python slidemaker.py get <presentation_id>
```

Note the number of slides and their content.

### 2. Download thumbnails

```bash
source venv/bin/activate && python slidemaker.py thumbnails <presentation_id>
```

### 3. Review each slide

Read every thumbnail image. For each slide, evaluate:

**Text quality:**
- Is text truncated or overflowing its container?
- Are there awkward line breaks (e.g., "ONBOARDIN G")?
- Is the font size readable but not too large?
- Is there too much text for the slide layout?

**Layout:**
- Do elements overlap in unintended ways?
- Is there good visual balance?
- Are decorative elements (icons, shapes) obscured by text?
- Is whitespace used effectively?

**Content:**
- Does the content match the slide's visual purpose?
- Is the information hierarchy clear (title → subtitle → body)?
- Are numbers and metrics prominently displayed?

**Consistency:**
- Is the header/title style consistent across slides?
- Are similar slides using similar text lengths?

### 4. Report findings

Present findings as a table:

| Slide | Issue | Severity | Fix |
|-------|-------|----------|-----|
| 2 | Title overflows | High | Shorten or reduce fontSize |
| 4 | Text overlaps icon | Medium | moveElement or shorten |

### 5. Ask before fixing

Present the findings and ask the user if they want you to apply fixes. Do not edit without confirmation.
