---
name: edit-presentation
description: Edit an existing Google Slides presentation — use when the user wants to change content, styling, layout, or structure of a deck
allowed-tools: Bash(source *), Read, Glob
argument-hint: <presentation_id or url> <changes>
---

# Edit Presentation

Modify an existing presentation based on the user's instructions.

## Arguments

`$ARGUMENTS` — presentation ID (or URL) followed by a description of desired changes.

## Steps

### 1. Read current state

```bash
source venv/bin/activate && python slidemaker.py get -d <presentation_id>
```

This returns all slides, elements, positions, sizes, and font info. Use this to understand the current layout before making changes.

### 2. Download thumbnails for context

```bash
source venv/bin/activate && python slidemaker.py thumbnails <presentation_id>
```

Read the thumbnails to visually understand the current state.

### 3. Plan the edits

Based on the user's request, determine which operations to use:

| User wants | Operation(s) |
|------------|-------------|
| Change text on a slide | `replaceText` |
| Change text everywhere | `replaceAllText` |
| Make text bigger/smaller | `textStyle` with `fontSize` |
| Bold, italic, color | `textStyle` |
| Center or align text | `paragraphStyle` with `alignment` |
| Move a shape | `moveElement` |
| Resize a shape | `resizeElement` |
| Change shape color | `shapeFill` |
| Add a border | `shapeOutline` |
| Change slide background | `slideBackground` |
| Add a new text box | `createShape` + `insertText` + `textStyle` |
| Add an image | `addImage` |
| Add a table | `createTable` + `insertTableText` |
| Remove an element | `deleteElement` |
| Remove a slide | `deleteSlide` |
| Reorder slides | `moveSlide` |
| Duplicate a slide | `duplicateSlide` |
| Add a line/arrow | `createLine` |

### 4. Apply edits

Batch as many operations as possible into a single `edit` call:

```bash
source venv/bin/activate && python slidemaker.py edit <presentation_id> '[op1, op2, ...]'
```

**Ordering matters**: operations are applied sequentially. Put `createShape` before `insertText` on that shape. Put `deleteText`/`replaceText` before `textStyle` if changing both content and style.

### 5. Verify with thumbnails

```bash
source venv/bin/activate && python slidemaker.py thumbnails <presentation_id>
```

Read each affected slide's thumbnail to confirm:
- Changes applied correctly
- No new layout issues introduced
- Overall slide still looks good

If issues remain, apply further fixes and re-verify.

## Tips

- When changing text, use `get -d` first to see the current `fontSize`. Match or reduce it to avoid overflow.
- Group related edits (e.g., `replaceText` + `textStyle` on the same element) in one `edit` call.
- For adding new content (text box + text + styling), you need 3 operations in order: `createShape`, `insertText`, `textStyle`.
- The slide coordinate system is 720 x 405 points. Use this to position elements sensibly.

> [!CAUTION]
> Confirm the edit plan with the user before applying changes, especially for destructive operations (deleting slides/elements).
