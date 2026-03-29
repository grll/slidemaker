---
name: slidemaker
description: Reference for the slidemaker CLI tool — use when working with Google Slides presentations (creating, editing, reviewing, or inspecting templates)
user-invocable: false
---

# Slidemaker CLI Reference

Slidemaker creates and edits Google Slides presentations from registered templates. It supports two backends (direct Google API or Apps Script web app), selected automatically via `.env`.

## Commands

| Command | Description |
|---------|-------------|
| `auth` | Test connection |
| `register <pres_id> <name>` | Register a template (downloads catalog + thumbnails) |
| `templates` | List registered templates |
| `inspect [pres_id] [--template name]` | Show slide elements and object IDs |
| `create [--template name] '<json>'` | Create presentation from template |
| `get <pres_id> [-d]` | Read presentation content (`-d` for position/size/font) |
| `edit <pres_id> '<json>'` | Edit a presentation |
| `thumbnails <pres_id> [-o dir]` | Download slide PNGs for visual review |

Always run commands with: `source venv/bin/activate && python slidemaker.py <command>`

## Create JSON format

```json
{
  "title": "Presentation Title",
  "template": "infographics",
  "keep_slides": [0, 5, 13],
  "replacements": {
    "element_object_id": "New text content"
  }
}
```

- `keep_slides`: template slide indices to keep, in desired order
- `replacements`: map element objectId → new text (IDs from catalog)

## Edit operations

All operations are passed as a JSON array to the `edit` command.

### Text operations
| Operation | Format |
|-----------|--------|
| `replaceText` | `{"replaceText": {"objectId": "id", "text": "new"}}` |
| `insertText` | `{"insertText": {"objectId": "id", "text": "new"}}` |
| `replaceAllText` | `{"replaceAllText": {"find": "old", "replace": "new"}}` |
| `textStyle` | `{"textStyle": {"objectId": "id", "fontSize": 12, "bold": true, "italic": false, "color": "#FF0000", "fontFamily": "Arial"}}` |
| `paragraphStyle` | `{"paragraphStyle": {"objectId": "id", "alignment": "CENTER", "lineSpacing": 115}}` |

### Shape operations
| Operation | Format |
|-----------|--------|
| `createShape` | `{"createShape": {"pageId": "slide_id", "shapeType": "TEXT_BOX", "x": 100, "y": 100, "width": 200, "height": 50}}` |
| `shapeFill` | `{"shapeFill": {"objectId": "id", "color": "#336699"}}` |
| `shapeOutline` | `{"shapeOutline": {"objectId": "id", "color": "#FFF", "weight": 2, "dashStyle": "SOLID"}}` |
| `moveElement` | `{"moveElement": {"objectId": "id", "x": 100, "y": 200}}` |
| `resizeElement` | `{"resizeElement": {"objectId": "id", "scaleX": 1.5, "scaleY": 1.5}}` |
| `deleteElement` | `{"deleteElement": {"objectId": "id"}}` |

### Slide operations
| Operation | Format |
|-----------|--------|
| `deleteSlide` | `{"deleteSlide": {"objectId": "slide_id"}}` |
| `duplicateSlide` | `{"duplicateSlide": {"objectId": "slide_id"}}` |
| `moveSlide` | `{"moveSlide": {"objectId": "slide_id", "insertionIndex": 2}}` |
| `slideBackground` | `{"slideBackground": {"pageId": "slide_id", "color": "#000000"}}` |

### Other operations
| Operation | Format |
|-----------|--------|
| `addImage` | `{"addImage": {"url": "https://...", "pageId": "slide_id", "size": {"width": 200, "height": 150}, "position": {"x": 50, "y": 50}}}` |
| `createLine` | `{"createLine": {"pageId": "slide_id", "category": "STRAIGHT", "x": 50, "y": 200, "width": 300, "height": 0}}` |
| `createTable` | `{"createTable": {"pageId": "slide_id", "rows": 3, "columns": 3, "x": 50, "y": 100, "width": 500, "height": 200}}` |
| `insertTableText` | `{"insertTableText": {"tableId": "id", "row": 0, "column": 0, "text": "Header"}}` |
| `groupObjects` | `{"groupObjects": {"objectIds": ["id1", "id2"]}}` |
| `ungroupObjects` | `{"ungroupObjects": {"objectId": "group_id"}}` |
| `raw` | `{"raw": {<any Slides API batchUpdate request>}}` |

## Tips

- Element object IDs are preserved when copying a template. Use the catalog to find them.
- Colors accept hex strings (`"#FF0000"`) or RGB dicts (`{"red": 1, "green": 0, "blue": 0}`).
- When text overflows, prefer reducing `fontSize` via `textStyle` over shortening the text.
- Use `get -d` to see element positions and font sizes before making layout adjustments.
- Always review changes with `thumbnails` after editing — Claude can read the PNGs to verify.
- The Google Slides coordinate system uses points (1 pt = 1/72 inch). Standard slide is 720x405 pt.

> [!CAUTION]
> Before creating or editing a presentation, confirm the plan with the user. Show which template slides you'll use and what content goes where.
