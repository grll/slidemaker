---
name: create-presentation
description: Create a new Google Slides presentation from a template — use when the user describes a presentation idea or topic
allowed-tools: Bash(source *), Read, Glob
---

# Create Presentation

Create a new Google Slides presentation from a registered template based on the user's description.

## Arguments

`$ARGUMENTS` — description of the presentation (topic, audience, key points, etc.)

## Registered templates

!`source venv/bin/activate && python slidemaker.py templates 2>/dev/null`

## Steps

### 1. Understand the request

Parse `$ARGUMENTS` to identify:
- Topic and title
- Target audience
- Key messages or data points
- Number of slides needed
- Tone (formal, casual, pitch, internal, etc.)

### 2. Browse template thumbnails

Look at the template slide thumbnails to pick the best layouts for the content. Thumbnails are at `templates/<name>/thumbnails/slide_XX.png`.

Read the catalog to see available slides and their text elements:

```bash
source venv/bin/activate && python slidemaker.py inspect --template <name>
```

Choose slides based on **visual fit** (look at thumbnails) and **content structure** (how many text elements, what kind of layout).

### 3. Design the storyline

Before creating, outline the narrative arc:
1. Hook/title
2. Problem or context
3. Solution or approach
4. Evidence/data/steps
5. Results or impact
6. Call to action or conclusion

Map each beat to a template slide.

### 4. Build the create command

Construct the JSON with:
- `title`: presentation title
- `template`: template name
- `keep_slides`: ordered list of slide indices
- `replacements`: element objectId → content text

**Important**: Look at element text lengths in the catalog. Template text boxes are sized for the original placeholder text. Keep replacement text similar in length, or plan to adjust font sizes in step 6.

```bash
source venv/bin/activate && python slidemaker.py create --template <name> '<json>'
```

### 5. Review with thumbnails

Download and visually inspect every slide:

```bash
source venv/bin/activate && python slidemaker.py thumbnails <presentation_id>
```

Read each thumbnail image and check for:
- Text overflow or truncation
- Awkward line breaks in titles
- Text overlapping decorative elements
- Empty areas that should have content
- Overall visual balance

### 6. Fix issues

For each problem found, apply the right fix:
- **Text overflow** → reduce font size with `textStyle` (prefer this over shortening text)
- **Bad positioning** → `moveElement` to adjust
- **Missing emphasis** → `textStyle` with bold/color
- **Wrong content** → `replaceText`

```bash
source venv/bin/activate && python slidemaker.py edit <presentation_id> '[...]'
```

### 7. Final review

Download thumbnails again and verify all issues are resolved. Repeat steps 6-7 until the deck looks clean.

Present the final URL to the user.

> [!CAUTION]
> Always show the user your slide plan (slide_plan.md) (which template slides, what content) before creating. Creating a presentation generates a new file in their Google Drive.
