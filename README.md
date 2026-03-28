# Slidemaker

Create and edit Google Slides presentations from templates, driven by Claude.

Supports multiple templates with visual thumbnails. Claude picks the right slides for your content by looking at the actual slide designs, populates them, and iterates by reviewing thumbnails of the result.

## Setup

Choose one of two backends depending on your Google account situation.

### Option A: Apps Script (recommended for Workspace / corporate accounts)

No Google Cloud project needed. Everything runs inside Google's own sandbox.

1. Go to [script.google.com](https://script.google.com) and create a new project
2. Replace the default code with the contents of `appscript/Code.gs`
3. In the left sidebar, click **Services** → **+** → add **Google Slides API** (listed as `Slides`, keep default version `v1`)
4. Click **Deploy** → **New deployment**
   - Type: **Web app**
   - Execute as: **Me**
   - Who has access: **Anyone** (or **Anyone in [your organization]**)
5. Click **Deploy**, authorize when prompted, and copy the web app URL
6. Create a `.env` file in the project root:
   ```
   WEBAPP_URL=https://script.google.com/macros/s/XXXX/exec
   ```
7. Test it:
   ```
   python slidemaker.py auth
   ```

### Option B: Direct Google API (for personal accounts)

Requires a Google Cloud project with OAuth credentials.

1. Go to [Google Cloud Console](https://console.cloud.google.com)
2. Create a project (or use an existing one)
3. Enable **Google Slides API** and **Google Drive API** (APIs & Services → Library)
4. Go to **APIs & Services** → **Credentials** → **Create Credentials** → **OAuth 2.0 Client ID**
   - Application type: **Desktop app**
5. Download the JSON file, rename it to `credentials.json`, and place it in the project root
6. Install dependencies:
   ```
   python -m venv venv
   source venv/bin/activate
   pip install -r requirements.txt
   ```
7. Authenticate (opens browser):
   ```
   python slidemaker.py auth
   ```

The backend is selected automatically: if `.env` contains `WEBAPP_URL`, the Apps Script backend is used. Otherwise, the direct API backend is used.

## Register a template

Before creating presentations, register at least one template:

```
python slidemaker.py register <presentation_id> <name>
```

This downloads the slide catalog (text elements + object IDs) and a PNG thumbnail for every slide. Thumbnails are stored in `templates/<name>/thumbnails/`.

List registered templates:

```
python slidemaker.py templates
```

## Usage

### Inspect a template

Show all slides and their text elements (with object IDs for targeting replacements):

```
python slidemaker.py inspect --template <name>
```

Or inspect any presentation by ID:

```
python slidemaker.py inspect <presentation_id>
```

### Create a presentation

```
python slidemaker.py create --template <name> '{
  "title": "Q1 Review",
  "keep_slides": [0, 5, 13],
  "replacements": {
    "element_object_id": "New text",
    "another_element_id": "More text"
  }
}'
```

- `keep_slides`: indices of template slides to keep (in order)
- `replacements`: map of element object ID → new text
- `template` can also be specified inside the JSON as `"template": "<name>"`

Returns `{"presentationId": "...", "url": "..."}`.

### Review a presentation (visual feedback)

Download thumbnails of a created presentation to visually review the result:

```
python slidemaker.py thumbnails <presentation_id>
```

Thumbnails are saved to `review/` by default (override with `--output`). Claude reads these images to spot text overflow, layout issues, or content problems, then edits the presentation to fix them.

### Read a presentation

```
python slidemaker.py get <presentation_id>
```

Returns all slides with their text elements and object IDs as JSON.

### Edit a presentation

```
python slidemaker.py edit <presentation_id> '[
  {"replaceText": {"objectId": "element_id", "text": "new text"}},
  {"deleteSlide": {"objectId": "slide_id"}},
  {"duplicateSlide": {"objectId": "slide_id"}},
  {"moveSlide": {"objectId": "slide_id", "insertionIndex": 2}}
]'
```

Raw Slides API requests can also be passed via `{"raw": {<batchUpdate request>}}`.

## Workflow

1. **Register** a template once (`register`)
2. **Browse** template thumbnails to pick the right slides for your content
3. **Create** a presentation with selected slides and text replacements (`create`)
4. **Review** the result by downloading thumbnails (`thumbnails`)
5. **Edit** to fix any issues spotted in the review (`edit`)
6. Repeat 4-5 until the slides look right
