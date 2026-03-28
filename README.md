# Slidemaker

Create and edit Google Slides presentations from a template, driven by Claude.

Uses a [Slidesgo infographic template](https://docs.google.com/presentation/d/1cWqfy4vpwbmlgPaN09ha02QILAj3phf_akY9C4VqtVE) with 35 pre-designed slide layouts. Claude picks the right slides for your content, populates them, and can refine them through conversation.

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

## Usage

### Inspect the template

Show all slides and their text elements (with object IDs for targeting replacements):

```
python slidemaker.py inspect
```

### Create a presentation

```
python slidemaker.py create '{
  "title": "Q1 Review",
  "keep_slides": [0, 5, 13],
  "replacements": {
    "g708a6ee8a1_0_59": "Q1 REVIEW\n2025",
    "g708a6ee8a1_0_60": "Company performance overview"
  }
}'
```

- `keep_slides`: indices of template slides to keep (in order)
- `replacements`: map of element object ID → new text

Returns `{"presentationId": "...", "url": "..."}`.

### Read a presentation

```
python slidemaker.py get <presentation_id>
```

Returns all slides with their text elements and object IDs.

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

## Template slides

| Index | Layout | Use for |
|-------|--------|---------|
| 0 | Title slide | Opening slide with title + subtitle |
| 1 | About/intro | Title + paragraph |
| 2 | SWOT | 4-quadrant analysis |
| 3 | Timeline | 4 time periods |
| 5 | 4-item comparison | Feature comparison, differentiators |
| 6 | What We Do | 4 items with percentages |
| 8 | Sales funnel | 4-stage funnel |
| 9 | Process | 4 numbered steps |
| 10 | User persona | Bio, motivations, traits |
| 11 | 6-item process | How it works |
| 13 | 3-step process | Simple steps |
| 14 | Phases | 3 numbered phases |
| 15 | Services | 5 numbered items |
| 22 | Strategy | 3 numbered strategy items |
| 24 | Solutions | 3 items |
| 26 | Events | 4 events |
| 27 | Percentages | 4 items with % bars |
| 29 | Predicted results | Chart + KPI numbers |

Run `python slidemaker.py inspect` for the full catalog with element IDs.
