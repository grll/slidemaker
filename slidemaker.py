#!/usr/bin/env python3
"""Slidemaker - Create and edit Google Slides from a template via the API."""

import argparse
import json
import os
import sys

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

SCOPES = [
    "https://www.googleapis.com/auth/presentations",
    "https://www.googleapis.com/auth/drive",
]

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CREDENTIALS_FILE = os.path.join(BASE_DIR, "credentials.json")
TOKEN_FILE = os.path.join(BASE_DIR, "token.json")
TEMPLATE_ID = "1cWqfy4vpwbmlgPaN09ha02QILAj3phf_akY9C4VqtVE"


def get_creds():
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(CREDENTIALS_FILE):
                print(f"ERROR: {CREDENTIALS_FILE} not found.", file=sys.stderr)
                sys.exit(1)
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, "w") as f:
            f.write(creds.to_json())
    return creds


def get_services(creds):
    slides = build("slides", "v1", credentials=creds)
    drive = build("drive", "v3", credentials=creds)
    return slides, drive


def extract_text(text_obj):
    parts = []
    for el in text_obj.get("textElements", []):
        if "textRun" in el:
            parts.append(el["textRun"]["content"])
    return "".join(parts)


def collect_text_elements(elements, result=None):
    """Recursively collect all text-bearing elements from a slide, including inside groups."""
    if result is None:
        result = []
    for elem in elements:
        if "elementGroup" in elem:
            collect_text_elements(elem["elementGroup"].get("children", []), result)
        elif "shape" in elem and "text" in elem.get("shape", {}):
            text = extract_text(elem["shape"]["text"]).strip()
            if text:
                result.append({
                    "objectId": elem["objectId"],
                    "text": text,
                    "shapeType": elem["shape"].get("shapeType", ""),
                })
    return result


# --- AUTH ---

def cmd_auth(_args):
    creds = get_creds()
    _, drive = get_services(creds)
    about = drive.about().get(fields="user").execute()
    print(f"Authenticated as: {about['user']['emailAddress']}")


# --- INSPECT ---

def cmd_inspect(args):
    """Show all template slides with their text elements."""
    creds = get_creds()
    slides_svc, _ = get_services(creds)

    pres_id = args.presentation_id if hasattr(args, 'presentation_id') and args.presentation_id else TEMPLATE_ID
    pres = slides_svc.presentations().get(presentationId=pres_id).execute()

    catalog = {
        "title": pres.get("title", ""),
        "presentationId": pres.get("presentationId", ""),
        "slideCount": len(pres.get("slides", [])),
        "slides": [],
    }

    for i, slide in enumerate(pres.get("slides", [])):
        text_elements = collect_text_elements(slide.get("pageElements", []))
        catalog["slides"].append({
            "index": i,
            "objectId": slide["objectId"],
            "elements": text_elements,
        })

    print(json.dumps(catalog, indent=2))


# --- CREATE ---

def cmd_create(args):
    """Create a new presentation from the template.

    Input JSON:
    {
      "title": "My Presentation",
      "keep_slides": [0, 5, 13],        // template slide indices to keep (in order)
      "replacements": {                   // element objectId -> new text
        "g708a6ee8a1_0_59": "NEW TITLE",
        "g708a6ee8a1_0_60": "New subtitle"
      }
    }
    """
    creds = get_creds()
    slides_svc, drive_svc = get_services(creds)
    content = json.loads(args.content)
    title = content.get("title", "Untitled Presentation")
    keep_indices = set(content.get("keep_slides", []))
    replacements = content.get("replacements", {})

    # 1. Copy template
    copy = drive_svc.files().copy(fileId=TEMPLATE_ID, body={"name": title}).execute()
    pres_id = copy["id"]
    url = f"https://docs.google.com/presentation/d/{pres_id}/edit"

    # 2. Get the copy
    pres = slides_svc.presentations().get(presentationId=pres_id).execute()
    all_slides = pres.get("slides", [])

    # 3. Delete unwanted slides (must keep at least one)
    slides_to_delete = [s["objectId"] for i, s in enumerate(all_slides) if i not in keep_indices]
    if slides_to_delete and len(slides_to_delete) < len(all_slides):
        delete_reqs = [{"deleteObject": {"objectId": sid}} for sid in slides_to_delete]
        slides_svc.presentations().batchUpdate(
            presentationId=pres_id, body={"requests": delete_reqs}
        ).execute()

    # 4. Reorder slides to match the order in keep_slides
    if "keep_slides" in content and len(content["keep_slides"]) > 1:
        # Re-read after deletion
        pres = slides_svc.presentations().get(presentationId=pres_id).execute()
        current_slides = pres.get("slides", [])
        current_ids = [s["objectId"] for s in current_slides]

        # Build desired order from keep_slides
        desired_ids = []
        for idx in content["keep_slides"]:
            sid = all_slides[idx]["objectId"]
            if sid in current_ids:
                desired_ids.append(sid)

        # Move slides to correct positions
        move_reqs = []
        for target_pos, sid in enumerate(desired_ids):
            move_reqs.append({
                "updateSlidesPosition": {
                    "slideObjectIds": [sid],
                    "insertionIndex": target_pos,
                }
            })
        if move_reqs:
            slides_svc.presentations().batchUpdate(
                presentationId=pres_id, body={"requests": move_reqs}
            ).execute()

    # 5. Apply text replacements
    if replacements:
        reqs = []
        for elem_id, new_text in replacements.items():
            reqs.append({"deleteText": {"objectId": elem_id, "textRange": {"type": "ALL"}}})
            reqs.append({"insertText": {"objectId": elem_id, "text": new_text, "insertionIndex": 0}})
        slides_svc.presentations().batchUpdate(
            presentationId=pres_id, body={"requests": reqs}
        ).execute()

    print(json.dumps({"presentationId": pres_id, "url": url}))


# --- GET ---

def cmd_get(args):
    """Get the current content of a presentation as JSON."""
    creds = get_creds()
    slides_svc, _ = get_services(creds)
    pres = slides_svc.presentations().get(presentationId=args.presentation_id).execute()

    output = {
        "title": pres.get("title", ""),
        "presentationId": pres.get("presentationId", ""),
        "slides": [],
    }

    for i, slide in enumerate(pres.get("slides", [])):
        text_elements = collect_text_elements(slide.get("pageElements", []))
        output["slides"].append({
            "index": i,
            "objectId": slide["objectId"],
            "elements": text_elements,
        })

    print(json.dumps(output, indent=2))


# --- EDIT ---

def cmd_edit(args):
    """Edit a presentation.

    Input JSON (list of operations):
    [
      {"replaceText": {"objectId": "elem_id", "text": "new text"}},
      {"deleteSlide": {"objectId": "slide_id"}},
      {"duplicateSlide": {"objectId": "slide_id"}},
      {"moveSlide": {"objectId": "slide_id", "insertionIndex": 2}},
      {"raw": {<any Slides API batchUpdate request>}}
    ]
    """
    creds = get_creds()
    slides_svc, _ = get_services(creds)
    pres_id = args.presentation_id
    ops = json.loads(args.requests)

    pres = slides_svc.presentations().get(presentationId=pres_id).execute()
    api_requests = []

    for op in ops:
        if "replaceText" in op:
            r = op["replaceText"]
            api_requests.append({"deleteText": {"objectId": r["objectId"], "textRange": {"type": "ALL"}}})
            api_requests.append({"insertText": {"objectId": r["objectId"], "text": r["text"], "insertionIndex": 0}})
        elif "deleteSlide" in op:
            api_requests.append({"deleteObject": {"objectId": op["deleteSlide"]["objectId"]}})
        elif "duplicateSlide" in op:
            api_requests.append({"duplicateObject": {"objectId": op["duplicateSlide"]["objectId"]}})
        elif "moveSlide" in op:
            m = op["moveSlide"]
            api_requests.append({"updateSlidesPosition": {"slideObjectIds": [m["objectId"]], "insertionIndex": m["insertionIndex"]}})
        elif "raw" in op:
            api_requests.append(op["raw"])

    if api_requests:
        result = slides_svc.presentations().batchUpdate(
            presentationId=pres_id, body={"requests": api_requests}
        ).execute()
        print(json.dumps({"applied": len(api_requests), "replies": len(result.get("replies", []))}))
    else:
        print(json.dumps({"applied": 0}))


# --- MAIN ---

def main():
    parser = argparse.ArgumentParser(description="Slidemaker - Google Slides automation")
    sub = parser.add_subparsers(dest="command")

    sub.add_parser("auth", help="Authenticate with Google")

    p_inspect = sub.add_parser("inspect", help="Inspect template or presentation")
    p_inspect.add_argument("presentation_id", nargs="?", default=None, help="Presentation ID (default: template)")

    p_create = sub.add_parser("create", help="Create presentation from template")
    p_create.add_argument("content", help="JSON: {title, keep_slides, replacements}")

    p_get = sub.add_parser("get", help="Get presentation content")
    p_get.add_argument("presentation_id", help="Presentation ID")

    p_edit = sub.add_parser("edit", help="Edit a presentation")
    p_edit.add_argument("presentation_id", help="Presentation ID")
    p_edit.add_argument("requests", help="JSON list of edit operations")

    args = parser.parse_args()
    if not args.command:
        parser.print_help()
        sys.exit(1)

    cmds = {"auth": cmd_auth, "inspect": cmd_inspect, "create": cmd_create, "get": cmd_get, "edit": cmd_edit}
    cmds[args.command](args)


if __name__ == "__main__":
    main()
