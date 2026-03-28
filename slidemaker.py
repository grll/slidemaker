#!/usr/bin/env python3
"""Slidemaker - Create and edit Google Slides from a template via the API.

Supports two backends:
  1. Direct Google API (requires credentials.json + OAuth)
  2. Apps Script web app (requires WEBAPP_URL in .env or --webapp flag)
"""

import argparse
import json
import os
import sys
import urllib.request
import urllib.error

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_ID = "1cWqfy4vpwbmlgPaN09ha02QILAj3phf_akY9C4VqtVE"
ENV_FILE = os.path.join(BASE_DIR, ".env")
CREDENTIALS_FILE = os.path.join(BASE_DIR, "credentials.json")
TOKEN_FILE = os.path.join(BASE_DIR, "token.json")


# --- BACKEND SELECTION ---

def get_webapp_url():
    """Read WEBAPP_URL from .env file."""
    if os.path.exists(ENV_FILE):
        with open(ENV_FILE) as f:
            for line in f:
                line = line.strip()
                if line.startswith("WEBAPP_URL="):
                    return line.split("=", 1)[1].strip().strip("'\"")
    return None


def use_webapp():
    """Check if we should use the Apps Script backend."""
    return get_webapp_url() is not None


# --- APPS SCRIPT BACKEND ---

def webapp_request(action, **kwargs):
    """Send a request to the Apps Script web app."""
    url = get_webapp_url()
    if not url:
        print("ERROR: WEBAPP_URL not set in .env", file=sys.stderr)
        sys.exit(1)

    payload = {"action": action, **kwargs}
    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(url, data=data, headers={"Content-Type": "application/json"})

    try:
        with urllib.request.urlopen(req, timeout=120) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8") if e.fp else ""
        print(f"ERROR: HTTP {e.code}: {body}", file=sys.stderr)
        sys.exit(1)


# --- DIRECT API BACKEND ---

def get_creds():
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow

    SCOPES = [
        "https://www.googleapis.com/auth/presentations",
        "https://www.googleapis.com/auth/drive",
    ]
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


def get_services():
    from googleapiclient.discovery import build
    creds = get_creds()
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


# --- COMMANDS ---

def cmd_auth(_args):
    if use_webapp():
        result = webapp_request("inspect", presentationId=TEMPLATE_ID)
        if "error" in result:
            print(f"ERROR: {result['error']}", file=sys.stderr)
            sys.exit(1)
        print(f"Apps Script backend OK. Template: {result.get('title', '?')}")
    else:
        creds = get_creds()
        from googleapiclient.discovery import build
        drive = build("drive", "v3", credentials=creds)
        about = drive.about().get(fields="user").execute()
        print(f"Authenticated as: {about['user']['emailAddress']}")


def cmd_inspect(args):
    pres_id = args.presentation_id or TEMPLATE_ID
    if use_webapp():
        result = webapp_request("inspect", presentationId=pres_id)
    else:
        slides_svc, _ = get_services()
        pres = slides_svc.presentations().get(presentationId=pres_id).execute()
        result = {
            "title": pres.get("title", ""),
            "presentationId": pres.get("presentationId", ""),
            "slideCount": len(pres.get("slides", [])),
            "slides": [],
        }
        for i, slide in enumerate(pres.get("slides", [])):
            text_elements = collect_text_elements(slide.get("pageElements", []))
            result["slides"].append({
                "index": i, "objectId": slide["objectId"], "elements": text_elements,
            })
    print(json.dumps(result, indent=2))


def cmd_create(args):
    content = json.loads(args.content)

    if use_webapp():
        result = webapp_request("create", **content)
    else:
        slides_svc, drive_svc = get_services()
        title = content.get("title", "Untitled Presentation")
        keep_indices = set(content.get("keep_slides", []))
        replacements = content.get("replacements", {})

        copy = drive_svc.files().copy(fileId=TEMPLATE_ID, body={"name": title}).execute()
        pres_id = copy["id"]
        url = f"https://docs.google.com/presentation/d/{pres_id}/edit"

        pres = slides_svc.presentations().get(presentationId=pres_id).execute()
        all_slides = pres.get("slides", [])

        slides_to_delete = [s["objectId"] for i, s in enumerate(all_slides) if i not in keep_indices]
        if slides_to_delete and len(slides_to_delete) < len(all_slides):
            delete_reqs = [{"deleteObject": {"objectId": sid}} for sid in slides_to_delete]
            slides_svc.presentations().batchUpdate(presentationId=pres_id, body={"requests": delete_reqs}).execute()

        if "keep_slides" in content and len(content["keep_slides"]) > 1:
            pres = slides_svc.presentations().get(presentationId=pres_id).execute()
            current_ids = [s["objectId"] for s in pres.get("slides", [])]
            desired_ids = [all_slides[idx]["objectId"] for idx in content["keep_slides"] if all_slides[idx]["objectId"] in current_ids]
            move_reqs = [{"updateSlidesPosition": {"slideObjectIds": [sid], "insertionIndex": pos}} for pos, sid in enumerate(desired_ids)]
            if move_reqs:
                slides_svc.presentations().batchUpdate(presentationId=pres_id, body={"requests": move_reqs}).execute()

        if replacements:
            reqs = []
            for elem_id, new_text in replacements.items():
                reqs.append({"deleteText": {"objectId": elem_id, "textRange": {"type": "ALL"}}})
                reqs.append({"insertText": {"objectId": elem_id, "text": new_text, "insertionIndex": 0}})
            slides_svc.presentations().batchUpdate(presentationId=pres_id, body={"requests": reqs}).execute()

        result = {"presentationId": pres_id, "url": url}

    if "error" in result:
        print(f"ERROR: {result['error']}", file=sys.stderr)
        sys.exit(1)
    print(json.dumps(result))


def cmd_get(args):
    if use_webapp():
        result = webapp_request("get", presentationId=args.presentation_id)
    else:
        slides_svc, _ = get_services()
        pres = slides_svc.presentations().get(presentationId=args.presentation_id).execute()
        result = {
            "title": pres.get("title", ""),
            "presentationId": pres.get("presentationId", ""),
            "slides": [],
        }
        for i, slide in enumerate(pres.get("slides", [])):
            text_elements = collect_text_elements(slide.get("pageElements", []))
            result["slides"].append({
                "index": i, "objectId": slide["objectId"], "elements": text_elements,
            })
    print(json.dumps(result, indent=2))


def cmd_edit(args):
    ops = json.loads(args.requests)

    if use_webapp():
        result = webapp_request("edit", presentationId=args.presentation_id, requests=ops)
    else:
        slides_svc, _ = get_services()
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
            res = slides_svc.presentations().batchUpdate(presentationId=args.presentation_id, body={"requests": api_requests}).execute()
            result = {"applied": len(api_requests), "replies": len(res.get("replies", []))}
        else:
            result = {"applied": 0}

    print(json.dumps(result))


# --- MAIN ---

def main():
    parser = argparse.ArgumentParser(description="Slidemaker - Google Slides automation")
    sub = parser.add_subparsers(dest="command")

    sub.add_parser("auth", help="Authenticate / test connection")

    p_inspect = sub.add_parser("inspect", help="Inspect template or presentation")
    p_inspect.add_argument("presentation_id", nargs="?", default=None)

    p_create = sub.add_parser("create", help="Create presentation from template")
    p_create.add_argument("content", help="JSON: {title, keep_slides, replacements}")

    p_get = sub.add_parser("get", help="Get presentation content")
    p_get.add_argument("presentation_id")

    p_edit = sub.add_parser("edit", help="Edit a presentation")
    p_edit.add_argument("presentation_id")
    p_edit.add_argument("requests", help="JSON list of edit operations")

    args = parser.parse_args()
    if not args.command:
        parser.print_help()
        sys.exit(1)

    cmds = {"auth": cmd_auth, "inspect": cmd_inspect, "create": cmd_create, "get": cmd_get, "edit": cmd_edit}
    cmds[args.command](args)


if __name__ == "__main__":
    main()
