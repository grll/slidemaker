#!/usr/bin/env python3
"""Slidemaker - Create and edit Google Slides from a template via the API.

Supports two backends:
  1. Direct Google API (requires credentials.json + OAuth)
  2. Apps Script web app (requires WEBAPP_URL in .env or --webapp flag)

Templates are registered in templates/<name>/ with catalog.json and thumbnails.
"""

import argparse
import json
import os
import sys
import urllib.request
import urllib.error

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")
ENV_FILE = os.path.join(BASE_DIR, ".env")
CREDENTIALS_FILE = os.path.join(BASE_DIR, "credentials.json")
TOKEN_FILE = os.path.join(BASE_DIR, "token.json")


# --- BACKEND SELECTION ---

def read_env():
    """Read all key=value pairs from .env file."""
    env = {}
    if os.path.exists(ENV_FILE):
        with open(ENV_FILE) as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#") and "=" in line:
                    k, v = line.split("=", 1)
                    env[k.strip()] = v.strip().strip("'\"")
    return env


def get_webapp_url():
    return read_env().get("WEBAPP_URL")


def get_api_key():
    return read_env().get("API_KEY")


def use_webapp():
    return get_webapp_url() is not None


# --- APPS SCRIPT BACKEND ---

def webapp_request(action, **kwargs):
    url = get_webapp_url()
    if not url:
        print("ERROR: WEBAPP_URL not set in .env", file=sys.stderr)
        sys.exit(1)
    api_key = get_api_key()
    payload = {"action": action, **kwargs}
    if api_key:
        payload["apiKey"] = api_key
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


def emu_to_pt(emu):
    """Convert EMU (English Metric Units) to points."""
    return round(emu / 12700, 1)


def collect_text_elements(elements, result=None, include_transform=False):
    if result is None:
        result = []
    for elem in elements:
        if "elementGroup" in elem:
            collect_text_elements(elem["elementGroup"].get("children", []), result, include_transform)
        elif "shape" in elem and "text" in elem.get("shape", {}):
            text = extract_text(elem["shape"]["text"]).strip()
            if text:
                entry = {
                    "objectId": elem["objectId"],
                    "text": text,
                    "shapeType": elem["shape"].get("shapeType", ""),
                }
                if include_transform:
                    t = elem.get("transform", {})
                    s = elem.get("size", {})
                    entry["transform"] = {
                        "x": emu_to_pt(t.get("translateX", 0)),
                        "y": emu_to_pt(t.get("translateY", 0)),
                        "scaleX": t.get("scaleX", 1),
                        "scaleY": t.get("scaleY", 1),
                    }
                    entry["size"] = {
                        "width": emu_to_pt(s.get("width", {}).get("magnitude", 0)),
                        "height": emu_to_pt(s.get("height", {}).get("magnitude", 0)),
                    }
                    # Extract font size from first text run
                    for te in elem["shape"]["text"].get("textElements", []):
                        if "textRun" in te:
                            ts = te["textRun"].get("style", {})
                            fs = ts.get("fontSize", {})
                            if fs:
                                entry["fontSize"] = fs.get("magnitude", 0)
                                entry["fontUnit"] = fs.get("unit", "PT")
                            break
                result.append(entry)
    return result


# --- TEMPLATE MANAGEMENT ---

def get_template_dir(name):
    return os.path.join(TEMPLATES_DIR, name)


def list_templates():
    if not os.path.exists(TEMPLATES_DIR):
        return []
    result = []
    for name in sorted(os.listdir(TEMPLATES_DIR)):
        config_path = os.path.join(TEMPLATES_DIR, name, "config.json")
        if os.path.exists(config_path):
            with open(config_path) as f:
                config = json.load(f)
            result.append({"name": name, **config})
    return result


def resolve_template(name):
    """Resolve template name to its config. Returns (template_dir, config)."""
    tdir = get_template_dir(name)
    config_path = os.path.join(tdir, "config.json")
    if not os.path.exists(config_path):
        print(f"ERROR: Template '{name}' not found. Run: slidemaker register <presentation_id> {name}", file=sys.stderr)
        sys.exit(1)
    with open(config_path) as f:
        return tdir, json.load(f)


def download_thumbnail(slides_svc, presentation_id, slide_object_id, output_path):
    """Download a slide thumbnail as PNG."""
    thumb = slides_svc.presentations().pages().getThumbnail(
        presentationId=presentation_id,
        pageObjectId=slide_object_id,
        thumbnailProperties_thumbnailSize="LARGE",
    ).execute()
    thumb_url = thumb.get("contentUrl")
    if thumb_url:
        urllib.request.urlretrieve(thumb_url, output_path)
        return True
    return False


def download_thumbnail_webapp(presentation_id, slide_object_id, output_path):
    """Download a slide thumbnail via Apps Script backend."""
    result = webapp_request("thumbnail", presentationId=presentation_id, pageObjectId=slide_object_id)
    if "error" in result:
        print(f"WARNING: Could not get thumbnail: {result['error']}", file=sys.stderr)
        return False
    thumb_url = result.get("contentUrl")
    if thumb_url:
        urllib.request.urlretrieve(thumb_url, output_path)
        return True
    return False


# --- COMMANDS ---

def cmd_auth(_args):
    if use_webapp():
        result = webapp_request("inspect", presentationId="dummy_test")
        print("Apps Script backend connected.")
    else:
        creds = get_creds()
        from googleapiclient.discovery import build
        drive = build("drive", "v3", credentials=creds)
        about = drive.about().get(fields="user").execute()
        print(f"Authenticated as: {about['user']['emailAddress']}")


def cmd_register(args):
    """Register a template: download catalog + thumbnails."""
    pres_id = args.presentation_id
    name = args.name

    tdir = get_template_dir(name)
    thumbs_dir = os.path.join(tdir, "thumbnails")
    os.makedirs(thumbs_dir, exist_ok=True)

    # Get presentation data
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

    # Save config
    config = {"presentationId": pres_id, "title": result.get("title", "")}
    with open(os.path.join(tdir, "config.json"), "w") as f:
        json.dump(config, f, indent=2)

    # Save catalog
    with open(os.path.join(tdir, "catalog.json"), "w") as f:
        json.dump(result, f, indent=2)

    # Download thumbnails
    print(f"Downloading {len(result['slides'])} thumbnails...")
    for slide in result["slides"]:
        idx = slide["index"]
        out_path = os.path.join(thumbs_dir, f"slide_{idx:02d}.png")
        if use_webapp():
            ok = download_thumbnail_webapp(pres_id, slide["objectId"], out_path)
        else:
            ok = download_thumbnail(slides_svc, pres_id, slide["objectId"], out_path)
        status = "ok" if ok else "FAILED"
        print(f"  slide {idx:2d}: {status}")

    print(f"\nTemplate '{name}' registered at {tdir}")


def cmd_templates(_args):
    """List registered templates."""
    templates = list_templates()
    if not templates:
        print("No templates registered. Run: slidemaker register <presentation_id> <name>")
        return
    for t in templates:
        tdir = get_template_dir(t["name"])
        thumbs = os.path.join(tdir, "thumbnails")
        n_thumbs = len([f for f in os.listdir(thumbs) if f.endswith(".png")]) if os.path.exists(thumbs) else 0
        print(f"  {t['name']}: {t.get('title', '?')} ({n_thumbs} thumbnails)")


def cmd_inspect(args):
    """Show catalog for a template or arbitrary presentation."""
    if args.template:
        tdir, config = resolve_template(args.template)
        cat_path = os.path.join(tdir, "catalog.json")
        with open(cat_path) as f:
            result = json.load(f)
    elif args.presentation_id:
        pres_id = args.presentation_id
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
    else:
        # Default: show first registered template
        templates = list_templates()
        if templates:
            tdir = get_template_dir(templates[0]["name"])
            with open(os.path.join(tdir, "catalog.json")) as f:
                result = json.load(f)
        else:
            print("No template specified and none registered.", file=sys.stderr)
            sys.exit(1)
    print(json.dumps(result, indent=2))


def cmd_create(args):
    content = json.loads(args.content)

    # Resolve template
    template_name = content.pop("template", args.template)
    if template_name:
        _, config = resolve_template(template_name)
        template_id = config["presentationId"]
    else:
        # Fall back to first registered template
        templates = list_templates()
        if templates:
            template_id = templates[0]["presentationId"]
            template_name = templates[0]["name"]
        else:
            print("ERROR: No template specified and none registered.", file=sys.stderr)
            sys.exit(1)

    if use_webapp():
        content["templateId"] = template_id
        result = webapp_request("create", **content)
    else:
        slides_svc, drive_svc = get_services()
        title = content.get("title", "Untitled Presentation")
        keep_indices = set(content.get("keep_slides", []))
        replacements = content.get("replacements", {})

        copy = drive_svc.files().copy(fileId=template_id, body={"name": title}).execute()
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
    detailed = getattr(args, 'detailed', False)
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
            text_elements = collect_text_elements(slide.get("pageElements", []), include_transform=detailed)
            result["slides"].append({
                "index": i, "objectId": slide["objectId"], "elements": text_elements,
            })
    print(json.dumps(result, indent=2))


def cmd_thumbnails(args):
    """Download thumbnails for a presentation (for reviewing created decks)."""
    pres_id = args.presentation_id
    out_dir = args.output or os.path.join(BASE_DIR, "review")
    os.makedirs(out_dir, exist_ok=True)

    # Get slide IDs
    if use_webapp():
        pres_data = webapp_request("get", presentationId=pres_id)
    else:
        slides_svc, _ = get_services()
        pres = slides_svc.presentations().get(presentationId=pres_id).execute()
        pres_data = {
            "slides": [{"index": i, "objectId": s["objectId"]} for i, s in enumerate(pres.get("slides", []))],
        }

    print(f"Downloading {len(pres_data['slides'])} thumbnails to {out_dir}/...")
    paths = []
    for slide in pres_data["slides"]:
        idx = slide["index"]
        out_path = os.path.join(out_dir, f"slide_{idx:02d}.png")
        if use_webapp():
            ok = download_thumbnail_webapp(pres_id, slide["objectId"], out_path)
        else:
            ok = download_thumbnail(slides_svc, pres_id, slide["objectId"], out_path)
        if ok:
            paths.append(out_path)
            print(f"  slide {idx}: {out_path}")
        else:
            print(f"  slide {idx}: FAILED")

    print(json.dumps({"paths": paths}))


def build_edit_requests(ops):
    """Convert high-level edit operations to Slides API batchUpdate requests."""
    api_requests = []
    for op in ops:
        if "replaceText" in op:
            r = op["replaceText"]
            api_requests.append({"deleteText": {"objectId": r["objectId"], "textRange": {"type": "ALL"}}})
            api_requests.append({"insertText": {"objectId": r["objectId"], "text": r["text"], "insertionIndex": 0}})

        elif "deleteSlide" in op:
            api_requests.append({"deleteObject": {"objectId": op["deleteSlide"]["objectId"]}})

        elif "deleteElement" in op:
            api_requests.append({"deleteObject": {"objectId": op["deleteElement"]["objectId"]}})

        elif "duplicateSlide" in op:
            api_requests.append({"duplicateObject": {"objectId": op["duplicateSlide"]["objectId"]}})

        elif "moveSlide" in op:
            m = op["moveSlide"]
            api_requests.append({"updateSlidesPosition": {"slideObjectIds": [m["objectId"]], "insertionIndex": m["insertionIndex"]}})

        elif "moveElement" in op:
            # Move an element to an absolute position (in points)
            m = op["moveElement"]
            transform = {"scaleX": m.get("scaleX", 1), "scaleY": m.get("scaleY", 1),
                         "shearX": 0, "shearY": 0, "unit": "PT",
                         "translateX": m.get("x", 0), "translateY": m.get("y", 0)}
            api_requests.append({
                "updatePageElementTransform": {
                    "objectId": m["objectId"],
                    "transform": transform,
                    "applyMode": "ABSOLUTE",
                }
            })

        elif "resizeElement" in op:
            # Scale an element relative to current size
            r = op["resizeElement"]
            api_requests.append({
                "updatePageElementTransform": {
                    "objectId": r["objectId"],
                    "transform": {
                        "scaleX": r.get("scaleX", 1), "scaleY": r.get("scaleY", 1),
                        "shearX": 0, "shearY": 0, "unit": "PT",
                        "translateX": 0, "translateY": 0,
                    },
                    "applyMode": "RELATIVE",
                }
            })

        elif "textStyle" in op:
            # Change text formatting: fontSize, bold, italic, foregroundColor
            s = op["textStyle"]
            style = {}
            fields = []
            if "fontSize" in s:
                style["fontSize"] = {"magnitude": s["fontSize"], "unit": "PT"}
                fields.append("fontSize")
            if "bold" in s:
                style["bold"] = s["bold"]
                fields.append("bold")
            if "italic" in s:
                style["italic"] = s["italic"]
                fields.append("italic")
            if "color" in s:
                # Accept hex like "#FF0000" or rgb dict
                c = s["color"]
                if isinstance(c, str) and c.startswith("#"):
                    r_val = int(c[1:3], 16) / 255
                    g_val = int(c[3:5], 16) / 255
                    b_val = int(c[5:7], 16) / 255
                    c = {"red": r_val, "green": g_val, "blue": b_val}
                style["foregroundColor"] = {"opaqueColor": {"rgbColor": c}}
                fields.append("foregroundColor")
            if "fontFamily" in s:
                style["fontFamily"] = s["fontFamily"]
                fields.append("fontFamily")
            if fields:
                api_requests.append({
                    "updateTextStyle": {
                        "objectId": s["objectId"],
                        "textRange": s.get("textRange", {"type": "ALL"}),
                        "style": style,
                        "fields": ",".join(fields),
                    }
                })

        elif "paragraphStyle" in op:
            # Change paragraph alignment etc.
            p = op["paragraphStyle"]
            style = {}
            fields = []
            if "alignment" in p:
                style["alignment"] = p["alignment"]  # START, CENTER, END, JUSTIFIED
                fields.append("alignment")
            if "lineSpacing" in p:
                style["lineSpacing"] = p["lineSpacing"]  # percentage, e.g. 115
                fields.append("lineSpacing")
            if "spaceAbove" in p:
                style["spaceAbove"] = {"magnitude": p["spaceAbove"], "unit": "PT"}
                fields.append("spaceAbove")
            if "spaceBelow" in p:
                style["spaceBelow"] = {"magnitude": p["spaceBelow"], "unit": "PT"}
                fields.append("spaceBelow")
            if fields:
                api_requests.append({
                    "updateParagraphStyle": {
                        "objectId": p["objectId"],
                        "textRange": p.get("textRange", {"type": "ALL"}),
                        "style": style,
                        "fields": ",".join(fields),
                    }
                })

        elif "shapeFill" in op:
            # Change shape background color
            f = op["shapeFill"]
            c = f["color"]
            if isinstance(c, str) and c.startswith("#"):
                r_val = int(c[1:3], 16) / 255
                g_val = int(c[3:5], 16) / 255
                b_val = int(c[5:7], 16) / 255
                c = {"red": r_val, "green": g_val, "blue": b_val}
            api_requests.append({
                "updateShapeProperties": {
                    "objectId": f["objectId"],
                    "shapeProperties": {
                        "shapeBackgroundFill": {
                            "solidFill": {"color": {"rgbColor": c}}
                        }
                    },
                    "fields": "shapeBackgroundFill.solidFill.color",
                }
            })

        elif "addImage" in op:
            # Add an image from URL
            img = op["addImage"]
            req = {
                "createImage": {
                    "url": img["url"],
                    "elementProperties": {
                        "pageObjectId": img["pageId"],
                    }
                }
            }
            if "size" in img:
                req["createImage"]["elementProperties"]["size"] = {
                    "width": {"magnitude": img["size"]["width"], "unit": "PT"},
                    "height": {"magnitude": img["size"]["height"], "unit": "PT"},
                }
            if "position" in img:
                req["createImage"]["elementProperties"]["transform"] = {
                    "scaleX": 1, "scaleY": 1, "shearX": 0, "shearY": 0,
                    "translateX": img["position"]["x"], "translateY": img["position"]["y"],
                    "unit": "PT",
                }
            api_requests.append(req)

        elif "raw" in op:
            api_requests.append(op["raw"])

    return api_requests


def cmd_edit(args):
    ops = json.loads(args.requests)

    if use_webapp():
        result = webapp_request("edit", presentationId=args.presentation_id, requests=ops)
    else:
        slides_svc, _ = get_services()
        api_requests = build_edit_requests(ops)

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

    p_reg = sub.add_parser("register", help="Register a template (download catalog + thumbnails)")
    p_reg.add_argument("presentation_id", help="Google Slides presentation ID")
    p_reg.add_argument("name", help="Template name (used as directory name)")

    sub.add_parser("templates", help="List registered templates")

    p_inspect = sub.add_parser("inspect", help="Inspect template or presentation")
    p_inspect.add_argument("presentation_id", nargs="?", default=None)
    p_inspect.add_argument("--template", "-t", help="Template name")

    p_create = sub.add_parser("create", help="Create presentation from template")
    p_create.add_argument("content", help="JSON: {title, keep_slides, replacements, template?}")
    p_create.add_argument("--template", "-t", help="Template name")

    p_get = sub.add_parser("get", help="Get presentation content")
    p_get.add_argument("presentation_id")
    p_get.add_argument("--detailed", "-d", action="store_true", help="Include position, size, and font info")

    p_edit = sub.add_parser("edit", help="Edit a presentation")
    p_edit.add_argument("presentation_id")
    p_edit.add_argument("requests", help="JSON list of edit operations")

    p_thumb = sub.add_parser("thumbnails", help="Download slide thumbnails for review")
    p_thumb.add_argument("presentation_id")
    p_thumb.add_argument("--output", "-o", help="Output directory (default: review/)")

    args = parser.parse_args()
    if not args.command:
        parser.print_help()
        sys.exit(1)

    cmds = {
        "auth": cmd_auth, "register": cmd_register, "templates": cmd_templates,
        "inspect": cmd_inspect, "create": cmd_create, "get": cmd_get,
        "edit": cmd_edit, "thumbnails": cmd_thumbnails,
    }
    cmds[args.command](args)


if __name__ == "__main__":
    main()
