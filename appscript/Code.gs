/**
 * Slidemaker Apps Script Backend
 *
 * Deploy as web app: Deploy > New deployment > Web app
 *   - Execute as: Me
 *   - Who has access: Anyone with the link (or "Anyone in [org]")
 *
 * All requests are POST with JSON body: {action: "...", ...params}
 */

const TEMPLATE_ID = '1cWqfy4vpwbmlgPaN09ha02QILAj3phf_akY9C4VqtVE';

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    let result;

    switch (req.action) {
      case 'inspect':
        result = doInspect(req.presentationId || TEMPLATE_ID);
        break;
      case 'create':
        result = doCreate(req);
        break;
      case 'get':
        result = doGet(req.presentationId);
        break;
      case 'edit':
        result = doEdit(req.presentationId, req.requests);
        break;
      default:
        return jsonResponse({error: 'Unknown action: ' + req.action});
    }

    return jsonResponse(result);
  } catch (err) {
    return jsonResponse({error: err.message, stack: err.stack});
  }
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// --- INSPECT ---

function doInspect(presentationId) {
  const pres = Slides.Presentations.get(presentationId);

  const catalog = {
    title: pres.title || '',
    presentationId: pres.presentationId,
    slideCount: pres.slides ? pres.slides.length : 0,
    slides: [],
  };

  (pres.slides || []).forEach(function(slide, i) {
    const elements = collectTextElements(slide.pageElements || []);
    catalog.slides.push({
      index: i,
      objectId: slide.objectId,
      elements: elements,
    });
  });

  return catalog;
}

// --- CREATE ---

function doCreate(req) {
  const title = req.title || 'Untitled Presentation';
  const keepIndices = new Set(req.keep_slides || []);
  const replacements = req.replacements || {};

  // 1. Copy template
  const copy = DriveApp.getFileById(TEMPLATE_ID).makeCopy(title);
  const presId = copy.getId();
  const url = 'https://docs.google.com/presentation/d/' + presId + '/edit';

  // 2. Get the copy to find slide IDs
  const pres = Slides.Presentations.get(presId);
  const allSlides = pres.slides || [];

  // 3. Delete unwanted slides
  const deleteReqs = [];
  allSlides.forEach(function(s, i) {
    if (!keepIndices.has(i)) {
      deleteReqs.push({deleteObject: {objectId: s.objectId}});
    }
  });
  if (deleteReqs.length > 0 && deleteReqs.length < allSlides.length) {
    Slides.Presentations.batchUpdate({requests: deleteReqs}, presId);
  }

  // 4. Reorder slides to match keep_slides order
  if (req.keep_slides && req.keep_slides.length > 1) {
    const presAfter = Slides.Presentations.get(presId);
    const currentIds = (presAfter.slides || []).map(function(s) { return s.objectId; });

    const desiredIds = [];
    req.keep_slides.forEach(function(idx) {
      const sid = allSlides[idx].objectId;
      if (currentIds.indexOf(sid) !== -1) {
        desiredIds.push(sid);
      }
    });

    const moveReqs = desiredIds.map(function(sid, pos) {
      return {updateSlidesPosition: {slideObjectIds: [sid], insertionIndex: pos}};
    });
    if (moveReqs.length > 0) {
      Slides.Presentations.batchUpdate({requests: moveReqs}, presId);
    }
  }

  // 5. Apply text replacements
  if (Object.keys(replacements).length > 0) {
    const textReqs = [];
    for (var elemId in replacements) {
      textReqs.push({deleteText: {objectId: elemId, textRange: {type: 'ALL'}}});
      textReqs.push({insertText: {objectId: elemId, text: replacements[elemId], insertionIndex: 0}});
    }
    Slides.Presentations.batchUpdate({requests: textReqs}, presId);
  }

  return {presentationId: presId, url: url};
}

// --- GET ---

function doGet(presentationId) {
  const pres = Slides.Presentations.get(presentationId);

  const output = {
    title: pres.title || '',
    presentationId: pres.presentationId,
    slides: [],
  };

  (pres.slides || []).forEach(function(slide, i) {
    const elements = collectTextElements(slide.pageElements || []);
    output.slides.push({
      index: i,
      objectId: slide.objectId,
      elements: elements,
    });
  });

  return output;
}

// --- EDIT ---

function doEdit(presentationId, ops) {
  const apiRequests = [];

  ops.forEach(function(op) {
    if (op.replaceText) {
      apiRequests.push({deleteText: {objectId: op.replaceText.objectId, textRange: {type: 'ALL'}}});
      apiRequests.push({insertText: {objectId: op.replaceText.objectId, text: op.replaceText.text, insertionIndex: 0}});
    } else if (op.deleteSlide) {
      apiRequests.push({deleteObject: {objectId: op.deleteSlide.objectId}});
    } else if (op.duplicateSlide) {
      apiRequests.push({duplicateObject: {objectId: op.duplicateSlide.objectId}});
    } else if (op.moveSlide) {
      apiRequests.push({updateSlidesPosition: {slideObjectIds: [op.moveSlide.objectId], insertionIndex: op.moveSlide.insertionIndex}});
    } else if (op.raw) {
      apiRequests.push(op.raw);
    }
  });

  if (apiRequests.length > 0) {
    const result = Slides.Presentations.batchUpdate({requests: apiRequests}, presentationId);
    return {applied: apiRequests.length, replies: (result.replies || []).length};
  }
  return {applied: 0};
}

// --- HELPERS ---

function collectTextElements(pageElements) {
  var result = [];
  (pageElements || []).forEach(function(elem) {
    if (elem.elementGroup) {
      result = result.concat(collectTextElements(elem.elementGroup.children || []));
    } else if (elem.shape && elem.shape.text) {
      var text = extractText(elem.shape.text).trim();
      if (text) {
        result.push({
          objectId: elem.objectId,
          text: text,
          shapeType: elem.shape.shapeType || '',
        });
      }
    }
  });
  return result;
}

function extractText(textObj) {
  var parts = [];
  (textObj.textElements || []).forEach(function(el) {
    if (el.textRun) {
      parts.push(el.textRun.content);
    }
  });
  return parts.join('');
}
