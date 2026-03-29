/**
 * Slidemaker Apps Script Backend
 *
 * Deploy as web app: Deploy > New deployment > Web app
 *   - Execute as: Me
 *   - Who has access: Anyone with the link (or "Anyone in [org]")
 *
 * All requests are POST with JSON body: {action: "...", apiKey: "...", ...params}
 *
 * SECURITY:
 *   - API_KEY: shared secret, must match on every request
 *   - ALLOWED_TEMPLATES: only these template IDs can be copied
 *   - CREATED_BY: tracks presentation IDs created by this app (Script Properties)
 *   - Only created presentations or allowed templates can be read/edited
 *   - SHARE_WITH: email to auto-share created presentations with
 */

// --- CONFIGURATION ---
// Set these in Script Properties (Project Settings > Script Properties):
//   API_KEY            - shared secret (required)
//   ALLOWED_TEMPLATES  - comma-separated list of template presentation IDs
//   SHARE_WITH         - email address to share created presentations with (optional)
//
// No hardcoded IDs — all template IDs come from ALLOWED_TEMPLATES.

function getConfig() {
  const props = PropertiesService.getScriptProperties();
  var config = {
    apiKey: props.getProperty('API_KEY') || '',
    allowedTemplates: (props.getProperty('ALLOWED_TEMPLATES') || '').split(',').map(function(s) { return s.trim(); }).filter(Boolean),
    shareWith: props.getProperty('SHARE_WITH') || '',
  };
  if (config.allowedTemplates.length === 0) {
    throw new Error('ALLOWED_TEMPLATES not configured in Script Properties');
  }
  return config;
}

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    const config = getConfig();

    // --- API KEY CHECK ---
    if (config.apiKey && req.apiKey !== config.apiKey) {
      return jsonResponse({error: 'Invalid API key'});
    }

    // --- ACTION ROUTING ---
    let result;

    switch (req.action) {
      case 'inspect':
        var inspectId = req.presentationId || config.allowedTemplates[0];
        assertAllowed(inspectId, config);
        result = doInspect(inspectId);
        break;
      case 'create':
        result = doCreate(req, config);
        break;
      case 'get':
        assertAllowed(req.presentationId, config);
        result = doGet(req.presentationId);
        break;
      case 'edit':
        assertAllowed(req.presentationId, config);
        result = doEdit(req.presentationId, req.requests);
        break;
      case 'thumbnail':
        assertAllowed(req.presentationId, config);
        result = doThumbnail(req.presentationId, req.pageObjectId);
        break;
      default:
        return jsonResponse({error: 'Unknown action: ' + req.action});
    }

    return jsonResponse(result);
  } catch (err) {
    return jsonResponse({error: err.message, stack: err.stack});
  }
}

// --- ACCESS CONTROL ---

function getCreatedIds() {
  var raw = PropertiesService.getScriptProperties().getProperty('CREATED_IDS') || '[]';
  return JSON.parse(raw);
}

function addCreatedId(presId) {
  var ids = getCreatedIds();
  ids.push(presId);
  // Keep last 500 to avoid hitting property size limits
  if (ids.length > 500) ids = ids.slice(-500);
  PropertiesService.getScriptProperties().setProperty('CREATED_IDS', JSON.stringify(ids));
}

function assertAllowed(presentationId, config) {
  if (!presentationId) throw new Error('No presentation ID provided');
  // Allow access to registered templates
  if (config.allowedTemplates.indexOf(presentationId) !== -1) return;
  // Allow access to presentations created by this app
  if (getCreatedIds().indexOf(presentationId) !== -1) return;
  throw new Error('Access denied: presentation ' + presentationId + ' is not an allowed template or created by this app');
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

function doCreate(req, config) {
  const title = req.title || 'Untitled Presentation';
  const keepIndices = new Set(req.keep_slides || []);
  const replacements = req.replacements || {};

  // 1. Copy template (must be in allowed list)
  const templateId = req.templateId || config.allowedTemplates[0];
  if (config.allowedTemplates.indexOf(templateId) === -1) {
    throw new Error('Template ' + templateId + ' is not in the allowed list');
  }
  const copy = DriveApp.getFileById(templateId).makeCopy(title);
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

  // 6. Track this presentation as created by the app
  addCreatedId(presId);

  // 7. Auto-share with configured email
  if (config.shareWith) {
    copy.addEditor(config.shareWith);
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

    } else if (op.deleteElement) {
      apiRequests.push({deleteObject: {objectId: op.deleteElement.objectId}});

    } else if (op.duplicateSlide) {
      apiRequests.push({duplicateObject: {objectId: op.duplicateSlide.objectId}});

    } else if (op.moveSlide) {
      apiRequests.push({updateSlidesPosition: {slideObjectIds: [op.moveSlide.objectId], insertionIndex: op.moveSlide.insertionIndex}});

    } else if (op.moveElement) {
      var m = op.moveElement;
      apiRequests.push({updatePageElementTransform: {
        objectId: m.objectId,
        transform: {scaleX: m.scaleX || 1, scaleY: m.scaleY || 1, shearX: 0, shearY: 0, unit: 'PT', translateX: m.x || 0, translateY: m.y || 0},
        applyMode: 'ABSOLUTE'
      }});

    } else if (op.resizeElement) {
      var r = op.resizeElement;
      apiRequests.push({updatePageElementTransform: {
        objectId: r.objectId,
        transform: {scaleX: r.scaleX || 1, scaleY: r.scaleY || 1, shearX: 0, shearY: 0, unit: 'PT', translateX: 0, translateY: 0},
        applyMode: 'RELATIVE'
      }});

    } else if (op.textStyle) {
      var s = op.textStyle;
      var style = {};
      var fields = [];
      if (s.fontSize) { style.fontSize = {magnitude: s.fontSize, unit: 'PT'}; fields.push('fontSize'); }
      if (s.bold !== undefined) { style.bold = s.bold; fields.push('bold'); }
      if (s.italic !== undefined) { style.italic = s.italic; fields.push('italic'); }
      if (s.fontFamily) { style.fontFamily = s.fontFamily; fields.push('fontFamily'); }
      if (s.color) {
        var c = s.color;
        if (typeof c === 'string' && c.charAt(0) === '#') {
          c = {red: parseInt(c.substr(1,2),16)/255, green: parseInt(c.substr(3,2),16)/255, blue: parseInt(c.substr(5,2),16)/255};
        }
        style.foregroundColor = {opaqueColor: {rgbColor: c}};
        fields.push('foregroundColor');
      }
      if (fields.length > 0) {
        apiRequests.push({updateTextStyle: {objectId: s.objectId, textRange: s.textRange || {type: 'ALL'}, style: style, fields: fields.join(',')}});
      }

    } else if (op.paragraphStyle) {
      var p = op.paragraphStyle;
      var pStyle = {};
      var pFields = [];
      if (p.alignment) { pStyle.alignment = p.alignment; pFields.push('alignment'); }
      if (p.lineSpacing) { pStyle.lineSpacing = p.lineSpacing; pFields.push('lineSpacing'); }
      if (p.spaceAbove !== undefined) { pStyle.spaceAbove = {magnitude: p.spaceAbove, unit: 'PT'}; pFields.push('spaceAbove'); }
      if (p.spaceBelow !== undefined) { pStyle.spaceBelow = {magnitude: p.spaceBelow, unit: 'PT'}; pFields.push('spaceBelow'); }
      if (pFields.length > 0) {
        apiRequests.push({updateParagraphStyle: {objectId: p.objectId, textRange: p.textRange || {type: 'ALL'}, style: pStyle, fields: pFields.join(',')}});
      }

    } else if (op.shapeFill) {
      var f = op.shapeFill;
      var fc = f.color;
      if (typeof fc === 'string' && fc.charAt(0) === '#') {
        fc = {red: parseInt(fc.substr(1,2),16)/255, green: parseInt(fc.substr(3,2),16)/255, blue: parseInt(fc.substr(5,2),16)/255};
      }
      apiRequests.push({updateShapeProperties: {
        objectId: f.objectId,
        shapeProperties: {shapeBackgroundFill: {solidFill: {color: {rgbColor: fc}}}},
        fields: 'shapeBackgroundFill.solidFill.color'
      }});

    } else if (op.addImage) {
      var img = op.addImage;
      var props = {pageObjectId: img.pageId};
      if (img.size) { props.size = {width: {magnitude: img.size.width, unit: 'PT'}, height: {magnitude: img.size.height, unit: 'PT'}}; }
      if (img.position) { props.transform = {scaleX:1, scaleY:1, shearX:0, shearY:0, translateX: img.position.x, translateY: img.position.y, unit: 'PT'}; }
      apiRequests.push({createImage: {url: img.url, elementProperties: props}});

    } else if (op.insertText) {
      apiRequests.push({insertText: {objectId: op.insertText.objectId, text: op.insertText.text, insertionIndex: op.insertText.insertionIndex || 0}});

    } else if (op.replaceAllText) {
      apiRequests.push({replaceAllText: {
        containsText: {text: op.replaceAllText.find, matchCase: op.replaceAllText.matchCase !== false},
        replaceText: op.replaceAllText.replace
      }});

    } else if (op.createShape) {
      var cs = op.createShape;
      var csReq = {createShape: {
        shapeType: cs.shapeType || 'TEXT_BOX',
        elementProperties: {
          pageObjectId: cs.pageId,
          size: {width: {magnitude: cs.width || 200, unit: 'PT'}, height: {magnitude: cs.height || 50, unit: 'PT'}},
          transform: {scaleX:1, scaleY:1, shearX:0, shearY:0, translateX: cs.x || 100, translateY: cs.y || 100, unit: 'PT'}
        }
      }};
      if (cs.objectId) csReq.createShape.objectId = cs.objectId;
      apiRequests.push(csReq);

    } else if (op.createLine) {
      var cl = op.createLine;
      apiRequests.push({createLine: {
        lineCategory: cl.category || 'STRAIGHT',
        elementProperties: {
          pageObjectId: cl.pageId,
          size: {width: {magnitude: cl.width || 200, unit: 'PT'}, height: {magnitude: cl.height || 0, unit: 'PT'}},
          transform: {scaleX:1, scaleY:1, shearX:0, shearY:0, translateX: cl.x || 100, translateY: cl.y || 100, unit: 'PT'}
        }
      }});

    } else if (op.createTable) {
      var ct = op.createTable;
      var ctReq = {createTable: {
        rows: ct.rows || 3, columns: ct.columns || 3,
        elementProperties: {
          pageObjectId: ct.pageId,
          size: {width: {magnitude: ct.width || 500, unit: 'PT'}, height: {magnitude: ct.height || 200, unit: 'PT'}},
          transform: {scaleX:1, scaleY:1, shearX:0, shearY:0, translateX: ct.x || 50, translateY: ct.y || 100, unit: 'PT'}
        }
      }};
      if (ct.objectId) ctReq.createTable.objectId = ct.objectId;
      apiRequests.push(ctReq);

    } else if (op.insertTableText) {
      var tt = op.insertTableText;
      apiRequests.push({insertText: {
        objectId: tt.tableId,
        cellLocation: {rowIndex: tt.row, columnIndex: tt.column},
        text: tt.text,
        insertionIndex: 0
      }});

    } else if (op.slideBackground) {
      var bg = op.slideBackground;
      var bgc = bg.color;
      if (typeof bgc === 'string' && bgc.charAt(0) === '#') {
        bgc = {red: parseInt(bgc.substr(1,2),16)/255, green: parseInt(bgc.substr(3,2),16)/255, blue: parseInt(bgc.substr(5,2),16)/255};
      }
      apiRequests.push({updatePageProperties: {
        objectId: bg.pageId,
        pageProperties: {pageBackgroundFill: {solidFill: {color: {rgbColor: bgc}}}},
        fields: 'pageBackgroundFill.solidFill.color'
      }});

    } else if (op.groupObjects) {
      apiRequests.push({groupObjects: {childrenObjectIds: op.groupObjects.objectIds}});

    } else if (op.ungroupObjects) {
      apiRequests.push({ungroupObjects: {objectIds: [op.ungroupObjects.objectId]}});

    } else if (op.updateImageProperties) {
      var ip = op.updateImageProperties;
      var iProps = {};
      var iFields = [];
      if (ip.transparency !== undefined) { iProps.transparency = ip.transparency; iFields.push('transparency'); }
      if (ip.brightness !== undefined) { iProps.brightness = ip.brightness; iFields.push('brightness'); }
      if (ip.contrast !== undefined) { iProps.contrast = ip.contrast; iFields.push('contrast'); }
      if (ip.recolor) { iProps.recolor = {recolorStops: ip.recolor}; iFields.push('recolor'); }
      if (iFields.length > 0) {
        apiRequests.push({updateImageProperties: {objectId: ip.objectId, imageProperties: iProps, fields: iFields.join(',')}});
      }

    } else if (op.shapeOutline) {
      var ol = op.shapeOutline;
      var outline = {};
      var olFields = [];
      if (ol.color) {
        var olc = ol.color;
        if (typeof olc === 'string' && olc.charAt(0) === '#') {
          olc = {red: parseInt(olc.substr(1,2),16)/255, green: parseInt(olc.substr(3,2),16)/255, blue: parseInt(olc.substr(5,2),16)/255};
        }
        outline.outlineFill = {solidFill: {color: {rgbColor: olc}}};
        olFields.push('outline.outlineFill.solidFill.color');
      }
      if (ol.weight !== undefined) { outline.weight = {magnitude: ol.weight, unit: 'PT'}; olFields.push('outline.weight'); }
      if (ol.dashStyle) { outline.dashStyle = ol.dashStyle; olFields.push('outline.dashStyle'); }
      if (olFields.length > 0) {
        apiRequests.push({updateShapeProperties: {objectId: ol.objectId, shapeProperties: {outline: outline}, fields: olFields.join(',')}});
      }

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

// --- THUMBNAIL ---

function doThumbnail(presentationId, pageObjectId) {
  var thumb = Slides.Presentations.Pages.getThumbnail(presentationId, pageObjectId, {
    'thumbnailProperties.thumbnailSize': 'LARGE'
  });
  return {contentUrl: thumb.contentUrl};
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
