import { createHash } from "node:crypto";

const digest = (value) => createHash("sha256").update(String(value ?? ""), "utf8").digest("hex");
const shortHash = (value) => digest(value).slice(0, 24);
const normalized = (value) => String(value ?? "").replace(/\s+/g, " ").trim();

function headingPaths(paragraphs) {
  const stack = [];
  return paragraphs.map((paragraph) => {
    const style = String(paragraph?.style || "");
    const match = style.match(/^Heading\s*([1-9])$/i);
    if (match && normalized(paragraph.text)) {
      const level = Number(match[1]);
      stack.length = level - 1;
      stack[level - 1] = normalized(paragraph.text);
    }
    return stack.filter(Boolean);
  });
}

function decorateWord(inspection) {
  const paragraphs = inspection.paragraphs || [];
  const paths = headingPaths(paragraphs);
  inspection.paragraphs = paragraphs.map((paragraph, index) => {
    const previous = paragraphs[index - 1];
    const next = paragraphs[index + 1];
    const text = normalized(paragraph.text);
    return {
      ...paragraph,
      anchor: {
        kind: "paragraph",
        part: paragraph.part || "word/document.xml",
        headingPath: paths[index],
        textHash: shortHash(text),
        beforeHash: previous ? shortHash(normalized(previous.text)) : null,
        afterHash: next ? shortHash(normalized(next.text)) : null,
        fingerprint: shortHash(JSON.stringify([paragraph.part || "word/document.xml", paths[index], text, normalized(previous?.text), normalized(next?.text)]))
      }
    };
  });
  inspection.tables = (inspection.tables || []).map((table, tableIndex) => {
    const headers = Array.isArray(table.rows?.[0]) ? table.rows[0].map(normalized) : [];
    return {
      ...table,
      anchor: {
        kind: "table",
        part: table.part || "word/document.xml",
        headers,
        fingerprint: shortHash(JSON.stringify([table.part || "word/document.xml", headers, table.rows?.length || 0, tableIndex]))
      }
    };
  });
  inspection.contentControls = (inspection.contentControls || []).map((control) => ({
    ...control,
    anchor: {
      kind: "contentControl",
      id: control.id ?? null,
      tag: control.tag ?? null,
      title: control.title ?? null,
      textHash: shortHash(normalized(control.text)),
      fingerprint: shortHash(JSON.stringify([control.id ?? null, control.tag ?? null, control.title ?? null, normalized(control.text)]))
    }
  }));
  return inspection;
}

function decorateExcel(inspection) {
  const sheetKey = Array.isArray(inspection.worksheets) ? "worksheets" : "sheets";
  inspection[sheetKey] = (inspection[sheetKey] || []).map((sheet) => ({
    ...sheet,
    anchor: {
      kind: "worksheet",
      name: sheet.name,
      relationshipId: sheet.relationshipId ?? null,
      part: sheet.part ?? null,
      fingerprint: shortHash(JSON.stringify([sheet.name, sheet.relationshipId ?? null, sheet.part ?? null]))
    },
    cells: (sheet.cells || []).map((cell) => ({
      ...cell,
      anchor: {
        kind: "range",
        sheet: sheet.name,
        address: cell.address,
        valueHash: shortHash(JSON.stringify(cell.value ?? cell.values ?? null)),
        formulaHash: shortHash(JSON.stringify(cell.formula ?? cell.formulas ?? null)),
        fingerprint: shortHash(JSON.stringify([sheet.relationshipId ?? sheet.name, cell.address, cell.value ?? cell.values ?? null, cell.formula ?? cell.formulas ?? null]))
      }
    }))
  }));
  inspection.tables = (inspection.tables || []).map((table) => ({
    ...table,
    anchor: {
      kind: "excelTable",
      name: table.name,
      displayName: table.displayName ?? null,
      sheet: table.sheet ?? table.worksheet ?? null,
      reference: table.reference ?? table.ref ?? null,
      fingerprint: shortHash(JSON.stringify([table.name, table.displayName ?? null, table.sheet ?? table.worksheet ?? null, table.reference ?? table.ref ?? null]))
    }
  }));
  return inspection;
}

function flattenShapes(shapes, output = []) {
  for (const shape of shapes || []) {
    output.push(shape);
    flattenShapes(shape.children, output);
  }
  return output;
}

function decoratePowerPoint(inspection) {
  inspection.slides = (inspection.slides || []).map((slide) => {
    const title = normalized((slide.shapes || []).find((shape) => /title/i.test(String(shape.placeholderType || shape.name || "")))?.text);
    const slideIdentity = slide.relationshipId ?? slide.part ?? slide.index;
    const decorateShape = (shape) => ({
      ...shape,
      anchor: {
        kind: "shape",
        slideId: slideIdentity,
        slideTitle: title || null,
        shapeId: shape.id ?? shape.shapeId ?? null,
        shapeName: shape.name ?? null,
        altText: shape.altText ?? shape.description ?? null,
        textHash: shortHash(normalized(shape.text)),
        fingerprint: shortHash(JSON.stringify([slideIdentity, shape.id ?? shape.shapeId ?? null, shape.name ?? null, shape.altText ?? shape.description ?? null, normalized(shape.text)]))
      },
      children: shape.children?.map(decorateShape)
    });
    return {
      ...slide,
      anchor: {
        kind: "slide",
        slideId: slideIdentity,
        title: title || null,
        fingerprint: shortHash(JSON.stringify([slideIdentity, title, (slide.shapes || []).length]))
      },
      shapes: (slide.shapes || []).map(decorateShape)
    };
  });
  return inspection;
}

export function addSemanticAnchors(kind, inspection) {
  if (!inspection || typeof inspection !== "object") return inspection;
  if (kind === "word") return decorateWord(inspection);
  if (kind === "excel") return decorateExcel(inspection);
  if (kind === "powerpoint") return decoratePowerPoint(inspection);
  return inspection;
}

function wordCandidates(inspection, anchor) {
  if (anchor.kind === "table") {
    return (inspection.tables || []).map((entry, index) => ({ entry, index })).filter(({ entry }) => {
      if (anchor.part && entry.anchor?.part !== anchor.part) return false;
      if (anchor.fingerprint && entry.anchor?.fingerprint === anchor.fingerprint) return true;
      return anchor.headers && JSON.stringify(entry.anchor?.headers || []) === JSON.stringify(anchor.headers);
    });
  }
  if (anchor.kind === "contentControl") {
    return (inspection.contentControls || []).map((entry, index) => ({ entry, index })).filter(({ entry }) =>
      (anchor.id != null && String(entry.id) === String(anchor.id))
      || (anchor.tag && entry.tag === anchor.tag)
      || (anchor.fingerprint && entry.anchor?.fingerprint === anchor.fingerprint));
  }
  const candidates = (inspection.paragraphs || []).map((entry, index) => ({ entry, index })).filter(({ entry }) => {
    if (anchor.part && entry.anchor?.part !== anchor.part) return false;
    if (anchor.fingerprint && entry.anchor?.fingerprint === anchor.fingerprint) return true;
    if (anchor.textHash && entry.anchor?.textHash !== anchor.textHash) return false;
    return Boolean(anchor.textHash || anchor.beforeHash || anchor.afterHash);
  });
  const exact = candidates.filter(({ entry }) => anchor.fingerprint && entry.anchor?.fingerprint === anchor.fingerprint);
  if (exact.length) return exact;
  if (candidates.length > 1 && (anchor.beforeHash || anchor.afterHash)) {
    const contextual = candidates.filter(({ entry }) => (!anchor.beforeHash || entry.anchor?.beforeHash === anchor.beforeHash) && (!anchor.afterHash || entry.anchor?.afterHash === anchor.afterHash));
    if (contextual.length) return contextual;
  }
  return candidates;
}

function excelCandidates(inspection, anchor) {
  if (anchor.kind === "excelTable") {
    return (inspection.tables || []).map((entry, index) => ({ entry, index })).filter(({ entry }) =>
      (anchor.name && entry.name === anchor.name)
      || (anchor.fingerprint && entry.anchor?.fingerprint === anchor.fingerprint));
  }
  const candidates = [];
  for (const sheet of inspection.worksheets || inspection.sheets || []) {
    if (anchor.sheet && sheet.name !== anchor.sheet) continue;
    if (anchor.kind === "worksheet") {
      if ((anchor.name && sheet.name === anchor.name) || (anchor.fingerprint && sheet.anchor?.fingerprint === anchor.fingerprint)) candidates.push({ entry: sheet, sheet });
      continue;
    }
    const cells = sheet.cells || [];
    const exact = cells.filter((cell) => anchor.fingerprint && cell.anchor?.fingerprint === anchor.fingerprint);
    if (exact.length) candidates.push(...exact.map((entry) => ({ entry, sheet })));
    else {
      const semantic = cells.filter((cell) => (!anchor.valueHash || cell.anchor?.valueHash === anchor.valueHash) && (!anchor.formulaHash || cell.anchor?.formulaHash === anchor.formulaHash));
      if (anchor.valueHash || anchor.formulaHash) candidates.push(...semantic.map((entry) => ({ entry, sheet })));
      else candidates.push(...cells.filter((cell) => anchor.address && cell.address === anchor.address).map((entry) => ({ entry, sheet })));
    }
  }
  return candidates;
}

function powerpointCandidates(inspection, anchor) {
  const candidates = [];
  for (const slide of inspection.slides || []) {
    const slideMatches = !anchor.slideId || String(slide.anchor?.slideId) === String(anchor.slideId);
    if (anchor.kind === "slide") {
      if ((anchor.fingerprint && slide.anchor?.fingerprint === anchor.fingerprint) || (slideMatches && (!anchor.title || slide.anchor?.title === anchor.title))) candidates.push({ entry: slide, slide });
      continue;
    }
    for (const shape of flattenShapes(slide.shapes)) {
      if (!slideMatches && anchor.slideId) continue;
      if (anchor.fingerprint && shape.anchor?.fingerprint === anchor.fingerprint) candidates.push({ entry: shape, slide });
      else if (anchor.shapeId != null && String(shape.id ?? shape.shapeId) === String(anchor.shapeId)) candidates.push({ entry: shape, slide });
      else if (anchor.shapeName && shape.name === anchor.shapeName && (!anchor.textHash || shape.anchor?.textHash === anchor.textHash)) candidates.push({ entry: shape, slide });
    }
  }
  return candidates;
}

function conflict(operationIndex, anchor, candidates, reason) {
  return {
    operationIndex,
    reason,
    anchor,
    candidateCount: candidates.length,
    candidates: candidates.slice(0, 10).map(({ entry, slide, sheet, index }) => ({
      index: index ?? entry.index ?? null,
      sheet: sheet?.name ?? null,
      slideIndex: slide?.index ?? null,
      id: entry.id ?? entry.shapeId ?? null,
      name: entry.name ?? null,
      text: normalized(entry.text).slice(0, 160)
    }))
  };
}

export function resolveSemanticOperations(kind, inspection, operations = [], rebasePolicy = "unique") {
  const resolved = [];
  const resolutions = [];
  const conflicts = [];
  for (const [operationIndex, rawOperation] of operations.entries()) {
    const operation = structuredClone(rawOperation);
    const anchor = operation.anchor;
    const operationPolicy = operation.rebasePolicy || rebasePolicy || "unique";
    const legacy = {
      paragraphIndex: operation.paragraphIndex, afterIndex: operation.afterIndex, afterParagraphIndex: operation.afterParagraphIndex,
      tableIndex: operation.tableIndex, contentControlIndex: operation.contentControlIndex,
      sheet: operation.sheet, address: operation.address, table: operation.table,
      slideIndex: operation.slideIndex, shapeId: operation.shapeId
    };
    delete operation.anchor;
    delete operation.rebasePolicy;
    if (!anchor) {
      resolved.push(operation);
      continue;
    }
    const candidates = kind === "word" ? wordCandidates(inspection, anchor)
      : kind === "excel" ? excelCandidates(inspection, anchor)
        : powerpointCandidates(inspection, anchor);
    if (candidates.length !== 1 || operationPolicy === "fail") {
      const exact = candidates.filter(({ entry }) => anchor.fingerprint && entry.anchor?.fingerprint === anchor.fingerprint);
      if (exact.length === 1) candidates.splice(0, candidates.length, exact[0]);
      else {
        conflicts.push(conflict(operationIndex, anchor, candidates, candidates.length ? "ambiguous_anchor" : "anchor_not_found"));
        resolved.push(operation);
        continue;
      }
    }
    const candidate = candidates[0];
    if (kind === "word") {
      if (anchor.kind === "table") operation.tableIndex = candidate.index;
      else if (anchor.kind === "contentControl") operation.contentControlIndex = candidate.index;
      else {
        operation.paragraphIndex = candidate.entry.index ?? candidate.index;
        operation.part = candidate.entry.part;
        if (Object.hasOwn(operation, "afterIndex")) operation.afterIndex = operation.paragraphIndex;
        if (Object.hasOwn(operation, "afterParagraphIndex")) operation.afterParagraphIndex = operation.paragraphIndex;
      }
    } else if (kind === "excel") {
      if (anchor.kind === "excelTable") operation.table = candidate.entry.name;
      else if (anchor.kind === "worksheet") operation.sheet = candidate.sheet?.name ?? candidate.entry.name;
      else {
        operation.sheet = candidate.sheet.name;
        operation.address = candidate.entry.address;
      }
    } else {
      operation.slideIndex = candidate.slide?.index ?? candidate.entry.index;
      if (anchor.kind !== "slide") operation.shapeId = candidate.entry.id ?? candidate.entry.shapeId;
    }
    const selectorPairs = kind === "word"
      ? [["paragraphIndex", operation.paragraphIndex], ["afterIndex", operation.afterIndex], ["afterParagraphIndex", operation.afterParagraphIndex], ["tableIndex", operation.tableIndex], ["contentControlIndex", operation.contentControlIndex]]
      : kind === "excel"
        ? [["sheet", operation.sheet], ["address", operation.address], ["table", operation.table]]
        : [["slideIndex", operation.slideIndex], ["shapeId", operation.shapeId]];
    const mismatches = selectorPairs.filter(([key, resolvedValue]) => legacy[key] !== undefined && String(legacy[key]) !== String(resolvedValue));
    if (mismatches.length) {
      conflicts.push({ ...conflict(operationIndex, anchor, candidates, "selector_anchor_mismatch"), mismatches: mismatches.map(([key, resolvedValue]) => ({ key, supplied: legacy[key], resolved: resolvedValue })) });
      resolved.push(operation);
      continue;
    }
    resolutions.push({ operationIndex, anchor, resolved: { paragraphIndex: operation.paragraphIndex, tableIndex: operation.tableIndex, contentControlIndex: operation.contentControlIndex, sheet: operation.sheet, address: operation.address, table: operation.table, slideIndex: operation.slideIndex, shapeId: operation.shapeId } });
    resolved.push(operation);
  }
  return { operations: resolved, resolutions, conflicts };
}

export function semanticAnchorHash(value) {
  return shortHash(typeof value === "string" ? value : JSON.stringify(value));
}
