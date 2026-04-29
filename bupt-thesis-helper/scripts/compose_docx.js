'use strict';

const fs = require('fs');
const path = require('path');
const { createRequire } = require('module');

function appendAncestorNodeModules(candidates, startPath) {
  let current = path.resolve(startPath || process.cwd());
  while (true) {
    candidates.push(path.join(current, 'node_modules'));
    const parent = path.dirname(current);
    if (parent === current) {
      break;
    }
    current = parent;
  }
}

function candidateNodeModulePaths() {
  const candidates = [];
  if (process.env.NODE_PATH) {
    candidates.push(...process.env.NODE_PATH.split(path.delimiter).filter(Boolean));
  }
  appendAncestorNodeModules(candidates, process.cwd());
  appendAncestorNodeModules(candidates, __dirname);
  return [...new Set(candidates.filter(Boolean))];
}

function loadPackage(packageName) {
  try {
    return require(packageName);
  } catch (directError) {
    for (const nodeModulesPath of candidateNodeModulePaths()) {
      try {
        const scopedRequire = createRequire(path.join(nodeModulesPath, '__skill_loader__.js'));
        return scopedRequire(packageName);
      } catch (error) {
      }
    }
    throw new Error(`无法加载依赖 ${packageName}。请先安装 skill 说明中的 Node 依赖。原始错误：${directError.message}`);
  }
}

const JSZip = loadPackage('jszip');
let xmlDomPackage;
try {
  xmlDomPackage = loadPackage('@xmldom/xmldom');
} catch (error) {
  xmlDomPackage = loadPackage('xmldom');
}
const { DOMParser, XMLSerializer } = xmlDomPackage;

function parseArgs(argv) {
  const args = { _: [] };
  for (let index = 0; index < argv.length; index += 1) {
    const token = argv[index];
    if (!token.startsWith('--')) {
      args._.push(token);
      continue;
    }
    const key = token.slice(2);
    const next = argv[index + 1];
    if (!next || next.startsWith('--')) {
      args[key] = true;
      continue;
    }
    args[key] = next;
    index += 1;
  }
  return args;
}

function parseXml(xmlText) {
  return new DOMParser().parseFromString(xmlText, 'text/xml');
}

function serializeXml(node) {
  return new XMLSerializer().serializeToString(node);
}

function getElementChildren(parent) {
  const elements = [];
  for (let child = parent.firstChild; child; child = child.nextSibling) {
    if (child.nodeType === 1) {
      elements.push(child);
    }
  }
  return elements;
}

function localName(node) {
  return node && node.localName ? node.localName : String(node.nodeName || '').split(':').pop();
}

function unionSpaceSeparated(baseValue, extraValue) {
  const merged = new Set((baseValue || '').split(/\s+/).filter(Boolean));
  (extraValue || '').split(/\s+/).filter(Boolean).forEach((item) => merged.add(item));
  return Array.from(merged).join(' ');
}

function mergeRootNamespaces(bodyDoc, coverDoc) {
  const bodyRoot = bodyDoc.documentElement;
  const coverRoot = coverDoc.documentElement;
  for (let index = 0; index < coverRoot.attributes.length; index += 1) {
    const attr = coverRoot.attributes.item(index);
    if (!attr) {
      continue;
    }
    if (attr.name === 'mc:Ignorable') {
      bodyRoot.setAttribute(attr.name, unionSpaceSeparated(bodyRoot.getAttribute(attr.name), attr.value));
      continue;
    }
    if ((attr.name === 'xmlns' || attr.name.startsWith('xmlns:')) && !bodyRoot.hasAttribute(attr.name)) {
      bodyRoot.setAttribute(attr.name, attr.value);
    }
  }
}

function collectRelationshipIds(node, refs = new Set()) {
  if (!node) {
    return refs;
  }
  if (node.nodeType === 1 && node.attributes) {
    for (let index = 0; index < node.attributes.length; index += 1) {
      const attr = node.attributes.item(index);
      if (attr && /^r:(id|embed|link)$/i.test(attr.name) && attr.value) {
        refs.add(attr.value);
      }
    }
  }
  for (let child = node.firstChild; child; child = child.nextSibling) {
    collectRelationshipIds(child, refs);
  }
  return refs;
}

function remapRelationshipIds(node, relationshipIdMap) {
  if (!node || !relationshipIdMap.size) {
    return;
  }
  if (node.nodeType === 1 && node.attributes) {
    for (let index = 0; index < node.attributes.length; index += 1) {
      const attr = node.attributes.item(index);
      if (attr && relationshipIdMap.has(attr.value)) {
        attr.value = relationshipIdMap.get(attr.value);
      }
    }
  }
  for (let child = node.firstChild; child; child = child.nextSibling) {
    remapRelationshipIds(child, relationshipIdMap);
  }
}

function nextRelationshipId(existingIds) {
  let counter = 1;
  while (existingIds.has(`rId${counter}`)) {
    counter += 1;
  }
  return `rId${counter}`;
}

function findDefaultContentType(typesDoc, extension) {
  const normalized = extension.replace(/^\./, '').toLowerCase();
  return getElementChildren(typesDoc.documentElement).find((node) => (
    localName(node) === 'Default'
      && String(node.getAttribute('Extension') || '').toLowerCase() === normalized
  )) || null;
}

function findOverrideContentType(typesDoc, partName) {
  return getElementChildren(typesDoc.documentElement).find((node) => (
    localName(node) === 'Override'
      && node.getAttribute('PartName') === partName
  )) || null;
}

function ensureContentTypeForPart(bodyTypesDoc, coverTypesDoc, sourcePartName, targetPartName) {
  const existingOverride = findOverrideContentType(bodyTypesDoc, targetPartName);
  if (existingOverride) {
    return;
  }

  const coverOverride = findOverrideContentType(coverTypesDoc, sourcePartName);
  if (coverOverride) {
    const overrideNode = coverOverride.cloneNode(true);
    overrideNode.setAttribute('PartName', targetPartName);
    bodyTypesDoc.documentElement.appendChild(overrideNode);
    return;
  }

  const extension = path.posix.extname(targetPartName).replace(/^\./, '').toLowerCase();
  if (!extension || findDefaultContentType(bodyTypesDoc, extension)) {
    return;
  }

  const coverDefault = findDefaultContentType(coverTypesDoc, extension);
  if (coverDefault) {
    bodyTypesDoc.documentElement.appendChild(coverDefault.cloneNode(true));
  }
}

function uniqueZipTarget(bodyZip, relativeTarget) {
  const parsed = path.posix.parse(relativeTarget);
  let counter = 1;
  let candidate = relativeTarget;
  while (bodyZip.file(path.posix.join('word', candidate))) {
    candidate = path.posix.join(parsed.dir, `${parsed.name}_cover${counter}${parsed.ext}`);
    counter += 1;
  }
  return candidate;
}

async function cloneUsedRelationships({ coverZip, bodyZip, coverRelsDoc, bodyRelsDoc, coverTypesDoc, bodyTypesDoc, usedRelationshipIds }) {
  const relationshipIdMap = new Map();
  const bodyRelationshipsRoot = bodyRelsDoc.documentElement;
  const existingIds = new Set(
    getElementChildren(bodyRelationshipsRoot)
      .map((node) => node.getAttribute('Id'))
      .filter(Boolean),
  );

  for (const relationshipNode of getElementChildren(coverRelsDoc.documentElement)) {
    const oldId = relationshipNode.getAttribute('Id');
    if (!oldId || !usedRelationshipIds.has(oldId)) {
      continue;
    }

    const clonedRelationship = relationshipNode.cloneNode(true);
    const newId = nextRelationshipId(existingIds);
    existingIds.add(newId);
    clonedRelationship.setAttribute('Id', newId);

    const targetMode = clonedRelationship.getAttribute('TargetMode');
    const target = clonedRelationship.getAttribute('Target');
    if (target && targetMode !== 'External' && !target.startsWith('/')) {
      const sourceZipPath = path.posix.normalize(path.posix.join('word', target));
      const sourcePart = coverZip.file(sourceZipPath);
      if (!sourcePart) {
        throw new Error(`封面文档缺少被引用资源：${sourceZipPath}`);
      }

      const newTarget = uniqueZipTarget(bodyZip, target);
      const targetZipPath = path.posix.join('word', newTarget);
      bodyZip.file(targetZipPath, await sourcePart.async('nodebuffer'));
      clonedRelationship.setAttribute('Target', newTarget);
      ensureContentTypeForPart(bodyTypesDoc, coverTypesDoc, `/${sourceZipPath}`, `/${targetZipPath}`);
    }

    bodyRelationshipsRoot.appendChild(clonedRelationship);
    relationshipIdMap.set(oldId, newId);
  }

  return relationshipIdMap;
}

function getRunTextNodes(run) {
  const nodes = [];
  for (let child = run.firstChild; child; child = child.nextSibling) {
    if (child.nodeType === 1 && localName(child) === 't') {
      nodes.push(child);
    }
  }
  return nodes;
}

function getRunText(run) {
  return getRunTextNodes(run).map((node) => node.textContent || '').join('');
}

function getParagraphText(paragraph) {
  const chunks = [];
  for (let child = paragraph.firstChild; child; child = child.nextSibling) {
    if (child.nodeType === 1 && localName(child) === 'r') {
      chunks.push(getRunText(child));
    }
  }
  return chunks.join('');
}

function normalizeCompactText(text) {
  return String(text || '').replace(/[ \t\r\n　]/g, '');
}


function mergeMissingStyles(bodyZip, coverZip) {
  const bodyStylesFile = bodyZip.file('word/styles.xml');
  const coverStylesFile = coverZip.file('word/styles.xml');
  if (!bodyStylesFile || !coverStylesFile) {
    return;
  }

  return Promise.all([
    bodyStylesFile.async('string'),
    coverStylesFile.async('string'),
  ]).then(([bodyStylesXml, coverStylesXml]) => {
    const bodyStylesDoc = parseXml(bodyStylesXml);
    const coverStylesDoc = parseXml(coverStylesXml);
    const bodyRoot = bodyStylesDoc.documentElement;
    const existingStyleIds = new Set(
      getElementChildren(bodyRoot)
        .filter((node) => localName(node) === 'style')
        .map((node) => node.getAttribute('w:styleId') || node.getAttribute('styleId'))
        .filter(Boolean),
    );

    let changed = false;
    for (const styleNode of getElementChildren(coverStylesDoc.documentElement)) {
      if (localName(styleNode) !== 'style') {
        continue;
      }
      const styleId = styleNode.getAttribute('w:styleId') || styleNode.getAttribute('styleId');
      if (!styleId || existingStyleIds.has(styleId)) {
        continue;
      }
      bodyRoot.appendChild(styleNode.cloneNode(true));
      existingStyleIds.add(styleId);
      changed = true;
    }

    if (changed) {
      bodyZip.file('word/styles.xml', serializeXml(bodyStylesDoc));
    }
  });
}

function findStyleById(stylesDoc, styleId) {
  return getElementChildren(stylesDoc.documentElement).find((node) => (
    localName(node) === 'style'
      && (node.getAttribute('w:styleId') || node.getAttribute('styleId')) === styleId
  )) || null;
}

function findChildElement(parent, targetLocalName) {
  return getElementChildren(parent).find((node) => localName(node) === targetLocalName) || null;
}

function ensureChildElement(doc, parent, qualifiedName) {
  let node = findChildElement(parent, qualifiedName.split(':').pop());
  if (!node) {
    node = doc.createElement(qualifiedName);
    parent.appendChild(node);
  }
  return node;
}

function setRunFonts(doc, runProperties, fonts) {
  const rFonts = ensureChildElement(doc, runProperties, 'w:rFonts');
  rFonts.setAttribute('w:ascii', fonts.ascii);
  rFonts.setAttribute('w:eastAsia', fonts.eastAsia);
  rFonts.setAttribute('w:hAnsi', fonts.hAnsi);
  rFonts.setAttribute('w:hint', 'eastAsia');
}

function setBooleanRunProperty(doc, runProperties, tagName, value) {
  const property = ensureChildElement(doc, runProperties, `w:${tagName}`);
  property.setAttribute('w:val', value ? 'true' : 'false');
}

function removeChildElement(parent, targetLocalName) {
  const node = findChildElement(parent, targetLocalName);
  if (node) {
    parent.removeChild(node);
    return true;
  }
  return false;
}

function setRunFontSize(doc, runProperties, halfPoints) {
  const size = String(halfPoints);
  const sz = ensureChildElement(doc, runProperties, 'w:sz');
  sz.setAttribute('w:val', size);
  const szCs = ensureChildElement(doc, runProperties, 'w:szCs');
  szCs.setAttribute('w:val', size);
}

function setRunColor(doc, runProperties, color) {
  const colorNode = ensureChildElement(doc, runProperties, 'w:color');
  colorNode.setAttribute('w:val', color);
  colorNode.removeAttribute('w:themeColor');
  colorNode.removeAttribute('w:themeTint');
  colorNode.removeAttribute('w:themeShade');
}

function setRunUnderline(doc, runProperties, value) {
  const underline = ensureChildElement(doc, runProperties, 'w:u');
  underline.setAttribute('w:val', value);
  underline.removeAttribute('w:color');
  underline.removeAttribute('w:themeColor');
  underline.removeAttribute('w:themeTint');
  underline.removeAttribute('w:themeShade');
}

function setParagraphSpacing(doc, paragraphProperties, options = {}) {
  const spacing = ensureChildElement(doc, paragraphProperties, 'w:spacing');
  if (options.before !== undefined) {
    spacing.setAttribute('w:before', String(options.before));
  }
  if (options.after !== undefined) {
    spacing.setAttribute('w:after', String(options.after));
  }
  if (options.beforeLines !== undefined) {
    spacing.setAttribute('w:beforeLines', String(options.beforeLines));
  }
  if (options.afterLines !== undefined) {
    spacing.setAttribute('w:afterLines', String(options.afterLines));
  }
  if (options.line !== undefined) {
    spacing.setAttribute('w:line', String(options.line));
  }
  if (options.lineRule !== undefined) {
    spacing.setAttribute('w:lineRule', String(options.lineRule));
  }
}

function setParagraphIndent(doc, paragraphProperties, options = {}) {
  const indent = ensureChildElement(doc, paragraphProperties, 'w:ind');
  const attrs = [
    'left',
    'right',
    'firstLine',
    'hanging',
    'leftChars',
    'rightChars',
    'firstLineChars',
    'hangingChars',
  ];
  for (const key of attrs) {
    const attrName = `w:${key}`;
    if (options[key] === undefined || options[key] === null) {
      indent.removeAttribute(attrName);
      continue;
    }
    indent.setAttribute(attrName, String(options[key]));
  }
}

function setParagraphAlignment(doc, paragraphProperties, value) {
  const jc = ensureChildElement(doc, paragraphProperties, 'w:jc');
  jc.setAttribute('w:val', value);
}

function setParagraphOutlineLevel(doc, paragraphProperties, level) {
  const outline = ensureChildElement(doc, paragraphProperties, 'w:outlineLvl');
  outline.setAttribute('w:val', String(level));
}

function setParagraphKeepNext(doc, paragraphProperties, enabled) {
  if (enabled) {
    ensureChildElement(doc, paragraphProperties, 'w:keepNext');
    return;
  }
  removeChildElement(paragraphProperties, 'keepNext');
}

function setParagraphKeepLines(doc, paragraphProperties, enabled) {
  if (enabled) {
    ensureChildElement(doc, paragraphProperties, 'w:keepLines');
    return;
  }
  removeChildElement(paragraphProperties, 'keepLines');
}

function ensureTableRowsCantSplit(doc, tableNode) {
  let changed = false;
  for (const rowNode of getElementChildren(tableNode)) {
    if (localName(rowNode) !== 'tr') {
      continue;
    }
    const rowProperties = ensureChildElement(doc, rowNode, 'w:trPr');
    ensureChildElement(doc, rowProperties, 'w:cantSplit');
    changed = true;
  }
  return changed;
}

function normalizeTableCaptionPagination(bodyDoc) {
  const body = bodyDoc.getElementsByTagName('w:body')[0];
  if (!body) {
    return false;
  }

  const nodes = getElementChildren(body);
  let changed = false;

  for (let index = 0; index < nodes.length - 1; index += 1) {
    const current = nodes[index];
    const next = nodes[index + 1];
    if (localName(current) !== 'p' || localName(next) !== 'tbl') {
      continue;
    }

    const paragraphText = normalizeCompactText(getParagraphText(current));
    if (!/^表\d+-\d+/.test(paragraphText)) {
      continue;
    }

    const paragraphProperties = ensureChildElement(bodyDoc, current, 'w:pPr');
    setParagraphKeepNext(bodyDoc, paragraphProperties, true);
    setParagraphKeepLines(bodyDoc, paragraphProperties, true);
    changed = true;

    if (ensureTableRowsCantSplit(bodyDoc, next)) {
      changed = true;
    }
  }

  return changed;
}

function dedupeStyles(stylesDoc) {
  const root = stylesDoc.documentElement;
  const styleNodes = getElementChildren(root).filter((node) => localName(node) === 'style');
  const seen = new Set();
  let changed = false;

  for (let index = styleNodes.length - 1; index >= 0; index -= 1) {
    const node = styleNodes[index];
    const styleId = node.getAttribute('w:styleId') || node.getAttribute('styleId');
    if (!styleId) {
      continue;
    }
    if (seen.has(styleId)) {
      root.removeChild(node);
      changed = true;
      continue;
    }
    seen.add(styleId);
  }

  return changed;
}

async function normalizeStyles(bodyZip) {
  const bodyStylesFile = bodyZip.file('word/styles.xml');
  if (!bodyStylesFile) {
    return;
  }

  const bodyStylesXml = await bodyStylesFile.async('string');
  const bodyStylesDoc = parseXml(bodyStylesXml);
  let changed = dedupeStyles(bodyStylesDoc);

  const normalFonts = { ascii: 'Times New Roman', eastAsia: '宋体', hAnsi: 'Times New Roman' };
  const headingFonts = { ascii: 'Times New Roman', eastAsia: '黑体', hAnsi: 'Times New Roman' };

  const normalizeParagraphStyle = (styleId, options) => {
    const styleNode = findStyleById(bodyStylesDoc, styleId);
    if (!styleNode) {
      return;
    }

    removeChildElement(styleNode, 'link');
    if (styleId.startsWith('BUPTHeading')) {
      removeChildElement(styleNode, 'basedOn');
    }
    const runProperties = ensureChildElement(bodyStylesDoc, styleNode, 'w:rPr');
    setRunFonts(bodyStylesDoc, runProperties, options.fonts);
    setRunFontSize(bodyStylesDoc, runProperties, options.size);
    setBooleanRunProperty(bodyStylesDoc, runProperties, 'b', Boolean(options.bold));
    setBooleanRunProperty(bodyStylesDoc, runProperties, 'bCs', Boolean(options.bold));
    setRunColor(bodyStylesDoc, runProperties, options.color || '000000');
    setRunUnderline(bodyStylesDoc, runProperties, options.underline || 'none');

    const paragraphProperties = ensureChildElement(bodyStylesDoc, styleNode, 'w:pPr');
    if (options.spacing) {
      setParagraphSpacing(bodyStylesDoc, paragraphProperties, options.spacing);
    }
    if (options.indent) {
      setParagraphIndent(bodyStylesDoc, paragraphProperties, options.indent);
    }
    if (options.alignment) {
      setParagraphAlignment(bodyStylesDoc, paragraphProperties, options.alignment);
    }
    if (options.outlineLevel !== undefined && options.outlineLevel !== null) {
      setParagraphOutlineLevel(bodyStylesDoc, paragraphProperties, options.outlineLevel);
    }
    changed = true;
  };

  normalizeParagraphStyle('FrontMatterTitle', {
    fonts: headingFonts,
    size: 32,
    bold: true,
    alignment: 'center',
    indent: {
      left: 0, right: 0, firstLine: 0, hanging: 0,
      leftChars: 0, rightChars: 0, firstLineChars: 0, hangingChars: 0,
    },
    spacing: { before: 240, after: 240, beforeLines: 100, afterLines: 100, line: 360, lineRule: 'auto' },
  });

  for (const styleId of ['Heading1', 'BUPTHeading1']) {
    normalizeParagraphStyle(styleId, {
      fonts: headingFonts,
      size: 32,
      bold: true,
      alignment: 'center',
      indent: {
        left: 0, right: 0, firstLine: 0, hanging: 0,
        leftChars: 0, rightChars: 0, firstLineChars: 0, hangingChars: 0,
      },
      spacing: { before: 0, after: 480, beforeLines: 0, afterLines: 200, line: 360, lineRule: 'auto' },
      outlineLevel: 0,
    });
  }
  for (const styleId of ['Heading2', 'BUPTHeading2']) {
    normalizeParagraphStyle(styleId, {
      fonts: headingFonts,
      size: 28,
      bold: true,
      alignment: 'left',
      indent: {
        left: 0, right: 0, firstLine: 0, hanging: 0,
        leftChars: 0, rightChars: 0, firstLineChars: 0, hangingChars: 0,
      },
      spacing: { before: 120, after: 120, beforeLines: 50, afterLines: 50, line: 360, lineRule: 'auto' },
      outlineLevel: 1,
    });
  }
  for (const styleId of ['Heading3', 'BUPTHeading3']) {
    normalizeParagraphStyle(styleId, {
      fonts: headingFonts,
      size: 24,
      bold: true,
      alignment: 'left',
      indent: {
        left: 0, right: 0, firstLine: 482, hanging: null,
        leftChars: 0, rightChars: 0, firstLineChars: 200, hangingChars: null,
      },
      spacing: { before: 120, after: 120, beforeLines: 50, afterLines: 50, line: 360, lineRule: 'auto' },
      outlineLevel: 2,
    });
  }

  for (const [styleId, outlineLevel] of [['Heading4', 3], ['Heading5', 4], ['Heading6', 5], ['Heading7', 6], ['Heading8', 7], ['Heading9', 8]]) {
    normalizeParagraphStyle(styleId, {
      fonts: headingFonts,
      size: 24,
      bold: true,
      alignment: 'left',
      indent: { left: 0, firstLine: 480, right: null, hanging: null },
      spacing: { before: 120, after: 120, beforeLines: 50, afterLines: 50, line: 360, lineRule: 'auto' },
      outlineLevel,
    });
  }

  normalizeParagraphStyle('TOCHeading', {
    fonts: headingFonts,
    size: 32,
    bold: true,
    alignment: 'center',
    indent: { left: 0, firstLine: 0, right: null, hanging: null },
    spacing: { before: 0, after: 300, line: 360, lineRule: 'auto' },
  });

  for (const styleId of ['TOC1', 'TOC2', 'TOC3']) {
    normalizeParagraphStyle(styleId, {
      fonts: styleId === 'TOC1' ? headingFonts : normalFonts,
      size: 24,
      bold: false,
      alignment: null,
      spacing: { before: 0, after: 0, line: 400, lineRule: 'exact' },
      indent: styleId === 'TOC1'
        ? { left: 0, right: 0, firstLine: 0, hanging: 0, leftChars: 0, rightChars: 0, firstLineChars: 0, hangingChars: 0 }
        : styleId === 'TOC2'
          ? { left: 420, right: 0, firstLine: 0, hanging: 0, leftChars: 200, rightChars: 0, firstLineChars: 0, hangingChars: 0 }
          : { left: 840, right: 0, firstLine: 0, hanging: 0, leftChars: 400, rightChars: 0, firstLineChars: 0, hangingChars: 0 },
    });
  }

  normalizeParagraphStyle('ImageBlock', {
    fonts: normalFonts,
    size: 24,
    bold: false,
    alignment: null,
    spacing: { before: 0, after: 0, line: 360, lineRule: 'auto' },
    indent: {
      left: 0, right: 0, firstLine: 0, hanging: 0,
      leftChars: 0, rightChars: 0, firstLineChars: 0, hangingChars: 0,
    },
  });

  normalizeParagraphStyle('ReferenceEntry', {
    fonts: normalFonts,
    size: 21,
    bold: false,
    alignment: null,
    spacing: { before: 0, after: 0, line: 360, lineRule: 'auto' },
    indent: {
      left: 340, right: 0, firstLine: 0, hanging: 340,
      leftChars: 0, rightChars: 0, firstLineChars: 0, hangingChars: 0,
    },
  });

  normalizeParagraphStyle('AlgorithmBlock', {
    fonts: { ascii: 'Times New Roman', eastAsia: '楷体', hAnsi: 'Times New Roman' },
    size: 21,
    bold: false,
    alignment: null,
    spacing: { before: 0, after: 0, line: 300, lineRule: 'auto' },
    indent: {
      left: 0, right: 0, firstLine: 0, hanging: 0,
      leftChars: 0, rightChars: 0, firstLineChars: 0, hangingChars: 0,
    },
  });

  const hyperlinkStyle = findStyleById(bodyStylesDoc, 'Hyperlink');
  if (hyperlinkStyle) {
    const runProperties = ensureChildElement(bodyStylesDoc, hyperlinkStyle, 'w:rPr');
    removeChildElement(runProperties, 'rFonts');
    removeChildElement(runProperties, 'b');
    removeChildElement(runProperties, 'bCs');
    setRunColor(bodyStylesDoc, runProperties, '000000');
    setRunUnderline(bodyStylesDoc, runProperties, 'none');
    changed = true;
  }

  if (changed) {
    bodyZip.file('word/styles.xml', serializeXml(bodyStylesDoc));
  }
}

function normalizeBodyHeadingParagraphs(bodyDoc) {
  let changed = false;
  const paragraphs = Array.from(bodyDoc.getElementsByTagName('w:p'));

  for (const paragraph of paragraphs) {
    const paragraphProperties = findChildElement(paragraph, 'pPr');
    if (!paragraphProperties) {
      continue;
    }
    const styleNode = findChildElement(paragraphProperties, 'pStyle');
    const styleId = styleNode && (styleNode.getAttribute('w:val') || styleNode.getAttribute('val'));
    if (!['BUPTHeading2', 'BUPTHeading3', 'Heading2', 'Heading3'].includes(styleId)) {
      continue;
    }
    setParagraphIndent(bodyDoc, paragraphProperties, styleId === 'BUPTHeading3' || styleId === 'Heading3'
      ? {
        left: 0,
        right: 0,
        firstLine: 482,
        hanging: null,
        leftChars: 0,
        rightChars: 0,
        firstLineChars: 200,
        hangingChars: null,
      }
      : {
        left: 0,
        right: 0,
        firstLine: 0,
        hanging: 0,
        leftChars: 0,
        rightChars: 0,
        firstLineChars: 0,
        hangingChars: 0,
      });
    changed = true;
  }

  return changed;
}

function normalizeCenteredParagraphIndents(bodyDoc) {
  let changed = false;
  const paragraphs = Array.from(bodyDoc.getElementsByTagName('w:p'));
  for (const paragraph of paragraphs) {
    const paragraphProperties = findChildElement(paragraph, 'pPr');
    if (!paragraphProperties) {
      continue;
    }
    const alignmentNode = findChildElement(paragraphProperties, 'jc');
    const alignment = alignmentNode && (alignmentNode.getAttribute('w:val') || alignmentNode.getAttribute('val'));
    if (alignment !== 'center') {
      continue;
    }
    setParagraphIndent(bodyDoc, paragraphProperties, {
      left: 0,
      right: 0,
      firstLine: 0,
      hanging: 0,
      leftChars: 0,
      rightChars: 0,
      firstLineChars: 0,
      hangingChars: 0,
    });
    changed = true;
  }
  return changed;
}

function normalizeDirectParagraphSpacing(bodyDoc) {
  let changed = false;
  const paragraphs = Array.from(bodyDoc.getElementsByTagName('w:p'));
  for (const paragraph of paragraphs) {
    const paragraphText = normalizeCompactText(getParagraphText(paragraph));
    if (!paragraphText) {
      continue;
    }
    const paragraphProperties = ensureChildElement(bodyDoc, paragraph, 'w:pPr');
    const styleNode = findChildElement(paragraphProperties, 'pStyle');
    const styleId = styleNode && (styleNode.getAttribute('w:val') || styleNode.getAttribute('val'));

    if (/^(图|表)\d+-\d+/.test(paragraphText)) {
      setParagraphSpacing(bodyDoc, paragraphProperties, { before: 120, after: 120, beforeLines: 50, afterLines: 50, line: 360, lineRule: 'auto' });
      changed = true;
      continue;
    }

    if (styleId === 'BUPTHeading1' || styleId === 'Heading1') {
      setParagraphSpacing(bodyDoc, paragraphProperties, { before: 0, after: 480, beforeLines: 0, afterLines: 200, line: 360, lineRule: 'auto' });
      changed = true;
      continue;
    }

    if (/^(摘要|ABSTRACT|目录)$/.test(paragraphText)) {
      setParagraphSpacing(bodyDoc, paragraphProperties, { before: 0, after: 0, beforeLines: 0, afterLines: 0, line: 360, lineRule: 'auto' });
      changed = true;
      continue;
    }
  }
  return changed;
}

function normalizeSectionPageSetup(bodyDoc) {
  let changed = false;
  const sectionProperties = Array.from(bodyDoc.getElementsByTagName('w:sectPr'));
  for (const sectPr of sectionProperties) {
    const pageSize = ensureChildElement(bodyDoc, sectPr, 'w:pgSz');
    pageSize.setAttribute('w:w', '11906');
    pageSize.setAttribute('w:h', '16838');
    const pageMargin = ensureChildElement(bodyDoc, sectPr, 'w:pgMar');
    pageMargin.setAttribute('w:top', '1418');
    pageMargin.setAttribute('w:right', '1417');
    pageMargin.setAttribute('w:bottom', '1418');
    pageMargin.setAttribute('w:left', '1417');
    pageMargin.setAttribute('w:header', '851');
    pageMargin.setAttribute('w:footer', '851');
    pageMargin.setAttribute('w:gutter', '0');
    changed = true;
  }
  return changed;
}

function prependCoverBody(bodyDoc, coverDoc, relationshipIdMap) {
  const bodyContainer = bodyDoc.getElementsByTagName('w:body')[0];
  const coverContainer = coverDoc.getElementsByTagName('w:body')[0];
  const anchor = bodyContainer.firstChild;
  let coverSectionProps = null;

  for (const child of getElementChildren(coverContainer)) {
    if (localName(child) === 'sectPr') {
      coverSectionProps = child.cloneNode(true);
      continue;
    }
    const clonedChild = child.cloneNode(true);
    remapRelationshipIds(clonedChild, relationshipIdMap);
    bodyContainer.insertBefore(clonedChild, anchor);
  }

  if (coverSectionProps) {
    const sectionBreakParagraph = bodyDoc.createElement('w:p');
    const paragraphProperties = bodyDoc.createElement('w:pPr');
    paragraphProperties.appendChild(coverSectionProps);
    sectionBreakParagraph.appendChild(paragraphProperties);
    bodyContainer.insertBefore(sectionBreakParagraph, anchor);
  }
}

function validateDocumentRelationships(zip, relsDoc) {
  const missingTargets = [];
  for (const relationshipNode of getElementChildren(relsDoc.documentElement)) {
    const targetMode = relationshipNode.getAttribute('TargetMode');
    const target = relationshipNode.getAttribute('Target');
    if (!target || targetMode === 'External' || target.startsWith('/')) {
      continue;
    }
    const zipPath = path.posix.normalize(path.posix.join('word', target));
    if (!zip.file(zipPath)) {
      missingTargets.push(`${relationshipNode.getAttribute('Id') || '(no-id)'} -> ${zipPath}`);
    }
  }
  if (missingTargets.length) {
    throw new Error(`DOCX 关系校验失败，存在缺失资源：\n${missingTargets.join('\n')}`);
  }
}

async function composeDocx({ coverPath, bodyPath, outputPath }) {
  const coverZip = await JSZip.loadAsync(fs.readFileSync(coverPath));
  const bodyZip = await JSZip.loadAsync(fs.readFileSync(bodyPath));

  const coverDocumentXml = await coverZip.file('word/document.xml').async('string');
  const bodyDocumentXml = await bodyZip.file('word/document.xml').async('string');
  const coverRelsXml = await coverZip.file('word/_rels/document.xml.rels').async('string');
  const bodyRelsXml = await bodyZip.file('word/_rels/document.xml.rels').async('string');
  const coverTypesXml = await coverZip.file('[Content_Types].xml').async('string');
  const bodyTypesXml = await bodyZip.file('[Content_Types].xml').async('string');

  const coverDoc = parseXml(coverDocumentXml);
  const bodyDoc = parseXml(bodyDocumentXml);
  const coverRelsDoc = parseXml(coverRelsXml);
  const bodyRelsDoc = parseXml(bodyRelsXml);
  const coverTypesDoc = parseXml(coverTypesXml);
  const bodyTypesDoc = parseXml(bodyTypesXml);

  mergeRootNamespaces(bodyDoc, coverDoc);
  const coverBody = coverDoc.getElementsByTagName('w:body')[0];
  const usedRelationshipIds = collectRelationshipIds(coverBody);
  const relationshipIdMap = await cloneUsedRelationships({
    coverZip,
    bodyZip,
    coverRelsDoc,
    bodyRelsDoc,
    coverTypesDoc,
    bodyTypesDoc,
    usedRelationshipIds,
  });

  prependCoverBody(bodyDoc, coverDoc, relationshipIdMap);
  normalizeBodyHeadingParagraphs(bodyDoc);
  normalizeCenteredParagraphIndents(bodyDoc);
  normalizeDirectParagraphSpacing(bodyDoc);
  normalizeSectionPageSetup(bodyDoc);
  normalizeTableCaptionPagination(bodyDoc);
  await mergeMissingStyles(bodyZip, coverZip);
  await normalizeStyles(bodyZip);

  bodyZip.file('word/document.xml', serializeXml(bodyDoc));
  bodyZip.file('word/_rels/document.xml.rels', serializeXml(bodyRelsDoc));
  bodyZip.file('[Content_Types].xml', serializeXml(bodyTypesDoc));

  validateDocumentRelationships(bodyZip, bodyRelsDoc);

  const outputBuffer = await bodyZip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 9 },
  });
  fs.writeFileSync(outputPath, outputBuffer);
}

async function main() {
  const skillRoot = path.resolve(__dirname, '..');
  const args = parseArgs(process.argv.slice(2));
  const coverInput = args.cover || path.join(skillRoot, 'assets', '论文封面+诚信声明.docx');
  const bodyInput = args.body || args._[0];
  if (!bodyInput) {
    console.error('错误: 请通过 --body <body-docx-path> 或位置参数显式指定正文 DOCX 路径。');
    process.exit(1);
  }
  const coverPath = path.resolve(coverInput);
  const bodyPath = path.resolve(bodyInput);
  const outputPath = path.resolve(args.output || path.join(path.dirname(bodyPath), `${path.parse(bodyPath).name}.docx`));

  if (!fs.existsSync(coverPath)) {
    console.error(`封面声明文件不存在: ${coverPath}`);
    process.exit(2);
  }
  if (!fs.existsSync(bodyPath)) {
    console.error(`正文 DOCX 不存在: ${bodyPath}`);
    process.exit(2);
  }
  console.log(`[compose] 前置注入封面声明: ${path.basename(coverPath)}`);
  await composeDocx({ coverPath, bodyPath, outputPath });
  console.log(`[compose] 输出完成: ${outputPath}`);
}

if (require.main === module) {
  main().catch((error) => {
    console.error(error && error.stack ? error.stack : String(error));
    process.exit(1);
  });
}

module.exports = {
  composeDocx,
};
