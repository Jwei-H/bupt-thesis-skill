'use strict';

const fs = require('fs');
const path = require('path');
const { createRequire } = require('module');

function candidateNodeModulePaths() {
  const candidates = [];
  if (process.env.NODE_PATH) {
    candidates.push(...process.env.NODE_PATH.split(path.delimiter).filter(Boolean));
  }
  candidates.push(path.join(process.env.HOME || '', '.workbuddy', 'binaries', 'node', 'workspace', 'node_modules'));
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
        // continue
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

function setRunText(run, text) {
  const doc = run.ownerDocument;
  const textNodes = getRunTextNodes(run);
  textNodes.forEach((node) => run.removeChild(node));
  const textNode = doc.createElement('w:t');
  if (/^[\s　]|[\s　]$/.test(text) || / {2,}/.test(text)) {
    textNode.setAttribute('xml:space', 'preserve');
  }
  textNode.appendChild(doc.createTextNode(text));
  run.appendChild(textNode);
}

function runHasUnderline(run) {
  for (let child = run.firstChild; child; child = child.nextSibling) {
    if (child.nodeType === 1 && localName(child) === 'rPr') {
      for (let item = child.firstChild; item; item = item.nextSibling) {
        if (item.nodeType === 1 && localName(item) === 'u') {
          return true;
        }
      }
    }
  }
  return false;
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

function normalizeCoverFieldLabel(text) {
  const normalizedMap = {
    '姓    名': '姓　　名',
    '学    院': '学　　院',
    '专    业': '专　　业',
    '班    级': '班　　级',
    '学    号': '学　　号',
  };
  return normalizedMap[text] || text;
}

function normalizeCompactText(text) {
  return String(text || '').replace(/[ \t\r\n　]/g, '');
}

function getParagraphRuns(paragraph) {
  return getElementChildren(paragraph).filter((node) => localName(node) === 'r');
}

function getFirstPageFieldKey(compactText) {
  if (compactText.startsWith('题目:') || compactText.startsWith('题目：')) {
    return 'title';
  }
  const fieldMap = {
    '姓名': 'name',
    '学院': 'school',
    '专业': 'major',
    '班级': 'class',
    '学号': 'studentId',
    '指导教师': 'advisor',
  };
  return fieldMap[compactText] || null;
}

function normalizeCoverDataValue(value) {
  if (value === undefined || value === null) {
    return '';
  }
  return String(value).trim();
}

function loadCoverData(coverDataPath) {
  if (!coverDataPath) {
    return null;
  }
  const rawText = fs.readFileSync(coverDataPath, 'utf8');
  let parsed;
  try {
    parsed = JSON.parse(rawText);
  } catch (error) {
    throw new Error(`封面信息 JSON 解析失败: ${coverDataPath}\n${error.message}`);
  }
  if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) {
    throw new Error(`封面信息 JSON 必须是对象: ${coverDataPath}`);
  }

  const now = new Date();
  return {
    title: normalizeCoverDataValue(parsed.title),
    name: normalizeCoverDataValue(parsed.name),
    school: normalizeCoverDataValue(parsed.school),
    major: normalizeCoverDataValue(parsed.major),
    class: normalizeCoverDataValue(parsed.class || parsed.className),
    studentId: normalizeCoverDataValue(parsed.studentId),
    advisor: normalizeCoverDataValue(parsed.advisor),
    year: String(now.getFullYear()),
    month: String(now.getMonth() + 1),
  };
}

function getRunProperties(run) {
  for (let child = run.firstChild; child; child = child.nextSibling) {
    if (child.nodeType === 1 && localName(child) === 'rPr') {
      return child;
    }
  }
  const rPr = run.ownerDocument.createElement('w:rPr');
  run.insertBefore(rPr, run.firstChild);
  return rPr;
}

function removeRunProperty(run, propertyLocalName) {
  const rPr = getRunProperties(run);
  const toRemove = [];
  for (let child = rPr.firstChild; child; child = child.nextSibling) {
    if (child.nodeType === 1 && localName(child) === propertyLocalName) {
      toRemove.push(child);
    }
  }
  toRemove.forEach((node) => rPr.removeChild(node));
}

function ensureRunProperty(run, propertyLocalName) {
  const rPr = getRunProperties(run);
  for (let child = rPr.firstChild; child; child = child.nextSibling) {
    if (child.nodeType === 1 && localName(child) === propertyLocalName) {
      return child;
    }
  }
  const propertyNode = run.ownerDocument.createElement(`w:${propertyLocalName}`);
  rPr.appendChild(propertyNode);
  return propertyNode;
}

function setRunBold(run, enabled) {
  if (enabled) {
    ensureRunProperty(run, 'b');
    ensureRunProperty(run, 'bCs');
    return;
  }
  removeRunProperty(run, 'b');
  removeRunProperty(run, 'bCs');
}

function setRunFontSize(run, halfPoints) {
  const sz = ensureRunProperty(run, 'sz');
  sz.setAttribute('w:val', String(halfPoints));
  const szCs = ensureRunProperty(run, 'szCs');
  szCs.setAttribute('w:val', String(halfPoints));
}

function setRunFonts(run, fonts) {
  const rFonts = ensureRunProperty(run, 'rFonts');
  Object.entries(fonts).forEach(([key, value]) => {
    if (value) {
      rFonts.setAttribute(`w:${key}`, value);
    }
  });
}

function removeUnderlineFormatting(run) {
  removeRunProperty(run, 'u');
}

function textDisplayWidth(text) {
  return Array.from(String(text || '')).reduce((sum, char) => sum + (char.charCodeAt(0) <= 0x7f ? 1 : 2), 0);
}

function spacePlaceholderWidth(text) {
  return Array.from(text).reduce((sum, char) => sum + (char === '　' ? 2 : 1), 0);
}

function clearRunContent(run) {
  const removableChildren = [];
  for (let child = run.firstChild; child; child = child.nextSibling) {
    if (!(child.nodeType === 1 && localName(child) === 'rPr')) {
      removableChildren.push(child);
    }
  }
  removableChildren.forEach((child) => run.removeChild(child));
}

function insertNodeAfter(referenceNode, newNode) {
  const parent = referenceNode.parentNode;
  if (!parent) {
    return;
  }
  if (referenceNode.nextSibling) {
    parent.insertBefore(newNode, referenceNode.nextSibling);
    return;
  }
  parent.appendChild(newNode);
}

function tokenizeWrapText(text) {
  return String(text || '').match(/[A-Za-z0-9]+|./gu) || [];
}

function findBalancedTwoLineWrap(text, maxWidth) {
  const limit = Math.max(1, maxWidth);
  const tokens = tokenizeWrapText(text);
  if (tokens.length < 2) {
    return null;
  }

  const tokenWidths = tokens.map((token) => textDisplayWidth(token));
  const totalWidth = tokenWidths.reduce((sum, width) => sum + width, 0);
  if (totalWidth <= limit || totalWidth > limit * 2) {
    return null;
  }

  const badLineStartPattern = /^[）】，。、；：！？]$/;
  const badLineEndPattern = /^[（【《“]$/;
  let leftWidth = 0;
  let best = null;

  for (let index = 1; index < tokens.length; index += 1) {
    leftWidth += tokenWidths[index - 1];
    const rightWidth = totalWidth - leftWidth;
    if (leftWidth > limit || rightWidth > limit) {
      continue;
    }

    const leftText = tokens.slice(0, index).join('');
    const rightText = tokens.slice(index).join('');
    let score = Math.abs(leftWidth - rightWidth);
    if (badLineEndPattern.test(tokens[index - 1])) {
      score += 100;
    }
    if (badLineStartPattern.test(tokens[index])) {
      score += 100;
    }
    if (/^[A-Za-z0-9]+$/.test(tokens[index - 1]) && /^[A-Za-z0-9]+$/.test(tokens[index])) {
      score += 100;
    }
    if (Math.min(leftWidth, rightWidth) < Math.max(4, Math.floor(limit * 0.35))) {
      score += 40;
    }

    if (!best || score < best.score) {
      best = { score, lines: [leftText, rightText] };
    }
  }

  return best ? best.lines : null;
}

function wrapTextByDisplayWidth(text, maxWidth) {
  const limit = Math.max(1, maxWidth);
  const balancedLines = findBalancedTwoLineWrap(text, limit);
  if (balancedLines) {
    return balancedLines;
  }

  const tokens = tokenizeWrapText(text);
  const lines = [];
  let currentLine = '';
  let currentWidth = 0;

  const pushCurrentLine = () => {
    if (currentLine) {
      lines.push(currentLine);
      currentLine = '';
      currentWidth = 0;
    }
  };

  const appendToken = (token) => {
    const tokenWidth = textDisplayWidth(token);
    if (tokenWidth > limit) {
      for (const char of Array.from(token)) {
        const charWidth = textDisplayWidth(char);
        if (currentWidth + charWidth > limit && currentLine) {
          pushCurrentLine();
        }
        currentLine += char;
        currentWidth += charWidth;
      }
      return;
    }

    if (currentWidth + tokenWidth > limit && currentLine) {
      pushCurrentLine();
    }
    currentLine += token;
    currentWidth += tokenWidth;
  };

  tokens.forEach(appendToken);
  pushCurrentLine();
  return lines.length ? lines : [''];
}

function formatFilledCoverValue(value, placeholderText) {
  if (!value) {
    return '';
  }
  const leftPadding = '  ';
  const availableWidth = Math.max(0, spacePlaceholderWidth(placeholderText));
  const paddedValue = `${leftPadding}${value}`;
  const trailingWidth = Math.max(0, availableWidth - textDisplayWidth(paddedValue));
  return `${paddedValue}${' '.repeat(trailingWidth)}`;
}

function buildContinuationParagraph(paragraph, placeholderIndex, lineText, placeholderText) {
  const continuationParagraph = paragraph.cloneNode(true);
  const continuationRuns = getParagraphRuns(continuationParagraph);
  continuationRuns.forEach((run, index) => {
    clearRunContent(run);
    if (index === placeholderIndex) {
      setRunText(run, formatFilledCoverValue(lineText, placeholderText));
    }
  });
  return continuationParagraph;
}

function fillFirstPageFieldParagraph(paragraph, fillValue) {
  if (!fillValue) {
    return false;
  }

  const runs = getParagraphRuns(paragraph);
  const placeholderIndex = runs.findIndex((run) => runHasUnderline(run));
  if (placeholderIndex === -1) {
    return false;
  }

  const placeholderRun = runs[placeholderIndex];
  const placeholderText = getRunText(placeholderRun);
  const placeholderWidth = Math.max(1, spacePlaceholderWidth(placeholderText));
  const lines = wrapTextByDisplayWidth(fillValue, placeholderWidth);

  clearRunContent(placeholderRun);
  setRunText(placeholderRun, formatFilledCoverValue(lines[0], placeholderText));

  let anchorParagraph = paragraph;
  for (let index = 1; index < lines.length; index += 1) {
    const continuationParagraph = buildContinuationParagraph(paragraph, placeholderIndex, lines[index], placeholderText);
    insertNodeAfter(anchorParagraph, continuationParagraph);
    anchorParagraph = continuationParagraph;
  }

  return true;
}

function isStatementTitleParagraph(compactText) {
  return compactText.startsWith('本人声明所呈交的毕业设计（论文），题目《》是本人在指导教师的指导下');
}

function fillStatementTitleParagraph(paragraph, coverData) {
  if (!coverData || !coverData.title) {
    return false;
  }
  const runs = getParagraphRuns(paragraph);
  const prefixIndex = runs.findIndex((run) => getRunText(run).includes('题目《'));
  const suffixIndex = runs.findIndex((run) => getRunText(run).includes('》是本人在指导教师'));
  if (prefixIndex === -1 || suffixIndex === -1 || suffixIndex <= prefixIndex) {
    return false;
  }
  const fillIndex = Math.min(prefixIndex + 1, suffixIndex - 1);
  setRunText(runs[fillIndex], ` ${coverData.title} `);
  for (let index = fillIndex + 1; index < suffixIndex; index += 1) {
    setRunText(runs[index], '');
  }
  return true;
}

function fillYearMonthParagraph(paragraph, coverData) {
  if (!coverData) {
    return false;
  }
  const runs = getParagraphRuns(paragraph).filter((run) => getRunText(run));
  if (!runs.length) {
    return false;
  }
  setRunText(runs[0], ` ${coverData.year}年`);
  if (runs[1]) {
    setRunText(runs[1], '    ');
  }
  if (runs[2]) {
    setRunText(runs[2], ` ${coverData.month}月`);
  }
  for (let index = 3; index < runs.length; index += 1) {
    setRunText(runs[index], '');
  }
  return true;
}

function polishCoverParagraphs(coverDoc, coverData = null) {
  const body = coverDoc.getElementsByTagName('w:body')[0];
  if (!body) {
    return;
  }

  for (const paragraph of getElementChildren(body)) {
    if (localName(paragraph) !== 'p') {
      continue;
    }
    const paragraphText = getParagraphText(paragraph);
    const compactText = normalizeCompactText(paragraphText);
    if (compactText === '年月') {
      fillYearMonthParagraph(paragraph, coverData);
      continue;
    }
    if (isStatementTitleParagraph(compactText)) {
      fillStatementTitleParagraph(paragraph, coverData);
      continue;
    }

    const firstPageFieldKey = getFirstPageFieldKey(compactText);
    if (!firstPageFieldKey) {
      continue;
    }

    const fillValue = coverData ? coverData[firstPageFieldKey] : '';
    fillFirstPageFieldParagraph(paragraph, fillValue);
  }
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
      spacing: { before: 180, after: 120, line: 360, lineRule: 'auto' },
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
      spacing: { before: 120, after: 90, line: 360, lineRule: 'auto' },
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
        left: 0, right: 0, firstLine: 480, hanging: 0,
        leftChars: 0, rightChars: 0, firstLineChars: 200, hangingChars: 0,
      },
      spacing: { before: 90, after: 60, line: 360, lineRule: 'auto' },
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
      spacing: { before: 90, after: 60, line: 360, lineRule: 'auto' },
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
      fonts: normalFonts,
      size: 24,
      bold: false,
      alignment: null,
      spacing: { before: 0, after: 0, line: 400, lineRule: 'exact' },
      indent: styleId === 'TOC1'
        ? { left: 0, firstLine: 0, right: null, hanging: null }
        : styleId === 'TOC2'
          ? { left: 360, firstLine: 0, right: null, hanging: null }
          : { left: 720, firstLine: 0, right: null, hanging: null },
    });
  }

  const hyperlinkStyle = findStyleById(bodyStylesDoc, 'Hyperlink');
  if (hyperlinkStyle) {
    const runProperties = ensureChildElement(bodyStylesDoc, hyperlinkStyle, 'w:rPr');
    setRunFonts(bodyStylesDoc, runProperties, normalFonts);
    setBooleanRunProperty(bodyStylesDoc, runProperties, 'b', false);
    setBooleanRunProperty(bodyStylesDoc, runProperties, 'bCs', false);
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
    if (styleId !== 'BUPTHeading2') {
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

async function composeDocx({ coverPath, bodyPath, outputPath, coverDataPath = null }) {
  const coverZip = await JSZip.loadAsync(fs.readFileSync(coverPath));
  const bodyZip = await JSZip.loadAsync(fs.readFileSync(bodyPath));
  const coverData = loadCoverData(coverDataPath);

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

  polishCoverParagraphs(coverDoc, coverData);
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
  const args = parseArgs(process.argv.slice(2));
  const coverPath = path.resolve(args.cover || '论文封面+诚信声明.docx');
  const bodyPath = path.resolve(args.body || 'thesis.body.tmp.docx');
  const outputPath = path.resolve(args.output || 'thesis.docx');
  const coverDataPath = args['cover-data'] ? path.resolve(args['cover-data']) : null;

  if (!fs.existsSync(coverPath)) {
    console.error(`封面声明文件不存在: ${coverPath}`);
    process.exit(2);
  }
  if (!fs.existsSync(bodyPath)) {
    console.error(`正文 DOCX 不存在: ${bodyPath}`);
    process.exit(2);
  }
  if (coverDataPath && !fs.existsSync(coverDataPath)) {
    console.error(`封面信息 JSON 不存在: ${coverDataPath}`);
    process.exit(2);
  }

  console.log(`[compose] 前置注入封面声明: ${path.basename(coverPath)}`);
  if (coverDataPath) {
    console.log(`[compose] 使用封面信息 JSON: ${coverDataPath}`);
  }
  await composeDocx({ coverPath, bodyPath, outputPath, coverDataPath });
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
