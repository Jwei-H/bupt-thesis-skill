/**
 * thesis.md -> thesis.docx
 * 北京邮电大学本科毕业论文 Word 生成脚本
 */

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
    throw directError;
  }
}

function parseArgs(argv) {
  const args = {};
  for (let index = 0; index < argv.length; index += 1) {
    const token = argv[index];
    if (!token.startsWith('--')) {
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

const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  ImageRun,
  Header,
  Footer,
  AlignmentType,
  BorderStyle,
  WidthType,
  ShadingType,
  VerticalAlign,
  PageBreak,
  PageNumber,
  TableOfContents,
  LevelFormat,
  InternalHyperlink,
  Math: DocxMath,
  MathRun,
  MathFraction,
  MathSubScript,
  MathSuperScript,
  MathRadical,
  ImportedXmlComponent,
} = loadPackage('docx');

const cliArgs = parseArgs(process.argv.slice(2));
const workspaceRoot = path.resolve(cliArgs.workspace || process.cwd());
const mdPath = path.resolve(workspaceRoot, cliArgs.input || 'thesis.md');
const outPath = path.resolve(workspaceRoot, cliArgs.output || 'thesis.docx');
const markdownDir = path.dirname(mdPath);
const krokiBaseUrl = String(cliArgs['kroki-base'] || 'https://kroki.io').replace(/\/+$/, '');

const PAGE_WIDTH = 11906;
const PAGE_HEIGHT = 16838;
// 官方要求：上、下、左、右各 2.5cm；页眉 1.5cm，页脚 1.5cm
// 1cm ≈ 567 twips
const MARGIN_TOP = 1418;    // 2.5cm
const MARGIN_BOTTOM = 1418; // 2.5cm
const MARGIN_LEFT = 1418;   // 2.5cm
const MARGIN_RIGHT = 1418;  // 2.5cm
const MARGIN_HEADER = 851;  // 1.5cm
const MARGIN_FOOTER = 851;  // 1.5cm
const CONTENT_WIDTH = PAGE_WIDTH - MARGIN_LEFT - MARGIN_RIGHT;

const SIZE = {
  SAN_HAO: 32,
  XIAO_SAN: 30,
  SI_HAO: 28,
  XIAO_SI: 24,
  WU_HAO: 21,
  XIAO_WU: 18,
};

const FONT_CN = '宋体';
const FONT_HEI = '黑体';
const FONT_KAI = '楷体';
const FONT_EN = 'Times New Roman';
const BODY_HEADING_STYLES = {
  1: 'BUPTHeading1',
  2: 'BUPTHeading2',
  3: 'BUPTHeading3',
};
const SUPPORTED_IMAGE_TYPES = {
  png: 'png',
  jpg: 'jpg',
  jpeg: 'jpg',
  gif: 'gif',
  bmp: 'bmp',
};
const imageAssetCache = new Map();

const mdText = fs.readFileSync(mdPath, 'utf-8');
const lines = mdText.split(/\r?\n/);

function makeRun(text, opts = {}) {
  return new TextRun({
    text,
    size: opts.size || SIZE.XIAO_SI,
    bold: opts.bold || false,
    italics: opts.italics || false,
    superScript: opts.superScript || false,
    underline: opts.underline,
    color: opts.color,
    break: opts.break || 0,
    font: opts.font || { ascii: FONT_EN, eastAsia: FONT_CN, hAnsi: FONT_EN },
  });
}

function lineSpacing(before = 0, after = 0, line = 360) {
  return {
    line,
    lineRule: 'auto',
    before,
    after,
  };
}

function indentTwoChars() {
  return { firstLine: 480 };
}

function zeroIndent() {
  return {
    left: 0,
    right: 0,
    firstLine: 0,
    hanging: 0,
  };
}

function normalizeImageSource(src) {
  const trimmed = String(src || '').trim();
  if (trimmed.startsWith('<') && trimmed.endsWith('>')) {
    return trimmed.slice(1, -1).trim();
  }
  return trimmed;
}

function isRemoteImageSource(src) {
  return /^https?:\/\//i.test(normalizeImageSource(src));
}

function getSourcePathname(src) {
  const normalized = normalizeImageSource(src);
  if (isRemoteImageSource(normalized)) {
    try {
      return new URL(normalized).pathname || normalized;
    } catch (error) {
      return normalized;
    }
  }
  return normalized;
}

function inferImageExtension(src, contentType = '') {
  const pathname = getSourcePathname(src);
  const ext = path.extname(pathname).replace('.', '').toLowerCase();
  if (SUPPORTED_IMAGE_TYPES[ext]) {
    return ext;
  }

  const normalizedType = String(contentType || '').split(';')[0].trim().toLowerCase();
  if (normalizedType === 'image/png') {
    return 'png';
  }
  if (normalizedType === 'image/jpeg' || normalizedType === 'image/jpg') {
    return 'jpg';
  }
  if (normalizedType === 'image/gif') {
    return 'gif';
  }
  if (normalizedType === 'image/bmp' || normalizedType === 'image/x-ms-bmp') {
    return 'bmp';
  }

  return ext;
}

function buildImagePlaceholderParagraph(label, spacing = lineSpacing(60, 60)) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing,
    indent: zeroIndent(),
    children: [makeRun(label, { size: SIZE.WU_HAO })],
  });
}

function loadLocalImageAssetSync(src) {
  const normalized = normalizeImageSource(src);
  const imagePath = path.resolve(markdownDir, normalized);
  const ext = inferImageExtension(normalized);
  const type = SUPPORTED_IMAGE_TYPES[ext];

  if (!fs.existsSync(imagePath) || !type) {
    const asset = {
      ok: false,
      src: normalized,
      reason: !fs.existsSync(imagePath) ? '图片文件不存在' : `暂不支持的图片类型：${ext || 'unknown'}`,
    };
    imageAssetCache.set(normalized, asset);
    return asset;
  }

  const asset = {
    ok: true,
    src: normalized,
    buffer: fs.readFileSync(imagePath),
    ext,
    type,
  };
  imageAssetCache.set(normalized, asset);
  return asset;
}

async function loadRemoteImageAsset(src) {
  const normalized = normalizeImageSource(src);
  try {
    const response = await fetch(normalized);
    if (!response.ok) {
      const asset = {
        ok: false,
        src: normalized,
        reason: `下载失败：HTTP ${response.status}`,
      };
      imageAssetCache.set(normalized, asset);
      return asset;
    }

    const contentType = response.headers.get('content-type') || '';
    const ext = inferImageExtension(normalized, contentType);
    const type = SUPPORTED_IMAGE_TYPES[ext];
    if (!type) {
      const asset = {
        ok: false,
        src: normalized,
        reason: `暂不支持的远程图片类型：${contentType || ext || 'unknown'}`,
      };
      imageAssetCache.set(normalized, asset);
      return asset;
    }

    const asset = {
      ok: true,
      src: normalized,
      buffer: Buffer.from(await response.arrayBuffer()),
      ext,
      type,
    };
    imageAssetCache.set(normalized, asset);
    return asset;
  } catch (error) {
    const asset = {
      ok: false,
      src: normalized,
      reason: `下载失败：${error.message}`,
    };
    imageAssetCache.set(normalized, asset);
    return asset;
  }
}

function getImageAssetSync(src) {
  const normalized = normalizeImageSource(src);
  if (imageAssetCache.has(normalized)) {
    return imageAssetCache.get(normalized);
  }
  if (isRemoteImageSource(normalized)) {
    return {
      ok: false,
      src: normalized,
      reason: '远程图片未预加载',
    };
  }
  return loadLocalImageAssetSync(normalized);
}

async function preloadImageAssets(parsedBlocks) {
  const sources = new Set();

  parsedBlocks.forEach((block) => {
    if (block.type === 'image' && block.src) {
      sources.add(normalizeImageSource(block.src));
      return;
    }

    if (block.type !== 'table') {
      return;
    }

    normalizeTableRows(block.lines).forEach((row) => {
      row.forEach((cell) => {
        splitByHtmlBreaks(cell).forEach((segment) => {
          const image = parseMarkdownImage(segment.trim());
          if (image && image.src) {
            sources.add(normalizeImageSource(image.src));
          }
        });
      });
    });
  });

  const remoteSources = Array.from(sources).filter((src) => isRemoteImageSource(src));
  if (!remoteSources.length) {
    return;
  }

  console.log(`[1.5/3] 正在预取远程图片：${remoteSources.length} 个`);
  await Promise.all(remoteSources.map((src) => loadRemoteImageAsset(src)));
}

function normalizeFenceLang(info) {
  return String(info || '').trim().split(/\s+/)[0].toLowerCase();
}

function normalizeDiagramLang(info) {
  const lang = normalizeFenceLang(info);
  if (['mermaid', 'plantuml', 'puml', 'uml'].includes(lang)) {
    return lang === 'mermaid' ? 'mermaid' : 'plantuml';
  }
  return '';
}

function normalizeMermaidSource(source) {
  const trimmed = String(source || '').trim();
  if (/^%%\{\s*init:/i.test(trimmed)) {
    return trimmed;
  }
  return `%%{init: {"theme":"base","themeVariables":{"fontFamily":"Times New Roman"}} }%%\n${trimmed}`;
}

function normalizePlantUmlSource(source) {
  const prefix = [
    'skinparam backgroundColor white',
    'skinparam defaultFontName "Times New Roman"',
    'scale 2',
  ].join('\n');
  const trimmed = String(source || '').trim();

  if (/^@startuml\b/i.test(trimmed)) {
    return trimmed.replace(/^@startuml\b[^\n]*\n?/i, (match) => `${match}${prefix}\n`);
  }

  return `@startuml\n${prefix}\n${trimmed}\n@enduml`;
}

function ensureDiagramTempDir() {
  const baseDir = path.join(workspaceRoot, '.workbuddy', 'tmp-diagrams');
  fs.mkdirSync(baseDir, { recursive: true });
  return fs.mkdtempSync(path.join(baseDir, 'render-'));
}

async function renderDiagramViaKroki(diagramType, source) {
  const response = await fetch(`${krokiBaseUrl}/${diagramType}/png`, {
    method: 'POST',
    headers: {
      'Content-Type': 'text/plain; charset=utf-8',
    },
    body: source,
  });

  if (!response.ok) {
    const detail = (await response.text()).slice(0, 300);
    throw new Error(`Kroki 渲染 ${diagramType} 失败：HTTP ${response.status}${detail ? ` - ${detail}` : ''}`);
  }

  return Buffer.from(await response.arrayBuffer());
}

async function renderDiagramBlocks(parsedBlocks) {
  const diagramIndexes = parsedBlocks
    .map((block, index) => ({ block, index }))
    .filter(({ block }) => block.type === 'codeblock' && normalizeDiagramLang(block.lang));

  if (!diagramIndexes.length) {
    return null;
  }

  const tempDir = ensureDiagramTempDir();
  console.log(`[1.25/3] 正在渲染图示代码块：${diagramIndexes.length} 个`);

  for (let seq = 0; seq < diagramIndexes.length; seq += 1) {
    const { block, index } = diagramIndexes[seq];
    const diagramLang = normalizeDiagramLang(block.lang);
    const source = diagramLang === 'mermaid'
      ? normalizeMermaidSource(block.code)
      : normalizePlantUmlSource(block.code);
    const imageBuffer = await renderDiagramViaKroki(diagramLang, source);
    const tempImagePath = path.join(tempDir, `${String(seq + 1).padStart(3, '0')}-${diagramLang}.png`);
    fs.writeFileSync(tempImagePath, imageBuffer);
    parsedBlocks[index] = {
      type: 'image',
      alt: `${diagramLang} diagram`,
      src: tempImagePath,
    };
  }

  return tempDir;
}

function cleanupTempDir(tempDir) {
  if (!tempDir) {
    return;
  }
  fs.rmSync(tempDir, { recursive: true, force: true });
}

function border(size = 4, color = '000000') {
  return { style: BorderStyle.SINGLE, size, color };
}

function allBorders(size = 4, color = '000000') {
  const b = border(size, color);
  return { top: b, bottom: b, left: b, right: b };
}

function sanitizeId(text) {
  const encoded = Array.from(String(text)).map((char) => (
    /[a-zA-Z0-9_]/.test(char)
      ? char
      : `_u${char.codePointAt(0).toString(16)}`
  )).join('');

  return encoded
    .replace(/^_+/, '')
    .slice(0, 80) || 'anchor';
}

let bookmarkNumericId = 1;

function nextBookmarkNumericId() {
  const current = bookmarkNumericId;
  bookmarkNumericId += 1;
  return current;
}

function headingBookmark(text) {
  return sanitizeId(`heading_${text}`);
}

function makeBookmarkRuns(id, child) {
  const bookmarkId = sanitizeId(id);
  const numericId = nextBookmarkNumericId();
  return [
    new TextRun({
      children: [new ImportedXmlComponent('w:bookmarkStart', {
        'w:id': String(numericId),
        'w:name': bookmarkId,
      })],
    }),
    child,
    new TextRun({
      children: [new ImportedXmlComponent('w:bookmarkEnd', {
        'w:id': String(numericId),
      })],
    }),
  ];
}

function normalSpacerParagraph() {
  return new Paragraph({
    spacing: lineSpacing(0, 0),
    indent: { left: 0, firstLine: 0 },
    children: [makeRun(' ', { size: SIZE.XIAO_SI })],
  });
}

const referenceMap = {};
const referenceLineIndexes = new Set();
lines.forEach((line, index) => {
  const match = line.match(/^\[\^(\d+)\]:\s*(.+)$/);
  if (match) {
    referenceMap[match[1]] = match[2].trim();
    referenceLineIndexes.add(index);
  }
});

const citationAnchorMap = {};
const citationCounter = {};
lines.forEach((line, index) => {
  if (referenceLineIndexes.has(index)) {
    return;
  }
  const regex = /\[\^(\d+)\]/g;
  let match;
  while ((match = regex.exec(line)) !== null) {
    const key = match[1];
    citationCounter[key] = (citationCounter[key] || 0) + 1;
    citationAnchorMap[key] = citationAnchorMap[key] || [];
    citationAnchorMap[key].push(sanitizeId(`cite_${key}_${citationCounter[key]}`));
  }
});

let listBlockCounter = 0;
function parseBlocks(sourceLines) {
  const blocks = [];
  let i = 0;

  while (i < sourceLines.length) {
    const line = sourceLines[i];

    if (referenceLineIndexes.has(i)) {
      i += 1;
      continue;
    }

    if (/^#\s+/.test(line)) {
      blocks.push({ type: 'title', text: line.replace(/^#\s+/, '').trim() });
      i += 1;
      continue;
    }

    if (/^##\s+/.test(line)) {
      blocks.push({ type: 'heading1', text: line.replace(/^##\s+/, '').trim() });
      i += 1;
      continue;
    }

    if (/^###\s+/.test(line)) {
      blocks.push({ type: 'heading2', text: line.replace(/^###\s+/, '').trim() });
      i += 1;
      continue;
    }

    if (/^####\s+/.test(line)) {
      blocks.push({ type: 'heading3', text: line.replace(/^####\s+/, '').trim() });
      i += 1;
      continue;
    }

    if (/^```/.test(line)) {
      const info = line.replace(/^```/, '').trim();
      const codeLines = [];
      i += 1;
      while (i < sourceLines.length && !/^```\s*$/.test(sourceLines[i])) {
        codeLines.push(sourceLines[i]);
        i += 1;
      }
      i += 1;
      blocks.push({ type: 'codeblock', lang: normalizeFenceLang(info), code: codeLines.join('\n') });
      continue;
    }

    if (line.trim() === '$$') {
      const formulaLines = [];
      i += 1;
      while (i < sourceLines.length && sourceLines[i].trim() !== '$$') {
        formulaLines.push(sourceLines[i]);
        i += 1;
      }
      i += 1;
      blocks.push({ type: 'formula', content: formulaLines.join('\n') });
      continue;
    }

    if (/^\|/.test(line)) {
      const tableLines = [];
      while (i < sourceLines.length && /^\|/.test(sourceLines[i])) {
        tableLines.push(sourceLines[i]);
        i += 1;
      }
      blocks.push({ type: 'table', lines: tableLines });
      continue;
    }

    const image = parseMarkdownImage(line.trim());
    if (image) {
      blocks.push({ type: 'image', alt: image.alt, src: image.src });
      i += 1;
      continue;
    }

    if (/^(图|表)\s*\d+/.test(line.trim())) {
      blocks.push({ type: 'caption', text: line.trim() });
      i += 1;
      continue;
    }

    const listMatch = line.match(/^(\s*)([-*]|\d+\.)\s+(.*)$/);
    if (listMatch) {
      const items = [];
      listBlockCounter += 1;
      while (i < sourceLines.length) {
        const current = sourceLines[i];
        const currentMatch = current.match(/^(\s*)([-*]|\d+\.)\s+(.*)$/);
        if (!currentMatch) {
          break;
        }
        const indentSpaces = currentMatch[1].replace(/\t/g, '    ').length;
        items.push({
          text: currentMatch[3],
          level: Math.min(Math.floor(indentSpaces / 2), 3),
          ordered: /\d+\./.test(currentMatch[2]),
        });
        i += 1;
      }
      blocks.push({ type: 'list', listId: listBlockCounter, items });
      continue;
    }

    if (line.trim() === '') {
      i += 1;
      continue;
    }

    blocks.push({ type: 'paragraph', text: line });
    i += 1;
  }

  return blocks;
}

const blocks = parseBlocks(lines);
console.log(`[1/3] Markdown 解析完成：${blocks.length} 个内容块`);
const renderCitationCounter = {};

function nextCitationAnchor(referenceKey) {
  renderCitationCounter[referenceKey] = (renderCitationCounter[referenceKey] || 0) + 1;
  return sanitizeId(`cite_${referenceKey}_${renderCitationCounter[referenceKey]}`);
}

function referenceBookmark(referenceKey) {
  return sanitizeId(`ref_${referenceKey}`);
}

function parseInlineRuns(text, baseSize = SIZE.XIAO_SI, opts = {}) {
  const runs = [];
  const baseFont = opts.font || { ascii: FONT_EN, eastAsia: FONT_CN, hAnsi: FONT_EN };
  const regex = /\*\*([^*]+)\*\*|\*([^*]+)\*|`([^`]+)`|\[\^(\d+)\]|<br\s*\/?>|\$([^$]+)\$/gi;
  let lastIndex = 0;
  let match;

  while ((match = regex.exec(text)) !== null) {
    if (match.index > lastIndex) {
      runs.push(makeRun(text.slice(lastIndex, match.index), {
        size: baseSize,
        bold: opts.bold || false,
        italics: opts.italics || false,
        font: baseFont,
      }));
    }

    if (match[1] !== undefined) {
      runs.push(makeRun(match[1], { size: baseSize, bold: true, font: baseFont }));
    } else if (match[2] !== undefined) {
      runs.push(makeRun(match[2], {
        size: baseSize,
        bold: opts.bold || false,
        italics: true,
        font: baseFont,
      }));
    } else if (match[3] !== undefined) {
      runs.push(makeRun(match[3], {
        size: baseSize,
        bold: opts.bold || false,
        font: { ascii: FONT_EN, eastAsia: FONT_CN, hAnsi: FONT_EN },
      }));
    } else if (match[4] !== undefined) {
      const key = match[4];
      const citeAnchor = nextCitationAnchor(key);
      runs.push(...makeBookmarkRuns(citeAnchor, makeRun('', { size: baseSize })));
      runs.push(new InternalHyperlink({
        anchor: referenceBookmark(key),
        children: [
          makeRun(`[${key}]`, {
            size: SIZE.XIAO_WU,
            superScript: true,
            color: '000000',
          }),
        ],
      }));
    } else if (/^<br/i.test(match[0])) {
      runs.push(makeRun('', { size: baseSize, break: 1 }));
    } else if (match[5] !== undefined) {
      runs.push(new DocxMath({ children: formulaToMathChildren(match[5]) }));
    }

    lastIndex = match.index + match[0].length;
  }

  if (lastIndex < text.length) {
    runs.push(makeRun(text.slice(lastIndex), {
      size: baseSize,
      bold: opts.bold || false,
      italics: opts.italics || false,
      font: baseFont,
    }));
  }

  return runs.length ? runs : [makeRun(text, { size: baseSize })];
}

const FORMULA_SYMBOL_MAP = {
  '\\times': '×',
  '\\oplus': '⊕',
  '\\neq': '≠',
  '\\varnothing': '∅',
  '\\in': '∈',
  '\\notin': '∉',
  '\\Leftrightarrow': '⇔',
  '\\leq': '≤',
  '\\geq': '≥',
  '\\cdot': '·',
};

function noBorder() {
  return { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
}

function noBorders() {
  const b = noBorder();
  return { top: b, bottom: b, left: b, right: b, insideH: b, insideV: b };
}

function parseMathArgument(state) {
  while (state.index < state.text.length && /\s/.test(state.text[state.index])) {
    state.index += 1;
  }

  if (state.text[state.index] === '{') {
    state.index += 1;
    const group = parseMathComponents(state, '}');
    if (state.text[state.index] === '}') {
      state.index += 1;
    }
    return group;
  }

  return parseSingleMathAtom(state);
}

function parseSingleMathAtom(state) {
  if (state.index >= state.text.length) {
    return [new MathRun('')];
  }

  const current = state.text[state.index];

  if (current === '\\') {
    if (state.text.startsWith('\\\\', state.index)) {
      state.index += 2;
      return [new MathRun(' ')];
    }

    const escaped = state.text[state.index + 1];
    if (['_', '{', '}', '&', '#', '$', '%'].includes(escaped)) {
      state.index += 2;
      return [new MathRun(escaped)];
    }
    if (escaped === ',') {
      state.index += 2;
      return [new MathRun(' ')];
    }

    const commandMatch = state.text.slice(state.index).match(/^\\[A-Za-z]+/);
    if (!commandMatch) {
      state.index += 1;
      return [new MathRun('\\')];
    }

    const command = commandMatch[0];
    state.index += command.length;

    if (command === '\\frac') {
      const numerator = parseMathArgument(state);
      const denominator = parseMathArgument(state);
      return [new MathFraction({ numerator, denominator })];
    }

    if (command === '\\sqrt') {
      const children = parseMathArgument(state);
      return [new MathRadical({ children })];
    }

    if (command === '\\text') {
      return parseMathArgument(state);
    }

    if (command === '\\left' || command === '\\right') {
      return parseSingleMathAtom(state);
    }

    return [new MathRun(FORMULA_SYMBOL_MAP[command] || command.replace(/^\\/, ''))];
  }

  if (current === '{') {
    state.index += 1;
    const group = parseMathComponents(state, '}');
    if (state.text[state.index] === '}') {
      state.index += 1;
    }
    return group;
  }

  if (/[A-Za-z0-9.]/.test(current)) {
    const start = state.index;
    while (state.index < state.text.length && /[A-Za-z0-9.]/.test(state.text[state.index])) {
      state.index += 1;
    }
    return [new MathRun(state.text.slice(start, state.index))];
  }

  if (/\s/.test(current)) {
    while (state.index < state.text.length && /\s/.test(state.text[state.index])) {
      state.index += 1;
    }
    return [new MathRun(' ')];
  }

  state.index += 1;
  return [new MathRun(current)];
}

function attachMathScript(components, kind, scriptComponents) {
  const base = components.pop() || new MathRun(' ');
  if (kind === '_') {
    components.push(new MathSubScript({ children: [base], subScript: scriptComponents }));
    return;
  }
  components.push(new MathSuperScript({ children: [base], superScript: scriptComponents }));
}

function parseMathComponents(state, stopChar = null) {
  const components = [];

  while (state.index < state.text.length) {
    const current = state.text[state.index];
    if (stopChar && current === stopChar) {
      break;
    }

    if (current === '_' || current === '^') {
      const kind = current;
      state.index += 1;
      const scriptComponents = parseMathArgument(state);
      attachMathScript(components, kind, scriptComponents);
      continue;
    }

    components.push(...parseSingleMathAtom(state));
  }

  return components;
}

function formulaToMathChildren(formulaText) {
  const normalized = formulaText
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  if (!normalized) {
    return [new MathRun(' ')];
  }

  const state = { text: normalized, index: 0 };
  return parseMathComponents(state);
}

function escapeXml(text) {
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function ommlRun(text) {
  return `<m:r><m:t xml:space="preserve">${escapeXml(text)}</m:t></m:r>`;
}

function parseOmmlArgument(state) {
  while (state.index < state.text.length && /\s/.test(state.text[state.index])) {
    state.index += 1;
  }

  if (state.text[state.index] === '{') {
    state.index += 1;
    const group = parseOmmlComponents(state, '}');
    if (state.text[state.index] === '}') {
      state.index += 1;
    }
    return group;
  }

  return parseSingleOmmlAtom(state);
}

function parseSingleOmmlAtom(state) {
  if (state.index >= state.text.length) {
    return [ommlRun('')];
  }

  const current = state.text[state.index];

  if (current === '\\') {
    if (state.text.startsWith('\\\\', state.index)) {
      state.index += 2;
      return [ommlRun(' ')];
    }

    const escaped = state.text[state.index + 1];
    if (['_', '{', '}', '&', '#', '$', '%'].includes(escaped)) {
      state.index += 2;
      return [ommlRun(escaped)];
    }
    if (escaped === ',') {
      state.index += 2;
      return [ommlRun(' ')];
    }

    const commandMatch = state.text.slice(state.index).match(/^\\[A-Za-z]+/);
    if (!commandMatch) {
      state.index += 1;
      return [ommlRun('\\')];
    }

    const command = commandMatch[0];
    state.index += command.length;

    if (command === '\\frac') {
      const numerator = parseOmmlArgument(state).join('');
      const denominator = parseOmmlArgument(state).join('');
      return [`<m:f><m:num>${numerator}</m:num><m:den>${denominator}</m:den></m:f>`];
    }

    if (command === '\\sqrt') {
      const children = parseOmmlArgument(state).join('');
      return [`<m:rad><m:radPr><m:degHide/></m:radPr><m:deg/><m:e>${children}</m:e></m:rad>`];
    }

    if (command === '\\text') {
      return parseOmmlArgument(state);
    }

    if (command === '\\left' || command === '\\right') {
      return parseSingleOmmlAtom(state);
    }

    return [ommlRun(FORMULA_SYMBOL_MAP[command] || command.replace(/^\\/, ''))];
  }

  if (current === '{') {
    state.index += 1;
    const group = parseOmmlComponents(state, '}');
    if (state.text[state.index] === '}') {
      state.index += 1;
    }
    return group;
  }

  if (/[A-Za-z0-9.]/.test(current)) {
    const start = state.index;
    while (state.index < state.text.length && /[A-Za-z0-9.]/.test(state.text[state.index])) {
      state.index += 1;
    }
    return [ommlRun(state.text.slice(start, state.index))];
  }

  if (/\s/.test(current)) {
    while (state.index < state.text.length && /\s/.test(state.text[state.index])) {
      state.index += 1;
    }
    return [ommlRun(' ')];
  }

  state.index += 1;
  return [ommlRun(current)];
}

function attachOmmlScript(components, kind, scriptComponents) {
  const base = components.pop() || ommlRun(' ');
  const script = scriptComponents.join('');
  if (kind === '_') {
    components.push(`<m:sSub><m:sSubPr/><m:e>${base}</m:e><m:sub>${script}</m:sub></m:sSub>`);
    return;
  }
  components.push(`<m:sSup><m:sSupPr/><m:e>${base}</m:e><m:sup>${script}</m:sup></m:sSup>`);
}

function parseOmmlComponents(state, stopChar = null) {
  const components = [];

  while (state.index < state.text.length) {
    const current = state.text[state.index];
    if (stopChar && current === stopChar) {
      break;
    }

    if (current === '_' || current === '^') {
      const kind = current;
      state.index += 1;
      const scriptComponents = parseOmmlArgument(state);
      attachOmmlScript(components, kind, scriptComponents);
      continue;
    }

    components.push(...parseSingleOmmlAtom(state));
  }

  return components;
}

function formulaToOmml(formulaText) {
  const normalized = formulaText
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  if (!normalized) {
    return ommlRun(' ');
  }

  const state = { text: normalized, index: 0 };
  return parseOmmlComponents(state).join('');
}

function importedXmlRoot(xml) {
  const wrapper = ImportedXmlComponent.fromXmlString(xml);
  return wrapper.root[0];
}

/**
 * 从公式文本中提取末尾的 LaTeX 注释形式的公式编号标签。
 * 约定：在 $$ 块内任意一行末尾写 % 式（X-Y） 即标记该公式编号。
 * 示例：`Score = W \times S  % 式（5-2）`
 * 返回 { formulaText, label } ，label 为 null 表示无编号。
 */
function extractFormulaLabel(rawText) {
  // 从最后一行提取 `% 式（X-Y）` 或 `% (X-Y)`
  const labelRe = /\s*%\s*(式[（(]\d+-\d+[）)]|[（(]\d+-\d+[）)])\s*$/m;
  const match = rawText.match(labelRe);
  if (!match) {
    return { formulaText: rawText.trim(), label: null };
  }
  const label = match[1].startsWith('式') ? match[1] : `式${match[1]}`;
  const formulaText = rawText.replace(labelRe, '').trim();
  return { formulaText, label };
}

/**
 * 构建带编号的公式行：三列无框表格
 *   左列（空白占位，约60%正文宽）| 中列（公式，居中）| 右列（编号，右对齐）
 * 官方要求：公式居中，序号标注在该行最右侧。
 */
function buildNumberedFormulaRow(mathChildren, label) {
  // 列宽分配：左占位 / 公式主体 / 右编号
  const labelWidth = 1200;   // ~2.1cm 给编号
  const leftWidth = labelWidth;
  const midWidth = CONTENT_WIDTH - leftWidth - labelWidth;

  const makeNoBorderCell = (children, align) => new TableCell({
    width: { size: 0, type: WidthType.AUTO },
    borders: {
      top: noBorder(), bottom: noBorder(), left: noBorder(), right: noBorder(),
      insideH: noBorder(), insideV: noBorder(),
    },
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 0, bottom: 0, left: 0, right: 0 },
    children: [new Paragraph({
      alignment: align,
      spacing: { line: 360, lineRule: 'auto', before: 60, after: 60 },
      indent: { left: 0, firstLine: 0 },
      children,
    })],
  });

  return new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: [leftWidth, midWidth, labelWidth],
    borders: {
      top: noBorder(), bottom: noBorder(), left: noBorder(), right: noBorder(),
      insideH: noBorder(), insideV: noBorder(),
    },
    alignment: AlignmentType.CENTER,
    rows: [new TableRow({
      children: [
        makeNoBorderCell([makeRun(' ', { size: SIZE.XIAO_SI })], AlignmentType.LEFT),
        makeNoBorderCell([new DocxMath({ children: mathChildren })], AlignmentType.CENTER),
        makeNoBorderCell(
          [makeRun(label, { size: SIZE.XIAO_SI })],
          AlignmentType.RIGHT,
        ),
      ],
    })],
  });
}

function buildMathParagraph(formulaText, label = null) {
  const mathChildren = formulaToMathChildren(formulaText);
  if (label) {
    return buildNumberedFormulaRow(mathChildren, label);
  }
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: lineSpacing(60, 60),
    indent: { left: 0, firstLine: 0 },
    children: [new DocxMath({ children: mathChildren })],
  });
}

function buildMathCellParagraph(formulaText, alignment = AlignmentType.LEFT) {
  const normalized = (formulaText || '').trim();
  return new Paragraph({
    alignment,
    spacing: lineSpacing(0, 0, 300),
    indent: { left: 0, firstLine: 0 },
    children: normalized
      ? [new DocxMath({ children: formulaToMathChildren(normalized) })]
      : [makeRun(' ', { size: SIZE.XIAO_SI })],
  });
}

function parseCasesFormula(formulaText) {
  const match = formulaText.trim().match(/^([\s\S]*?)\\begin\{cases\}([\s\S]*?)\\end\{cases\}$/);
  if (!match) {
    return null;
  }

  const prefix = match[1].trim();
  const rows = match[2]
    .split(/\\\\/)
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line) => {
      const [expression, condition = ''] = line.split(/\s*&\s*/);
      return {
        expression: expression.trim(),
        condition: condition.trim(),
      };
    });

  return rows.length ? { prefix, rows } : null;
}

function formulaVisualText(formulaText) {
  return formulaText
    .replace(/\\notin/g, '∉')
    .replace(/\\in/g, '∈')
    .replace(/\\times/g, '×')
    .replace(/\\Leftrightarrow/g, '⇔')
    .replace(/\\[A-Za-z]+/g, '')
    .replace(/[{}_]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function buildCaseFormulaTable(formulaText, label = null) {
  const parsed = parseCasesFormula(formulaText);
  if (!parsed) {
    return [buildMathParagraph(formulaText, label)];
  }

  const expressionLengths = parsed.rows.map((row) => measureVisualLength(formulaVisualText(row.expression)));
  const maxExpressionLength = Math.max(...expressionLengths, 0);

  const rowXml = parsed.rows.map((row, index) => {
    const expressionLength = expressionLengths[index];
    const gapCount = Math.max(3, Math.ceil((maxExpressionLength - expressionLength) / 1.5) + 3);
    const gap = ' '.repeat(gapCount);
    return `<m:e>${formulaToOmml(row.expression)}${ommlRun(gap)}${formulaToOmml(row.condition)}</m:e>`;
  }).join('');

  const prefixXml = parsed.prefix ? `${formulaToOmml(parsed.prefix)}${ommlRun(' ')}` : '';
  const xml = `<m:oMath>${prefixXml}<m:d><m:dPr><m:begChr m:val="{"/><m:endChr m:val=""/></m:dPr><m:e><m:eqArr>${rowXml}</m:eqArr></m:e></m:d></m:oMath>`;

  if (label) {
    const labelWidth = 1200;
    const leftWidth = labelWidth;
    const midWidth = CONTENT_WIDTH - leftWidth - labelWidth;
    const makeNoBorderCell = (children, align) => new TableCell({
      width: { size: 0, type: WidthType.AUTO },
      borders: {
        top: noBorder(), bottom: noBorder(), left: noBorder(), right: noBorder(),
        insideH: noBorder(), insideV: noBorder(),
      },
      verticalAlign: VerticalAlign.CENTER,
      margins: { top: 0, bottom: 0, left: 0, right: 0 },
      children: [new Paragraph({
        alignment: align,
        spacing: { line: 360, lineRule: 'auto', before: 60, after: 60 },
        indent: { left: 0, firstLine: 0 },
        children,
      })],
    });
    return [new Table({
      width: { size: CONTENT_WIDTH, type: WidthType.DXA },
      columnWidths: [leftWidth, midWidth, labelWidth],
      borders: {
        top: noBorder(), bottom: noBorder(), left: noBorder(), right: noBorder(),
        insideH: noBorder(), insideV: noBorder(),
      },
      alignment: AlignmentType.CENTER,
      rows: [new TableRow({
        children: [
          makeNoBorderCell([makeRun(' ', { size: SIZE.XIAO_SI })], AlignmentType.LEFT),
          makeNoBorderCell([importedXmlRoot(xml)], AlignmentType.CENTER),
          makeNoBorderCell([makeRun(label, { size: SIZE.XIAO_SI })], AlignmentType.RIGHT),
        ],
      })],
    })];
  }

  return [new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: lineSpacing(60, 60),
    indent: { left: 0, firstLine: 0 },
    children: [importedXmlRoot(xml)],
  })];
}

function makeHeadingRun(text, opts = {}) {
  return new TextRun({
    text,
    size: opts.size || SIZE.XIAO_SI,
    bold: opts.bold !== false,
    color: '000000',
    font: { ascii: FONT_EN, eastAsia: FONT_HEI, hAnsi: FONT_EN },
  });
}

function buildHeading(text, level, size, alignment, indent) {
  // 不再依赖 Word 内置 Heading1/2/3。
  // 用户机器上的 Word 会把内置 heading 样式重新套回主题字体（等线）和主题色（蓝色），
  // 尤其在标题又被自动目录引用时更容易触发。这里改用自定义段落样式供正文与 TOC 同时引用。
  const styleId = BODY_HEADING_STYLES[level] || BODY_HEADING_STYLES[1];
  return new Paragraph({
    style: styleId,
    alignment,
    spacing: lineSpacing(180, 120),
    indent: indent || { left: 0, firstLine: 0 },
    children: makeBookmarkRuns(headingBookmark(text), new TextRun({ text })),
  });
}

function buildTitle(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: lineSpacing(240, 240),
    children: [makeRun(text, { size: SIZE.SAN_HAO, bold: true })],
  });
}

function buildFrontMatterHeading(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: lineSpacing(180, 120),
    indent: { left: 0, firstLine: 0 },
    children: makeBookmarkRuns(`heading_${text}`, makeHeadingRun(text, { size: SIZE.SAN_HAO, bold: true })),
  });
}

function pageBreakParagraph() {
  return new Paragraph({ children: [new PageBreak()] });
}

function shouldStartOnNewPage(block, context = {}) {
  if (!context.previousBlock) {
    return false;
  }

  if (block.type === 'title') {
    return true;
  }

  return block.type === 'heading1' && !['摘要', 'ABSTRACT'].includes(block.text);
}

function buildTableOfContentsBlock() {
  return new TableOfContents('', {
    hyperlink: true,
    stylesWithLevels: [
      { styleName: BODY_HEADING_STYLES[1], level: 1 },
      { styleName: BODY_HEADING_STYLES[2], level: 2 },
      { styleName: BODY_HEADING_STYLES[3], level: 3 },
    ],
  });
}

function buildDefaultHeader() {
  return new Header({
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: lineSpacing(0, 0, 240),
        children: [makeRun('北京邮电大学本科毕业设计（论文）', {
          size: SIZE.XIAO_WU,
          font: { ascii: FONT_EN, eastAsia: FONT_CN, hAnsi: FONT_EN },
        })],
      }),
    ],
  });
}

function buildPageNumberFooter() {
  return new Footer({
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: lineSpacing(0, 0, 240),
        indent: { left: 0, firstLine: 0 },
        children: [new TextRun({
          children: [PageNumber.CURRENT],
        })],
      }),
    ],
  });
}

function splitByHtmlBreaks(text) {
  return text.split(/<br\s*\/?>/i);
}

function isMarkdownImage(text) {
  return Boolean(parseMarkdownImage(text));
}

function parseMarkdownImage(text) {
  const match = (text || '').trim().match(/^!\[([^\]]*)\]\((.+)\)$/);
  if (!match) {
    return null;
  }

  let src = match[2].trim();
  if (src.startsWith('<')) {
    const closingIndex = src.lastIndexOf('>');
    if (closingIndex > 0) {
      src = src.slice(1, closingIndex).trim();
    }
  }

  return {
    alt: match[1].trim(),
    src: normalizeImageSource(src),
  };
}

function buildTableCellImageParagraph(image, columnWidth) {
  const asset = getImageAssetSync(image.src);
  if (!asset.ok) {
    return buildImagePlaceholderParagraph(`[图像占位：${image.alt || image.src}]`, lineSpacing(0, 0, 240));
  }

  const dimensions = getImageDimensions(asset.buffer, asset.ext);
  const usableWidth = Math.max(120, Math.floor((columnWidth - 240) / 15));
  const scaled = scaleImageSizeToFit(dimensions.width, dimensions.height, usableWidth, 300);

  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: lineSpacing(0, 0, 240),
    indent: zeroIndent(),
    children: [new ImageRun({
      type: asset.type,
      data: asset.buffer,
      transformation: scaled,
      altText: {
        title: image.alt || image.src,
        description: image.alt || image.src,
        name: image.alt || image.src,
      },
    })],
  });
}

function buildTableCellParagraphs(text, isHeader, columnWidth) {
  const paragraphs = [];

  splitByHtmlBreaks(text).forEach((segment) => {
    const trimmed = segment.trim();

    if (!trimmed) {
      if (paragraphs.length) {
        paragraphs.push(new Paragraph({
          style: 'TableCell',
          spacing: lineSpacing(0, 0, 180),
          indent: zeroIndent(),
          children: [makeRun(' ', { size: SIZE.XIAO_SI })],
        }));
      }
      return;
    }

    const image = parseMarkdownImage(trimmed);
    if (image) {
      paragraphs.push(buildTableCellImageParagraph(image, columnWidth));
      return;
    }

    if (/^图\s*\d+/.test(trimmed)) {
      paragraphs.push(new Paragraph({
        style: 'TableCell',
        alignment: AlignmentType.CENTER,
        spacing: lineSpacing(0, 0, 240),
        indent: zeroIndent(),
        children: parseInlineRuns(trimmed, SIZE.XIAO_SI),
      }));
      return;
    }

    paragraphs.push(new Paragraph({
      style: 'TableCell',
      alignment: isHeader ? AlignmentType.CENTER : AlignmentType.LEFT,
      spacing: lineSpacing(0, 0, 300),
      indent: zeroIndent(),
      children: parseInlineRuns(trimmed, SIZE.XIAO_SI, { bold: isHeader }),
    }));
  });

  return paragraphs.length ? paragraphs : [new Paragraph({
    style: 'TableCell',
    spacing: lineSpacing(0, 0, 300),
    indent: zeroIndent(),
    children: [makeRun(' ', { size: SIZE.XIAO_SI })],
  })];
}

function cleanCellText(text) {
  return text
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/^!\[[^\]]*\]\([^\)]+\)$/gm, '')
    .replace(/\*\*/g, '')
    .replace(/`/g, '')
    .replace(/\[\^(\d+)\]/g, '[$1]')
    .replace(/\$([^$]+)\$/g, '$1');
}

function measureVisualLength(text) {
  let total = 0;
  for (const ch of text) {
    if (/\s/.test(ch)) {
      total += 0.5;
    } else if (/[\u4e00-\u9fff]/.test(ch)) {
      total += 2;
    } else {
      total += 1;
    }
  }
  return total;
}

function calcTableColumnWidths(rows) {
  const colCount = rows[0].length;
  const weights = new Array(colCount).fill(1);

  rows.forEach((row, rowIndex) => {
    row.forEach((cell, columnIndex) => {
      const linesOfCell = cleanCellText(cell).split(/\n+/).filter(Boolean);
      const longestLine = linesOfCell.length
        ? Math.max(...linesOfCell.map((segment) => measureVisualLength(segment)))
        : 1;
      const headerBoost = rowIndex === 0 ? 2.2 : 1;
      weights[columnIndex] = Math.max(weights[columnIndex], longestLine * headerBoost + (rowIndex === 0 ? 3 : 0));
    });
  });

  const preferredMinWidth = colCount <= 3 ? 1600 : colCount <= 5 ? 1200 : colCount <= 7 ? 920 : 720;
  const minWidth = Math.min(preferredMinWidth, Math.max(520, Math.floor(CONTENT_WIDTH / colCount)));
  const totalWeight = weights.reduce((sum, current) => sum + current, 0) || 1;

  const widths = weights.map((weight) => Math.max(minWidth, Math.round((weight / totalWeight) * CONTENT_WIDTH)));

  let overflow = widths.reduce((sum, current) => sum + current, 0) - CONTENT_WIDTH;
  while (overflow > 0) {
    const adjustableIndexes = widths
      .map((value, index) => ({ value, index }))
      .filter(({ value }) => value > minWidth)
      .sort((a, b) => b.value - a.value)
      .map(({ index }) => index);

    if (!adjustableIndexes.length) {
      break;
    }

    for (const index of adjustableIndexes) {
      if (overflow <= 0) {
        break;
      }
      const reducible = widths[index] - minWidth;
      if (reducible <= 0) {
        continue;
      }
      const reduceBy = Math.min(reducible, Math.max(1, Math.ceil(overflow / adjustableIndexes.length)));
      widths[index] -= reduceBy;
      overflow -= reduceBy;
    }
  }

  const diff = CONTENT_WIDTH - widths.reduce((sum, current) => sum + current, 0);
  widths[widths.length - 1] += diff;
  return widths;
}

function normalizeTableRows(tableLines) {
  const rows = tableLines
    .filter((line) => !/^\|[\s\-:|]+\|\s*$/.test(line))
    .map((line) => line.replace(/^\|/, '').replace(/\|\s*$/, '').split('|').map((cell) => cell.trim()));

  const filteredRows = rows.filter((row) => !row.every((cell) => /^:?-+:?$/.test(cell) || cell === ''));
  const maxCols = filteredRows.reduce((max, row) => Math.max(max, row.length), 0);

  return filteredRows.map((row) => {
    const normalized = row.slice();
    while (normalized.length < maxCols) {
      normalized.push('');
    }
    return normalized;
  });
}

function buildTable(tableLines) {
  const rows = normalizeTableRows(tableLines);
  if (!rows.length) {
    return new Paragraph({ children: [] });
  }

  const columnWidths = calcTableColumnWidths(rows);
  const tableRows = rows.map((row, rowIndex) => {
    const isHeader = rowIndex === 0;
    return new TableRow({
      tableHeader: isHeader,
      children: row.map((cellText, columnIndex) => new TableCell({
        borders: allBorders(),
        width: { size: columnWidths[columnIndex], type: WidthType.DXA },
        verticalAlign: VerticalAlign.CENTER,
        margins: { top: 80, bottom: 80, left: 60, right: 60 },
        shading: isHeader
          ? { fill: 'F2F2F2', type: ShadingType.CLEAR }
          : { fill: 'FFFFFF', type: ShadingType.CLEAR },
        children: buildTableCellParagraphs(cellText, isHeader, columnWidths[columnIndex]),
      })),
    });
  });

  return new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths,
    rows: tableRows,
    borders: {
      top: border(),
      bottom: border(),
      left: border(),
      right: border(),
      insideH: border(),
      insideV: border(),
    },
    alignment: AlignmentType.CENTER,
  });
}

function getImageDimensions(buffer, ext) {
  if (ext === 'png' && buffer.length >= 24) {
    return {
      width: buffer.readUInt32BE(16),
      height: buffer.readUInt32BE(20),
    };
  }

  if ((ext === 'jpg' || ext === 'jpeg') && buffer.length > 4) {
    let offset = 2;
    while (offset < buffer.length) {
      if (buffer[offset] !== 0xff) {
        offset += 1;
        continue;
      }
      const marker = buffer[offset + 1];
      const size = buffer.readUInt16BE(offset + 2);
      const isSOF = [0xc0, 0xc1, 0xc2, 0xc3, 0xc5, 0xc6, 0xc7, 0xc9, 0xca, 0xcb, 0xcd, 0xce, 0xcf].includes(marker);
      if (isSOF) {
        return {
          height: buffer.readUInt16BE(offset + 5),
          width: buffer.readUInt16BE(offset + 7),
        };
      }
      offset += 2 + size;
    }
  }

  return { width: 900, height: 520 };
}

function scaleImageSizeToFit(width, height, maxWidth, maxHeight) {
  const ratio = Math.min(maxWidth / width, maxHeight / height, 1);
  return {
    width: Math.max(1, Math.round(width * ratio)),
    height: Math.max(1, Math.round(height * ratio)),
  };
}

function scaleImageSize(width, height) {
  return scaleImageSizeToFit(width, height, 520, 720);
}

function buildReferenceParagraphs() {
  return Object.keys(referenceMap)
    .sort((a, b) => Number(a) - Number(b))
    .map((key) => {
      const children = [
        ...makeBookmarkRuns(referenceBookmark(key), makeRun(`[${key}] `, { size: SIZE.WU_HAO })),
        ...parseInlineRuns(referenceMap[key], SIZE.WU_HAO),
      ];
      const firstCitationAnchor = citationAnchorMap[key] && citationAnchorMap[key][0];
      if (firstCitationAnchor) {
        children.push(makeRun(' ', { size: SIZE.WU_HAO }));
        children.push(new InternalHyperlink({
          anchor: firstCitationAnchor,
          children: [makeRun('↩', { size: SIZE.XIAO_WU, color: '000000' })],
        }));
      }
      // 官方：参考文献内容五号宋体/TNR，行间距1.5倍
      return new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: lineSpacing(0, 0, 360),
        indent: { left: 420, hanging: 420 },
        children,
      });
    });
}

function buildNumberingConfig(parsedBlocks) {
  const configs = [];
  parsedBlocks
    .filter((block) => block.type === 'list')
    .forEach((block) => {
      const levels = [0, 1, 2, 3].map((level) => ({
        level,
        alignment: AlignmentType.LEFT,
        style: {
          paragraph: {
            indent: {
              left: 480 * (level + 1),
              hanging: 360,
            },
          },
        },
      }));

      if (block.items.some((item) => !item.ordered)) {
        configs.push({
          reference: `bullet_${block.listId}`,
          levels: levels.map((levelConfig) => ({
            ...levelConfig,
            format: LevelFormat.BULLET,
            text: '•',
          })),
        });
      }

      if (block.items.some((item) => item.ordered)) {
        configs.push({
          reference: `ordered_${block.listId}`,
          levels: levels.map((levelConfig) => ({
            ...levelConfig,
            format: LevelFormat.DECIMAL,
            text: levelConfig.level === 0
              ? '%1.'
              : levelConfig.level === 1
                ? '%1.%2.'
                : levelConfig.level === 2
                  ? '%1.%2.%3.'
                  : '%1.%2.%3.%4.',
          })),
        });
      }
    });

  return configs;
}

function isTableCaptionBlock(block) {
  return Boolean(block && block.type === 'caption' && /^表\s*\d+/.test(block.text));
}

function blockToElements(block, context = {}) {
  switch (block.type) {
    case 'title': {
      const elements = [];
      if (shouldStartOnNewPage(block, context)) {
        elements.push(pageBreakParagraph());
      }
      elements.push(buildTitle(block.text));
      return elements;
    }

    case 'heading1': {
      const elements = [];
      if (shouldStartOnNewPage(block, context)) {
        elements.push(pageBreakParagraph());
      }

      if (['摘要', 'ABSTRACT', '目录'].includes(block.text)) {
        elements.push(buildFrontMatterHeading(block.text));
        if (block.text === '目录') {
          elements.push(buildTableOfContentsBlock());
        }
        return elements;
      }

      const heading = buildHeading(block.text, 1, SIZE.SAN_HAO, AlignmentType.CENTER);
      if (block.text === '参考文献') {
        elements.push(heading, ...buildReferenceParagraphs());
        return elements;
      }
      elements.push(heading);
      return elements;
    }

    case 'heading2':
      // 官方：二级标题黑体四号，加粗，顶格，不缩进
      return [buildHeading(block.text, 2, SIZE.SI_HAO, AlignmentType.LEFT, zeroIndent())];

    case 'heading3':
      // 官方：三级标题黑体小四号，加粗，首行缩进2字符
      return [buildHeading(block.text, 3, SIZE.XIAO_SI, AlignmentType.LEFT, indentTwoChars())];

    case 'paragraph':
      return [new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: lineSpacing(0, 0),
        indent: indentTwoChars(),
        children: parseInlineRuns(block.text, SIZE.XIAO_SI),
      })];

    case 'list':
      return block.items.map((item) => new Paragraph({
        numbering: {
          reference: item.ordered ? `ordered_${block.listId}` : `bullet_${block.listId}`,
          level: item.level,
        },
        alignment: AlignmentType.JUSTIFIED,
        spacing: lineSpacing(0, 0),
        children: parseInlineRuns(item.text, SIZE.XIAO_SI),
      }));

    case 'codeblock':
      return block.code.split('\n').map((line) => new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { line: 300, lineRule: 'auto', before: 0, after: 0 },
        indent: { left: 240, firstLine: 0 },
        border: {
          left: { style: BorderStyle.SINGLE, size: 8, color: 'BFBFBF', space: 1 },
        },
        shading: { fill: 'F7F7F7', type: ShadingType.CLEAR },
        children: [makeRun(line || ' ', {
          size: SIZE.WU_HAO,
          font: { ascii: FONT_EN, eastAsia: FONT_CN, hAnsi: FONT_EN },
        })],
      }));

    case 'formula': {
      const { formulaText, label } = extractFormulaLabel(block.content);
      return buildCaseFormulaTable(formulaText, label);
    }

    case 'image': {
      const asset = getImageAssetSync(block.src);
      if (!asset.ok) {
        return [buildImagePlaceholderParagraph(`[图像占位：${block.alt || block.src}]`, lineSpacing(60, 60))];
      }

      const dimensions = getImageDimensions(asset.buffer, asset.ext);
      const scaled = scaleImageSize(dimensions.width, dimensions.height);

      return [new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: lineSpacing(60, 60),
        indent: { left: 0, firstLine: 0 },
        children: [new ImageRun({
          type: asset.type,
          data: asset.buffer,
          transformation: scaled,
          altText: {
            title: block.alt || block.src,
            description: block.alt || block.src,
            name: block.alt || block.src,
          },
        })],
      })];
    }

    case 'caption':
      return [new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: lineSpacing(0, context.nextBlock && context.nextBlock.type === 'table' ? 0 : 120),
        indent: { left: 0, firstLine: 0 },
        // 官方：图/表题注中文楷体五号，英文 Times New Roman 五号
        children: parseInlineRuns(block.text, SIZE.WU_HAO, { font: { ascii: FONT_EN, eastAsia: FONT_KAI, hAnsi: FONT_EN } }),
      })];

    case 'table': {
      const elements = [];
      if (!isTableCaptionBlock(context.previousBlock)) {
        elements.push(normalSpacerParagraph());
      }
      elements.push(buildTable(block.lines));
      elements.push(normalSpacerParagraph());
      return elements;
    }

    default:
      return [];
  }
}

async function main() {
  let diagramTempDir = null;
  try {
    diagramTempDir = await renderDiagramBlocks(blocks);
    await preloadImageAssets(blocks);

    const frontMatterElements = [];
  const bodyElements = [];
  let bodyStarted = false;

  blocks.forEach((block, index) => {
    const elements = blockToElements(block, {
      previousBlock: index > 0 ? blocks[index - 1] : null,
      nextBlock: index < blocks.length - 1 ? blocks[index + 1] : null,
    });

    const isBodyStartHeading = !bodyStarted
      && block.type === 'heading1'
      && /^第[一二三四五六七八九十百零〇两]+章\s+/.test(block.text);

    if (isBodyStartHeading) {
      bodyStarted = true;
      const normalizedElements = elements.slice();
      if (normalizedElements.length > 1) {
        normalizedElements.shift();
      }
      normalizedElements.forEach((element) => bodyElements.push(element));
    } else if (bodyStarted) {
      elements.forEach((element) => bodyElements.push(element));
    } else {
      elements.forEach((element) => frontMatterElements.push(element));
    }

    if ((index + 1) % 100 === 0 || index === blocks.length - 1) {
      console.log(`[2/3] 正在构建文档元素：${index + 1}/${blocks.length}`);
    }
  });

  console.log(`[2/3] 文档元素构建完成：前置 ${frontMatterElements.length} 个元素，正文 ${bodyElements.length} 个元素`);

  const doc = new Document({
    features: {
      updateFields: true,
    },
    numbering: {
      config: buildNumberingConfig(blocks),
    },
    styles: {
      default: {
        document: {
          run: {
            font: { ascii: FONT_EN, eastAsia: FONT_CN, hAnsi: FONT_EN },
            size: SIZE.XIAO_SI,
          },
          paragraph: {
            spacing: lineSpacing(),
            indent: { left: 0, firstLine: 0 },
          },
        },
      },
      paragraphStyles: [
        {
          // 显式定义 Normal 样式，彻底清零首行缩进
          id: 'Normal',
          name: 'Normal',
          quickFormat: true,
          run: {
            font: { ascii: FONT_EN, eastAsia: FONT_CN, hAnsi: FONT_EN },
            size: SIZE.XIAO_SI,
          },
          paragraph: {
            spacing: lineSpacing(),
            indent: { left: 0, firstLine: 0 },
          },
        },
        {
          // 表格单元格段落专用样式，与正文段落样式隔离，确保无任何首行/悬挂缩进
          id: 'TableCell',
          name: 'Table Cell',
          quickFormat: true,
          run: {
            font: { ascii: FONT_EN, eastAsia: FONT_CN, hAnsi: FONT_EN },
            size: SIZE.XIAO_SI,
          },
          paragraph: {
            spacing: lineSpacing(0, 0),
            indent: zeroIndent(),
          },
        },
        {
          id: BODY_HEADING_STYLES[1],
          name: BODY_HEADING_STYLES[1],
          basedOn: 'Normal',
          next: 'Normal',
          quickFormat: true,
          run: {
            size: SIZE.SAN_HAO,
            bold: true,
            font: { ascii: FONT_EN, eastAsia: FONT_HEI, hAnsi: FONT_EN },
            color: '000000',
          },
          paragraph: {
            alignment: AlignmentType.CENTER,
            spacing: lineSpacing(180, 120),
            indent: { left: 0, firstLine: 0 },
            outlineLevel: 0,
          },
        },
        {
          id: BODY_HEADING_STYLES[2],
          name: BODY_HEADING_STYLES[2],
          basedOn: 'Normal',
          next: 'Normal',
          quickFormat: true,
          run: {
            size: SIZE.SI_HAO,
            bold: true,
            font: { ascii: FONT_EN, eastAsia: FONT_HEI, hAnsi: FONT_EN },
            color: '000000',
          },
          paragraph: {
            alignment: AlignmentType.LEFT,
            spacing: lineSpacing(120, 90),
            indent: zeroIndent(),
            outlineLevel: 1,
          },
        },
        {
          id: BODY_HEADING_STYLES[3],
          name: BODY_HEADING_STYLES[3],
          basedOn: 'Normal',
          next: 'Normal',
          quickFormat: true,
          run: {
            size: SIZE.XIAO_SI,
            bold: true,
            font: { ascii: FONT_EN, eastAsia: FONT_HEI, hAnsi: FONT_EN },
            color: '000000',
          },
          paragraph: {
            alignment: AlignmentType.LEFT,
            spacing: lineSpacing(90, 60),
            indent: indentTwoChars(),
            outlineLevel: 2,
          },
        },
        {
          id: 'TOCHeading',
          name: 'TOC Heading',
          basedOn: 'Normal',
          quickFormat: true,
          run: {
            size: SIZE.SAN_HAO,
            bold: true,
            font: { ascii: FONT_EN, eastAsia: FONT_HEI, hAnsi: FONT_EN },
          },
          paragraph: {
            alignment: AlignmentType.CENTER,
            spacing: lineSpacing(0, 120, 300),
            indent: { left: 0, firstLine: 0 },
          },
        },
        {
          id: 'TOC1',
          name: 'TOC 1',
          basedOn: 'Normal',
          run: {
            size: SIZE.XIAO_SI,
            bold: false,
            // 目录字体与正文保持一致：一级目录使用宋体小四
            font: { ascii: FONT_EN, eastAsia: FONT_CN, hAnsi: FONT_EN },
          },
          paragraph: {
            spacing: { line: 400, lineRule: 'exact', before: 0, after: 0 },
            indent: { left: 0, firstLine: 0 },
          },
        },
        {
          id: 'TOC2',
          name: 'TOC 2',
          basedOn: 'Normal',
          run: {
            size: SIZE.XIAO_SI,
            bold: false,
            // 目录分级保留缩进：二级目录缩进一级
            font: { ascii: FONT_EN, eastAsia: FONT_CN, hAnsi: FONT_EN },
          },
          paragraph: {
            spacing: { line: 400, lineRule: 'exact', before: 0, after: 0 },
            indent: { left: 360, firstLine: 0 },
          },
        },
        {
          id: 'TOC3',
          name: 'TOC 3',
          basedOn: 'Normal',
          run: {
            size: SIZE.XIAO_SI,
            bold: false,
            font: { ascii: FONT_EN, eastAsia: FONT_CN, hAnsi: FONT_EN },
          },
          paragraph: {
            spacing: { line: 400, lineRule: 'exact', before: 0, after: 0 },
            indent: { left: 720, firstLine: 0 },
          },
        },
      ],
    },
    sections: [
      {
        properties: {
          page: {
            size: { width: PAGE_WIDTH, height: PAGE_HEIGHT },
            margin: {
              top: MARGIN_TOP,
              right: MARGIN_RIGHT,
              bottom: MARGIN_BOTTOM,
              left: MARGIN_LEFT,
              header: MARGIN_HEADER,
              footer: MARGIN_FOOTER,
            },
          },
        },
        headers: {
          default: buildDefaultHeader(),
        },
        children: frontMatterElements,
      },
      {
        properties: {
          page: {
            size: { width: PAGE_WIDTH, height: PAGE_HEIGHT },
            margin: {
              top: MARGIN_TOP,
              right: MARGIN_RIGHT,
              bottom: MARGIN_BOTTOM,
              left: MARGIN_LEFT,
              header: MARGIN_HEADER,
              footer: MARGIN_FOOTER,
            },
            pageNumbers: {
              start: 1,
            },
          },
        },
        headers: {
          default: buildDefaultHeader(),
        },
        footers: {
          default: buildPageNumberFooter(),
        },
        children: bodyElements,
      },
    ],
  });

    console.log('[3/3] 开始打包 Word 文档...');
    const buffer = await Packer.toBuffer(doc);
    fs.writeFileSync(outPath, buffer);
    console.log(`✅ 生成成功: ${outPath}`);
  } finally {
    cleanupTempDir(diagramTempDir);
  }
}

main().catch((error) => {
  console.error('❌ 生成失败:', error);
  process.exit(1);
});
