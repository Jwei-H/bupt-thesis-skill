'use strict';

const fs = require('fs');
const path = require('path');

const HEADING_RE = /^(#{1,4})\s+(.+?)\s*$/;
const IMAGE_RE = /^!\[([^\]]*)\]\(([^)]+)\)\s*$/;
const TABLE_CAPTION_RE = /^表\s+(\d+(?:-\d+)+)\s+.+表\s*$/;
const IMAGE_CAPTION_RE = /^图\s+(\d+(?:-\d+)+)\s+.+图\s*$/;
const SECTION2_RE = /^(\d+)\.(\d+)\s+.+$/;
const SECTION3_RE = /^(\d+)\.(\d+)\.(\d+)\s+.+$/;
const KEYWORDS_ZH_RE = /^(\*\*关键词\*\*|关键词)/;
const KEYWORDS_EN_RE = /^(\*\*KEY WORDS\*\*|KEY WORDS)/i;
const LIST_RE = /^(\s*)([-*]|\d+\.)\s+/;

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

const CN_NUM_MAP = new Map([
  ['一', 1], ['二', 2], ['三', 3], ['四', 4], ['五', 5], ['六', 6], ['七', 7], ['八', 8], ['九', 9],
  ['十', 10], ['十一', 11], ['十二', 12], ['十三', 13], ['十四', 14], ['十五', 15], ['十六', 16],
  ['十七', 17], ['十八', 18], ['十九', 19], ['二十', 20],
]);

function addIssue(issues, severity, rule, line, message, suggestion = '') {
  issues.push({ severity, rule, line, message, suggestion });
}

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

function parseBlocks(lines) {
  const blocks = [];
  let index = 0;

  while (index < lines.length) {
    const raw = lines[index];
    const stripped = raw.trim();
    const lineNo = index + 1;

    if (raw.startsWith('```')) {
      const start = index;
      const info = raw.replace(/^```/, '').trim();
      index += 1;
      while (index < lines.length && !lines[index].startsWith('```')) {
        index += 1;
      }
      if (index < lines.length) {
        index += 1;
      }
      const lang = normalizeFenceLang(info);
      const diagramLang = normalizeDiagramLang(info);
      blocks.push({
        type: diagramLang ? 'diagram' : 'codeblock',
        line: start + 1,
        endLine: index,
        lang,
        diagramLang,
        lines: lines.slice(start, index),
      });
      continue;
    }

    if (stripped === '$$') {
      const start = index;
      index += 1;
      while (index < lines.length && lines[index].trim() !== '$$') {
        index += 1;
      }
      if (index < lines.length) {
        index += 1;
      }
      blocks.push({ type: 'formula', line: start + 1, endLine: index, lines: lines.slice(start, index) });
      continue;
    }

    if (!stripped) {
      index += 1;
      continue;
    }

    const headingMatch = raw.match(HEADING_RE);
    if (headingMatch) {
      blocks.push({
        type: 'heading',
        line: lineNo,
        endLine: lineNo,
        text: headingMatch[2].trim(),
        level: headingMatch[1].length,
      });
      index += 1;
      continue;
    }

    if (raw.startsWith('|')) {
      const start = index;
      const tableLines = [];
      while (index < lines.length && lines[index].startsWith('|')) {
        tableLines.push(lines[index]);
        index += 1;
      }
      blocks.push({ type: 'table', line: start + 1, endLine: index, lines: tableLines });
      continue;
    }

    if (IMAGE_RE.test(stripped)) {
      blocks.push({ type: 'image', line: lineNo, endLine: lineNo, text: stripped });
      index += 1;
      continue;
    }

    if (LIST_RE.test(raw)) {
      blocks.push({ type: 'list', line: lineNo, endLine: lineNo, text: stripped });
      index += 1;
      continue;
    }

    blocks.push({ type: 'paragraph', line: lineNo, endLine: lineNo, text: stripped });
    index += 1;
  }

  return blocks;
}

function extractHeadings(blocks) {
  return blocks.filter((block) => block.type === 'heading');
}

function parseChapterNumber(text) {
  const match = text.match(/^第([一二三四五六七八九十]{1,3})章\s+/);
  return match ? CN_NUM_MAP.get(match[1]) || null : null;
}

function getSectionLineRange(blocks, headingText) {
  const headings = extractHeadings(blocks);
  for (let index = 0; index < headings.length; index += 1) {
    if (headings[index].text !== headingText) {
      continue;
    }
    const start = headings[index].line + 1;
    const end = index + 1 < headings.length ? headings[index + 1].line - 1 : blocks[blocks.length - 1].endLine;
    return { start, end };
  }
  return null;
}

function previousNonemptyBlock(blocks, index) {
  return index > 0 ? blocks[index - 1] : null;
}

function nextNonemptyBlock(blocks, index) {
  return index + 1 < blocks.length ? blocks[index + 1] : null;
}

function splitTableRow(line) {
  return line.trim().replace(/^\|/, '').replace(/\|\s*$/, '').split('|').map((cell) => cell.trim());
}

function iterTableCells(tableBlock) {
  const results = [];
  let rowIndex = 0;
  for (const raw of tableBlock.lines || []) {
    if (/^\|[\s\-:|]+\|\s*$/.test(raw)) {
      rowIndex += 1;
      continue;
    }
    const cells = splitTableRow(raw);
    cells.forEach((cell, columnIndex) => {
      results.push({ row: rowIndex + 1, column: columnIndex + 1, cell });
    });
    rowIndex += 1;
  }
  return results;
}

function nearestChapter(headings, lineNo) {
  let current = null;
  for (const heading of headings) {
    if (heading.line > lineNo) {
      break;
    }
    if (heading.level === 2) {
      const chapter = parseChapterNumber(heading.text);
      if (chapter !== null) {
        current = chapter;
      }
    }
  }
  return current;
}

function checkHeadingRules(blocks, issues) {
  const headings = extractHeadings(blocks);
  if (!headings.length) {
    addIssue(issues, 'error', 'heading.missing', 1, '未找到任何标题。');
    return;
  }

  let prevLevel = null;
  let currentChapter = null;
  let currentSection2 = null;
  const lastSection2Index = new Map();
  const lastSection3Index = new Map();
  const chapterNumbers = [];

  if (!headings.some((heading) => heading.level === 2 && heading.text === '目录')) {
    addIssue(issues, 'warning', 'heading.toc_missing', 1, '未发现 `## 目录` 占位标题。', '建议保留 `## 目录`，由导出脚本自动生成目录。');
  }

  for (const heading of headings) {
    if (prevLevel !== null && heading.level > prevLevel + 1) {
      addIssue(
        issues,
        'error',
        'heading.level_jump',
        heading.line,
        `标题层级跳跃：前一个标题级别为 H${prevLevel}，当前为 H${heading.level}。`,
        '避免从 `##` 直接跳到 `####`。',
      );
    }
    prevLevel = heading.level;

    if (heading.level === 2) {
      const chapter = parseChapterNumber(heading.text);
      if (chapter !== null) {
        chapterNumbers.push(chapter);
        currentChapter = chapter;
        currentSection2 = null;
        if (chapterNumbers.length >= 2 && chapterNumbers[chapterNumbers.length - 1] !== chapterNumbers[chapterNumbers.length - 2] + 1) {
          addIssue(issues, 'warning', 'heading.chapter_sequence', heading.line, `章标题序号可能不连续：当前为第 ${chapter} 章。`);
        }
      } else if (!['摘要', 'ABSTRACT', '目录', '参考文献', '致谢', '附录'].includes(heading.text)) {
        addIssue(
          issues,
          'warning',
          'heading.level1_format',
          heading.line,
          `一级标题 \`${heading.text}\` 不属于常见章标题格式。`,
          '若为正文一级标题，建议使用“第 X 章 ...”格式。',
        );
      }
      continue;
    }

    if (heading.level === 3) {
      const match = heading.text.match(SECTION2_RE);
      if (!match) {
        addIssue(issues, 'error', 'heading.level2_format', heading.line, `二级标题 \`${heading.text}\` 不符合 \`X.Y 标题\` 格式。`);
        continue;
      }
      const chapterNo = Number(match[1]);
      const sectionNo = Number(match[2]);
      if (currentChapter === null) {
        addIssue(issues, 'error', 'heading.level2_orphan', heading.line, '二级标题前缺少所属章标题。');
      } else if (chapterNo !== currentChapter) {
        addIssue(issues, 'error', 'heading.level2_chapter_mismatch', heading.line, `二级标题章号为 ${chapterNo}，但当前章为 ${currentChapter}。`);
      }
      const last = lastSection2Index.get(chapterNo);
      if (last !== undefined && sectionNo <= last) {
        addIssue(issues, 'warning', 'heading.level2_order', heading.line, `二级标题编号可能未递增：当前为 ${chapterNo}.${sectionNo}。`);
      }
      lastSection2Index.set(chapterNo, sectionNo);
      currentSection2 = [chapterNo, sectionNo];
      continue;
    }

    if (heading.level === 4) {
      const match = heading.text.match(SECTION3_RE);
      if (!match) {
        addIssue(issues, 'error', 'heading.level3_format', heading.line, `三级标题 \`${heading.text}\` 不符合 \`X.Y.Z 标题\` 格式。`);
        continue;
      }
      const chapterNo = Number(match[1]);
      const sectionNo = Number(match[2]);
      const subNo = Number(match[3]);
      if (!currentSection2) {
        addIssue(issues, 'error', 'heading.level3_orphan', heading.line, '三级标题前缺少所属二级标题。');
      } else {
        if (chapterNo !== currentSection2[0] || sectionNo !== currentSection2[1]) {
          addIssue(
            issues,
            'error',
            'heading.level3_parent_mismatch',
            heading.line,
            `三级标题编号为 ${chapterNo}.${sectionNo}.${subNo}，但当前二级标题为 ${currentSection2[0]}.${currentSection2[1]}。`,
          );
        }
        const key = `${chapterNo}.${sectionNo}`;
        const last = lastSection3Index.get(key);
        if (last !== undefined && subNo <= last) {
          addIssue(issues, 'warning', 'heading.level3_order', heading.line, `三级标题编号可能未递增：当前为 ${chapterNo}.${sectionNo}.${subNo}。`);
        }
        lastSection3Index.set(key, subNo);
      }
    }
  }

  blocks.forEach((block, index) => {
    if (block.type !== 'heading' || ![3, 4].includes(block.level)) {
      return;
    }
    const nextBlock = nextNonemptyBlock(blocks, index);
    if (!nextBlock) {
      addIssue(issues, 'error', 'heading.trailing_empty', block.line, `标题 \`${block.text}\` 后缺少正文描述。`);
      return;
    }
    if (nextBlock.type !== 'paragraph') {
      addIssue(
        issues,
        'error',
        'heading.no_intro_text',
        block.line,
        `标题 \`${block.text}\` 后未先出现正文描述。`,
        '在标题后补一两句引导性文字，再放子标题、图表、公式或列表。',
      );
    }
  });
}

function checkAbstractRules(lines, blocks, issues) {
  [['摘要', KEYWORDS_ZH_RE, false], ['ABSTRACT', KEYWORDS_EN_RE, true]].forEach(([headingText, keywordRe, isEn]) => {
    const range = getSectionLineRange(blocks, headingText);
    if (!range) {
      addIssue(issues, 'warning', 'abstract.missing_section', 1, `未找到 \`${headingText}\` 段落。`);
      return;
    }
    const sectionLines = [];
    for (let index = range.start - 1; index < range.end && index < lines.length; index += 1) {
      sectionLines.push({ line: index + 1, text: lines[index] });
    }

    const keywordEntry = sectionLines.find((entry) => keywordRe.test(entry.text.trim()));
    if (!keywordEntry) {
      addIssue(issues, 'error', 'abstract.keyword_missing', range.start, `\`${headingText}\` 后未找到关键词行。`);
      return;
    }

    // 检查关键词数量（3-5个）
    const keywordLine = keywordEntry.text.trim();
    // 提取关键词部分（去掉"关键词"/"KEY WORDS"前缀和加粗标记）
    const kwContent = keywordLine
      .replace(/^\*\*关键词\*\*|^关键词/, '')
      .replace(/^\*\*KEY WORDS\*\*|^KEY WORDS/i, '')
      .trim();
    if (kwContent) {
      // 中文关键词用空格或中文间隔符分隔，英文关键词用2个以上空格或分号分隔
      const kwList = isEn
        ? kwContent.split(/\s{2,}|;/).map((k) => k.trim()).filter(Boolean)
        : kwContent.split(/[\s；;，,]+/).map((k) => k.trim()).filter(Boolean);
      if (kwList.length < 3) {
        addIssue(issues, 'warning', 'abstract.keyword_count_low', keywordEntry.line, `\`${headingText}\` 关键词数量为 ${kwList.length} 个，官方要求 3-5 个。`);
      } else if (kwList.length > 5) {
        addIssue(issues, 'warning', 'abstract.keyword_count_high', keywordEntry.line, `\`${headingText}\` 关键词数量为 ${kwList.length} 个，官方要求 3-5 个。`);
      }
    }

    let prevNonEmpty = null;
    for (let index = sectionLines.length - 1; index >= 0; index -= 1) {
      const entry = sectionLines[index];
      if (entry.line >= keywordEntry.line) {
        continue;
      }
      if (entry.text.trim()) {
        prevNonEmpty = entry;
        break;
      }
    }

    if (!prevNonEmpty || (prevNonEmpty.text.trim() !== '<br>' && !prevNonEmpty.text.trim().endsWith('<br>'))) {
      addIssue(
        issues,
        'error',
        'abstract.br_before_keywords',
        keywordEntry.line,
        `\`${headingText}\` 的关键词前缺少 \`<br>\` 换行。`,
        '在摘要正文最后一行末尾追加 `<br>`，或单独写一行 `<br>` 后再写关键词。',
      );
    }
  });
}

function recordCaption(issues, captionMap, captionType, lineNo, captionText, expectedChapter) {
  const regex = captionType === 'image' ? IMAGE_CAPTION_RE : TABLE_CAPTION_RE;
  const match = captionText.match(regex);
  if (!match) {
    addIssue(
      issues,
      'error',
      `caption.${captionType}_format`,
      lineNo,
      `${captionType === 'image' ? '图片' : '表格'}题注格式不合法：\`${captionText}\`。`,
    );
    return;
  }
  const number = match[1];
  if (captionMap.has(number)) {
    addIssue(issues, 'warning', `caption.${captionType}_duplicate`, lineNo, `${captionType === 'image' ? '图片' : '表格'}题注编号 \`${number}\` 重复。`);
  } else {
    captionMap.set(number, lineNo);
  }
  if (expectedChapter !== null && Number(number.split('-')[0]) !== expectedChapter) {
    addIssue(issues, 'warning', `caption.${captionType}_chapter_mismatch`, lineNo, `题注编号 \`${number}\` 的章号与当前位置所在章节 ${expectedChapter} 不一致。`);
  }
}

function checkMediaRules(markdownPath, blocks, issues) {
  const headings = extractHeadings(blocks);
  const imageNumbers = new Map();
  const tableNumbers = new Map();

  blocks.forEach((block, index) => {
    if (block.type === 'table') {
      const prevBlock = previousNonemptyBlock(blocks, index);
      if (!prevBlock || prevBlock.type !== 'paragraph') {
        addIssue(issues, 'error', 'table.caption_missing', block.line, '表格上方缺少题注。', '请在表格前一段添加 `表 X-Y ...表` 形式的题注。');
      } else {
        recordCaption(issues, tableNumbers, 'table', prevBlock.line, prevBlock.text, nearestChapter(headings, prevBlock.line));
      }

      iterTableCells(block).forEach(({ row, column, cell }) => {
        const segments = cell.split(/<br\s*\/?>/i).map((segment) => segment.trim()).filter(Boolean);
        segments.forEach((segment, segmentIndex) => {
          const imageMatch = segment.match(IMAGE_RE);
          if (!imageMatch) {
            return;
          }
          const src = normalizeImageSource(imageMatch[2]);
          if (!isRemoteImageSource(src)) {
            const imagePath = path.resolve(path.dirname(markdownPath), src);
            if (!fs.existsSync(imagePath)) {
              addIssue(issues, 'error', 'table.image_missing_file', block.line, `表格单元格中的图片文件不存在：\`${src}\`（第 ${row} 行第 ${column} 列）。`);
            }
          }
          const caption = segments.slice(segmentIndex + 1).find((item) => IMAGE_CAPTION_RE.test(item));
          if (!caption) {
            addIssue(issues, 'error', 'table.image_caption_missing', block.line, `表格单元格中的图片 \`${src}\` 缺少同单元格题注（第 ${row} 行第 ${column} 列）。`);
          } else {
            recordCaption(issues, imageNumbers, 'image', block.line, caption, nearestChapter(headings, block.line));
          }
        });
      });
      return;
    }

    if (block.type === 'image') {
      const match = block.text.match(IMAGE_RE);
      if (!match) {
        return;
      }
      const src = normalizeImageSource(match[2]);
      if (!isRemoteImageSource(src)) {
        const imagePath = path.resolve(path.dirname(markdownPath), src);
        if (!fs.existsSync(imagePath)) {
          addIssue(issues, 'error', 'image.file_missing', block.line, `图片文件不存在：\`${src}\`.`);
        }
      }
      const nextBlock = nextNonemptyBlock(blocks, index);
      if (!nextBlock || nextBlock.type !== 'paragraph') {
        addIssue(issues, 'error', 'image.caption_missing', block.line, `图片 \`${src}\` 下方缺少题注。`, '请在图片后补充 `图 X-Y ...图` 形式的题注。');
      } else {
        recordCaption(issues, imageNumbers, 'image', nextBlock.line, nextBlock.text, nearestChapter(headings, nextBlock.line));
      }
      return;
    }

    if (block.type === 'diagram') {
      const label = block.diagramLang === 'mermaid' ? 'Mermaid 图示' : 'PlantUML 图示';
      const nextBlock = nextNonemptyBlock(blocks, index);
      if (!nextBlock || nextBlock.type !== 'paragraph') {
        addIssue(issues, 'error', 'diagram.caption_missing', block.line, `${label} 下方缺少题注。`, '请在图示代码块后补充 `图 X-Y ...图` 形式的题注。');
      } else {
        recordCaption(issues, imageNumbers, 'image', nextBlock.line, nextBlock.text, nearestChapter(headings, nextBlock.line));
      }
    }
  });
}

function checkFormulaRules(blocks, issues) {
  const headings = extractHeadings(blocks);
  const formulaBlocks = blocks.filter((b) => b.type === 'formula');
  const formulaNumbers = new Map();
  // 匹配行末的 % 式（X-Y） 标签（也支持半角括号）
  const FORMULA_LABEL_LINE_RE = /%\s*(式[（(]\d+-\d+[）)]|[（(]\d+-\d+[）)])\s*$/;

  formulaBlocks.forEach((block) => {
    const blockLines = block.lines || [];
    let labelMatch = null;
    for (const line of blockLines) {
      const m = line.match(FORMULA_LABEL_LINE_RE);
      if (m) {
        labelMatch = m;
        break;
      }
    }

    if (!labelMatch) {
      addIssue(
        issues,
        'warning',
        'formula.no_label',
        block.line,
        '公式块缺少编号标注。',
        '在 $$ 块内任意一行末尾追加 `% 式（X-Y）` 形式的编号，如 `% 式（5-1）`。',
      );
      return;
    }

    const label = labelMatch[1].startsWith('式') ? labelMatch[1] : `式${labelMatch[1]}`;
    const numPart = label.replace(/^式[（(]/, '').replace(/[）)]$/, '');
    const parts = numPart.split('-');
    const labelChapter = Number(parts[0]);
    const key = numPart;

    if (formulaNumbers.has(key)) {
      addIssue(issues, 'warning', 'formula.duplicate_label', block.line, `公式编号 \`${label}\` 重复。`);
    } else {
      formulaNumbers.set(key, block.line);
    }

    const chapter = nearestChapter(headings, block.line);
    if (chapter !== null && labelChapter !== chapter) {
      addIssue(
        issues,
        'warning',
        'formula.chapter_mismatch',
        block.line,
        `公式编号 \`${label}\` 的章号为 ${labelChapter}，但当前位置在第 ${chapter} 章。`,
      );
    }
  });
}

function checkReferenceRules(lines, blocks, issues) {
  // 统计参考文献定义（脚注格式 [^N]: ...）
  const refDefs = [];
  lines.forEach((line, index) => {
    const match = line.match(/^\[\^(\d+)\]:\s*(.+)$/);
    if (match) {
      refDefs.push({ key: match[1], content: match[2].trim(), line: index + 1 });
    }
  });

  // 检查参考文献数量（至少10篇）
  if (refDefs.length > 0 && refDefs.length < 10) {
    addIssue(
      issues,
      'warning',
      'reference.count_low',
      1,
      `参考文献共 ${refDefs.length} 篇，官方要求每篇论文至少查阅 10 篇文献资料。`,
    );
  }

  // 检查正文中的引用是否使用 [^N] 格式（上角标方括号）
  // 同时检查是否存在裸露的 [数字] 格式（没有^）
  const bareRefRe = /(?<!\[)\[(\d+)\](?!\:)/g;
  lines.forEach((line, index) => {
    if (/^\[\^\d+\]:/.test(line)) {
      return; // 跳过参考文献定义行
    }
    let match;
    while ((match = bareRefRe.exec(line)) !== null) {
      addIssue(
        issues,
        'warning',
        'reference.bare_citation',
        index + 1,
        `疑似裸引用 \`[${match[1]}]\`，官方要求使用上角标方括号格式 \`[^${match[1]}]\`。`,
        `将 \`[${match[1]}]\` 改为 \`[^${match[1]}]\`。`,
      );
    }
  });

  // 检查参考文献标题是否存在
  const hasRefSection = blocks.some((block) => block.type === 'heading' && block.level === 2 && block.text === '参考文献');
  if (refDefs.length > 0 && !hasRefSection) {
    addIssue(issues, 'warning', 'reference.section_missing', 1, '存在参考文献定义但未找到 `## 参考文献` 标题。');
  }
}

function checkSpecialSections(blocks, issues) {
  // 检查致谢、附录等特殊章节
  const headings = extractHeadings(blocks);
  const specialSections = ['致谢', '附录'];
  specialSections.forEach((name) => {
    const found = headings.find((h) => h.level === 2 && h.text.startsWith(name));
    if (found) {
      // 致谢/附录应为二级标题（## 致谢）
    }
  });
}

function runChecks(markdownPathInput) {
  const markdownPath = path.resolve(markdownPathInput);
  const text = fs.readFileSync(markdownPath, 'utf-8');
  const lines = text.split(/\r?\n/);
  const blocks = parseBlocks(lines);
  const issues = [];

  checkHeadingRules(blocks, issues);
  checkAbstractRules(lines, blocks, issues);
  checkMediaRules(markdownPath, blocks, issues);
  checkFormulaRules(blocks, issues);
  checkReferenceRules(lines, blocks, issues);
  checkSpecialSections(blocks, issues);

  const headings = extractHeadings(blocks);
  return {
    markdown: markdownPath,
    total_headings: headings.length,
    headings: headings.map((heading) => ({ line: heading.line, level: heading.level, text: heading.text })),
    error_count: issues.filter((issue) => issue.severity === 'error').length,
    warning_count: issues.filter((issue) => issue.severity === 'warning').length,
    issues,
  };
}

function writeLog(logPath, result, lines) {
  const out = [];

  out.push(`# 检查报告`);
  out.push(`文件: ${result.markdown}`);
  out.push(`时间: ${new Date().toLocaleString('zh-CN')}`);
  out.push(`标题数: ${result.total_headings}  错误: ${result.error_count}  警告: ${result.warning_count}`);
  out.push('');

  out.push('## 标题结构');
  result.headings.forEach((h) => {
    const indent = '  '.repeat(Math.max(0, h.level - 1));
    out.push(`${indent}L${h.level} 行${h.line}: ${h.text}`);
  });
  out.push('');

  if (!result.issues.length) {
    out.push('## 问题');
    out.push('无问题。');
  } else {
    const errors = result.issues.filter((i) => i.severity === 'error');
    const warnings = result.issues.filter((i) => i.severity === 'warning');

    if (errors.length) {
      out.push(`## 错误 (${errors.length})`);
      errors.forEach((issue) => {
        const raw = (lines[issue.line - 1] || '').trim().slice(0, 80);
        out.push(`行${issue.line}: ${raw}`);
        out.push(`  [${issue.rule}] ${issue.message}`);
        if (issue.suggestion) out.push(`  → ${issue.suggestion}`);
      });
      out.push('');
    }

    if (warnings.length) {
      out.push(`## 警告 (${warnings.length})`);
      warnings.forEach((issue) => {
        const raw = (lines[issue.line - 1] || '').trim().slice(0, 80);
        out.push(`行${issue.line}: ${raw}`);
        out.push(`  [${issue.rule}] ${issue.message}`);
        if (issue.suggestion) out.push(`  → ${issue.suggestion}`);
      });
    }
  }

  fs.writeFileSync(logPath, out.join('\n') + '\n', 'utf-8');
}

function printTextReport(result) {
  const status = result.error_count > 0 ? '❌' : result.warning_count > 0 ? '⚠️ ' : '✓';
  console.log(`${status}  错误 ${result.error_count}  警告 ${result.warning_count}  标题 ${result.total_headings}  →  ${result.logPath || ''}`);
}

function main() {
  const args = parseArgs(process.argv.slice(2));
  const markdown = args._[0];
  if (!markdown) {
    console.error('错误: 请指定输入的 Markdown 文件路径。');
    process.exit(1);
  }
  const markdownPath = path.resolve(markdown);
  if (!fs.existsSync(markdownPath)) {
    console.error(`文件不存在: ${markdownPath}`);
    process.exit(2);
  }

  const result = runChecks(markdownPath);

  if (args.json) {
    console.log(JSON.stringify(result, null, 2));
  } else {
    const lines = fs.readFileSync(markdownPath, 'utf-8').split(/\r?\n/);
    const logPath = markdownPath.replace(/\.md$/, '') + '.check.log';
    writeLog(logPath, result, lines);
    result.logPath = logPath;
    printTextReport(result);
  }

  process.exit(result.error_count > 0 ? 1 : 0);
}

if (require.main === module) {
  main();
}

module.exports = {
  runChecks,
  printTextReport,
  parseBlocks,
};
