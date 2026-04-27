'use strict';

const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');

function ensureDir(dir) {
  fs.mkdirSync(dir, { recursive: true });
}

function decodeXml(text) {
  return String(text || '')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}

function extractParagraphText(paragraphXml) {
  const parts = [];
  const re = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g;
  let match;
  while ((match = re.exec(paragraphXml)) !== null) {
    parts.push(decodeXml(match[1]));
  }
  return parts.join('').trim();
}

function extractImageRIds(paragraphXml) {
  const ids = [];
  const re = /<a:blip[^>]+r:embed="([^"]+)"/g;
  let match;
  while ((match = re.exec(paragraphXml)) !== null) {
    ids.push(match[1]);
  }
  return ids;
}

function parseRelationships(relsXml) {
  const relationships = new Map();
  const re = /<Relationship\b([^>]*)\/>/g;
  let match;
  while ((match = re.exec(relsXml)) !== null) {
    const attrs = match[1];
    const id = (attrs.match(/\bId="([^"]+)"/) || [])[1];
    const type = (attrs.match(/\bType="([^"]+)"/) || [])[1];
    const target = (attrs.match(/\bTarget="([^"]+)"/) || [])[1];
    if (id && target && /\/image$/i.test(type || '')) {
      relationships.set(id, target);
    }
  }
  return relationships;
}

async function main() {
  const docxPath = process.argv[2];
  const outDir = process.argv[3];
  if (!docxPath || !outDir) {
    throw new Error('Usage: node scripts/extract_docx_images.js <docx-path> <output-dir>');
  }

  ensureDir(outDir);
  const zip = await JSZip.loadAsync(fs.readFileSync(docxPath));
  const documentXml = await zip.file('word/document.xml').async('string');
  const relsXml = await zip.file('word/_rels/document.xml.rels').async('string');
  const rels = parseRelationships(relsXml);

  const paragraphs = [];
  const paragraphRe = /<w:p\b[\s\S]*?<\/w:p>/g;
  let paragraphMatch;
  while ((paragraphMatch = paragraphRe.exec(documentXml)) !== null) {
    const xml = paragraphMatch[0];
    paragraphs.push({
      text: extractParagraphText(xml),
      imageRIds: extractImageRIds(xml),
    });
  }

  const manifest = [];
  let imageIndex = 0;
  for (let pIndex = 0; pIndex < paragraphs.length; pIndex += 1) {
    const paragraph = paragraphs[pIndex];
    for (const rId of paragraph.imageRIds) {
      const target = rels.get(rId);
      if (!target) {
        continue;
      }
      const normalizedTarget = target.replace(/^\.\.\//, '');
      const mediaPath = normalizedTarget.startsWith('word/')
        ? normalizedTarget
        : `word/${normalizedTarget}`;
      const file = zip.file(mediaPath);
      if (!file) {
        continue;
      }

      imageIndex += 1;
      const ext = path.extname(mediaPath) || '.png';
      const outName = `midterm-fig-${String(imageIndex).padStart(2, '0')}${ext}`;
      const outPath = path.join(outDir, outName);
      fs.writeFileSync(outPath, await file.async('nodebuffer'));

      const before = [];
      for (let i = pIndex - 1; i >= 0 && before.length < 3; i -= 1) {
        if (paragraphs[i].text) before.unshift(paragraphs[i].text);
      }
      const after = [];
      for (let i = pIndex + 1; i < paragraphs.length && after.length < 3; i += 1) {
        if (paragraphs[i].text) after.push(paragraphs[i].text);
      }
      manifest.push({
        index: imageIndex,
        file: outName,
        source: mediaPath,
        paragraph: pIndex + 1,
        before,
        after,
      });
    }
  }

  const manifestPath = path.join(outDir, 'manifest.json');
  fs.writeFileSync(manifestPath, JSON.stringify(manifest, null, 2), 'utf-8');
  console.log(`extracted=${manifest.length}`);
  console.log(`manifest=${manifestPath}`);
}

main().catch((error) => {
  console.error(error && error.stack ? error.stack : String(error));
  process.exit(1);
});
