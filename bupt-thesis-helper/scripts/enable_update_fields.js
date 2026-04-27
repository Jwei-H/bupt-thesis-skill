'use strict';

const fs = require('fs');
const JSZip = require('jszip');

async function main() {
  const docxPath = process.argv[2];
  if (!docxPath) {
    throw new Error('Usage: node scripts/enable_update_fields.js <docx-path>');
  }

  const zip = await JSZip.loadAsync(fs.readFileSync(docxPath));
  const settingsPath = 'word/settings.xml';
  const settingsFile = zip.file(settingsPath);
  if (!settingsFile) {
    throw new Error('word/settings.xml not found');
  }

  let settings = await settingsFile.async('string');
  if (!settings.includes('<w:updateFields')) {
    settings = settings.replace('</w:settings>', '<w:updateFields w:val="true"/></w:settings>');
  }
  zip.file(settingsPath, settings);

  const documentPath = 'word/document.xml';
  const documentFile = zip.file(documentPath);
  if (documentFile) {
    let documentXml = await documentFile.async('string');
    documentXml = documentXml.replace(
      /<w:fldChar w:fldCharType="begin"(?![^>]*w:dirty=)/g,
      '<w:fldChar w:fldCharType="begin" w:dirty="true"',
    );
    // WPS recognizes document-location hyperlinks more reliably when bookmarks
    // are direct paragraph children rather than wrapped in an empty run.
    documentXml = documentXml
      .replace(/<w:r><w:bookmarkStart([^>]*)\/><\/w:r>/g, '<w:bookmarkStart$1/>')
      .replace(/<w:r><w:bookmarkEnd([^>]*)\/><\/w:r>/g, '<w:bookmarkEnd$1/>');
    zip.file(documentPath, documentXml);
  }

  const output = await zip.generateAsync({ type: 'nodebuffer' });
  fs.writeFileSync(docxPath, output);
  console.log(`Enabled updateFields and normalized bookmarks for ${docxPath}`);
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
