'use strict';

const fs = require('fs');
const path = require('path');
const { spawnSync } = require('child_process');
const { runChecks, printTextReport } = require('./check_markdown');

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

function runNodeScript(scriptPath, scriptArgs, options = {}) {
  const result = spawnSync(process.execPath, [scriptPath, ...scriptArgs], {
    cwd: options.cwd,
    stdio: 'inherit',
    env: { ...process.env, ...(options.env || {}) },
  });
  if (result.status !== 0) {
    process.exit(result.status || 1);
  }
}

function resolveCliPath(baseDir, targetPath) {
  return path.isAbsolute(targetPath) ? targetPath : path.resolve(baseDir, targetPath);
}

async function main() {
  const skillRoot = path.resolve(__dirname, '..');
  const args = parseArgs(process.argv.slice(2));
  const baseDir = path.resolve(args.workspace || process.cwd());
  const markdownInput = args.input || args.markdown || args._[0];
  if (!markdownInput) {
    console.error('错误: 请指定输入的 Markdown 文件路径。');
    process.exit(1);
  }
  const markdownPath = resolveCliPath(baseDir, markdownInput);
  const generatorPath = path.resolve(skillRoot, 'scripts', 'generate_thesis.js');
  const composerPath = path.resolve(skillRoot, 'scripts', 'compose_docx.js');
  const coverPath = resolveCliPath(baseDir, args.cover || path.join(skillRoot, 'assets', '论文封面+诚信声明.docx'));
  const outputPath = args.output
    ? resolveCliPath(baseDir, args.output)
    : path.join(path.dirname(markdownPath), `${path.parse(markdownPath).name || 'document'}.docx`);
  const bodyTempName = `${path.parse(outputPath).name}.body.tmp.docx`;
  const bodyTempPath = path.join(path.dirname(outputPath), bodyTempName);

  if (!fs.existsSync(markdownPath)) {
    console.error(`Markdown 文件不存在: ${markdownPath}`);
    process.exit(2);
  }
  if (!fs.existsSync(generatorPath)) {
    console.error(`generate_thesis.js 不存在: ${generatorPath}`);
    process.exit(2);
  }
  if (!fs.existsSync(composerPath)) {
    console.error(`compose_docx.js 不存在: ${composerPath}`);
    process.exit(2);
  }
  if (!fs.existsSync(coverPath)) {
    console.error(`封面声明文件不存在: ${coverPath}`);
    process.exit(2);
  }
  if (!args['skip-check']) {
    const result = runChecks(markdownPath);
    printTextReport(result);
    if (result.error_count > 0 && !args.force) {
      console.error('\n检查未通过，已阻止导出。若确需继续，可追加 --force。');
      process.exit(1);
    }
  }

  console.log(`\n[step 1/3] 生成正文 DOCX: ${generatorPath}`);
  runNodeScript(generatorPath, ['--input', markdownPath, '--output', bodyTempPath], { cwd: path.dirname(markdownPath) });

  console.log(`[step 2/3] 组装封面与正文: ${composerPath}`);
  runNodeScript(composerPath, ['--cover', coverPath, '--body', bodyTempPath, '--output', outputPath], { cwd: path.dirname(markdownPath) });

  console.log(`[step 3/3] 输出完成: ${outputPath}`);
  fs.rmSync(bodyTempPath, { force: true });
}

main().catch((error) => {
  console.error(error && error.stack ? error.stack : String(error));
  process.exit(1);
});
