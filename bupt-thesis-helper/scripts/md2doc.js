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

const DEFAULT_COVER_DATA_FILENAME = 'thesis.cover.json';

function resolveWorkspacePath(workspace, targetPath) {
  return path.resolve(workspace, targetPath);
}

function ensureCoverDataTemplate({ templatePath, workspaceCoverDataPath }) {
  if (fs.existsSync(workspaceCoverDataPath)) {
    return { created: false };
  }
  fs.copyFileSync(templatePath, workspaceCoverDataPath);
  return { created: true };
}

async function main() {
  const skillRoot = path.resolve(__dirname, '..');
  const args = parseArgs(process.argv.slice(2));
  const workspace = path.resolve(args.workspace || process.cwd());
  const markdownInput = args.input || args.markdown || args._[0] || 'thesis.md';
  const markdownPath = path.resolve(workspace, markdownInput);
  const generatorPath = path.resolve(skillRoot, 'scripts', 'generate_thesis.js');
  const composerPath = path.resolve(skillRoot, 'scripts', 'compose_docx.js');
  const coverPath = path.resolve(args.cover || path.join(skillRoot, 'assets', '论文封面+诚信声明.docx'));
  const coverDataTemplatePath = path.resolve(skillRoot, 'assets', 'thesis.cover.example.json');
  const coverDataPath = resolveWorkspacePath(workspace, args['cover-data'] || DEFAULT_COVER_DATA_FILENAME);
  const outputDefaultName = `${path.parse(markdownPath).name || 'thesis'}.docx`;
  const outputPath = path.resolve(workspace, args.output || outputDefaultName);
  const bodyTempName = `${path.parse(outputPath).name}.body.tmp.docx`;
  const bodyTempPath = path.resolve(path.dirname(outputPath), bodyTempName);

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
  if (!fs.existsSync(coverDataTemplatePath)) {
    console.error(`封面信息 JSON 模板不存在: ${coverDataTemplatePath}`);
    process.exit(2);
  }

  const coverDataTemplateStatus = ensureCoverDataTemplate({
    templatePath: coverDataTemplatePath,
    workspaceCoverDataPath: coverDataPath,
  });
  if (coverDataTemplateStatus.created) {
    console.log(`[info] 未发现封面信息 JSON，已复制模板到工作区: ${coverDataPath}`);
    console.log('[info] 请按需填写其中字段；后续再次执行 md2doc 时会自动把第一页封面信息写入 DOCX。');
  } else {
    console.log(`[info] 使用工作区封面信息 JSON: ${coverDataPath}`);
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
  runNodeScript(generatorPath, ['--workspace', workspace, '--input', markdownPath, '--output', bodyTempPath], { cwd: workspace });

  console.log(`[step 2/3] 组装封面与正文: ${composerPath}`);
  runNodeScript(composerPath, ['--cover', coverPath, '--body', bodyTempPath, '--output', outputPath, '--cover-data', coverDataPath], { cwd: workspace });

  console.log(`[step 3/3] 输出完成: ${outputPath}`);
  fs.rmSync(bodyTempPath, { force: true });
}

main().catch((error) => {
  console.error(error && error.stack ? error.stack : String(error));
  process.exit(1);
});
