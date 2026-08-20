#!/usr/bin/env node
/**
 * Typechecks every *.gs file against the real Apps Script API surface
 * (@types/google-apps-script), so a bad method name (e.g. the setVerticalAlign /
 * setVerticalAlignment mix-up this script exists to catch) is a check-time error
 * instead of a runtime TypeError in front of a user.
 *
 * The Apps Script editor only accepts .gs/.html files, and tsc only reads .js/.ts,
 * so this copies *.gs into a gitignored .typecheck/ dir as .js, runs tsc against
 * that (per tsconfig.json), and rewrites the reported paths back to the original
 * .gs filenames so errors are actionable without extra translation.
 */

const fs = require('fs');
const path = require('path');
const { spawnSync } = require('child_process');

const ROOT = path.resolve(__dirname, '..');
const TYPECHECK_DIR = path.join(ROOT, '.typecheck');

function main() {
  const gsFiles = fs.readdirSync(ROOT).filter((f) => f.endsWith('.gs'));
  if (gsFiles.length === 0) {
    console.error('No .gs files found in ' + ROOT + ' - nothing to check.');
    process.exit(1);
  }

  fs.rmSync(TYPECHECK_DIR, { recursive: true, force: true });
  fs.mkdirSync(TYPECHECK_DIR);

  gsFiles.forEach((gsFile) => {
    const jsName = gsFile.replace(/\.gs$/, '.js');
    fs.copyFileSync(path.join(ROOT, gsFile), path.join(TYPECHECK_DIR, jsName));
  });

  const tscBin = path.join(ROOT, 'node_modules', '.bin', process.platform === 'win32' ? 'tsc.cmd' : 'tsc');
  if (!fs.existsSync(tscBin)) {
    console.error('typescript is not installed - run `npm install` first.');
    process.exit(1);
  }

  const result = spawnSync(tscBin, ['-p', 'tsconfig.json'], { cwd: ROOT, encoding: 'utf8' });
  const output = (result.stdout || '') + (result.stderr || '');

  const rewritten = output
    .split('\n')
    .map((line) => line.replace(/\.typecheck[\\/](\S+?)\.js/g, '$1.gs'))
    .join('\n');

  process.stdout.write(rewritten);

  if (result.status === 0) {
    console.log('\ntypecheck: OK - ' + gsFiles.length + ' file(s) checked against the real Apps Script API.');
  } else {
    console.error('\ntypecheck: FAILED - see errors above (file:line refers to the .gs source).');
  }

  process.exit(result.status === null ? 1 : result.status);
}

main();
