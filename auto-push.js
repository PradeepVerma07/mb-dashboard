/**
 * auto-push.js
 * Watches the project for file changes and automatically git add + commit + push.
 * Run once with:  node auto-push.js
 * Stop with:      Ctrl + C
 */

const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

const ROOT = __dirname;

// Folders / patterns to ignore
const IGNORE = [
  '.git', 'node_modules', 'dist', 'build', '.vite',
  'auto-push.js'
];

let debounceTimer = null;
const DEBOUNCE_MS = 2000; // wait 2 s of quiet before pushing

function shouldIgnore(filePath) {
  return IGNORE.some(seg => filePath.includes(path.sep + seg) || filePath.includes('/' + seg));
}

function gitPush(changedFile) {
  try {
    const rel = path.relative(ROOT, changedFile);
    const timestamp = new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' });
    const message = Auto-save:  - ;

    execSync('git add -A', { cwd: ROOT, stdio: 'pipe' });

    // Only commit if there is something staged
    const status = execSync('git status --porcelain', { cwd: ROOT }).toString().trim();
    if (!status) {
      console.log('[auto-push] No changes to commit.');
      return;
    }

    execSync(git commit -m "", { cwd: ROOT, stdio: 'pipe' });
    execSync('git push', { cwd: ROOT, stdio: 'pipe' });
    console.log([auto-push] Pushed: );
  } catch (err) {
    const msg = err.stderr ? err.stderr.toString().trim() : err.message;
    console.error('[auto-push] Error:', msg);
  }
}

function onChange(filePath) {
  if (shouldIgnore(filePath)) return;
  console.log([auto-push] Change detected: );
  clearTimeout(debounceTimer);
  debounceTimer = setTimeout(() => gitPush(filePath), DEBOUNCE_MS);
}

function watchDir(dir) {
  try {
    fs.watch(dir, { recursive: true }, (event, filename) => {
      if (!filename) return;
      const full = path.join(dir, filename);
      if (!shouldIgnore(full)) onChange(full);
    });
    console.log([auto-push] Watching: );
  } catch (e) {
    console.error('[auto-push] Could not watch:', dir, e.message);
  }
}

console.log('');
console.log('MB Dash - Auto Push Watcher');
console.log('Every save => git commit + push to GitHub');
console.log('Press Ctrl+C to stop.');
console.log('');

watchDir(ROOT);
