// Applies YouCoded's two changes to euro-office-lite's built editor folder.
// WHY at bundle time: the app then serves plain files; nothing is rewritten per request.
import { readFile, writeFile, copyFile } from 'node:fs/promises';
import path from 'node:path';

const dir = process.argv[2];
if (!dir) { console.error('usage: patch.mjs <editors dir>'); process.exit(2); }

async function replaceOnce(file, from, to) {
  const p = path.join(dir, file);
  const text = await readFile(p, 'utf8');
  if (!text.includes(from)) { console.error(`patch: pattern not found in ${file}: ${from}`); process.exit(1); }
  await writeFile(p, text.replace(from, to));
}

// 1. Relay first (bridge.js calls __TAURI__ at load), then the theme bridge.
await replaceOnce('index.html', '<head>', '<head><script src="tauri-relay.js"></script><script src="yc-bridge.js"></script>');
// 2. Media and dictionaries under the document's own origin (design §3a), never a shared scheme.
await replaceOnce('bridge.js',
  "var ASC_PROTO_BASE = _isWindows ? 'http://ascdesktop.localhost/' : 'ascdesktop://';",
  "var ASC_PROTO_BASE = location.origin + '/asc/';");
const here = path.dirname(new URL(import.meta.url).pathname);
await copyFile(path.join(here, '..', 'bridge', 'tauri-relay.js'), path.join(dir, 'tauri-relay.js'));
await copyFile(path.join(here, '..', 'bridge', 'yc-bridge.js'), path.join(dir, 'yc-bridge.js'));
console.log('patch: ok');
