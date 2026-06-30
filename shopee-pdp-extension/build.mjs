// Build minify cho extension: gộp lib/shared.js vào background & popup, minify content scripts,
// rồi copy các file tĩnh sang dist/. Source vẫn giữ dễ đọc; chỉ phát hành thư mục dist/.
import { build } from 'esbuild';
import { cp, rm, mkdir } from 'node:fs/promises';

const OUT = 'dist';
const TARGET = 'chrome111';
const common = { bundle: true, minify: true, target: TARGET, legalComments: 'none' };

await rm(OUT, { recursive: true, force: true });
await mkdir(`${OUT}/content`, { recursive: true });
await mkdir(`${OUT}/popup`, { recursive: true });

// Background & popup: bundle (inline shared.js) + minify, giữ ESM để khớp manifest/popup.html
await build({ ...common, entryPoints: ['background.js'], outfile: `${OUT}/background.js`, format: 'esm' });
await build({ ...common, entryPoints: ['popup/popup.js'], outfile: `${OUT}/popup/popup.js`, format: 'esm' });

// Content scripts: IIFE độc lập, chỉ minify
await build({ ...common, bundle: false, entryPoints: ['content/pageHook.js'], outfile: `${OUT}/content/pageHook.js`, format: 'iife' });
await build({ ...common, bundle: false, entryPoints: ['content/bridge.js'], outfile: `${OUT}/content/bridge.js`, format: 'iife' });

// File tĩnh (drop-in, không cần sửa)
await cp('manifest.json', `${OUT}/manifest.json`);
await cp('popup/popup.html', `${OUT}/popup/popup.html`);
await cp('README.md', `${OUT}/README.md`).catch(() => {});

console.log(`✅ Build xong -> ${OUT}/ (load unpacked thư mục này)`);
