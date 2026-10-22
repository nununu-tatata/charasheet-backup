// ==========================================
// ブックマークレットのビルド
// mylist-urls.js を javascript: URL に変換し、使い方画面（ルートの guide.html）と配布ページ（docs/）に埋め込む
// 使い方: node bookmarklet/build.js
// ==========================================
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, 'mylist-urls.js'), 'utf8');

// 行全体がコメントの行と、行頭のインデントを取り除く
const code = src
  .split(/\r?\n/)
  .filter(line => !/^\s*\/\//.test(line))
  .map(line => line.trim())
  .filter(line => line !== '')
  .join('\n');

const href = 'javascript:' + encodeURIComponent(code);
const escapeHtml = s => s.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

const template = fs.readFileSync(path.join(__dirname, 'guide.template.html'), 'utf8');
const html = template
  .replace('{{HREF}}', escapeHtml(href))
  .replace('{{HREF_ESCAPED}}', escapeHtml(href));

fs.writeFileSync(path.join(__dirname, '..', 'guide.html'), html);

// 配布ページ（docs/）用
fs.writeFileSync(
  path.join(__dirname, '..', 'docs', 'bookmarklet-code.js'),
  `// bookmarklet/build.js で自動生成（直接編集しない）\nconst BOOKMARKLET_HREF = ${JSON.stringify(href)};\n`
);
console.log(`guide.html と docs/bookmarklet-code.js を生成しました（ブックマークレット ${href.length} 文字）`);
