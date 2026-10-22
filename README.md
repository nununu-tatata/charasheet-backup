# CoCキャラシ保管スプレッドシート

キャラクター保管所（CoC6版）のキャラクターシートを、Googleスプレッドシートにまとめてバックアップするツールです。

配布ページ：https://nununu-tatata.github.io/charasheet-backup/

## フォルダ構成

| パス | 内容 |
|---|---|
| `main.js` / `chatPalette.js` / `guide.html` / `appsscript.json` | スプレッドシートのGAS（clasp で push） |
| `docs/` | 配布ページ（GitHub Pages） |
| `bookmarklet/` | マイリストからURLを取得するブックマークレットと、そのビルドスクリプト |
| `admin/` | 管理者用スクリプト（通常は push しない） |

## 開発メモ

- ブックマークレットを変更したら `node bookmarklet/build.js` で `guide.html` と `docs/bookmarklet-code.js` を作り直す
- 新しい版を配布するときは `main.js` の `SCRIPT_VERSION` と `docs/version.json` を両方更新する
- GAS の push 先
  - 開発用：`clasp push`（`.clasp.json`）
  - 配布用：`clasp push -P .clasp.dist.json`
