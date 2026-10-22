function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('📥 コンバート')
        .addItem('URLからコンバート', 'updateCharacterSheet')
        .addSeparator()
        .addItem('全シートを一括更新', 'updateAllSheets')
        .addSeparator()
        .addItem('古いスプレッドシートから引っ越す', 'importFromOldSpreadsheet')
        .addSeparator()
        .addItem('マイリストからURLをまとめて取得する方法', 'showBookmarkletGuide')
        .addSeparator()
        .addItem('バージョン情報・更新の確認', 'showVersionInfo')
        .addToUi();
}

// ============================================================
// バージョン管理・更新のお知らせ
// 配布ページの version.json と比較し、新しい版があれば知らせる
// 新しい版を配布するときは SCRIPT_VERSION と docs/version.json の両方を更新する
// ============================================================
const SCRIPT_VERSION = '1.1.0';
const DISTRIBUTION_PAGE_URL = 'https://nununu-tatata.github.io/charasheet-backup/';
const UPDATE_INFO_URL = DISTRIBUTION_PAGE_URL + 'version.json';

// 最新版の情報を取得する（取得できなければ null）
function fetchLatestVersionInfo_() {
    try {
        const res = UrlFetchApp.fetch(UPDATE_INFO_URL, { muteHttpExceptions: true });
        if (res.getResponseCode() !== 200) return null;
        return JSON.parse(res.getContentText());
    } catch (e) {
        return null;
    }
}

// "1.2.0" 形式のバージョンを比較する（a が新しければ正の数）
function compareVersions_(a, b) {
    const pa = a.split('.').map(Number);
    const pb = b.split('.').map(Number);
    for (let i = 0; i < Math.max(pa.length, pb.length); i++) {
        const diff = (pa[i] || 0) - (pb[i] || 0);
        if (diff !== 0) return diff;
    }
    return 0;
}

// 新しい版があれば完了メッセージの末尾に付けるお知らせ文を返す（なければ空文字）
function getUpdateNotice_() {
    const latest = fetchLatestVersionInfo_();
    if (!latest || !latest.version || compareVersions_(latest.version, SCRIPT_VERSION) <= 0) return '';
    return `\n\n🆕 新しいバージョン（v${latest.version}）が公開されています。\n${latest.url || DISTRIBUTION_PAGE_URL}`;
}

function showVersionInfo() {
    const latest = fetchLatestVersionInfo_();
    let msg = `現在のバージョン：v${SCRIPT_VERSION}\n\n`;
    if (!latest || !latest.version) {
        msg += '最新版の情報を取得できませんでした。\n配布ページで確認してください。';
    } else if (compareVersions_(latest.version, SCRIPT_VERSION) > 0) {
        msg += `🆕 新しいバージョン（v${latest.version}）が公開されています。`;
        if (latest.notes) msg += `\n\n【主な変更点】\n${latest.notes}`;
    } else {
        msg += '✅ 最新版を使っています。';
    }
    msg += `\n\n配布ページ：\n${DISTRIBUTION_PAGE_URL}`;
    SpreadsheetApp.getUi().alert(msg);
}

// ============================================================
// ブックマークレットの使い方画面を表示する
// guide.html は bookmarklet/build.js で生成する
// ============================================================
function showBookmarkletGuide() {
    const html = HtmlService.createHtmlOutputFromFile('guide').setWidth(620).setHeight(640);
    SpreadsheetApp.getUi().showModalDialog(html, 'マイリストからURLをまとめて取得する');
}

// ============================================================
// チャパレ・プロフィール欄の位置
// テンプレートでの行番号。持ち物・武器・追加技能が多いキャラは途中で行を挿入するため、
// 実際の位置は「挿入した行数（AB3に保存）」だけ下にずれる
// ============================================================
const PALETTE_SETTING_ROW = 164; // チャパレ設定（B/E/H/K/N/Q列）
const PROFILE_SETTING_ROW = 168; // プロフィール設定（B/E/H/K/N/Q列）
const PALETTE_OUTPUT_ROW = 171;  // チャパレ（B列）・プロフィール（N列）
const CCFOLIA_OUTPUT_ROW = 180;  // CCFOLIAデータ（B列）

// 設定セルからチャパレ・プロフィールの設定を読み取る
function readPaletteSettings_(sheet, rowOffset) {
    const palRow = PALETTE_SETTING_ROW + rowOffset;
    const profRow = PROFILE_SETTING_ROW + rowOffset;
    const get = cell => sheet.getRange(cell).getValue().toString().trim();
    const tekeyMode = get(`K${palRow}`);
    return {
        diceCommand:             get(`B${palRow}`) || 'CCB',
        isOutputAll:             get(`E${palRow}`) === '全技能',
        isAddBrackets:           get(`H${palRow}`) === 'ON',
        isUseFolding:            tekeyMode === 'Tekey向け',
        includeStatVariables:    tekeyMode === 'Tekey向け',
        includeStatx5:           get(`N${palRow}`) === 'ON',
        includeStatRes:          get(`Q${palRow}`) === 'ON',
        includeStatusProfile:    get(`B${profRow}`) !== 'OFF',
        includeSkillListProfile: get(`E${profRow}`) !== 'OFF',
        includeTraits:           get(`H${profRow}`) === 'ON',
        includeWeaponsProfile:   get(`K${profRow}`) !== 'OFF',
        includeItems:            get(`N${profRow}`) !== 'OFF',
        includePersonalData:     get(`Q${profRow}`) === 'ON',
    };
}

// 生成したチャパレ・プロフィール・CCFOLIAデータを書き込む
function writePaletteResult_(sheet, rowOffset, result) {
    sheet.getRange(`B${PALETTE_OUTPUT_ROW + rowOffset}`).setValue(result.palette);
    sheet.getRange(`N${PALETTE_OUTPUT_ROW + rowOffset}`).setValue(result.profile);
    sheet.getRange(`B${CCFOLIA_OUTPUT_ROW + rowOffset}`).setValue(result.ccfoliaJson);
}

// ============================================================
// onEdit トリガー：設定セルが変更されたらチャパレ・プロフを自動再生成
// 対象セル: チャパレ設定行（B/E/H/K列）、プロフィール設定行（B/E/H/K/N/Q列）
// ============================================================
function onEdit(e) {
    if (!e) return;
    const row = e.range.getRow();
    const col = e.range.getColumn();
    const sheet = e.range.getSheet();

    // AB1に保存されたJSONを取得（コンバートしたキャラのシート以外は対象外）
    const storedJson = sheet.getRange('AB1').getValue();
    if (!storedJson || storedJson === '') return;
    const rowOffset = Number(sheet.getRange('AB3').getValue()) || 0;

    // 設定エリアのセルかどうか判定
    const isPalRow = (row === PALETTE_SETTING_ROW + rowOffset && [2,5,8,11].includes(col));
    const isProfRow = (row === PROFILE_SETTING_ROW + rowOffset && [2,5,8,11,14,17].includes(col));
    if (!isPalRow && !isProfRow) return;

    try {
        const data = JSON.parse(storedJson);
        const storedUrl = sheet.getRange('AB2').getValue();
        const result = generateChatPaletteData(data, storedUrl.toString(), readPaletteSettings_(sheet, rowOffset));
        writePaletteResult_(sheet, rowOffset, result);
    } catch (err) {
        Logger.log('onEdit チャパレ再生成エラー: ' + err.message);
    }
}

// ============================================================
// 実行時間の上限対策
// GASは1回の実行が6分で強制終了されるため、余裕を持って次の1件に入る前に打ち切る
// ============================================================
const TIME_LIMIT_MS = 4.5 * 60 * 1000;

// ============================================================
// キャラクター保管所のURLからシートのキーを取り出す
// 数字ID（5219700）と MDID（m + 32桁の英数字）の両方に対応
// ============================================================
function parseSheetKey(url) {
    const str = url.toString().trim();
    const m = str.match(/vampire-blood\.net\/(m[0-9a-f]{32}|\d+)/i) || str.match(/^(m[0-9a-f]{32}|\d+)$/i);
    if (!m) throw new Error('キャラクター保管所のURLではありません');
    return m[1].toLowerCase();
}

// JSONテキストのハッシュ（MD5）を計算する（差分検知用）
function computeHash(text) {
    const hashBytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, text);
    return hashBytes.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

// ============================================================
// キャラクター保管所からキャラシのJSONを取得する
// 非公開のキャラシはログインしていないと 404 になるため、分かりやすいメッセージにする
// ============================================================
function fetchCharacterJson_(key) {
    const res = UrlFetchApp.fetch(`https://charasheet.vampire-blood.net/${key}.json`, { muteHttpExceptions: true });
    const code = res.getResponseCode();
    if (code === 404) throw new Error('キャラシが見つかりません（非公開設定か、削除された可能性があります）');
    if (code !== 200) throw new Error(`キャラクター保管所からの取得に失敗しました（エラー: ${code}）`);
    return res.getContentText();
}

// ============================================================
// GASの「1日に外部サイトから取得できる回数」の上限エラーかどうか
// この状態では残りも全部失敗するため、すぐに処理を止める
// ============================================================
function isUrlFetchQuotaError_(e) {
    const message = (e && e.message) || '';
    return /urlfetch/i.test(message) && /(too many times|多すぎ)/i.test(message);
}
const URLFETCH_QUOTA_MESSAGE =
    '⛔ Googleの制限により、このアカウントでは今日はこれ以上キャラシを取得できません。\n' +
    '（GASで外部サイトからデータを取得できる1日の回数の上限に達しました）\n\n' +
    '時間をおいて（目安：24時間後）、もう一度実行してください。\n' +
    '処理できなかったURLはそのまま残っているので、続きから処理されます。';

// ============================================================
// 二重実行の防止
// メニューを続けて押すと同じ処理が同時に動き、同じキャラのシートを取り合って壊れるため、
// 実行中は次の実行を受け付けない
// ============================================================
function withDocumentLock_(fn) {
    const lock = LockService.getDocumentLock();
    if (!lock.tryLock(1000)) {
        SpreadsheetApp.getUi().alert('別のコンバート・更新を実行中です。\n終わってから、もう一度実行してください。');
        return;
    }
    try {
        fn();
    } finally {
        lock.releaseLock();
    }
}

function updateAllSheets() {
    withDocumentLock_(updateAllSheets_);
}

function updateCharacterSheet() {
    withDocumentLock_(updateCharacterSheet_);
}

// ============================================================
// 全シートを一括更新する関数（差分検知あり）
// シート一覧のB列にあるURL全件に対して処理する
// 時間切れで止まっても、更新済みのキャラはハッシュ一致でスキップされるため再実行で続きから処理できる
// ============================================================
function updateAllSheets_() {
    const startTime = Date.now();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const LIST_SHEET_NAME = 'シート一覧';
    const listSheet = ss.getSheetByName(LIST_SHEET_NAME);

    if (!listSheet) {
        SpreadsheetApp.getUi().alert(`「${LIST_SHEET_NAME}」シートが見つかりません。`);
        return;
    }

    const lastRow = listSheet.getLastRow();
    if (lastRow < 2) {
        SpreadsheetApp.getUi().alert('シート一覧にデータがありません。先に「URLからコンバート」で1件以上登録してください。');
        return;
    }

    // CB列のヘッダーを確認・設定
    const HASH_COL = 80; // CB列（A-T:20列 + 技能59列 + 1）
    if (listSheet.getRange(1, HASH_COL).getValue() === '') {
        listSheet.getRange(1, HASH_COL).setValue('hash(変更検知用)');
    }

    // B列のURL一覧と既存ハッシュ（CB列）を取得
    const dataRange = listSheet.getRange(2, 1, lastRow - 1, HASH_COL);
    const allRows = dataRange.getValues();

    let updatedCount = 0;
    let skippedCount = 0;
    let errorCount = 0;
    let timedOut = false;
    let quotaExceeded = false;
    const errorMessages = [];

    for (let i = 0; i < allRows.length; i++) {
        if (Date.now() - startTime > TIME_LIMIT_MS) {
            timedOut = true;
            break;
        }

        const url = allRows[i][1]; // B列 = index 1
        const savedHash = allRows[i][HASH_COL - 1].toString();

        if (!url || url.toString().trim() === '') continue;

        const urlStr = url.toString().trim();

        // URLのバリデーション：httpで始まらない行（ヘッダーや誤入力）をスキップ
        if (!urlStr.startsWith('http')) continue;

        const listRowNum = i + 2;
        let key;
        try {
            key = parseSheetKey(urlStr);
        } catch (e) {
            errorCount++;
            errorMessages.push(`${listRowNum}行目: ${e.message}`);
            continue;
        }

        // サーバー負荷軽減のためリクエスト前に少し待機（500ms）
        Utilities.sleep(500);

        let responseText = '';
        try {
            responseText = fetchCharacterJson_(key);
        } catch (e) {
            if (isUrlFetchQuotaError_(e)) {
                quotaExceeded = true;
                break;
            }
            errorCount++;
            errorMessages.push(`ID ${key}: 取得失敗 - ${e.message}`);
            continue;
        }

        if (computeHash(responseText) === savedHash) {
            // ハッシュ一致 → 変更なし → スキップ
            skippedCount++;
            continue;
        }

        // ハッシュが異なる → 変更あり → コンバート実行（ハッシュはconvertOneUrl内で一覧に保存される）
        try {
            convertOneUrl(ss, urlStr, responseText);
            updatedCount++;
        } catch (e) {
            errorCount++;
            errorMessages.push(`ID ${key}: コンバート失敗 - ${e.message}`);
        }
    }

    let msg = `一括更新が完了しました！\n\n更新: ${updatedCount}件\nスキップ（変更なし）: ${skippedCount}件`;
    if (errorCount > 0) {
        msg += `\nエラー: ${errorCount}件\n\n${errorMessages.join('\n')}`;
    }
    if (timedOut) {
        msg = msg.replace('一括更新が完了しました！', '⏱ 時間の上限に近づいたため、途中で止めました。');
        msg += '\n\nもう一度「全シートを一括更新」を実行してください。\n更新済みのキャラは自動でスキップされ、続きから処理されます。';
    }
    if (quotaExceeded) {
        msg = msg.replace('一括更新が完了しました！', URLFETCH_QUOTA_MESSAGE);
        SpreadsheetApp.getUi().alert(msg);
        return;
    }
    SpreadsheetApp.getUi().alert(msg + getUpdateNotice_());
}

// ============================================================
// URL 1件のコンバート処理（共通コア関数）
// updateCharacterSheet / updateAllSheets の両方から呼ばれる
// responseText を渡すと再フェッチを省略できる（一括更新時に活用）
// ============================================================
function convertOneUrl(ss, inputUrl, responseText) {
    const TEMPLATE_SHEET_NAME = 'テンプレート';
    const LIST_SHEET_NAME = 'シート一覧';

    const templateSheet = ss.getSheetByName(TEMPLATE_SHEET_NAME);
    if (!templateSheet) throw new Error(`「${TEMPLATE_SHEET_NAME}」シートが見つかりません。`);
    const listSheet = ss.getSheetByName(LIST_SHEET_NAME);
    if (!listSheet) throw new Error(`「${LIST_SHEET_NAME}」シートが見つかりません。`);

    // --- 1. JSONデータの取得（responseTextが渡されなければ自分でフェッチ） ---
    const sheetKey = parseSheetKey(inputUrl);
    if (!responseText) {
        responseText = fetchCharacterJson_(sheetKey);
    }

    const data = JSON.parse(responseText);
    if (data.game && data.game !== 'coc') {
        throw new Error(`CoC6版のキャラクターシートではありません（${data.game}）`);
    }

    // MDIDで渡されても、JSON内の data_id（数字ID）に統一する
    const id = data.data_id ? data.data_id.toString() : sheetKey;
    const url = `https://charasheet.vampire-blood.net/${id}`;
    // 一覧やシートに古い形式（MDID）で保存されていても同じキャラと判定できるようにする
    const knownKeys = [id, sheetKey];
    if (data.phrase) knownKeys.push('m' + data.phrase.toString().toLowerCase());

    // --- 2. シートの用意 ---
    const listRow = findListRow_(listSheet, knownKeys);
    const oldSheet = findExistingSheet_(ss, listSheet, listRow, knownKeys, data);
    const sheetName = decideSheetName_(ss, data, id, oldSheet);

    // 新しいシートへの書き込みが成功してから旧シートと差し替える（途中で失敗しても旧シートが消えないように）
    const targetSheet = templateSheet.copyTo(ss);
    try {
        const insertedRows = writeCharacterData_(targetSheet, data);
        // 書き込みは後でまとめて反映されることがあり、入力規則違反などのエラーが遅れて出るため、
        // チャパレ（エラーでも続行する部分）に進む前と、旧シートと差し替える前に確定させる
        SpreadsheetApp.flush();
        writeChatPalette_(targetSheet, data, url, responseText, insertedRows);
        SpreadsheetApp.flush();
    } catch (e) {
        ss.deleteSheet(targetSheet);
        throw e;
    }

    let originalPosition = -1; // -1 = 新規（位置復元不要）
    if (oldSheet) {
        originalPosition = oldSheet.getIndex(); // 1始まりの位置を記録
        ss.deleteSheet(oldSheet);
    }
    targetSheet.setName(sheetName);

    // 既存シートの更新だった場合、元の位置に戻す
    if (originalPosition > 0) {
        const totalSheets = ss.getSheets().length;
        const restorePos = Math.min(originalPosition, totalSheets);
        ss.setActiveSheet(targetSheet);
        ss.moveActiveSheet(restorePos);
    }

    // --- 3. 一覧シートの更新 ---
    updateListSheet_(listSheet, listRow, targetSheet, data, id, url, sheetName, responseText);

    return sheetName;
}

// シート一覧のA列（ID）から該当キャラの行番号を探す（見つからなければ -1）
function findListRow_(listSheet, knownKeys) {
    const lastRow = listSheet.getLastRow();
    if (lastRow < 2) return -1;
    const ids = listSheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
        if (knownKeys.includes(ids[i][0].toString().trim().toLowerCase())) return i + 2;
    }
    return -1;
}

// 更新対象の既存シートを探す（同じ名前の別キャラを上書きしないよう、シート名ではなくIDで特定する）
function findExistingSheet_(ss, listSheet, listRow, knownKeys, data) {
    // 一覧のC列のリンク（#gid=xxx）からシートを特定する
    if (listRow > 0) {
        const m = listSheet.getRange(listRow, 3).getFormula().match(/#gid=(\d+)/);
        if (m) {
            const gid = Number(m[1]);
            const sheet = ss.getSheets().find(s => s.getSheetId() === gid);
            if (sheet) return sheet;
        }
    }
    // 一覧から見つからない場合は、同名シートのAB2（保存済みURL）で同じキャラか確認する
    const title = data.data_title || data.pc_name;
    const sameNameSheet = title ? ss.getSheetByName(title) : null;
    if (sameNameSheet) {
        try {
            if (knownKeys.includes(parseSheetKey(sameNameSheet.getRange('AB2').getValue()))) return sameNameSheet;
        } catch (e) {
            // AB2にURLがない＝別のシートなので対象外
        }
    }
    return null;
}

// シート名の決定 (data_title を優先、なければ pc_name)
// 別キャラや管理用シートと名前がかぶる場合は末尾にIDを付ける
function decideSheetName_(ss, data, id, oldSheet) {
    const RESERVED_NAMES = ['テンプレート', 'シート一覧', 'コンバートURL入力'];
    let name = (data.data_title || data.pc_name || '').toString().trim().slice(0, 100) || `無名_${id}`;
    const existing = ss.getSheetByName(name);
    const isOtherSheet = existing && (!oldSheet || existing.getSheetId() !== oldSheet.getSheetId());
    if (isOtherSheet || RESERVED_NAMES.includes(name)) {
        name = `${name.slice(0, 90)}_${id}`;
    }
    return name;
}

// ============================================================
// キャラクターデータをシートに書き込む
// 戻り値：挿入した行数（その下にあるチャパレ欄がずれる行数）
// ============================================================
function writeCharacterData_(targetSheet, data) {
    let insertedRows = 0;

    // --- 3. 基礎ステータス・ポイント・プロフィールの書き込み ---
    function getDiff(baseKey, currentKey) {
        const base = parseInt(data[baseKey]) || 0;
        const current = parseInt(data[currentKey]) || 0;
        const diff = current - base;
        return diff !== 0 ? diff : "";
    }

    const statUpdates = [
        { cell: 'F137', value: data.pc_name }, // キャラクター名
        { cell: 'E3', value: parseInt(data.NA1) || 0 }, { cell: 'E4', value: getDiff('NA1', 'NP1') }, // STR
        { cell: 'F3', value: parseInt(data.NA2) || 0 }, { cell: 'F4', value: getDiff('NA2', 'NP2') }, // CON
        { cell: 'G3', value: parseInt(data.NA3) || 0 }, { cell: 'G4', value: getDiff('NA3', 'NP3') }, // POW
        { cell: 'H3', value: parseInt(data.NA4) || 0 }, { cell: 'H4', value: getDiff('NA4', 'NP4') }, // DEX
        { cell: 'I3', value: parseInt(data.NA5) || 0 }, { cell: 'I4', value: getDiff('NA5', 'NP5') }, // APP
        { cell: 'J3', value: parseInt(data.NA6) || 0 }, { cell: 'J4', value: getDiff('NA6', 'NP6') }, // SIZ
        { cell: 'K3', value: parseInt(data.NA7) || 0 }, { cell: 'K4', value: getDiff('NA7', 'NP7') }, // INT
        { cell: 'L3', value: parseInt(data.NA8) || 0 }, { cell: 'L4', value: getDiff('NA8', 'NP8') }, // EDU

        // 派生ステータス・SAN関連
        { cell: 'M3', value: parseInt(data.NA9) || 0 },  // HP最大値
        { cell: 'N3', value: parseInt(data.NA10) || 0 }, // MP最大値
        { cell: 'O3', value: parseInt(data.NA11) || 0 }, // 初期SAN
        { cell: 'F10', value: parseInt(data.SAN_Left) || 0 }, // 現在SAN値
        { cell: 'P3', value: parseInt(data.NA12) || 0 }, // アイデア
        { cell: 'Q3', value: parseInt(data.NA13) || 0 }, // 幸運
        { cell: 'R3', value: parseInt(data.NA14) || 0 }, // 知識
        { cell: 'K10', value: data.SAN_Danger }, // 不定領域
        { cell: 'H10', value: data.SAN_max || data.SAN_Max }, // 現在SAN上限

        // ポイント関連
        { cell: 'H13', value: data.TS_Maximum || data.ShokugyoPt }, // 職業P総計
        { cell: 'K13', value: data.TS_Add || data.ShokugyoPtAdd },  // 職業P追加
        { cell: 'F13', value: data.TS_Total },                      // 職業P使用数
        { cell: 'H14', value: data.TK_Maximum || data.KyoumiPt },   // 興味P総計
        { cell: 'K14', value: data.TK_Add || data.KyoumiPtAdd },    // 興味P追加
        { cell: 'F14', value: data.TK_Total },                      // 興味P使用数

        // プロフィール・その他
        { cell: 'E102', value: data.dmg_bonus }, // ダメージボーナス(DB)
        { cell: 'E139', value: data.shokugyo || data.shuzoku },  // 職業
        { cell: 'E140', value: data.pc_height }, // 身長
        { cell: 'J140', value: data.pc_weight }, // 体重
        { cell: 'O140', value: data.pc_kigen },  // 出身
        { cell: 'E141', value: data.color_hair },// 髪の色
        { cell: 'J141', value: data.color_eye }, // 瞳の色
        { cell: 'O141', value: data.color_skin },// 肌の色
        { cell: 'J139', value: data.age },       // 年齢
        { cell: 'O139', value: data.sex },       // 性別

        // 所持金・メモ・タグ
        { cell: 'N134', value: data.price_item_sum }, // 価格総計
        { cell: 'S134', value: data.money },          // 現在の所持金
        { cell: 'W134', value: data.debt },           // 預金・借金
        { cell: 'B144', value: data.pc_making_memo }, // その他メモ
        { cell: 'F138', value: data.pc_tags || "" },  // タグ

        // 自由記述サブ技能(D列)
        { cell: 'D52', value: data.unten_bunya },
        { cell: 'D57', value: data.seisaku_bunya },
        { cell: 'D58', value: data.main_souju_norimono },
        { cell: 'D72', value: data.mylang_name },
        { cell: 'D82', value: data.geijutu_bunya }
    ];

    statUpdates.forEach(u => {
        if (u.value !== undefined && u.value !== "") targetSheet.getRange(u.cell).setValue(u.value);
    });

    // --- ★配列の安全取得機能 ---
    function getSafeArray(key) {
        if (data[key] == null || data[key] === "") return [];
        if (Array.isArray(data[key])) return data[key];
        if (typeof data[key] === 'object') return Object.values(data[key]);
        return [data[key]];
    }

    // --- 4. 技能データの範囲読み取り（行がズレる前に初期位置を把握する） ---
    const skillRangeVals = targetSheet.getRange("B18:C120").getValues();

    // --- 5. 所持品・所持金の書き込みと行挿入（一番下から処理） ---
    const itemClearCols = ['B', 'J', 'L', 'N', 'P'];
    itemClearCols.forEach(col => targetSheet.getRange(`${col}113:${col}132`).clearContent());

    const itemNames = getSafeArray('item_name');
    const itemTankas = getSafeArray('item_tanka');
    const itemNums = getSafeArray('item_num');
    const itemPrices = getSafeArray('item_price');
    const itemMemos = getSafeArray('item_memo');

    let writtenItemCount = 0;
    for (let i = 0; i < itemNames.length; i++) {
        if (!itemNames[i]) continue; // 空欄はスキップ

        let rowNum = 113 + writtenItemCount;
        if (writtenItemCount >= 20) { // 21個目以降は行を挿入
            targetSheet.insertRowAfter(rowNum - 1);
            insertedRows++;
            let prevRowRange = targetSheet.getRange(`${rowNum - 1}:${rowNum - 1}`);
            let newRowRange = targetSheet.getRange(`${rowNum}:${rowNum}`);
            prevRowRange.copyTo(newRowRange, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
            // コピーした行の不要なデータをクリア
            itemClearCols.forEach(col => targetSheet.getRange(`${col}${rowNum}`).clearContent());
        }

        targetSheet.getRange(`B${rowNum}`).setValue(itemNames[i] || "");
        targetSheet.getRange(`J${rowNum}`).setValue(itemTankas[i] || "");
        targetSheet.getRange(`L${rowNum}`).setValue(itemNums[i] || "");
        targetSheet.getRange(`N${rowNum}`).setValue(itemPrices[i] || "");
        targetSheet.getRange(`P${rowNum}`).setValue(itemMemos[i] || "");

        writtenItemCount++;
    }

    // --- 6. 戦闘・武器・防具の書き込みと行挿入 ---
    // テンプレートのG列にある「%」などを消さないため、書き込む列のみクリアする
    const weaponClearCols = ['B', 'F', 'H', 'K', 'M', 'O', 'Q', 'S'];
    weaponClearCols.forEach(col => targetSheet.getRange(`${col}104:${col}109`).clearContent());

    const armsNames = getSafeArray('arms_name');
    const armsHits = getSafeArray('arms_hit');
    const armsDamages = getSafeArray('arms_damage');
    const armsRanges = getSafeArray('arms_range');
    const armsAttackCounts = getSafeArray('arms_attack_count');
    const armsLastShots = getSafeArray('arms_last_shot');
    const armsVitalities = getSafeArray('arms_vitality');
    const armsSonotas = getSafeArray('arms_sonota');

    let writtenWeaponCount = 0;
    for (let i = 0; i < armsNames.length; i++) {
        if (!armsNames[i]) continue;

        let rowNum = 104 + writtenWeaponCount;
        if (writtenWeaponCount >= 6) { // 7個目以降は行を挿入
            targetSheet.insertRowAfter(rowNum - 1);
            insertedRows++;
            let prevRowRange = targetSheet.getRange(`${rowNum - 1}:${rowNum - 1}`);
            let newRowRange = targetSheet.getRange(`${rowNum}:${rowNum}`);
            prevRowRange.copyTo(newRowRange, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
            // コピーした行の不要なデータをクリア
            weaponClearCols.forEach(col => targetSheet.getRange(`${col}${rowNum}`).clearContent());
        }

        targetSheet.getRange(`B${rowNum}`).setValue(armsNames[i] || "");
        targetSheet.getRange(`F${rowNum}`).setValue(armsHits[i] || "");
        targetSheet.getRange(`H${rowNum}`).setValue(armsDamages[i] || "");
        targetSheet.getRange(`K${rowNum}`).setValue(armsRanges[i] || "");
        targetSheet.getRange(`M${rowNum}`).setValue(armsAttackCounts[i] || "");
        targetSheet.getRange(`O${rowNum}`).setValue(armsLastShots[i] || "");
        targetSheet.getRange(`Q${rowNum}`).setValue(armsVitalities[i] || "");
        targetSheet.getRange(`S${rowNum}`).setValue(armsSonotas[i] || "");

        writtenWeaponCount++;
    }

    // --- 7. 技能データの解析（辞書作成）と行挿入処理 ---
    const categories = [
        { prefix: "TB", defaults: ["回避", "キック", "組み付き", "こぶし（パンチ）", "頭突き", "投擲", "マーシャルアーツ", "拳銃", "サブマシンガン", "ショットガン", "マシンガン", "ライフル"] },
        { prefix: "TF", defaults: ["応急手当", "鍵開け", "隠す", "隠れる", "聞き耳", "忍び歩き", "写真術", "精神分析", "追跡", "登攀", "図書館", "目星"] },
        { prefix: "TA", defaults: ["運転", "機械修理", "重機械操作", "乗馬", "水泳", "製作", "操縦", "跳躍", "電気修理", "ナビゲート", "変装"] },
        { prefix: "TC", defaults: ["言いくるめ", "信用", "説得", "値切り", "母国語"] },
        { prefix: "TK", defaults: ["医学", "オカルト", "化学", "クトゥルフ神話", "芸術", "経理", "考古学", "コンピューター", "心理学", "人類学", "生物学", "地質学", "電子工学", "天文学", "博物学", "物理学", "法律", "薬学", "歴史"] }
    ];

    // 各カテゴリーの最後のデフォルト技能の行インデックスと、その下の空欄を特定する
    let catInfo = categories.map(cat => {
        let lastIdx = -1;
        for (let i = 0; i < skillRangeVals.length; i++) {
            if (cat.defaults.includes(skillRangeVals[i][1].toString().trim())) {
                lastIdx = i;
            }
        }

        let slots = [];
        if (lastIdx !== -1) {
            let idx = lastIdx + 1;
            while (idx < skillRangeVals.length) {
                let rawB = skillRangeVals[idx][0];
                let valC = skillRangeVals[idx][1].toString().trim();

                // C列が空欄であり、かつB列が文字列（カテゴリーヘッダーなど）ではない場合を空き枠と判定
                if (valC === "" && (typeof rawB === "boolean" || rawB.toString().trim() === "")) {
                    slots.push(idx);
                    idx++;
                } else {
                    break;
                }
            }
        }

        // 追加技能（カスタム技能）を抽出
        let customSkills = [];
        const arrJob = getSafeArray(cat.prefix + "AS");
        const arrInterest = getSafeArray(cat.prefix + "AK");
        const arrCheckbox = getSafeArray(cat.prefix + "AU"); // AUは成長チェックボックス
        const arrGrowth = getSafeArray(cat.prefix + "AA");   // AAは成長P
        const arrOther = getSafeArray(cat.prefix + "AO");    // AOはその他P
        const arrDefault = getSafeArray(cat.prefix + "AD");
        const arrTotal = getSafeArray(cat.prefix + "AP");
        const arrNames = getSafeArray(cat.prefix + "AName");

        const defaultLen = cat.defaults.length;
        for (let i = 0; i < arrNames.length; i++) {
            let name = arrNames[i];
            if (name) {
                const targetIndex = defaultLen + i;
                let jobVal = parseInt(arrJob[targetIndex]) || 0;
                let intVal = parseInt(arrInterest[targetIndex]) || 0;
                const checkboxRaw = arrCheckbox[targetIndex];
                const checkboxVal = (checkboxRaw === undefined || checkboxRaw === "") ? "" : checkboxRaw.toString();
                let growthVal = parseInt(arrGrowth[targetIndex]) || 0;
                let otherVal = parseInt(arrOther[targetIndex]) || 0;
                const defVal = parseInt(arrDefault[targetIndex]) || 0;

                // JSONで成長値が出力されないバグ対策：合計値(AP)から逆算して補完する
                const totalVal = parseInt(arrTotal[targetIndex]) || 0;
                const currentSum = defVal + jobVal + intVal + growthVal + otherVal;
                if (totalVal !== 0 && currentSum !== totalVal) {
                    const diff = totalVal - currentSum;
                    if (diff > 0) growthVal += diff;
                    else otherVal += diff;
                }

                // 名前が入力されていれば、ポイントが0でも抽出する
                customSkills.push({
                    name: name, job: jobVal, interest: intVal, growth: growthVal,
                    other: otherVal, defaultVal: defVal, checkbox: checkboxVal
                });
            }
        }

        return { cat: cat, lastIdx: lastIdx, slots: slots, customSkills: customSkills };
    });

    // --- 6-1. デフォルト技能の書き込み ---
    const defaultSkillDict = {};
    categories.forEach(cat => {
        const arrJob = getSafeArray(cat.prefix + "AS");
        const arrInterest = getSafeArray(cat.prefix + "AK");
        const arrCheckbox = getSafeArray(cat.prefix + "AU");
        const arrGrowth = getSafeArray(cat.prefix + "AA");
        const arrOther = getSafeArray(cat.prefix + "AO");
        const arrDefault = getSafeArray(cat.prefix + "AD");
        const arrTotal = getSafeArray(cat.prefix + "AP");

        cat.defaults.forEach((name, index) => {
            let jobVal = parseInt(arrJob[index]) || 0;
            let intVal = parseInt(arrInterest[index]) || 0;
            const checkboxRaw = arrCheckbox[index];
            const checkboxVal = (checkboxRaw === undefined || checkboxRaw === "") ? "" : checkboxRaw.toString();
            let growthVal = parseInt(arrGrowth[index]) || 0;
            let otherVal = parseInt(arrOther[index]) || 0;
            const defVal = parseInt(arrDefault[index]) || 0;

            // JSONで成長値が出力されないバグ対策：合計値(AP)から逆算して補完する
            const totalVal = parseInt(arrTotal[index]) || 0;
            const currentSum = defVal + jobVal + intVal + growthVal + otherVal;
            if (totalVal !== 0 && currentSum !== totalVal) {
                const diff = totalVal - currentSum;
                if (diff > 0) growthVal += diff;
                else otherVal += diff;
            }

            if (jobVal > 0 || intVal > 0 || growthVal > 0 || otherVal > 0 || checkboxVal !== "") {
                defaultSkillDict[name] = {
                    job: jobVal, interest: intVal, growth: growthVal,
                    other: otherVal, checkbox: checkboxVal
                };
            }
        });
    });

    for (let i = 0; i < skillRangeVals.length; i++) {
        let skillName = skillRangeVals[i][1].toString().trim();
        if (skillName && defaultSkillDict[skillName]) {
            const row = 18 + i;
            let cbVal = defaultSkillDict[skillName].checkbox;
            if (cbVal === "1") {
                targetSheet.getRange(`B${row}`).setValue(true);
            } else if (cbVal !== "" && cbVal !== undefined) {
                let cell = targetSheet.getRange(`B${row}`);
                cell.clearDataValidations();
                cell.setValue(cbVal);
            }

            if (defaultSkillDict[skillName].job > 0) targetSheet.getRange(`G${row}`).setValue(defaultSkillDict[skillName].job);
            if (defaultSkillDict[skillName].interest > 0) targetSheet.getRange(`H${row}`).setValue(defaultSkillDict[skillName].interest);
            if (defaultSkillDict[skillName].growth > 0) targetSheet.getRange(`I${row}`).setValue(defaultSkillDict[skillName].growth);
            if (defaultSkillDict[skillName].other > 0) targetSheet.getRange(`J${row}`).setValue(defaultSkillDict[skillName].other);
        }
    }

    // --- 6-2. 追加技能の書き込みと行挿入（下にあるカテゴリーから順に処理） ---
    // 下の行を挿入しても上の行番号には影響を与えないため、降順で処理する
    catInfo.sort((a, b) => b.lastIdx - a.lastIdx);

    catInfo.forEach(info => {
        let customSkills = info.customSkills;
        if (customSkills.length === 0) return;

        let slots = info.slots;

        // 空き枠が足りない場合、必要な数だけ事前に「最後の枠の上」に行を挿入して枠を確保する
        if (customSkills.length > slots.length && slots.length > 0) {
            let diff = customSkills.length - slots.length;
            let insertRow = slots[slots.length - 1] + 18;
            let firstSlotRow = slots[0] + 18;

            for (let k = 0; k < diff; k++) {
                // 最後の空き枠の「上」に挿入することで、グループ化の範囲内に強制的に含める
                targetSheet.insertRowBefore(insertRow);
                insertedRows++;

                // コピー元を決定。枠の途中（firstSlotRow）からコピーすることで、末尾の太枠が連続するのを防ぎチェックボックスを維持する
                // slotsが1つしかない場合は、挿入によって元の行が下にズレていくため補正する
                let sourceRow = (firstSlotRow === insertRow) ? (insertRow + k + 1) : firstSlotRow;

                let sourceRange = targetSheet.getRange(`${sourceRow}:${sourceRow}`);
                let newRange = targetSheet.getRange(`${insertRow}:${insertRow}`);
                sourceRange.copyTo(newRange, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

                // コピーした行の不要なデータをクリア
                targetSheet.getRange(`B${insertRow}:C${insertRow}`).clearContent();
                targetSheet.getRange(`F${insertRow}:J${insertRow}`).clearContent();
            }

            // 枠が拡張されたのでslotsを再定義（開始位置から連続した配列にする）
            let baseSlot = slots[0];
            slots = [];
            for (let k = 0; k < customSkills.length; k++) {
                slots.push(baseSlot + k);
            }
        } else if (customSkills.length > slots.length && slots.length === 0) {
            // 空き枠が一切ない場合は、デフォルト技能の直下に追加していく
            let insertAtRow = info.lastIdx + 18;
            let diff = customSkills.length;

            for (let k = 0; k < diff; k++) {
                let rowNum = insertAtRow + k + 1;
                targetSheet.insertRowAfter(rowNum - 1);
                insertedRows++;

                // 最後のデフォルト技能が太枠を持っている場合を避けるため、1つ上（通常の枠）から書式をコピー
                let sourceRow = (insertAtRow > 18) ? (insertAtRow - 1) : insertAtRow;
                let sourceRange = targetSheet.getRange(`${sourceRow}:${sourceRow}`);
                let newRange = targetSheet.getRange(`${rowNum}:${rowNum}`);
                sourceRange.copyTo(newRange, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

                targetSheet.getRange(`B${rowNum}:C${rowNum}`).clearContent();
                targetSheet.getRange(`F${rowNum}:J${rowNum}`).clearContent();

                slots.push(info.lastIdx + k + 1);
            }
        }

        // 確保された枠（slots）に対してデータを書き込む
        for (let i = 0; i < customSkills.length; i++) {
            let skill = customSkills[i];
            let rowNum = slots[i] + 18;

            let cbVal = skill.checkbox;
            if (cbVal === "1") {
                targetSheet.getRange(`B${rowNum}`).setValue(true);
            } else if (cbVal !== "" && cbVal !== undefined) {
                let cell = targetSheet.getRange(`B${rowNum}`);
                cell.clearDataValidations();
                cell.setValue(cbVal);
            }

            targetSheet.getRange(`C${rowNum}`).setValue(skill.name);
            targetSheet.getRange(`F${rowNum}`).setValue(skill.defaultVal);
            if (skill.job > 0) targetSheet.getRange(`G${rowNum}`).setValue(skill.job);
            if (skill.interest > 0) targetSheet.getRange(`H${rowNum}`).setValue(skill.interest);
            if (skill.growth > 0) targetSheet.getRange(`I${rowNum}`).setValue(skill.growth);
            if (skill.other > 0) targetSheet.getRange(`J${rowNum}`).setValue(skill.other);
        }
    });

    return insertedRows;
}

// ============================================================
// 一覧シートの更新（ハイパーリンク追加）
// ============================================================
function updateListSheet_(listSheet, listRow, targetSheet, data, id, url, sheetName, responseText) {
    function getSafeArray(key) {
        if (data[key] == null || data[key] === "") return [];
        if (Array.isArray(data[key])) return data[key];
        if (typeof data[key] === 'object') return Object.values(data[key]);
        return [data[key]];
    }
    // 技能配列の先頭から n 件を取り出す（足りない分は空欄）
    function pickSkills(key, n) {
        const arr = getSafeArray(key);
        return Array.from({ length: n }, (_, i) => arr[i] || "");
    }

    const pcName = data.pc_name || "";
    const occupation = data.shokugyo || data.shuzoku || "";
    const formattedDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");

    // HYPERLINK関数を使ってシートへのリンクを作成（名前に " が含まれても数式が壊れないようにエスケープ）
    const linkLabel = (pcName || sheetName).toString().replace(/"/g, '""');
    const hyperlinkFormula = `=HYPERLINK("#gid=${targetSheet.getSheetId()}", "${linkLabel}")`;

    const listData = [
        id,                 // A: ID
        url,                // B: URL
        hyperlinkFormula,   // C: PC名 (リンク)
        occupation,         // D: 職業
        formattedDate,      // E: 更新日
        data.pc_tags || "", // F: タグ
        data.SAN_Left || "",// G: 現在SAN
        data.NA1 || "",     // H: STR
        data.NA2 || "",     // I: CON
        data.NA3 || "",     // J: POW
        data.NA4 || "",     // K: DEX
        data.NA5 || "",     // L: APP
        data.NA6 || "",     // M: SIZ
        data.NA7 || "",     // N: INT
        data.NA8 || "",     // O: EDU
        data.NA9 || "",     // P: HP
        data.NA10 || "",    // Q: MP
        data.NA11 || "",    // R: 初期SAN
        data.NA12 || "",    // S: アイデア
        data.NA13 || "",    // T: 幸運
        ...pickSkills('TBAP', 12), // U〜AF: 戦闘技能
        ...pickSkills('TFAP', 12), // AG〜AR: 探索技能
        ...pickSkills('TAAP', 11), // AS〜BC: 行動技能
        ...pickSkills('TCAP', 5),  // BD〜BH: 交渉技能
        ...pickSkills('TKAP', 19), // BI〜CA: 知識技能
    ];

    // 技能ヘッダーを確認・設定（U列=21列目〜, 初回コンバート時のみ）
    if (listSheet.getRange(1, 21).getValue() === '') {
        const skillHeaderList = [
            '回避','キック','組み付き','こぶし(パンチ)','頭突き','投擲','マーシャルアーツ','拳銃','サブマシンガン','ショットガン','マシンガン','ライフル',
            '応急手当','鍵開け','隠す','隠れる','聞き耳','忍び歩き','写真術','精神分析','追跡','登攀','図書館','目星',
            '運転','機械修理','重機械操作','乗馬','水泳','製作','操縦','跳躍','電気修理','ナビゲート','変装',
            '言いくるめ','信用','説得','値切り','母国語',
            '医学','オカルト','化学','クトゥルフ神話','芸術','経理','考古学','コンピューター','心理学','人類学','生物学','地質学','電子工学','天文学','博物学','物理学','法律','薬学','歴史'
        ];
        listSheet.getRange(1, 21, 1, skillHeaderList.length).setValues([skillHeaderList]);
    }

    // CB列（80列目）のヘッダーを確認・設定
    if (listSheet.getRange(1, 80).getValue() === '') {
        listSheet.getRange(1, 80).setValue('hash(変更検知用)');
    }

    let row = listRow;
    if (row > 0) {
        listSheet.getRange(row, 1, 1, listData.length).setValues([listData]);
    } else {
        listSheet.appendRow(listData);
        row = listSheet.getLastRow();
    }
    listSheet.getRange(row, 3).setFormula(hyperlinkFormula);
    // コンバート時のJSONハッシュをCB列に保存（次回の差分検知用）
    listSheet.getRange(row, 80).setValue(computeHash(responseText));
}

// ============================================================
// チャットパレット・プロフィール・CCFOLIAデータの生成
// ============================================================
function writeChatPalette_(targetSheet, data, url, responseText, rowOffset) {
    try {
        // テンプレートから設定を読み取る（行を挿入した分だけ下にずれている）
        const palSettings = readPaletteSettings_(targetSheet, rowOffset);
        const paletteResult = generateChatPaletteData(data, url, palSettings);
        writePaletteResult_(targetSheet, rowOffset, paletteResult);

        // onEdit再生成用にURL・行のずれ・JSONをAB2/AB3/AB1に保存（ネット不要でオフライン対応）
        targetSheet.getRange('AB2').setValue(url);
        targetSheet.getRange('AB3').setValue(rowOffset);
        // 1セルに入るのは5万文字まで。超える場合は保存せず、設定変更時の自動再生成だけ使えない状態にする
        const CELL_CHAR_LIMIT = 50000;
        if (responseText.length <= CELL_CHAR_LIMIT) {
            targetSheet.getRange('AB1').setValue(responseText);
        }
        // チャパレ部分のエラーはここで確定させ、本体のコンバートは続ける
        SpreadsheetApp.flush();
    } catch (palErr) {
        // チャパレ生成に失敗しても本体のコンバートは完了しているのでエラーログのみ
        Logger.log('チャパレ生成エラー: ' + palErr.message);
    }
}

// ============================================================
// メニュー「URLからコンバート」の実行関数
// 「コンバートURL入力」シートのB列（2行目以降）に複数URLを入力できる
// 成功したURLは入力欄から消え、エラーと時間切れで未処理のURLだけが残る（再実行で続きから処理できる）
// ============================================================
function updateCharacterSheet_() {
    const startTime = Date.now();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const INPUT_SHEET_NAME = 'コンバートURL入力';

    const inputSheet = ss.getSheetByName(INPUT_SHEET_NAME);
    if (!inputSheet) {
        SpreadsheetApp.getUi().alert(`「${INPUT_SHEET_NAME}」シートが見つかりません。`);
        return;
    }

    const lastRow = inputSheet.getLastRow();
    const urlRange = lastRow >= 2 ? inputSheet.getRange(2, 2, lastRow - 1, 1) : null;
    // 1セルに改行区切りで複数URLが入っていても（セル編集中に貼り付けた場合など）1件ずつに分ける
    const urls = urlRange
        ? urlRange.getValues().flatMap(r => r[0].toString().split(/\s+/)).filter(u => u !== '')
        : [];

    if (urls.length === 0) {
        SpreadsheetApp.getUi().alert('B2セルにURLが入力されていません。');
        return;
    }

    const successNames = [];
    const errors = [];
    let remainingUrls = [];
    let quotaExceeded = false;
    for (let i = 0; i < urls.length; i++) {
        if (Date.now() - startTime > TIME_LIMIT_MS) {
            remainingUrls = urls.slice(i);
            break;
        }
        try {
            successNames.push(convertOneUrl(ss, urls[i], null));
        } catch (e) {
            if (isUrlFetchQuotaError_(e)) {
                // 失敗したURLも含めて未処理として残す
                quotaExceeded = true;
                remainingUrls = urls.slice(i);
                break;
            }
            errors.push({ url: urls[i], message: e.message });
        }
    }

    // 入力欄にはエラーと未処理のURLだけを残す
    const leftUrls = errors.map(e => e.url).concat(remainingUrls);
    urlRange.clearContent();
    if (leftUrls.length > 0) {
        inputSheet.getRange(2, 2, leftUrls.length, 1).setValues(leftUrls.map(u => [u]));
    }

    let msg = quotaExceeded ? `${URLFETCH_QUOTA_MESSAGE}\n\n`
        : remainingUrls.length > 0 ? `⏱ 時間の上限に近づいたため、途中で止めました。\n\n`
        : `コンバート・更新が完了しました！\n\n`;
    if (successNames.length > 0) msg += `✅ 成功 ${successNames.length}件:\n` + successNames.map(n => `　${n}`).join('\n');
    if (errors.length > 0) {
        msg += `\n\n❌ エラー ${errors.length}件:\n` + errors.map(e => `${e.url}\n  → ${e.message}`).join('\n');
        msg += `\n\nエラーになったURLは入力欄に残しています。確認して、不要なら削除してください。`;
    }
    if (remainingUrls.length > 0) {
        msg += `\n\n⏳ 未処理 ${remainingUrls.length}件`;
        if (!quotaExceeded) msg += `\nもう一度「URLからコンバート」を実行すると続きから処理します。`;
    }
    // 上限に達しているときは更新確認の取得も失敗するため行わない
    SpreadsheetApp.getUi().alert(quotaExceeded ? msg : msg + getUpdateNotice_());
}

// ============================================================
// メニュー「古いスプレッドシートから引っ越す」の実行関数
// 古いスプレッドシートの「シート一覧」B列からキャラシのURLを読み取り、
// このスプレッドシートの「コンバートURL入力」に追加する（古い方は読むだけで変更しない）
// ============================================================
function importFromOldSpreadsheet() {
    withDocumentLock_(importFromOldSpreadsheet_);
}

function importFromOldSpreadsheet_() {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const INPUT_SHEET_NAME = 'コンバートURL入力';
    const LIST_SHEET_NAME = 'シート一覧';

    const inputSheet = ss.getSheetByName(INPUT_SHEET_NAME);
    if (!inputSheet) {
        ui.alert(`「${INPUT_SHEET_NAME}」シートが見つかりません。`);
        return;
    }

    const answer = ui.prompt(
        '古いスプレッドシートから引っ越す',
        '古いスプレッドシートのURLを貼り付けてください。\n（古いスプレッドシートを開いたときの、ブラウザのアドレス欄のURLです）\n\n' +
        '古いスプレッドシートの「シート一覧」にあるキャラシのURLを、このスプレッドシートの「コンバートURL入力」に移します。古いスプレッドシートは変更しません。',
        ui.ButtonSet.OK_CANCEL
    );
    if (answer.getSelectedButton() !== ui.Button.OK) return;

    const m = answer.getResponseText().match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
    if (!m) {
        ui.alert('スプレッドシートのURLではないようです。\n「https://docs.google.com/spreadsheets/d/…」で始まるURLを貼り付けてください。');
        return;
    }
    if (m[1] === ss.getId()) {
        ui.alert('貼り付けたのは、今開いているスプレッドシートのURLです。\n古い方のスプレッドシートのURLを貼り付けてください。');
        return;
    }

    let oldSpreadsheet;
    try {
        oldSpreadsheet = SpreadsheetApp.openById(m[1]);
    } catch (e) {
        ui.alert('古いスプレッドシートを開けませんでした。\nURLが正しいか、今と同じGoogleアカウントで作ったスプレッドシートか確認してください。');
        return;
    }

    const oldListSheet = oldSpreadsheet.getSheetByName(LIST_SHEET_NAME);
    if (!oldListSheet || oldListSheet.getLastRow() < 2) {
        ui.alert(`古いスプレッドシートに「${LIST_SHEET_NAME}」が見つからないか、キャラが登録されていません。`);
        return;
    }

    // 古い一覧のB列から、キャラクター保管所のURLだけを取り出す（キャラ単位で重複を除く）
    const keyOf = url => { try { return parseSheetKey(url); } catch (e) { return null; } };
    const oldUrls = oldListSheet.getRange(2, 2, oldListSheet.getLastRow() - 1, 1).getValues()
        .map(r => r[0].toString().trim())
        .filter(u => u !== '' && keyOf(u));

    // すでにこのスプレッドシートに登録済み・入力済みのキャラは追加しない（2回実行しても重複しないように）
    const knownKeys = new Set();
    const listSheet = ss.getSheetByName(LIST_SHEET_NAME);
    if (listSheet && listSheet.getLastRow() >= 2) {
        listSheet.getRange(2, 1, listSheet.getLastRow() - 1, 2).getValues().forEach(r => {
            [r[0], r[1]].forEach(v => { const k = keyOf(v); if (k) knownKeys.add(k); });
        });
    }
    const inputLastRow = inputSheet.getLastRow();
    if (inputLastRow >= 2) {
        inputSheet.getRange(2, 2, inputLastRow - 1, 1).getValues()
            .flatMap(r => r[0].toString().split(/\s+/))
            .forEach(v => { const k = keyOf(v); if (k) knownKeys.add(k); });
    }

    const newUrls = [];
    oldUrls.forEach(u => {
        const k = keyOf(u);
        if (knownKeys.has(k)) return;
        knownKeys.add(k);
        newUrls.push(u);
    });

    if (newUrls.length === 0) {
        ui.alert(`古いスプレッドシートの ${oldUrls.length} 件は、すべてこのスプレッドシートに登録済み（または入力済み）です。`);
        return;
    }

    // 入力欄の最後のURLの下に追加する
    const startRow = Math.max(inputLastRow + 1, 2);
    inputSheet.getRange(startRow, 2, newUrls.length, 1).setValues(newUrls.map(u => [u]));

    let msg = `「${oldSpreadsheet.getName()}」から ${newUrls.length} 件のURLを「${INPUT_SHEET_NAME}」に移しました。`;
    const skipped = oldUrls.length - newUrls.length;
    if (skipped > 0) msg += `\n（${skipped} 件は登録済み・入力済みのため追加していません）`;
    msg += '\n\n続けて、メニューの「URLからコンバート」を押してください。\nキャラが多いときは、何回か実行すると続きから処理されます。';
    ui.alert(msg);
}
