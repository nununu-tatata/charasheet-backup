// ==========================================
// 配布用スプレッドシートの初期化（管理者用・配布用コピーで一度だけ実行する）
// テンプレート・シート一覧・コンバートURL入力 以外のシートを削除し、一覧と入力欄を空にする
// 通常は push 対象外。使うときだけルートにコピーして push し、実行後に削除して push し直す
// ==========================================
function resetForDistribution() {
    const KEEP_SHEETS = ['テンプレート', 'シート一覧', 'コンバートURL入力'];
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();

    const targets = ss.getSheets().filter(s => !KEEP_SHEETS.includes(s.getName()));
    const answer = ui.alert(
        '配布用に初期化します',
        `「${ss.getName()}」から ${targets.length} 件のシートを削除し、シート一覧と入力欄を空にします。\n元に戻せません。本当に実行しますか？`,
        ui.ButtonSet.YES_NO
    );
    if (answer !== ui.Button.YES) return;

    targets.forEach(s => ss.deleteSheet(s));

    // シート一覧はヘッダー（1行目）だけ残す
    const listSheet = ss.getSheetByName('シート一覧');
    if (listSheet && listSheet.getLastRow() > 1) {
        listSheet.deleteRows(2, listSheet.getLastRow() - 1);
    }
    // 入力欄（B2以降）を空にする
    const inputSheet = ss.getSheetByName('コンバートURL入力');
    if (inputSheet && inputSheet.getLastRow() > 1) {
        inputSheet.getRange(2, 2, inputSheet.getLastRow() - 1, 1).clearContent();
    }

    ui.alert(`初期化が完了しました。${targets.length} 件のシートを削除しました。`);
}
