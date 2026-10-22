// ==========================================
// 配布ページの動作
// ==========================================

// 配布用スプレッドシートのコピー用URL（共有リンクの /edit... を /copy に置き換えたもの）
// 空欄のあいだは「準備中」と表示する
const SPREADSHEET_COPY_URL = 'https://docs.google.com/spreadsheets/d/1hEHd0JRUxbp1wAb3k7IVArZElsJWIxQZFwCcSxBGsws/copy';

// --- アクセス解析のイベント記録（広告ブロック等で gtag がない場合は何もしない） ---
function trackEvent(name, params) {
    if (typeof gtag === 'function') gtag('event', name, params || {});
}

// --- コピーボタン ---
const copyBtn = document.getElementById('copy-btn');
if (SPREADSHEET_COPY_URL) {
    copyBtn.href = SPREADSHEET_COPY_URL;
    copyBtn.addEventListener('click', () => trackEvent('copy_spreadsheet'));
} else {
    copyBtn.classList.add('disabled');
    copyBtn.textContent = '📄 準備中です';
}

// --- バージョン情報・更新履歴（version.json から読み込む） ---
fetch('version.json', { cache: 'no-cache' })
    .then(res => res.json())
    .then(info => {
        document.getElementById('version-badge').textContent = `最新版 v${info.version}`;
        document.getElementById('last-updated').textContent = info.date;
        const body = document.getElementById('changelog-body');
        body.innerHTML = (info.history || []).map(h => `
            <div style="margin-bottom: 12px;">
                <div style="font-weight: 700;">v${h.version}（${h.date}）</div>
                <ul>${h.changes.map(c => `<li>${c}</li>`).join('')}</ul>
            </div>`).join('');
    })
    .catch(() => {
        document.getElementById('version-badge').style.display = 'none';
        document.getElementById('changelog-body').textContent = '更新履歴を読み込めませんでした。';
    });

// --- ブックマークレット（コードは bookmarklet-code.js。bookmarklet/build.js で生成） ---
const bookmarkletLink = document.getElementById('bookmarklet-link');
const bookmarkletCode = document.getElementById('bookmarklet-code');
bookmarkletLink.href = BOOKMARKLET_HREF;
bookmarkletCode.value = BOOKMARKLET_HREF;
bookmarkletLink.addEventListener('dragstart', () => trackEvent('drag_bookmarklet'));
bookmarkletLink.addEventListener('click', e => {
    e.preventDefault();
    alert('このボタンはクリックせず、ブックマークバーへドラッグして登録してください。\n\n登録したら、キャラクター保管所のマイリストを開いてからブックマークを押してください。');
});
document.getElementById('bookmarklet-copy').addEventListener('click', e => {
    navigator.clipboard.writeText(BOOKMARKLET_HREF).then(() => {
        e.target.textContent = '✅ コピーしました';
        trackEvent('copy_bookmarklet_code');
    });
});

// --- テーマ切替 ---
const btnTheme = document.getElementById('btn-theme-toggle');
btnTheme.addEventListener('click', () => {
    const html = document.documentElement;
    const isDark = html.getAttribute('data-theme') === 'dark';
    html.setAttribute('data-theme', isDark ? '' : 'dark');
    btnTheme.textContent = isDark ? '🌙' : '☀️';
});

// --- UI制御 ---
function openModal(modalId) { const m = document.getElementById(modalId); if (m) m.classList.add('open'); }
function closeModal(modalId) { const m = document.getElementById(modalId); if (m) m.classList.remove('open'); }
window.addEventListener('click', (e) => { if (e.target.classList.contains('modal-overlay')) e.target.classList.remove('open'); });

function toggleShareMenu(event) {
    event.stopPropagation();
    const menu = document.getElementById('shareMenu');
    if (menu) menu.classList.toggle('show');
}
window.addEventListener('click', () => { const m = document.getElementById('shareMenu'); if (m) m.classList.remove('show'); });

function shareCopyLink() {
    const url = window.location.href;
    navigator.clipboard.writeText(url).then(() => {
        const t = document.getElementById('shareTooltip');
        t.classList.add('show');
        setTimeout(() => t.classList.remove('show'), 2000);
    });
}
function shareToX() {
    const url = encodeURIComponent(window.location.href);
    const text = encodeURIComponent(document.title);
    window.open(`https://twitter.com/intent/tweet?url=${url}&text=${text}`, '_blank');
}
function shareToBluesky() {
    const url = encodeURIComponent(window.location.href);
    const text = encodeURIComponent(document.title);
    window.open(`https://bsky.app/intent/compose?text=${text}%0A${url}`, '_blank');
}
function shareToMisskey() {
    const url = encodeURIComponent(window.location.href);
    const text = encodeURIComponent(document.title);
    window.open(`https://misskey-hub.net/share/?text=${text}&url=${url}`, '_blank');
}
function shareToDiscord() {
    shareCopyLink();
    alert("Discordへの直接シェア機能はありません。リンクをコピーしましたので貼り付けてください！");
}
