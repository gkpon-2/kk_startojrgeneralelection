(function () {
  'use strict';
  const $ = (id) => document.getElementById(id);
  const R = window.JTRender;
  const SHEET_VOTE = '投票内容';
  const SHEET_VOTER = '投票者';
  const SHEET_HELP = '記入方法';
  const OUT_DPI = 200;

  let cfg = null, img = null;
  let voteRows = [], voters = [];

  init().catch((e) => showMessages([{ kind: 'err', text: '設定の読み込みに失敗しました：' + e.message }]));

  async function init() {
    cfg = await (await fetch('template/config.json', { cache: 'no-store' })).json();
    $('cfgTitle').textContent = cfg.title || '';
    img = await loadImage(cfg.image);
    for (const f of JTFonts.FONTS) {
      const o = document.createElement('option');
      o.value = f.family; o.textContent = f.label;
      $('fontSel').appendChild(o);
    }
    $('dlTemplate').onclick = downloadTemplate;
    $('xlsxInput').onchange = onFile;
    $('fontSel').onchange = drawPreview;
    $('colorSel').onchange = drawPreview;
    $('paperSel').onchange = () => { $('colorSel').disabled = $('paperSel').value !== 'color'; drawPreview(); };
    $('rowSel').onchange = drawPreview;
    $('exportBtn').onclick = exportPdf;
  }

  function loadImage(src) {
    return new Promise((res, rej) => {
      const i = new Image();
      i.onload = () => res(i);
      i.onerror = () => rej(new Error('用紙画像を読み込めません：' + src));
      i.src = src;
    });
  }

  // ---------- Excelテンプレート ----------
  async function downloadTemplate() {
    const cats = R.categories(cfg);
    // 名前欄の列は設定にある分だけ（コンビ・3人組は 名前2・名前3 が増える）
    const nameParts = ['名前', '名前2', '名前3'].filter((p) => cats.some((c) => c.parts[p]));
    const cols = ['名前', '理由', ...nameParts.filter((p) => p !== '名前')];
    const vrows = [['部門番号', '部門名', ...cols, '備考']];
    const grey = new Set();
    cats.forEach((c, i) => {
      const r = i + 1;
      const notes = [];
      const row = [c.no, c.category];
      cols.forEach((p, j) => {
        const fixed = c.fixed[p] || '';
        if (fixed) notes.push(`${p}は固定（書き換えても反映されません）`);
        if (fixed || !c.parts[p]) grey.add(`${r},${j + 2}`);
        row.push(fixed);
      });
      if (!c.parts['理由']) notes.push('理由欄なし');
      if (c.parts['名前2']) notes.push(c.parts['名前3'] ? '3人分の名前を入力' : '2人分の名前を入力');
      row.push(notes.join('／'));
      vrows.push(row);
    });

    const vf = R.voterFields(cfg);
    const voterHead = vf.map((f) => f.key);
    const textCols = vf.map((f, i) => (f.type === 'digits' || f.format === 'zip' ? i : -1)).filter((i) => i >= 0);

    const help = [['項目', '書き方']];
    help.push(['投票内容シート', '全ページ共通。部門番号と部門名は変更しないでください']);
    help.push(['投票者シート', '1行が応募用紙1枚になります。空の行は無視します']);
    help.push(['名前欄', '左ページの一覧の番号だけでもOK。一覧にないジュニアは「141+名前」と書きます']);
    for (const f of vf) {
      if (f.type === 'choice') help.push([f.key, f.options.map((o, i) => `${i + 1}=${o.label}`).join('　')]);
      else if (f.type === 'digits') help.push([f.key, `数字${f.count}桁（ハイフンは入れても入れなくてもOK）`]);
      else if (f.format === 'zip') help.push([f.key, '数字7桁。ハイフンはあってもなくてもOK（出力は 123-4567 の形になります）']);
    }

    const blob = await XlsxMini.write([
      { name: SHEET_VOTE, rows: vrows, widths: [10, 30, 16, 44, ...nameParts.slice(1).map(() => 16), 30], headerStyle: true, greyCells: grey },
      { name: SHEET_VOTER, rows: [voterHead], widths: vf.map((f) => (f.key === '住所' ? 48 : 14)), headerStyle: true, textCols },
      { name: SHEET_HELP, rows: help, widths: [18, 80], headerStyle: true },
    ]);
    try { await download(blob, `ジュニア大賞_入力テンプレート_${cfg.year || ''}.xlsx`); }
    catch (e) { showMessages([{ kind: 'err', text: 'ダウンロードできませんでした：' + e.message }]); }
  }

  // ---------- 読み込みと検証 ----------
  async function onFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    $('xlsxName').textContent = file.name;
    const msgs = [];
    try {
      const book = await XlsxMini.read(await file.arrayBuffer());
      if (!book[SHEET_VOTE]) msgs.push({ kind: 'err', text: `「${SHEET_VOTE}」シートがありません` });
      if (!book[SHEET_VOTER]) msgs.push({ kind: 'err', text: `「${SHEET_VOTER}」シートがありません` });
      if (msgs.length) return fail(msgs);

      voteRows = XlsxMini.toObjects(book[SHEET_VOTE]);
      voters = XlsxMini.toObjects(book[SHEET_VOTER]);

      const vf = R.voterFields(cfg);
      const head = (book[SHEET_VOTER][0] || []).map((h) => String(h).trim());
      const missing = vf.map((f) => f.key).filter((k) => !head.includes(k));
      if (missing.length) msgs.push({ kind: 'err', text: `投票者シートに列がありません：${missing.join('、')}` });
      if (!voters.length) msgs.push({ kind: 'err', text: '投票者シートにデータがありません' });

      const rowErrs = [], rowWarns = [];
      for (const v of voters) {
        for (const f of vf) {
          const val = String(v[f.key] ?? '').trim();
          if (!val) { rowWarns.push(`${v._row}行目：${f.key}が空欄です`); continue; }
          if (f.type === 'choice') {
            const n = Number(val);
            if (!Number.isInteger(n) || n < 1 || n > f.options.length) rowErrs.push(`${v._row}行目：${f.key}は1〜${f.options.length}の数字で入力してください（「${val}」）`);
          }
        }
      }
      if (rowErrs.length) msgs.push({ kind: 'err', text: '入力エラー', list: rowErrs });

      const { warnings, notices, values } = R.resolveValues(cfg, voteRows, voters[0] || {});
      const fixedWarns = warnings;
      if (fixedWarns.length) msgs.push({ kind: 'warn', text: '固定部門の書き換えを検出しました（出力には反映されません）', list: fixedWarns });
      if (notices.length) msgs.push({ kind: 'warn', text: '投票内容シートの確認', list: notices });
      const voteFields = cfg.fields.filter((f) => f.sheet === 'vote' && !f.fixed);
      const filled = (part) => voteFields.filter((f) => f.part === part && values[f.id]).length;
      const total = (part) => voteFields.filter((f) => f.part === part).length;
      msgs.push({ kind: 'ok', text: `投票内容：名前 ${filled('名前')}/${total('名前')}件、理由 ${filled('理由')}/${total('理由')}件を読み込みました` });
      if (rowWarns.length) msgs.push({ kind: 'warn', text: '空欄があります（その欄は何も書かずに出力します）', list: rowWarns });

      if (msgs.some((m) => m.kind === 'err')) return fail(msgs);
      msgs.unshift({ kind: 'ok', text: `${voters.length}枚分を読み込みました` });
      showMessages(msgs);

      const sel = $('rowSel');
      sel.innerHTML = '';
      voters.forEach((v, i) => {
        const o = document.createElement('option');
        o.value = i;
        o.textContent = `${i + 1}枚目：${v['氏名'] || v[vf[0] && vf[0].key] || ''}`;
        sel.appendChild(o);
      });
      $('previewCard').hidden = false;
      drawPreview();
    } catch (err) {
      fail([{ kind: 'err', text: '読み込みに失敗しました：' + err.message }]);
    } finally {
      e.target.value = '';
    }
  }

  function fail(msgs) {
    $('previewCard').hidden = true;
    showMessages(msgs);
  }

  function showMessages(msgs, target) {
    const box = target || $('messages');
    box.innerHTML = '';
    for (const m of msgs) {
      const d = document.createElement('div');
      d.className = 'msg ' + m.kind;
      d.textContent = m.text;
      if (m.list && m.list.length) {
        const ul = document.createElement('ul');
        m.list.slice(0, 30).forEach((t) => { const li = document.createElement('li'); li.textContent = t; ul.appendChild(li); });
        if (m.list.length > 30) { const li = document.createElement('li'); li.textContent = `ほか${m.list.length - 30}件`; ul.appendChild(li); }
        d.appendChild(ul);
      }
      box.appendChild(d);
    }
  }

  // ---------- プレビュー ----------
  let drawToken = 0;
  async function drawPreview() {
    const token = ++drawToken;
    const family = await JTFonts.ensure($('fontSel').value);
    if (token !== drawToken) return;
    const i = Number($('rowSel').value) || 0;
    const c = $('preview');
    const W = Math.round(cfg.page.widthMm / 25.4 * 150);
    const H = Math.round(cfg.page.heightMm / 25.4 * 150);
    c.width = W; c.height = H;
    const { values } = R.resolveValues(cfg, voteRows, voters[i]);
    const warns = R.renderPage(c.getContext('2d'), W, H, cfg, img, values, { font: family, color: $('colorSel').value, paper: $('paperSel').value, seed: i + 1 });
    showMessages(warns.length ? [{ kind: 'warn', text: 'この1枚の注意点', list: warns }] : [], $('rowWarnings'));
  }

  // ---------- PDF出力 ----------
  async function exportPdf() {
    const btn = $('exportBtn');
    btn.disabled = true;
    const mode = document.querySelector('input[name=outMode]:checked').value;
    const family = await JTFonts.ensure($('fontSel').value);
    const color = $('colorSel').value;
    const paper = $('paperSel').value;
    const { PDFDocument } = PDFLib;
    const W = Math.round(cfg.page.widthMm / 25.4 * OUT_DPI);
    const H = Math.round(cfg.page.heightMm / 25.4 * OUT_DPI);
    const pw = cfg.page.widthMm / 25.4 * 72, ph = cfg.page.heightMm / 25.4 * 72;
    const canvas = document.createElement('canvas');
    canvas.width = W; canvas.height = H;
    const ctx = canvas.getContext('2d');
    const allWarn = [];
    try {
      const single = mode === 'single' ? await PDFDocument.create() : null;
      const zip = mode === 'zip' ? new JSZip() : null;
      for (let i = 0; i < voters.length; i++) {
        $('progress').textContent = `${i + 1} / ${voters.length} 枚目を作成中…`;
        const { values } = R.resolveValues(cfg, voteRows, voters[i]);
        const w = R.renderPage(ctx, W, H, cfg, img, values, { font: family, color, paper, seed: i + 1 });
        w.forEach((t) => allWarn.push(`${i + 1}枚目：${t}`));
        const jpg = await new Promise((res) => canvas.toBlob((b) => res(b), 'image/jpeg', 0.9));
        const bytes = new Uint8Array(await jpg.arrayBuffer());
        const doc = single || await PDFDocument.create();
        const im = await doc.embedJpg(bytes);
        const page = doc.addPage([pw, ph]);
        page.drawImage(im, { x: 0, y: 0, width: pw, height: ph });
        if (zip) {
          const name = String(voters[i]['氏名'] || '').replace(/[\\/:*?"<>|\s]/g, '') || 'noname';
          zip.file(`${String(i + 1).padStart(3, '0')}_${name}.pdf`, await doc.save());
        }
        await new Promise((r) => setTimeout(r, 0));
      }
      const stamp = new Date().toISOString().slice(0, 10);
      if (single) await download(new Blob([await single.save()], { type: 'application/pdf' }), `ジュニア大賞_${voters.length}枚_${stamp}.pdf`);
      else await download(await zip.generateAsync({ type: 'blob' }), `ジュニア大賞_${voters.length}枚_${stamp}.zip`);
      $('progress').textContent = `${voters.length}枚を出力しました`;
      if (allWarn.length) showMessages([{ kind: 'warn', text: '出力時の注意点', list: allWarn }], $('rowWarnings'));
    } catch (e) {
      $('progress').textContent = '出力に失敗しました：' + e.message;
    } finally {
      btn.disabled = false;
    }
  }

  async function download(blob, name) {
    const r = await JTSave(blob, name);
    if (r === 'declined') $('progress').textContent = '保存をキャンセルしました';
  }
})();
