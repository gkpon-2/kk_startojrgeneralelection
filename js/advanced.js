(function () {
  'use strict';
  const $ = (id) => document.getElementById(id);
  const R = window.JTRender, P = window.JTPlan;
  const OUT_DPI = 200;
  const SHEET_VOTER = '投票者', SHEET_REASON = '理由', SHEET_VOTE = '投票内容', SHEET_HELP = '記入方法';

  let cfg = null, img = null;
  let csv = null;            // {head, list}
  let voters = [], voteRows = [], reasons = {};
  let planRes = null;
  let cancel = false;

  init().catch((e) => show([{ kind: 'err', text: '設定の読み込みに失敗しました：' + e.message }]));

  async function init() {
    cfg = await (await fetch('template/config.json', { cache: 'no-store' })).json();
    $('cfgTitle').textContent = cfg.title || '';
    img = await new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = () => rej(new Error('用紙画像を読み込めません')); i.src = cfg.image; });
    for (const f of JTFonts.FONTS) { const o = document.createElement('option'); o.value = f.family; o.textContent = f.label; $('fontSel').appendChild(o); }
    // 応募用紙があるのは12月号と1月号だけ
    for (const m of [12, 1]) { const o = document.createElement('option'); o.value = m; o.textContent = `${m}月号`; $('monthSel').appendChild(o); }
    $('monthSel').value = '12';

    $('dlTemplate').onclick = downloadTemplate;
    $('csvInput').onchange = onCsv;
    $('xlsxInput').onchange = onXlsx;
    $('monthSel').onchange = replan;
    for (const id of ['fontSel', 'colorSel']) $(id).onchange = drawPreview;
    $('paperSel').onchange = () => { $('colorSel').disabled = $('paperSel').value !== 'color'; drawPreview(); };
    $('pageNo').onchange = drawPreview;
    $('prevPage').onclick = () => { $('pageNo').value = Math.max(1, Number($('pageNo').value) - 1); drawPreview(); };
    $('nextPage').onclick = () => { $('pageNo').value = Math.min(voters.length, Number($('pageNo').value) + 1); drawPreview(); };
    $('dlMapping').onclick = () => saveCsv(P.mapping(planRes, chunkSize(), fileName), `対応表_${month()}月号.csv`);
    $('dlRemaining').onclick = () => saveCsv(P.remainingCsv(csv.head, planRes), `残り_${month()}月号のあと.csv`);
    $('exportBtn').onclick = exportZip;
    $('cancelBtn').onclick = () => { cancel = true; };
  }

  const month = () => Number($('monthSel').value);
  const chunkSize = () => Math.max(10, Math.min(500, Number($('chunkSize').value) || 100));
  const pad = (n) => String(n).padStart(String(voters.length).length, '0');
  // lo,hi：今回出力する範囲（0始まり・hiは含まない）。範囲の一部だけ出すときは実際に入るページで名前を付ける
  function fileName(c, lo = 0, hi = voters.length) {
    const a = Math.max(lo, c * chunkSize()) + 1, b = Math.min(hi, voters.length, (c + 1) * chunkSize());
    return `ジュニア大賞_${month()}月号_${pad(a)}-${pad(b)}.pdf`;
  }

  // ---------- テンプレート ----------
  async function downloadTemplate() {
    const vf = R.voterFields(cfg);
    const textCols = vf.map((f, i) => (f.format === 'zip' ? i : -1)).filter((i) => i >= 0);
    // 投票内容シートは通常版と同じ形（全部門）。交換CSVで埋まらないページはこの内容になる
    const cats = R.categories(cfg);
    const nameParts = ['名前', '名前2', '名前3'].filter((p) => cats.some((c) => c.parts[p]));
    const cols = ['名前', '理由', ...nameParts.filter((p) => p !== '名前')];
    const vrows = [['部門番号', '部門名', ...cols, '備考']];
    const grey = new Set();
    cats.forEach((c, i) => {
      const notes = [];
      const row = [c.no, c.category];
      cols.forEach((p, j) => {
        const fixed = c.fixed[p] || '';
        if (fixed) notes.push(`${p}は固定（書き換えても反映されません）`);
        if (fixed || !c.parts[p]) grey.add(`${i + 1},${j + 2}`);
        row.push(fixed);
      });
      if (!c.parts['理由']) notes.push('理由欄なし');
      if (c.parts['名前2']) notes.push(c.parts['名前3'] ? '3人分の名前を入力' : '2人分の名前を入力');
      row.push(notes.join('／'));
      vrows.push(row);
    });
    const help = [['項目', '書き方'],
      ['交換CSV', '交換管理アプリの書き出しをそのまま読み込みます。進行状況が「記入中」の行だけ使います。項目B（部門）・項目C（ジュニア）・数量を使います'],
      ['理由シート', '交換先アカウント・部門番号・ジュニア番号・理由を1行に1つ。交換先アカウントは交換CSVのURLをそのまま貼るか、@なしのIDでもOK。同じジュニアでも交換相手ごとに別の理由を使えます。アカウントを空欄にした理由は、その部門・ジュニアで専用の理由がない交換すべてに使います。アカウントを「自分」にした理由は、投票内容シートの内容になるページ（自分の票）で、名前が一致するときにループします。同じ部門・ジュニアの理由は1枚ごとに順番にループします。1行だけなら全部同じ理由になります。登録がない交換は投票内容シートの理由を使います。コンビ・3人組は1人目の番号で登録します'],
      ['投票者シート', '1行が応募用紙1枚。今回作る枚数ぶんの行を入れます'],
      ['投票内容シート', '通常版と同じです。交換CSVの票が入らないページ（票数が投票者数に満たない部門の残りのページ、CSVにない部門）は、この内容で埋まります'],
      ['固定の部門', '恋人・先生・キャスター・MC・手がきれいは39番固定です。交換CSVでも39番宛ての交換だけ使え、理由は理由シートから入ります。39番以外の交換はエラーになります'],
      ['残りCSV', '今回の枚数に入りきらなかった票です。次の号ではこれを交換CSVとして読み込むと、理由のループも続きから使います']];
    for (const f of vf) if (f.type === 'choice') help.push([f.key, f.options.map((o, i) => `${i + 1}=${o.label}`).join('　')]);
    const blob = await XlsxMini.write([
      { name: SHEET_VOTER, rows: [vf.map((f) => f.key)], widths: vf.map((f) => (f.key === '住所' ? 48 : 14)), headerStyle: true, textCols },
      { name: SHEET_REASON, rows: [['交換先アカウント', '部門番号', 'ジュニア番号', '理由']], widths: [28, 10, 12, 60], headerStyle: true },
      { name: SHEET_VOTE, rows: vrows, widths: [10, 30, 16, 44, ...nameParts.slice(1).map(() => 16), 30], headerStyle: true, greyCells: grey },
      { name: SHEET_HELP, rows: help, widths: [16, 90], headerStyle: true },
    ]);
    await JTSave(blob, `ジュニア大賞_上級モード_テンプレート_${cfg.year || ''}.xlsx`);
  }

  // ---------- 読み込み ----------
  async function onCsv(e) {
    const f = e.target.files[0]; if (!f) return;
    $('csvName').textContent = f.name;
    try { csv = P.readExchanges(P.parseCsv(P.decode(new Uint8Array(await f.arrayBuffer())))); replan(); }
    catch (err) { csv = null; show([{ kind: 'err', text: 'CSVを読めませんでした：' + err.message }]); hidePlan(); }
    finally { e.target.value = ''; }
  }

  async function onXlsx(e) {
    const f = e.target.files[0]; if (!f) return;
    $('xlsxName').textContent = f.name;
    try {
      const book = await XlsxMini.read(await f.arrayBuffer());
      if (!book[SHEET_VOTER]) throw new Error(`「${SHEET_VOTER}」シートがありません`);
      voters = XlsxMini.toObjects(book[SHEET_VOTER]);
      voteRows = book[SHEET_VOTE] ? XlsxMini.toObjects(book[SHEET_VOTE]) : [];
      reasons = {};
      for (const r of XlsxMini.toObjects(book[SHEET_REASON] || [])) {
        const no = P.leadNum(r['部門番号']);
        const jr = P.leadNum(r['ジュニア番号']) || String(r['ジュニア番号'] ?? '').trim();
        const t = String(r['理由'] ?? '').trim();
        const k = `${P.acctKey(r['交換先アカウント'])}|${no}|${jr}`;
        if (no && jr && t) (reasons[k] = reasons[k] || []).push(t);
      }
      replan();
    } catch (err) { voters = []; show([{ kind: 'err', text: 'Excelを読めませんでした：' + err.message }]); hidePlan(); }
    finally { e.target.value = ''; }
  }

  function hidePlan() { $('planCard').hidden = true; $('previewCard').hidden = true; }

  // ---------- 割り当て ----------
  function replan() {
    const msgs = [];
    if (!csv) msgs.push({ kind: 'warn', text: '交換CSVを読み込んでください' });
    if (!voters.length) msgs.push({ kind: 'warn', text: 'Excel（投票者シートに1行以上）を読み込んでください' });
    if (msgs.length) { show(msgs); hidePlan(); return; }

    // 投票者のチェック
    const vf = R.voterFields(cfg), errs = [];
    for (const v of voters) for (const f of vf) {
      const val = String(v[f.key] ?? '').trim();
      if (f.type === 'choice' && val) {
        const n = Number(val);
        if (!Number.isInteger(n) || n < 1 || n > f.options.length) errs.push(`投票者${v._row}行目：${f.key}は1〜${f.options.length}で入力してください`);
      }
    }
    const nReasons = Object.values(reasons).reduce((a, b) => a + b.length, 0);
    const active = csv.hasStatus ? csv.list.filter((x) => x.status === P.TARGET_STATUS).length : csv.list.length;
    msgs.push({ kind: 'ok', text: `交換${active}件（進行状況「${P.TARGET_STATUS}」）・投票者${voters.length}人・理由${nReasons}件（${Object.keys(reasons).length}組）を読み込みました` + (csv.hasStatus && active < csv.list.length ? `。ほか${csv.list.length - active}件は「${P.TARGET_STATUS}」ではないので使いません` : '') });
    if (!csv.hasStatus) msgs.push({ kind: 'warn', text: 'CSVに「進行状況」の列がないので、全行を使います' });
    const ownN = Object.keys(reasons).filter((k) => k.startsWith(P.OWN + '|'));
    if (ownN.length) msgs.push({ kind: 'ok', text: `自分の票の理由ループ：${ownN.map((k) => { const [, no, jr] = k.split('|'); return `部門${no}・${jr}番（${reasons[k].length}種類）`; }).join('、')}` });
    const unused = P.unusedOwn(cfg, reasons, R.resolveValues(cfg, voteRows, voters[0] || {}).values);
    if (unused.length) msgs.push({ kind: 'warn', text: '理由シートの「自分」の行の確認', list: unused });
    else if (!active) msgs.push({ kind: 'warn', text: `進行状況が「${P.TARGET_STATUS}」の交換がありません` });
    if (errs.length) msgs.push({ kind: 'err', text: '投票者シートの入力エラー', list: errs });
    show(msgs);

    planRes = P.plan(cfg, csv.list, voters.length, month(), reasons);
    renderPlan();
    $('planCard').hidden = false;
    $('previewCard').hidden = !!errs.length;
    $('pageNo').max = voters.length;
    $('fromPage').max = $('toPage').max = voters.length;
    $('toPage').value = voters.length;
    if (Number($('fromPage').value) > voters.length) $('fromPage').value = 1;
    drawPreview();
  }

  function renderPlan() {
    const N = voters.length;
    const carry = planRes.summary.reduce((a, s) => a + s.carry, 0);
    const maxNeed = Math.max(0, ...planRes.summary.map((s) => s.need));
    $('planInfo').innerHTML = '';
    const stat = document.createElement('div'); stat.className = 'stat';
    const add = (label, v) => { const s = document.createElement('span'); s.append(label + ' '); const b = document.createElement('b'); b.textContent = v; s.append(b); stat.append(s); };
    add('今回の枚数', `${N}枚`);
    add('いちばん多い部門', `${maxNeed}票`);
    add('次の号へ繰り越し', `${carry}票`);
    $('planInfo').append(stat);

    const t = $('summaryTable');
    t.innerHTML = '';
    const tr = (cells, th) => { const r = document.createElement('tr'); cells.forEach((c) => { const el = document.createElement(th ? 'th' : 'td'); if (c && c.cls) el.className = c.cls; if (c && c.node) el.append(c.node); else el.textContent = c && c.text != null ? c.text : c; r.append(el); }); t.append(r); };
    tr(['部門', '部門名', '必要', '今回', '繰り越し', '投票内容で埋める', ''], true);
    for (const s of planRes.summary) {
      const pill = document.createElement('span');
      pill.className = 'pill ' + (s.carry ? 'carry' : 'ok');
      pill.textContent = s.carry ? '次の号へ' : '今回で完了';
      tr([s.no, { text: s.category, cls: 'cat' }, { text: s.need, cls: 'n' }, { text: s.used, cls: 'n' }, { text: s.carry, cls: 'n' }, { text: s.fallback ? `${s.fallback}枚` : '—', cls: 'n' }, { node: pill }]);
    }

    const errs = planRes.issues.filter((i) => i.level === 'err').map((i) => i.text);
    const warns = planRes.issues.filter((i) => i.level === 'warn').map((i) => i.text);
    const box = [];
    if (errs.length) box.push({ kind: 'err', text: '要確認（この交換は今回使わず、残りCSVにそのまま残します）', list: errs });
    if (warns.length) box.push({ kind: 'warn', text: '注意', list: warns });
    show(box, $('issues'));
  }

  // ---------- 1ページ分の値 ----------
  function valuesFor(i) {
    const { values } = R.resolveValues(cfg, voteRows, voters[i] || {});
    Object.assign(values, P.pageVotes(planRes, reasons, i));
    return Object.assign(values, P.ownReasons(cfg, planRes, reasons, values, i));
  }

  let token = 0;
  async function drawPreview() {
    if (!planRes || !voters.length) return;
    const t = ++token;
    const family = await JTFonts.ensure($('fontSel').value);
    if (t !== token) return;
    const i = Math.min(voters.length, Math.max(1, Number($('pageNo').value) || 1)) - 1;
    $('pageNo').value = i + 1;
    const c = $('preview');
    c.width = Math.round(cfg.page.widthMm / 25.4 * 150);
    c.height = Math.round(cfg.page.heightMm / 25.4 * 150);
    const w = R.renderPage(c.getContext('2d'), c.width, c.height, cfg, img, valuesFor(i), opts(family, i));
    show(w.length ? [{ kind: 'warn', text: `${i + 1}枚目の注意点`, list: w }] : [], $('pageWarnings'));
  }

  function opts(family, i, transparent) {
    return { font: family, color: $('colorSel').value, paper: $('paperSel').value, seed: i + 1, transparent };
  }

  // ---------- 出力 ----------
  async function exportZip() {
    const N = voters.length;
    const from = Math.max(1, Number($('fromPage').value) || 1) - 1;
    const to = Math.min(N, Number($('toPage').value) || N);
    if (from >= to) { $('progress').textContent = '「何枚目から」「何枚目まで」を確認してください'; return; }
    cancel = false;
    $('exportBtn').disabled = true; $('cancelBtn').hidden = false;
    document.querySelector('.progress').hidden = false;
    const family = await JTFonts.ensure($('fontSel').value);
    const { PDFDocument } = PDFLib;
    const W = Math.round(cfg.page.widthMm / 25.4 * OUT_DPI), H = Math.round(cfg.page.heightMm / 25.4 * OUT_DPI);
    const pw = cfg.page.widthMm / 25.4 * 72, ph = cfg.page.heightMm / 25.4 * 72;

    // 背景は1回だけ作って、各PDFで全ページ共有する
    const bgC = document.createElement('canvas'); bgC.width = W; bgC.height = H;
    const bctx = bgC.getContext('2d'); bctx.fillStyle = '#fff'; bctx.fillRect(0, 0, W, H);
    bctx.drawImage(R.paperOf(img, $('paperSel').value), 0, 0, W, H);
    const bgBytes = new Uint8Array(await (await blobOf(bgC, 'image/jpeg', 0.88)).arrayBuffer());

    const layer = document.createElement('canvas'); layer.width = W; layer.height = H;
    const lctx = layer.getContext('2d');
    const zip = new JSZip();
    const size = chunkSize();
    const warns = [];
    const started = Date.now();
    try {
      for (let c = Math.floor(from / size); c * size < to; c++) {
        const a = Math.max(from, c * size), b = Math.min(to, (c + 1) * size);
        const doc = await PDFDocument.create();
        const bg = await doc.embedJpg(bgBytes);
        for (let i = a; i < b; i++) {
          if (cancel) throw new Error('中止しました');
          const w = R.renderPage(lctx, W, H, cfg, null, valuesFor(i), opts(family, i, true));
          w.forEach((t) => warns.push(`${i + 1}枚目：${t}`));
          const png = await doc.embedPng(new Uint8Array(await (await blobOf(layer, 'image/png')).arrayBuffer()));
          const page = doc.addPage([pw, ph]);
          page.drawImage(bg, { x: 0, y: 0, width: pw, height: ph });
          page.drawImage(png, { x: 0, y: 0, width: pw, height: ph });
          const done = i - from + 1, total = to - from;
          if (done % 5 === 0 || done === total) {
            const sec = (Date.now() - started) / 1000, rest = Math.round(sec / done * (total - done));
            $('progress').textContent = `${i + 1}枚目まで作成（${done}/${total}）・残り約${rest}秒`;
            $('bar').style.width = `${(done / total) * 100}%`;
            await new Promise((r) => setTimeout(r, 0));
          }
        }
        zip.file(fileName(c, from, to), await doc.save());
      }
      zip.file(`対応表_${month()}月号.csv`, P.toCsv(P.mapping(planRes, size, (c) => fileName(c, from, to))));
      zip.file(`残り_${month()}月号のあと.csv`, P.toCsv(P.remainingCsv(csv.head, planRes)));
      $('progress').textContent = 'ZIPにまとめています…';
      const blob = await zip.generateAsync({ type: 'blob' });
      const r = await JTSave(blob, `ジュニア大賞_${month()}月号_${pad(from + 1)}-${pad(to)}.zip`);
      $('progress').textContent = r === 'declined' ? '保存をキャンセルしました' : `${to - from}枚を出力しました（${(blob.size / 1048576).toFixed(1)}MB）`;
      if (warns.length) show([{ kind: 'warn', text: '出力時の注意点', list: warns }], $('pageWarnings'));
    } catch (e) {
      $('progress').textContent = e.message === '中止しました' ? '中止しました。「何枚目から」を変えると続きから作れます' : '出力に失敗しました：' + e.message;
    } finally {
      $('exportBtn').disabled = false; $('cancelBtn').hidden = true;
    }
  }

  function blobOf(canvas, type, q) { return new Promise((res) => canvas.toBlob(res, type, q)); }
  async function saveCsv(rows, name) {
    try { await JTSave(new Blob([P.toCsv(rows)], { type: 'text/csv' }), name); }
    catch (e) { $('progress').textContent = '保存できませんでした：' + e.message; }
  }

  function show(msgs, target) {
    const box = target || $('messages');
    box.innerHTML = '';
    for (const m of msgs) {
      const d = document.createElement('div');
      d.className = 'msg ' + m.kind; d.textContent = m.text;
      if (m.list && m.list.length) {
        const ul = document.createElement('ul');
        m.list.slice(0, 40).forEach((t) => { const li = document.createElement('li'); li.textContent = t; ul.append(li); });
        if (m.list.length > 40) { const li = document.createElement('li'); li.textContent = `ほか${m.list.length - 40}件`; ul.append(li); }
        d.append(ul);
      }
      box.append(d);
    }
  }
})();
