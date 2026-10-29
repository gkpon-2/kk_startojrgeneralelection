// 上級モードの割り当てロジック（画面に依存しない）
(function (g) {
  'use strict';

  // ---------- CSV ----------
  function parseCsv(text) {
    text = text.replace(/^﻿/, '');
    const rows = [];
    let row = [], cell = '', q = false;
    for (let i = 0; i < text.length; i++) {
      const c = text[i];
      if (q) {
        if (c === '"') { if (text[i + 1] === '"') { cell += '"'; i++; } else q = false; }
        else cell += c;
      } else if (c === '"') q = true;
      else if (c === ',') { row.push(cell); cell = ''; }
      else if (c === '\n' || c === '\r') {
        if (c === '\r' && text[i + 1] === '\n') i++;
        row.push(cell); cell = '';
        if (row.some((v) => v !== '')) rows.push(row);
        row = [];
      } else cell += c;
    }
    row.push(cell);
    if (row.some((v) => v !== '')) rows.push(row);
    return rows;
  }

  function toCsv(rows) {
    return '﻿' + rows.map((r) => r.map((v) => `"${String(v ?? '').replace(/"/g, '""')}"`).join(',')).join('\r\n') + '\r\n';
  }

  // UTF-8で読めなければShift_JISとして読む
  function decode(buf) {
    try { return new TextDecoder('utf-8', { fatal: true }).decode(buf); }
    catch (_) { return new TextDecoder('shift_jis').decode(buf); }
  }

  // 「https://x.com/abc?s=11」「@abc」「abc」→ 'abc'
  function acctKey(s) {
    let t = String(s ?? '').trim();
    const m = t.match(/(?:x|twitter)\.com\/([^/?#\s]+)/i);
    if (m) t = m[1];
    return t.replace(/^@/, '').toLowerCase();
  }

  // 交換相手ごとの理由 → なければ部門・ジュニア共通の理由
  function reasonsFor(reasons, ex) {
    const own = reasons[`${acctKey(ex.acct)}|${ex.no}|${ex.names[0]}`];
    if (ex.acct && own && own.length) return own;
    return reasons[`|${ex.no}|${ex.names[0]}`];
  }

  const toHalf = (s) => String(s ?? '').replace(/[０-９]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xfee0)).trim();
  const leadNum = (s) => { const m = toHalf(s).match(/^(\d+)/); return m ? String(parseInt(m[1], 10)) : ''; };

  // 「21 内村颯太, 126 元木湧」→ ['21','126']。番号がなければ書かれた文字のまま
  function splitNames(s) {
    return toHalf(s).split(/[,，、]/).map((p) => p.trim()).filter(Boolean).map((p) => leadNum(p) || p);
  }

  // ---------- 交換CSV → 交換リスト ----------
  const COL = { status: '進行状況', cat: '項目B', name: '項目C', qty: '数量', acct: '交換先アカウント', note: '備考', used: '使用済み' };

  function readExchanges(rows) {
    if (!rows.length) throw new Error('CSVが空です');
    const head = rows[0].map((h) => String(h).trim());
    // 大賞タスクトラッカーの列名の言い換えにも対応する
    const ALIAS = { '記入項目': COL.cat, '記入者': COL.name };
    const keys = head.map((h) => ALIAS[h] || h);
    const need = [COL.cat, COL.name, COL.qty];
    const miss = need.filter((k) => !keys.includes(k));
    if (miss.length) throw new Error(`CSVに列がありません：${miss.join('、')}`);
    const ix = (k) => keys.indexOf(k);
    const list = rows.slice(1).map((r, i) => {
      const get = (k) => (ix(k) >= 0 ? String(r[ix(k)] ?? '').trim() : '');
      const m = toHalf(get(COL.note)).match(/(\d{1,2})\s*月号のみ/);
      return {
        line: i + 2,
        raw: r,
        status: get(COL.status),
        catText: get(COL.cat),
        no: leadNum(get(COL.cat)),
        nameText: get(COL.name),
        names: splitNames(get(COL.name)),
        qty: parseInt(toHalf(get(COL.qty)).replace(/[^0-9]/g, ''), 10) || 0,
        acct: get(COL.acct),
        note: get(COL.note),
        onlyMonth: m ? parseInt(m[1], 10) : null,
        used: parseInt(toHalf(get(COL.used)), 10) || 0,
      };
    });
    const hasStatus = head.includes(COL.status);
    list.forEach((ex) => { ex.hasStatusCol = hasStatus; });
    return { head, list, hasStatus };
  }

  // 使う状況（交換管理アプリの「進行状況」）
  const TARGET_STATUS = '記入中';

  // ---------- 割り当て ----------
  // cfg：用紙設定、N：今回の枚数、month：今回の号（月）、reasons：{"部門|番号":[理由...]}
  function plan(cfg, exchanges, N, month, reasons) {
    const byNo = {};
    for (const f of cfg.fields) {
      if (f.sheet !== 'vote' || !/^\d+$/.test(String(f.no))) continue;
      const c = (byNo[f.no] = byNo[f.no] || { no: f.no, category: f.category, parts: {}, fixed: false });
      c.parts[f.part] = f;
      if (f.fixed && f.part === '名前') c.fixed = leadNum(f.fixed) || String(f.fixed);
    }
    const issues = [];   // {level:'err'|'warn', ex, text}
    const usable = [];
    let skipped = 0;
    for (const ex of exchanges) {
      const who = ex.acct ? `（${ex.acct}）` : '';
      // 進行状況が「記入中」の行だけ使う（列がないCSVは全行）。それ以外は残りCSVにそのまま残す
      if (ex.hasStatusCol && ex.status !== TARGET_STATUS) { skipped++; continue; }
      if (!ex.no || !byNo[ex.no]) { issues.push({ level: 'err', ex, text: `${ex.line}行目「${ex.catText}」：用紙にない部門です${who}` }); continue; }
      const c = byNo[ex.no];
      if (!ex.names.length) { issues.push({ level: 'err', ex, text: `${ex.line}行目 部門${ex.no}：ジュニアが空欄です${who}` }); continue; }
      if (c.fixed && ex.names[0] !== c.fixed) { issues.push({ level: 'err', ex, text: `${ex.line}行目 部門${ex.no}：この部門は${c.fixed}番で固定です。「${ex.nameText}」には投票できません${who}` }); continue; }
      if (ex.qty <= 0) { issues.push({ level: 'warn', ex, text: `${ex.line}行目 部門${ex.no}：数量が0です` }); continue; }
      const slots = ['名前', '名前2', '名前3'].filter((p) => c.parts[p]).length;
      if (ex.names.length > slots) issues.push({ level: 'warn', ex, text: `${ex.line}行目 部門${ex.no}：名前が${ex.names.length}人分ありますが、欄は${slots}つです。先頭から${slots}人分だけ書きます` });
      if (ex.names.length < slots) issues.push({ level: 'warn', ex, text: `${ex.line}行目 部門${ex.no}：${slots}人分の欄に${ex.names.length}人分しかありません` });
      if (ex.onlyMonth && ex.onlyMonth !== month) {
        // 号の順番は12月号→1月号。過ぎた号の指定が残っていたら知らせる
        const order = (m) => (m === 12 ? 0 : m === 1 ? 1 : -1);
        if (order(ex.onlyMonth) < order(month)) issues.push({ level: 'err', ex, text: `${ex.line}行目 部門${ex.no}：「${ex.onlyMonth}月号のみ」の交換が残っています。${month}月号では使えません${who}` });
        continue; // 別の号の分。今回は使わず残りへ
      }
      usable.push(ex);
    }

    // 部門ごとに、今回の号のみの行 → CSVの順 で詰める
    const segments = {}; // no -> [{ex, start, count, used0}]
    const keyUsed = new Map(); // 交換（交換先・部門・ジュニア）ごとの通しの使用数
    const keyOf_ = new Map();  // 行 → 交換のキー
    const taken = new Map();
    const groups = {};
    for (const ex of usable) (groups[ex.no] = groups[ex.no] || []).push(ex);
    const summary = [];
    for (const no of Object.keys(groups).sort((a, b) => a - b)) {
      // 交換先・部門・ジュニアが同じ行は1つの交換として扱い、ページを連続させて理由も通しでループする
      const keyOf = (ex) => `${acctKey(ex.acct)}|${no}|${ex.names.join(',')}`;
      const first = new Map();
      for (const ex of groups[no].slice().sort((a, b) => a.line - b.line)) if (!first.has(keyOf(ex))) first.set(keyOf(ex), ex.line);
      const pri = (ex) => (ex.onlyMonth === month ? 0 : 1);
      const list = groups[no].slice().sort((a, b) => pri(a) - pri(b) || first.get(keyOf(a)) - first.get(keyOf(b)) || a.line - b.line);
      let p = 0, need = 0;
      segments[no] = [];
      const warned = new Set();
      for (const ex of list) {
        const k = keyOf(ex);
        // 前の号までに使った数（同じ交換の行は同じ値を持つので最大値を使う）
        if (!keyUsed.has(k)) keyUsed.set(k, Math.max(...list.filter((x) => keyOf(x) === k).map((x) => x.used)));
        keyOf_.set(ex, k);
        need += ex.qty;
        const n = Math.min(ex.qty, Math.max(0, N - p));
        if (n > 0) {
          segments[no].push({ ex, start: p, count: n, used0: keyUsed.get(k) });
          keyUsed.set(k, keyUsed.get(k) + n);
          taken.set(ex, n);
          p += n;
        }
        const rl = reasonsFor(reasons, ex);
        if (!(rl && rl.length) && !warned.has(k)) { warned.add(k); issues.push({ level: 'warn', ex, text: `部門${no}・${ex.names[0]}番${ex.acct ? `（${acctKey(ex.acct)}）` : ''}：理由シートに登録がありません（投票内容シートの理由を使います）` }); }
        if (ex.onlyMonth === month && n < ex.qty) issues.push({ level: 'err', ex, text: `${ex.line}行目 部門${no}：「${month}月号のみ」ですが、今回の枚数に入りきりません（${ex.qty - n}票あふれ）` });
      }
      summary.push({ no, category: byNo[no].category, need, used: p, carry: need - p, fallback: Math.max(0, N - p) });
    }

    // 残りCSV（入りきらなかった票・別の号の分・エラーの行）
    // 同じ交換の行は、次の号で理由が続きから回るよう「使用済み」に交換全体で使った数を書く
    const remaining = exchanges.map((ex) => ({ ex, left: ex.qty - (taken.get(ex) || 0), used: keyOf_.has(ex) ? keyUsed.get(keyOf_.get(ex)) : ex.used + (taken.get(ex) || 0) })).filter((r) => r.left > 0);

    return { segments, summary, issues, remaining, byNo, skipped };
  }

  // 1ページ分の部門の値（番号と理由）を返す
  function pageVotes(planRes, reasons, i) {
    const out = {}; // fieldId -> value
    for (const [no, segs] of Object.entries(planRes.segments)) {
      const s = segs.find((x) => i >= x.start && i < x.start + x.count);
      if (!s) continue;
      const c = planRes.byNo[no];
      ['名前', '名前2', '名前3'].forEach((p, k) => {
        const f = c.parts[p];
        if (f && !f.fixed) out[f.id] = s.ex.names[k] || '';
      });
      const list = reasonsFor(reasons, s.ex);
      // 理由が登録されていない交換は、投票内容シートの理由をそのまま使う
      if (c.parts['理由'] && list && list.length) out[c.parts['理由'].id] = list[(s.used0 + (i - s.start)) % list.length];
    }
    return out;
  }

  // 自分の票（投票内容シートの内容になるページ）の理由ループ
  // 理由シートで交換先アカウントを「自分」にした行を、部門・ジュニアが投票内容の名前と一致するときに使う
  const OWN = '自分';
  function ownReasons(cfg, planRes, reasons, values, i) {
    const out = {};
    const used = {};
    for (const [no, segs] of Object.entries(planRes.segments)) used[no] = segs.reduce((a, x) => a + x.count, 0);
    const parts = {};
    for (const f of cfg.fields) {
      if (f.sheet !== 'vote' || !/^\d+$/.test(String(f.no))) continue;
      (parts[f.no] = parts[f.no] || {})[f.part] = f;
    }
    for (const [no, p] of Object.entries(parts)) {
      if (!p['理由'] || !p['名前']) continue;
      if (i < (used[no] || 0)) continue; // 交換の票が入るページ
      const name = String(values[p['名前'].id] ?? '').trim();
      if (!name) continue;
      const list = reasons[`${OWN}|${no}|${leadNum(name) || name}`];
      if (list && list.length) out[p['理由'].id] = list[(i - (used[no] || 0)) % list.length];
    }
    return out;
  }

  // 「自分」の行のうち、投票内容の名前と合わずに使われないものを知らせる
  function unusedOwn(cfg, reasons, values) {
    const msgs = [];
    const nameOf = {};
    for (const f of cfg.fields) if (f.sheet === 'vote' && f.part === '名前' && /^\d+$/.test(String(f.no))) nameOf[f.no] = String(values[f.id] ?? '').trim();
    for (const k of Object.keys(reasons)) {
      const [a, no, jr] = k.split('|');
      if (a !== OWN) continue;
      const n = nameOf[no];
      if (n === undefined) msgs.push(`部門${no}は用紙にありません（理由シートの「自分」の行）`);
      else if ((leadNum(n) || n) !== jr) msgs.push(`部門${no}・${jr}番：投票内容シートの名前（${n || '空欄'}）と一致しないので、この「自分」の理由は使いません`);
    }
    return msgs;
  }

  // 対応表：交換先ごとに、今回何枚目〜何枚目に入ったか
  function mapping(planRes, chunk, fileName) {
    const rows = [['交換先アカウント', '部門', 'ジュニア', '今回の枚数', '何枚目から', '何枚目まで', 'ファイル', '残り']];
    const all = [];
    for (const segs of Object.values(planRes.segments)) for (const s of segs) all.push(s);
    all.sort((a, b) => (a.ex.acct || '').localeCompare(b.ex.acct || '') || a.ex.no - b.ex.no);
    for (const s of all) {
      const from = s.start + 1, to = s.start + s.count;
      const files = [];
      for (let c = Math.floor(s.start / chunk); c <= Math.floor((to - 1) / chunk); c++) files.push(fileName(c));
      rows.push([s.ex.acct, s.ex.catText, s.ex.nameText, s.count, from, to, files.join(' / '), s.ex.qty - s.count]);
    }
    return rows;
  }

  function remainingCsv(head, planRes) {
    const h = head.includes(COL.used) ? head.slice() : [...head, COL.used];
    const qi = h.indexOf(COL.qty), ui = h.indexOf(COL.used);
    const rows = [h];
    for (const r of planRes.remaining) {
      const row = h.map((_, i) => r.ex.raw[i] ?? '');
      row[qi] = String(r.left);
      row[ui] = String(r.used);
      rows.push(row);
    }
    return rows;
  }

  g.JTPlan = { OWN, ownReasons, unusedOwn, TARGET_STATUS, acctKey, parseCsv, toCsv, decode, readExchanges, plan, pageVotes, mapping, remainingCsv, leadNum, splitNames };
})(typeof window !== 'undefined' ? window : globalThis);
