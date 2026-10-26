// 共通描画ロジック：プレビューとPDF出力の両方でこの関数だけを使う
(function (g) {
  'use strict';

  const FALLBACK_FONT = 'BIZ UDGothic';
  const MM_PER_PT = 25.4 / 72;

  // Excelの値 → 欄IDごとの描画値。固定部門の名前はExcelを読まず、設定の値だけを使う
  // notices：取り込み時に利用者へ伝える注意（固定の書き換え、使われない列への入力など）
  // loop：{ loops: {部門番号: [理由...]}, index: 何枚目(0始まり) }。ループがある部門は枚ごとに理由を順番に回す
  function resolveValues(cfg, voteRows, voterRow, loop) {
    const values = {};
    const warnings = [];
    const notices = [];
    const voteMap = {};
    for (const r of voteRows || []) {
      const no = normNo(r['部門番号']);
      if (no) voteMap[no] = r;
    }
    // 部門ごとに存在する欄
    const partsOf = {};
    for (const f of cfg.fields) {
      if (f.sheet !== 'vote') continue;
      (partsOf[normNo(f.no)] = partsOf[normNo(f.no)] || new Set()).add(f.part);
    }
    // 使われない名前列（名前2・名前3）に入っている文字。理由が空ならそれを理由として使う
    const stray = {};
    for (const [no, r] of Object.entries(voteMap)) {
      const parts = partsOf[no];
      if (!parts) {
        if (['名前', '名前2', '名前3', '理由'].some((k) => str(r[k]))) notices.push(`部門番号「${no}」は今年の用紙にありません（${r._row}行目）。この行は使いません`);
        continue;
      }
      const extra = ['名前2', '名前3'].filter((k) => !parts.has(k) && str(r[k])).map((k) => str(r[k]));
      if (!extra.length) continue;
      if (parts.has('理由') && !str(r['理由'])) {
        stray[no] = extra.join(' ');
        notices.push(`部門${no}：「名前2/名前3」の列に入っていた「${short(stray[no])}」を理由として描画します（理由は「理由」列へ）`);
      } else {
        notices.push(`部門${no}：「名前2/名前3」の列はこの部門では使いません（「${short(extra.join(' '))}」は描画しません）`);
      }
    }
    for (const f of cfg.fields) {
      if (f.sheet === 'vote') {
        const no = normNo(f.no);
        const row = voteMap[no] || {};
        if (f.fixed != null && f.fixed !== '') {
          values[f.id] = f.fixed;
          const xl = str(row[f.part]);
          if (xl && xl !== f.fixed) {
            warnings.push(`部門${f.no}「${f.category}」の${f.part}は固定です。Excelの「${xl}」は使わず「${f.fixed}」を描画します`);
          }
        } else {
          const list = f.part === '理由' && loop && loop.loops && loop.loops[no];
          if (list && list.length) values[f.id] = list[(loop.index || 0) % list.length];
          else values[f.id] = str(row[f.part]) || (f.part === '理由' ? stray[no] || '' : '');
        }
      } else {
        values[f.id] = str((voterRow || {})[f.key]);
      }
    }
    return { values, warnings, notices };
  }

  // 「1」「1.0」「１」「①」のような表記ゆれを揃える
  function normNo(v) {
    let s = str(v).replace(/[０-９]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xfee0));
    s = s.replace(/[\u2460-\u2473]/g, (c) => String(c.charCodeAt(0) - 0x245f));
    if (/^\d+(\.0+)?$/.test(s)) s = String(parseInt(s, 10));
    return s;
  }

  function str(v) {
    if (v == null) return '';
    return String(v).trim();
  }

  // 1ページ描画。W,H はキャンバスのピクセル寸法
  function renderPage(ctx, W, H, cfg, img, values, opts) {
    opts = opts || {};
    const warnings = [];
    ctx.save();
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, W, H);
    if (img) ctx.drawImage(paperOf(img, opts.paper), 0, 0, W, H);

    const pxPerPt = W / (cfg.page.widthMm / MM_PER_PT);
    // 白黒出力では文字も黒に揃える
    const color = opts.paper && opts.paper !== 'color' ? '#111111' : (opts.color || '#1d1d1f');
    const family = opts.font || 'Klee One';
    const fontOf = (pt) => `${(pt * pxPerPt).toFixed(2)}px "${family}", "${FALLBACK_FONT}", sans-serif`;
    const d = cfg.defaults || {};
    let seed = (opts.seed || 1) >>> 0;
    const rand = () => { seed = (seed * 1664525 + 1013904223) >>> 0; return seed / 4294967296; };

    ctx.fillStyle = color;
    ctx.strokeStyle = color;
    ctx.textBaseline = 'middle';

    for (const f of cfg.fields) {
      const v = values[f.id];
      if (v == null || v === '') continue;
      const label = fieldLabel(f);

      if (f.type === 'choice') {
        const n = parseInt(v, 10);
        const o = (f.options || [])[n - 1];
        if (!o) { warnings.push(`${label}：「${v}」に対応する選択肢がありません（1〜${(f.options || []).length}）`); continue; }
        const rx = (f.rx || 0.012) * W, ry = (f.ry || 0.009) * H;
        ctx.lineWidth = Math.max(1, (cfg.defaults && cfg.defaults.circleLineMm || 0.35) / cfg.page.widthMm * W);
        const start = rand() * Math.PI * 2;
        const tilt = (rand() - 0.5) * 0.25;
        ctx.beginPath();
        ctx.ellipse(o.x * W, o.y * H, rx, ry, tilt, start, start + Math.PI * 2.12);
        ctx.stroke();
        continue;
      }

      const [rx0, ry0, rw0, rh0] = f.rect;
      const x = rx0 * W, y = ry0 * H, w = rw0 * W, h = rh0 * H;

      if (f.type === 'digits' && f.count) {
        let s = String(v).replace(/[^0-9０-９]/g, '');
        if (s.length < f.count && /^\d+$/.test(s)) s = s.padStart(f.count, '0'); // Excelで消えた先頭0を補う
        if (s.length > f.count) warnings.push(`${label}：${f.count}桁を超えています（${s}）`);
        const cw = w / f.count;
        const size = f.size || d.fontSize || 10;
        ctx.font = fontOf(size);
        ctx.textAlign = 'center';
        for (let i = 0; i < Math.min(s.length, f.count); i++) {
          ctx.fillText(s[i], x + cw * (i + 0.5), y + h / 2);
        }
        continue;
      }

      // 郵便番号：数字7桁なら 123-4567 に整える（Excelで消えた先頭0も補う）
      let text = String(v);
      if (f.format === 'zip') text = formatZip(text);

      // text：枠に収まるまで縮小し、必要なら折り返す
      const fit = fitText(ctx, text, w, h, f.size || d.fontSize || 10, f.min || d.minFontSize || 5, f.lines || 1, fontOf, pxPerPt);
      if (!fit.fit) warnings.push(`${label}：最小サイズでも枠に収まりません（「${short(v)}」）`);
      ctx.font = fontOf(fit.size);
      const lineH = fit.size * pxPerPt * 1.15;
      const total = lineH * fit.lines.length;
      // auto：基準サイズの1行に収まれば中央、収まらなければ左寄せ（番号は中央、長い名前は左から）
      let align = f.align || 'left';
      if (align === 'auto') align = fit.size === (f.size || d.fontSize || 10) && fit.lines.length === 1 ? 'center' : 'left';
      ctx.textAlign = align;
      const tx = align === 'center' ? x + w / 2 : align === 'right' ? x + w - 2 : x + 2;
      fit.lines.forEach((ln, i) => {
        ctx.fillText(ln, tx, y + (h - total) / 2 + lineH * (i + 0.5));
      });
    }
    ctx.restore();
    return warnings;
  }

  function formatZip(v) {
    const digits = v.replace(/[０-９]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xfee0)).replace(/[^0-9]/g, '');
    if (/^\d{4,7}$/.test(digits) && !/[^0-9０-９\-－ー‐\s〒]/.test(v)) {
      const z = digits.padStart(7, '0');
      return z.slice(0, 3) + '-' + z.slice(3);
    }
    return v;
  }

  // 用紙画像の色処理。color=そのまま、gray=白黒、light=白黒で薄く（インク節約）。モードごとに1回だけ作って使い回す
  const paperCache = new WeakMap();
  function paperOf(img, mode) {
    if (!mode || mode === 'color') return img;
    let m = paperCache.get(img);
    if (!m) { m = {}; paperCache.set(img, m); }
    if (m[mode]) return m[mode];
    const c = document.createElement('canvas');
    c.width = img.naturalWidth || img.width; c.height = img.naturalHeight || img.height;
    const x = c.getContext('2d');
    x.drawImage(img, 0, 0);
    const data = x.getImageData(0, 0, c.width, c.height);
    const p = data.data;
    for (let i = 0; i < p.length; i += 4) {
      let g = 0.299 * p[i] + 0.587 * p[i + 1] + 0.114 * p[i + 2];
      if (mode === 'light') g = 255 - (255 - g) * 0.4;
      p[i] = p[i + 1] = p[i + 2] = g;
    }
    x.putImageData(data, 0, 0);
    m[mode] = c;
    return c;
  }

  function fitText(ctx, text, w, h, size, min, maxLines, fontOf, pxPerPt) {
    let last = null;
    for (let s = size; s >= min - 1e-6; s -= 0.5) {
      ctx.font = fontOf(s);
      const lines = wrap(ctx, text, w - 4);
      const ok = lines.length <= maxLines && lines.length * s * pxPerPt * 1.15 <= h + 0.5;
      last = { size: s, lines: lines.slice(0, maxLines), fit: ok && lines.length <= maxLines };
      if (ok) return last;
    }
    return last;
  }

  // 日本語は1文字単位で折り返す
  function wrap(ctx, text, maxW) {
    const out = [];
    for (const para of text.split(/\r?\n/)) {
      let line = '';
      for (const ch of Array.from(para)) {
        if (line && ctx.measureText(line + ch).width > maxW) { out.push(line); line = ch; }
        else line += ch;
      }
      out.push(line);
    }
    return out;
  }

  function short(s) { s = String(s); return s.length > 12 ? s.slice(0, 12) + '…' : s; }

  function fieldLabel(f) {
    if (f.sheet === 'vote') return `部門${f.no}「${f.category}」${f.part}`;
    return f.key;
  }

  // 投票内容シートに出す部門の一覧（設定の並び順）
  function categories(cfg) {
    const map = new Map();
    for (const f of cfg.fields) {
      if (f.sheet !== 'vote') continue;
      const k = str(f.no);
      if (!map.has(k)) map.set(k, { no: k, category: f.category, parts: {}, fixed: {} });
      const c = map.get(k);
      c.parts[f.part] = true;
      if (f.fixed) c.fixed[f.part] = f.fixed;
    }
    return [...map.values()];
  }

  function voterFields(cfg) {
    const seen = new Set(), out = [];
    for (const f of cfg.fields) {
      if (f.sheet === 'voter' && !seen.has(f.key)) { seen.add(f.key); out.push(f); }
    }
    return out;
  }

  g.JTRender = { resolveValues, normNo, renderPage, categories, voterFields, fieldLabel, FALLBACK_FONT, MM_PER_PT };
})(window);
