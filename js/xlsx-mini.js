// 最小限の .xlsx 読み書き（JSZipのみ使用）
(function (g) {
  'use strict';

  const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

  function colName(i) {
    let s = '';
    i += 1;
    while (i > 0) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); }
    return s;
  }
  function colIndex(ref) {
    const letters = ref.replace(/[0-9]/g, '');
    let n = 0;
    for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n - 1;
  }

  // sheets: [{name, rows: [[cell...]], widths: [num], headerStyle: true}]
  async function write(sheets) {
    const zip = new JSZip();
    zip.file('[Content_Types].xml',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
      sheets.map((s, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('') +
      '</Types>');
    zip.file('_rels/.rels',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
      '</Relationships>');
    zip.file('xl/workbook.xml',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' +
      sheets.map((s, i) => `<sheet name="${esc(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('') +
      '</sheets></workbook>');
    zip.file('xl/_rels/workbook.xml.rels',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      sheets.map((s, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('') +
      `<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
      '</Relationships>');
    // スタイル：0=標準、1=見出し（太字・灰色）、2=文字列書式（郵便番号の先頭0対策）、3=固定（グレー）
    zip.file('xl/styles.xml',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<fonts count="3"><font><sz val="11"/><name val="游ゴシック"/></font><font><b/><sz val="11"/><name val="游ゴシック"/></font><font><sz val="11"/><color rgb="FF888888"/><name val="游ゴシック"/></font></fonts>' +
      '<fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
      '<fill><patternFill patternType="solid"><fgColor rgb="FFE8E8EC"/><bgColor indexed="64"/></patternFill></fill>' +
      '<fill><patternFill patternType="solid"><fgColor rgb="FFF3F3F3"/><bgColor indexed="64"/></patternFill></fill></fills>' +
      '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      '<cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
      '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>' +
      '<xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
      '<xf numFmtId="0" fontId="2" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs>' +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
      '</styleSheet>');
    sheets.forEach((s, i) => {
      const cols = (s.widths || []).map((w, c) => `<col min="${c + 1}" max="${c + 1}" width="${w}" customWidth="1"${s.textCols && s.textCols.includes(c) ? ' style="2"' : ''}/>`).join('');
      const rows = s.rows.map((row, r) =>
        `<row r="${r + 1}">` + row.map((cell, c) => {
          if (cell == null || cell === '') return '';
          const ref = colName(c) + (r + 1);
          let style = r === 0 && s.headerStyle ? 1 : 0;
          if (s.greyCells && s.greyCells.has(`${r},${c}`)) style = 3;
          if (typeof cell === 'number') return `<c r="${ref}" s="${style}"><v>${cell}</v></c>`;
          return `<c r="${ref}" t="inlineStr" s="${style}"><is><t xml:space="preserve">${esc(cell)}</t></is></c>`;
        }).join('') + '</row>').join('');
      zip.file(`xl/worksheets/sheet${i + 1}.xml`,
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
        (s.headerStyle ? '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>' : '') +
        (cols ? `<cols>${cols}</cols>` : '') +
        `<sheetData>${rows}</sheetData></worksheet>`);
    });
    return zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }

  // 戻り値: { シート名: [[セル値...], ...] }
  async function read(arrayBuffer) {
    const zip = await JSZip.loadAsync(arrayBuffer);
    const parse = async (path) => {
      const f = zip.file(path);
      if (!f) return null;
      return new DOMParser().parseFromString(await f.async('string'), 'application/xml');
    };
    const tags = (node, name) => Array.from(node.getElementsByTagNameNS('*', name));
    // Excelが保存するふりがな（<rPh>の中の<t>）は本文ではないので読まない
    const inPhonetic = (el, root) => { for (let p = el.parentNode; p && p !== root; p = p.parentNode) if (p.localName === 'rPh') return true; return false; };
    const textOf = (node) => tags(node, 't').filter((t) => !inPhonetic(t, node)).map((t) => t.textContent).join('');

    const wb = await parse('xl/workbook.xml');
    const rels = await parse('xl/_rels/workbook.xml.rels');
    if (!wb || !rels) throw new Error('Excelファイル（.xlsx）として読めませんでした');
    const relMap = {};
    for (const r of tags(rels, 'Relationship')) relMap[r.getAttribute('Id')] = r.getAttribute('Target');

    const sst = await parse('xl/sharedStrings.xml');
    const shared = sst ? tags(sst, 'si').map(textOf) : [];

    const out = {};
    for (const s of tags(wb, 'sheet')) {
      const rid = s.getAttribute('r:id') || s.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
      let target = relMap[rid] || '';
      target = target.replace(/^\//, '');
      if (!target.startsWith('xl/')) target = 'xl/' + target;
      const doc = await parse(target);
      if (!doc) continue;
      const rows = [];
      for (const row of tags(doc, 'row')) {
        const ri = parseInt(row.getAttribute('r'), 10) - 1;
        const arr = [];
        for (const c of tags(row, 'c')) {
          const ci = colIndex(c.getAttribute('r'));
          const t = c.getAttribute('t');
          const vNode = tags(c, 'v')[0];
          let val;
          if (t === 's') val = shared[parseInt(vNode && vNode.textContent, 10)] ?? '';
          else if (t === 'inlineStr') val = textOf(c);
          else if (t === 'b') val = vNode && vNode.textContent === '1' ? 'TRUE' : 'FALSE';
          else if (t === 'str' || t === 'e') val = vNode ? vNode.textContent : '';
          else val = vNode ? Number(vNode.textContent) : '';
          arr[ci] = val;
        }
        rows[ri] = arr;
      }
      out[s.getAttribute('name')] = Array.from(rows, (r) => Array.from(r || [], (v) => (v === undefined ? '' : v)));
    }
    return out;
  }

  // 1行目を見出しとしてオブジェクト配列に。全セル空の行は除外し、元の行番号を _row に残す
  function toObjects(rows) {
    if (!rows || !rows.length) return [];
    const head = rows[0].map((h) => String(h ?? '').trim());
    const out = [];
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i] || [];
      if (!r.some((v) => String(v ?? '').trim() !== '')) continue;
      const o = { _row: i + 1 };
      head.forEach((h, c) => { if (h) o[h] = r[c] ?? ''; });
      out.push(o);
    }
    return out;
  }

  g.XlsxMini = { write, read, toObjects };
})(window);
