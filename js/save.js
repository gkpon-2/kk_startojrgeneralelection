// ファイル保存：claude.ai上では保存確認を経由し、通常のサーバーではそのままダウンロード
(function (g) {
  'use strict';
  let dl = null;
  const ready = (g.claude && typeof g.claude.use === 'function')
    ? g.claude.use('downloads').then((d) => { dl = d; return d; }).catch(() => null)
    : Promise.resolve(null);

  async function save(blob, name) {
    await ready;
    if (dl) {
      try { await dl.save({ filename: name, data: blob }); return 'saved'; }
      catch (e) { if (e && e.code === 'declined') return 'declined'; throw new Error((e && e.message) || '保存できませんでした'); }
    }
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
    return 'saved';
  }

  g.JTSave = save;
})(window);
