// 号ごとの用紙（画像と座標）。部門・欄のIDはどの号も同じなので、Excelテンプレートは共通
(function (g) {
  'use strict';
  const LIST = [
    { month: 12, label: '12月号', config: 'template/config.json' },
    { month: 1, label: '1月号', config: 'template/config-01.json' },
  ];

  function loadImage(src) {
    return new Promise((res, rej) => {
      const i = new Image();
      i.onload = () => res(i);
      i.onerror = () => rej(new Error('用紙画像を読み込めません：' + src));
      i.src = src;
    });
  }

  // { 12: {cfg, img, label}, 1: {...} }
  async function loadAll() {
    const out = {};
    await Promise.all(LIST.map(async (it) => {
      const cfg = await (await fetch(it.config, { cache: 'no-store' })).json();
      out[it.month] = { cfg, img: await loadImage(cfg.image), label: it.label };
    }));
    return out;
  }

  g.JTIssues = { LIST, loadAll };
})(window);
