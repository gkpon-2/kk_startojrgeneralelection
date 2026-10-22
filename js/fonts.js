// 同梱フォント。選んだものだけ読み込む（日本語フォントは1種4〜9MB）
(function (g) {
  'use strict';
  const FONTS = [
    { family: 'Klee One', file: 'fonts/KleeOne-Regular.ttf', label: 'Klee One（丁寧な鉛筆書き）' },
    { family: 'Zen Kurenaido', file: 'fonts/ZenKurenaido-Regular.ttf', label: 'Zen Kurenaido（ペン字風）' },
    { family: 'Yomogi', file: 'fonts/Yomogi-Regular.ttf', label: 'Yomogi（ゆるい手書き）' },
    { family: 'Hachi Maru Pop', file: 'fonts/HachiMaruPop-Regular.ttf', label: 'Hachi Maru Pop（丸文字）' },
    { family: 'Yusei Magic', file: 'fonts/YuseiMagic-Regular.ttf', label: 'Yusei Magic（マーカー風）' },
  ];
  const FALLBACK = { family: 'BIZ UDGothic', file: 'fonts/BIZUDGothic-Regular.ttf' };
  const loaded = new Map();

  function load(f) {
    if (!loaded.has(f.family)) {
      const ff = new FontFace(f.family, `url(${f.file})`);
      loaded.set(f.family, ff.load().then((x) => { document.fonts.add(x); return x; }));
    }
    return loaded.get(f.family);
  }

  // 手書き風フォントにない文字はブラウザが1文字ずつ代替フォントで描くので、代替も必ず読む
  async function ensure(family) {
    const f = FONTS.find((x) => x.family === family) || FONTS[0];
    await Promise.all([load(f), load(FALLBACK)]);
    return f.family;
  }

  g.JTFonts = { FONTS, FALLBACK, ensure };
})(window);
