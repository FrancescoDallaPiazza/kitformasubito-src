const { chromium } = require('playwright');
(async () => {
  const b = await chromium.launch();
  const p = await b.newPage({ viewport: { width: 980, height: 860 }, deviceScaleFactor: 2 });
  await p.goto('file://' + __dirname + '/tu-formi.html');
  await p.evaluate(() => document.fonts.ready);
  await p.locator('.g').screenshot({ path: __dirname + '/../tu-formi.png', omitBackground: true });
  await b.close();
})();
