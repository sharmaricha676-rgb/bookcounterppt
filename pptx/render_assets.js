// Renders every image and animation frame the interactive .pptx needs from the live three.js model.
// The page's animation clock is replaced so each frame is captured at an exact simulated time.
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

const OUT = process.argv[2];
const ONLY = process.argv[3] ? process.argv[3].split(',') : null;
const want = k => !ONLY || ONLY.includes(k);
const URL = 'file:///home/user/bookcounterppt/index.html';
const HIDE_UI = '.top,.nav,.hint,.card,.tip,.loading,.help{display:none!important}';

async function open(browser, w, h, dpr, panelMode) {
  const page = await browser.newPage({ viewport: { width: w, height: h }, deviceScaleFactor: dpr });
  page.on('pageerror', e => console.log('[pageerror]', e.message));
  await page.addInitScript(() => {
    window.__fake = { on: false, t: 0, step: 0, frames: 0 };
    const raf = window.requestAnimationFrame.bind(window);
    window.requestAnimationFrame = cb => raf(ts => {
      const f = window.__fake; f.frames++;
      if (f.on) { f.t += f.step; cb(f.t); } else { f.t = ts; cb(ts); }
    });
  });
  await page.goto(URL);
  await page.waitForFunction(() => window.BookCounter, null, { timeout: 60000 });
  await page.waitForTimeout(2500);
  // 'right' keeps the (invisible) slide panel so the model sits right of centre, like the live deck
  const panelCss = panelMode === 'right' ? '.panel{opacity:0!important}' : '.panel{display:none!important}';
  await page.addStyleTag({ content: HIDE_UI + panelCss });
  await page.waitForTimeout(600);
  await page.evaluate(() => {
    window.__fake.on = true; window.__fake.step = 0;
    window.BookCounter.scene.traverse(o => { if (o.isSpotLight && o.shadow) { o.shadow.mapSize.set(1024, 1024); if (o.shadow.map) { o.shadow.map.dispose(); o.shadow.map = null; } } });
  });
  return page;
}
const frames = async (page, n) => {
  await page.evaluate(() => { window.__fake.mark = window.__fake.frames; });
  await page.waitForFunction(n => window.__fake.frames - window.__fake.mark >= n, n, { timeout: 120000, polling: 30 });
};
// advance simulated time by one 50 ms tick and let it render
const tick = async (page, ms) => {
  await page.evaluate(ms => { window.__fake.step = ms; window.__fake.mark = window.__fake.frames; }, ms);
  await page.waitForFunction(() => window.__fake.frames - window.__fake.mark >= 1, null, { timeout: 120000, polling: 20 });
  await page.evaluate(() => { window.__fake.step = 0; });
};
const labels = (page, on) => page.evaluate(on => { document.querySelector('#labels').style.display = on ? '' : 'none'; }, on);
async function setSlide(page, n, extra) {
  await page.evaluate(([n, extra]) => {
    const B = window.BookCounter;
    B.go(n); B.settle(); B.state.spin = false;
    if (extra && extra.pages !== undefined) B.fw.pages = extra.pages;
    if (extra && extra.orbit) Object.assign(B.orbit, extra.orbit);
  }, [n, extra || null]);
  await frames(page, 3);
}
const mkdir = d => fs.mkdirSync(d, { recursive: true });

(async () => {
  mkdir(OUT);
  const browser = await chromium.launch({ args: ['--use-angle=swiftshader', '--enable-unsafe-swiftshader', '--ignore-gpu-blocklist'] });

  // ---------- stills, model right of centre (text goes on the left of the slide) ----------
  if (want('stills2')) {
    const p2 = await open(browser, 1600, 900, 1, 'right');
    await labels(p2, true);
    await setSlide(p2, 7); await p2.screenshot({ path: `${OUT}/wiring.png` });
    await setSlide(p2, 8); await p2.screenshot({ path: `${OUT}/dims.png` });
    await labels(p2, false);
    await setSlide(p2, 10); await p2.screenshot({ path: `${OUT}/hero2.png` });
    console.log('stills2 done');
    await p2.close();
  }
  if (want('stills')) {
    const page = await open(browser, 1280, 720, 2, 'right');
    await labels(page, false);
    await setSlide(page, 2);                       // the device, no labels: hotspots are drawn in PowerPoint
    await page.screenshot({ path: `${OUT}/hub.png` });
    const spots = await page.evaluate(() => {
      const B = window.BookCounter, THREE = B.THREE;
      const groups = id => B.deviceRoot.children.filter(c => c.userData.part === id);
      const proj = v => { const p = v.clone().project(B.camera); return [+((p.x + 1) / 2).toFixed(4), +((1 - p.y) / 2).toFixed(4)]; };
      const centre = id => { const b = new THREE.Box3(); groups(id).forEach(g => b.expandByObject(g)); return b.getCenter(new THREE.Vector3()); };
      const worldOf = o => o.getWorldPosition(new THREE.Vector3());
      const oled = groups('oled')[0], button = groups('button')[0], base = groups('shell')[0];
      const decal = base.children.find(m => m.material && m.material.alphaTest > 0);
      const screen = oled.children.find(m => m.material && m.material.type === 'MeshBasicMaterial');
      const sens = id => { const b = new THREE.Box3().setFromObject(groups(id)[0]); const c = b.getCenter(new THREE.Vector3()); c.y = b.min.y + (b.max.y - b.min.y) * 0.35; return c; };
      return { oled: proj(worldOf(screen)), button: proj(worldOf(button).add(new THREE.Vector3(0, 0.3, 0))), shell: proj(worldOf(decal)), sensorL: proj(sens('sensorL')), sensorR: proj(sens('sensorR')) };
    });
    fs.writeFileSync(`${OUT}/hub_spots.json`, JSON.stringify(spots, null, 2));
    console.log('hub + spots', JSON.stringify(spots));
    await page.close();

    const p2 = await open(browser, 1600, 900, 1, 'right');
    await labels(p2, false);
    await setSlide(p2, 1); await p2.screenshot({ path: `${OUT}/problem.png` });
    await setSlide(p2, 6, { pages: 27 }); await p2.screenshot({ path: `${OUT}/oled.png` });
    await labels(p2, true);
    await setSlide(p2, 7); await p2.screenshot({ path: `${OUT}/wiring.png` });
    await setSlide(p2, 8); await p2.screenshot({ path: `${OUT}/dims.png` });
    await labels(p2, false);
    await setSlide(p2, 10); await p2.screenshot({ path: `${OUT}/hero2.png` });
    await setSlide(p2, 0); await p2.screenshot({ path: `${OUT}/hero.png` });
    console.log('stills done');
    await p2.close();
  }

  // ---------- looping GIF frames, model centred ----------
  if (want('spin')) {
    const page = await open(browser, 960, 540, 1, 'center');
    await labels(page, false);
    await setSlide(page, 0);
    const base = await page.evaluate(() => ({ theta: window.BookCounter.orbit.theta, r: window.BookCounter.orbit.r }));
    const dir = `${OUT}/spin`; mkdir(dir);
    for (let i = 0; i < 36; i++) {
      await page.evaluate(([t, r]) => { const o = window.BookCounter.orbit; o.theta = t; o.r = r; o.vt = 0; }, [base.theta + i * Math.PI * 2 / 36, 24]);
      await frames(page, 1);
      await page.screenshot({ path: `${dir}/${String(i).padStart(3, '0')}.png` });
    }
    console.log('spin done');
    await page.close();
  }

  if (want('explode')) {
    const page = await open(browser, 960, 540, 1, 'center');
    await labels(page, false);
    await setSlide(page, 3);
    await page.evaluate(() => window.BookCounter.setExplode(0));
    const dir = `${OUT}/explode`; mkdir(dir);
    const seq = [];
    for (let i = 1; i <= 14; i++) seq.push(i / 14);
    seq.push(1);
    for (let i = 13; i >= 0; i--) seq.push(i / 14);
    let n = 0;
    for (const v of seq) {
      await page.evaluate(v => window.BookCounter.setExplode(v), v);
      await frames(page, 1);
      await page.screenshot({ path: `${dir}/${String(n++).padStart(3, '0')}.png` });
    }
    console.log('explode done', n);
    await page.close();
  }

  // ---------- page turns (simulated clock, exact frames) ----------
  async function flipFrames(page, dir, dirn, count) {
    mkdir(dir);
    await page.screenshot({ path: `${dir}/000.png` });
    await page.evaluate(d => window.BookCounter.flip(d), dirn);
    for (let i = 1; i < count; i++) {
      await tick(page, 70);
      await page.screenshot({ path: `${dir}/${String(i).padStart(3, '0')}.png` });
    }
  }
  if (want('sensor')) {
    const page = await open(browser, 960, 540, 1, 'center');
    await labels(page, false);
    await setSlide(page, 4);
    await flipFrames(page, `${OUT}/sensor`, 1, 24);
    console.log('sensor done');
    await page.close();
  }
  if (want('flips')) {
    const page = await open(browser, 960, 540, 1, 'center');
    await labels(page, false);
    await setSlide(page, 5, { pages: 0 });
    for (let k = 1; k <= 5; k++) { await flipFrames(page, `${OUT}/fwd${k}`, 1, 26); console.log('fwd', k, await page.evaluate(() => window.BookCounter.fw.pages)); }
    await page.close();
  }

  await browser.close();
})().catch(e => { console.error(e); process.exit(1); });
