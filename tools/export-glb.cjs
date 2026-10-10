// Regenerates assets/book-counter.glb from the live 3D model in index.html,
// so the same model can be inserted into PowerPoint (Insert > 3D Models > This Device).
//
// Usage (from the repo root):
//   npm install --no-save playwright && npx playwright install chromium
//   node tools/export-glb.cjs
//
// Edit CONFIG in index.html first if you change the shell size; the export follows it.

const path = require('path');
const fs = require('fs');
const { chromium } = require('playwright');

const root = path.resolve(__dirname, '..');
const out = path.join(root, 'assets', 'book-counter.glb');

(async () => {
  const browser = await chromium.launch({ args: ['--use-angle=swiftshader', '--enable-unsafe-swiftshader', '--ignore-gpu-blocklist'] });
  const page = await browser.newPage({ viewport: { width: 1280, height: 800 } });
  page.on('pageerror', e => console.error('page error:', e.message));
  await page.goto('file://' + path.join(root, 'index.html'));
  await page.waitForFunction(() => window.BookCounter, null, { timeout: 60000 });
  await page.waitForTimeout(2500); // let fonts load and the canvases redraw
  await page.addScriptTag({ path: path.join(root, 'vendor', 'GLTFExporter.js') });

  const b64 = await page.evaluate(() => new Promise(resolve => {
    const B = window.BookCounter, THREE = B.THREE;
    B.setExplode(0);
    B.deviceRoot.updateMatrixWorld(true);
    const model = B.deviceRoot.clone(true);
    model.position.set(0, 0, 0);
    model.scale.setScalar(0.01); // the scene is in centimetres, glTF is in metres
    model.traverse(o => {
      o.userData = {};
      if (o.isLineSegments) o.visible = false;                                         // dimension lines
      if (o.isMesh && o.material && o.material.blending === THREE.AdditiveBlending) o.visible = false; // IR beams
    });
    model.updateMatrixWorld(true);
    // centre the visible model on the origin so PowerPoint (and any viewer) rotates it about its middle;
    // hidden helpers (beams, dimension lines) are left out of the measurement because they are not exported
    const box = new THREE.Box3();
    model.traverse(o => { if (o.isMesh && o.visible) box.expandByObject(o); });
    model.position.sub(box.getCenter(new THREE.Vector3()));
    const scene = new THREE.Scene();
    scene.add(model);
    scene.updateMatrixWorld(true);
    new THREE.GLTFExporter().parse(scene, buf => {
      const bytes = new Uint8Array(buf);
      let s = '';
      for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
      resolve(btoa(s));
    }, { binary: true, onlyVisible: true, maxTextureSize: 1024 });
  }));

  fs.mkdirSync(path.dirname(out), { recursive: true });
  fs.writeFileSync(out, Buffer.from(b64, 'base64'));
  console.log(`wrote ${path.relative(root, out)} (${(fs.statSync(out).size / 1024).toFixed(0)} KB)`);
  await browser.close();
})().catch(e => { console.error(e); process.exit(1); });
