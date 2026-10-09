# Book Counter: 3D interactive presentation

An interactive, 3D slide deck for **Book Counter**, a 3D-printed bookmark that counts the pages you turn.
It runs in any modern browser. The whole deck is one file, `index.html`.

The model is built in code to match the real device: the white printed shell with **BOOK COUNTER**
embossed on the front, the OLED window and reset button on top, and the two blue IR sensor modules
(FC-51 / LM393) on the front corners, wired in with jumper leads.

## Present it

1. Open `index.html` in Chrome, Edge, Firefox or Safari (double-clicking it works).
2. Press **F** for full screen.
3. Use **→ / ←**, **Space** or a presentation clicker to move between the 12 slides.

It works offline. The 3D engine loads from a CDN, and falls back to the copy in `vendor/` when there is no internet.
Without internet the custom fonts fall back to system fonts.

| Key | Does |
| --- | --- |
| → · Space · PageDown | Next slide |
| ← · PageUp | Previous slide |
| Home · End | First or last slide |
| T · B | Turn a page forward or back (on slides with the book) |
| E | Open the shell (exploded view) |
| X | X-ray: see through the shell |
| S | Turntable on or off |
| R | Reset the camera for this slide |
| F | Full screen |
| ? | Show the shortcuts |

Mouse: drag to rotate, scroll to zoom, click any part for its details. Touch: drag, pinch, tap.

### What's on the slides

1. Title
2. The problem: paper books keep no score
3. The device: OLED, two IR sensors, reset button, printed shell (labelled on the model)
4. Inside: exploded view with the microcontroller, battery, OLED and wiring
5. How an IR sensor sees a page (beam shown in red)
6. **Live demo:** turn pages and watch the OLED count, with a live trace of both sensor signals
7. The counting logic (simplified Arduino code; the demo runs this same logic)
8. Wiring table
9. Design and build, with dimension lines on the model
10. Design challenges
11. Next version
12. Questions, plus how many pages were turned during the talk

You can link straight to a slide: `index.html#6` opens the live demo.

## Make it match your build

All the facts live in one `CONFIG` block near the top of the `<script>` in `index.html`:

```js
const CONFIG = {
  presenters: '',                      // your names, shown on the first and last slide
  sizeMM: { w: 105, d: 60, h: 35 },    // outer size of the shell; the 3D model is built from these
  material: 'PLA',
  mcu: 'Arduino Nano',
  battery: '9 V battery',
  ...
  wiring: [ ['IR sensor, left', 'OUT', 'D2'], ... ],
};
```

Change a value, save, and reload the page. The slides, part cards and the model all update.
Slide text is in the `SLIDES` list just below it.

## Use the model in PowerPoint

`assets/book-counter.glb` is the same 3D model as a standard glTF file.

1. In PowerPoint: **Insert → 3D Models → This Device…** and pick `assets/book-counter.glb`.
2. Drag the model to rotate it.
3. For a moving 3D effect, duplicate the slide, rotate or zoom the model on the copy, and apply the **Morph** transition.

If you change `CONFIG`, regenerate the file:

```sh
npm install --no-save playwright && npx playwright install chromium
node tools/export-glb.cjs
```

## Put it online (optional)

Push this repository to GitHub, then **Settings → Pages → Deploy from a branch**, and choose the branch and `/ (root)`.
The deck is then available at `https://<user>.github.io/<repo>/`.

## Files

| Path | What it is |
| --- | --- |
| `index.html` | The presentation: slides, 3D model, demo logic |
| `vendor/three.min.js` | three.js r128, the offline fallback for the 3D engine |
| `vendor/GLTFExporter.js` | three.js exporter, used only by `tools/export-glb.cjs` |
| `assets/book-counter.glb` | The device model for PowerPoint or any 3D viewer |
| `tools/export-glb.cjs` | Rebuilds the `.glb` from the live model |

three.js is MIT licensed; see `vendor/LICENSE-three.txt`.
