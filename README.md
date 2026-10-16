# Book Counter: 3D interactive presentation

An interactive, 3D slide deck for **Book Counter**, a 3D-printed bookmark that counts the pages you turn.
It runs in any modern browser. The whole deck is one file, `index.html`.

The model is built in code to match the real device: the white printed shell with **BOOK COUNTER**
embossed on the front, the OLED window and turn button on top, and the two blue IR sensor modules
(FC-51 / LM393) on the front corners, wired in with jumper leads.

It also has an **automatic page turner**: press the button and the Arduino tells the motor driver to swing
the servo, whose rubber wheel pushes the page over; the IR sensors then count it like any other turn.
The servo hangs off a hollow 3D-printed beam (180 × 16 × 13 mm) that carries its cable from the shell out over the book,
and the whole bookmark sits on a cardboard base. There is no battery: the device runs from the computer's USB port.

Both decks share a Y2K, early-2000s look: a glossy sky-blue backdrop, glass panels, Aqua-style buttons and
Windows XP-style window frames.

## Present it

1. Open `index.html` in Chrome, Edge, Firefox or Safari (double-clicking it works).
2. Press **F** for full screen.
3. Use **→ / ←**, **Space** or a presentation clicker to move between the 13 slides.

It works offline. The 3D engine loads from a CDN, and falls back to the copy in `vendor/` when there is no internet.
Without internet the custom fonts fall back to system fonts.

| Key | Does |
| --- | --- |
| → · Space · PageDown | Next slide |
| ← · PageUp | Previous slide |
| Home · End | First or last slide |
| A | Press the turn button: the page turner flips a page |
| T · B | Turn a page forward or back by hand (on slides with the book) |
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
3. The device: OLED, two IR sensors, turn button, hollow beam, page turner, printed shell, cardboard base (labelled on the model)
4. Inside: exploded view with the Arduino, motor driver, OLED, USB power and wiring
5. How an IR sensor sees a page (beam shown in red)
6. **The page turner:** button → Arduino → motor driver → servo → wheel → page, step by step
7. **Live demo:** press the turn button or flip by hand, and watch the OLED count, with a live trace of both sensor signals
8. The counting logic (simplified Arduino code; the demo runs this same logic)
9. Wiring table
10. Design and build, with dimension lines on the model
11. Design challenges
12. Next version
13. Questions, plus how many pages were turned during the talk

You can link straight to a slide: `index.html#7` opens the live demo.

## Make it match your build

All the facts live in one `CONFIG` block near the top of the `<script>` in `index.html`:

```js
const CONFIG = {
  presenters: '',                      // your names, shown on the first and last slide
  sizeMM: { w: 105, d: 60, h: 35 },    // outer size of the shell; the 3D model is built from these
  material: 'PLA',
  mcu: 'Arduino Nano',
  power: 'USB cable from the computer · 5 V',
  servo: 'Servo motor (SG90 size)',
  driver: 'Motor driver board',
  pipeMM: { l: 180, w: 16, h: 13 },    // the hollow beam that holds the page turner
  ...
  wiring: [ ['IR sensor, left', 'OUT', 'D2'], ... ],
};
```

Change a value, save, and reload the page. The slides, part cards and the model all update.
Slide text is in the `SLIDES` list just below it.

## PowerPoint version

`BookCounter-interactive.pptx` is the same talk as a native PowerPoint file. It is clickable in slideshow mode:
a menu slide links to every section, the live demo has **Press turn button ►** and **◄ Flip back by hand**
buttons that step the OLED count, the hotspots on the device open part details, and the 3D model can be
rotated on the slide. It uses Morph and Page Curl transitions and animated GIFs rendered from `index.html`.

The generator is in `pptx/` (`render_assets.js` renders the frames, `make_gifs.sh` builds the GIFs,
`build_deck.py` writes the deck, `check_links.py` checks every button lands on a slide).

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
