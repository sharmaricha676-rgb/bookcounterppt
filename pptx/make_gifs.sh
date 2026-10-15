#!/usr/bin/env bash
# Turns the rendered frame folders into GIFs for the deck.
# Loops (spin, explode, sensor) repeat forever; page turns play once and stop on the new count.
set -euo pipefail
A="$1"; OUT="$2"; mkdir -p "$OUT"
PAL='split[a][b];[a]palettegen=max_colors=200:stats_mode=full[p];[b][p]paletteuse=dither=sierra2_4a'

loop_gif() { # dir fps out [hold_seconds]
  local hold="${4:-0}"
  ffmpeg -y -loglevel error -framerate "$2" -i "$A/$1/%03d.png" \
    -vf "tpad=stop_mode=clone:stop_duration=${hold},${PAL}" -loop 0 "$OUT/$3"
}
once_gif() { # dir fps out [reverse]
  local pre=""
  [ "${4:-}" = "reverse" ] && pre="reverse,"
  ffmpeg -y -loglevel error -framerate "$2" -i "$A/$1/%03d.png" -vf "${pre}${PAL}" -loop 0 "$OUT/$3"
  python3 "$(dirname "$0")/gif_hold.py" "$OUT/$3"   # stop on the last frame, whatever the viewer does with looping
}

loop_gif spin 12 spin.gif
loop_gif explode 10 explode.gif 0
loop_gif sensor 14 sensor.gif 1.2
for k in 1 2 3 4 5; do once_gif "fwd$k" 14 "fwd$k.gif"; done
# flips back by hand (no wheel): left sensor first, count goes down
for k in 0 1 2 3 4; do once_gif "back$k" 14 "back$k.gif"; done
# the turner view sits left of centre in the render; crop off the empty right side
ffmpeg -y -loglevel error -framerate 14 -i "$A/turner/%03d.png" \
  -vf "crop=760:540:0:0,tpad=stop_mode=clone:stop_duration=1.0,${PAL}" -loop 0 "$OUT/turner.gif"
ls -la "$OUT"
