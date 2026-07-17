#!/usr/bin/env bash
# Record a demo GIF of Markdown Hub for the README.
#
# Approach: capture a screen region as a PNG sequence, then assemble
# the PNGs into an animated GIF with ffmpeg. Run this on the same OS
# the README screenshots are taken on (Windows recommended — see notes).
#
# Prerequisites:
#   - ffmpeg in PATH (https://ffmpeg.org/download.html)
#   - gifsicle (optional, for size optimization; https://www.lcdf.org/gifsicle/)
#   - For Windows: ScreenToGif is a much easier alternative
#       (https://www.screentogif.com/) — skip this script entirely.
#
# Usage:
#   1. Open VS Code with a sample .md file in the editor.
#   2. Set up the screen region you want to record (suggest 900x600
#      centered on the VS Code window, showing the Explorer + Editor + right-click menu).
#   3. Run: scripts/record-gif.sh <output_name> [duration_seconds]
#   4. The script will start capturing after a 3-second countdown.
#   5. Perform the action (e.g. right-click → Convert to DOCX → file appears).
#   6. Press Ctrl+C when done. The script assembles the PNGs into a GIF.
#
# Examples:
#   scripts/record-gif.sh media/demo-md-to-docx 8
#   scripts/record-gif.sh media/demo-batch 15

set -e

OUTPUT_NAME="${1:-media/demo}"
DURATION="${2:-10}"
FRAME_DIR=$(mktemp -d -t mh-record-XXXXXX)
TRIMMED_DIR=$(mktemp -d -t mh-trim-XXXXXX)

# Detect OS
case "$(uname -s)" in
    Linux*)     OS="linux" ;;
    Darwin*)    OS="mac" ;;
    CYGWIN*|MINGW*|MSYS*) OS="windows" ;;
    *)          OS="unknown" ;;
esac

echo "Markdown Hub GIF recorder"
echo "=========================="
echo "Output:    $OUTPUT_NAME.gif"
echo "Duration: ~$DURATION seconds (Ctrl+C to stop earlier)"
echo "Platform:  $OS"
echo "Frames:    $FRAME_DIR"
echo ""

# Pick the right screen capture tool
case "$OS" in
    linux)
        # Try ffmpeg x11grab first, then gst-launch, then ImageMagick
        if command -v ffmpeg >/dev/null 2>&1; then
            CAPTURE_CMD="ffmpeg"
            # Detect display
            DISPLAY_ID="${DISPLAY:-:0}"
            # Use 900x600 around center of primary monitor
            FFMPEG_INPUT="-f x11grab -video_size 900x600 -i ${DISPLAY_ID}+100,100"
        else
            echo "ERROR: ffmpeg not found in PATH. Install it or use ScreenToGif on Windows."
            exit 1
        fi
        ;;
    mac)
        if command -v ffmpeg >/dev/null 2>&1; then
            CAPTURE_CMD="ffmpeg"
            FFMPEG_INPUT="-f avfoundation -i 0"  # macOS screen capture
        else
            echo "ERROR: ffmpeg not found. Install via: brew install ffmpeg"
            exit 1
        fi
        ;;
    windows)
        echo "ERROR: This bash script is awkward on Windows. Use ScreenToGif instead:"
        echo "  1. Download from https://www.screentogif.com/"
        echo "  2. Run it, select the VS Code window, hit Record"
        echo "  3. Perform the action, then Stop and Save as '$OUTPUT_NAME.gif'"
        echo "  4. Move the resulting .gif to the media/ folder"
        exit 1
        ;;
    *)
        echo "ERROR: Unknown OS: $OS"
        exit 1
        ;;
esac

# 3-second countdown so the user can switch to the target window
echo "Starting in 3 seconds — switch to your VS Code window now..."
sleep 1; echo "2..."; sleep 1; echo "1..."; sleep 1; echo "Recording!"

# Capture frames
ffmpeg $FFMPEG_INPUT \
    -framerate 10 \
    -t "$DURATION" \
    "$FRAME_DIR/frame_%04d.png" \
    -y -loglevel error

echo ""
echo "Capture complete: $(ls $FRAME_DIR/frame_*.png 2>/dev/null | wc -l) frames."

# Auto-trim first/last frame (often contain mouse artifact from clicking Record)
first_frame=$(ls $FRAME_DIR/frame_*.png | head -1)
last_frame=$(ls $FRAME_DIR/frame_*.png | tail -1)
echo "Trimming edges: dropping $first_frame and $last_frame"
mkdir -p "$TRIMMED_DIR"
for f in $(ls $FRAME_DIR/frame_*.png | sed '1d;$d'); do
    cp "$f" "$TRIMMED_DIR/"
done

# Trim the canvas: find the smallest bounding box that contains
# all non-uniform pixels (i.e. drop solid white margins around the demo).
# We use a heuristic: just resize to 900x600 to keep things simple.
# (Manual cropping in ffmpeg would be -vf "crop=W:H:X:Y".)

# Assemble the GIF
echo "Assembling GIF..."
ffmpeg -framerate 10 -i "$TRIMMED_DIR/frame_%04d.png" \
    -vf "scale=900:-1:flags=lanczos" \
    -loop 0 \
    -y "$OUTPUT_NAME.gif" \
    -loglevel error

# Optional: optimize with gifsicle
if command -v gifsicle >/dev/null 2>&1; then
    echo "Optimizing with gifsicle..."
    gifsicle -O3 --lossy=20 -k 256 "$OUTPUT_NAME.gif" -o "$OUTPUT_NAME.opt.gif"
    mv "$OUTPUT_NAME.opt.gif" "$OUTPUT_NAME.gif"
else
    echo "(Install gifsicle for size optimization; current GIF may be > 5 MB)"
fi

# Cleanup
rm -rf "$FRAME_DIR" "$TRIMMED_DIR"

echo ""
echo "Done: $OUTPUT_NAME.gif"
ls -lh "$OUTPUT_NAME.gif"
echo ""
echo "Next: review the GIF (does it show the action clearly?), then move it to media/."
