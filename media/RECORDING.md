# Recording demo GIFs

The placeholders in `media/*.svg` need to be replaced with real
screen recordings. Here's the workflow.

## Quick start (Windows, recommended)

[ScreenToGif](https://www.screentogif.com/) is the easiest tool:

1. Download & install ScreenToGif
2. Open the file `media/sample.md` (or any markdown file) in VS Code
3. Open ScreenToGif → Recorder
4. Drag a region around the VS Code window
5. Click Record
6. Perform the action (e.g. right-click → Convert to DOCX)
7. Click Stop → save as `media/md-to-docx.gif`
8. Repeat for the other demos

## Quick start (Linux/macOS, scripted)

`scripts/record-gif.sh` automates the capture with `ffmpeg`:

```bash
# Install ffmpeg
#   Linux:  apt install ffmpeg
#   macOS:  brew install ffmpeg
#   Win:    https://www.gyan.dev/ffmpeg/builds/

# Optional: install gifsicle for size optimization
#   Linux:  apt install gifsicle
#   macOS:  brew install gifsicle

# Record (script handles 3-second countdown)
./scripts/record-gif.sh media/md-to-docx 8   # ~8-second clip
./scripts/record-gif.sh media/batch 12
./scripts/record-gif.sh media/pptx 10
```

## Recommended demos (3 short clips, each < 1 MB)

### 1. `md-to-docx.gif` (8s)
- Show a `.md` file open in the editor
- Right-click → "Convert to DOCX"
- Toast notification: "Conversion complete"
- Status bar / output panel showing the success message

### 2. `batch.gif` (12s)
- Show Explorer with a folder of 5-10 `.md` files
- Right-click the folder → "Batch: Markdown → DOCX"
- Output panel scrolling through progress
- Final output folder shown with all the new `.docx` files

### 3. `pptx.gif` (10s)
- Show a `.md` file with a heading, table, code block, and list
- Right-click → "Convert to PPTX"
- The output `.pptx` file in Explorer (or thumbnail)
- Optional: a quick second click to open the .pptx in PowerPoint and
  show the first slide

## Tips for good demos

- **Resolution**: 900×600 is ideal (small file, fits GitHub README)
- **Framerate**: 10-15 fps is enough (lowers file size 5-10x vs 30 fps)
- **Length**: 5-12 seconds max — longer and viewers bounce
- **Mouse**: hide the mouse cursor if possible (ScreenToGif has this option)
- **Font**: use 14-16px in VS Code editor for readability in the GIF
- **Theme**: use a light or dark theme consistently — don't switch mid-recording
- **No audio**: GitHub README GIFs are silent (no point recording sound)

## File size budget

GitHub README renders images at ~600px wide. Aim for:
- Each GIF: **< 1 MB** (preferably < 500 KB)
- Total of 3 GIFs: **< 2 MB**

If a GIF is too big:
1. Reduce framerate (15 → 10 fps)
2. Reduce resolution (900 → 720 wide)
3. Reduce colors with gifsicle:
   ```bash
   gifsicle -O3 --lossy=30 -k 128 input.gif -o output.gif
   ```
4. Trim repeated frames (ScreenToGif does this automatically)
