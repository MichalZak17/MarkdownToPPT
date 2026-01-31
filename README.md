# Markdown to PPTX Generator

Because life is too short to drag textboxes around in PowerPoint like it's 2003.

Write your presentation in Markdown. Run one command. Get a clean `.pptx` file. No clicking, no dragging, no existential dread.

## Why Does This Exist?

You know the drill. Someone asks for "a quick presentation." Three hours later you're manually aligning bullet points, wondering where it all went wrong. Your coffee is cold. Your will to live is fading. PowerPoint has won again.

Not anymore. Write Markdown. Get slides. Go home early.

## Installation

```bash
pip install python-pptx typer rich Pillow
```

Yes, that's it. No 47-step setup guide. No Docker. No YAML config files the size of a novel.

## Quick Start

```bash
python generator.py build markdown/ -t "My Presentation" -a "Your Name"
```

Your presentation lands in the `exports/` folder with a unique timestamp, so you'll never accidentally overwrite your masterpiece with a worse version (we've all been there).

Want to see what you're about to generate without actually generating it?

```bash
python generator.py preview markdown/
```

## Project Structure

```
your-project/
├── generator.py          <- the brain
├── markdown/             <- your content goes here
│   ├── 00_intro.md
│   ├── 01_basics.md
│   ├── 02_advanced.md
│   └── images/           <- your images go here
│       ├── diagram.png
│       ├── screenshot.webp
│       └── whatever.jpg
└── exports/              <- generated presentations end up here
```

Files are sorted by the number at the start of their filename. `00_` comes first, `01_` comes second. Revolutionary stuff, I know.

## How to Write Your Markdown

Each `.md` file = one chapter. Each chapter gets a fancy title slide and an optional agenda slide for free.

### The Frontmatter

Put this at the top of each file. It's optional, but skipping it is like showing up to a meeting without pants -- technically possible, but people will notice.

```markdown
---
title: My Chapter Title
agenda: What we'll cover in this chapter. This shows up on its own slide.
---
```

| Field | What It Does |
|-------|--------------|
| `title` | Chapter title. If you skip it, the filename gets promoted (and `02_advanced` becomes "Advanced" -- not terrible, honestly) |
| `agenda` | A one-liner describing the chapter. Gets its own slide. Leave it out if you like mystery |

### Slides

Separate slides with `---` on its own line. Each slide needs a heading (`#`) as its title.

```markdown
# First Slide

- Point one
- Point two
- Point three that nobody will read

---

# Second Slide

- More points
- Because bullet points are how adults communicate now
```

### Code Blocks

Wrap code in triple backticks with a language tag. It gets rendered on a dark background with the language label, looking significantly more professional than your actual code.

````markdown
# Code Example

```python
def solve_everything():
    return 42
```
````

### Images

Drop your image files into `markdown/images/` and reference them by filename. Supports PNG, JPEG, GIF, BMP, TIFF, WebP, and basically anything Pillow can open. Exotic formats get auto-converted to PNG behind the scenes.

#### Image-only slide

Just the image, no bullet points. It gets centered on the slide.

```markdown
# Architecture Diagram

![](architecture.png)
```

#### Image + text (two-column layout)

Add an image alongside bullet points and it automatically goes into a two-column layout -- text on the left, image on the right. No overlap, no manual positioning, no tears.

```markdown
# How It Works

- Step one: write markdown
- Step two: run the generator
- Step three: mass adoration from colleagues

![](workflow.png)
```

You reference images by bare filename (e.g., `screenshot.webp`) and they're resolved from the `images/` folder automatically. Full relative paths like `images/screenshot.webp` also work if you enjoy typing.

### Inline Formatting

Your bullet points support inline markdown:

| Syntax | Result |
|--------|--------|
| `` `code` `` | Rendered in monospace with accent color |
| `**bold**` | Bold text |
| `*italic*` | Italic text |

So `<div>` will show up as a nice orange monospace snippet in the presentation instead of the raw backtick mess.

## Commands

### `build`

The main event. Turns your markdown into a presentation.

```bash
python generator.py build [INPUT_DIR] [OPTIONS]
```

| Option | Description | Default |
|--------|-------------|---------|
| `-o`, `--output` | Custom output path | `exports/<title>_<timestamp>.pptx` |
| `-t`, `--title` | Presentation title (shown on the title slide) | `Presentation` |
| `-a`, `--author` | Author name (shown below the title) | (none) |

Output goes to the `exports/` folder by default, with a timestamp in the filename. Every generation is unique. Like snowflakes, but useful.

### `preview`

See the structure without generating anything. Good for checking you didn't accidentally put 47 slides in one chapter.

```bash
python generator.py preview [INPUT_DIR]
```

## What Gets Generated

For the curious (or suspicious):

1. **Title slide** -- your presentation title and author on a dark blue background
2. **For each markdown file (chapter):**
   - Chapter title slide with chapter number
   - Agenda slide (if you wrote an `agenda` in frontmatter)
   - All your content slides, code slides, and image slides in order

Everything is 16:9 widescreen. Because 4:3 died with overhead projectors.

## Theming

The default theme uses dark blue headings, orange accents, and a dark code background. It looks professional enough that nobody will question it.

Want different colors? Edit the `Theme` class in `generator.py`:

```python
class Theme:
    PRIMARY = RGBColor(0x1E, 0x3A, 0x5F)      # headings, header bar
    SECONDARY = RGBColor(0x3D, 0x5A, 0x80)    # chapter title backgrounds
    ACCENT = RGBColor(0xE0, 0x8E, 0x45)       # accent line, inline code
    TEXT_DARK = RGBColor(0x2D, 0x3A, 0x4A)     # body text
    TEXT_LIGHT = RGBColor(0xFF, 0xFF, 0xFF)    # text on dark backgrounds
    BG_CODE = RGBColor(0x2D, 0x2D, 0x2D)      # code block background
```

Change the hex values, run the build again, enjoy your new corporate-approved color scheme.

## Tips and Tricks

- **Chapter ordering**: Prefix filenames with numbers (`00_`, `01_`, `02_`). They sort lexicographically, so `10_` comes after `09_`, not after `1_`.
- **No frontmatter?** No problem. The filename becomes the chapter title. `03_javascript_basics.md` turns into "Javascript Basics."
- **Empty slides**: If a section between `---` separators has no content, it gets skipped. The generator judges silently but moves on.
- **Missing images**: If an image file doesn't exist, you get a placeholder text saying so. The presentation still builds. We're not monsters.
- **Multiple builds**: Each build creates a new timestamped file. Your `exports/` folder is your version history. Old school, but it works.

## License

MIT -- do whatever you want with it. No warranty, no support hotline, no refunds on time spent making presentations.
