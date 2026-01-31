<div align="center">

# Markdown to PPTX Generator

**Because life is too short to drag textboxes around in PowerPoint like it's 2003.**

Write your presentation in Markdown. Run one command. Get a clean `.pptx` file.<br>
No clicking. No dragging. No existential dread.

[![Python](https://img.shields.io/badge/Python-3.10+-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://python.org)
[![License](https://img.shields.io/badge/License-MIT-E08E45?style=for-the-badge)](LICENSE)
[![Platform](https://img.shields.io/badge/Platform-Linux%20%7C%20macOS%20%7C%20Windows-1E3A5F?style=for-the-badge)](#)

[![python-pptx](https://img.shields.io/badge/python--pptx-Presentation%20Engine-blue?style=flat-square)](https://python-pptx.readthedocs.io/)
[![Typer](https://img.shields.io/badge/Typer-CLI%20Framework-blue?style=flat-square)](https://typer.tiangolo.com/)
[![Pillow](https://img.shields.io/badge/Pillow-Image%20Processing-blue?style=flat-square)](https://pillow.readthedocs.io/)

</div>

---

## Table of Contents

- [Why Does This Exist?](#-why-does-this-exist)
- [Installation](#-installation)
- [Quick Start](#-quick-start)
- [Project Structure](#-project-structure)
- [Writing Markdown](#-writing-your-markdown)
  - [Frontmatter](#the-frontmatter)
  - [Slides](#slides)
  - [Code Blocks](#code-blocks)
  - [Images](#images)
  - [Inline Formatting](#inline-formatting)
- [Commands](#-commands)
- [What Gets Generated](#-what-gets-generated)
- [Theming](#-theming)
- [Tips and Tricks](#-tips-and-tricks)
- [License](#-license)

---

## 🤔 Why Does This Exist?

You know the drill. Someone asks for "a quick presentation." Three hours later you're manually aligning bullet points, wondering where it all went wrong. Your coffee is cold. Your will to live is fading. PowerPoint has won again.

**Not anymore.** Write Markdown. Get slides. Go home early.

> *"I used to mass-produce presentations by hand. Now I mass-produce them with a single command. Progress."*
> — Every user of this tool, probably

---

## 📦 Installation

```bash
pip install python-pptx typer rich Pillow
```

Yes, that's it. No 47-step setup guide. No Docker compose files. No YAML configs the size of a novel.

| Dependency | Why |
|:-----------|:----|
| `python-pptx` | The thing that actually makes PowerPoint files |
| `typer` | CLI framework so you can feel like a hacker |
| `rich` | Pretty terminal output, because we have standards |
| `Pillow` | Image processing — handles WebP, AVIF, and other formats PowerPoint has never heard of |

---

## 🚀 Quick Start

Generate a presentation:

```bash
python generator.py build markdown/ -t "My Presentation" -a "Your Name"
```

Your presentation lands in the `exports/` folder with a unique timestamp, so you'll never accidentally overwrite your masterpiece with a worse version (we've all been there).

Want to see what you're about to generate without actually generating it?

```bash
python generator.py preview markdown/
```

<details>
<summary><strong>Example output</strong></summary>

```
Presentation Preview

Chapter 0: HTML and CSS – Part 1
  Agenda: HTML basics, images, hyperlinks, lists, tables, and forms.
    TXT Developer Tools
    CODE Example – Hello World
    TXT Document Structure
    TXT Text and Formatting

Chapter 1: HTML and CSS – Part 2
  Agenda: Cascading Style Sheets (CSS), page layout...
    TXT CSS Basics
    TXT Selectors and Cascading
    TXT Page Layout

Total slides: 39
```

</details>

---

## 🗂️ Project Structure

```
your-project/
├── generator.py              ← the brain
├── markdown/                 ← your content goes here
│   ├── 00_intro.md
│   ├── 01_basics.md
│   ├── 02_advanced.md
│   └── images/               ← your images go here
│       ├── diagram.png
│       ├── screenshot.webp
│       └── whatever.jpg
└── exports/                  ← generated presentations end up here
    └── my_talk_20260201_143052.pptx
```

Files are sorted by the number at the start of their filename. `00_` comes first, `01_` comes second. Revolutionary stuff, I know.

---

## ✍️ Writing Your Markdown

Each `.md` file = one chapter. Each chapter gets a fancy title slide and an optional agenda slide for free.

### The Frontmatter

Put this at the top of each file. It's optional, but skipping it is like showing up to a meeting without pants — technically possible, but people will notice.

```markdown
---
title: My Chapter Title
agenda: What we'll cover in this chapter. This shows up on its own slide.
---
```

| Field | What It Does |
|:------|:-------------|
| `title` | Chapter title. If you skip it, the filename gets promoted (and `02_advanced` becomes "Advanced" — not terrible, honestly) |
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

Drop your image files into `markdown/images/` and reference them by filename.

> **Supported formats:** PNG, JPEG, GIF, BMP, TIFF, WebP — basically anything Pillow can open. Exotic formats get auto-converted to PNG behind the scenes, because PowerPoint is still catching up with the 21st century.

#### Image-only slide

Just the image, no bullet points. It gets centered on the slide.

```markdown
# Architecture Diagram

![](architecture.png)
```

#### Image + text (two-column layout)

Add an image alongside bullet points and it automatically goes into a two-column layout — text on the left, image on the right. No overlap, no manual positioning, no tears.

```markdown
# How It Works

- Step one: write markdown
- Step two: run the generator
- Step three: mass adoration from colleagues

![](workflow.png)
```

You reference images by bare filename (e.g., `screenshot.webp`) and they're resolved from the `images/` folder automatically. Full relative paths like `images/screenshot.webp` also work if you enjoy typing.

### Inline Formatting

Your bullet points support inline markdown formatting:

| Syntax | Renders As |
|:-------|:-----------|
| `` `code` `` | Monospace text in accent color |
| `**bold**` | **Bold text** |
| `*italic*` | *Italic text* |

So `` `<div>` `` will show up as a styled monospace snippet in the presentation instead of raw backtick garbage.

---

## ⚙️ Commands

### `build`

The main event. Turns your markdown into a presentation.

```bash
python generator.py build [INPUT_DIR] [OPTIONS]
```

| Option | Description | Default |
|:-------|:------------|:--------|
| `-o`, `--output` | Custom output path | `exports/<title>_<timestamp>.pptx` |
| `-t`, `--title` | Presentation title (shown on the title slide) | `Presentation` |
| `-a`, `--author` | Author name (shown below the title) | *(none)* |

Output goes to the `exports/` folder by default, with a timestamp in the filename. Every generation is unique. Like snowflakes, but useful.

### `preview`

See the structure without generating anything. Good for checking you didn't accidentally put 47 slides in one chapter.

```bash
python generator.py preview [INPUT_DIR]
```

---

## 📊 What Gets Generated

For the curious (or suspicious):

```
┌─────────────────────────────────────────────┐
│              TITLE SLIDE                     │
│         Your title + author name             │
├─────────────────────────────────────────────┤
│  For each .md file:                          │
│  ┌────────────────────────────────────────┐  │
│  │  Chapter Title Slide (number + title)  │  │
│  │  Agenda Slide (if frontmatter has it)  │  │
│  │  Content Slides (bullets, code, imgs)  │  │
│  └────────────────────────────────────────┘  │
└─────────────────────────────────────────────┘
```

Everything is **16:9 widescreen**. Because 4:3 died with overhead projectors and we don't speak of it.

---

## 🎨 Theming

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

<details>
<summary><strong>Color reference</strong></summary>

| Variable | Color | Used For |
|:---------|:------|:---------|
| `PRIMARY` | ![#1E3A5F](https://via.placeholder.com/12/1E3A5F/1E3A5F.png) `#1E3A5F` | Headings, accent bar |
| `SECONDARY` | ![#3D5A80](https://via.placeholder.com/12/3D5A80/3D5A80.png) `#3D5A80` | Chapter title slide backgrounds |
| `ACCENT` | ![#E08E45](https://via.placeholder.com/12/E08E45/E08E45.png) `#E08E45` | Top accent line, inline code |
| `TEXT_DARK` | ![#2D3A4A](https://via.placeholder.com/12/2D3A4A/2D3A4A.png) `#2D3A4A` | Body text |
| `BG_CODE` | ![#2D2D2D](https://via.placeholder.com/12/2D2D2D/2D2D2D.png) `#2D2D2D` | Code block background |

</details>

---

## 💡 Tips and Tricks

- **Chapter ordering** — Prefix filenames with numbers (`00_`, `01_`, `02_`). They sort lexicographically, so `10_` comes after `09_`, not after `1_`. Math is hard.
- **No frontmatter?** No problem. The filename becomes the chapter title. `03_javascript_basics.md` turns into "Javascript Basics." Lazy, but effective.
- **Empty slides** — If a section between `---` separators has no content, it gets skipped. The generator judges silently but moves on.
- **Missing images** — If an image file doesn't exist, you get a placeholder text saying so. The presentation still builds. We're not monsters.
- **Multiple builds** — Each build creates a new timestamped file. Your `exports/` folder is your version history. Old school, but it works.
- **WebP and friends** — Throw any modern image format at it. If PowerPoint can't handle it natively (spoiler: it can't handle most things), Pillow converts it to PNG automatically.

---

## 📃 License

**MIT** — do whatever you want with it.

No warranty. No support hotline. No refunds on time spent making presentations.

```
THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND.
ALSO WITHOUT SYMPATHY FOR YOUR DEADLINE.
```

---

<div align="center">

**Made with mild frustration and a distaste for GUI presentation editors.**

*If this saved you time, consider starring the repo. It won't fix PowerPoint, but it'll make me feel something.*

</div>
