---
title: HTML and CSS – Part 1
agenda: HTML basics, images, hyperlinks, lists, tables, and forms.
---

# Developer Tools
- Firefox: Menu -> Tools -> Toggle Tools (Ctrl + Shift + I / F12)
- Chrome: Menu -> More Tools -> Developer Tools (Ctrl + Shift + I / F12)
- Edge: More -> Developer Tools (F12)

---

# Example – Hello World
```html
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Hello World!</title>
</head>
<body>
    Hello World!
</body>
</html>
```

---

# Document Structure
- `<!DOCTYPE html>`: Declares the document type as HTML 5.
- `<html>`: Beginning and end of the document.
- `<head>`: Metadata, page title, linking styles and scripts.
- `<body>`: All visible content of the page.

![](html_structure.png)

---

# Text and Formatting
- `<p></p>`: Creating paragraphs.
- `<b>` / `<strong>`: Bold and emphasis.
- `<i>` / `<em>`: Italics and emphasis.
- `<h1>` - `<h6>`: Six heading levels.

---

# Images (img)
- `src` attribute: Path to the file (relative or absolute).
- `alt` attribute: Alternative text for search engines and screen readers.
- Styling: Width and height can be set in px or %.

---

# Forms
- `<form action="..." method="...">`: Form container.
- `<input type="text">`: Text field.
- `<input type="password">`: Password field.
- `<textarea>`: Large text areas.
- `<button>`: Custom and submit buttons.
