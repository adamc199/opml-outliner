# OPML Outliner

A desktop OPML outliner built with Python and PyQt6. Designed for managing hierarchical outlines in the OPML 2.0 format with inline HTML rendering, rich text formatting, and live include nodes.

![Platform](https://img.shields.io/badge/platform-Linux%20%7C%20macOS%20%7C%20Windows-blue)
![Python](https://img.shields.io/badge/python-3.10%2B-blue)
![License](https://img.shields.io/badge/license-MIT-green)

## Features

- **OPML 2.0** — full read/write support
- **Inline HTML rendering** — bold, italic, images, and links displayed directly in the tree
- **Live include nodes** — embed another OPML file (local or remote URL) as a subtree; refresh on demand
- **Rich text editing** — toggle `<b>` and `<i>` tags with `Ctrl+B` / `Ctrl+I` while editing
- **Hyperlinks** — attach a URL to any node; `Ctrl+Enter` opens it in the browser
- **Import from URL** — fetch and parse any OPML document from the web
- **HTML export** — export the full outline as a numbered, formatted HTML document
- **Paste as nodes** — paste multi-line text from the clipboard; each line becomes a child node
- **Undo/redo** — full history for all edits
- **Customisable appearance** — font, background colour, text colour, and line spacing; all preferences are persisted
- **Recent files** menu
- **Per-file state** — expand/collapse, current selection, scroll position, and window geometry are persisted inside the OPML on save and restored on open
- **Launch restore** — the last-opened file is reopened on start
- **Dark UI** by default; fully themeable

## Requirements

- Python 3.10+
- PyQt6

```bash
pip install PyQt6
```

No other dependencies are required.

## Running

```bash
python opml-outliner.py
```

Or make it executable:

```bash
chmod +x opml-outliner.py
./opml-outliner.py
```

### Opening a file at launch

```bash
python opml-outliner.py /path/to/my-outline.opml
```

## Keyboard Shortcuts

### Navigation

| Shortcut | Action |
|----------|--------|
| `Arrow keys` | Move between nodes |
| `Ctrl+Enter` | Expand/collapse node — or open URL if node has a link |
| `F2` | Edit current node inline |

### Editing

| Shortcut | Action |
|----------|--------|
| `Enter` | Add sibling node below current |
| `Ctrl+Shift+Enter` | Add sibling node below current (alternate) |
| `Tab` | Indent node (make child of previous sibling) |
| `Shift+Tab` | Outdent node (promote one level up) |
| `Ctrl+D` | Move node down |
| `Ctrl+U` | Move node up |
| `Ctrl+C` | Copy node (with all children) |
| `Ctrl+X` | Cut node |
| `Ctrl+V` | Paste node |
| `Delete` | Delete node and its children |

### Formatting (while editing a node)

| Shortcut | Action |
|----------|--------|
| `Ctrl+B` | Toggle **bold** on selected text |
| `Ctrl+I` | Toggle *italic* on selected text |

### File

| Shortcut | Action |
|----------|--------|
| `Ctrl+N` | New outline |
| `Ctrl+O` | Open OPML file |
| `Ctrl+S` | Save |
| `Ctrl+Shift+S` | Save As |
| `Ctrl+Z` | Undo |
| `Ctrl+Shift+Z` | Redo |

## Include Nodes

Include nodes embed the contents of another OPML file as a live subtree.

1. Use **Node > Add Include Node** from the menu, or add an `includeUrl` attribute manually
2. The include source can be a local file path or an `http://` / `https://` URL
3. Press `Ctrl+Enter` on an include node to fetch and expand its contents
4. Each expansion fetches fresh content — useful for shared outline libraries

Include nodes are marked with a `📄` badge in the tree.

## Adding Links

1. Select a node and choose **Node > Add/Edit Link**
2. Enter the URL
3. Press `Ctrl+Enter` to open the link in the default browser

Nodes with links show a `🔗` badge.

## Paste as Nodes

Copy any multi-line text to the clipboard, then use **Edit > Paste as Nodes**. Each non-blank line is inserted as a child node under the currently selected node.

## HTML Export

**File > Export as HTML** generates a self-contained HTML file with:
- Numbered outline hierarchy
- Preserved bold/italic formatting
- Clickable hyperlinks
- Inline images

## OPML Format

The outliner reads and writes standard OPML 2.0. A few non-standard attributes and a `<state>` element are used to persist UI state inside the file itself; all are optional and ignored by any standards-compliant OPML reader.

### `<outline>` attributes

| Attribute | Purpose |
|-----------|---------|
| `text` | Node display text (may contain HTML tags) |
| `url` | Hyperlink attached to the node |
| `xmlUrl` | Path or URL of an OPML file to embed (include node) |
| `_expanded` | `"true"` when the node has children and is expanded |

### `<head><state>` attributes

| Attribute | Purpose |
|-----------|---------|
| `selectedPath` | 0-based index path to the currently selected node (e.g. `0/2/1`) |
| `scrollY` | Vertical scroll position of the tree view |
| `windowGeometry` | Qt-encoded hex blob of window position/size; multi-monitor and maximized state aware |

### Where state lives

- **Inside the OPML file:** expand/collapse, selection, scroll, window geometry (written on every save)
- **In `~/.config/opml-outliner/prefs.json`:** appearance preferences, last-opened file path, recent files

## License

MIT — see [LICENSE](LICENSE).
