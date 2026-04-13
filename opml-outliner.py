#!/usr/bin/env python3
"""
OPML Outliner - Tree-based with proper collapse/expand
Beautiful fonts and readable text with line wrapping
"""

import sys
import math

# Increase recursion limit to prevent deep QTextDocument layouts from blowing up
sys.setrecursionlimit(10000)
import xml.etree.ElementTree as ET
import html
import webbrowser
import textwrap
import json
import re
import urllib.request
import gzip
import threading
from pathlib import Path
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QTreeWidget, QTreeWidgetItem,
    QVBoxLayout, QHBoxLayout, QWidget, QPushButton, QFileDialog,
    QMessageBox, QInputDialog, QLabel, QStatusBar, QMenu, QMenuBar,
    QColorDialog, QFontDialog, QDialog, QLineEdit, QDialogButtonBox,
    QStyledItemDelegate, QStyleOptionViewItem, QTextEdit, QAbstractItemView,
    QSizePolicy
)
from PyQt6.QtCore import Qt, QTimer, QRect, QSize, QUrl
from PyQt6.QtGui import QKeySequence, QShortcut, QFont, QColor, QPainter, QTextDocument, QUndoStack, QUndoCommand


class FormattingTextEdit(QTextEdit):
    """Custom QTextEdit that handles Ctrl+B and Ctrl+I"""
    def keyPressEvent(self, event):
        if event.modifiers() == Qt.KeyboardModifier.ControlModifier:
            if event.key() == Qt.Key.Key_B:
                self.toggle_format('b')
                return
            elif event.key() == Qt.Key.Key_I:
                self.toggle_format('i')
                return
        super().keyPressEvent(event)
    
    def toggle_format(self, tag):
        """Toggle HTML formatting tag - add or remove"""
        cursor = self.textCursor()
        full_text = self.toPlainText()
        
        if cursor.hasSelection():
            start = cursor.selectionStart()
            end = cursor.selectionEnd()
            selected = full_text[start:end]
            
            # Check if already formatted - if so, remove tags
            open_tag = f'<{tag}>'
            close_tag = f'</{tag}>'
            
            if open_tag in selected and close_tag in selected:
                # Remove tags
                new_text = full_text[:start] + selected.replace(open_tag, '').replace(close_tag, '') + full_text[end:]
            else:
                # Add tags
                new_text = full_text[:start] + f'<{tag}>{selected}</{tag}>' + full_text[end:]
            
            self.setPlainText(new_text)
            cursor.setPosition(start)
            cursor.setPosition(start + len(selected), cursor.MoveMode.KeepAnchor)
            self.setTextCursor(cursor)
        else:
            pos = cursor.position()
            # Insert tags at cursor
            new_text = full_text[:pos] + f'<{tag}></{tag}>' + full_text[pos:]
            self.setPlainText(new_text)
            cursor.setPosition(pos + len(tag) + 2)
            self.setTextCursor(cursor)


class EditableTreeWidget(QTreeWidget):
    """Custom tree widget that handles Escape key properly during editing"""
    def keyPressEvent(self, event):
        from PyQt6.QtWidgets import QTextEdit
        if event.key() == Qt.Key.Key_Escape:
            # Find the active editor
            editor = self.findChild(QTextEdit)
            if editor:
                # Get the text FIRST
                text = editor.toPlainText()
                # Close the editor with Accepted state - this will trigger setModelData
                self.closeEditor(editor, QTreeWidget.EditStrategy.Accepted)
                # NOW set the text on the item (after close so it doesn't get overwritten)
                item = self.currentItem()
                if item:
                    item.setText(0, text)
                    self.itemChanged.emit(item, 0)
                event.accept()
                return
        super().keyPressEvent(event)
    
    def closeEditor(self, editor, hint=None):
        # First save the data from editor to model
        from PyQt6.QtWidgets import QAbstractItemDelegate
        self.commitData(editor)
        super().closeEditor(editor, hint)


# Module-level image cache shared across all document instances.
# Maps URL string -> QImage, or None if the fetch failed.
_image_cache: dict = {}
_data_url_cache: dict = {}   # url_str -> "data:image/...;base64,..." string
_image_loading: set = set()


_repaint_callback = None


def _fetch_image(url_str):
    """Start a background fetch for url_str if not already cached or in flight."""
    if url_str in _image_cache or url_str in _image_loading:
        return
    _image_loading.add(url_str)

    def _worker(u=url_str):
        from PyQt6.QtGui import QImage
        try:
            if u.startswith('file://') or (len(u) > 1 and u[1] == ':') or u.startswith('/'):
                # Local path
                local = u[7:] if u.startswith('file://') else u
                img = QImage(local)
            else:
                req = urllib.request.Request(u, headers={'User-Agent': 'Mozilla/5.0'})
                with urllib.request.urlopen(req, timeout=10) as resp:
                    data = resp.read()
                img = QImage()
                img.loadFromData(data)
            if not img.isNull():
                if img.width() > 800:
                    img = img.scaledToWidth(800, Qt.TransformationMode.SmoothTransformation)
                _image_cache[u] = img
                # Build a data: URI so Qt doesn't need to do its own resource lookup
                import base64
                from PyQt6.QtCore import QBuffer, QIODevice
                buf = QBuffer()
                buf.open(QIODevice.OpenModeFlag.WriteOnly)
                img.save(buf, 'PNG')
                buf.close()
                b64 = base64.b64encode(bytes(buf.data())).decode('ascii')
                _data_url_cache[u] = f'data:image/png;base64,{b64}'
            else:
                _image_cache[u] = None
        except Exception:
            _image_cache[u] = None
        finally:
            _image_loading.discard(u)
            cb = _repaint_callback
            if cb:
                QTimer.singleShot(0, cb)

    threading.Thread(target=_worker, daemon=True).start()


class MultiLineDelegate(QStyledItemDelegate):
    """Delegate that renders and sizes multi-line text with HTML formatting"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self._editor_font = None
        self._text_color = "#000000"
        self._bg_color = "#ffffff"
        self._disable_rich = False
        self._force_plain_role = Qt.ItemDataRole.UserRole + 9

    def setRepaintCallback(self, callback):
        global _repaint_callback
        _repaint_callback = callback

    def setEditorFont(self, font):
        self._editor_font = font

    def setTextColor(self, color_str: str):
        self._text_color = color_str

    def setBgColor(self, color_str: str):
        self._bg_color = color_str
    
    def paint(self, painter, option, index):
        text = index.data(Qt.ItemDataRole.DisplayRole) or ""
        is_html = bool(index.data(Qt.ItemDataRole.UserRole + 3))

        painter.save()
        painter.setClipRect(option.rect)  # Prevent text overflow into adjacent rows
        from PyQt6.QtWidgets import QStyle
        is_selected = option.state & QStyle.StateFlag.State_Selected

        if is_selected:
            painter.fillRect(option.rect, option.palette.highlight())
            text_color = "#ffffff"
        else:
            text_color = self._text_color

        doc = QTextDocument()
        doc.setDefaultFont(option.font)
        doc.setDocumentMargin(0)

        if is_html:
            # Replace remote src URLs with data: URIs for reliable inline rendering.
            # Kick off background fetch for any not yet cached.
            srcs = re.findall(r'<img\b[^>]+src=["\']([^"\']+)["\']', text, re.IGNORECASE)
            html_text = text
            for src in srcs:
                data_url = _data_url_cache.get(src)
                if data_url:
                    html_text = html_text.replace(src, data_url, 1)
                else:
                    _fetch_image(src)
                    # Remove the img tag until loaded so layout isn't broken
                    html_text = re.sub(
                        r'<img\b[^>]+src=["\']' + re.escape(src) + r'["\'][^>]*>',
                        '', html_text, flags=re.IGNORECASE)
            doc.setHtml(f'<span style="color: {text_color};">{html_text}</span>')
        else:
            doc.setHtml(f'<span style="color: {text_color};">{html.escape(text)}</span>')

        doc.setTextWidth(max(120, option.rect.width() - 8))
        painter.translate(option.rect.x() + 4, option.rect.y() + 2)
        doc.drawContents(painter)
        painter.restore()

    def sizeHint(self, option, index):
        # While editing this item, match the editor's actual height so the tree
        # makes room and doesn't let the editor overlap the row below.
        if getattr(self, '_editing_index', None) is not None and index == self._editing_index:
            # Find the live editor widget and use its current height
            widget = option.widget
            if widget:
                editor = widget.findChild(FormattingTextEdit)
                if editor:
                    return QSize(option.rect.width() or 400, editor.height())

        text = index.data(Qt.ItemDataRole.DisplayRole)
        if text is None:
            return super().sizeHint(option, index)
        is_html = bool(index.data(Qt.ItemDataRole.UserRole + 3))
        viewport_w = (option.widget.viewport().width()
                      if option.widget and option.widget.viewport() else 400)
        if is_html:
            # Find first <img src="..."> and use the loaded image's height if cached
            img_match = re.search(r'<img\b[^>]+src=["\']([^"\']+)["\']', text, re.IGNORECASE)
            if img_match:
                url_str = img_match.group(1)
                cached = _image_cache.get(url_str)
                if cached is not None and not cached.isNull():
                    return QSize(viewport_w, cached.height() + 4)
                return QSize(viewport_w, 300)  # placeholder until image loads
        plain_text = re.sub(r'<[^>]+>', '', text)
        return self._plain_size_hint(plain_text, option, index)

    def _plain_size_hint(self, text, option, index=None):
        viewport_w = option.widget.viewport().width() if option.widget and option.widget.viewport() else 400
        doc = QTextDocument()
        doc.setDefaultFont(option.font)
        doc.setDocumentMargin(0)
        doc.setHtml(f'<span style="color: {self._text_color};">{html.escape(text[:50000])}</span>')
        doc.setTextWidth(self._size_hint_text_width(option, index))
        return QSize(viewport_w, math.ceil(doc.size().height()) + 4)

    def _item_depth(self, index):
        depth = 0
        p = index.parent()
        while p.isValid():
            depth += 1
            p = p.parent()
        return depth

    def _size_hint_text_width(self, option, index=None):
        """Compute the text wrap width for sizeHint — derived from current viewport + depth.

        Always use the live viewport width rather than option.rect.width(), which can be
        stale immediately after a resize and causes sizeHint to return a height calculated
        for the wrong wrap width, leaving rows too short and text bleeding into the row below.
        """
        if option.widget and option.widget.viewport():
            viewport_w = option.widget.viewport().width()
            tree_indent = option.widget.indentation() if hasattr(option.widget, 'indentation') else 20
            if index is not None and index.isValid():
                depth = self._item_depth(index)
                # (depth + 1): the root decorator expand arrow consumes one indentation
                # unit even at depth=0, so each item's text rect is (depth+1)*indent narrower.
                return max(120, viewport_w - (depth + 1) * tree_indent - 8)
            return max(120, viewport_w - tree_indent - 8)
        # Last resort when no widget is available
        if option.rect.width() > 0:
            return max(120, option.rect.width() - 8)
        return 400
    
    def createEditor(self, parent, option, index):
        self._editing_index = index
        editor = FormattingTextEdit(parent)
        editor.setAcceptRichText(False)
        editor.setLineWrapMode(QTextEdit.LineWrapMode.WidgetWidth)
        editor.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        editor.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        editor.setStyleSheet(
            "QTextEdit {"
            "  border: none;"
            "  border-left: 2px solid #0066cc;"
            "  padding: 2px 4px;"
            f" color: {self._text_color};"
            f" background: {self._bg_color};"
            "}"
        )
        if self._editor_font:
            editor.setFont(self._editor_font)
        editor.document().contentsChanged.connect(
            lambda: self._auto_resize_editor(editor, index)
        )
        return editor

    def destroyEditor(self, editor, index):
        self._editing_index = None
        super().destroyEditor(editor, index)

    def _auto_resize_editor(self, editor, index):
        doc_h = int(editor.document().size().height()) + 6
        min_h = editor.fontMetrics().lineSpacing() + 6
        new_h = max(doc_h, min_h)
        editor.setFixedHeight(new_h)
        # Tell the tree the item needs more space — triggers updateEditorGeometry
        self.sizeHintChanged.emit(index)

    def updateEditorGeometry(self, editor, option, index):
        rect = option.rect
        doc_h = int(editor.document().size().height()) + 6
        min_h = max(rect.height(), editor.fontMetrics().lineSpacing() + 6)
        height = max(doc_h, min_h)
        from PyQt6.QtCore import QRect
        editor.setGeometry(QRect(rect.x(), rect.y(), rect.width(), height))
    
    def setEditorData(self, editor, index):
        text = index.data(Qt.ItemDataRole.UserRole + 2) or index.data(Qt.ItemDataRole.DisplayRole)
        if text is None:
            text = ""
        editor.setPlainText(text)
        editor.setProperty("_index", index)

    def setModelData(self, editor, model, index):
        # Get text and convert back to single line for storage
        text = editor.toPlainText()
        model.setData(index, text, Qt.ItemDataRole.EditRole)


class OPMLOutliner(QMainWindow):
    def __init__(self):
        super().__init__()
        self.current_file = None
        self.wrap_width = 80  # Characters to wrap at
        self.clipboard_item = None  # For copy/cut/paste
        self.undo_stack = QUndoStack(self)
        self.is_dirty = False  # Track unsaved changes
        
        # Load preferences
        self.prefs_file = Path.home() / ".config" / "opml-outliner" / "prefs.json"
        self.load_preferences()
        
        self.init_ui()
        self.setup_menu()
        self.setup_shortcuts()
    
    def load_preferences(self):
        """Load saved preferences"""
        defaults = {
            "bg_color": "#ffffff",
            "text_color": "#1a1a1a",
            "font_family": "Georgia",
            "font_size": 13,
            "node_spacing": 3,  # padding in pixels
            "last_file": "",
            "last_folder": str(Path.home()),
            "recent_files": [],
        }

        try:
            if self.prefs_file.exists():
                with open(self.prefs_file, 'r') as f:
                    prefs = json.load(f)
                    self.bg_color = prefs.get("bg_color", defaults["bg_color"])
                    self.text_color = prefs.get("text_color", defaults["text_color"])
                    self.font_family = prefs.get("font_family", defaults["font_family"])
                    self.font_size = prefs.get("font_size", defaults["font_size"])
                    self.node_spacing = prefs.get("node_spacing", defaults["node_spacing"])
                    self.last_file = prefs.get("last_file", defaults["last_file"])
                    self.last_folder = prefs.get("last_folder", defaults["last_folder"])
                    self.recent_files = prefs.get("recent_files", defaults["recent_files"])
            else:
                self.bg_color = defaults["bg_color"]
                self.text_color = defaults["text_color"]
                self.font_family = defaults["font_family"]
                self.font_size = defaults["font_size"]
                self.node_spacing = defaults["node_spacing"]
                self.last_file = defaults["last_file"]
                self.last_folder = defaults["last_folder"]
                self.recent_files = defaults["recent_files"]
        except:
            self.bg_color = defaults["bg_color"]
            self.text_color = defaults["text_color"]
            self.font_family = defaults["font_family"]
            self.font_size = defaults["font_size"]
            self.node_spacing = defaults["node_spacing"]
            self.last_file = defaults["last_file"]
            self.last_folder = defaults["last_folder"]
            self.recent_files = defaults["recent_files"]
    
    def save_preferences(self):
        """Save current preferences"""
        prefs = {
            "bg_color": self.bg_color,
            "text_color": self.text_color,
            "font_family": self.font_family,
            "font_size": self.font_size,
            "node_spacing": self.node_spacing,
            "last_file": getattr(self, 'last_file', ''),
            "last_folder": getattr(self, 'last_folder', str(Path.home())),
            "recent_files": getattr(self, 'recent_files', []),
        }
        
        try:
            self.prefs_file.parent.mkdir(parents=True, exist_ok=True)
            with open(self.prefs_file, 'w') as f:
                json.dump(prefs, f, indent=2)
        except Exception as e:
            print(f"Failed to save preferences: {e}")

    def _add_recent_file(self, path):
        """Add path to the top of the recent files list and persist it."""
        path = str(path)
        recent = getattr(self, 'recent_files', [])
        if path in recent:
            recent.remove(path)
        recent.insert(0, path)
        self.recent_files = recent[:10]
        self.save_preferences()
        self._rebuild_recent_menu()

    def _rebuild_recent_menu(self):
        """Repopulate the Open Recent submenu from self.recent_files."""
        menu = getattr(self, 'recent_menu', None)
        if menu is None:
            return
        menu.clear()
        recent = [p for p in getattr(self, 'recent_files', []) if Path(p).exists()]
        if not recent:
            placeholder = menu.addAction("No recent files")
            placeholder.setEnabled(False)
            return
        for path in recent:
            label = Path(path).name
            action = menu.addAction(label)
            action.setToolTip(path)
            action.triggered.connect(lambda checked, p=path: self._open_recent(p))
        menu.addSeparator()
        clear_action = menu.addAction("Clear Recent Files")
        clear_action.triggered.connect(self._clear_recent_files)

    def _open_recent(self, path):
        if not Path(path).exists():
            QMessageBox.warning(self, "File Not Found", f"Could not find:\n{path}")
            self.recent_files = [p for p in self.recent_files if p != path]
            self.save_preferences()
            self._rebuild_recent_menu()
            return
        self.load_opml(path)
        self.current_file = path
        self.last_folder = str(Path(path).parent)
        self.setWindowTitle(f"OPML Outliner - {Path(path).name}")
        self._add_recent_file(path)

    def _clear_recent_files(self):
        self.recent_files = []
        self.save_preferences()
        self._rebuild_recent_menu()

    def init_ui(self):
        self.setWindowTitle("OPML Outliner")
        self.setGeometry(100, 100, 1600, 1000)

        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # File path label
        self.file_path_label = QLabel("")
        self.file_path_label.setStyleSheet("color: #666666; padding: 4px 16px 4px; font-size: 11pt;")
        self.file_path_label.setMinimumWidth(0)
        self.file_path_label.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Preferred)
        layout.addWidget(self.file_path_label)

        # Tree widget
        self.tree = EditableTreeWidget()
        self.tree.setHeaderLabels(["Outline"])
        self.tree.setHeaderHidden(True)
        
        # Custom delegate for multi-line height and editing
        tree_font = QFont(self.font_family, self.font_size)
        tree_font.setStyleHint(QFont.StyleHint.Serif)
        # Use custom delegate for HTML rendering
        self.delegate = MultiLineDelegate()
        self.delegate.setEditorFont(tree_font)
        self.delegate.setTextColor(self.text_color)
        self.delegate.setBgColor(self.bg_color)
        def _on_image_loaded():
            self.tree.viewport().update()
            self.tree.scheduleDelayedItemsLayout()
        self.delegate.setRepaintCallback(_on_image_loaded)
        self.tree.setItemDelegate(self.delegate)
        self.tree.setEditTriggers(
            QTreeWidget.EditTrigger.DoubleClicked | 
            QTreeWidget.EditTrigger.EditKeyPressed
        )
        self.tree.setIndentation(25)
        self.tree.setTextElideMode(Qt.TextElideMode.ElideNone)
        self.tree.setAnimated(False)
        self.tree.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.tree.setRootIsDecorated(True)  # Show expand/collapse triangles
        self.tree.setItemsExpandable(True)  # Enable expand/collapse
        self.tree.setExpandsOnDoubleClick(False)  # Don't expand on double-click (for editing)
        self.tree.setAllColumnsShowFocus(True)
        
        # Beautiful font
        tree_font = QFont(self.font_family, self.font_size)
        tree_font.setStyleHint(QFont.StyleHint.Serif)
        self.tree.setFont(tree_font)
        
        # Connect expand/collapse events to update triangles
        self.tree.itemExpanded.connect(lambda item: self.update_node_display(item, True))
        self.tree.itemCollapsed.connect(lambda item: self.update_node_display(item, True))
        
        # Connect item changed to update stored text after editing
        self.tree.itemChanged.connect(self.on_item_changed)
        
        # Auto-scroll to current item
        self.tree.currentItemChanged.connect(self.on_current_item_changed)

        layout.addWidget(self.tree)

        # Status bar
        self.status_bar = QStatusBar()
        self.mode_label = QLabel("Use ← → to collapse/expand, ↑↓ to navigate")
        self.mode_label.setStyleSheet("color: #0066cc; padding: 4px 8px;")
        self.status_bar.addPermanentWidget(self.mode_label)
        self.setStatusBar(self.status_bar)

        self.apply_initial_style()

    def on_current_item_changed(self, current, previous):
        if current:
            self.tree.scrollToItem(current, QAbstractItemView.ScrollHint.EnsureVisible)
    
    def apply_initial_style(self):
        """Apply initial window styling (not tree-specific)"""
        self.setStyleSheet(f"""
            QMainWindow {{
                background-color: #f9f9f7;
            }}
            
            QTreeWidget {{
                background-color: {self.bg_color};
                color: {self.text_color};
                border: none;
                outline: none;
                padding: 12px;
                font-size: {self.font_size}pt;
                show-decoration-selected: 1;
            }}
            
            QTreeWidget::item {{
                padding: {self.node_spacing}px 6px;
                border: none;
                min-height: {(self.node_spacing * 2) + 18}px;
            }}
            
            QTreeWidget::item:selected {{
                background-color: #e8f4f8;
                color: #0066cc;
                border-left: 3px solid #0066cc;
            }}
            
            QTreeWidget::item:hover {{
                background-color: transparent;
            }}
            
            QPushButton {{
                background-color: #0066cc;
                color: white;
                border: none;
                padding: 10px 20px;
                border-radius: 6px;
                font-size: 12pt;
            }}
            
            QPushButton:hover {{
                background-color: #0052a3;
            }}
            
            QPushButton:pressed {{
                background-color: #003d7a;
            }}
            
            QStatusBar {{
                background-color: #f0f0f0;
                border-top: 1px solid #d0d0d0;
            }}

            QMenuBar {{
                background-color: #2d2d2d;
                color: #f0f0f0;
                padding: 2px 0;
                font-size: 12pt;
            }}

            QMenuBar::item {{
                padding: 4px 12px;
                background: transparent;
            }}

            QMenuBar::item:selected {{
                background-color: #4a4a4a;
                border-radius: 4px;
            }}

            QMenu {{
                background-color: #2d2d2d;
                color: #f0f0f0;
                border: 1px solid #555;
                padding: 4px 0;
                font-size: 12pt;
            }}

            QMenu::item {{
                padding: 6px 24px 6px 16px;
            }}

            QMenu::item:selected {{
                background-color: #0066cc;
            }}

            QMenu::separator {{
                height: 1px;
                background: #555;
                margin: 4px 8px;
            }}
        """)
    
    def setup_menu(self):
        """Setup menu bar"""
        print("Setting up menu bar...")
        menubar = self.menuBar()
        menubar.setNativeMenuBar(False)

        # ── File ──────────────────────────────────────────────────────────────
        file_menu = menubar.addMenu("File")

        new_action = file_menu.addAction("New")
        new_action.triggered.connect(self.new_file)
        new_action.setShortcut(QKeySequence("Ctrl+N"))

        open_action = file_menu.addAction("Open…")
        open_action.triggered.connect(self.open_file)
        open_action.setShortcut(QKeySequence("Ctrl+O"))

        save_action = file_menu.addAction("Save")
        save_action.triggered.connect(self.save_file)
        save_action.setShortcut(QKeySequence("Ctrl+S"))
        save_action.setShortcutContext(Qt.ShortcutContext.ApplicationShortcut)

        save_as_action = file_menu.addAction("Save As…")
        save_as_action.triggered.connect(self.save_file_as)
        save_as_action.setShortcut(QKeySequence("Ctrl+Shift+S"))

        file_menu.addSeparator()

        self.recent_menu = file_menu.addMenu("Open Recent")
        self._rebuild_recent_menu()

        file_menu.addSeparator()

        import_action = file_menu.addAction("Import from URL…")
        import_action.triggered.connect(self.import_from_url)

        file_menu.addSeparator()

        export_html_action = file_menu.addAction("Export HTML…")
        export_html_action.triggered.connect(self.export_html)

        # ── Edit ──────────────────────────────────────────────────────────────
        edit_menu = menubar.addMenu("Edit")

        undo_action = edit_menu.addAction("Undo")
        undo_action.triggered.connect(self.undo)
        undo_action.setShortcut(QKeySequence("Ctrl+Z"))

        redo_action = edit_menu.addAction("Redo")
        redo_action.triggered.connect(self.redo)
        redo_action.setShortcut(QKeySequence("Ctrl+Y"))

        edit_menu.addSeparator()

        link_action = edit_menu.addAction("Add Link…")
        link_action.triggered.connect(self.add_link_to_node)
        link_action.setShortcut(QKeySequence("Ctrl+L"))

        edit_menu.addSeparator()

        include_action = edit_menu.addAction("Include from File/URL…")
        include_action.triggered.connect(self.add_include_node)

        refresh_action = edit_menu.addAction("Refresh Include")
        refresh_action.triggered.connect(self.refresh_include_node)
        refresh_action.setShortcut(QKeySequence("Ctrl+R"))

        paste_nodes_action = edit_menu.addAction("Paste as Nodes…")
        paste_nodes_action.triggered.connect(self.paste_as_nodes)
        paste_nodes_action.setShortcut(QKeySequence("Ctrl+Shift+V"))

        # ── View ──────────────────────────────────────────────────────────────
        view_menu = menubar.addMenu("View")

        font_action = view_menu.addAction("Change Font…")
        font_action.triggered.connect(self.change_font)

        view_menu.addSeparator()

        bg_action = view_menu.addAction("Background Color…")
        bg_action.triggered.connect(self.change_bg_color)

        text_action = view_menu.addAction("Text Color…")
        text_action.triggered.connect(self.change_text_color)

        spacing_action = view_menu.addAction("Node Spacing…")
        spacing_action.triggered.connect(self.change_spacing)

        view_menu.addSeparator()

        reset_action = view_menu.addAction("Reset to Defaults")
        reset_action.triggered.connect(self.reset_appearance)

    def change_font(self):
        """Change editor font"""
        current_font = QFont(self.font_family, self.font_size)
        font, ok = QFontDialog.getFont(current_font, self)
        if ok:
            self.font_family = font.family()
            self.font_size = font.pointSize()
            self.tree.setFont(font)
            self.apply_colors()
            self.save_preferences()

    def show_color_menu(self):
        """Show color options menu"""
        menu = QMenu(self)
        
        bg_action = menu.addAction("Background Color...")
        bg_action.triggered.connect(self.change_bg_color)
        
        text_action = menu.addAction("Text Color...")
        text_action.triggered.connect(self.change_text_color)
        
        menu.addSeparator()
        
        reset_action = menu.addAction("Reset Colors")
        reset_action.triggered.connect(self.reset_appearance)
        
        # Show menu at button
        menu.exec(self.sender().mapToGlobal(self.sender().rect().bottomLeft()))

    def change_bg_color(self):
        """Change background color"""
        color = QColorDialog.getColor(QColor(self.bg_color), self)
        if color.isValid():
            self.bg_color = color.name()
            self.apply_colors()
            self.save_preferences()

    def change_text_color(self):
        """Change text color"""
        color = QColorDialog.getColor(QColor(self.text_color), self)
        if color.isValid():
            self.text_color = color.name()
            self.apply_colors()
            self.save_preferences()
    
    def change_spacing(self):
        """Change node spacing"""
        spacing, ok = QInputDialog.getInt(
            self, "Node Spacing",
            "Enter padding in pixels (1-20):",
            value=self.node_spacing,
            min=1,
            max=20
        )
        if ok:
            self.node_spacing = spacing
            self.apply_colors()
            self.save_preferences()

    def reset_appearance(self):
        """Reset to default appearance"""
        self.bg_color = "#ffffff"
        self.text_color = "#1a1a1a"
        self.font_family = "Georgia"
        self.font_size = 13
        self.node_spacing = 3
        self.tree.setFont(QFont(self.font_family, self.font_size))
        self.apply_colors()
        self.save_preferences()

    def apply_colors(self):
        """Apply current color scheme"""
        min_height = (self.node_spacing * 2) + 18  # Calculate based on padding
        
        self.tree.setStyleSheet(f"""
            QTreeWidget {{
                background-color: {self.bg_color};
                color: {self.text_color};
                border: none;
                outline: none;
                padding: 12px;
                font-size: {self.font_size}pt;
                show-decoration-selected: 1;
            }}
            
            QTreeWidget::item {{
                padding: {self.node_spacing}px 6px;
                border: none;
                min-height: {min_height}px;
            }}
            
            QTreeWidget::item:selected {{
                background-color: #e8f4f8;
                color: #0066cc;
                border-left: 3px solid #0066cc;
            }}
            
            QTreeWidget::item:hover {{
                background-color: transparent;
            }}
        """)
        self.delegate.setTextColor(self.text_color)
        self.delegate.setBgColor(self.bg_color)

    def add_link_to_node(self):
        """Add or edit link on current node"""
        item = self.tree.currentItem()
        if not item:
            QMessageBox.warning(self, "No Node Selected", "Please select a node first")
            return
        
        current_url = item.data(0, Qt.ItemDataRole.UserRole) or ""
        
        url, ok = QInputDialog.getText(
            self, "Add Link", 
            "Enter URL (use .opml for inline include):",
            text=current_url
        )
        
        if ok:
            if url.strip():
                # Check if it's an OPML file
                url_lower = url.strip().lower()
                is_opml = url_lower.endswith('.opml') or 'opml' in url_lower
                
                # Add link
                item.setData(0, Qt.ItemDataRole.UserRole, url.strip())
                # Remove existing icons
                text = item.text(0).replace('🔗 ', '').replace('📄 ', '')
                
                if is_opml:
                    item.setText(0, f"📄 {text}")
                    item.setForeground(0, QColor("#008800"))
                else:
                    item.setText(0, f"🔗 {text}")
                    item.setForeground(0, QColor("#0066cc"))
            else:
                # Remove link
                item.setData(0, Qt.ItemDataRole.UserRole, None)
                text = item.text(0).replace('🔗 ', '').replace('📄 ', '')
                item.setText(0, text)
                item.setForeground(0, QColor(self.text_color))

    def paste_as_nodes(self):
        """Paste multi-line text and turn it into an outline node tree.

        Rules:
          • Blank lines separate top-level sibling groups.
          • Within each group the FIRST line becomes the parent node.
          • Every subsequent line in that group becomes a child node.
          • A group with only one line produces a childless node.
        """
        dialog = QDialog(self)
        dialog.setWindowTitle("Paste as Nodes")
        dialog.resize(640, 420)
        layout = QVBoxLayout(dialog)

        hint = QLabel(
            "Paste text below.\n"
            "Blank lines → new top-level node.  "
            "First line of each block → parent.  "
            "Remaining lines → children."
        )
        hint.setWordWrap(True)
        layout.addWidget(hint)

        text_area = QTextEdit()
        text_area.setPlaceholderText("Paste text here…")
        layout.addWidget(text_area)

        btn_row = QHBoxLayout()
        ok_btn = QPushButton("Create Nodes")
        ok_btn.setDefault(True)
        cancel_btn = QPushButton("Cancel")
        ok_btn.clicked.connect(dialog.accept)
        cancel_btn.clicked.connect(dialog.reject)
        btn_row.addWidget(ok_btn)
        btn_row.addWidget(cancel_btn)
        layout.addLayout(btn_row)

        if dialog.exec() != QDialog.DialogCode.Accepted:
            return

        raw = text_area.toPlainText()
        if not raw.strip():
            return

        # Split into blocks on one-or-more blank lines
        blocks = re.split(r'\n[ \t]*\n', raw)
        new_items = []
        for block in blocks:
            lines = [l.rstrip() for l in block.split('\n') if l.strip()]
            if not lines:
                continue
            parent_item = QTreeWidgetItem([lines[0]])
            parent_item.setFlags(parent_item.flags() | Qt.ItemFlag.ItemIsEditable)
            parent_item.setData(0, Qt.ItemDataRole.UserRole + 2, lines[0])
            parent_item.setData(0, Qt.ItemDataRole.UserRole + 8, True)
            parent_item.setData(0, Qt.ItemDataRole.UserRole + 9, False)
            for line in lines[1:]:
                child = QTreeWidgetItem(parent_item, [line])
                child.setFlags(child.flags() | Qt.ItemFlag.ItemIsEditable)
                child.setData(0, Qt.ItemDataRole.UserRole + 2, line)
                child.setData(0, Qt.ItemDataRole.UserRole + 8, True)
                child.setData(0, Qt.ItemDataRole.UserRole + 9, False)
            new_items.append(parent_item)

        if not new_items:
            return

        self.save_state()

        # Insert after the current item (or append if nothing selected)
        current = self.tree.currentItem()
        self._loading = True
        self.tree.setUpdatesEnabled(False)
        try:
            if current:
                par = current.parent()
                if par:
                    idx = par.indexOfChild(current)
                    for i, it in enumerate(new_items):
                        par.insertChild(idx + 1 + i, it)
                else:
                    idx = self.tree.indexOfTopLevelItem(current)
                    for i, it in enumerate(new_items):
                        self.tree.insertTopLevelItem(idx + 1 + i, it)
            else:
                for it in new_items:
                    self.tree.addTopLevelItem(it)
        finally:
            self.tree.setUpdatesEnabled(True)
            self._loading = False

        # Update display for all new items
        for it in new_items:
            self.update_node_display(it, it.childCount() > 0)
            for j in range(it.childCount()):
                ch = it.child(j)
                self.update_node_display(ch, False)

        self.tree.setCurrentItem(new_items[0])
        if new_items[0].childCount() > 0:
            new_items[0].setExpanded(True)
        self.is_dirty = True

    def refresh_include_node(self):
        """Re-fetch the selected include node from its source URL."""
        item = self.tree.currentItem()
        if not item:
            return

        url = item.data(0, Qt.ItemDataRole.UserRole + 1)
        if not url:
            url, ok = QInputDialog.getText(
                self, "Refresh Include",
                "Enter the URL or file path to re-fetch:"
            )
            if not ok or not url.strip():
                return
            url = url.strip()

        try:
            root = self._load_opml_root(url)
            body = root.find('.//body')
            if body is None:
                QMessageBox.warning(self, "Refresh Include", "No body found in OPML.")
                return
            # Remove stale children
            item.takeChildren()
            # Store URL for future saves
            item.setData(0, Qt.ItemDataRole.UserRole + 1, url)
            self._loading = True
            self.tree.setUpdatesEnabled(False)
            for outline in body.findall('outline'):
                self.add_outline_to_tree(outline, item, from_include=True)
            self.tree.setUpdatesEnabled(True)
            self._loading = False
            item.setExpanded(True)
            self.tree.scheduleDelayedItemsLayout()
            self.is_dirty = True
        except Exception as e:
            self.tree.setUpdatesEnabled(True)
            self._loading = False
            QMessageBox.critical(self, "Error", f"Failed to refresh include:\n{e}")

    def add_include_node(self):
        """Add a node that includes content from file or URL"""
        dialog = QDialog(self)
        dialog.setWindowTitle("Include from File/URL")
        dialog.resize(520, 120)
        layout = QVBoxLayout(dialog)

        row = QHBoxLayout()
        path_edit = QLineEdit()
        path_edit.setPlaceholderText("Enter file path or URL…")
        browse_btn = QPushButton("Browse…")

        def browse():
            start_dir = self.last_folder if Path(self.last_folder).is_dir() else str(Path.home())
            filename, _ = QFileDialog.getOpenFileName(
                dialog, "Select OPML File", start_dir,
                "OPML Files (*.opml);;All Files (*)"
            )
            if filename:
                path_edit.setText(filename)

        browse_btn.clicked.connect(browse)
        row.addWidget(path_edit)
        row.addWidget(browse_btn)
        layout.addLayout(row)

        btn_row = QHBoxLayout()
        ok_btn = QPushButton("Include")
        ok_btn.setDefault(True)
        cancel_btn = QPushButton("Cancel")
        ok_btn.clicked.connect(dialog.accept)
        cancel_btn.clicked.connect(dialog.reject)
        btn_row.addStretch()
        btn_row.addWidget(ok_btn)
        btn_row.addWidget(cancel_btn)
        layout.addLayout(btn_row)

        if dialog.exec() != QDialog.DialogCode.Accepted:
            return

        path = path_edit.text().strip()
        if path:
            try:
                root = self._load_opml_root(path.strip())
                body = root.find('.//body')

                if body is not None:
                    if path.strip().startswith('http'):
                        label = path.strip().rstrip('/').split('/')[-1] or path.strip()
                    else:
                        label = Path(path.strip()).name
                    include_root = QTreeWidgetItem([f"📄 Included: {label}"])
                    include_root.setFlags(include_root.flags() | Qt.ItemFlag.ItemIsEditable)
                    include_root.setData(0, Qt.ItemDataRole.UserRole + 1, path.strip())
                    include_root.setData(0, Qt.ItemDataRole.UserRole + 9, True)

                    self._loading = True
                    for outline in body.findall('outline'):
                        self.add_outline_to_tree(outline, include_root, from_include=True)
                    self._loading = False

                    current = self.tree.currentItem()
                    if current:
                        parent = current.parent()
                        if parent:
                            index = parent.indexOfChild(current)
                            parent.insertChild(index + 1, include_root)
                        else:
                            index = self.tree.indexOfTopLevelItem(current)
                            self.tree.insertTopLevelItem(index + 1, include_root)
                    else:
                        self.tree.addTopLevelItem(include_root)

                    include_root.setExpanded(True)
                    self.tree.setCurrentItem(include_root)
                    self.is_dirty = True
            except Exception as e:
                self.tree.setUpdatesEnabled(True)
                QMessageBox.critical(self, "Error", f"Failed to include:\n{e}")
    
    def import_from_url(self):
        """Import an OPML file from a URL"""
        url, ok = QInputDialog.getText(
            self, "Import from URL",
            "Enter URL to OPML file:"
        )
        
        if ok and url.strip():
            try:
                # Download the OPML file
                req = urllib.request.Request(url.strip(), headers={'User-Agent': 'Mozilla/5.0'})
                with urllib.request.urlopen(req, timeout=30) as response:
                    data = response.read()
                
                # Check if content is gzip compressed (magic bytes 0x1f 0x8b)
                if len(data) >= 2 and data[0] == 0x1f and data[1] == 0x8b:
                    # Decompress gzip
                    content = gzip.decompress(data).decode('utf-8')
                else:
                    content = data.decode('utf-8')
                
                # Parse the OPML
                root = ET.fromstring(content)
                body = root.find('.//body')
                
                if body is not None:
                    # Create a root item for the imported content
                    import_root = QTreeWidgetItem([f"📥 Imported: {Path(url).name}"])
                    import_root.setFlags(import_root.flags() | Qt.ItemFlag.ItemIsEditable)
                    import_root.setData(0, Qt.ItemDataRole.UserRole + 3, url.strip())  # Store source URL
                    import_root.setData(0, Qt.ItemDataRole.UserRole + 9, True)  # Force plain rendering for imports
                    
                    self.tree.setUpdatesEnabled(False)
                    for outline in body.findall('outline'):
                        self.add_outline_to_tree(outline, import_root, from_include=True)
                    self.tree.setUpdatesEnabled(True)
                    
                    # Add to tree
                    self.tree.addTopLevelItem(import_root)
                    import_root.setExpanded(True)
                    self.tree.setCurrentItem(import_root)
                    self.is_dirty = True
                    
                    QMessageBox.information(self, "Success", "OPML imported successfully!")
                else:
                    QMessageBox.warning(self, "Error", "No body content found in OPML")
            except Exception as e:
                self.tree.setUpdatesEnabled(True)
                QMessageBox.critical(self, "Error", f"Failed to import:\n{e}")

    def setup_shortcuts(self):
        # Ctrl+Enter: Open link or toggle expand
        shortcut = QShortcut(QKeySequence("Ctrl+Return"), self)
        shortcut.activated.connect(self.handle_ctrl_enter)
        
        # Ctrl+B: Bold (when editing)
        shortcut = QShortcut(QKeySequence("Ctrl+B"), self)
        shortcut.activated.connect(self.format_bold)
        
        # Ctrl+I: Italic (when editing)
        shortcut = QShortcut(QKeySequence("Ctrl+I"), self)
        shortcut.activated.connect(self.format_italic)
        
        # Ctrl+d: Move node down
        shortcut = QShortcut(QKeySequence("Ctrl+d"), self)
        shortcut.activated.connect(self.move_node_down)
        
        # Ctrl+u: Move node up
        shortcut = QShortcut(QKeySequence("Ctrl+u"), self)
        shortcut.activated.connect(self.move_node_up)
        
        # Ctrl+Up: Move node up
        shortcut = QShortcut(QKeySequence("Ctrl+Up"), self)
        shortcut.activated.connect(self.move_node_up)
        
        # Ctrl+Down: Move node down
        shortcut = QShortcut(QKeySequence("Ctrl+Down"), self)
        shortcut.activated.connect(self.move_node_down)
        
        # Tab: Indent
        shortcut = QShortcut(QKeySequence(Qt.Key.Key_Tab), self)
        shortcut.activated.connect(self.indent_node)
        
        # Shift+Tab: Outdent
        shortcut = QShortcut(QKeySequence("Shift+Tab"), self)
        shortcut.activated.connect(self.outdent_node)
        
        # Left arrow: Collapse
        shortcut = QShortcut(QKeySequence(Qt.Key.Key_Left), self)
        shortcut.activated.connect(self.collapse_node)
        
        # Right arrow: Expand
        shortcut = QShortcut(QKeySequence(Qt.Key.Key_Right), self)
        shortcut.activated.connect(self.expand_node)
        
        # Enter: Create new sibling
        shortcut = QShortcut(QKeySequence(Qt.Key.Key_Return), self)
        shortcut.activated.connect(self.add_sibling_node)
        
        # Escape: Edit current node (nav mode -> edit mode)
        shortcut = QShortcut(QKeySequence(Qt.Key.Key_Escape), self)
        shortcut.activated.connect(self.edit_current_node)
        
        # Ctrl+C: Copy node
        shortcut = QShortcut(QKeySequence.StandardKey.Copy, self)
        shortcut.activated.connect(self.copy_node)
        
        # Ctrl+X: Cut node
        shortcut = QShortcut(QKeySequence.StandardKey.Cut, self)
        shortcut.activated.connect(self.cut_node)
        
        # Ctrl+V: Paste node
        shortcut = QShortcut(QKeySequence.StandardKey.Paste, self)
        shortcut.activated.connect(self.paste_node)
        
        # Delete: Delete node
        shortcut = QShortcut(QKeySequence(Qt.Key.Key_Delete), self)
        shortcut.activated.connect(self.delete_node)
        
        # Ctrl+Z: Undo
        shortcut = QShortcut(QKeySequence.StandardKey.Undo, self)
        shortcut.activated.connect(self.undo)

        # Ctrl+Y: Redo
        shortcut = QShortcut(QKeySequence.StandardKey.Redo, self)
        shortcut.activated.connect(self.redo)
    
    def resizeEvent(self, event):
        super().resizeEvent(event)
        # Re-query sizeHint for all items so row heights adapt to new width
        self.tree.scheduleDelayedItemsLayout()
        # Re-elide the file path label at the new width
        tooltip = self.file_path_label.toolTip()
        if tooltip:
            self._set_file_path_label(tooltip)

    def closeEvent(self, event):
        """Handle window close - check for unsaved changes"""
        if self.is_dirty:
            reply = QMessageBox.question(
                self, 'Unsaved Changes',
                'You have unsaved changes. Do you want to save before closing?',
                QMessageBox.StandardButton.Save | 
                QMessageBox.StandardButton.Discard | 
                QMessageBox.StandardButton.Cancel
            )
            if reply == QMessageBox.StandardButton.Save:
                self.save_file()
                event.accept()
            elif reply == QMessageBox.StandardButton.Discard:
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()
    
    def save_state(self):
        """Save current tree state for undo"""
        # Use a list as a stack for multiple undo levels
        if not hasattr(self, '_undo_stack'):
            self._undo_stack = []
        if not hasattr(self, '_redo_stack'):
            self._redo_stack = []
        
        # Save current state
        state_data = {
            'tree': self.tree_to_opml_string(),
            'expanded': self.get_expanded_items()
        }
        self._undo_stack.append(state_data)
        # Limit stack size
        if len(self._undo_stack) > 50:
            self._undo_stack.pop(0)
        # Clear redo stack on new action
        self._redo_stack = []
    
    def get_expanded_items(self):
        """Get list of expanded item texts"""
        expanded = []
        def get_expanded_recursive(item, path):
            if item.isExpanded():
                path_str = '/'.join(path + [item.text(0)[:30]])
                expanded.append(path_str)
            for i in range(item.childCount()):
                get_expanded_recursive(item.child(i), path + [item.text(0)[:30]])
        for i in range(self.tree.topLevelItemCount()):
            get_expanded_recursive(self.tree.topLevelItem(i), [])
        return expanded
    
    def restore_state(self, state_data):
        """Restore tree from saved state"""
        if not state_data:
            return
            
        state = state_data.get('tree', '')
        expanded_paths = state_data.get('expanded', [])
        
        if state:
            # Save current expanded state before clearing
            current_expanded = self.get_expanded_items()
            
            self.tree.clear()
            root = ET.fromstring(state)
            body = root.find('.//body')
            if body is not None:
                for outline in body.findall('outline'):
                    self.add_outline_to_tree(outline, None, from_include=False)

            # First try to restore from saved paths
            self.restore_expanded_items(expanded_paths)
            
            # If no expanded paths saved, try to restore current state
            if not expanded_paths and current_expanded:
                self.restore_expanded_items(current_expanded)
    
    def restore_expanded_items(self, expanded_paths):
        """Restore expanded state based on paths"""
        if not expanded_paths:
            return
            
        # Build a map of item text to item
        text_to_items = {}
        def build_map(item, parent_path):
            key = parent_path + item.text(0)[:50]
            if key not in text_to_items:
                text_to_items[key] = item
            for i in range(item.childCount()):
                build_map(item.child(i), key + '/')
        
        for i in range(self.tree.topLevelItemCount()):
            build_map(self.tree.topLevelItem(i), '')
        
        # Now expand items
        for path in expanded_paths:
            parts = path.split('/')
            if not parts:
                continue
            # Find the top-level item
            key = parts[0][:50]
            for i in range(self.tree.topLevelItemCount()):
                item = self.tree.topLevelItem(i)
                if item.text(0).startswith(parts[0][:50]):
                    item.setExpanded(True)
                    # Navigate down the path
                    current = item
                    for j, part in enumerate(parts[1:], 1):
                        found = False
                        for k in range(current.childCount()):
                            child = current.child(k)
                            if child.text(0).startswith(part[:50]):
                                child.setExpanded(True)
                                current = child
                                found = True
                                break
                        if not found:
                            break
    
    def tree_to_opml_string(self):
        """Convert tree to OPML string for state saving"""
        root = ET.Element('opml', version="2.0")
        head = ET.SubElement(root, 'head')
        title = ET.SubElement(head, 'title')
        title.text = "saved"
        body = ET.SubElement(root, 'body')
        for i in range(self.tree.topLevelItemCount()):
            item = self.tree.topLevelItem(i)
            self.item_to_outline(item, body)
        return ET.tostring(root, encoding='unicode')
    
    def undo(self):
        if hasattr(self, '_undo_stack') and len(self._undo_stack) > 0:
            # Save current state to redo stack
            current_state = {
                'tree': self.tree_to_opml_string(),
                'expanded': self.get_expanded_items()
            }
            self._redo_stack.append(current_state)
            
            # Restore previous state
            state_data = self._undo_stack.pop()
            self.restore_state(state_data)
            self.is_dirty = True
    
    def redo(self):
        if hasattr(self, '_redo_stack') and len(self._redo_stack) > 0:
            # Save current state to undo stack
            current_state = {
                'tree': self.tree_to_opml_string(),
                'expanded': self.get_expanded_items()
            }
            self._undo_stack.append(current_state)
            
            # Restore redo state
            state_data = self._redo_stack.pop()
            self.restore_state(state_data)
            self.is_dirty = True

    def update_node_display(self, item, has_children):
        """Update node text with appropriate triangle indicator and wrapped text"""
        original_text = item.data(0, Qt.ItemDataRole.UserRole + 2) or item.text(0).lstrip('▶▼ ')
        # Strip any icon prefix that may have been inadvertently stored (e.g. from a
        # previous buggy save or from _on_item_changed reading a dirty DisplayRole).
        for _pfx in ('📄 ', '🔗 '):
            if original_text.startswith(_pfx):
                original_text = original_text[len(_pfx):]
                break
        url = item.data(0, Qt.ItemDataRole.UserRole)
        force_plain = bool(item.data(0, Qt.ItemDataRole.UserRole + 9))

        # Check if it's an OPML link
        is_opml_link = bool(url and (url.lower().endswith('.opml') or 'opml' in url.lower()))

        display_text = original_text
        user_edited = bool(item.data(0, Qt.ItemDataRole.UserRole + 8))

        if force_plain or not user_edited:
            # Strip all HTML except <img> tags so images render inline
            display_text = re.sub(r'<(?!/?img\b)[^>]*>', '', display_text, flags=re.IGNORECASE)
            is_html = bool(re.search(r'<img\b', display_text, re.IGNORECASE))
        else:
            # Only user-authored text can keep HTML rendering
            is_html = bool(re.search(r'<[a-zA-Z]', display_text))

        item.setData(0, Qt.ItemDataRole.UserRole + 3, is_html)

        wrapped_text = display_text  # delegate handles word-wrap via TextWordWrap

        def set_display(text, color=None):
            item.setData(0, Qt.ItemDataRole.DisplayRole, text)
            if color:
                item.setForeground(0, QColor(color))
            else:
                item.setForeground(0, QColor(self.text_color))

        # Add icons for links (no triangles)
        if is_opml_link:
            set_display(f"📄 {wrapped_text}", "#008800")
        elif url:
            set_display(f"🔗 {wrapped_text}", "#0066cc")
        else:
            set_display(wrapped_text, self.text_color)

    def on_item_changed(self, item, column):
        """Handle text edits - update stored original text"""
        # Ignore signals while loading or while we're already updating display
        if getattr(self, '_loading', False) or getattr(self, '_in_item_changed', False):
            return
        self._in_item_changed = True
        try:
            self._on_item_changed_inner(item, column)
        finally:
            self._in_item_changed = False

    def _on_item_changed_inner(self, item, column):
        if column == 0:
            current_text = item.text(0)
            # Replace newlines with spaces (but keep HTML tags intact)
            original_text = current_text.replace('\n', ' ')
            
            # Check if text actually changed - save undo state if so
            if hasattr(self, '_editing_item') and self._editing_item is item:
                prev_text = getattr(self, '_editing_prev_text', '')
                if prev_text != original_text and prev_text != "":
                    # Text changed - save undo state
                    self.save_state()
            
            # Store the original text (with HTML tags preserved)
            item.setData(0, Qt.ItemDataRole.UserRole + 2, original_text)
            item.setData(0, Qt.ItemDataRole.UserRole + 8, True)  # mark user edited
            # Update display with proper wrapping and triangle
            has_children = item.childCount() > 0
            self.update_node_display(item, has_children)
            # Mark as dirty (but not during initial load)
            if hasattr(self, 'current_file') and self.current_file:
                self.is_dirty = True
            
            # Clear editing state
            self._editing_item = None

    def collapse_node(self):
        """Collapse current node"""
        item = self.tree.currentItem()
        if item:
            if item.isExpanded() and item.childCount() > 0:
                # If expanded with children, collapse it
                item.setExpanded(False)
            else:
                # If already collapsed or no children, go to parent
                parent = item.parent()
                if parent:
                    self.tree.setCurrentItem(parent)

    def expand_node(self):
        """Expand current node, auto-including OPML URLs when needed"""
        item = self.tree.currentItem()
        if not item:
            return
        url = item.data(0, Qt.ItemDataRole.UserRole)
        if url and item.childCount() == 0:
            url_lower = url.lower()
            if url_lower.endswith('.opml') or 'opml' in url_lower:
                self.include_opml_file(url, item)
                return
        if item.childCount() > 0:
            if not item.isExpanded():
                item.setExpanded(True)
            else:
                self.tree.setCurrentItem(item.child(0))

    def format_bold(self):
        editor = self.tree.findChild(FormattingTextEdit)
        if editor:
            editor.toggle_format('b')
        else:
            self._toggle_format_on_item('b')

    def format_italic(self):
        editor = self.tree.findChild(FormattingTextEdit)
        if editor:
            editor.toggle_format('i')
        else:
            self._toggle_format_on_item('i')

    def _toggle_format_on_item(self, tag):
        item = self.tree.currentItem()
        if not item:
            return
        text = item.data(0, Qt.ItemDataRole.UserRole + 2) or item.text(0)
        open_tag, close_tag = f'<{tag}>', f'</{tag}>'
        if text.startswith(open_tag) and text.endswith(close_tag):
            text = text[len(open_tag):-len(close_tag)]
        else:
            text = f'{open_tag}{text}{close_tag}'
        item.setData(0, Qt.ItemDataRole.UserRole + 2, text)
        item.setData(0, Qt.ItemDataRole.UserRole + 8, True)
        self.update_node_display(item, item.childCount() > 0)
        self.is_dirty = True
    
    def handle_ctrl_enter(self):
        """Open link, or add link if none exists, or toggle expand/collapse"""
        item = self.tree.currentItem()
        if not item:
            return
        
        url = item.data(0, Qt.ItemDataRole.UserRole)
        if url:
            # Check if it's an OPML file
            url_lower = url.lower()
            is_opml = url_lower.endswith('.opml') or 'opml' in url_lower
            
            if is_opml:
                # It's an OPML file - include it inline
                self.include_opml_file(url, item)
            else:
                # Has link - open it
                webbrowser.open(url)
        elif item.childCount() > 0:
            # No link but has children - toggle expand
            item.setExpanded(not item.isExpanded())
        else:
            # No link, no children - offer to add link
            self.add_link_to_node()
    
    def include_opml_file(self, path_or_url, parent_item=None):
        """Include an OPML file inline"""
        try:
            root = self._load_opml_root(path_or_url)
            
            body = root.find('.//body')
            if body is not None:
                # Save undo state
                self.save_state()
                
                # Create include node
                if parent_item is None:
                    include_root = QTreeWidgetItem([f"📄 Included: {Path(path_or_url).name}"])
                    include_root.setFlags(include_root.flags() | Qt.ItemFlag.ItemIsEditable)
                    include_root.setData(0, Qt.ItemDataRole.UserRole + 1, path_or_url)
                    include_root.setData(0, Qt.ItemDataRole.UserRole + 9, True)  # Force plain
                    self.tree.addTopLevelItem(include_root)
                    target = include_root
                else:
                    target = parent_item
                
                # Batch add
                self._loading = True
                try:
                    for outline in body.findall('outline'):
                        self.add_outline_to_tree(outline, target, from_include=True)
                finally:
                    self._loading = False
                # Defer expand so Qt processes the newly added rows first
                QTimer.singleShot(0, lambda t=target: (t.setExpanded(True), self.tree.setCurrentItem(t)))
                self.is_dirty = True
        except Exception as e:
            self._loading = False
            QMessageBox.critical(self, "Error", f"Failed to include OPML:\n{e}")

    def _load_opml_root(self, path_or_url):
        """Load an OPML root element from file or URL with gzip + encoding fallbacks."""
        if path_or_url.startswith('http'):
            req = urllib.request.Request(path_or_url, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req, timeout=30) as response:
                data = response.read()
            if len(data) >= 2 and data[0] == 0x1f and data[1] == 0x8b:
                data = gzip.decompress(data)
            content = data.decode('utf-8', errors='replace')
            return ET.fromstring(content)
        else:
            tree = ET.parse(path_or_url)
            return tree.getroot()

    def move_node_down(self):
        item = self.tree.currentItem()
        if not item:
            return
        
        self.save_state()
        
        parent = item.parent()
        if parent:
            index = parent.indexOfChild(item)
            if index < parent.childCount() - 1:
                parent.takeChild(index)
                parent.insertChild(index + 1, item)
                self.tree.setCurrentItem(item)
                self.tree.scrollToItem(item)
                self.is_dirty = True
        else:
            index = self.tree.indexOfTopLevelItem(item)
            if index < self.tree.topLevelItemCount() - 1:
                self.tree.takeTopLevelItem(index)
                self.tree.insertTopLevelItem(index + 1, item)
                self.tree.setCurrentItem(item)
                self.tree.scrollToItem(item)
                self.is_dirty = True

    def move_node_up(self):
        item = self.tree.currentItem()
        if not item:
            return
        
        self.save_state()
        
        parent = item.parent()
        if parent:
            index = parent.indexOfChild(item)
            if index > 0:
                parent.takeChild(index)
                parent.insertChild(index - 1, item)
                self.tree.setCurrentItem(item)
                self.tree.scrollToItem(item)
                self.is_dirty = True
        else:
            index = self.tree.indexOfTopLevelItem(item)
            if index > 0:
                self.tree.takeTopLevelItem(index)
                self.tree.insertTopLevelItem(index - 1, item)
                self.tree.setCurrentItem(item)
                self.tree.scrollToItem(item)
                self.is_dirty = True

    def indent_node(self):
        item = self.tree.currentItem()
        if not item:
            return
        
        self.save_state()
        
        parent = item.parent()
        if parent:
            index = parent.indexOfChild(item)
            if index > 0:
                prev_sibling = parent.child(index - 1)
                parent.takeChild(index)
                prev_sibling.addChild(item)
                prev_sibling.setExpanded(True)
                self.tree.setCurrentItem(item)
                self.is_dirty = True
        else:
            index = self.tree.indexOfTopLevelItem(item)
            if index > 0:
                prev_sibling = self.tree.topLevelItem(index - 1)
                self.tree.takeTopLevelItem(index)
                prev_sibling.addChild(item)
                prev_sibling.setExpanded(True)
                self.tree.setCurrentItem(item)
                self.is_dirty = True

    def outdent_node(self):
        item = self.tree.currentItem()
        if not item:
            return
        
        self.save_state()
        
        parent = item.parent()
        if not parent:
            return
        
        grandparent = parent.parent()
        parent_index = parent.indexOfChild(item)
        parent.takeChild(parent_index)
        
        if grandparent:
            gp_index = grandparent.indexOfChild(parent)
            grandparent.insertChild(gp_index + 1, item)
        else:
            p_index = self.tree.indexOfTopLevelItem(parent)
            self.tree.insertTopLevelItem(p_index + 1, item)
        
        self.tree.setCurrentItem(item)
        self.is_dirty = True

    def edit_current_node(self):
        """Enter edit mode on current node with cursor at end"""
        # Check if already editing - if so, do nothing (let Qt handle Escape)
        editor = self.tree.findChild(QLineEdit)
        if editor:
            return  # Already editing, let default behavior handle Escape
        
        item = self.tree.currentItem()
        if item:
            # Save previous text for undo
            prev_text = item.data(0, Qt.ItemDataRole.UserRole + 2) or item.text(0)
            self._editing_item = item
            self._editing_prev_text = prev_text
            
            self.tree.scrollToItem(item)
            self.tree.editItem(item, 0)
            # Move cursor to end after edit starts
            QTimer.singleShot(0, self._move_cursor_to_end)
    
    def _move_cursor_to_end(self):
        """Move text cursor to end in the current editor"""
        from PyQt6.QtWidgets import QTextEdit
        editor = self.tree.findChild(QTextEdit)
        if editor:
            cursor = editor.textCursor()
            cursor.movePosition(cursor.MoveOperation.End)
            editor.setTextCursor(cursor)

    def add_sibling_node(self):
        self.save_state()
        
        item = self.tree.currentItem()
        new_item = QTreeWidgetItem([""])
        new_item.setFlags(new_item.flags() | Qt.ItemFlag.ItemIsEditable)
        new_item.setData(0, Qt.ItemDataRole.UserRole + 2, "")  # Store original text
        
        if item:
            parent = item.parent()
            if parent:
                index = parent.indexOfChild(item)
                parent.insertChild(index + 1, new_item)
                # Update parent's triangle if it's the first child
                if parent.childCount() == 1:
                    self.update_node_display(parent, True)
            else:
                index = self.tree.indexOfTopLevelItem(item)
                self.tree.insertTopLevelItem(index + 1, new_item)
        else:
            self.tree.addTopLevelItem(new_item)
        
        self.tree.setCurrentItem(new_item)
        self.tree.scrollToItem(new_item)
        QTimer.singleShot(0, lambda: self.tree.editItem(new_item, 0))
        self.is_dirty = True

    def copy_item_recursive(self, item):
        """Deep copy a tree item with all children"""
        new_item = QTreeWidgetItem([item.text(0)])
        new_item.setFlags(item.flags())
        
        # Copy data
        url = item.data(0, Qt.ItemDataRole.UserRole)
        if url:
            new_item.setData(0, Qt.ItemDataRole.UserRole, url)
            new_item.setForeground(0, item.foreground(0))
        
        # Copy children
        for i in range(item.childCount()):
            child_copy = self.copy_item_recursive(item.child(i))
            new_item.addChild(child_copy)
        
        return new_item

    def copy_node(self):
        """Copy current node and its children"""
        item = self.tree.currentItem()
        if item:
            self.clipboard_item = self.copy_item_recursive(item)
            self.mode_label.setText("📋 Node copied - Press Ctrl+V to paste")
            self.mode_label.setStyleSheet("color: #00aa00; padding: 4px 8px;")
            QTimer.singleShot(2000, lambda: self.mode_label.setText("Use ← → to collapse/expand, ↑↓ to navigate"))

    def cut_node(self):
        """Cut current node"""
        item = self.tree.currentItem()
        if item:
            self.save_state()
            self.clipboard_item = self.copy_item_recursive(item)
            
            # Delete the original
            parent = item.parent()
            if parent:
                parent.removeChild(item)
            else:
                index = self.tree.indexOfTopLevelItem(item)
                self.tree.takeTopLevelItem(index)
            
            self.mode_label.setText("✂️ Node cut - Press Ctrl+V to paste")
            self.mode_label.setStyleSheet("color: #ff6600; padding: 4px 8px;")
            QTimer.singleShot(2000, lambda: self.mode_label.setText("Use ← → to collapse/expand, ↑↓ to navigate"))
            self.is_dirty = True

    def paste_node(self):
        """Paste node as sibling of current"""
        if not self.clipboard_item:
            return
        
        self.save_state()
        
        item = self.tree.currentItem()
        pasted_item = self.copy_item_recursive(self.clipboard_item)
        
        if item:
            parent = item.parent()
            if parent:
                index = parent.indexOfChild(item)
                parent.insertChild(index + 1, pasted_item)
            else:
                index = self.tree.indexOfTopLevelItem(item)
                self.tree.insertTopLevelItem(index + 1, pasted_item)
        else:
            self.tree.addTopLevelItem(pasted_item)
        
        self.tree.setCurrentItem(pasted_item)
        self.mode_label.setText("✅ Node pasted")
        self.is_dirty = True
        self.mode_label.setStyleSheet("color: #00aa00; padding: 4px 8px;")
        QTimer.singleShot(2000, lambda: self.mode_label.setText("Use ← → to collapse/expand, ↑↓ to navigate"))

    def delete_node(self):
        """Delete current node"""
        item = self.tree.currentItem()
        if not item:
            return
        
        # Confirm if node has children
        if item.childCount() > 0:
            reply = QMessageBox.question(
                self, 'Confirm Delete',
                f'Delete this node and its {item.childCount()} children?',
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply != QMessageBox.StandardButton.Yes:
                return
        
        self.save_state()
        
        # Delete the node
        parent = item.parent()
        if parent:
            parent.removeChild(item)
        else:
            index = self.tree.indexOfTopLevelItem(item)
            self.tree.takeTopLevelItem(index)
        
        self.is_dirty = True

    def wrap_text(self, text, width=80):
        """Wrap text to specified width (skip HTML)"""
        if '<' in text and '>' in text:
            return text
        if not text:
            return ''

        # Avoid Python textwrap recursion on extremely long unbreakable strings
        try:
            tokens = text.split()
            if tokens and max(len(t) for t in tokens) > width * 4:
                return textwrap.fill(text, width=width, break_long_words=True, break_on_hyphens=False)

            wrapped = textwrap.wrap(text, width=width, break_long_words=False, break_on_hyphens=False)
            return '\n'.join(wrapped) if wrapped else text
        except RecursionError:
            return text

    def _set_file_path_label(self, text):
        """Set the file path label, storing full path as tooltip and eliding on display."""
        self.file_path_label.setToolTip(text)
        fm = self.file_path_label.fontMetrics()
        available = self.file_path_label.width() - 32  # account for padding
        if available < 20:
            available = 600
        elided = fm.elidedText(text, Qt.TextElideMode.ElideLeft, available)
        self.file_path_label.setText(elided)

    def new_file(self):
        self.tree.clear()
        self.current_file = None
        self.is_dirty = False
        self.setWindowTitle("OPML Outliner - New File")
        self._set_file_path_label("New File")
        
        # Disable updates during item creation
        self.tree.setUpdatesEnabled(False)
        first_item = QTreeWidgetItem(["Start typing..."])
        first_item.setFlags(first_item.flags() | Qt.ItemFlag.ItemIsEditable)
        first_item.setData(0, Qt.ItemDataRole.UserRole + 9, True)  # Force plain
        first_item.setData(0, Qt.ItemDataRole.UserRole + 3, False)  # Not HTML
        self.tree.addTopLevelItem(first_item)
        
        self.tree.setUpdatesEnabled(True)
        self.tree.setCurrentItem(first_item)
        self.tree.setFocus()

    def open_file(self):
        start_dir = self.last_folder if Path(self.last_folder).is_dir() else str(Path.home())
        filename, _ = QFileDialog.getOpenFileName(
            self, "Open File", start_dir,
            "OPML Files (*.opml);;All Files (*)"
        )
        if filename:
            self.last_folder = str(Path(filename).parent)
            self.load_opml(filename)
            self.current_file = filename
            self.setWindowTitle(f"OPML Outliner - {Path(filename).name}")
            self._add_recent_file(filename)

    def save_file(self):
        if not self.current_file:
            self.save_file_as()
        else:
            self.save_opml(self.current_file)
            self.is_dirty = False

    def save_file_as(self):
        start_dir = self.last_folder if Path(self.last_folder).is_dir() else str(Path.home())
        filename, _ = QFileDialog.getSaveFileName(
            self, "Save File", start_dir,
            "OPML Files (*.opml);;All Files (*)"
        )
        if filename:
            self.last_folder = str(Path(filename).parent)
            self.save_opml(filename)
            self.current_file = filename
            self.last_file = filename
            self.setWindowTitle(f"OPML Outliner - {Path(filename).name}")
            self._set_file_path_label(filename)
            self._add_recent_file(filename)
            self.is_dirty = False

    def load_opml(self, filename):
        try:
            self._loading = True
            tree = ET.parse(filename)
            root = tree.getroot()
            self.tree.clear()
            
            self.tree.setUpdatesEnabled(False)
            body = root.find('.//body')
            if body is not None:
                for outline in body.findall('outline'):
                    self.add_outline_to_tree(outline, None, from_include=False)
            
            self.tree.setUpdatesEnabled(True)

            if self.tree.topLevelItemCount() > 0:
                first_item = self.tree.topLevelItem(0)
                self.tree.setCurrentItem(first_item)
                self.tree.setFocus()
            
            # Update file tracking
            self.current_file = filename
            self.last_file = filename
            self._set_file_path_label(filename)
            self.setWindowTitle(f"OPML Outliner - {Path(filename).name}")
            self.save_preferences()
            self.is_dirty = False
            self._refresh_pending_includes()
        except Exception as e:
            self.tree.setUpdatesEnabled(True)
            QMessageBox.critical(self, "Error", f"Failed to load file:\n{e}")
        finally:
            self._loading = False

    def add_outline_to_tree(self, outline_elem, parent_item, from_include=False):
        """Iterative (non-recursive) tree builder to avoid stack overflow on deep OPML."""
        # Stack entries: (elem, parent_widget_item)
        stack = [(outline_elem, parent_item)]
        # For normal-render nodes we must call update_node_display after children
        # are added, so track them in post-order: (item, has_children)
        normal_items = []  # [(item, elem)] to update after the full subtree is built

        while stack:
            elem, par = stack.pop()
            raw_text = elem.get('text', '')
            text = html.unescape(raw_text)
            url = elem.get('url')

            # Extract URLs from anchor tags when no explicit url attribute
            extracted_urls = re.findall(r'<a href="([^"]+)"[^>]*>([^<]*)</a>', raw_text)
            if extracted_urls and not url:
                url = extracted_urls[0][0]
                text = raw_text
                for href, link_text in extracted_urls:
                    text = re.sub(r'<a href="[^"]*"[^>]*>[^<]*</a>', link_text, text, count=1)
                text = html.unescape(text).strip()

            if from_include:
                item = QTreeWidgetItem(par if par else self.tree, [text])
                item.setFlags(item.flags() | Qt.ItemFlag.ItemIsEditable)
                item.setData(0, Qt.ItemDataRole.UserRole + 9, True)
                item.setData(0, Qt.ItemDataRole.UserRole + 8, False)
                item.setData(0, Qt.ItemDataRole.UserRole + 2, text)
                if url:
                    item.setData(0, Qt.ItemDataRole.UserRole, url)
                normal_items.append((item, elem))
            else:
                item = QTreeWidgetItem(par if par else self.tree, [text])
                item.setFlags(item.flags() | Qt.ItemFlag.ItemIsEditable)
                item.setData(0, Qt.ItemDataRole.UserRole + 9, False)
                item.setData(0, Qt.ItemDataRole.UserRole + 8, True)
                item.setData(0, Qt.ItemDataRole.UserRole + 2, text)
                if url:
                    item.setData(0, Qt.ItemDataRole.UserRole, url)
                xml_url = elem.get('xmlUrl')
                if xml_url:
                    item.setData(0, Qt.ItemDataRole.UserRole + 1, xml_url)
                    # Queue a fresh re-fetch; skip the stale saved children
                    if not hasattr(self, '_pending_include_reloads'):
                        self._pending_include_reloads = []
                    self._pending_include_reloads.append((item, xml_url))
                    normal_items.append((item, elem))
                    continue  # do NOT push saved children — fresh load will replace them
                normal_items.append((item, elem))

            # Push children in reverse order so left-to-right order is preserved
            children = elem.findall('outline')
            for child in reversed(children):
                stack.append((child, item))

        # Update display for normal-render items (needs children already attached)
        for item, elem in normal_items:
            self.update_node_display(item, item.childCount() > 0)

    def _item_to_html_lines(self, item, lines, line_counter):
        """Recursively render one tree item into the outline HTML format."""
        text = item.data(0, Qt.ItemDataRole.UserRole + 2) or item.text(0)
        for pfx in ('📄 ', '🔗 '):
            if text.startswith(pfx):
                text = text[len(pfx):]
                break
        if text.startswith('Included: '):
            text = text[len('Included: '):]

        url = item.data(0, Qt.ItemDataRole.UserRole)
        child_count = item.childCount()
        has_img = bool(re.search(r'<img\b', text, re.IGNORECASE))
        has_anchor = bool(url) or bool(re.search(r'<a\b', text, re.IGNORECASE))
        is_expanded = item.isExpanded()

        if child_count > 0:
            line_counter[0] += 1
            collapsed_cls = '' if is_expanded else ' collapsed'
            lines.append(f'<ul class="outline">')
            lines.append(f'<li data-line="{line_counter[0]}" class="owedge{collapsed_cls}">')
            lines.append(f'<span class="tog"></span>')
            if url:
                lines.append(f'<span class="lbl"><a href="{html.escape(url)}" target="_blank">{text}</a></span>')
            else:
                lines.append(f'<span class="lbl">{text}</span>')
            lines.append('<div class="children">')
            for i in range(child_count):
                self._item_to_html_lines(item.child(i), lines, line_counter)
            lines.append('</div>')
            lines.append('</li>')
            lines.append('</ul>')
        else:
            ul_cls = 'outline'
            li_cls = 'ou'
            if url:
                ul_cls += ' link wanchor'
                li_cls += ' link wanchor'
            elif has_anchor:
                ul_cls += ' wanchor'
                li_cls += ' wanchor'
            if has_img:
                ul_cls += ' wimg'
                li_cls += ' wimg'
            if url:
                content = f'<a href="{html.escape(url)}" target="_blank">{text}</a>'
            else:
                content = text
            lines.append(f'<ul class="{ul_cls}">')
            lines.append(f'<li class="{li_cls}">{content}</li>')
            lines.append('</ul>')

    _HTML_PAGE = '''\
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{title}</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;
     font-size:14px;line-height:1.6;color:#222;background:#fff;padding:20px 28px}}
ul.outline{{list-style:none;padding:0;margin:1px 0}}
li.ou{{padding:1px 0 1px 22px;}}
li.owedge{{padding:1px 0 1px 22px;position:relative}}
li.owedge>.tog{{
  position:absolute;left:2px;top:4px;
  width:14px;height:14px;cursor:pointer;
  display:inline-flex;align-items:center;justify-content:center;
  font-size:10px;color:#666;user-select:none;
  transition:transform .15s ease}}
li.owedge>.tog::before{{content:"▾"}}
li.owedge.collapsed>.tog{{transform:rotate(-90deg)}}
li.owedge>.lbl{{cursor:pointer;font-weight:500}}
li.owedge>.lbl:hover{{color:#0055cc}}
li.owedge>.children{{
  overflow:hidden;
  max-height:20000px;
  transition:max-height .25s ease-in, opacity .2s ease-in;
  opacity:1}}
li.owedge.collapsed>.children{{
  max-height:0;
  opacity:0;
  transition:max-height .18s ease-out, opacity .15s ease-out}}
a{{color:#0066cc;text-decoration:none}}
a:hover{{text-decoration:underline}}
img{{max-width:100%;height:auto;display:block;margin:4px 0}}
audio{{vertical-align:middle}}
</style>
</head>
<body>
{body}
<script>
(function(){{
  document.querySelectorAll("li.owedge").forEach(function(li){{
    ["tog","lbl"].forEach(function(cls){{
      var el=li.querySelector(":scope > ."+cls);
      if(el) el.addEventListener("click",function(e){{
        e.stopPropagation();
        li.classList.toggle("collapsed");
      }});
    }});
  }});
}})();
</script>
</body>
</html>'''

    def export_html(self):
        """Generate a self-contained interactive HTML page and show in a copyable dialog."""
        body_lines = []
        line_counter = [0]
        for i in range(self.tree.topLevelItemCount()):
            self._item_to_html_lines(self.tree.topLevelItem(i), body_lines, line_counter)

        title = (self.current_file and Path(self.current_file).stem) or 'Outline'
        full_page = self._HTML_PAGE.format(title=html.escape(title), body='\n'.join(body_lines))
        fragment = '\n'.join(body_lines)

        dialog = QDialog(self)
        dialog.setWindowTitle("Export HTML")
        dialog.resize(920, 640)
        layout = QVBoxLayout(dialog)

        editor = QTextEdit()
        editor.setPlainText(full_page)
        editor.setFont(QFont("Monospace", 10))
        editor.setReadOnly(True)
        layout.addWidget(editor)

        btn_row = QHBoxLayout()
        copy_page_btn = QPushButton("Copy Full Page")
        copy_page_btn.clicked.connect(lambda: QApplication.clipboard().setText(full_page))
        copy_frag_btn = QPushButton("Copy Fragment Only")
        copy_frag_btn.clicked.connect(lambda: QApplication.clipboard().setText(fragment))
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(dialog.close)
        btn_row.addWidget(copy_page_btn)
        btn_row.addWidget(copy_frag_btn)
        btn_row.addStretch()
        btn_row.addWidget(close_btn)
        layout.addLayout(btn_row)

        dialog.exec()

    def _refresh_pending_includes(self):
        """Re-fetch include nodes queued during file load, replacing stale saved children."""
        items = getattr(self, '_pending_include_reloads', [])
        if not items:
            return
        self._pending_include_reloads = []
        prev_loading = getattr(self, '_loading', False)
        self._loading = True
        try:
            for item, url in items:
                try:
                    root = self._load_opml_root(url)
                    body = root.find('.//body')
                    if body is not None:
                        self.tree.setUpdatesEnabled(False)
                        for outline in body.findall('outline'):
                            self.add_outline_to_tree(outline, item, from_include=True)
                        self.tree.setUpdatesEnabled(True)
                        item.setExpanded(True)
                except Exception:
                    pass  # keep item but with no children; offline/error scenario
        finally:
            self._loading = prev_loading
        self.tree.viewport().update()
        self.tree.scheduleDelayedItemsLayout()

    def save_opml(self, filename):
        try:
            root = ET.Element('opml', version="2.0")
            head = ET.SubElement(root, 'head')
            title = ET.SubElement(head, 'title')
            title.text = Path(filename).stem
            
            body = ET.SubElement(root, 'body')
            
            for i in range(self.tree.topLevelItemCount()):
                item = self.tree.topLevelItem(i)
                self.item_to_outline(item, body)
            
            tree = ET.ElementTree(root)
            ET.indent(tree, space='  ')
            tree.write(filename, encoding='UTF-8', xml_declaration=True)
            
            QMessageBox.information(self, "Success", f"File saved!")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save file:\n{e}")

    def item_to_outline(self, item, parent_elem):
        # Get original text (without triangles or icons)
        original_text = item.data(0, Qt.ItemDataRole.UserRole + 2)
        if not original_text:
            # Fallback: strip triangles and icons
            text = item.text(0).replace('🔗 ', '').replace('📄 ', '').replace('▶ ', '').replace('▼ ', '').replace('\n', ' ')
        else:
            text = original_text.replace('\n', ' ')
            # Strip icon prefixes that may have been stored in UserRole+2
            for _pfx in ('📄 ', '🔗 '):
                if text.startswith(_pfx):
                    text = text[len(_pfx):]
                    break

        outline = ET.SubElement(parent_elem, 'outline')
        outline.set('text', html.escape(text))

        url = item.data(0, Qt.ItemDataRole.UserRole)
        if url:
            outline.set('type', 'link')
            outline.set('url', url)

        include_url = item.data(0, Qt.ItemDataRole.UserRole + 1)
        if include_url:
            outline.set('xmlUrl', include_url)

        for i in range(item.childCount()):
            child = item.child(i)
            self.item_to_outline(child, outline)


def main():
    app = QApplication(sys.argv)
    app.setApplicationName("OPML Outliner")
    
    outliner = OPMLOutliner()
    outliner.show()

    if outliner.last_file and Path(outliner.last_file).exists():
        outliner.load_opml(outliner.last_file)
    else:
        outliner.new_file()
    
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
