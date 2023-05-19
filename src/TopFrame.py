"""
@author Jeremy G. Wilson

Copyright 2023 Jeremy G. Wilson

This file is a part of the Sermon Prep Database program (v.3.4.3)

Sermon Prep Database is free software: you can redistribute it and/or
modify it under the terms of the GNU General Public License (GNU GPL)
published by the Free Software Foundation, either version 3 of the
License, or (at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <http://www.gnu.org/licenses/>.

The Sermon Prep Database program includes Artifex Software's GhostScript,
licensed under the GNU Affero General Public License (GNU AGPL). See
https://www.ghostscript.com/licensing/index.html for more information.
"""
from PyQt5.QtCore import QSize, Qt
from PyQt5.QtGui import QTextCharFormat, QFont, QTextListFormat, QTextCursor, QIcon
from PyQt5.QtWidgets import QWidget, QHBoxLayout, QPushButton, QLabel, QComboBox, QLineEdit, QTextEdit, QMessageBox


class TopFrame(QWidget):

    def __init__(self, win, gui, spd):
        super().__init__()
        self.win = win
        self.gui = gui
        self.spd = spd

        icon_size = QSize(24, 24)
        button_frame = QWidget()
        self.setStyleSheet('background-color: white;')

        button_frame_layout = QHBoxLayout()
        self.setLayout(button_frame_layout)

        undo_button = QPushButton()
        undo_button.setFixedSize(icon_size)
        undo_button.setStyleSheet('QPushButton { icon: url(' + self.spd.cwd + 'resources/spUndoIcon.png); border: 0 } QPushButton:pressed { background-color: white; } QPushButton:hover { background-color: ' + self.gui.background_color + '; }')
        undo_button.setFocusPolicy(Qt.NoFocus)
        undo_button.clicked.connect(self.gui.menu_bar.press_ctrl_z)
        undo_button.setToolTip('Undo')
        button_frame_layout.addWidget(undo_button)

        redo_button = QPushButton()
        redo_button.setFixedSize(icon_size)
        redo_button.setStyleSheet('QPushButton { icon: url(' + self.spd.cwd + 'resources/spRedoIcon.png); border: 0 } QPushButton:checked { background-color: white; } QPushButton:hover { background-color: ' + self.gui.background_color + '; }')
        redo_button.setFocusPolicy(Qt.NoFocus)
        redo_button.clicked.connect(self.gui.menu_bar.press_ctrl_y)
        redo_button.setToolTip('Redo')
        button_frame_layout.addWidget(redo_button)
        button_frame_layout.addSpacing(20)

        self.bold_button = QPushButton()
        self.bold_button.setCheckable(True)
        self.bold_button.setFocusPolicy(Qt.NoFocus)
        self.bold_button.clicked.connect(self.set_bold)
        self.bold_button.setFixedSize(icon_size)
        self.bold_button.setStyleSheet('QPushButton { icon: url(' + self.spd.cwd + 'resources/spBoldIcon.png); border: 0 } QPushButton:checked { background-color: ' + self.gui.background_color + '; } QPushButton:hover { background-color: ' + self.gui.background_color + '; }')
        self.bold_button.setToolTip('Bold\n(Ctrl+B)')
        button_frame_layout.addWidget(self.bold_button)

        self.italic_button = QPushButton()
        self.italic_button.setCheckable(True)
        self.italic_button.setFocusPolicy(Qt.NoFocus)
        self.italic_button.clicked.connect(self.set_italic)
        self.italic_button.setFixedSize(icon_size)
        self.italic_button.setStyleSheet('QPushButton { icon: url(' + self.spd.cwd + 'resources/spItalicIcon.png); border: 0 } QPushButton:checked { background-color: ' + self.gui.background_color + '; } QPushButton:hover { background-color: ' + self.gui.background_color + ';  }')
        self.italic_button.setToolTip('Italic\n(Ctrl+I)')
        button_frame_layout.addWidget(self.italic_button)

        self.underline_button = QPushButton()
        self.underline_button.setCheckable(True)
        self.underline_button.setFocusPolicy(Qt.NoFocus)
        self.underline_button.clicked.connect(self.set_underline)
        self.underline_button.setFixedSize(icon_size)
        self.underline_button.setStyleSheet('QPushButton { icon: url(' + self.spd.cwd + 'resources/spUnderLineIcon.png); border: 0 } QPushButton:checked { background-color: ' + self.gui.background_color + '; } QPushButton:hover { background-color: ' + self.gui.background_color + '; }')
        self.underline_button.setToolTip('Underline\n(Ctrl+U)')
        button_frame_layout.addWidget(self.underline_button)

        self.bullet_button = QPushButton()
        self.bullet_button.setCheckable(True)
        self.bullet_button.setFocusPolicy(Qt.NoFocus)
        self.bullet_button.clicked.connect(self.set_bullet)
        self.bullet_button.setFixedSize(icon_size)
        self.bullet_button.setStyleSheet('QPushButton { icon: url(' + self.spd.cwd + 'resources/spBulletIcon.png); border: 0 } QPushButton:checked { background-color: ' + self.gui.background_color + '; } QPushButton:hover { background-color: ' + self.gui.background_color + '; }')
        self.bullet_button.setToolTip('Add Bullets\n(Ctrl+Shift+B)')
        button_frame_layout.addWidget(self.bullet_button)

        text_visible = QPushButton()
        text_visible.setCheckable(True)
        text_visible.setFixedSize(QSize(icon_size.width() * 3, icon_size.height()))
        text_visible.setIconSize(QSize(round(icon_size.width() * 3 / 2), round(icon_size.height() / 2)))
        text_visible.setFocusPolicy(Qt.NoFocus)
        text_visible.setToolTip('Show Sermon Text on All Tabs')
        text_visible.setStyleSheet('QPushButton { icon: url(' + self.spd.cwd + 'resources/spShowText.png); border: 0; padding: 0;} QPushButton:checked { icon: url(' + self.spd.cwd + 'resources/spShowTextChecked.png); }')
        text_visible.clicked.connect(lambda: self.keep_text_visible(text_visible.isChecked()))
        button_frame_layout.addWidget(text_visible)

        button_frame_layout.addStretch(1)

        dates_label = QLabel('Choose Record by Date:')
        button_frame_layout.addWidget(dates_label)

        self.dates_cb = QComboBox()
        self.dates_cb.setStyleSheet('selection-background-color: ' + self.gui.background_color + '; selection-color: black')
        self.dates_cb.addItems(self.spd.dates)
        self.dates_cb.currentIndexChanged.connect(lambda: self.spd.get_by_index(self.dates_cb.currentIndex()))
        button_frame_layout.addWidget(self.dates_cb)

        references_label = QLabel('Choose Record by Sermon Reference:')
        button_frame_layout.addWidget(references_label)

        self.references_cb = QComboBox()
        self.references_cb.setStyleSheet(
            'selection-background-color: ' + self.gui.background_color + '; selection-color: black')
        for item in self.spd.references:
            self.references_cb.addItem(item[0])
        self.references_cb.currentIndexChanged.connect(
            lambda: self.get_index_of_reference(self.references_cb.currentIndex()))
        button_frame_layout.addWidget(self.references_cb)

        search_label = QLabel('Keyword Search:')
        button_frame_layout.addWidget(search_label)

        search_field = QLineEdit()
        search_field.setMinimumWidth(200)
        search_field.returnPressed.connect(lambda: self.do_search(search_field.text()))
        button_frame_layout.addWidget(search_field)

        button_frame_layout.addStretch(1)

        self.first_rec_button = QPushButton()
        self.first_rec_button.setFixedSize(icon_size)
        self.first_rec_button.setStyleSheet('QPushButton { icon: url(' + self.spd.cwd + 'resources/spFirstRecIcon.png); border: 0 } QPushButton:pressed { background-color: white; } QPushButton:hover { background-color: ' + self.gui.background_color + '; }')
        self.first_rec_button.clicked.connect(lambda: self.spd.first_rec())
        self.first_rec_button.setToolTip('Jump to First Record')
        button_frame_layout.addWidget(self.first_rec_button)

        self.prev_rec_button = QPushButton()
        self.prev_rec_button.setFixedSize(icon_size)
        self.prev_rec_button.setStyleSheet('QPushButton { icon: url(' + self.spd.cwd + 'resources/spPrevRecIcon.png); border: 0 } QPushButton:pressed { background-color: white; } QPushButton:hover { background-color: ' + self.gui.background_color + '; }')
        self.prev_rec_button.clicked.connect(lambda: self.spd.prev_rec())
        self.prev_rec_button.setToolTip('Go to Previous Record')
        button_frame_layout.addWidget(self.prev_rec_button)

        self.next_rec_button = QPushButton()
        self.next_rec_button.setFixedSize(icon_size)
        self.next_rec_button.setStyleSheet('QPushButton { icon: url(' + self.spd.cwd + 'resources/spNextRecIcon.png); border: 0 } QPushButton:pressed { background-color: white; } QPushButton:hover { background-color: ' + self.gui.background_color + '; }')
        self.next_rec_button.clicked.connect(lambda: self.spd.next_rec())
        self.next_rec_button.setToolTip('Go to Next Record')
        button_frame_layout.addWidget(self.next_rec_button)

        self.last_rec_button = QPushButton()
        self.last_rec_button.setFixedSize(icon_size)
        self.last_rec_button.setStyleSheet('QPushButton { icon: url(' + self.spd.cwd + 'resources/spLastRecIcon.png); border: 0 } QPushButton:pressed { background-color: white; } QPushButton:hover { background-color: ' + self.gui.background_color + '; }')
        self.last_rec_button.clicked.connect(lambda: self.spd.last_rec())
        self.last_rec_button.setToolTip('Jump to Last Record')
        button_frame_layout.addWidget(self.last_rec_button)

        new_rec_button = QPushButton()
        new_rec_button.setFixedSize(icon_size)
        new_rec_button.setStyleSheet('QPushButton { icon: url(' + self.spd.cwd + 'resources/spNewRecIcon.png); border: 0 } QPushButton:pressed { background-color: white; } QPushButton:hover { background-color: ' + self.gui.background_color + '; }')
        new_rec_button.clicked.connect(lambda: self.spd.new_rec())
        new_rec_button.setToolTip('Create a New Record')
        button_frame_layout.addWidget(new_rec_button)
        button_frame_layout.addSpacing(20)

        save_button = QPushButton()
        save_button.setFixedSize(icon_size)
        save_button.setStyleSheet('QPushButton { icon: url(' + self.spd.cwd + 'resources/spSaveIcon.png); border: 0 } QPushButton:pressed { background-color: white; } QPushButton:hover { background-color: ' + self.gui.background_color + '; }')
        save_button.clicked.connect(lambda: self.spd.save_rec())
        save_button.setToolTip('Save this Record')
        button_frame_layout.addWidget(save_button)

        print_button = QPushButton()
        print_button.setFixedSize(icon_size)
        print_button.setStyleSheet('QPushButton { icon: url(' + self.spd.cwd + 'resources/spPrintIcon.png); border: 0 } QPushButton:pressed { background-color: white; } QPushButton:hover { background-color: ' + self.gui.background_color + '; }')
        print_button.clicked.connect(self.gui.menu_bar.print_rec)
        print_button.setToolTip('Print this Record')
        button_frame_layout.addWidget(print_button)

        self.id_label = QLabel()
        button_frame_layout.addWidget(self.id_label)

        self.gui.layout.addWidget(button_frame)

    def keep_text_visible(self, check_state):
        if check_state:
            num_tabs = self.gui.tabbed_frame.count()
            for i in range(num_tabs):
                if i > 0:
                    frame = self.gui.tabbed_frame.widget(i)
                    widget = frame.findChild(QWidget, 'text_box')
                    text_title = widget.findChild(QLabel, 'text_title')
                    text_edit = widget.findChild(QTextEdit, 'text_edit')
                    text_edit.setText(self.gui.sermon_text_edit.toPlainText())
                    text_title.setText(self.gui.sermon_reference_field.text())
                    if widget:
                        widget.show()
        else:
            num_tabs = self.gui.tabbed_frame.count()
            for i in range(num_tabs):
                if i > 0:
                    frame = self.gui.tabbed_frame.widget(i)
                    widget = frame.findChild(QWidget, 'text_box')
                    if widget:
                        widget.hide()

    def get_index_of_reference(self, index):
        id_to_find = self.spd.references[index][1]
        counter = 0
        for item in self.spd.ids:
            if item == id_to_find:
                break
            counter += 1
        self.spd.get_by_index(counter)

    def do_search(self, text):
        result_list = self.spd.get_search_results(text)
        if len(result_list) == 0:
            QMessageBox.information(
                None,
                'No Results',
                'No results were found. Please try your search again.',
                QMessageBox.Ok
            )
        else:
            from gui import SearchBox
            search_box = SearchBox(self.gui)
            self.gui.tabbed_frame.addTab(search_box, QIcon(self.gui.spd.cwd + 'resources/searchIcon.png'), 'Search')
            search_box.show_results(result_list)
            self.gui.tabbed_frame.setCurrentWidget(search_box)

    def set_bold(self):
        component = self.win.focusWidget()
        if isinstance(component, QTextEdit):
            cursor = component.textCursor()
            if cursor.hasSelection():
                selection_start = cursor.selectionStart()
                selection_end = cursor.selectionEnd()
                cursor.setPosition(selection_start, QTextCursor.MoveMode.MoveAnchor)
                cursor.setPosition(selection_end, QTextCursor.MoveMode.KeepAnchor)
                char_format = cursor.charFormat()
                if char_format.font().bold():
                    char_format.setFontWeight(QFont.Normal)
                    cursor.mergeCharFormat(char_format)
                else:
                    char_format.setFontWeight(QFont.Bold)
                    cursor.mergeCharFormat(char_format)
            else:
                font = cursor.charFormat().font()
                if font.weight() == QFont.Normal:
                    font.setWeight(QFont.Bold)
                    component.setCurrentFont(font)
                else:
                    font.setWeight(QFont.Normal)
                    component.setCurrentFont(font)
        else:
            self.bold_button.setChecked(False)

    def set_italic(self):
        component = self.win.focusWidget()
        if isinstance(component, QTextEdit):
            cursor = component.textCursor()
            if cursor.hasSelection():
                font = QTextCharFormat(cursor.charFormat())
                if not font.fontItalic():
                    font.setFontItalic(True)
                    cursor.setCharFormat(QTextCharFormat(font))
                else:
                    font.setFontItalic(False)
                    cursor.mergeCharFormat(QTextCharFormat(font))
            else:
                font = cursor.charFormat().font()
                if not font.italic():
                    font.setItalic(True)
                    component.setCurrentFont(font)
                else:
                    font.setItalic(False)
                    component.setCurrentFont(font)
        else:
            self.italic_button.setChecked(False)

    def set_underline(self):
        component = self.win.focusWidget()
        if isinstance(component, QTextEdit):
            cursor = component.textCursor()
            if cursor.hasSelection():
                font = QTextCharFormat(cursor.charFormat())
                if not font.fontUnderline():
                    font.setFontUnderline(True)
                    cursor.setCharFormat(QTextCharFormat(font))
                else:
                    font.setFontUnderline(False)
                    cursor.mergeCharFormat(QTextCharFormat(font))
            else:
                font = cursor.charFormat().font()
                if not font.underline():
                    font.setUnderline(True)
                    component.setCurrentFont(font)
                else:
                    font.setUnderline(False)
                    component.setCurrentFont(font)
        else:
            self.underline_button.setChecked(False)

    def set_bullet(self):
        component = self.win.focusWidget()
        if isinstance(component, QTextEdit):
            cursor = component.textCursor()
            text_list = cursor.currentList()
            if text_list:
                start = cursor.selectionStart()
                end = cursor.selectionEnd()
                removed = 0
                for i in range(text_list.count()):
                    item = text_list.item(i - removed)
                    if (item.position() <= end and
                            item.position() + item.length() > start):
                        text_list.remove(item)
                        block_cursor = QTextCursor(item)
                        block_format = block_cursor.blockFormat()
                        block_format.setIndent(0)
                        block_cursor.mergeBlockFormat(block_format)
                        removed += 1
                component.setTextCursor(cursor)
            else:
                list_format = QTextListFormat()
                style = QTextListFormat.Style.ListDisc
                list_format.setStyle(style)
                cursor.createList(list_format)
                component.setTextCursor(cursor)
        else:
            self.bullet_button.setChecked(False)
