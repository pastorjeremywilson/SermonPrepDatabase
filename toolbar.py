"""
@author Jeremy G. Wilson

Copyright 2024 Jeremy G. Wilson

This file is a part of the Sermon Prep Database program (v.5.0.1)

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

from PyQt6.QtCore import QSize, Qt
from PyQt6.QtGui import QTextCharFormat, QFont, QTextListFormat, QTextCursor, QIcon
from PyQt6.QtWidgets import QWidget, QHBoxLayout, QPushButton, QLabel, QComboBox, QLineEdit, QTextEdit, QMessageBox


class Toolbar(QWidget):
    """
    Toolbar creates the uppermost QWidget of the GUI that holds formatting, search, and navigation elements.
    """
    def __init__(self, gui, spd):
        super().__init__()
        self.gui = gui
        self.spd = spd
        self.setObjectName('toolbar')

        icon_size = QSize(16, 16)

        toolbar_layout = QHBoxLayout(self)
        toolbar_layout.setContentsMargins(5, 5, 5, 5)

        self.undo_button = QPushButton()
        self.undo_button.setIcon(QIcon('resources/svg/spUndoIcon.svg'))
        self.undo_button.setIconSize(icon_size)
        self.undo_button.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.undo_button.clicked.connect(self.gui.menu_bar.press_ctrl_z)
        self.undo_button.setToolTip('Undo')
        toolbar_layout.addWidget(self.undo_button)

        self.redo_button = QPushButton()
        self.redo_button.setIcon(QIcon('resources/svg/spRedoIcon.svg'))
        self.redo_button.setIconSize(icon_size)
        self.redo_button.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.redo_button.clicked.connect(self.gui.menu_bar.press_ctrl_y)
        self.redo_button.setToolTip('Redo')
        toolbar_layout.addWidget(self.redo_button)
        toolbar_layout.addSpacing(20)

        self.bold_button = QPushButton()
        self.bold_button.setCheckable(True)
        self.bold_button.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.bold_button.clicked.connect(self.set_bold)
        self.bold_button.setIcon(QIcon('resources/svg/spBoldIcon.svg'))
        self.bold_button.setIconSize(icon_size)
        self.bold_button.setToolTip('Bold\n(Ctrl+B)')
        toolbar_layout.addWidget(self.bold_button)

        self.italic_button = QPushButton()
        self.italic_button.setCheckable(True)
        self.italic_button.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.italic_button.clicked.connect(self.set_italic)
        self.italic_button.setIcon(QIcon('resources/svg/spItalicIcon.svg'))
        self.italic_button.setIconSize(icon_size)
        self.italic_button.setToolTip('Italic\n(Ctrl+I)')
        toolbar_layout.addWidget(self.italic_button)

        self.underline_button = QPushButton()
        self.underline_button.setCheckable(True)
        self.underline_button.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.underline_button.clicked.connect(self.set_underline)
        self.underline_button.setIcon(QIcon('resources/svg/spUnderlineIcon.svg'))
        self.underline_button.setIconSize(icon_size)
        self.underline_button.setToolTip('Underline\n(Ctrl+U)')
        toolbar_layout.addWidget(self.underline_button)

        self.bullet_button = QPushButton()
        self.bullet_button.setCheckable(True)
        self.bullet_button.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.bullet_button.clicked.connect(self.set_bullet)
        self.bullet_button.setIcon(QIcon('resources/svg/spBulletIcon.svg'))
        self.bullet_button.setIconSize(icon_size)
        self.bullet_button.setToolTip('Add Bullets\n(Ctrl+Shift+B)')
        toolbar_layout.addWidget(self.bullet_button)
        toolbar_layout.addSpacing(20)

        self.text_visible = QPushButton()
        self.text_visible.setObjectName('text_visible')
        self.text_visible.setCheckable(True)
        self.text_visible.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.text_visible.setToolTip('Show Sermon Text on All Tabs')
        self.text_visible.setIcon(QIcon('resources/svg/spShowText.svg'))
        self.text_visible.setIconSize(QSize(round(icon_size.width() * 2.5), icon_size.height()))
        self.text_visible.clicked.connect(self.keep_text_visible)
        toolbar_layout.addWidget(self.text_visible)

        toolbar_layout.addStretch(1)

        dates_label = QLabel('Choose Record by Date:')
        dates_label.setAutoFillBackground(False)
        toolbar_layout.addWidget(dates_label)

        self.dates_cb = QComboBox()
        self.dates_cb.addItems(self.spd.dates)
        self.dates_cb.currentIndexChanged.connect(lambda: self.spd.get_by_index(self.dates_cb.currentIndex()))
        toolbar_layout.addWidget(self.dates_cb)

        references_label = QLabel('Choose Record by Sermon Reference:')
        references_label.setAutoFillBackground(False)
        toolbar_layout.addWidget(references_label)

        self.references_cb = QComboBox()
        for item in self.spd.references:
            self.references_cb.addItem(item[0])
        self.references_cb.currentIndexChanged.connect(
            lambda: self.get_index_of_reference(self.references_cb.currentIndex()))
        toolbar_layout.addWidget(self.references_cb)

        search_label = QLabel('Keyword Search:')
        search_label.setAutoFillBackground(False)
        toolbar_layout.addWidget(search_label)

        search_field = QLineEdit()
        search_field.setMinimumWidth(200)
        search_field.returnPressed.connect(lambda: self.do_search(search_field.text()))
        toolbar_layout.addWidget(search_field)

        toolbar_layout.addStretch(1)

        self.first_rec_button = QPushButton()
        self.first_rec_button.setIcon(QIcon('resources/svg/spFirstRecIcon.svg'))
        self.first_rec_button.setIconSize(icon_size)
        self.first_rec_button.clicked.connect(lambda: self.spd.first_rec())
        self.first_rec_button.setToolTip('Jump to First Record')
        toolbar_layout.addWidget(self.first_rec_button)

        self.prev_rec_button = QPushButton()
        self.prev_rec_button.setIcon(QIcon('resources/svg/spPrevRecIcon.svg'))
        self.prev_rec_button.setIconSize(icon_size)
        self.prev_rec_button.clicked.connect(lambda: self.spd.prev_rec())
        self.prev_rec_button.setToolTip('Go to Previous Record')
        toolbar_layout.addWidget(self.prev_rec_button)

        self.next_rec_button = QPushButton()
        self.next_rec_button.setIcon(QIcon('resources/svg/spNextRecIcon.svg'))
        self.next_rec_button.clicked.connect(lambda: self.spd.next_rec())
        self.next_rec_button.setToolTip('Go to Next Record')
        toolbar_layout.addWidget(self.next_rec_button)

        self.last_rec_button = QPushButton()
        self.last_rec_button.setIcon(QIcon('resources/svg/spLastRecIcon.svg'))
        self.last_rec_button.clicked.connect(lambda: self.spd.last_rec())
        self.last_rec_button.setToolTip('Jump to Last Record')
        toolbar_layout.addWidget(self.last_rec_button)

        self.new_rec_button = QPushButton()
        self.new_rec_button.setIcon(QIcon('resources/svg/spNewIcon.svg'))
        self.new_rec_button.clicked.connect(lambda: self.spd.new_rec())
        self.new_rec_button.setToolTip('Create a New Record')
        toolbar_layout.addWidget(self.new_rec_button)
        toolbar_layout.addSpacing(20)

        self.save_button = QPushButton()
        self.save_button.setIcon(QIcon('resources/svg/spSaveIcon.svg'))
        self.save_button.setIconSize(icon_size)
        self.save_button.clicked.connect(lambda: self.spd.save_rec())
        self.save_button.setToolTip('Save this Record')
        toolbar_layout.addWidget(self.save_button)

        self.print_button = QPushButton()
        self.print_button.setIcon(QIcon('resources/svg/spPrintIcon.svg'))
        self.print_button.setIconSize(icon_size)
        self.print_button.clicked.connect(self.gui.menu_bar.print_rec)
        self.print_button.setToolTip('Print this Record')
        toolbar_layout.addWidget(self.print_button)

        self.id_label = QLabel()
        toolbar_layout.addWidget(self.id_label)

    def keep_text_visible(self):
        """
        Handle the user's toggling of the text_visible button.
        """
        check_state = self.text_visible.isChecked()

        if check_state:
            # add the reference and passage text to each tab's text_box then make it show
            num_tabs = self.gui.tabbed_frame.count()
            for i in range(num_tabs):
                if i > 0:
                    frame = self.gui.tabbed_frame.widget(i)
                    widget = frame.findChild(QWidget, 'text_box')
                    text_title = widget.findChild(QLabel, 'text_title')
                    text_edit = widget.findChild(QTextEdit, 'text_box_text_edit')
                    text_edit.setText(self.gui.sermon_text_edit.toPlainText())
                    text_title.setText(self.gui.sermon_reference_field.text())
                    if widget:
                        widget.show()
        else:
            # hide the text_box on each tab
            num_tabs = self.gui.tabbed_frame.count()
            for i in range(num_tabs):
                if i > 0:
                    frame = self.gui.tabbed_frame.widget(i)
                    widget = frame.findChild(QWidget, 'text_box')
                    if widget:
                        widget.hide()

    def get_index_of_reference(self, index):
        """
        Method to find the index number of the user's chosen reference.

        :param int index: The index of the reference combo box's chosen reference
        """
        id_to_find = self.spd.references[index][1]
        counter = 0
        for item in self.spd.ids:
            if item == id_to_find:
                break
            counter += 1
        self.spd.get_by_index(counter)

    def do_search(self, text):
        """
        Method to call get_search_results from SermonPrepDatabase and display the results

        :param str text: The user's search term(s)
        """
        result_list = self.spd.get_search_results(text)
        if len(result_list) == 0:
            QMessageBox.information(
                None,
                'No Results',
                'No results were found. Please try your search again.',
                QMessageBox.StandardButton.Ok
            )
        else:
            from gui import SearchBox
            search_box = SearchBox(self.gui)
            self.gui.tabbed_frame.addTab(search_box, QIcon('resources/svg/spSearchIcon.svg'), 'Search')
            search_box.show_results(result_list)
            self.gui.tabbed_frame.setCurrentWidget(search_box)

    def set_bold(self):
        """
        Method to toggle the bold state of the text at the user's cursor or selection.
        """
        component = self.gui.focusWidget()
        if isinstance(component, QTextEdit):
            cursor = component.textCursor()
            # handle this differently if the user has a section of text selected
            if cursor.hasSelection():
                selection_start = cursor.selectionStart()
                selection_end = cursor.selectionEnd()
                cursor.setPosition(selection_start, QTextCursor.MoveMode.MoveAnchor)
                cursor.setPosition(selection_end, QTextCursor.MoveMode.KeepAnchor)
                char_format = cursor.charFormat()
                if char_format.font().bold():
                    char_format.setFontWeight(QFont.Weight.Normal)
                    cursor.mergeCharFormat(char_format)
                else:
                    char_format.setFontWeight(QFont.Weight.Bold)
                    cursor.mergeCharFormat(char_format)
            else:
                font = cursor.charFormat().font()
                if font.weight() == QFont.Weight.Normal:
                    font.setWeight(QFont.Weight.Bold)
                    component.setCurrentFont(font)
                else:
                    font.setWeight(QFont.Weight.Normal)
                    component.setCurrentFont(font)
        else:
            self.bold_button.setChecked(False)

    def set_italic(self):
        """
        Method to toggle the italic state of the text at the user's cursor or selection.
        """
        component = self.gui.focusWidget()
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
        """
        Method to toggle the underline state of the text at the user's cursor or selection.
        """
        component = self.gui.focusWidget()
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
        """
        Method to toggle the bulleted list state of the text at the user's cursor or selection.
        """
        component = self.gui.focusWidget()
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
