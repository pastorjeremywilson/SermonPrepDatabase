"""
@author Jeremy G. Wilson

Copyright 2023 Jeremy G. Wilson

This file is a part of the Sermon Prep Database program (v.4.0.0)

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

import logging
import re
import shutil
import sys
from os.path import exists

from PyQt5.QtCore import Qt, QSize, QDate, QDateTime, pyqtSignal, QObject
from PyQt5.QtGui import QIcon, QFont, QKeyEvent, QTextCursor, QStandardItemModel, QStandardItem, QPixmap
from PyQt5.QtWidgets import QBoxLayout, QWidget, QUndoStack, QMessageBox, QTabWidget, QGridLayout, QLabel, QLineEdit, \
    QCheckBox, QDateEdit, QTextEdit, QAction, QMainWindow, QVBoxLayout, QHBoxLayout, QPushButton, QTableView
from symspellpy import Verbosity

from getScripture import GetScripture
from MenuBar import MenuBar
from TopFrame import TopFrame


class GUI(QObject):
    spd = None
    win = None
    undo_stack = None
    changes = False
    font_family = None
    font_size = None
    gs = None

    open_import_splash = pyqtSignal()
    change_import_splash_dir = pyqtSignal(str)
    change_import_splash_file = pyqtSignal(str)
    close_import_splash = pyqtSignal()
    
    def __init__(self, spd):
        super().__init__()
        self.spd = spd
        self.check_for_db()
        self.init_components()

    def init_components(self):
        self.spd.get_ids()
        self.spd.get_date_list()
        self.spd.get_scripture_list()
        self.spd.get_user_settings()
        self.spd.backup_db()

        self.accent_color = self.spd.user_settings[1]
        self.background_color = self.spd.user_settings[2]
        self.font_family = self.spd.user_settings[3]
        self.font_size = self.spd.user_settings[4]
        self.standard_font = QFont(self.font_family, int(self.font_size))
        self.bold_font = QFont(self.font_family, int(self.font_size), QFont.Bold)

        self.win = Win(self)
        icon_pixmap = QPixmap(self.spd.cwd + 'resources/icon.png')
        self.win.setWindowIcon(QIcon(icon_pixmap))

        self.layout = QBoxLayout(QBoxLayout.TopToBottom)
        self.main_widget = QWidget()
        self.main_widget.setStyleSheet('background-color: white;')
        self.win.setCentralWidget(self.main_widget)
        self.main_widget.setLayout(self.layout)

        self.undo_stack = QUndoStack(self.main_widget)

        self.menu_bar = MenuBar(self.win, self, self.spd)
        self.top_frame = TopFrame(self.win, self, self.spd)
        self.layout.addWidget(self.top_frame)

        self.build_tabbed_frame()
        self.build_scripture_tab()
        self.build_exegesis_tab()
        self.build_outline_tab()
        self.build_research_tab()
        self.build_sermon_tab()

        self.set_style_sheets()

        self.win.showMaximized()

    # check if the database file exists, and that the user_settings table exists. Prompt to create new if not.
    def check_for_db(self):
        if not exists(self.spd.db_loc):
            response = QMessageBox.question(
                None,
                'Database Not Found',
                'It looks like this is the first time you\'ve run Sermon Prep Database v3.3.4.\n'
                'Would you like to import an old database?\n'
                '(Choose "No" to create a new database)',
                QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel
            )

            if response == QMessageBox.Yes:
                from ConvertDatabase import ConvertDatabase
                ConvertDatabase(self.spd)
            elif response == QMessageBox.No:
                shutil.copy(self.spd.cwd + 'resources/database_template.db', self.spd.db_loc)
                QMessageBox.information(None, 'Database Created', 'A new database has been created.', QMessageBox.Ok)
                self.spd.app.processEvents()
            else:
                quit(0)

        self.spd.write_to_log('checkForDB completed')
    
    def build_tabbed_frame(self):
        self.tabbed_frame = QTabWidget()
        self.tabbed_frame.setTabPosition(QTabWidget.West)
        self.tabbed_frame.setIconSize(QSize(24, 24))
        self.layout.addWidget(self.tabbed_frame)
        
    def build_scripture_tab(self, insert=False):
        self.scripture_frame = QWidget()
        self.scripture_frame.setStyleSheet('background-color: ' + self.background_color)
        self.scripture_frame_layout = QGridLayout()
        self.scripture_frame_layout.setColumnStretch(0, 1)
        self.scripture_frame_layout.setColumnStretch(1, 1)
        self.scripture_frame_layout.setColumnStretch(2, 0)
        self.scripture_frame.setLayout(self.scripture_frame_layout)
        
        pericope_label = QLabel(self.spd.user_settings[5])
        self.scripture_frame_layout.addWidget(pericope_label, 0, 0)
        
        pericope_field = QLineEdit()
        pericope_field.textEdited.connect(self.changes_detected)
        self.scripture_frame_layout.addWidget(pericope_field, 1, 0)
        
        pericope_text_label = QLabel(self.spd.user_settings[6])
        self.scripture_frame_layout.addWidget(pericope_text_label, 2, 0)

        pericope_text_edit = CustomTextEdit(self.win, self)
        pericope_text_edit.cursorPositionChanged.connect(lambda: self.set_style_buttons(pericope_text_edit))
        self.scripture_frame_layout.addWidget(pericope_text_edit, 3, 0)
        
        sermon_reference_label = QLabel(self.spd.user_settings[7])
        self.scripture_frame_layout.addWidget(sermon_reference_label, 0, 1)

        self.sermon_reference_field = QLineEdit()

        if exists(self.spd.app_dir + '/my_bible.xml'):
            self.scripture_frame_layout.addWidget(self.sermon_reference_field, 1, 1)

            self.auto_fill_checkbox = QCheckBox('Auto-fill ' + self.spd.user_settings[8])
            self.auto_fill_checkbox.setChecked(True)
            self.scripture_frame_layout.addWidget(self.auto_fill_checkbox, 1, 2)
            if self.spd.auto_fill:
                self.auto_fill_checkbox.setChecked(True)
            else:
                self.auto_fill_checkbox.setChecked(False)
            self.auto_fill_checkbox.stateChanged.connect(self.auto_fill)
        else:
            self.scripture_frame_layout.addWidget(self.sermon_reference_field, 1, 1, 1, 2)
        
        sermon_text_label = QLabel(self.spd.user_settings[8])
        self.scripture_frame_layout.addWidget(sermon_text_label, 2, 1, 1, 2)
        
        self.sermon_text_edit = CustomTextEdit(self.win, self)
        self.sermon_text_edit.cursorPositionChanged.connect(lambda: self.set_style_buttons(self.sermon_text_edit))
        self.scripture_frame_layout.addWidget(self.sermon_text_edit, 3, 1, 1, 2)

        self.sermon_reference_field.textChanged.connect(self.reference_changes)
        self.sermon_reference_field.textEdited.connect(self.reference_changes)

        if insert:
            self.tabbed_frame.insertTab(
                0,
                self.scripture_frame,
                QIcon(self.spd.cwd + 'resources/scriptureIcon.png'),
                'Scripture'
            )
        else:
            self.tabbed_frame.addTab(
                self.scripture_frame,
                QIcon(self.spd.cwd + 'resources/scriptureIcon.png'),
                'Scripture'
            )
        
    def build_exegesis_tab(self):
        self.exegesis_frame = QWidget()
        self.exegesis_frame_layout = QGridLayout()
        self.exegesis_frame_layout.setColumnMinimumWidth(1, 20)
        self.exegesis_frame_layout.setColumnMinimumWidth(3, 20)
        self.exegesis_frame_layout.setColumnMinimumWidth(5, 20)
        self.exegesis_frame_layout.setRowMinimumHeight(2, 50)
        self.exegesis_frame_layout.setRowMinimumHeight(5, 50)
        self.exegesis_frame_layout.setRowStretch(8, 100)
        self.exegesis_frame.setLayout(self.exegesis_frame_layout)
        
        fcft_label = QLabel(self.spd.user_settings[9])
        self.exegesis_frame_layout.addWidget(fcft_label, 0, 0)
        
        fcft_text = CustomTextEdit(self.win, self)
        fcft_text.cursorPositionChanged.connect(lambda: self.set_style_buttons(fcft_text))
        self.exegesis_frame_layout.addWidget(fcft_text, 1, 0)
        
        gat_label = QLabel(self.spd.user_settings[10])
        self.exegesis_frame_layout.addWidget(gat_label, 3, 0)
        
        gat_text = CustomTextEdit(self.win, self)
        gat_text.cursorPositionChanged.connect(lambda: self.set_style_buttons(gat_text))
        self.exegesis_frame_layout.addWidget(gat_text, 4, 0)
        
        cpt_label = QLabel(self.spd.user_settings[11])
        self.exegesis_frame_layout.addWidget(cpt_label, 6, 0)
        
        cpt_text = CustomTextEdit(self.win, self)
        cpt_text.cursorPositionChanged.connect(lambda: self.set_style_buttons(cpt_text))
        self.exegesis_frame_layout.addWidget(cpt_text, 7, 0)
        
        pb_label = QLabel(self.spd.user_settings[12])
        self.exegesis_frame_layout.addWidget(pb_label, 3, 2)
        
        pb_text = CustomTextEdit(self.win, self)
        pb_text.cursorPositionChanged.connect(lambda: self.set_style_buttons(pb_text))
        self.exegesis_frame_layout.addWidget(pb_text, 4, 2)
        
        fcfs_label = QLabel(self.spd.user_settings[13])
        self.exegesis_frame_layout.addWidget(fcfs_label, 0, 4)
        
        fcfs_text = CustomTextEdit(self.win, self)
        fcfs_text.cursorPositionChanged.connect(lambda: self.set_style_buttons(fcfs_text))
        self.exegesis_frame_layout.addWidget(fcfs_text, 1, 4)
        
        gas_label = QLabel(self.spd.user_settings[14])
        self.exegesis_frame_layout.addWidget(gas_label, 3, 4)
        
        gas_text = CustomTextEdit(self.win, self)
        gas_text.cursorPositionChanged.connect(lambda: self.set_style_buttons(gas_text))
        self.exegesis_frame_layout.addWidget(gas_text, 4, 4)
        
        cps_label = QLabel(self.spd.user_settings[15])
        self.exegesis_frame_layout.addWidget(cps_label, 6, 4)
        
        cps_text = CustomTextEdit(self.win, self)
        cps_text.cursorPositionChanged.connect(lambda: self.set_style_buttons(cps_text))
        self.exegesis_frame_layout.addWidget(cps_text, 7, 4)

        scripture_box = ScriptureBox(self.background_color)
        self.exegesis_frame_layout.addWidget(scripture_box, 0, 6, 8, 1)
        
        self.tabbed_frame.addTab(self.exegesis_frame, QIcon(self.spd.cwd + 'resources/exegIcon.png'), 'Exegesis')
        
    def build_outline_tab(self):
        self.outline_frame = QWidget()
        self.outline_frame_layout = QGridLayout()
        self.outline_frame_layout.setColumnMinimumWidth(1, 20)
        self.outline_frame_layout.setColumnMinimumWidth(3, 20)
        self.outline_frame_layout.setColumnMinimumWidth(5, 20)
        self.outline_frame.setLayout(self.outline_frame_layout)
        
        scripture_outline_label = QLabel(self.spd.user_settings[16])
        self.outline_frame_layout.addWidget(scripture_outline_label, 0, 0)
        
        scripture_outline_text = CustomTextEdit(self.win, self)
        scripture_outline_text.cursorPositionChanged.connect(lambda: self.set_style_buttons(scripture_outline_text))
        self.outline_frame_layout.addWidget(scripture_outline_text, 1, 0)
        
        sermon_outline_label = QLabel(self.spd.user_settings[17])
        self.outline_frame_layout.addWidget(sermon_outline_label, 0, 2)
        
        sermon_outline_text = CustomTextEdit(self.win, self)
        sermon_outline_text.cursorPositionChanged.connect(lambda: self.set_style_buttons(sermon_outline_text))
        self.outline_frame_layout.addWidget(sermon_outline_text, 1, 2)
        
        illustration_label = QLabel(self.spd.user_settings[18])
        self.outline_frame_layout.addWidget(illustration_label, 0, 4)
        
        illustration_text = CustomTextEdit(self.win, self)
        illustration_text.cursorPositionChanged.connect(lambda: self.set_style_buttons(illustration_text))
        self.outline_frame_layout.addWidget(illustration_text, 1, 4)

        scripture_box = ScriptureBox(self.background_color)
        self.outline_frame_layout.addWidget(scripture_box, 0, 6, 5, 1)

        self.tabbed_frame.addTab(self.outline_frame, QIcon(self.spd.cwd + 'resources/outlineIcon.png'), 'Outlines')
        
    def build_research_tab(self):
        self.research_frame = QWidget()
        self.research_frame_layout = QGridLayout()
        self.research_frame.setLayout(self.research_frame_layout)
        self.research_frame_layout.setColumnMinimumWidth(1, 20)
        
        research_label = QLabel(self.spd.user_settings[19])
        self.research_frame_layout.addWidget(research_label, 0, 0)
        
        research_text = CustomTextEdit(self.win, self)
        research_text.setObjectName('research')
        research_text.cursorPositionChanged.connect(lambda: self.set_style_buttons(research_text))
        self.research_frame_layout.addWidget(research_text, 1, 0)

        scripture_box = ScriptureBox(self.background_color)
        self.research_frame_layout.addWidget(scripture_box, 0, 2, 2, 1)
        
        self.tabbed_frame.addTab(self.research_frame, QIcon(self.spd.cwd + 'resources/researchIcon.png'), 'Research')
        
    def build_sermon_tab(self):
        self.sermon_frame = QWidget()
        self.sermon_frame_layout = QGridLayout()
        self.sermon_frame.setLayout(self.sermon_frame_layout)
        self.sermon_frame_layout.setColumnMinimumWidth(2, 20)
        self.sermon_frame_layout.setColumnMinimumWidth(5, 20)
        self.sermon_frame_layout.setColumnMinimumWidth(8, 20)
        self.sermon_frame_layout.setRowMinimumHeight(2, 20)
        self.sermon_frame_layout.setRowMinimumHeight(5, 20)
        
        sermon_title_label = QLabel(self.spd.user_settings[20])
        self.sermon_frame_layout.addWidget(sermon_title_label, 0, 0, 1, 2)
        
        sermon_title_field = QLineEdit()
        sermon_title_field.textEdited.connect(self.changes_detected)
        self.sermon_frame_layout.addWidget(sermon_title_field, 1, 0, 1, 2)
        
        sermon_date_label = QLabel(self.spd.user_settings[21])
        self.sermon_frame_layout.addWidget(sermon_date_label, 0, 3, 1, 2)

        self.sermon_date_edit = QDateEdit()
        self.sermon_date_edit.setCalendarPopup(True)
        self.sermon_date_edit.setMinimumHeight(30)
        self.sermon_date_edit.dateChanged.connect(self.date_changes)
        self.sermon_frame_layout.addWidget(self.sermon_date_edit, 1, 3, 1, 2)
        
        sermon_location_label = QLabel(self.spd.user_settings[22])
        self.sermon_frame_layout.addWidget(sermon_location_label, 0, 6, 1, 2)
        
        sermon_location_field = QLineEdit()
        sermon_location_field.textEdited.connect(self.changes_detected)
        self.sermon_frame_layout.addWidget(sermon_location_field, 1, 6, 1, 2)
        
        ctw_label = QLabel(self.spd.user_settings[23])
        self.sermon_frame_layout.addWidget(ctw_label, 3, 0, 1, 2)
        
        ctw_field = QLineEdit()
        ctw_field.textEdited.connect(self.changes_detected)
        self.sermon_frame_layout.addWidget(ctw_field, 4, 0, 1, 2)
        
        hr_label = QLabel(self.spd.user_settings[24])
        self.sermon_frame_layout.addWidget(hr_label, 3, 6, 1, 2)
        
        hr_field = QLineEdit()
        hr_field.textEdited.connect(self.changes_detected)
        self.sermon_frame_layout.addWidget(hr_field, 4, 6, 1, 2)
        
        sermon_label = QLabel(self.spd.user_settings[25])
        self.sermon_frame_layout.addWidget(sermon_label, 6, 0)
        
        sermon_text = CustomTextEdit(self.win, self)
        sermon_text.cursorPositionChanged.connect(lambda: self.set_style_buttons(sermon_text))
        self.sermon_frame_layout.addWidget(sermon_text, 7, 0, 1, 8)

        scripture_box = ScriptureBox(self.background_color)
        self.sermon_frame_layout.addWidget(scripture_box, 0, 9, 8, 1)
        
        self.tabbed_frame.addTab(self.sermon_frame, QIcon(self.spd.cwd + 'resources/sermonIcon.png'), 'Sermon')

    def set_style_sheets(self):
        self.tabbed_frame.setStyleSheet('''
            QTabWidget::pane {
                border: 50px solid ''' + self.background_color + ''';}
            QTabBar::tab {
                background-color: ''' + self.accent_color + ''';
                color: white;
                font-family: "''' + self.font_family + '''";
                font-size: ''' + str(self.font_size) + '''pt;
                font-weight: bold;
                width: 40px;
                height: 120px;
                padding: 10px;
                margin-bottom: 5px;}
            QTabBar::tab:selected {
                background-color: ''' + self.background_color + ''';
                color: black;
                font-family: "''' + self.font_family + '''";
                font-size: 20px;
                font-weight: bold;
                width: 50px;}
            ''')

        standard_style_sheet = ('''
            QWidget {
                background-color: ''' + self.background_color + ''';}
            QLabel {
                font-family: "''' + self.font_family + '''";
                font-size: ''' + str(self.font_size) + '''pt;}
            QLineEdit {
                background-color: white;
                font-family: "''' + self.font_family + '''";
                font-size: ''' + str(self.font_size) + '''pt;
                padding: 3px;
                border: 1px solid ''' + self.accent_color + ''';}
            QTextEdit {
                background-color: white;
                font-family: "''' + self.font_family + '''";
                font-size: ''' + str(self.font_size) + '''pt;
                padding: 3px;
                border: 1px solid ''' + self.accent_color + ''';}
            ''')

        self.scripture_frame.setStyleSheet(standard_style_sheet)
        self.exegesis_frame.setStyleSheet(standard_style_sheet)
        self.outline_frame.setStyleSheet(standard_style_sheet)
        self.research_frame.setStyleSheet(standard_style_sheet)
        self.sermon_frame.setStyleSheet(standard_style_sheet)

        for component in self.tabbed_frame.findChildren(CustomTextEdit, 'custom_text_edit'):
            component.document().setDefaultFont(QFont(self.font_family, int(self.font_size)))

    def set_style_buttons(self, component):
        cursor = component.textCursor()
        font = cursor.charFormat().font()
        if font.weight() == QFont.Normal:
            self.top_frame.bold_button.setChecked(False)
        else:
            self.top_frame.bold_button.setChecked(True)
        if font.italic():
            self.top_frame.italic_button.setChecked(True)
        else:
            self.top_frame.italic_button.setChecked(False)
        if font.underline():
            self.top_frame.underline_button.setChecked(True)
        else:
            self.top_frame.underline_button.setChecked(False)

        text_list = cursor.currentList()
        if text_list:
            self.top_frame.bullet_button.setChecked(True)
        else:
            self.top_frame.bullet_button.setChecked(False)
        
    def fill_values(self, record):
        index = 1
        self.win.setWindowTitle('Sermon Prep Database - ' + str(record[0][17]) + ' - ' + str(record[0][3]))
        for i in range(self.scripture_frame_layout.count()):
            component = self.scripture_frame_layout.itemAt(i).widget()
            if isinstance(component, QLineEdit):
                if record[0][index]:
                    component.setText(str(record[0][index].replace('&quot', '"')).strip())
                else:
                    component.clear()
                index += 1
            elif isinstance(component, CustomTextEdit):
                component.clear()
                if record[0][index]:
                    component.setMarkdown(record[0][index].replace('&quot', '"').strip())
                    component.check_whole_text()
                index += 1
        for i in range(self.exegesis_frame_layout.count()):
            component = self.exegesis_frame_layout.itemAt(i).widget()
            if isinstance(component, QTextEdit) and not component.objectName() == 'text_box':
                if record[0][index]:
                    component.setMarkdown(record[0][index].replace('&quot', '"').strip())
                    component.check_whole_text()
                else:
                    component.clear()
                index += 1
        for i in range(self.outline_frame_layout.count()):
            component = self.outline_frame_layout.itemAt(i).widget()
            if isinstance(component, QTextEdit) and not component.objectName() == 'text_box':
                if record[0][index]:
                    component.setMarkdown(record[0][index].replace('&quot', '"').strip())
                    component.check_whole_text()
                else:
                    component.clear()
                index += 1
        for i in range(self.research_frame_layout.count()):
            component = self.research_frame_layout.itemAt(i).widget()
            if isinstance(component, QTextEdit):
                if record[0][index]:
                    text = self.spd.reformat_string_for_load(record[0][index])
                    component.setMarkdown(text.strip())
                    component.check_whole_text()
                else:
                    component.clear()
                index += 1
        for i in range(self.sermon_frame_layout.count()):
            component = self.sermon_frame_layout.itemAt(i).widget()
            if isinstance(component, QLineEdit):
                if record[0][index]:
                    component.setText(record[0][index].replace('&quot', '"'))
                else:
                    component.clear()
                index += 1

            if isinstance(component, QDateEdit):
                date = record[0][index]
                unusable_date = False
                if '/' in date:
                    date_split = date.split('/')
                elif '\\' in date:
                    date_split = date.split('\\')
                elif '-' in date:
                    date_split = date.split('-')
                else:
                    unusable_date = True
                if not unusable_date:
                    if int(date_split[0]) > 31:
                        component.setDate(QDate(int(date_split[0]), int(date_split[1]), int(date_split[2])))
                    elif date_split[2] > 31:
                        component.setDate(QDate(int(date_split[2]), int(date_split[0]), int(date_split[1])))
                else:
                    self.spd.write_to_log('unusable date in record #' + str(record[0][0]))
                    component.setDate(QDateTime.currentDateTime().date())
                index += 1

            if isinstance(component, CustomTextEdit):
                if record[0][index]:
                    component.setMarkdown(record[0][index].replace('&quot', '"').strip())
                    component.check_whole_text()
                else:
                    component.clear()
                index += 1

            if component.objectName() == 'text_box':
                component.clear()
                index += 1

        num_tabs = self.tabbed_frame.count()
        for i in range(1, num_tabs):
            frame = self.tabbed_frame.widget(i)
            widget = frame.findChild(QWidget, 'text_box')
            if widget:
                text_title = widget.findChild(QLabel, 'text_title')
                text_edit = widget.findChild(QTextEdit, 'text_edit')
                text_edit.setText(self.sermon_text_edit.toPlainText())
                text_title.setText(self.sermon_reference_field.text())

        self.top_frame.id_label.setText('ID: ' + str(record[0][0]))
        self.changes = False

    def changes_detected(self):
        self.changes = True

    def reference_changes(self):
        try:
            self.top_frame.references_cb.setItemText(self.top_frame.references_cb.currentIndex(), self.sermon_reference_field.text())
            self.win.setWindowTitle('Sermon Prep Database - ' + self.sermon_date_edit.text() + ' - ' + self.sermon_reference_field.text())

            num_tabs = self.tabbed_frame.count()
            for i in range(num_tabs):
                if i > 0:
                    frame = self.tabbed_frame.widget(i)
                    widget = frame.findChild(QWidget, 'text_box')
                    text_title = widget.findChild(QLabel, 'text_title')
                    text_title.setText(self.sermon_reference_field.text())

            if self.spd.auto_fill:
                if not self.gs: # only create one instance of GetScripture
                    self.gs = GetScripture(self.spd)

                if ':' in self.sermon_reference_field.text(): # only attempt to get the text if there's enough to work with
                    passage = self.gs.get_passage(self.sermon_reference_field.text())
                    if passage and not passage == -1:
                        self.sermon_text_edit.setText(passage)

            self.changes = True
        except Exception as ex:
            logging.exception(str(ex), True)

    def auto_fill(self):
        if self.auto_fill_checkbox.isChecked():
            self.spd.auto_fill = True
            self.spd.write_auto_fill_changes()
        else:
            self.spd.auto_fill = False
            self.spd.write_auto_fill_changes()

    def text_changes(self):
        num_tabs = self.tabbed_frame.count()
        for i in range(num_tabs):
            if i > 0:
                frame = self.tabbed_frame.widget(i)
                widget = frame.findChild(QWidget, 'text_box')
                text_edit = widget.findChild(QTextEdit, 'text_edit')
                text_edit.setText(self.sermon_text_edit.toPlainText())

    def date_changes(self):
        self.top_frame.dates_cb.setItemText(self.top_frame.dates_cb.currentIndex(), self.sermon_date_edit.text())
        self.win.setWindowTitle(
            'Sermon Prep Database - ' + self.sermon_date_edit.text() + ' - ' + self.sermon_reference_field.text())
        self.changes = True

    def do_exit(self):
        goon = True
        if self.changes:
            goon = self.spd.ask_save()
        if goon:
            sys.exit(0)

class CustomTextEdit(QTextEdit):
    def __init__(self, win, gui):
        super().__init__()
        self.setObjectName('custom_text_edit')
        self.win = win
        self.gui = gui
        self.textChanged.connect(self.gui.changes_detected)

    def keyReleaseEvent(self, evt):
        if evt.key() == Qt.Key_Space or evt.key() == Qt.Key_Return or evt.key() == Qt.Key_Enter:
            self.check_previous_word()

    def changeEvent(self, evt):
        self.gui.changes = True

    def check_previous_word(self):
        self.blockSignals(True)
        punctuations = [',', '.', '?', '!', ')', ';', ':', '-']

        cursor = self.textCursor()
        cursor.movePosition(QTextCursor.PreviousWord)
        cursor.select(cursor.WordUnderCursor)

        word = cursor.selection().toPlainText()
        for punctuation in punctuations:
            if word == punctuation:
                print('Word is punctuation')
                cursor.clearSelection()
                cursor.movePosition(QTextCursor.PreviousWord)
                cursor.movePosition(QTextCursor.PreviousWord)
                cursor.select(cursor.WordUnderCursor)
                word = cursor.selection().toPlainText()
                break

        print('Previous Word:', word)

        # if there's an apostrophe, check the next two characters for contraction letters
        if cursor.selection().toPlainText().endswith('\''):
            cursor.movePosition(QTextCursor.NextCharacter, cursor.KeepAnchor)
            if re.search('[a-z]$', cursor.selection().toPlainText()):
                word = cursor.selection().toPlainText()
                cursor.movePosition(QTextCursor.NextCharacter, cursor.KeepAnchor)
                if re.search('[a-z]$', cursor.selection().toPlainText()):
                    word = cursor.selection().toPlainText()

        cleaned_word = self.clean_word(word)

        suggestions = None
        if len(cleaned_word) > 0 and not any(c.isnumeric() for c in cleaned_word):
            if any(h.isalpha() for h in cleaned_word):
                suggestions = self.gui.spd.sym_spell.lookup(cleaned_word, Verbosity.CLOSEST, max_edit_distance=2,
                                                            include_unknown=True)

            if suggestions:
                char_format = cursor.charFormat()
                if not suggestions[0].term == cleaned_word:
                    char_format.setForeground(Qt.red)
                    cursor.mergeCharFormat(char_format)

                else:
                    if char_format.foreground() == Qt.red:
                        char_format.setForeground(Qt.black)
                        cursor.mergeCharFormat(char_format)

        cursor.clearSelection()
        cursor.movePosition(QTextCursor.NextWord)

        self.blockSignals(False)

    def check_whole_text(self):
        self.blockSignals(True)

        cursor = self.textCursor()
        cursor.movePosition(QTextCursor.Start)

        while not self.gui.spd.disable_spell_check:
            last_word = False
            cursor.select(cursor.WordUnderCursor)
            word = cursor.selection().toPlainText()

            #if the position doesn't change after moving to the next character, this is the last word
            pos_before_move = cursor.position()
            cursor.movePosition(QTextCursor.NextCharacter, cursor.KeepAnchor)
            pos_after_move = cursor.position()
            if pos_before_move == pos_after_move:
                last_word = True

            if cursor.selection().toPlainText().endswith('\''): # if there's an apostrophe, check the next two characters for contraction letters
                cursor.movePosition(QTextCursor.NextCharacter, cursor.KeepAnchor)
                if re.search('[a-z]$', cursor.selection().toPlainText()):
                    word = cursor.selection().toPlainText()
                    cursor.movePosition(QTextCursor.NextCharacter, cursor.KeepAnchor)
                    if re.search('[a-z]$', cursor.selection().toPlainText()):
                        word = cursor.selection().toPlainText()

            cursor.movePosition(QTextCursor.PreviousCharacter, cursor.KeepAnchor)

            cleaned_word = self.clean_word(word)

            suggestions = None
            if len(cleaned_word) > 0 and not any(c.isnumeric() for c in cleaned_word):
                if any(h.isalpha() for h in cleaned_word):
                    suggestions = self.gui.spd.sym_spell.lookup(cleaned_word, Verbosity.CLOSEST, max_edit_distance=2,
                                                            include_unknown=True)

                if suggestions:
                    char_format = cursor.charFormat()
                    if not suggestions[0].term == cleaned_word:
                        char_format.setForeground(Qt.red)
                        cursor.mergeCharFormat(char_format)

                    else:
                        if char_format.foreground() == Qt.red:
                            char_format.setForeground(Qt.black)
                            cursor.mergeCharFormat(char_format)

            cursor.clearSelection()
            cursor.movePosition(QTextCursor.NextWord)

            if last_word:
                break

        self.blockSignals(False)

    def contextMenuEvent(self, e):
        menu = self.createStandardContextMenu()

        clean_whitespace_action = QAction("Remove extra whitespace")
        clean_whitespace_action.triggered.connect(self.clean_whitespace)
        menu.insertAction(menu.actions()[0], clean_whitespace_action)
        menu.insertSeparator(menu.actions()[1])

        if not self.gui.spd.disable_spell_check:
            cursor = self.cursorForPosition(e.pos())
            cursor.select(QTextCursor.WordUnderCursor)
            word = cursor.selection().toPlainText()
            if len(word) > 0:
                cursor.movePosition(QTextCursor.NextCharacter, cursor.KeepAnchor)
                if cursor.selection().toPlainText().endswith('\''): # if there's an apostrophe, check the next two characters for contraction letters
                    cursor.movePosition(QTextCursor.NextCharacter, cursor.KeepAnchor)
                    if re.search('[a-z]$', cursor.selection().toPlainText()):
                        word = cursor.selection().toPlainText()
                        cursor.movePosition(QTextCursor.NextCharacter, cursor.KeepAnchor)
                        if re.search('[a-z]$', cursor.selection().toPlainText()):
                            word = cursor.selection().toPlainText()

                upper = False
                if word[0].isupper():
                    upper = True

                cleaned_word = self.clean_word(word)

                suggestions = self.gui.spd.sym_spell.lookup(cleaned_word, Verbosity.CLOSEST, max_edit_distance=2,
                                                            include_unknown=True)

                if not suggestions[0].term == cleaned_word:
                    spell_actions = {}

                    number_of_suggestions = len(suggestions)
                    if number_of_suggestions > 10: number_of_suggestions = 11

                    for i in range(number_of_suggestions):
                        term = suggestions[i].term
                        if upper:
                            term = term[0].upper() + term[1:]
                        spell_actions['action% s' % str(i)] = QAction(term)
                        spell_actions['action% s' % str(i)].setData((cursor, term))
                        spell_actions['action% s' % str(i)].triggered.connect(self.replace_word)
                        menu.insertAction(menu.actions()[i], spell_actions['action% s' % str(i)])

                    menu.insertSeparator(menu.actions()[i + 1])
                    action = QAction('Add to dictionary')
                    action.triggered.connect(lambda: self.gui.spd.add_to_dictionary(self, cleaned_word))
                    menu.insertAction(menu.actions()[i + 2], action)
                    menu.insertSeparator(menu.actions()[i + 3])

        menu.exec(e.globalPos())
        menu.close()

    def clean_word(self, word):
        chars = ['.', ',', ';', ':', '?', '!', '"', '...', '*', '-', '_',
                 '\n', '\u2026', '\u201c', '\u201d']
        single_quotes = ['\u2018', '\u2019']

        cleaned_word = word.lower().strip()

        for char in chars:
            cleaned_word = cleaned_word.replace(char, '')
        for single_quote in single_quotes:
            cleaned_word = cleaned_word.replace(single_quote, '\'')

        cleaned_word = cleaned_word.replace('\'s', '')
        cleaned_word = cleaned_word.replace('s\'', 's')
        cleaned_word = cleaned_word.replace("<[.?*]>", '')

        if cleaned_word.startswith('\''):
            cleaned_word = cleaned_word[1:len(cleaned_word)]
        if cleaned_word.endswith('\''):
            cleaned_word = cleaned_word[0:len(cleaned_word) - 1]

        # there's a chance that utf-8-sig artifacts will be attatched to the word
        # encoding to utf-8 then decoding as ascii removes them
        cleaned_word = cleaned_word.encode('utf-8').decode('ascii', errors='ignore')

        return cleaned_word

    def replace_word(self):
        sender = self.sender()
        cursor = sender.data()[0]
        term = sender.data()[1]

        self.setTextCursor(cursor)
        self.textCursor().removeSelectedText()
        cursor = self.textCursor()
        cursor.movePosition(QTextCursor.NextCharacter, QTextCursor.KeepAnchor)
        self.setTextCursor(cursor)

        chars = ['.', ',', ';', ':', '?', '!', '"', '*', '-', '_', '\n', ' ', '\u2026', '\u201c', '\u201d', '\u2018', '\u2019']
        add_space = True
        for char in chars:
            if self.textCursor().selectedText() == char:
                add_space = False

        cursor = self.textCursor()
        cursor.movePosition(QTextCursor.PreviousCharacter, QTextCursor.MoveAnchor)
        self.setTextCursor(cursor)

        if add_space:
            self.textCursor().insertText(term + ' ')
        else:
            self.textCursor().insertText(term)
        self.changes()

    def clean_whitespace(self):
        component = self.win.focusWidget()
        if isinstance(component, QTextEdit):
            string = component.toMarkdown()
            string = re.sub(' +', ' ', string)
            string = re.sub('\t+', '\t', string)

            component.setMarkdown(string)
        self.changes()

class Win(QMainWindow):
    from PyQt5.QtGui import QCloseEvent

    def __init__(self, gui):
        QMainWindow.__init__(self)
        self.gui = gui
        self.setWindowTitle('Sermon Prep Database')
        self.setStyleSheet('background-color: white')
        self.resize(1000, 800)
        self.move(50, 50)

    def closeEvent(self, event:QCloseEvent) -> None:
        event.ignore()
        self.gui.do_exit()

    def keyPressEvent(self, event:QKeyEvent):
        if (event.modifiers() & Qt.ControlModifier) and (event.modifiers() & Qt.ShiftModifier) and event.key() == Qt.Key_B:
            self.gui.top_frame.set_bullet()
            self.gui.top_frame.bullet_button.blockSignals(True)
            if self.gui.top_frame.bullet_button.isChecked():
                self.gui.top_frame.bullet_button.setChecked(False)
            else:
                self.gui.top_frame.bullet_button.setChecked(True)
            self.gui.top_frame.bold_button.blockSignals(False)
        elif event.modifiers() & Qt.ControlModifier and event.key() == Qt.Key_B:
            self.gui.top_frame.set_bold()
            self.gui.top_frame.bold_button.blockSignals(True)
            if self.gui.top_frame.bold_button.isChecked():
                self.gui.top_frame.bold_button.setChecked(False)
            else:
                self.gui.top_frame.bold_button.setChecked(True)
            self.gui.top_frame.bold_button.blockSignals(False)
        elif event.modifiers() & Qt.ControlModifier and event.key() == Qt.Key_I:
            self.gui.top_frame.set_italic()
            self.gui.top_frame.italic_button.blockSignals(True)
            if self.gui.top_frame.italic_button.isChecked():
                self.gui.top_frame.italic_button.setChecked(False)
            else:
                self.gui.top_frame.italic_button.setChecked(True)
            self.gui.top_frame.italic_button.blockSignals(False)
        elif event.modifiers() & Qt.ControlModifier and event.key() == Qt.Key_U:
            self.gui.top_frame.set_underline()
            self.gui.top_frame.underline_button.blockSignals(True)
            if self.gui.top_frame.underline_button.isChecked():
                self.gui.top_frame.underline_button.setChecked(False)
            else:
                self.gui.top_frame.underline_button.setChecked(True)
            self.gui.top_frame.underline_button.blockSignals(False)
        elif event.modifiers() & Qt.ControlModifier and event.key() == Qt.Key_S:
            self.gui.spd.save_rec()
        elif event.modifiers() & Qt.ControlModifier and event.key() == Qt.Key_P:
            self.gui.menu_bar.print_rec()
        elif event.modifiers() & Qt.ControlModifier and event.key() == Qt.Key_Q:
            self.gui.do_exit()
        event.accept()

class ScriptureBox(QWidget):
    def __init__(self, bgcolor):
        super().__init__()
        self.setObjectName('text_box')
        self.setMaximumWidth(300)
        text_layout = QVBoxLayout()
        self.setLayout(text_layout)
        text_title = QLabel()
        text_title.setObjectName('text_title')
        text_title.setStyleSheet('font-weight: bold; text-decoration: underline;')
        text_layout.addWidget(text_title)
        self.text_edit = QTextEdit()
        self.text_edit.setObjectName('text_edit')
        self.text_edit.setStyleSheet('background-color: ' + bgcolor + '; border: 0;')
        self.text_edit.setReadOnly(True)
        text_layout.addWidget(self.text_edit)
        self.hide()

    def clear(self):
        self.text_edit.clear()

class SearchBox(QWidget):
    def __init__(self, gui):
        self.gui = gui
        super().__init__()

    def show_results(self, result_list):
        results_widget_layout = QVBoxLayout()
        self.setLayout(results_widget_layout)
        self.setStyleSheet('''
                        QWidget {
                            background-color: ''' + self.gui.background_color + ''';}
                        QLabel {
                            font-family: "Helvetica";
                            font-size: 16px;
                            padding: 10px;
                            font-color: ''' + self.gui.accent_color + ''';}
                        QTableView {
                            background-color: white;
                            font-family: "Helvetica";
                            font-size: 16px;
                            padding: 3px;}
                        ''')

        results_header = QWidget()
        header_layout = QHBoxLayout()
        results_header.setLayout(header_layout)

        results_label = QLabel()
        header_layout.addWidget(results_label)

        close_button = QPushButton()
        close_button.setIcon(QIcon(self.gui.spd.cwd + 'resources/closeIcon.png'))
        close_button.setStyleSheet('background-color: ' + self.gui.accent_color)
        close_button.setToolTip('Close the search tab')
        close_button.pressed.connect(self.remove_self)
        header_layout.addStretch()
        header_layout.addWidget(close_button)

        results_widget_layout.addWidget(results_header)

        filtered_results = []
        for line in result_list:

            counter = 0
            for item in line[0]:
                counter += 1

            words_found = str(line[1])
            words_found = words_found.replace('[', '')
            words_found = words_found.replace(']', '')
            words_found = words_found.replace("\'", '')
            filtered_results.append((
                str(line[0][0]),
                str(line[2]),
                words_found,
                line[0][3],
                line[0][16],
                line[0][17],
                line[0][21][0:100] + '...'))

        model = QStandardItemModel(len(filtered_results), 5)
        for i in range(len(filtered_results)):
            for n in range(len(filtered_results[i])):
                item = QStandardItem(filtered_results[i][n])
                item.setEditable(False)
                model.setItem(i, n, item)
        model.setHeaderData(0, Qt.Horizontal, 'ID')
        model.setHeaderData(1, Qt.Horizontal, '# of\r\nMatches')
        model.setHeaderData(2, Qt.Horizontal, 'Word(s) Found')
        model.setHeaderData(3, Qt.Horizontal, 'Sermon Text')
        model.setHeaderData(4, Qt.Horizontal, 'Sermon Title')
        model.setHeaderData(5, Qt.Horizontal, 'Sermon Date')
        model.setHeaderData(6, Qt.Horizontal, 'Sermon Snippet')

        results_table_view = QTableView()
        results_table_view.setModel(model)
        results_table_view.setColumnWidth(0, 30)
        results_table_view.setColumnWidth(1, 60)
        results_table_view.setColumnWidth(2, 150)
        results_table_view.setColumnWidth(3, 200)
        results_table_view.setColumnWidth(4, 200)
        results_table_view.setColumnWidth(5, 100)
        results_table_view.setColumnWidth(6, 500)
        results_table_view.setShowGrid(False)
        results_table_view.setSelectionBehavior(QTableView.SelectRows)
        results_table_view.doubleClicked.connect(
            lambda: self.retrieve_selection(model, results_table_view.selectionModel().currentIndex().row()))
        results_widget_layout.addWidget(results_table_view)

        if len(filtered_results) == 1:
            results_label.setText(str(len(
                filtered_results)) + ' result found.\nDouble-click a result below to open it.')
        else:
            results_label.setText(str(len(
                filtered_results)) + ' results found.\nDouble-click a result below to open it.')

    def retrieve_selection(self, model, selection):
        goon = True
        if self.gui.changes:
            goon = self.gui.spd.ask_save()
        if goon:
            id_ = model.index(selection, 0).data()
            index = 0
            for item in self.gui.spd.ids:
                if int(item) == int(id_):
                    break
                index += 1
            self.gui.spd.get_by_index(index)
            self.gui.tabbed_frame.setCurrentWidget(self.gui.tabbed_frame.widget(0))

    def remove_self(self):
        self.gui.tabbed_frame.removeTab(5)
        self.gui.tabbed_frame.setCurrentWidget(self.gui.tabbed_frame.widget(0))
        self.destroy()
