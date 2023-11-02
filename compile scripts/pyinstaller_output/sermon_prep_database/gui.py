"""
@author Jeremy G. Wilson

Copyright 2023 Jeremy G. Wilson

This file is a part of the Sermon Prep Database program (v.4.0.8)

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
from PyQt5.QtGui import QIcon, QFont, QTextCursor, QStandardItemModel, QStandardItem, QPixmap, QColor, QPalette
from PyQt5.QtWidgets import QBoxLayout, QWidget, QUndoStack, QMessageBox, QTabWidget, QGridLayout, QLabel, QLineEdit, \
    QCheckBox, QDateEdit, QTextEdit, QAction, QMainWindow, QVBoxLayout, QHBoxLayout, QPushButton, QTableView
from symspellpy import Verbosity

from get_scripture import GetScripture
from menu_bar import MenuBar
from top_frame import TopFrame


class GUI(QObject):
    """
    GUI handles all the operations from the user interface. It builds the QT window and elements and requires the
    SermonPrepDatabase object in order to access its methods.

    :param SermonPrepDatabase spd: SermonPrepDatabase object
    """
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
        """
        Attach SermonPrepDatabase object to self, run a check on the database file, and initiate the QT GUI
        """
        super().__init__()
        self.spd = spd
        self.check_for_db()
        self.init_components()

    def init_components(self):
        """
        Builds the QT GUI, also using elements from menu_bar.py, top_frame.py, and print_dialog.py
        """
        self.spd.get_ids()
        self.spd.get_date_list()
        self.spd.get_scripture_list()
        self.spd.get_user_settings()
        self.spd.backup_db()

        self.accent_color = self.spd.user_settings[1]
        self.background_color = self.spd.user_settings[2]
        self.font_family = self.spd.user_settings[3]
        self.font_size = self.spd.user_settings[4]
        self.font_color = self.spd.user_settings[29]
        self.text_background = self.spd.user_settings[30]
        self.standard_font = QFont(self.font_family, int(self.font_size))
        self.bold_font = QFont(self.font_family, int(self.font_size), QFont.Bold)
        try:
            self.spd.line_spacing = str(self.spd.user_settings[28])
        except IndexError:
            self.spd.line_spacing = '1'

        self.win = Win(self)
        icon_pixmap = QPixmap(self.spd.cwd + 'resources/svg/spIcon.svg')
        self.win.setWindowIcon(QIcon(icon_pixmap))

        self.layout = QBoxLayout(QBoxLayout.TopToBottom)
        self.main_widget = QWidget()
        self.win.setCentralWidget(self.main_widget)
        self.main_widget.setLayout(self.layout)

        self.undo_stack = QUndoStack(self.main_widget)

        self.menu_bar = MenuBar(self.win, self, self.spd)
        self.top_frame = TopFrame(self.win, self, self.spd)
        self.top_frame.setObjectName('topFrame')
        self.layout.addWidget(self.top_frame)

        self.build_tabbed_frame()
        self.build_scripture_tab()
        self.build_exegesis_tab()
        self.build_outline_tab()
        self.build_research_tab()
        self.build_sermon_tab()

        self.set_style_sheets()

        self.win.showMaximized()

    def check_for_db(self):
        """
        Check if the database file exists. Prompt to import an existing database or create a new database if not. If
        it exists, but does not include the line_spacing column in the user_settings table, then it is an old version.
        """
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
                from convert_database import ConvertDatabase
                ConvertDatabase(self.spd)
            elif response == QMessageBox.No:
                # Create a new database in the user's App Data directory by copying the existing database template
                shutil.copy(self.spd.cwd + 'resources/database_template.db', self.spd.db_loc)
                QMessageBox.information(None, 'Database Created', 'A new database has been created.', QMessageBox.Ok)
                self.spd.app.processEvents()
            else:
                quit(0)
        else:
            # check that the existing database is in the new format
            result = self.spd.check_line_spacing()
            if result == -1:
                response = QMessageBox.question(
                    None,
                    'Old Database Found',
                    'It appears that you are upgrading from a previous version of Sermon Prep Database. Your database'
                    'file will need to be upgraded before you can continue. Upgrade now?',
                    QMessageBox.Yes | QMessageBox.No
                )

                if response == QMessageBox.Yes:
                    from convert_database import ConvertDatabase
                    ConvertDatabase(self.spd, 'existing')
                else:
                    quit(0)

        self.spd.write_to_log('checkForDB completed')
    
    def build_tabbed_frame(self):
        """
        Create a QTabWidget
        """
        self.tabbed_frame = QTabWidget()
        self.tabbed_frame.setObjectName('tabbedFrame')
        self.tabbed_frame.setTabPosition(QTabWidget.West)
        self.tabbed_frame.setIconSize(QSize(24, 24))
        self.layout.addWidget(self.tabbed_frame)
        
    def build_scripture_tab(self, insert=False):
        """
        Create a QWidget to hold the scripture tab's elements. Adds the elements.

        :param boolean insert: If other tabs already exist, insert tab at position 0.
        """
        self.scripture_frame = QWidget()
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
                QIcon(self.spd.cwd + 'resources/svg/spScriptureIcon.svg'),
                'Scripture'
            )
        else:
            self.tabbed_frame.addTab(
                self.scripture_frame,
                QIcon(self.spd.cwd + 'resources/svg/spScriptureIcon.svg'),
                'Scripture'
            )
        
    def build_exegesis_tab(self):
        """
        Create a QWidget to hold the exegesis tab's elements. Adds the elements.
        """
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
        
        self.tabbed_frame.addTab(self.exegesis_frame, QIcon(self.spd.cwd + 'resources/svg/spExegIcon.svg'), 'Exegesis')
        
    def build_outline_tab(self):
        """
        Create a QWidget to hold the outline tab's elements. Adds the elements.
        """
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

        self.tabbed_frame.addTab(self.outline_frame, QIcon(self.spd.cwd + 'resources/svg/spOutlineIcon.svg'), 'Outlines')
        
    def build_research_tab(self):
        """
        Create a QWidget to hold the research tab's elements. Adds the elements.
        """
        self.research_frame = QWidget()
        self.research_frame_layout = QGridLayout()
        self.research_frame.setLayout(self.research_frame_layout)
        self.research_frame_layout.setColumnMinimumWidth(1, 20)
        
        research_label = QLabel(self.spd.user_settings[19])
        self.research_frame_layout.addWidget(research_label, 0, 0)
        
        research_text = CustomTextEdit(self.win, self)
        research_text.setObjectName('custom_text_edit')
        research_text.cursorPositionChanged.connect(lambda: self.set_style_buttons(research_text))
        self.research_frame_layout.addWidget(research_text, 1, 0)

        scripture_box = ScriptureBox(self.background_color)
        self.research_frame_layout.addWidget(scripture_box, 0, 2, 2, 1)
        
        self.tabbed_frame.addTab(self.research_frame, QIcon(self.spd.cwd + 'resources/svg/spResearchIcon.svg'), 'Research')
        
    def build_sermon_tab(self):
        """
        Create a QWidget to hold the sermon tab's elements. Adds the elements.
        """
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
        
        self.tabbed_frame.addTab(self.sermon_frame, QIcon(self.spd.cwd + 'resources/svg/spSermonIcon.svg'), 'Sermon')

    def set_style_sheets(self):
        """
        Applies predetermined style sheets to self.tabbed_frame as well as each tab's QWidget. Also makes font changes
        to the TopFrame. Customizes the styles with the user's background color, accent color, font family, and font
        size.
        """
        current_changes_state = self.changes

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
                background-color: ''' + self.background_color + ''';
                color: ''' + self.font_color + ''';}
            QLabel {
                font-family: "''' + self.font_family + '''";
                font-size: ''' + str(self.font_size) + '''pt;
                color: ''' + self.font_color + ''';
                background-color: ''' + self.background_color + ''';}
            QLineEdit {
                font-family: "''' + self.font_family + '''";
                font-size: ''' + str(self.font_size) + '''pt;
                padding: 3px;
                border: 1px solid ''' + self.accent_color + ''';
                color: ''' + self.font_color + ''';
                background-color: ''' + self.text_background + ''';}
            QTextEdit {
                font-family: "''' + self.font_family + '''";
                font-size: ''' + str(self.font_size) + '''pt;
                padding: 3px;
                border: 1px solid ''' + self.accent_color + ''';
                color: ''' + self.font_color + ''';
                background-color: ''' + self.text_background + ''';}
            QDateEdit {
                font-family: "''' + self.font_family + '''";
                font-size: "''' + str(self.font_size) + '''";
                padding: 3px;
                border: 1px solid ''' + self.accent_color + ''';
                color: ''' + self.font_color + ''';
                background-color: ''' + self.text_background + ''';}
            ''')
        # anything larger than 12 is too big for the top frame
        if int(self.font_size) <= 12:
            size = self.font_size
        else:
            size = 12

        top_frame_style_sheet = ('''
            QLabel {
                font-family: "''' + self.font_family + '''";
                font-size: ''' + str(size) + '''pt;
                color: ''' + self.font_color + ''';}
            QLineEdit {
                font-family: "''' + self.font_family + '''";
                font-size: ''' + str(size) + '''pt;
                padding: 3px;
                border: 1px solid ''' + self.accent_color + ''';
                color: ''' + self.font_color + ''';
                background-color: ''' + self.text_background + ''';}
            QComboBox {
                font-family: "''' + self.font_family + '''";
                font-size: ''' + str(size) + '''pt;
                padding: 3px;
                border: 1px solid ''' + self.accent_color + ''';
                background-color: ''' + self.text_background + ''';
                color: ''' + self.font_color + ''';
                selection-background-color: ''' + self.background_color + ''';
                selection-color: ''' + self.font_color + ''';}
            QPushButton { 
                padding: 5;
                border: 0;}
            QPushButton:pressed {
                background-color: white;}
            QPushButton:hover {
                background-color: ''' + self.background_color + ''';}
            QPushButton:checked { 
                background-color: ''' + self.background_color + ''';}
            QPushButton#text_visible:checked {
                icon: url(''' + self.spd.cwd + '''resources/svg/spShowTextChecked.svg); 
                background-color: none;}
            QPushButton#text_visible:hover {
                background-color: ''' + self.background_color + ''';}
        ''')

        menu_style_sheet = '''
            QMenu { background: ''' + self.text_background + '''; } 
            QMenu:separator:hr { 
                background-color: ''' + self.text_background + ''';
                height: 0px;
                border-top: 1px solid ''' + self.accent_color + '''; 
                margin: 5px 
            } 
            QMenu::item:unselected { 
                background-color: ''' + self.text_background + ''';
                color: ''' + self.font_color + ''';
            } 
            QMenu::item:selected {
                background: ''' + self.background_color + ''';
                color: ''' + self.font_color + ''';
            }'''

        self.scripture_frame.setStyleSheet(standard_style_sheet)
        self.exegesis_frame.setStyleSheet(standard_style_sheet)
        self.outline_frame.setStyleSheet(standard_style_sheet)
        self.research_frame.setStyleSheet(standard_style_sheet)
        self.sermon_frame.setStyleSheet(standard_style_sheet)
        self.top_frame.setStyleSheet(top_frame_style_sheet)
        self.win.menuBar().setStyleSheet(menu_style_sheet)
        self.sermon_date_edit.setFont(QFont(self.font_family, int(size)))

        palette = self.main_widget.palette()
        palette.setColor(QPalette.ColorRole.Window, QColor(self.text_background))
        palette.setColor(QPalette.ColorRole.WindowText, QColor(self.font_color))
        self.main_widget.setPalette(palette)
        self.main_widget.setAutoFillBackground(True)

        palette = self.top_frame.palette()
        palette.setColor(QPalette.ColorRole.Window, QColor(self.text_background))
        palette.setColor(QPalette.ColorRole.WindowText, QColor(self.font_color))
        self.top_frame.setPalette(palette)
        self.top_frame.setAutoFillBackground(True)

        palette = self.win.menuBar().palette()
        palette.setColor(QPalette.ColorRole.Window, QColor(self.text_background))
        palette.setColor(QPalette.ColorRole.WindowText, QColor(self.font_color))
        self.win.menuBar().setPalette(palette)
        self.win.menuBar().setAutoFillBackground(True)

        for component in self.tabbed_frame.findChildren(CustomTextEdit, 'custom_text_edit'):
            component.document().setDefaultFont(QFont(self.font_family, int(self.font_size)))

        self.spd.app.processEvents()

        self.changes = current_changes_state

    def set_style_buttons(self, component):
        """
        Method for changing the GUI's style buttons (bold, italic, underline, bullets) based on where the cursor is
        located.

        :param QObject component: the QObject (CustomTextEdit) that is currently being used.
        """
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
        """
        Takes all the values from the currently accessed record and places them in their proper elements in the GUI.

        :param list of str record: a list whose first element is a list of values in their proper order.
        """
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
                    component.setMarkdown(self.spd.reformat_string_for_load(record[0][index]))
                    component.check_whole_text()
                index += 1
        for i in range(self.exegesis_frame_layout.count()):
            component = self.exegesis_frame_layout.itemAt(i).widget()
            if isinstance(component, CustomTextEdit) and not component.objectName() == 'text_box':
                component.clear()
                if record[0][index]:
                    component.setMarkdown(self.spd.reformat_string_for_load(record[0][index]))
                    component.check_whole_text()
                index += 1
        for i in range(self.outline_frame_layout.count()):
            component = self.outline_frame_layout.itemAt(i).widget()
            if isinstance(component, CustomTextEdit) and not component.objectName() == 'text_box':
                component.clear()
                if record[0][index]:
                    component.setMarkdown(self.spd.reformat_string_for_load(record[0][index]))
                    component.check_whole_text()
                index += 1
        for i in range(self.research_frame_layout.count()):
            component = self.research_frame_layout.itemAt(i).widget()
            if isinstance(component, CustomTextEdit):
                component.clear()
                if record[0][index]:
                    component.setMarkdown(self.spd.reformat_string_for_load(record[0][index]))
                    component.check_whole_text()
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

                # Check for common date delimiters. If they don't exist, don't use the date.
                # Switch this to Python's native date parsing in the future.
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
                component.clear()
                if record[0][index]:
                    component.setMarkdown(self.spd.reformat_string_for_load(record[0][index]))
                    component.check_whole_text()
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
        """
        Simply sets self.changes to True
        """
        self.changes = True

    def reference_changes(self):
        """
        When the sermon text reference TextEdit is changed, reflect those changes in the references combobox, the
        optional sermon text box on each tab, and on the MainWindow's title. If user has imported a bible and has auto
        fill turned on, use GetScripture to fill the sermon text TextEdit.
        """
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
        """
        Method to change the self.spd.auto_fill value based on user input then save that change to the database.
        """
        if self.auto_fill_checkbox.isChecked():
            self.spd.auto_fill = True
            self.spd.write_auto_fill_changes()
        else:
            self.spd.auto_fill = False
            self.spd.write_auto_fill_changes()

    def text_changes(self):
        """
        Method to fill the optional sermon text box on each tab with scripture when the Scripture Text text is
        changed
        """
        num_tabs = self.tabbed_frame.count()
        for i in range(num_tabs):
            if i > 0:
                frame = self.tabbed_frame.widget(i)
                widget = frame.findChild(QWidget, 'text_box')
                text_edit = widget.findChild(QTextEdit, 'text_edit')
                text_edit.setText(self.sermon_text_edit.toPlainText())

    def date_changes(self):
        """
        Method to change the dates combobox and window title to reflect changes made to the sermon date.
        """
        self.top_frame.dates_cb.setItemText(self.top_frame.dates_cb.currentIndex(), self.sermon_date_edit.text())
        self.win.setWindowTitle(
            'Sermon Prep Database - ' + self.sermon_date_edit.text() + ' - ' + self.sermon_reference_field.text())
        self.changes = True

    def do_exit(self):
        """
        Method to ask if changes are to be saved before exiting the program.
        """
        goon = True
        if self.changes:
            goon = self.spd.ask_save()
        if goon:
            sys.exit(0)

class CustomTextEdit(QTextEdit):
    """
    CustomTextEdit is an implementation of QTextEdit that adds spell-checking capabilities. Two different types of
    spell checking are done depending on user input and current spell-check state: check previous word, or check
    whole text. Also modifies the standard context menu to include spell check suggestions and an option to remove
    extra white space from the text.

    :param QMainWindow win: GUI's Main Window
    :param GUI gui: The GUI object
    """
    def __init__(self, win, gui):
        super().__init__()
        self.setObjectName('custom_text_edit')
        self.win = win
        self.gui = gui
        self.textChanged.connect(self.gui.changes_detected)
        self.spelling_errors_present = False

    def keyReleaseEvent(self, evt):
        """
        @override
        During normal word completion events like space or enter, check the previous word's spelling. If user is
        navigating using arrow keys, check the whole text to see if user has changed previously misspelled words.
        """
        arrow_keys = [Qt.Key_Left, Qt.Key_Right, Qt.Key_Up, Qt.Key_Down]
        if not self.gui.spd.disable_spell_check:
            if evt.key() == Qt.Key_Space or evt.key() == Qt.Key_Return or evt.key() == Qt.Key_Enter:
                self.check_previous_word()

            if self.spelling_errors_present and evt.key() in arrow_keys:
                self.check_whole_text()

    def mouseReleaseEvent(self, evt):
        """
        @override
        If user is navigating using mouse clicks, check the whole text to see if user has changed previously misspelled
        words.
        """
        if evt.button() == 1 and self.spelling_errors_present and not self.gui.spd.disable_spell_check:
            self.check_whole_text()

    def changeEvent(self, evt):
        """
        @override
        Set self.gui.changes to true.
        """
        self.gui.changes = True

    def check_previous_word(self):
        """
        Method to spell-check the word previous to the user's cursor. Skips over punctuation if that is the first
        previous "word" found.
        """
        self.blockSignals(True) # we don't want changeevent fired while manipulating for spell check
        punctuations = [',', '.', '?', '!', ')', ';', ':', '-']

        cursor = self.textCursor()
        cursor.movePosition(QTextCursor.PreviousWord)
        cursor.select(cursor.WordUnderCursor)

        word = cursor.selection().toPlainText()
        for punctuation in punctuations:
            if word == punctuation:
                cursor.clearSelection()
                cursor.movePosition(QTextCursor.PreviousWord)
                cursor.movePosition(QTextCursor.PreviousWord)
                cursor.select(cursor.WordUnderCursor)
                word = cursor.selection().toPlainText()
                break

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
                # if the first suggestion is the same as the word, then it's not spelled wrong
                if not suggestions[0].term == cleaned_word:
                    char_format.setForeground(Qt.red)
                    cursor.mergeCharFormat(char_format)
                    self.spelling_errors_present = True
                else:
                    char_format.setForeground(Qt.black)
                    cursor.mergeCharFormat(char_format)

        cursor.clearSelection()
        cursor.movePosition(QTextCursor.NextWord)

        char_format = cursor.charFormat()
        char_format.setForeground(Qt.black)
        cursor.mergeCharFormat(char_format)
        self.setTextCursor(cursor)

        self.blockSignals(False)

    def check_whole_text(self):
        """
        Method to check every word in this QTextEdit for spelling errors.
        """
        self.blockSignals(True)
        self.spelling_errors_present = False

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
                        self.spelling_errors_present = True

                    else:
                        if char_format.foreground() == Qt.red:
                            char_format.setForeground(Qt.black)
                            cursor.mergeCharFormat(char_format)

            cursor.clearSelection()
            cursor.movePosition(QTextCursor.NextWord)

            if last_word:
                break

        self.blockSignals(False)

    def contextMenuEvent(self, evt):
        """
        @override
        Alters the standard context menu of this QTextEdit to include a list of spell-check words as well as an option
        to remove extraneous whitespace from the text.
        """
        menu = self.createStandardContextMenu()
        menu.setStyleSheet('QMenu::item:selected { background: white; color: black; }')

        clean_whitespace_action = QAction("Remove extra whitespace")
        clean_whitespace_action.triggered.connect(self.clean_whitespace)
        menu.insertAction(menu.actions()[0], clean_whitespace_action)
        menu.insertSeparator(menu.actions()[1])

        if not self.gui.spd.disable_spell_check:
            cursor = self.cursorForPosition(evt.pos())
            cursor.select(QTextCursor.WordUnderCursor)
            word = cursor.selection().toPlainText()
            if len(word) > 0:
                cursor.movePosition(QTextCursor.NextCharacter, cursor.KeepAnchor)

                # if there's an apostrophe, check the next two characters for contraction letters
                if cursor.selection().toPlainText().endswith('\''):
                    cursor.movePosition(QTextCursor.NextCharacter, cursor.KeepAnchor)
                    if re.search('[a-z]$', cursor.selection().toPlainText()):
                        word = cursor.selection().toPlainText()
                        cursor.movePosition(QTextCursor.NextCharacter, cursor.KeepAnchor)
                        if re.search('[a-z]$', cursor.selection().toPlainText()):
                            word = cursor.selection().toPlainText()

                # Check to see if the original word is capitalized so that the replacement can also be capitalized
                upper = False
                if word[0].isupper():
                    upper = True

                cleaned_word = self.clean_word(word)
                suggestions = self.gui.spd.sym_spell.lookup(cleaned_word, Verbosity.CLOSEST, max_edit_distance=2,
                                                            include_unknown=True)

                next_menu_index = 1
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

                        next_menu_index = i + 1

                menu.insertSeparator(menu.actions()[next_menu_index])
                action = QAction('Add to dictionary')
                action.triggered.connect(lambda: self.gui.spd.add_to_dictionary(self, cleaned_word))
                menu.insertAction(menu.actions()[next_menu_index + 2], action)
                menu.insertSeparator(menu.actions()[next_menu_index + 3])

        menu.exec(evt.globalPos())
        menu.close()

    def clean_word(self, word):
        """
        Method to strip any non-word characters as well as pluralizing apostrophes out of a word to be spell-checked.

        :param str word: Word to be cleaned.
        """
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

        # there's a chance that utf-8-sig artifacts will be attached to the word
        # encoding to utf-8 then decoding as ascii removes them
        cleaned_word = cleaned_word.encode('utf-8').decode('ascii', errors='ignore')

        return cleaned_word

    def replace_word(self):
        """
        Method to replace a misspelled word if the user chooses a replacement from the context menu.
        """
        sender = self.sender()
        cursor = sender.data()[0]
        term = sender.data()[1]

        self.setTextCursor(cursor)
        selected_word = self.textCursor().selectedText().strip()

        # check if the original word includes punctuation so that it can be preserved
        punctuation = ''
        if not selected_word[len(selected_word) - 1].isalpha():
            punctuation = selected_word[len(selected_word) - 1]

        self.textCursor().removeSelectedText()
        cursor = self.textCursor()
        cursor.movePosition(QTextCursor.NextCharacter, QTextCursor.KeepAnchor)
        self.setTextCursor(cursor)

        cursor = self.textCursor()
        cursor.movePosition(QTextCursor.PreviousCharacter, QTextCursor.MoveAnchor)
        self.setTextCursor(cursor)

        if len(punctuation) > 0:
            self.textCursor().insertText(term + punctuation)
        else:
            self.textCursor().insertText(term + punctuation + ' ')

        self.gui.changes = True

    def clean_whitespace(self):
        """
        Method to remove duplicate spaces and tabs from the document.
        """
        component = self.win.focusWidget()
        if isinstance(component, QTextEdit):
            string = component.toMarkdown()
            string = re.sub(' +', ' ', string)
            string = re.sub('\t+', '\t', string)

            component.setMarkdown(string)
        self.gui.changes = True

class Win(QMainWindow):
    """
    Win customizes QMainWindow in order to differently handle the closeEvent as well as add keyboard shortcut
    functionality for the user.
    """
    from PyQt5.QtGui import QCloseEvent

    def __init__(self, gui):
        QMainWindow.__init__(self)
        self.gui = gui
        self.setWindowTitle('Sermon Prep Database')
        self.resize(1000, 800)
        self.move(50, 50)

    def closeEvent(self, event:QCloseEvent) -> None:
        """
        @override
        Ignore the closeEvent and run GUI's do_exit method instead.
        """
        event.ignore()
        self.gui.do_exit()

    def keyPressEvent(self, event):
        """
        Add keyboard shortcuts for common user tasks:
            Ctrl-Shift-B: Turn on bullets
            Ctrl-B: Bold
            Ctrl-I: Italic
            Ctrl-U: Underline
            Ctrl-S: Save
            Ctrl-P: Print
            Ctrl-Q: Exit
        """
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
    """
    Creates an independent QWidget that can be added or removed from layouts based on user's input.

    :param str bgcolor: User's chosen background color
    """
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
    """
    Creates an independent QWidget to be added to the main tabbed widget when the user performs a search.

    :param GUI gui: The GUI object
    """
    def __init__(self, gui):
        self.gui = gui
        super().__init__()

    def show_results(self, result_list):
        """
        Method to build the results widget.

        :param list of str result_list: The list containing each list of results from the search.
        """
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
                            padding: 3px;
                        }
                        ''')

        results_header = QWidget()
        header_layout = QHBoxLayout()
        results_header.setLayout(header_layout)

        results_label = QLabel()
        header_layout.addWidget(results_label)

        close_button = QPushButton()
        close_button.setIcon(QIcon(self.gui.spd.cwd + 'resources/svg/spCloseIcon.svg'))
        close_button.setStyleSheet('background-color: ' + self.gui.accent_color)
        close_button.setToolTip('Close the search tab')
        close_button.pressed.connect(self.remove_self)
        header_layout.addStretch()
        header_layout.addWidget(close_button)

        results_widget_layout.addWidget(results_header)

        # count the number of times the search term(s) was/were found in the record so that they can be sorted
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
        """
        Method to pull up whichever record the user selects

        :param QStandardItemModel model: The model applied to the results_table_view
        :param int selection: The row number of the results_table_view that was double-clicked
        """
        # be sure to check for changes before pulling up the new record
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
        """
        Method to remove this widget's tab from the GUI's tabbed widget.
        """
        self.gui.tabbed_frame.removeTab(5)
        self.gui.tabbed_frame.setCurrentWidget(self.gui.tabbed_frame.widget(0))
        self.destroy()
