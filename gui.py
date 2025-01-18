import re
from os.path import exists

from PyQt6.QtCore import Qt, QSize, QDate, QDateTime, QObject, QRunnable, pyqtSignal, QSizeF
from PyQt6.QtGui import QIcon, QFont, QStandardItemModel, QStandardItem, QPixmap, QColor, \
    QCloseEvent, QAction, QUndoStack, QTextCursor, QPainter, QTextDocument, QTextOption, QTextCharFormat
from PyQt6.QtWidgets import QWidget, QTabWidget, QGridLayout, QLabel, QLineEdit, \
    QCheckBox, QDateEdit, QTextEdit, QMainWindow, QVBoxLayout, QHBoxLayout, QPushButton, QTableView, QDialog, \
    QApplication, QProgressBar, QTabBar

from get_scripture import GetScripture
from menu_bar import MenuBar
from toolbar import Toolbar
from runnables import SpellCheck, LoadDictionary


class GUI(QMainWindow):
    spd = None
    undo_stack = None
    changes = False
    gs = None
    create_main_gui = pyqtSignal()
    clear_changes_signal = pyqtSignal()
    set_text_cursor_signal = pyqtSignal(QWidget, QTextCursor)
    set_text_color_signal = pyqtSignal(QTextEdit, list, QColor)
    reset_cursor_color_signal = pyqtSignal(QTextEdit)
    spell_check = None
    
    def __init__(self):
        """
        GUI handles all the operations from the user interface. It builds the QT window and elements and requires the
        SermonPrepDatabase object in order to access its methods.
        """

        super().__init__()
        self.create_main_gui.connect(self.create_gui)
        self.clear_changes_signal.connect(self.clear_changes)
        self.set_text_cursor_signal.connect(self.set_text_cursor)
        self.set_text_color_signal.connect(self.set_text_color)
        self.reset_cursor_color_signal.connect(self.reset_cursor_color)

        self.standard_font = None
        self.bold_font = None
        self.main_widget = None
        self.layout = None
        self.menu_bar = None
        self.toolbar = None
        self.tabbed_frame = None
        self.scripture_frame = None
        self.scripture_frame_layout = None
        self.sermon_reference_field = None
        self.auto_fill_checkbox = None
        self.sermon_text_edit = None
        self.exegesis_frame = None
        self.exegesis_frame_layout = None
        self.outline_frame = None
        self.outline_frame_layout = None
        self.research_frame = None
        self.research_frame_layout = None
        self.sermon_frame = None
        self.sermon_frame_layout = None
        self.sermon_date_edit = None

        self.light_tab_icons = [
            QIcon('resources/svg/spScriptureIcon.svg'),
            QIcon('resources/svg/spExegIcon.svg'),
            QIcon('resources/svg/spOutlineIcon.svg'),
            QIcon('resources/svg/spResearchIcon.svg'),
            QIcon('resources/svg/spSermonIcon.svg')
        ]
        self.dark_tab_icons = [
            QIcon('resources/svg/spScriptureIconDark.svg'),
            QIcon('resources/svg/spExegIconDark.svg'),
            QIcon('resources/svg/spOutlineIconDark.svg'),
            QIcon('resources/svg/spResearchIconDark.svg'),
            QIcon('resources/svg/spSermonIconDark.svg')
        ]

        from main import SermonPrepDatabase
        self.spd = SermonPrepDatabase(self)

        initial_startup = InitialStartup(self)
        self.spd.spell_check_thread_pool.start(initial_startup)

    def create_gui(self):
        """
        Builds the QT GUI, also using elements from menu_bar.py, toolbar.py, and print_dialog.py
        """
        self.setWindowTitle('Sermon Prep Database')

        self.standard_font = QFont(self.spd.user_settings['font_family'], int(self.spd.user_settings['font_size']))
        self.bold_font = QFont(
            self.spd.user_settings['font_family'], int(self.spd.user_settings['font_size']), QFont.Weight.Bold)
        self.spd.line_spacing = str(self.spd.user_settings['line_spacing'])

        icon_pixmap = QPixmap('resources/svg/spIcon.svg')
        self.setWindowIcon(QIcon(icon_pixmap))

        self.main_widget = QWidget()
        self.setCentralWidget(self.main_widget)
        self.layout = QVBoxLayout(self.main_widget)

        self.undo_stack = QUndoStack(self.main_widget)

        self.menu_bar = MenuBar(self, self.spd)
        self.toolbar = Toolbar(self, self.spd)
        self.layout.addWidget(self.toolbar)

        self.build_tabbed_frame()
        self.build_scripture_tab()
        self.build_exegesis_tab()
        self.build_outline_tab()
        self.build_research_tab()
        self.build_sermon_tab()

        self.menu_bar.color_change(self.spd.user_settings['theme'])

        self.showMaximized()

        self.spd.current_rec_index = len(self.spd.ids) - 1
        self.spd.get_by_index(self.spd.current_rec_index)
        self.apply_font(self.spd.user_settings['font_family'], self.spd.user_settings['font_size'])
        self.apply_line_spacing()
    
    def build_tabbed_frame(self):
        """
        Create a QTabWidget
        """
        tab_container = QWidget()
        tab_container.setObjectName('tab_container')
        tab_container.setAutoFillBackground(True)
        tab_container_layout = QVBoxLayout(tab_container)
        tab_container_layout.setContentsMargins(0, 10, 0, 0)
        self.layout.addWidget(tab_container)

        self.tabbed_frame = QTabWidget()
        self.tabbed_frame.setAutoFillBackground(True)
        self.tabbed_frame.setIconSize(QSize(24, 24))
        self.tabbed_frame.tabBar().currentChanged.connect(self.current_tab_changed)
        tab_container_layout.addWidget(self.tabbed_frame)

    def current_tab_changed(self):
        if self.spd.user_settings['theme'] == 'dark':
            return
        index = self.sender().currentIndex()
        for i in range(5):
            if i == index:
                self.tabbed_frame.setTabIcon(i, self.dark_tab_icons[i])
            else:
                self.tabbed_frame.setTabIcon(i, self.light_tab_icons[i])
        
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
        
        pericope_label = QLabel(self.spd.user_settings['label1'])
        self.scripture_frame_layout.addWidget(pericope_label, 0, 0)
        
        pericope_field = QLineEdit()
        pericope_field.textEdited.connect(self.changes_detected)
        self.scripture_frame_layout.addWidget(pericope_field, 1, 0)
        
        pericope_text_label = QLabel(self.spd.user_settings['label2'])
        self.scripture_frame_layout.addWidget(pericope_text_label, 2, 0)

        pericope_text_edit = CustomTextEdit(self)
        pericope_text_edit.cursorPositionChanged.connect(lambda: self.set_style_buttons(pericope_text_edit))
        self.scripture_frame_layout.addWidget(pericope_text_edit, 3, 0)
        
        sermon_reference_label = QLabel(self.spd.user_settings['label3'])
        self.scripture_frame_layout.addWidget(sermon_reference_label, 0, 1)

        self.sermon_reference_field = QLineEdit()
        self.sermon_reference_field.textEdited.connect(self.reference_changes)

        if exists(self.spd.app_dir + '/my_bible.xml'):
            self.scripture_frame_layout.addWidget(self.sermon_reference_field, 1, 1)

            self.auto_fill_checkbox = QCheckBox('Auto-fill ' + self.spd.user_settings['label4'])
            self.auto_fill_checkbox.setChecked(True)
            self.scripture_frame_layout.addWidget(self.auto_fill_checkbox, 1, 2)
            if self.spd.user_settings['auto_fill']:
                self.auto_fill_checkbox.setChecked(True)
            else:
                self.auto_fill_checkbox.setChecked(False)
            self.auto_fill_checkbox.stateChanged.connect(self.auto_fill)
        else:
            self.scripture_frame_layout.addWidget(self.sermon_reference_field, 1, 1, 1, 2)
        
        sermon_text_label = QLabel(self.spd.user_settings['label4'])
        self.scripture_frame_layout.addWidget(sermon_text_label, 2, 1, 1, 2)
        
        self.sermon_text_edit = CustomTextEdit(self)
        self.sermon_text_edit.cursorPositionChanged.connect(lambda: self.set_style_buttons(self.sermon_text_edit))
        self.scripture_frame_layout.addWidget(self.sermon_text_edit, 3, 1, 1, 2)

        self.sermon_reference_field.textChanged.connect(self.reference_changes)
        self.sermon_reference_field.textEdited.connect(self.reference_changes)

        if insert:
            self.tabbed_frame.insertTab(
                0,
                self.scripture_frame,
                self.light_tab_icons[0],
                'Scripture'
            )
        else:
            self.tabbed_frame.addTab(
                self.scripture_frame,
                self.light_tab_icons[0],
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
        
        fcft_label = QLabel(self.spd.user_settings['label5'])
        self.exegesis_frame_layout.addWidget(fcft_label, 0, 0)
        
        fcft_text = CustomTextEdit(self)
        fcft_text.cursorPositionChanged.connect(lambda: self.set_style_buttons(fcft_text))
        self.exegesis_frame_layout.addWidget(fcft_text, 1, 0)
        
        gat_label = QLabel(self.spd.user_settings['label6'])
        self.exegesis_frame_layout.addWidget(gat_label, 3, 0)
        
        gat_text = CustomTextEdit(self)
        gat_text.cursorPositionChanged.connect(lambda: self.set_style_buttons(gat_text))
        self.exegesis_frame_layout.addWidget(gat_text, 4, 0)
        
        cpt_label = QLabel(self.spd.user_settings['label7'])
        self.exegesis_frame_layout.addWidget(cpt_label, 6, 0)
        
        cpt_text = CustomTextEdit(self)
        cpt_text.cursorPositionChanged.connect(lambda: self.set_style_buttons(cpt_text))
        self.exegesis_frame_layout.addWidget(cpt_text, 7, 0)
        
        pb_label = QLabel(self.spd.user_settings['label8'])
        self.exegesis_frame_layout.addWidget(pb_label, 3, 2)
        
        pb_text = CustomTextEdit(self)
        pb_text.cursorPositionChanged.connect(lambda: self.set_style_buttons(pb_text))
        self.exegesis_frame_layout.addWidget(pb_text, 4, 2)
        
        fcfs_label = QLabel(self.spd.user_settings['label9'])
        self.exegesis_frame_layout.addWidget(fcfs_label, 0, 4)
        
        fcfs_text = CustomTextEdit(self)
        fcfs_text.cursorPositionChanged.connect(lambda: self.set_style_buttons(fcfs_text))
        self.exegesis_frame_layout.addWidget(fcfs_text, 1, 4)
        
        gas_label = QLabel(self.spd.user_settings['label10'])
        self.exegesis_frame_layout.addWidget(gas_label, 3, 4)
        
        gas_text = CustomTextEdit(self)
        gas_text.cursorPositionChanged.connect(lambda: self.set_style_buttons(gas_text))
        self.exegesis_frame_layout.addWidget(gas_text, 4, 4)
        
        cps_label = QLabel(self.spd.user_settings['label11'])
        self.exegesis_frame_layout.addWidget(cps_label, 6, 4)
        
        cps_text = CustomTextEdit(self)
        cps_text.cursorPositionChanged.connect(lambda: self.set_style_buttons(cps_text))
        self.exegesis_frame_layout.addWidget(cps_text, 7, 4)

        scripture_box = ScriptureBox()
        self.exegesis_frame_layout.addWidget(scripture_box, 0, 6, 8, 1)
        
        self.tabbed_frame.addTab(self.exegesis_frame, self.light_tab_icons[1], 'Exegesis')

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
        
        scripture_outline_label = QLabel(self.spd.user_settings['label12'])
        self.outline_frame_layout.addWidget(scripture_outline_label, 0, 0)
        
        scripture_outline_text = CustomTextEdit(self)
        scripture_outline_text.cursorPositionChanged.connect(lambda: self.set_style_buttons(scripture_outline_text))
        self.outline_frame_layout.addWidget(scripture_outline_text, 1, 0)
        
        sermon_outline_label = QLabel(self.spd.user_settings['label13'])
        self.outline_frame_layout.addWidget(sermon_outline_label, 0, 2)
        
        sermon_outline_text = CustomTextEdit(self)
        sermon_outline_text.cursorPositionChanged.connect(lambda: self.set_style_buttons(sermon_outline_text))
        self.outline_frame_layout.addWidget(sermon_outline_text, 1, 2)
        
        illustration_label = QLabel(self.spd.user_settings['label14'])
        self.outline_frame_layout.addWidget(illustration_label, 0, 4)
        
        illustration_text = CustomTextEdit(self)
        illustration_text.cursorPositionChanged.connect(lambda: self.set_style_buttons(illustration_text))
        self.outline_frame_layout.addWidget(illustration_text, 1, 4)

        scripture_box = ScriptureBox()
        self.outline_frame_layout.addWidget(scripture_box, 0, 6, 5, 1)

        self.tabbed_frame.addTab(self.outline_frame, self.light_tab_icons[2], 'Outlines')
        
    def build_research_tab(self):
        """
        Create a QWidget to hold the research tab's elements. Adds the elements.
        """
        self.research_frame = QWidget()
        self.research_frame_layout = QGridLayout()
        self.research_frame.setLayout(self.research_frame_layout)
        self.research_frame_layout.setColumnMinimumWidth(1, 20)
        
        research_label = QLabel(self.spd.user_settings['label15'])
        self.research_frame_layout.addWidget(research_label, 0, 0)
        
        research_text = CustomTextEdit(self)
        research_text.setObjectName('custom_text_edit')
        research_text.cursorPositionChanged.connect(lambda: self.set_style_buttons(research_text))
        self.research_frame_layout.addWidget(research_text, 1, 0)

        scripture_box = ScriptureBox()
        self.research_frame_layout.addWidget(scripture_box, 0, 2, 2, 1)
        
        self.tabbed_frame.addTab(self.research_frame, self.light_tab_icons[3], 'Research')
        
    def build_sermon_tab(self):
        """
        Create a QWidget to hold the sermon tab's elements. Adds the elements.
        """
        self.sermon_frame = QWidget()
        self.sermon_frame_layout = QGridLayout()
        self.sermon_frame.setLayout(self.sermon_frame_layout)
        self.sermon_frame_layout.setColumnStretch(0, 5)
        self.sermon_frame_layout.setColumnStretch(1, 2)
        self.sermon_frame_layout.setColumnStretch(2, 5)
        self.sermon_frame_layout.setColumnStretch(3, 2)
        
        sermon_title_label = QLabel(self.spd.user_settings['label16'])
        self.sermon_frame_layout.addWidget(sermon_title_label, 0, 0)
        
        sermon_title_field = QLineEdit()
        sermon_title_field.textEdited.connect(self.changes_detected)
        self.sermon_frame_layout.addWidget(sermon_title_field, 1, 0)
        
        sermon_date_label = QLabel(self.spd.user_settings['label17'])
        self.sermon_frame_layout.addWidget(sermon_date_label, 0, 1)

        self.sermon_date_edit = QDateEdit()
        self.sermon_date_edit.setCalendarPopup(True)
        self.sermon_date_edit.setMinimumHeight(30)
        self.sermon_date_edit.dateChanged.connect(self.date_changes)
        self.sermon_frame_layout.addWidget(self.sermon_date_edit, 1, 1)
        
        sermon_location_label = QLabel(self.spd.user_settings['label18'])
        self.sermon_frame_layout.addWidget(sermon_location_label, 0, 2)
        
        sermon_location_field = QLineEdit()
        sermon_location_field.textEdited.connect(self.changes_detected)
        self.sermon_frame_layout.addWidget(sermon_location_field, 1, 2)
        
        ctw_label = QLabel(self.spd.user_settings['label19'])
        self.sermon_frame_layout.addWidget(ctw_label, 2, 0)
        
        ctw_field = QLineEdit()
        ctw_field.textEdited.connect(self.changes_detected)
        self.sermon_frame_layout.addWidget(ctw_field, 3, 0)
        
        hr_label = QLabel(self.spd.user_settings['label20'])
        self.sermon_frame_layout.addWidget(hr_label, 2, 2)
        
        hr_field = QLineEdit()
        hr_field.textEdited.connect(self.changes_detected)
        self.sermon_frame_layout.addWidget(hr_field, 3, 2)

        sermon_text = CustomTextEdit(self)
        sermon_text.cursorPositionChanged.connect(lambda: self.set_style_buttons(sermon_text))
        self.sermon_frame_layout.addWidget(sermon_text, 5, 0, 1, 4)

        self.sermon_view_button = QPushButton()
        if self.spd.user_settings['theme'] == 'dark':
            self.sermon_view_button.setIcon(QIcon('resources/svg/spSermonViewIconLight.svg'))
        else:
            self.sermon_view_button.setIcon(QIcon('resources/svg/spSermonViewIconDark.svg'))
        self.sermon_view_button.setIconSize(QSize(36, 36))
        self.sermon_view_button.setFixedSize(QSize(56, 56))
        self.sermon_view_button.setToolTip('Open a window that shows your sermon')
        self.sermon_view_button.released.connect(lambda: SermonView(self, sermon_text.toSimplifiedHtml()))
        self.sermon_frame_layout.addWidget(self.sermon_view_button, 0, 3, 4, 1)
        
        sermon_label = QLabel(self.spd.user_settings['label21'])
        self.sermon_frame_layout.addWidget(sermon_label, 4, 0)

        scripture_box = ScriptureBox()
        self.sermon_frame_layout.addWidget(scripture_box, 0, 5, 6, 1)
        
        self.tabbed_frame.addTab(self.sermon_frame, self.light_tab_icons[4], 'Sermon')

    def clear_changes(self):
        self.changes = False

    def apply_font(self, font_family, font_size, font_chooser=None, save_config=False):
        """
        Method to apply the user's font changes to the GUI

        :param str font_family: The family name of the font the user chose.
        :param int font_size: The font size the user chose.
        :param QWidget font_chooser: The widget created in the change_font method.
        :param boolean save_config: True if this was called when the user clicked 'OK', so close fontChooser.
        """
        current_changes_status = self.changes

        self.spd.user_settings['font_family'] = font_family
        self.spd.user_settings['font_size'] = font_size
        self.standard_font = QFont(font_family, int(font_size))
        self.bold_font = QFont(font_family, int(font_size), QFont.Weight.Bold)
        font = QFont(self.spd.user_settings['font_family'], int(self.spd.user_settings['font_size']))
        for label in self.findChildren(QLabel):
            label.setFont(font)
        for line_edit in self.findChildren(QLineEdit):
            line_edit.setFont(font)
        for text_edit in self.findChildren(CustomTextEdit):
            html = text_edit.toSimplifiedHtml()
            text_edit.document().setDefaultStyleSheet(
                'p, li {'
                'font-family: ' + self.spd.user_settings['font_family'] + ';'
                'font-size: ' + self.spd.user_settings['font_size'] + 'pt;'
                'line-height: ' + self.spd.user_settings['line_spacing'] + ';'
                '}'
            )
            text_edit.setHtml(html)

            cursor = text_edit.textCursor()
            char_format = cursor.charFormat()
            cursor.select(cursor.SelectionType.Document)
            char_format.setFont(QFont(self.spd.user_settings['font_family'], int(self.spd.user_settings['font_size'])))
            cursor.setCharFormat(char_format)
            cursor.clearSelection()
            text_edit.setTextCursor(cursor)
        for tab_bar in self.findChildren(QTabBar):
            tab_bar.setFont(QFont(self.spd.user_settings['font_family'], int(self.spd.user_settings['font_size']) + 2))

        if font_chooser:
            font_chooser.deleteLater()

        if save_config:
            self.spd.save_user_settings()

        self.changes = current_changes_status

    def apply_line_spacing(self):
        current_changes_status = self.changes
        line_height = self.spd.user_settings['line_spacing']
        for text_edit in self.findChildren(CustomTextEdit):
            html = text_edit.toSimplifiedHtml()
            text_edit.document().setDefaultStyleSheet(
                'p, li {'
                'font-family: ' + self.spd.user_settings['font_family'] + ';'
                'font-size: ' + self.spd.user_settings['font_size'] + 'pt;'
                'line-height: ' + self.spd.user_settings['line_spacing'] + ';'
                '}'
            )
            text_edit.setHtml(html)

            cursor = text_edit.textCursor()
            char_format = cursor.charFormat()
            cursor.select(cursor.SelectionType.Document)
            char_format.setFont(QFont(self.spd.user_settings['font_family'], int(self.spd.user_settings['font_size'])))
            cursor.setCharFormat(char_format)
            cursor.clearSelection()
            text_edit.setTextCursor(cursor)
        self.menuBar().findChild(QAction, 'compact').setChecked(str(line_height) == '1.0')
        self.menuBar().findChild(QAction, 'regular').setChecked(str(line_height) == '1.2')
        self.menuBar().findChild(QAction, 'wide').setChecked(str(line_height) == '1.5')
        self.changes = current_changes_status

    def set_style_buttons(self, component):
        """
        Method for changing the GUI's style buttons (bold, italic, underline, bullets) based on where the cursor is
        located.

        :param QObject component: the QObject (CustomTextEdit) that is currently being used.
        """
        cursor = component.textCursor()
        font = cursor.charFormat().font()
        if font.weight() == QFont.Weight.Normal:
            self.toolbar.bold_button.setChecked(False)
        else:
            self.toolbar.bold_button.setChecked(True)
        if font.italic():
            self.toolbar.italic_button.setChecked(True)
        else:
            self.toolbar.italic_button.setChecked(False)
        if font.underline():
            self.toolbar.underline_button.setChecked(True)
        else:
            self.toolbar.underline_button.setChecked(False)

        text_list = cursor.currentList()
        if text_list:
            self.toolbar.bullet_button.setChecked(True)
        else:
            self.toolbar.bullet_button.setChecked(False)
        
    def fill_values(self, record):
        """
        Takes all the values from the currently accessed record and places them in their proper elements in the GUI.

        :param list of str record: a list whose first element is a list of values in their proper order.
        """
        index = 1
        self.setWindowTitle('Sermon Prep Database - ' + str(record[0][17]) + ' - ' + str(record[0][3]))

        for i in range(self.scripture_frame_layout.count()):
            component = self.scripture_frame_layout.itemAt(i).widget()

            if isinstance(component, QLineEdit):
                if record[0][index]:
                    component.setText(str(record[0][index].replace('&quot;', '"')).strip())
                else:
                    component.clear()
                index += 1

            elif isinstance(component, CustomTextEdit):
                component.clear()
                component.full_spell_check_done = False

                if record[0][index]:
                    component.setHtml(self.spd.reformat_string_for_load(record[0][index]))
                    """spell_check = SpellCheck(component, 'whole', self)
                    self.spd.spell_check_thread_pool.start(spell_check)"""

                index += 1

        for i in range(self.exegesis_frame_layout.count()):
            component = self.exegesis_frame_layout.itemAt(i).widget()

            if isinstance(component, CustomTextEdit) and not component.objectName() == 'text_box':
                component.clear()
                component.full_spell_check_done = False

                if record[0][index]:
                    component.setHtml(self.spd.reformat_string_for_load(record[0][index]))
                    """spell_check = SpellCheck(component, 'whole', self)
                    self.spd.spell_check_thread_pool.start(spell_check)"""

                index += 1

        for i in range(self.outline_frame_layout.count()):
            component = self.outline_frame_layout.itemAt(i).widget()

            if isinstance(component, CustomTextEdit) and not component.objectName() == 'text_box':
                component.clear()
                component.full_spell_check_done = False

                if record[0][index]:
                    component.setHtml(self.spd.reformat_string_for_load(record[0][index]))
                    """spell_check = SpellCheck(component, 'whole', self)
                    self.spd.spell_check_thread_pool.start(spell_check)"""

                index += 1

        for i in range(self.research_frame_layout.count()):
            component = self.research_frame_layout.itemAt(i).widget()

            if isinstance(component, CustomTextEdit):
                component.clear()
                component.full_spell_check_done = False

                if record[0][index]:
                    component.setHtml(self.spd.reformat_string_for_load(record[0][index]))
                    """spell_check = SpellCheck(component, 'whole', self)
                    self.spd.spell_check_thread_pool.start(spell_check)"""

                index += 1

        for i in range(self.sermon_frame_layout.count()):
            component = self.sermon_frame_layout.itemAt(i).widget()

            if isinstance(component, QLineEdit):
                if record[0][index]:
                    component.setText(record[0][index].replace('&quot;', '"'))
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
                    date_split = []
                    unusable_date = True

                if not unusable_date:
                    if int(date_split[0]) > 31:
                        component.setDate(QDate(int(date_split[0]), int(date_split[1]), int(date_split[2])))
                    elif int(date_split[2]) > 31:
                        component.setDate(QDate(int(date_split[2]), int(date_split[0]), int(date_split[1])))

                else:
                    self.spd.write_to_log('unusable date in record #' + str(record[0][0]))
                    component.setDate(QDateTime.currentDateTime().date())
                index += 1

            if isinstance(component, CustomTextEdit):
                component.clear()
                component.full_spell_check_done = False

                if record[0][index]:
                    component.setHtml(self.spd.reformat_string_for_load(record[0][index]))
                    """spell_check = SpellCheck(component, 'whole', self)
                    self.spd.spell_check_thread_pool.start(spell_check)"""

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
                text_edit = widget.findChild(QTextEdit, 'text_box_text_edit')
                text_edit.setText(self.sermon_text_edit.toPlainText())
                text_title.setText(self.sermon_reference_field.text())

        self.toolbar.id_label.setText('ID: ' + str(record[0][0]))

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
            self.toolbar.references_cb.setItemText(self.toolbar.references_cb.currentIndex(), self.sermon_reference_field.text())
            self.setWindowTitle('Sermon Prep Database - ' + self.sermon_date_edit.text() + ' - ' + self.sermon_reference_field.text())

            num_tabs = self.tabbed_frame.count()
            for i in range(num_tabs):
                if i > 0:
                    frame = self.tabbed_frame.widget(i)
                    widget = frame.findChild(QWidget, 'text_box')
                    text_title = widget.findChild(QLabel, 'text_title')
                    text_title.setText(self.sermon_reference_field.text())

            if self.spd.user_settings['auto_fill']:
                if not self.gs: # only create one instance of GetScripture
                    self.gs = GetScripture(self.spd)

                if ':' in self.sermon_reference_field.text(): # only attempt to get the text if there's enough to work with
                    passage = self.gs.get_passage(self.sermon_reference_field.text())
                    if passage and not passage == -1:
                        self.sermon_text_edit.setText(passage)

            self.changes = True
        except Exception as ex:
            self.spd.write_to_log(str(ex))

    def auto_fill(self):
        """
        Method to change the self.spd.auto_fill value based on user input then save that change to the database.
        """
        self.spd.user_settings['auto_fill'] = self.auto_fill_checkbox.isChecked()
        self.spd.save_user_settings()
        if self.auto_fill_checkbox.isChecked():
            self.reference_changes()

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
        self.toolbar.dates_cb.setItemText(self.toolbar.dates_cb.currentIndex(), self.sermon_date_edit.text())
        self.setWindowTitle(
            'Sermon Prep Database - ' + self.sermon_date_edit.text() + ' - ' + self.sermon_reference_field.text())
        self.changes = True

    def set_text_cursor(self, widget, cursor):
        widget.setTextCursor(cursor)

    def set_text_color(self, widget, indices, color):
        try:
            for index in indices:
                cursor = widget.textCursor()
                cursor.setPosition(index, QTextCursor.MoveMode.MoveAnchor)
                cursor.select(QTextCursor.SelectionType.WordUnderCursor)
                char_format = cursor.charFormat()
                char_format.setForeground(color)
                cursor.mergeCharFormat(char_format)
                cursor.clearSelection()
                char_format.setForeground(color)
                cursor.mergeCharFormat(char_format)
        except Exception as ex:
            print(str(ex))

    def reset_cursor_color(self, widget):
        cursor = widget.textCursor()
        char_format = cursor.charFormat()
        char_format.setForeground(Qt.GlobalColor.black)
        widget.mergeCurrentCharFormat(char_format)

    def do_exit(self, evt):
        """
        Method to ask if changes are to be saved before exiting the program.
        """
        goon = True
        if self.changes:
            goon = self.spd.ask_save()
        if goon:
            self.deleteLater()
            evt.accept()

    def closeEvent(self, evt:QCloseEvent):
        """
        @override
        Ignore the closeEvent and run GUI's do_exit method instead.
        """
        evt.ignore()
        self.do_exit(evt)

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
        if ((event.modifiers() & Qt.KeyboardModifier.ControlModifier)
                and (event.modifiers() & Qt.KeyboardModifier.ShiftModifier)
                and event.key() == Qt.Key.Key_B):
            self.toolbar.set_bullet()
            self.toolbar.bullet_button.blockSignals(True)
            if self.toolbar.bullet_button.isChecked():
                self.toolbar.bullet_button.setChecked(False)
            else:
                self.toolbar.bullet_button.setChecked(True)
            self.toolbar.bold_button.blockSignals(False)
        elif event.modifiers() & Qt.KeyboardModifier.ControlModifier and event.key() == Qt.Key.Key_B:
            self.toolbar.set_bold()
            self.toolbar.bold_button.blockSignals(True)
            if self.toolbar.bold_button.isChecked():
                self.toolbar.bold_button.setChecked(False)
            else:
                self.toolbar.bold_button.setChecked(True)
            self.toolbar.bold_button.blockSignals(False)
        elif event.modifiers() & Qt.KeyboardModifier.ControlModifier and event.key() == Qt.Key.Key_I:
            self.toolbar.set_italic()
            self.toolbar.italic_button.blockSignals(True)
            if self.toolbar.italic_button.isChecked():
                self.toolbar.italic_button.setChecked(False)
            else:
                self.toolbar.italic_button.setChecked(True)
            self.toolbar.italic_button.blockSignals(False)
        elif event.modifiers() & Qt.KeyboardModifier.ControlModifier and event.key() == Qt.Key.Key_U:
            self.toolbar.set_underline()
            self.toolbar.underline_button.blockSignals(True)
            if self.toolbar.underline_button.isChecked():
                self.toolbar.underline_button.setChecked(False)
            else:
                self.toolbar.underline_button.setChecked(True)
            self.toolbar.underline_button.blockSignals(False)
        elif event.modifiers() & Qt.KeyboardModifier.ControlModifier and event.key() == Qt.Key.Key_S:
            self.spd.save_rec()
        elif event.modifiers() & Qt.KeyboardModifier.ControlModifier and event.key() == Qt.Key.Key_P:
            self.menu_bar.print_rec()
        elif event.modifiers() & Qt.KeyboardModifier.ControlModifier and event.key() == Qt.Key.Key_Q:
            self.close()
        event.accept()


class InitialStartup(QRunnable):
    """
    QDialog class that shows the user the process of starting up the application.

    :param GUI gui: The GUI class.
    """
    def __init__(self, gui):
        super().__init__()

        self.gui = gui
        self.startup_splash = StartupSplash(gui, 6)

    def run(self):
        self.startup_splash.update_text.emit('Getting System Info')
        self.gui.spd.get_system_info()

        self.startup_splash.update_text.emit('Getting User Settings')
        self.gui.spd.get_user_settings()

        self.startup_splash.update_text.emit('Loading Dictionaries')
        ld = LoadDictionary(self.gui.spd)
        self.gui.spd.load_dictionary_thread_pool.start(ld)

        self.startup_splash.update_text.emit('Getting Indices')
        self.gui.spd.get_ids()
        self.gui.spd.get_date_list()
        self.gui.spd.get_scripture_list()
        self.gui.spd.backup_db()

        self.startup_splash.update_text.emit('Finishing Up')
        self.gui.create_main_gui.emit()
        self.startup_splash.end.emit()

        self.gui.changes = False


class StartupSplash(QDialog):
    update_text = pyqtSignal(str)
    end = pyqtSignal()

    def __init__(self, gui, progress_end):
        super().__init__()
        self.gui = gui
        self.update_text.connect(self.change_text)
        self.end.connect(lambda: self.done(0))

        self.setWindowFlag(Qt.WindowType.FramelessWindowHint)
        self.setModal(True)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setMinimumWidth(300)

        layout = QGridLayout()
        self.setLayout(layout)

        self.working_label = QLabel()
        self.working_label.setAutoFillBackground(False)
        self.working_label.setPixmap(QPixmap('resources/icon.png'))
        #movie = QMovie('resources/waitIcon.webp')
        #self.working_label.setMovie(movie)
        layout.addWidget(self.working_label, 0, 0, Qt.AlignmentFlag.AlignHCenter)
        #movie.start()

        self.status_label = QLabel('Starting...')
        self.status_label.setFont(QFont('Helvetica', 16, QFont.Weight.Bold))
        self.status_label.setStyleSheet('color: #d7d7f4; text-align: center;')
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.status_label, 1, 0, Qt.AlignmentFlag.AlignCenter)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(1, progress_end)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setStyleSheet(
            'QProgressBar {'
                'border: 1px solid gray;'
            '}'
            'QProgressBar::chunk {'
                'border: none;'
                'background: #d7d7f4;'
            '}'
        )
        layout.addWidget(self.progress_bar)

        self.show()
        QApplication.processEvents()

    def change_text(self, text):
        """
        Method to change the text shown on the splash screen.

        :param str text: The text to display.
        """
        self.status_label.setText(text)
        self.progress_bar.setValue(self.progress_bar.value() + 1)
        QApplication.processEvents()


class CustomTextEdit(QTextEdit):
    """
    CustomTextEdit is an implementation of QTextEdit that adds spell-checking capabilities. Two different types of
    spell checking are done depending on user input and current spell-check state: check previous word, or check
    whole text. Also modifies the standard context menu to include spell check suggestions and an option to remove
    extra white space from the text.

    :param GUI gui: The GUI object
    """
    def __init__(self, gui):
        super().__init__()
        self.setObjectName('custom_text_edit')
        self.gui = gui
        self.full_spell_check_done = False
        self.document().setDefaultStyleSheet(
            'p {'
            'font-family: ' + self.gui.spd.user_settings['font_family'] + ';'
            'font-size: ' + self.gui.spd.user_settings['font_size'] + 'pt;'
            'line-height: ' + self.gui.spd.user_settings['line_spacing'] + ';'
            '}'
        )

        self.textChanged.connect(self.text_changed)
        self.word_spell_check = SpellCheck(self, 'previous', self.gui)
        self.word_spell_check.setAutoDelete(False)

    def keyReleaseEvent(self, evt):
        cursor = self.textCursor()
        if cursor.charFormat().foreground() == Qt.GlobalColor.red:
            self.word_spell_check.type = 'current'
            self.gui.spd.spell_check_thread_pool.start(self.word_spell_check)
        elif (evt.key() == Qt.Key.Key_Space
                or evt.key() == Qt.Key.Key_Return
                or evt.key() == Qt.Key.Key_Enter):
            if not self.gui.spd.user_settings['disable_spell_check']:
                self.word_spell_check.type = 'previous'
                self.gui.spd.spell_check_thread_pool.start(self.word_spell_check)

    def keyPressEvent(self, evt):
        if evt.key() == Qt.Key.Key_Enter or evt.key() == Qt.Key.Key_Return:
            self.append('')
        else:
            super().keyPressEvent(evt)

    def contextMenuEvent(self, evt):
        """
        @override
        Alters the standard context menu of this QTextEdit to include a list of spell-check words as well as an option
        to remove extraneous whitespace from the text.
        """
        menu = self.createStandardContextMenu()

        clean_whitespace_action = QAction("Remove extra whitespace")
        clean_whitespace_action.triggered.connect(self.clean_whitespace)
        menu.insertAction(menu.actions()[0], clean_whitespace_action)
        menu.insertSeparator(menu.actions()[1])

        if not self.gui.spd.disable_spell_check:
            cursor = self.cursorForPosition(evt.pos())
            cursor.select(QTextCursor.SelectionType.WordUnderCursor)
            word = cursor.selection().toPlainText()

            if len(word) > 0:
                upper = False
                if word[0].isupper():
                    upper = True

                spell_check = SpellCheck(None, None, self.gui)
                suggestions = spell_check.check_single_word(word)

                next_menu_index = 0
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
                action.triggered.connect(lambda: self.gui.spd.add_to_dictionary(self, spell_check.clean_word(word)))
                menu.insertAction(menu.actions()[next_menu_index + 2], action)
                menu.insertSeparator(menu.actions()[next_menu_index + 3])

            menu.exec(evt.globalPos())
            menu.close()

    def toSimplifiedHtml(self):
        """
        Method to strip unneeded html tags and convert others to simpler tags
        """
        string = self.toHtml()

        string = re.split('<body.*?>', string)[1]

        # preserve any desired tags from getting removed during the wholesale <.*?> removal
        string = re.sub('<p.*?>', '{p}', string)
        string = re.sub('</p>', '{/p}', string)
        string = re.sub('<ul.*?>', '{ul}', string)
        string = re.sub('</ul>', '{/ul}', string)
        string = re.sub('<li.*?>', '{li}', string)
        string = re.sub('</li>', '{/li}', string)

        bold_texts = re.findall('<span.*?font-weight.*?</span>', string)
        for text in bold_texts:
            new_text = re.sub('<.*?>', '', text)
            new_text = '{b}' + new_text + '{/b}'
            string = string.replace(text, new_text)

        italic_texts = re.findall('<span.*?font-style.*?</span>', string)
        for text in italic_texts:
            new_text = re.sub('<.*?>', '', text)
            new_text = '{i}' + new_text + '{/i}'
            string = string.replace(text, new_text)

        underline_texts = re.findall('<span.*?text-decoration.*?</span>', string)
        for text in underline_texts:
            new_text = re.sub('<.*?>', '', text)
            new_text = '{u}' + new_text + '{/u}'
            string = string.replace(text, new_text)

        # convert preserved tags back to their original form
        string = re.sub('<.*?>', '', string)
        string = string.replace('{p}', '<p>')
        string = string.replace('{/p}', '</p>\n')
        string = string.replace('{ul}', '<ul>')
        string = string.replace('{/ul}', '</ul>')
        string = string.replace('{li}', '<li>')
        string = string.replace('{/li}', '</li>\n')
        string = string.replace('{b}', '<b>')
        string = string.replace('{/b}', '</b>')
        string = string.replace('{i}', '<i>')
        string = string.replace('{/i}', '</i>')
        string = string.replace('{u}', '<u>')
        string = string.replace('{/u}', '</u>')

        return string.strip()

    def text_changed(self):
        self.gui.changes = True

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
        cursor.movePosition(QTextCursor.MoveOperation.NextCharacter, QTextCursor.MoveMode.KeepAnchor)
        self.setTextCursor(cursor)

        cursor = self.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.PreviousCharacter, QTextCursor.MoveMode.MoveAnchor)
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
        component = self.gui.focusWidget()
        if isinstance(component, QTextEdit):
            string = component.toHtml()
            string = re.sub(' +', ' ', string)
            string = re.sub('\t+', '\t', string)

            component.setHtml(string)
        self.gui.changes = True

    def paintEvent(self, evt):
        if not self.full_spell_check_done:
            spell_check = SpellCheck(self, 'whole', self.gui)
            self.gui.spd.spell_check_thread_pool.start(spell_check)
            self.full_spell_check_done = True
        super().paintEvent(evt)


class ScriptureBox(QWidget):
    """
    Creates an independent QWidget that can be added or removed from layouts based on user's input.
    """
    def __init__(self):
        super().__init__()
        self.setObjectName('text_box')
        self.setMaximumWidth(300)
        text_layout = QVBoxLayout()
        self.setLayout(text_layout)
        text_title = QLabel()
        text_title.setObjectName('text_title')
        text_layout.addWidget(text_title)
        self.text_edit = QTextEdit()
        self.text_edit.setObjectName('text_box_text_edit')
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

        results_header = QWidget()
        header_layout = QHBoxLayout()
        results_header.setLayout(header_layout)

        results_label = QLabel()
        header_layout.addWidget(results_label)

        close_button = QPushButton()
        close_button.setIcon(QIcon('resources/svg/spCloseIcon.svg'))
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
        model.setHeaderData(0, Qt.Orientation.Horizontal, 'ID')
        model.setHeaderData(1, Qt.Orientation.Horizontal, '# of\r\nMatches')
        model.setHeaderData(2, Qt.Orientation.Horizontal, 'Word(s) Found')
        model.setHeaderData(3, Qt.Orientation.Horizontal, 'Sermon Text')
        model.setHeaderData(4, Qt.Orientation.Horizontal, 'Sermon Title')
        model.setHeaderData(5, Qt.Orientation.Horizontal, 'Sermon Date')
        model.setHeaderData(6, Qt.Orientation.Horizontal, 'Sermon Snippet')

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
        results_table_view.setSelectionBehavior(QTableView.SelectionBehavior.SelectRows)
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


class SermonView(QWidget):
    def __init__(self, gui, text):
        super().__init__()
        self.gui = gui
        self.text = text
        self.button_bar = QWidget()
        self.next_page_button = QPushButton('next')
        self.previous_page_button = QPushButton('previous')
        self.sermon_label = QLabel()
        self.num_pages = 0
        self.current_page = 0
        self.current_font = QFont(
            self.gui.spd.user_settings['font_family'], int(self.gui.spd.user_settings['font_size']))
        self.page_label = None

        self.init_components()
        self.showMaximized()
        self.gui.spd.app.processEvents()

        self.page_width = self.width() - 40
        self.page_height = self.height() - self.button_bar.height() - 40

        self.page_pixmaps = []
        self.create_document_pages()
        self.set_page()

        self.previous_page_button.setEnabled(False)
        if len(self.page_pixmaps) < 2:
            self.next_page_button.setEnabled(False)
        self.add_page_buttons()

    def init_components(self):
        self.setParent(self.gui)
        self.setWindowFlag(Qt.WindowType.Window)
        self.setWindowTitle('Sermon Viewer')

        layout = QVBoxLayout(self)

        button_layout = QHBoxLayout(self.button_bar)
        layout.addWidget(self.button_bar)

        zoom_out_button = QPushButton()
        zoom_out_button.setObjectName('zoom_out')
        zoom_out_button.setIcon(QIcon('resources/svg/spZoomOut.svg'))
        zoom_out_button.setIconSize(QSize(36, 36))
        zoom_out_button.pressed.connect(self.zoom)
        button_layout.addWidget(zoom_out_button)
        button_layout.addSpacing(20)

        zoom_in_button = QPushButton()
        zoom_in_button.setObjectName('zoom_in')
        zoom_in_button.setIcon(QIcon('resources/svg/spZoomIn.svg'))
        zoom_in_button.setIconSize(QSize(36, 36))
        zoom_in_button.pressed.connect(self.zoom)
        button_layout.addWidget(zoom_in_button)
        button_layout.addStretch()

        self.page_label = QLabel()
        self.page_label.setFont(self.gui.bold_font)
        button_layout.addWidget(self.page_label)
        button_layout.addStretch()

        layout.addWidget(self.sermon_label)
        layout.addStretch()

    def add_page_buttons(self):
        page_button_next_stylesheet = (
            'QPushButton {'
            '   border: none;'
            '   background: transparent;'
            '   color: #00000000;'
            '}'
            'QPushButton:hover {'
            '   background-color: rgba(255, 255, 255, 150);'
            '   background-image: url("resources/svg/spNextPageIcon.svg");'
            '   background-repeat: no-repeat;'
            '   background-position: center;'
            '}'
            'QPushButton:pressed {'
            '   background-color: #ffffff;'
            '   color: #000000;'
            '}'
        )
        page_button_previous_stylesheet = (
            'QPushButton {'
            '   border: none;'
            '   background: transparent;'
            '   color: #00000000;'
            '}'
            'QPushButton:hover {'
            '   background-color: rgba(255, 255, 255, 150);'
            '   background-image: url("resources/svg/spPrevPageIcon.svg");'
            '   background-repeat: no-repeat;'
            '   background-position: center;'
            '}'
            'QPushButton:pressed {'
            '   background-color: #ffffff;'
            '   color: #000000;'
            '}'
        )

        self.previous_page_button.setParent(self)
        self.previous_page_button.setObjectName('previous')
        self.previous_page_button.setGeometry(
            0, self.height() - self.page_height - 20, 200, self.page_height)
        self.previous_page_button.setStyleSheet(page_button_previous_stylesheet)
        self.previous_page_button.pressed.connect(self.set_page)
        self.previous_page_button.show()

        self.next_page_button.setParent(self)
        self.next_page_button.setObjectName('next')
        self.next_page_button.setGeometry(self.width() - 200, self.height() - self.page_height - 20, 200, self.page_height)
        self.next_page_button.setStyleSheet(page_button_next_stylesheet)
        self.next_page_button.pressed.connect(self.set_page)
        self.next_page_button.show()

    def create_document_pages(self):
        document = QTextDocument()
        document.setPageSize(QSizeF(self.page_width, self.page_height))
        document.setDefaultFont(self.current_font)

        text_option = QTextOption()
        text_option.setWrapMode(QTextOption.WrapMode.WordWrap)
        document.setDefaultTextOption(text_option)

        document.setHtml(self.text)

        self.page_pixmaps = []
        current_y = 0
        painter = QPainter()
        self.num_pages = document.pageCount()
        for i in range(document.pageCount()):
            pixmap = QPixmap(self.page_width, self.page_height)
            pixmap.fill(Qt.GlobalColor.white)

            painter.begin(pixmap)
            painter.translate(0, -current_y)
            document.drawContents(painter)
            painter.end()

            self.page_pixmaps.append(pixmap)
            current_y += self.page_height

    def zoom(self):
        if self.sender().objectName() == 'zoom_in':
            self.current_font.setPointSize(self.current_font.pointSize() + 2)
        else:
            self.current_font.setPointSize(self.current_font.pointSize() - 2)
        self.create_document_pages()
        self.set_page()

    def set_page(self):
        if self.sender().objectName() == 'next':
            self.current_page += 1
        elif self.sender().objectName() == 'previous':
            self.current_page -= 1

        if self.current_page == 0:
            self.previous_page_button.setEnabled(False)
        else:
            self.previous_page_button.setEnabled(True)

        if self.current_page >= len(self.page_pixmaps) - 1:
            self.current_page = len(self.page_pixmaps) - 1
            self.next_page_button.setEnabled(False)
        else:
            self.next_page_button.setEnabled(True)

        self.sermon_label.setPixmap(self.page_pixmaps[self.current_page])
        self.page_label.setText(f'Page {self.current_page + 1} of {self.num_pages}')
