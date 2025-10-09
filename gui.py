from os.path import exists

from PyQt6.QtCore import Qt, QSize, QDate, QDateTime, pyqtSignal, QThreadPool
from PyQt6.QtGui import QIcon, QFont, QPixmap, QCloseEvent, QAction, QUndoStack, QTextCursor, QTextBlockFormat
from PyQt6.QtWidgets import QWidget, QTabWidget, QGridLayout, QLabel, QCheckBox, QDateEdit, QTextEdit, QMainWindow, \
    QVBoxLayout, QPushButton, QTabBar

from get_scripture import GetScripture
from spell_check_widgets import SpellCheckTextEdit, SpellCheckLineEdit
from widgets import MenuBar, StartupSplash
from runnables import LoadDictionary
from widgets import Toolbar
from widgets import ScriptureBox, SermonView


class GUI(QMainWindow):
    clear_changes_signal = pyqtSignal()
    undo_stack = None
    changes = False
    gs = None
    spell_check = None
    
    def __init__(self, main):
        """
        GUI handles all the operations from the user interface. It builds the QT window and elements and requires the
        SermonPrepDatabase object in order to access its methods.
        """
        super().__init__()
        self.main = main
        self.clear_changes_signal.connect(self.clear_changes)

        self.startup_splash = StartupSplash(self, 6)
        self.startup_splash.show()
        self.main.app.processEvents()

        self.change_startup_splash_text('Getting System Info')
        self.main.get_system_info()

        self.change_startup_splash_text('Getting User Settings')
        self.main.get_user_settings()

        self.change_startup_splash_text('Loading Dictionaries')
        self.main.spell_check_thread_pool = QThreadPool()
        self.main.spell_check_thread_pool.setStackSize(256000000)
        self.main.load_dictionary_thread_pool = QThreadPool()
        ld = LoadDictionary(self.main)
        self.main.load_dictionary_thread_pool.start(ld)
        self.main.load_dictionary_thread_pool.waitForDone()

        self.change_startup_splash_text('Getting Indices')
        self.main.get_ids()
        self.main.get_date_list()
        self.main.get_scripture_list()
        self.main.backup_db()

        self.change_startup_splash_text('Finishing Up')

        self.standard_font = QFont(self.main.user_settings['font_family'], int(self.main.user_settings['font_size']))
        self.bold_font = QFont(
            self.main.user_settings['font_family'], int(self.main.user_settings['font_size']), QFont.Weight.Bold)
        self.central_widget = QWidget()
        self.central_layout = QVBoxLayout(self.central_widget)
        self.tab_widget = QTabWidget()
        self.menu_bar = MenuBar(self, self.main)
        self.toolbar = Toolbar(self, self.main)
        self.scripture_widget = QWidget()
        self.scripture_layout = QGridLayout(self.scripture_widget)
        self.sermon_reference_field = SpellCheckLineEdit(self)
        self.auto_fill_checkbox = QCheckBox('Auto-fill ' + self.main.user_settings['label4'])
        self.sermon_text_edit = SpellCheckTextEdit(self)
        self.exegesis_widget = QWidget()
        self.exegesis_layout = QGridLayout(self.exegesis_widget)
        self.outline_widget = QWidget()
        self.outline_layout = QGridLayout(self.outline_widget)
        self.research_widget = QWidget()
        self.research_layout = QGridLayout(self.research_widget)
        self.sermon_widget = QWidget()
        self.sermon_layout = QGridLayout(self.sermon_widget)
        self.sermon_date_edit = QDateEdit()

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

        self.create_gui()

        self.startup_splash.deleteLater()

    def change_startup_splash_text(self, text):
        """
        Method to change the text shown on the splash screen.

        :param str text: The text to display.
        """
        self.startup_splash.status_label.setText(text)
        self.startup_splash.progress_bar.setValue(self.startup_splash.progress_bar.value() + 1)
        self.main.app.processEvents()

    def create_gui(self):
        """
        Builds the QT GUI, also using elements from menu_bar.py, toolbar.py, and print_dialog.py
        """
        self.setWindowTitle('Sermon Prep Database')

        self.standard_font = QFont(self.main.user_settings['font_family'], int(self.main.user_settings['font_size']))
        self.bold_font = QFont(
            self.main.user_settings['font_family'], int(self.main.user_settings['font_size']), QFont.Weight.Bold)
        self.main.line_spacing = str(self.main.user_settings['line_spacing'])

        icon_pixmap = QPixmap('resources/svg/spIcon.svg')
        self.setWindowIcon(QIcon(icon_pixmap))

        self.setCentralWidget(self.central_widget)

        self.undo_stack = QUndoStack(self.central_widget)

        self.central_layout.addWidget(self.toolbar)

        self.build_tabbed_frame()
        self.build_scripture_tab()
        self.build_exegesis_tab()
        self.build_outline_tab()
        self.build_research_tab()
        self.build_sermon_tab()

        self.menu_bar.color_change(self.main.user_settings['theme'])

        self.main.current_rec_index = len(self.main.ids) - 1
        self.apply_font(self.main.user_settings['font_family'], self.main.user_settings['font_size'])
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
        self.central_layout.addWidget(tab_container)

        self.tab_widget.setAutoFillBackground(True)
        self.tab_widget.setIconSize(QSize(24, 24))
        self.tab_widget.tabBar().currentChanged.connect(self.current_tab_changed)
        tab_container_layout.addWidget(self.tab_widget)

    def current_tab_changed(self):
        if self.main.user_settings['theme'] == 'dark':
            return
        index = self.sender().currentIndex()
        for i in range(5):
            if i == index:
                self.tab_widget.setTabIcon(i, self.dark_tab_icons[i])
            else:
                self.tab_widget.setTabIcon(i, self.light_tab_icons[i])
        
    def build_scripture_tab(self, insert=False):
        """
        Create a QWidget to hold the scripture tab's elements. Adds the elements.

        :param boolean insert: If other tabs already exist, insert tab at position 0.
        """
        self.scripture_layout.setColumnStretch(0, 1)
        self.scripture_layout.setColumnStretch(1, 1)
        self.scripture_layout.setColumnStretch(2, 0)
        
        pericope_label = QLabel(self.main.user_settings['label1'])
        self.scripture_layout.addWidget(pericope_label, 0, 0)
        
        pericope_line_edit = SpellCheckLineEdit(self)
        self.scripture_layout.addWidget(pericope_line_edit, 1, 0)
        
        pericope_text_label = QLabel(self.main.user_settings['label2'])
        self.scripture_layout.addWidget(pericope_text_label, 2, 0)

        pericope_text_edit = SpellCheckTextEdit(self)
        pericope_text_edit.cursorPositionChanged.connect(self.set_style_buttons)
        self.scripture_layout.addWidget(pericope_text_edit, 3, 0)
        
        sermon_reference_label = QLabel(self.main.user_settings['label3'])
        self.scripture_layout.addWidget(sermon_reference_label, 0, 1)

        if exists(self.main.app_dir + '/my_bible.xml'):
            self.scripture_layout.addWidget(self.sermon_reference_field, 1, 1)

            self.auto_fill_checkbox.setChecked(True)
            self.scripture_layout.addWidget(self.auto_fill_checkbox, 1, 2)
            if self.main.user_settings['auto_fill']:
                self.auto_fill_checkbox.setChecked(True)
            else:
                self.auto_fill_checkbox.setChecked(False)
            self.auto_fill_checkbox.stateChanged.connect(self.auto_fill)
        else:
            self.scripture_layout.addWidget(self.sermon_reference_field, 1, 1, 1, 2)
        
        sermon_text_label = QLabel(self.main.user_settings['label4'])
        self.scripture_layout.addWidget(sermon_text_label, 2, 1, 1, 2)

        self.sermon_text_edit.cursorPositionChanged.connect(self.set_style_buttons)
        self.scripture_layout.addWidget(self.sermon_text_edit, 3, 1, 1, 2)

        if insert:
            self.tab_widget.insertTab(
                0,
                self.scripture_widget,
                self.light_tab_icons[0],
                'Scripture'
            )
        else:
            self.tab_widget.addTab(
                self.scripture_widget,
                self.light_tab_icons[0],
                'Scripture'
            )
        
    def build_exegesis_tab(self):
        """
        Create a QWidget to hold the exegesis tab's elements. Adds the elements.
        """
        self.exegesis_layout.setColumnMinimumWidth(1, 20)
        self.exegesis_layout.setColumnMinimumWidth(3, 20)
        self.exegesis_layout.setColumnMinimumWidth(5, 20)
        self.exegesis_layout.setRowMinimumHeight(2, 50)
        self.exegesis_layout.setRowMinimumHeight(5, 50)
        self.exegesis_layout.setRowStretch(8, 100)
        
        fcft_label = QLabel(self.main.user_settings['label5'])
        self.exegesis_layout.addWidget(fcft_label, 0, 0)
        
        fcft_text = SpellCheckTextEdit(self)
        fcft_text.cursorPositionChanged.connect(self.set_style_buttons)
        self.exegesis_layout.addWidget(fcft_text, 1, 0)
        
        gat_label = QLabel(self.main.user_settings['label6'])
        self.exegesis_layout.addWidget(gat_label, 3, 0)
        
        gat_text = SpellCheckTextEdit(self)
        gat_text.cursorPositionChanged.connect(self.set_style_buttons)
        self.exegesis_layout.addWidget(gat_text, 4, 0)
        
        cpt_label = QLabel(self.main.user_settings['label7'])
        self.exegesis_layout.addWidget(cpt_label, 6, 0)
        
        cpt_text = SpellCheckTextEdit(self)
        cpt_text.cursorPositionChanged.connect(self.set_style_buttons)
        self.exegesis_layout.addWidget(cpt_text, 7, 0)
        
        pb_label = QLabel(self.main.user_settings['label8'])
        self.exegesis_layout.addWidget(pb_label, 3, 2)
        
        pb_text = SpellCheckTextEdit(self)
        pb_text.cursorPositionChanged.connect(self.set_style_buttons)
        self.exegesis_layout.addWidget(pb_text, 4, 2)
        
        fcfs_label = QLabel(self.main.user_settings['label9'])
        self.exegesis_layout.addWidget(fcfs_label, 0, 4)
        
        fcfs_text = SpellCheckTextEdit(self)
        fcfs_text.cursorPositionChanged.connect(self.set_style_buttons)
        self.exegesis_layout.addWidget(fcfs_text, 1, 4)
        
        gas_label = QLabel(self.main.user_settings['label10'])
        self.exegesis_layout.addWidget(gas_label, 3, 4)
        
        gas_text = SpellCheckTextEdit(self)
        gas_text.cursorPositionChanged.connect(self.set_style_buttons)
        self.exegesis_layout.addWidget(gas_text, 4, 4)
        
        cps_label = QLabel(self.main.user_settings['label11'])
        self.exegesis_layout.addWidget(cps_label, 6, 4)
        
        cps_text = SpellCheckTextEdit(self)
        cps_text.cursorPositionChanged.connect(self.set_style_buttons)
        self.exegesis_layout.addWidget(cps_text, 7, 4)

        scripture_box = ScriptureBox()
        self.exegesis_layout.addWidget(scripture_box, 0, 6, 8, 1)
        
        self.tab_widget.addTab(self.exegesis_widget, self.light_tab_icons[1], 'Exegesis')

    def build_outline_tab(self):
        """
        Create a QWidget to hold the outline tab's elements. Adds the elements.
        """
        self.outline_layout.setColumnMinimumWidth(1, 20)
        self.outline_layout.setColumnMinimumWidth(3, 20)
        self.outline_layout.setColumnMinimumWidth(5, 20)
        
        scripture_outline_label = QLabel(self.main.user_settings['label12'])
        self.outline_layout.addWidget(scripture_outline_label, 0, 0)
        
        scripture_outline_text = SpellCheckTextEdit(self)
        scripture_outline_text.cursorPositionChanged.connect(self.set_style_buttons)
        self.outline_layout.addWidget(scripture_outline_text, 1, 0)
        
        sermon_outline_label = QLabel(self.main.user_settings['label13'])
        self.outline_layout.addWidget(sermon_outline_label, 0, 2)
        
        sermon_outline_text = SpellCheckTextEdit(self)
        sermon_outline_text.cursorPositionChanged.connect(self.set_style_buttons)
        self.outline_layout.addWidget(sermon_outline_text, 1, 2)
        
        illustration_label = QLabel(self.main.user_settings['label14'])
        self.outline_layout.addWidget(illustration_label, 0, 4)
        
        illustration_text = SpellCheckTextEdit(self)
        illustration_text.cursorPositionChanged.connect(self.set_style_buttons)
        self.outline_layout.addWidget(illustration_text, 1, 4)

        scripture_box = ScriptureBox()
        self.outline_layout.addWidget(scripture_box, 0, 6, 5, 1)

        self.tab_widget.addTab(self.outline_widget, self.light_tab_icons[2], 'Outlines')
        
    def build_research_tab(self):
        """
        Create a QWidget to hold the research tab's elements. Adds the elements.
        """
        self.research_layout.setColumnMinimumWidth(1, 20)
        
        research_label = QLabel(self.main.user_settings['label15'])
        self.research_layout.addWidget(research_label, 0, 0)
        
        research_text_edit = SpellCheckTextEdit(self)
        research_text_edit.cursorPositionChanged.connect(self.set_style_buttons)
        self.research_layout.addWidget(research_text_edit, 1, 0)

        scripture_box = ScriptureBox()
        self.research_layout.addWidget(scripture_box, 0, 2, 2, 1)
        
        self.tab_widget.addTab(self.research_widget, self.light_tab_icons[3], 'Research')
        
    def build_sermon_tab(self):
        """
        Create a QWidget to hold the sermon tab's elements. Adds the elements.
        """
        self.sermon_layout.setColumnStretch(0, 5)
        self.sermon_layout.setColumnStretch(1, 2)
        self.sermon_layout.setColumnStretch(2, 5)
        self.sermon_layout.setColumnStretch(3, 2)
        
        sermon_title_label = QLabel(self.main.user_settings['label16'])
        self.sermon_layout.addWidget(sermon_title_label, 0, 0)
        
        sermon_title_field = SpellCheckLineEdit(self)
        self.sermon_layout.addWidget(sermon_title_field, 1, 0)
        
        sermon_date_label = QLabel(self.main.user_settings['label17'])
        self.sermon_layout.addWidget(sermon_date_label, 0, 1)

        self.sermon_date_edit.setCalendarPopup(True)
        self.sermon_date_edit.setFixedHeight(30)
        self.sermon_date_edit.dateChanged.connect(self.date_changes)
        self.sermon_layout.addWidget(self.sermon_date_edit, 1, 1)
        
        sermon_location_label = QLabel(self.main.user_settings['label18'])
        self.sermon_layout.addWidget(sermon_location_label, 0, 2)
        
        sermon_location_field = SpellCheckLineEdit(self)
        self.sermon_layout.addWidget(sermon_location_field, 1, 2)
        
        ctw_label = QLabel(self.main.user_settings['label19'])
        self.sermon_layout.addWidget(ctw_label, 2, 0)
        
        ctw_field = SpellCheckLineEdit(self)
        self.sermon_layout.addWidget(ctw_field, 3, 0)
        
        hr_label = QLabel(self.main.user_settings['label20'])
        self.sermon_layout.addWidget(hr_label, 2, 2)
        
        hr_field = SpellCheckLineEdit(self)
        self.sermon_layout.addWidget(hr_field, 3, 2)

        sermon_text = SpellCheckTextEdit(self)
        sermon_text.cursorPositionChanged.connect(self.set_style_buttons)
        self.sermon_layout.addWidget(sermon_text, 5, 0, 1, 4)

        self.sermon_view_button = QPushButton()
        if self.main.user_settings['theme'] == 'dark':
            self.sermon_view_button.setIcon(QIcon('resources/svg/spSermonViewIconLight.svg'))
        else:
            self.sermon_view_button.setIcon(QIcon('resources/svg/spSermonViewIconDark.svg'))
        self.sermon_view_button.setIconSize(QSize(36, 36))
        self.sermon_view_button.setFixedSize(QSize(56, 56))
        self.sermon_view_button.setToolTip('Open a window that shows your sermon')
        self.sermon_view_button.released.connect(lambda: SermonView(self, sermon_text.toSimplifiedHtml()))
        self.sermon_layout.addWidget(self.sermon_view_button, 0, 3, 4, 1)
        
        sermon_label = QLabel(self.main.user_settings['label21'])
        self.sermon_layout.addWidget(sermon_label, 4, 0)

        scripture_box = ScriptureBox()
        self.sermon_layout.addWidget(scripture_box, 0, 5, 6, 1)
        
        self.tab_widget.addTab(self.sermon_widget, self.light_tab_icons[4], 'Sermon')

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

        self.main.user_settings['font_family'] = font_family
        self.main.user_settings['font_size'] = font_size
        self.standard_font = QFont(font_family, int(font_size))
        self.bold_font = QFont(font_family, int(font_size), QFont.Weight.Bold)
        font = QFont(self.main.user_settings['font_family'], int(self.main.user_settings['font_size']))
        for label in self.findChildren(QLabel):
            label.setFont(font)
        for line_edit in self.findChildren(SpellCheckLineEdit):
            line_edit.setFont(font)
        for text_edit in self.findChildren(SpellCheckTextEdit):
            html = text_edit.toSimplifiedHtml()
            text_edit.setFont(font)
            text_edit.document().setDefaultStyleSheet(
                'p, li {'
                'font-family: ' + self.main.user_settings['font_family'] + ';'
                'font-size: ' + self.main.user_settings['font_size'] + 'pt;'
                'line-height: ' + self.main.user_settings['line_spacing'] + ';'
                '}'
            )
            text_edit.setHtml(html)

        for tab_bar in self.findChildren(QTabBar):
            tab_bar.setFont(QFont(self.main.user_settings['font_family'], int(self.main.user_settings['font_size']) + 2))

        if font_chooser:
            font_chooser.deleteLater()

        if save_config:
            self.main.save_user_settings()

        self.changes = current_changes_status

    def apply_line_spacing(self):
        current_changes_status = self.changes
        line_height = self.main.user_settings['line_spacing']
        for text_edit in self.findChildren(SpellCheckTextEdit):
            block = text_edit.document().begin()
            font_metrics = text_edit.fontMetrics()
            font_height = font_metrics.height()
            line_height_pixels = font_height * float(line_height)
            cursor = text_edit.textCursor()
            block_count = 0
            while block.isValid():
                block_format = block.blockFormat()
                block_format.setLineHeight(line_height_pixels, QTextBlockFormat.LineHeightTypes.FixedHeight.value)
                block_format.setTopMargin(0)
                block_format.setBottomMargin(line_height_pixels / 2)

                # Apply the format using a cursor
                cursor.setPosition(block.position())
                cursor.movePosition(QTextCursor.MoveOperation.EndOfBlock, QTextCursor.MoveMode.KeepAnchor)
                cursor.setBlockFormat(block_format)

                block = block.next()
                block_count += 1

        self.menuBar().findChild(QAction, 'compact').setChecked(str(line_height) == '1.0')
        self.menuBar().findChild(QAction, 'regular').setChecked(str(line_height) == '1.2')
        self.menuBar().findChild(QAction, 'wide').setChecked(str(line_height) == '1.5')
        self.changes = current_changes_status

    def set_style_buttons(self):
        """
        Method for changing the GUI's style buttons (bold, italic, underline, bullets) based on where the cursor is
        located.
        """
        cursor = self.sender().textCursor()
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

        for i in range(self.scripture_layout.count()):
            component = self.scripture_layout.itemAt(i).widget()

            if isinstance(component, SpellCheckLineEdit):
                if record[0][index]:
                    component.setText(str(record[0][index].replace('&quot;', '"')).strip())
                else:
                    component.clear()
                index += 1

            elif isinstance(component, SpellCheckTextEdit):
                component.clear()
                component.full_spell_check_done = False

                if record[0][index]:
                    component.setHtml(self.main.reformat_string_for_load(record[0][index]))

                index += 1

        for i in range(self.exegesis_layout.count()):
            component = self.exegesis_layout.itemAt(i).widget()

            if isinstance(component, SpellCheckTextEdit) and not component.objectName() == 'text_box':
                component.clear()
                component.full_spell_check_done = False

                if record[0][index]:
                    component.setHtml(self.main.reformat_string_for_load(record[0][index]))
                    """spell_check = SpellCheck(component, 'whole', self)
                    self.spd.spell_check_thread_pool.start(spell_check)"""

                index += 1

        for i in range(self.outline_layout.count()):
            component = self.outline_layout.itemAt(i).widget()

            if isinstance(component, SpellCheckTextEdit) and not component.objectName() == 'text_box':
                component.clear()
                component.full_spell_check_done = False

                if record[0][index]:
                    component.setHtml(self.main.reformat_string_for_load(record[0][index]))
                    """spell_check = SpellCheck(component, 'whole', self)
                    self.spd.spell_check_thread_pool.start(spell_check)"""

                index += 1

        for i in range(self.research_layout.count()):
            component = self.research_layout.itemAt(i).widget()

            if isinstance(component, SpellCheckTextEdit):
                component.clear()
                component.full_spell_check_done = False

                if record[0][index]:
                    component.setHtml(self.main.reformat_string_for_load(record[0][index]))

                index += 1

        for i in range(self.sermon_layout.count()):
            component = self.sermon_layout.itemAt(i).widget()

            if isinstance(component, SpellCheckLineEdit):
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
                    self.main.write_to_log('unusable date in record #' + str(record[0][0]))
                    component.setDate(QDateTime.currentDateTime().date())
                index += 1

            if isinstance(component, SpellCheckTextEdit):
                component.clear()
                component.full_spell_check_done = False

                if record[0][index]:
                    component.setHtml(self.main.reformat_string_for_load(record[0][index]))

                index += 1

            if component.objectName() == 'text_box':
                component.clear()
                index += 1

        num_tabs = self.tab_widget.count()
        for i in range(1, num_tabs):
            frame = self.tab_widget.widget(i)
            widget = frame.findChild(QWidget, 'text_box')
            if widget:
                text_title = widget.findChild(QLabel, 'text_title')
                text_edit = widget.findChild(QTextEdit, 'text_box_text_edit')
                text_edit.setText(self.sermon_text_edit.toPlainText())
                text_title.setText(self.sermon_reference_field.text())

        self.toolbar.id_label.setText('ID: ' + str(record[0][0]))

        self.apply_line_spacing()

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

            num_tabs = self.tab_widget.count()
            for i in range(num_tabs):
                if i > 0:
                    frame = self.tab_widget.widget(i)
                    widget = frame.findChild(QWidget, 'text_box')
                    text_title = widget.findChild(QLabel, 'text_title')
                    text_title.setText(self.sermon_reference_field.text())

            if self.main.user_settings['auto_fill']:
                if not self.gs: # only create one instance of GetScripture
                    self.gs = GetScripture(self.main)

                if ':' in self.sermon_reference_field.text(): # only attempt to get the text if there's enough to work with
                    passage = self.gs.get_passage(self.sermon_reference_field.text())
                    if passage and not passage == -1:
                        self.sermon_text_edit.setText(passage)

            self.changes = True
        except Exception as ex:
            self.main.write_to_log(str(ex))

    def auto_fill(self):
        """
        Method to change the self.spd.auto_fill value based on user input then save that change to the database.
        """
        self.main.user_settings['auto_fill'] = self.auto_fill_checkbox.isChecked()
        self.main.save_user_settings()
        if self.auto_fill_checkbox.isChecked():
            self.reference_changes()

    def text_changes(self):
        """
        Method to fill the optional sermon text box on each tab with scripture when the Scripture Text text is
        changed
        """
        num_tabs = self.tab_widget.count()
        for i in range(num_tabs):
            if i > 0:
                frame = self.tab_widget.widget(i)
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

    def do_exit(self, evt):
        """
        Method to ask if changes are to be saved before exiting the program.
        """
        goon = True
        if self.changes:
            goon = self.main.ask_save()
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
            self.main.save_rec()
        elif event.modifiers() & Qt.KeyboardModifier.ControlModifier and event.key() == Qt.Key.Key_P:
            self.menu_bar.print_rec()
        elif event.modifiers() & Qt.KeyboardModifier.ControlModifier and event.key() == Qt.Key.Key_Q:
            self.close()
        event.accept()
