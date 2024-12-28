import logging
import os
import re
import shutil
import sys
from os.path import exists

import wmi
from PyQt6.QtCore import Qt, QSizeF, QRectF
from PyQt6.QtGui import QStandardItemModel, QFontDatabase, QStandardItem, QPixmap, QTextCursor, QFont, QIcon, \
    QTextDocument, QTextOption, QPainter
from PyQt6.QtPrintSupport import QPrinter
from PyQt6.QtWidgets import QFileDialog, QWidget, QVBoxLayout, QLabel, QTableView, QPushButton, \
    QTabWidget, QHBoxLayout, QComboBox, QTextBrowser, QLineEdit, QTextEdit, QDateEdit, QMessageBox, QSpinBox
from pynput.keyboard import Key, Controller

from runnables import LoadDictionary, SpellCheck
from toolbar import Toolbar


class MenuBar:
    """
    Builds the QMainWindow's menuBar and handles the actions performed by the menu.
    """
    def __init__(self, gui, spd):
        """
        :param GUI gui: The program's GUI Object
        :param SermonPrepDatabase spd: The program's SermonPrepDatabase object
        """
        self.keyboard = Controller()
        self.gui = gui
        self.spd = spd
        menu_bar = self.gui.menuBar()

        file_menu = menu_bar.addMenu('File')
        file_menu.setToolTipsVisible(True)

        save_action = file_menu.addAction('Save (Ctrl-S)')
        save_action.triggered.connect(self.spd.save_rec)

        print_action = file_menu.addAction('Print (Ctrl-P)')
        print_action.setToolTip('Print the current record')
        print_action.triggered.connect(self.print_rec)

        file_menu.addSeparator()

        backup_action = file_menu.addAction('Create Backup')
        backup_action.setToolTip('Manually save a backup of your database')
        backup_action.triggered.connect(self.do_backup)

        restore_action = file_menu.addAction('Restore from Backup')
        restore_action.setToolTip('Restore a previous backup of your database')
        restore_action.triggered.connect(self.restore_backup)

        file_menu.addSeparator()

        import_action = file_menu.addAction('Import Sermons from Files')
        import_action.setToolTip('Import sermons that have been saved as .docx, .odt, or .txt')
        import_action.triggered.connect(self.import_from_files)

        bible_action = file_menu.addAction('Import Zefania XML Bible')
        bible_action.setToolTip('Import a bible file saved in the Zefania XML format to use with your program')
        bible_action.triggered.connect(self.import_bible)

        file_menu.addSeparator()

        exit_action = file_menu.addAction('Exit (Ctrl-Q)')
        exit_action.setToolTip('Quit the program')
        exit_action.triggered.connect(self.do_exit)

        edit_menu = menu_bar.addMenu('Edit')
        edit_menu.setToolTipsVisible(True)

        cut_action = edit_menu.addAction('Cut (Ctrl-X)')
        cut_action.triggered.connect(self.press_ctrl_x)

        copy_action = edit_menu.addAction('Copy (Ctrl-C)')
        copy_action.triggered.connect(self.press_ctrl_c)

        paste_action = edit_menu.addAction('Paste (Ctrl-V)')
        paste_action.triggered.connect(self.press_ctrl_v)

        edit_menu.addSeparator()

        config_menu = edit_menu.addMenu('Configure')
        config_menu.setToolTipsVisible(True)

        color_menu = config_menu.addMenu('Change Theme')

        red_color_action = color_menu.addAction('Wine')
        red_color_action.triggered.connect(lambda: self.color_change('red'))

        green_color_action = color_menu.addAction('Forest')
        green_color_action.triggered.connect(lambda: self.color_change('green'))

        blue_color_action = color_menu.addAction('Sky')
        blue_color_action.triggered.connect(lambda: self.color_change('blue'))

        yellow_color_action = color_menu.addAction('Leather')
        yellow_color_action.triggered.connect(lambda: self.color_change('gold'))

        surf_color_action = color_menu.addAction('Surf')
        surf_color_action.triggered.connect(lambda: self.color_change('surf'))

        royal_color_action = color_menu.addAction('Royal')
        royal_color_action.triggered.connect(lambda: self.color_change('royal'))

        dark_color_action = color_menu.addAction('Dark')
        dark_color_action.triggered.connect(lambda: self.color_change('dark'))

        """custom_color_menu = color_menu.addMenu('Custom Colors')

        bg_color_action = custom_color_menu.addAction('Change Accent Color')
        bg_color_action.setToolTip('Choose a different color for accents and borders')
        bg_color_action.triggered.connect(lambda: self.color_change('bg'))

        fg_color_action = custom_color_menu.addAction('Change Background Color')
        fg_color_action.setToolTip('Choose a different color for the background')
        fg_color_action.triggered.connect(lambda: self.color_change('fg'))"""

        font_action = config_menu.addAction('Change Font')
        font_action.setToolTip('Change the font and font size used in the program')
        font_action.triggered.connect(self.change_font)

        line_spacing_menu = config_menu.addMenu('Change Line Spacing')
        line_spacing_menu.setToolTip('Increase or decrease the line spacing of the text')

        compact_spacing_action = line_spacing_menu.addAction('Compact')
        compact_spacing_action.setObjectName('compact')
        compact_spacing_action.setCheckable(True)
        compact_spacing_action.setChecked(False)
        compact_spacing_action.setToolTip('Set line spacing to compact')
        compact_spacing_action.triggered.connect(lambda: self.change_line_spacing('compact'))

        regular_spacing_action = line_spacing_menu.addAction('Regular')
        regular_spacing_action.setObjectName('regular')
        regular_spacing_action.setCheckable(True)
        regular_spacing_action.setChecked(False)
        regular_spacing_action.setToolTip('Set line spacing to regular')
        regular_spacing_action.triggered.connect(lambda: self.change_line_spacing('regular'))

        wide_spacing_action = line_spacing_menu.addAction('Wide')
        wide_spacing_action.setObjectName('wide')
        wide_spacing_action.setCheckable(True)
        wide_spacing_action.setChecked(False)
        wide_spacing_action.setToolTip('Set line spacing to wide')
        wide_spacing_action.triggered.connect(lambda: self.change_line_spacing('wide'))

        config_menu.addSeparator()

        rename_action = config_menu.addAction('Rename Labels')
        rename_action.setToolTip('Rename the labels in this program')
        rename_action.triggered.connect(self.rename_labels)

        config_menu.addSeparator()

        remove_words_item = config_menu.addAction('Remove custom words from dictionary')
        remove_words_item.setToolTip('Choose words to remove from the dictionary that have been added by you')
        remove_words_item.triggered.connect(self.remove_words)

        self.disable_spell_check_action = config_menu.addAction('Disable Spell Check')
        self.disable_spell_check_action.setToolTip('Disabling spell check improves performance and memory usage')
        self.disable_spell_check_action.setCheckable(True)
        if self.spd.disable_spell_check:
            self.disable_spell_check_action.setChecked(True)
        else:
            self.disable_spell_check_action.setChecked(False)
        self.disable_spell_check_action.triggered.connect(self.disable_spell_check)

        record_menu = menu_bar.addMenu('Record')
        record_menu.setToolTipsVisible(True)

        first_rec_action = record_menu.addAction('Jump to First Record')
        first_rec_action.triggered.connect(self.spd.first_rec)

        prev_rec_action = record_menu.addAction('Go to Previous Record')
        prev_rec_action.triggered.connect(self.spd.prev_rec)

        next_rec_action = record_menu.addAction('Go to Next Record')
        next_rec_action.triggered.connect(self.spd.next_rec)

        last_rec_action = record_menu.addAction('Jump to Last Record')
        last_rec_action.triggered.connect(self.spd.last_rec)

        record_menu.addSeparator()

        new_rec_action = record_menu.addAction('Create New Record')
        new_rec_action.triggered.connect(self.spd.new_rec)

        del_rec_action = record_menu.addAction('Delete Current Record')
        del_rec_action.triggered.connect(self.spd.del_rec)

        help_menu = menu_bar.addMenu('Help')

        help_action = help_menu.addAction('Help Topics')
        help_action.triggered.connect(self.show_help)

        about_action = help_menu.addAction('About')
        about_action.triggered.connect(self.show_about)

    def print_rec(self):
        PrintHandler(self.gui)

    def do_backup(self):
        """
        Creates a QFileDialog where the user can save a custom backup of their database
        """
        try:
            import os
            user_dir = os.path.expanduser('~')

            fileName = QFileDialog.getSaveFileName(self.gui, 'Create Backup',
                                                   user_dir + '/sermon_prep_database_backup.db', 'Database File (*.db)')
            if len(fileName[0]) == 0:
                return

            import shutil
            shutil.copy(self.spd.db_loc, fileName[0])
            self.spd.write_to_log('Created Backup as ' + fileName[0])

            QMessageBox.information(
                None,
                'Backup Created',
                'Backup successfully created as ' + fileName[0],
                QMessageBox.StandardButton.Ok
            )
        # make this more precise you lazy turd
        except Exception as ex:
            self.spd.write_to_log('There was a problem creating the backup:\n\n' + str(ex), True)

    def restore_backup(self):
        """
        Method to restore the user's database from a backup file. Creates a QFileDialog for the user to choose
        their backup then copies that backup to the user's app data location.
        """
        dialog = QFileDialog()
        dialog.setWindowTitle('Restore from Backup')
        dialog.setNameFilter('Database File (*.db)')
        dialog.setFileMode(QFileDialog.FileMode.ExistingFile)
        dialog.setDirectory(self.spd.app_dir)

        dialog.exec()

        if dialog.selectedFiles():
            db_file = dialog.selectedFiles()[0]
            import shutil
            shutil.copy(self.spd.db_loc, self.spd.app_dir + '/active-database-backup.db')
            os.remove(self.spd.db_loc)
            shutil.copy(db_file, self.spd.db_loc)

            QMessageBox.information(
                None,
                'Attempting Restore',
                'The program will now attempt to restore the data from your backup.',
                QMessageBox.StandardButton.Ok
            )

            try:
                self.spd.get_ids()
                self.spd.get_date_list()
                self.spd.get_scripture_list()
                self.spd.get_user_settings()
                self.gui.toolbar.dates_cb.addItems(self.spd.dates)
                for item in self.spd.references:
                    self.gui.toolbar.references_cb.addItem(item[0])
                self.spd.last_rec()
            except Exception as err:
                shutil.copy(self.spd.app_dir + '/active-database-backup.db', self.spd.db_loc)
                self.spd.write_to_log('MenuBar.restore_backup: ' + str(err), True)

                QMessageBox.critical(
                    None,
                    'Error Loading Database Data',
                    'There was a problem loading the data from your backup. Your prior database has not been changed.  '
                    'The error is as follows:\r\n' + str(err),
                    QMessageBox.StandardButton.Ok
                )

                self.spd.get_ids()
                self.spd.get_date_list()
                self.spd.get_scripture_list()
                self.spd.get_user_settings()
                self.gui.toolbar.dates_cb.addItems(self.spd.dates)
                for item in self.spd.references:
                    self.gui.references_cb.addItem(item[0])
                self.spd.last_rec()
            else:
                QMessageBox.information(
                    None,
                    'Backup Restored',
                    'Backup successfully restored.\n\nA copy of your prior database has been saved as ' + self.spd.app_dir
                    + '/active-database-backup.db',
                    QMessageBox.StandardButton.Ok
                )

    def import_from_files(self):
        """
        Method to inform user about the best format for imported file names and to begin the import by calling
        GetFromDocx.
        """
        QMessageBox.information(
            self.gui,
            'Import from Files',
            'For best results, the files you are importing should be named according to this syntax:\n\n'
            'YYYY-MM-DD.book.chapter.verse-verse\n\n'
            'For example, a sermon preached on May 20th, 2011 on Mark 3:1-12, saved as a Microsoft Word document,'
            'would be named:\n\n'
            '2011-05-11.mark.3.1-12.docx', QMessageBox.StandardButton.Ok)
        from get_from_docx import GetFromDocx
        GetFromDocx(self.gui)

    def import_bible(self):
        """
        Method to import and save a user's xml bible for use in the program.
        """
        file = QFileDialog.getOpenFileName(
            self.gui,
            'Choose Bible File',
            os.path.expanduser('~'),
            'XML Bible File (*.xml)'
        )
        try:
            if file[0]:
                shutil.copy(file[0], self.spd.app_dir + '/my_bible.xml')
                self.spd.bible_file = self.spd.app_dir + '/my_bible.xml'

                # verify the file by attempting to get a passage from the new file
                from get_scripture import GetScripture
                self.gui.gs = GetScripture(self.spd)
                passage = self.gui.gs.get_passage('John 3:16')

                if not passage or passage == -1 or passage == '':
                    QMessageBox.warning(
                        self.gui,
                        'Bad Format',
                        'There is a problem with your XML bible: ' + file[0] + '. Try downloading it again or ensuring '
                        'that it is formatted according to Zefania standards.',
                        QMessageBox.StandardButton.Ok
                    )

                    # we're just not going to worry about the option to have multiple bibles
                    if exists(self.spd.app_dir + '/my_bible.xml'):
                        os.remove(self.spd.app_dir + '/my_bible.xml')
                else:
                    QMessageBox.information(
                        self.gui,
                        'Import Complete',
                        'Bible file has been successfully imported',
                        QMessageBox.StandardButton.Ok
                    )
                    try:
                        self.gui.tabbed_frame.removeTab(0)
                        self.gui.build_scripture_tab(True)
                        self.gui.auto_fill_checkbox.setChecked(True)
                        self.gui.set_style_sheets()
                        self.gui.tabbed_frame.setCurrentIndex(0)
                        self.spd.get_by_index(self.spd.current_rec_index)
                    except Exception:
                        logging.exception('')

        except Exception as ex:
            self.spd.write_to_log(str(ex))
            QMessageBox.warning(
                self.gui,
                'Import Error',
                'An error occurred while importing the file ' + file[0] + ':\n\n' + str(ex),
                QMessageBox.StandardButton.Ok
            )

            if exists(self.spd.app_dir + '/my_bible.xml'):
                os.remove(self.spd.app_dir + '/my_bible.xml')

    def rename_labels(self):
        """
        Method to allow the user to change the text of heading labels used in the program (i.e. 'Sermon Text Reference'
        or 'Fallen Condition Focus of the Text'.
        """
        self.rename_widget = QWidget()
        self.rename_widget.setWindowFlag(Qt.WindowType.Window)
        self.rename_widget.resize(920, 400)
        rename_widget_layout = QVBoxLayout()
        self.rename_widget.setLayout(rename_widget_layout)

        rename_label = QLabel('Use this table to set new names for any of the labels in this program.\n'
                             'Double-click a label under "New Label" to rename it')
        rename_widget_layout.addWidget(rename_label)

        # provide two columns in the QTableView; one that will contain the labels as they are, the other to provide
        # an area to change the label's name
        model = QStandardItemModel(len(self.spd.user_settings), 2)
        for i in range(1, 22):
            item = QStandardItem(self.spd.user_settings[f'label{i}'])
            item.setEditable(False)
            model.setItem(i, 0, item)
            item2 = QStandardItem(self.spd.user_settings[f'label{i}'])
            model.setItem(i, 1, item2)
        model.setHeaderData(0, Qt.Orientation.Horizontal, 'Current Label')
        model.setHeaderData(1, Qt.Orientation.Horizontal, 'New Label')

        rename_table_view = QTableView()
        rename_table_view.setModel(model)
        rename_table_view.setColumnWidth(0, 400)
        rename_table_view.setColumnWidth(1, 400)
        rename_widget_layout.addWidget(rename_table_view)

        save_button = QPushButton('Save Changes')
        save_button.clicked.connect(lambda: self.write_label_changes(model))
        rename_widget_layout.addWidget(save_button)

        self.rename_widget.show()

    def write_label_changes(self, model):
        """
        Method to write the user's heading label changes to the database's user_settings table.
        Perhaps this belongs in the SermonPrepDatabase class.

        :param QStandardItemModel model: The QStandardItemModel that contains the user's new label names.
        """
        for i in range(model.rowCount()):
            self.spd.user_settings[f'label{i + 1}'] = model.data(model.index(i, 1))
        print(self.spd.user_settings)
        return
        self.spd.save_user_settings()

        self.rename_widget.destroy()

        for widget in self.gui.main_widget.children():
            if isinstance(widget, QWidget):
                self.gui.layout.removeWidget(widget)

        self.gui.menu_bar = self
        self.gui.toolbar = Toolbar(self.gui, self.spd)
        self.gui.layout.addWidget(self.gui.toolbar)
        self.gui.build_tabbed_frame()
        self.gui.build_scripture_tab()
        self.gui.build_exegesis_tab()
        self.gui.build_outline_tab()
        self.gui.build_research_tab()
        self.gui.build_sermon_tab()
        self.gui.set_style_sheets()

        self.spd.get_by_index(self.spd.current_rec_index)

    def color_change(self, type):
        """
        Method to apply the user's chosen theme or custom colors to the GUI.
        :param str type: The theme or custom color chosen by the user.
        """
        if type == 'red':
            style_sheet = open('resources/style_sheets/spd-red.qss').read()
            self.gui.setStyleSheet(style_sheet)
        elif type == 'green':
            style_sheet = open('resources/style_sheets/spd-green.qss').read()
            self.gui.setStyleSheet(style_sheet)
        elif type == 'blue':
            style_sheet = open('resources/style_sheets/spd-blue.qss').read()
            self.gui.setStyleSheet(style_sheet)
        elif type == 'gold':
            style_sheet = open('resources/style_sheets/spd-leather.qss').read()
            self.gui.setStyleSheet(style_sheet)
        elif type == 'surf':
            style_sheet = open('resources/style_sheets/spd-surf.qss').read()
            self.gui.setStyleSheet(style_sheet)
        elif type == 'royal':
            style_sheet = open('resources/style_sheets/spd-royal.qss').read()
            self.gui.setStyleSheet(style_sheet)
        elif type == 'dark':
            style_sheet = open('resources/style_sheets/spd-dark.qss').read()
            self.gui.setStyleSheet(style_sheet)
        else:
            type = 'blue'
            style_sheet = open('resources/style_sheets/spd-blue.qss').read()
            self.gui.setStyleSheet(style_sheet)

        if type == 'dark':
            self.gui.toolbar.undo_button.setIcon(QIcon('resources/svg/spUndoIconDark.svg'))
            self.gui.toolbar.redo_button.setIcon(QIcon('resources/svg/spRedoIconDark.svg'))
            self.gui.toolbar.bold_button.setIcon(QIcon('resources/svg/spBoldIconDark.svg'))
            self.gui.toolbar.italic_button.setIcon(QIcon('resources/svg/spItalicIconDark.svg'))
            self.gui.toolbar.underline_button.setIcon(QIcon('resources/svg/spUnderlineIconDark.svg'))
            self.gui.toolbar.bullet_button.setIcon(QIcon('resources/svg/spBulletIconDark.svg'))
            self.gui.toolbar.text_visible.setIcon(QIcon('resources/svg/spShowTextDark.svg'))
            self.gui.toolbar.first_rec_button.setIcon(QIcon('resources/svg/spFirstRecIconDark.svg'))
            self.gui.toolbar.prev_rec_button.setIcon(QIcon('resources/svg/spPrevRecIconDark.svg'))
            self.gui.toolbar.next_rec_button.setIcon(QIcon('resources/svg/spNextRecIconDark.svg'))
            self.gui.toolbar.last_rec_button.setIcon(QIcon('resources/svg/spLastRecIconDark.svg'))
            self.gui.toolbar.new_rec_button.setIcon(QIcon('resources/svg/spNewIconDark.svg'))
            self.gui.toolbar.save_button.setIcon(QIcon('resources/svg/spSaveIconDark.svg'))
            self.gui.toolbar.print_button.setIcon(QIcon('resources/svg/spPrintIconDark.svg'))
            self.gui.sermon_view_button.setIcon(QIcon('resources/svg/spSermonViewIconLight.svg'))
        else:
            self.gui.toolbar.undo_button.setIcon(QIcon('resources/svg/spUndoIcon.svg'))
            self.gui.toolbar.redo_button.setIcon(QIcon('resources/svg/spRedoIcon.svg'))
            self.gui.toolbar.bold_button.setIcon(QIcon('resources/svg/spBoldIcon.svg'))
            self.gui.toolbar.italic_button.setIcon(QIcon('resources/svg/spItalicIcon.svg'))
            self.gui.toolbar.underline_button.setIcon(QIcon('resources/svg/spUnderlineIcon.svg'))
            self.gui.toolbar.bullet_button.setIcon(QIcon('resources/svg/spBulletIcon.svg'))
            self.gui.toolbar.text_visible.setIcon(QIcon('resources/svg/spShowText.svg'))
            self.gui.toolbar.first_rec_button.setIcon(QIcon('resources/svg/spFirstRecIcon.svg'))
            self.gui.toolbar.prev_rec_button.setIcon(QIcon('resources/svg/spPrevRecIcon.svg'))
            self.gui.toolbar.next_rec_button.setIcon(QIcon('resources/svg/spNextRecIcon.svg'))
            self.gui.toolbar.last_rec_button.setIcon(QIcon('resources/svg/spLastRecIcon.svg'))
            self.gui.toolbar.new_rec_button.setIcon(QIcon('resources/svg/spNewIcon.svg'))
            self.gui.toolbar.save_button.setIcon(QIcon('resources/svg/spSaveIcon.svg'))
            self.gui.toolbar.print_button.setIcon(QIcon('resources/svg/spPrintIcon.svg'))
            self.gui.sermon_view_button.setIcon(QIcon('resources/svg/spSermonViewIconDark.svg'))
        for i in range(self.gui.tabbed_frame.count()):
            if i == self.gui.tabbed_frame.currentIndex() and not type == 'dark':
                self.gui.tabbed_frame.setTabIcon(i, self.gui.dark_tab_icons[i])
            else:
                self.gui.tabbed_frame.setTabIcon(i, self.gui.light_tab_icons[i])

        self.spd.user_settings['theme'] = type
        self.spd.save_user_settings()

    def disable_spell_check(self):
        """
        Method to turn on or off the spell checking capabilities of the CustomTextEdit.
        """
        if self.disable_spell_check_action.isChecked():
            self.spd.user_settings['disable_spell_check'] = True
            self.spd.save_user_settings()

            # remove any red formatting created by spell check
            for widget in self.gui.tabbed_frame.findChildren(QTextEdit, 'custom_text_edit'):
                widget.blockSignals(True)

                cursor = widget.textCursor()
                cursor.select(QTextCursor.SelectionType.Document)

                char_format = cursor.charFormat()
                char_format.setForeground(Qt.GlobalColor.black)
                cursor.mergeCharFormat(char_format)
                cursor.clearSelection()

                widget.setTextCursor(cursor)

                widget.blockSignals(False)

        else:
            # if spell check is enabled after having been disabled at startup, the dictionary will need to be loaded
            if not self.spd.sym_spell:
                ld = LoadDictionary(self.spd)
                self.spd.load_dictionary_thread_pool.start(ld)
            self.spd.user_settings['disable_spell_check'] = False
            self.spd.save_user_settings()

            # run spell check on all CustomTextEdits
            for widget in self.gui.tabbed_frame.findChildren(QTextEdit, 'custom_text_edit'):
                spell_check = SpellCheck(widget, 'whole', self.gui)
                self.gui.spd.spell_check_thread_pool.start(spell_check)

    def show_help(self):
        """
        Call the ShowHelp class on user's input
        """
        self.help_widget = QWidget()
        self.help_widget.setObjectName('help_widget')
        self.help_widget.setParent(self.gui)
        self.help_widget.setWindowFlag(Qt.WindowType.Window)
        self.help_widget.setWindowTitle('Help Topics')
        help_layout = QVBoxLayout(self.help_widget)

        help_label = QLabel('Help Topics')
        help_label.setFont(QFont(self.gui.spd.user_settings['font_family'], 20))
        help_layout.addWidget(help_label)

        sh = ShowHelp(self.gui, self.spd)
        help_layout.addWidget(sh)

        self.help_widget.showMaximized()

    def show_about(self):
        """
        Display the 'about' text on user's input
        """
        self.about_win = QWidget()
        about_layout = QVBoxLayout()
        self.about_win.setLayout(about_layout)

        about_label = QLabel('Sermon Prep Database v.5.0.0.005')
        about_layout.addWidget(about_label)

        about_text = QTextBrowser()
        about_text.setHtml('''
            Sermon Prep Database is free software: you can redistribute it and/or
            modify it under the terms of the GNU General Public License (GNU GPL)
            published by the Free Software Foundation, either version 3 of the
            License, or (at your option) any later version.<br><br>
    
            This program is distributed in the hope that it will be useful,
            but WITHOUT ANY WARRANTY; without even the implied warranty of
            MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
            GNU General Public License for more details.<br><br>
    
            You should have received a copy of the GNU General Public License
            along with this program.  If not, see <a href="http://www.gnu.org/licenses/">http://www.gnu.org/licenses/</a>.<br><br>
    
            The Sermon Prep Database program includes Artifex Software's GhostScript,
            licensed under the GNU Affero General Public License (GNU AGPL). See
            <a href="https://www.ghostscript.com/licensing/index.html">https://www.ghostscript.com/licensing/index.html</a> for more information.<br><br>
    
            This program is a work-in-progress by a guy who is not, in no way, a
            professional programmer. If you run into any problems, unexpected behavior,
            missing features, or attempts to assimilate your unique biological and
            technological distinctiveness, email <a href="mailto:pastorjeremywilson@gmail.com">pastorjeremywilson@gmail.com</a>
        ''')
        about_text.setOpenExternalLinks(True)
        about_text.setReadOnly(True)
        about_layout.addWidget(about_text)

        self.about_win.show()

    def remove_words(self):
        """
        Instantiate the RemoveCustomWords class from dialogs.py
        """
        from dialogs import RemoveCustomWords
        self.rcm = RemoveCustomWords(self.spd)

    def change_font(self):
        """
        Method to provide a font choosing widget to the user by obtaining a list of all the fonts available to
        the system.
        """
        current_font = self.spd.user_settings['font_family']
        current_size = self.spd.user_settings['font_size']

        global font_chooser
        font_chooser = QWidget()
        font_layout = QVBoxLayout()
        font_chooser.setLayout(font_layout)

        top_panel = QWidget()
        top_layout = QHBoxLayout()
        top_panel.setLayout(top_layout)
        font_layout.addWidget(top_panel)

        family_combo_box = FontFaceComboBox(self.gui)
        family_combo_box.setCurrentText(current_font)
        family_combo_box.currentIndexChanged.connect(
            lambda: self.gui.apply_font(family_combo_box.currentText(), size_combo_box.currentText()))
        top_layout.addWidget(family_combo_box)

        size_combo_box = QComboBox()
        sizes = ['10', '12', '14', '16', '18', '20', '24', '28', '32']
        size_combo_box.addItems(sizes)
        size_combo_box.currentIndexChanged.connect(
            lambda: self.gui.apply_font(family_combo_box.currentText(), size_combo_box.currentText()))
        top_layout.addWidget(size_combo_box)
        size_combo_box.setCurrentText(current_size)

        bottom_panel = QWidget()
        bottom_layout = QHBoxLayout()
        bottom_panel.setLayout(bottom_layout)
        font_layout.addWidget(bottom_panel)

        ok_button = QPushButton('OK')
        ok_button.clicked.connect(
            lambda: self.gui.apply_font(family_combo_box.currentText(), size_combo_box.currentText(), font_chooser, True))
        bottom_layout.addWidget(ok_button)

        cancel_button = QPushButton('Cancel')
        cancel_button.clicked.connect(lambda: self.gui.apply_font(current_font, current_size, font_chooser))
        bottom_layout.addWidget(cancel_button)

        font_chooser.show()

    def change_line_spacing(self, spacing):
        """
        Method to change the line spacing in the custom text edits. Since the data needs to be reloaded for the changes
        to be effected, ask for save first.
        """
        if spacing == 'compact':
            self.spd.user_settings['line_spacing'] = '1.0'
        elif spacing == 'regular':
            self.spd.user_settings['line_spacing'] = '1.2'
        elif spacing == 'wide':
            self.spd.user_settings['line_spacing'] = '1.5'

        self.gui.apply_line_spacing()
        self.spd.save_user_settings()

    def press_ctrl_z(self):
        """
        Method to programmatically execute this key combination
        """
        self.keyboard.press(Key.ctrl)
        self.keyboard.press('z')
        self.keyboard.release('z')
        self.keyboard.release(Key.ctrl)

    def press_ctrl_y(self):
        """
        Method to programmatically execute this key combination
        """
        self.keyboard.press(Key.ctrl)
        self.keyboard.press('y')
        self.keyboard.release('y')
        self.keyboard.release(Key.ctrl)

    def press_ctrl_x(self):
        """
        Method to programmatically execute this key combination
        """
        self.keyboard.press(Key.ctrl)
        self.keyboard.press('x')
        self.keyboard.release('x')
        self.keyboard.release(Key.ctrl)

    def press_ctrl_c(self):
        """
        Method to programmatically execute this key combination
        """
        self.keyboard.press(Key.ctrl)
        self.keyboard.press('c')
        self.keyboard.release('c')
        self.keyboard.release(Key.ctrl)

    def press_ctrl_v(self):
        """
        Method to programmatically execute this key combination
        """
        self.keyboard.press(Key.ctrl)
        self.keyboard.press('v')
        self.keyboard.release('v')
        self.keyboard.release(Key.ctrl)

    def do_exit(self):
        """
        Method to ask if changes should be saved, then exit the program.
        """
        goon = True
        if self.gui.changes:
            goon = self.spd.ask_save()
        if goon:
            sys.exit(0)


class ShowHelp(QTabWidget):
    """
    Method to create a QTabbedWidget to hold the different help topics.
    """
    def __init__(self, gui, spd):
        super().__init__()
        self.gui = gui
        self.spd = spd

        self.setObjectName('help_tab_widget')
        self.tabBar().setObjectName('help_tab_widget')

        self.resize(1200, 800)
        self.make_intro()
        self.make_menu()
        self.make_tool()
        self.make_scripture()
        self.make_exeg()
        self.make_outline()
        self.make_research()
        self.make_sermon()

    def make_intro(self):
        """
        Method to create the intro widget for the help tabs.
        """
        intro_widget = QWidget()
        intro_layout = QVBoxLayout()
        intro_widget.setLayout(intro_layout)

        intro_label = QLabel('Introduction')
        intro_layout.addWidget(intro_label)

        intro_text = QTextBrowser()
        intro_text.setFont(self.gui.standard_font)
        intro_text.setText(
            u'Thank-you for trying out this Sermon Prep Database. Its purpose is to provide an easy-to-use program to '
            'organize and store the many different facets of the sermon preparation process. Whatever information you '
            'save in this program is stored on your hard drive in the widely-used SQLite Database format. This allows '
            'for secure storage and quick retrieval of your important information.<br><br>Should you ever need to find '
            'the database file that stores all your sermons, it is located at'
            '<br>C:\\Users\\YourUserName\\AppData\\Roaming\\Sermon Prep Database\\sermon_prep_database.db<br><br>'
            'The program is laid out primarily in a tab-style window. To the left of the window are tabs that you can '
            'use to navigate the different categories of information you may want to store. As each tab is selected, '
            'the contents of the window will change to reflect that category. At the top of the screen, you will find '
            'a tool bar that contains useful buttons and other tools you may want to make use of as you use the '
            'program.<br><br> It is my hope and prayer that you will find this program both a useful and valuable '
            'tool.<br><br>This program is a work-in-progress by a guy who is not, in no way, a professional '
            'programmer. If you run into any problems, unexpected behavior, missing features, or attempts to '
            'assimilate your unique biological and technological distinctiveness, email '
            '<a href='"'mailto:pastorjeremywilson@gmail.com'"'>pastorjeremywilson@gmail.com</a>'
        )
        intro_text.setOpenExternalLinks(True)
        intro_text.setMinimumWidth(750)
        intro_text.setReadOnly(True)
        intro_layout.addWidget(intro_text)

        self.addTab(intro_widget, 'Introduction')

    def make_menu(self):
        """
        Method to create the menu widget for the help tabs.
        """
        menu_widget = QWidget()
        menu_layout = QVBoxLayout()
        menu_widget.setLayout(menu_layout)

        menu_label = QLabel('Menu Bar')
        menu_label.setFont(self.gui.bold_font)
        menu_layout.addWidget(menu_label)
        menu_pic = QPixmap('resources/menuPic.png')
        menu_pic_label = QLabel()
        menu_pic_label.setPixmap(menu_pic)
        menu_layout.addWidget(menu_pic_label)

        menu_text = QTextBrowser()
        menu_text.setHtml(
            'At the very top of the screen is a menu bar containing the File, Edit, Record, and Help menus<br><br>'
            '<strong>File</strong><br>In the File menu, you will find <strong>Save</strong> and <strong>Print'
            '</strong> commands that perform the same functions as the corresponding buttons, as well as an <strong>'
            'Exit</strong> command. Here, you will also find a <strong>Create Backup</strong> command. While the '
            'program automatically saves a backup of your database every time it is opened, there may be times where '
            'you would like to manually create a backup. Choose this option and navigate to the folder where you want '
            'to save your backup.<br><br>Should you accidentally delete a record, or if you need to restore from a '
            'backup for any reason, choose the option to <strong>Restore from Backup</strong>. This will open a file '
            'dialog where you can choose the backup file you would like to restore.<br><br>You also have an option to '
            '<strong>import</strong> sermons from Microsoft Word .docx files, LibreOffice/OpenOffice .odt '
            'files, or plain text .txt files.<br><br>This function has the ability to import the sermon\'s date, '
            'scripture reference, and sermon manuscript into the Sermon Prep Database. In order to do so, a little '
            'legwork is needed. It is best to copy all of your sermon files into one folder, as this will go through '
            'all the files in the folder you choose, pulling the manuscript from each file. In order to also import '
            'the date and reference for each sermon, the files should be named in a particular way. The file name '
            'should start with the date in <strong>YYYY-MM-DD</strong> format folowed by a period. After the period, '
            'the scripture reference follows in the format book.chapter.verse-verse. Finally is the proper suffix '
            'for your file type (.docx, .odt, or .txt).<br><br>For example, a sermon preached on Sunday, March 19th, '
            '2024 on Mark 9:1-12, saved as a Microsoft Word document would have the file name <strong>2024-03-19.mark.'
            '9.1-12.docx</strong>. If the file names aren\'t renamed in this way, the program can still import your '
            'sermons, it just won\'t be able to automatically save the corresponding date and reference information '
            'as well<br><br>Next, if you have your favorite bible downloaded as a Zefaniah XML file '
            '(<a href="https://sourceforge.net/projects/zefania-sharp/files/Bibles/">'
            'https://sourceforge.net/projects/zefania-sharp/files/Bibles/</a>), you can import that file '
            'into the program by clicking <strong>Import XML Bible</strong>. This will allow the program to '
            'automatically insert your sermon text into the "Sermon Text" area of the "Scripture" tab. This has only '
            'been tested with Zephania bible files in particular, but others may work.<br><br>'
            '<strong>Edit</strong><br>In the Edit menu, you will find the customary <strong>Cut</strong>, <strong>Copy'
            '</strong>, and <strong>Paste</strong> commands. Again, <strong>Ctrl-X</strong>, <strong>Ctrl-C</strong>, '
            'and <strong>Ctrl-V</strong> will also perform these same functions.<br><br>In this menu you will '
            'also find a <strong>Configure</strong> menu. This contains commands to change the colors used in the '
            'program as well as what the labels say, the font that is used, and the line spacing of the editors.'
            '<br><br>First, you\'ll see a menu that lets you set the color theme of the program, or to change '
            'the accent and background colors of the program to anything you\'d like. Changing the accent color will '
            'affect the background color of the left-side tabs, while changing the background colors will change the '
            'color surrounding the different boxes where you enter information.<br><br>Under this, there is a command '
            'to <strong>Change Font</strong>. '
            'This allows you to choose a custom font and font size for the labels and entry boxes throughout the '
            'program. You may have a particular font or size that you find easier to read, and you can change to that '
            'here. Following that, there is a menu where you can set the <strong>Line Spacing</strong> of the editing '
            'boxes options are \'compact\', \'regular\', and \'wide\'.<br><br>Next, the Edit menu contains a '
            'command to <strong>Rename Labels</strong>. Choosing this will open up a new window where you can '
            'customize the labels that appear above each data entry box. This new window will have two columns, the '
            'first showing the current label, and the second being where you can type in what you would rather have '
            'that label say. For example, if you like to use, "Big Idea of the Text," instead of, "Central Proposition '
            'of the Text," you would click inside the right-hand "Central Proposition of the Text" box and change it '
            'to read, "Big Idea of the Text."<br><br>The last two items in the Edit menu have to do with spell-checking. '
            'As you work with the program, you may find yourself adding words to the dictionary that spell check thinks '
            'are wrong (this is done by right-clicking a red, spell-checked word and choosing "Add to Dictionary"). '
            'should you want to remove any of these added words later, choose <strong>Remove custom words from '
            'dictionary</strong>. A window will pop up where you can select from the list of added words to remove. '
            'Following this is an option to <strong>Disable Spell Check</strong>. You may want to do this if you don\'t'
            'need spell-checking, or if you are on an older computer where performance is an issue.'
            '<br><br><strong>Record</strong><br>In the Record menu are the same navigation controls that appear '
            'in the upper bar of the window. In addition is a <strong>Delete Record</strong> option. Use this if you '
            'are currently viewing a record you would like to delete; if it\'s a duplicate or unneeded, for example. '
            'Deleting a record cannot be undone, so you will be prompted to make sure you really want to delete the '
            'current record.'
        )
        menu_text.setFont(self.gui.standard_font)
        menu_text.setReadOnly(True)
        menu_text.setOpenExternalLinks(True)
        menu_layout.addWidget(menu_text)

        self.addTab(menu_widget, 'Menu Bar')

    def make_tool(self):
        """
        Method to create the toolbar widget for the help tabs.
        """
        tool_widget = QWidget()
        tool_layout = QVBoxLayout()
        tool_widget.setLayout(tool_layout)

        tool_label = QLabel('Toolbar')
        tool_label.setFont(self.gui.bold_font)
        tool_layout.addWidget(tool_label)

        tool_text = QTextEdit()
        tool_text.setHtml(
            'Along the top of the window, you\'ll see a toolbar containing different formatting and navigation tools.'
            '<br><br><img src="resources/undoPic.png"><br>'
            'First, you\'ll see Undo and Redo buttons that you can use to undo or redo any editing you perform. Of '
            'course, Ctrl-Z and Ctrl-Y will also perform these functions.<br><br><img src="'
            'resources/formatPic.png"><br>Beside those are the buttons you can use to format your text with bold, '
            'italic, underline, and bullet-point options.<br><br><img src="'
            'resources/showScripturePic.png"><br>Right beside these is an option to show the sermon text on all tabs. '
            'If you would like to keep your sermon text handy while you are entering, for example, your exegesis notes '
            'or research notes, click this icon. The sermon text will be shown to the right of the Exegesis, Outline, '
            'Research, and Sermon tabs.<br><br><img src="resources/searchPic.png"><br>In the '
            'middle of this bar are two pull-down menus where you can bring up past sermons by the sermon date '
            'or the sermon\'s scripture. This becomes increasingly useful as you add more and more sermons to the '
            'database, allowing you to see what you\'ve preached on in the past or how you\'ve dealt with particular '
            'texts.<br><br>Next to these is a search box. By entering a passage or keyword into this box and pressing '
            'Enter, you can search for scripture passages or words you\'ve used in past sermons.<br><br>To search for '
            'phrases, use double quotes. For example, "empty tomb" or "Matthew 1:". The search will create a new tab '
            'showing any records where your search term(s) appear, and double-clicking any of the results '
            'will bring up that particular sermon\'s record. The search results are sorted with exact matches first,'
            ' followed by results in order of how many search terms were found. This new tab can be closed by pressing '
            'the "X" icon.<br><br><img src="resources/navPic.png"><br>To the right of this upper bar are the navigation '
            'buttons, as well as the save and print buttons. These '
            'buttons will allow you to navigate to the first or last sermons in your database or switch to the '
            'previous or next sermons. Just to the right of these navigation buttons is the "New Record" button. '
            'This will create a new, blank record to start recording information for your next sermon.<br><br>After '
            'this is the save and print buttons. The save button will save all your current info to the database, '
            'while the print button will print a basic hard-copy of the information you have for the currently '
            'showing sermon.'
        )
        tool_text.setFont(self.gui.standard_font)
        tool_text.setReadOnly(True)
        tool_layout.addWidget(tool_text)

        self.addTab(tool_widget, 'Toolbar')

    def make_scripture(self):
        """
        Method to create the scripture widget for the help tabs.
        """
        scripture_widget = QWidget()
        scripture_layout = QVBoxLayout()
        scripture_widget.setLayout(scripture_layout)
        scripture_label = QLabel('Scripture Tab')
        scripture_label.setFont(self.gui.bold_font)
        scripture_layout.addWidget(scripture_label)

        scripture_text = QTextEdit()
        scripture_text.setHtml(
            'The scripture tab contains information about the week\'s pericope and the text you\'ll use for the '
            'sermon. If you don\'t follow the pericopes, you can store any texts you\'ll be using for the week. You '
            'can always change the headings of any of the entries on the scripture tab by selecting "Rename Labels" '
            'under the "Edit" menu.<br><br>The first box you can enter text into is the "Pericope" box. An example of '
            'what goes here would be, "Second Sunday of Easter". Below that is where you can enter the recommended '
            'texts for that Pericope.<br><br>Next, you can enter the passage of the text you\'ll be using for your '
            'sermon, entering the text of that passage underneath.<br><br>If you have previously imported a Zefania '
            'XML bible file, the text of the passage you typed in will be automatically filled in. To turn this '
            'feature off, simply uncheck the box labeled "Auto-fill ' + self.spd.user_settings[8] + '".'
        )
        scripture_text.setFont(self.gui.standard_font)
        scripture_text.setReadOnly(True)
        scripture_layout.addWidget(scripture_text)

        self.addTab(scripture_widget, 'Scripture Tab')

    def make_exeg(self):
        """
        Method to create the exegesis widget for the help tabs.
        """
        exeg_widget = QWidget()
        exeg_layout = QVBoxLayout()
        exeg_widget.setLayout(exeg_layout)
        exeg_label = QLabel('Exegesis Tab')
        exeg_label.setFont(self.gui.bold_font)
        exeg_layout.addWidget(exeg_label)

        exeg_text = QTextEdit()
        exeg_text.setHtml(
            'It is in the Exegesis where you will record the work of exegeting your text.<br><br>On the left is where '
            'you will consider the text itself. What is the Fallen Condition Focus of the text? That is, is there a '
            'particular sin problem being addressed in the text? Is it a specific sin, or a result of the world being '
            'broken by sin? Then, what is the gospel answer of the text? That is, how has the person and work of '
            'Christ remedied that sin condition? Finally, what is the central proposition of the text? That is, what '
            'is the "big idea" that the text is communicating?<br><br>Next comes the Purpose Bridge. This is a '
            'statement that usually starts with, "I want my congregation to..." This is about exegeting your '
            'congregation. What does this text have to say, in particular, to your congregation?<br><br>The, on the '
            'right, are the same ideas that you gleaned from the text, but with an eye to your sermon and your '
            'congregation. What is the sin problem that the sermon will address? What is the gospel answer to that '
            'sin problem? What is the big idea that your sermon is going to convey?'
        )
        exeg_text.setFont(self.gui.standard_font)
        exeg_text.setReadOnly(True)
        exeg_layout.addWidget(exeg_text)

        self.addTab(exeg_widget, 'Exegesis Tab')

    def make_outline(self):
        """
        Method to create the outlines widget for the help tabs.
        """
        outline_widget = QWidget()
        outline_layout = QVBoxLayout()
        outline_widget.setLayout(outline_layout)

        outline_label = QLabel('Outline Tab')
        outline_label.setFont(self.gui.bold_font)
        outline_layout.addWidget(outline_label)

        outline_text = QTextEdit()
        outline_text.setHtml(
            'In the outline tab, you can record outlines for the sermon text as well as for your sermon. You can also '
            'save any illustration ideas that come up as you study'
        )
        outline_text.setFont(self.gui.standard_font)
        outline_text.setReadOnly(True)
        outline_layout.addWidget(outline_text)

        self.addTab(outline_widget, 'Outline Tab')

    def make_research(self):
        """
        Method to create the research widget for the help tabs.
        """
        research_widget = QWidget()
        research_layout = QVBoxLayout()
        research_widget.setLayout(research_layout)

        research_label = QLabel('Research Tab')
        research_label.setFont(self.gui.bold_font)
        research_layout.addWidget(research_label)

        research_text = QTextEdit()
        research_text.setHtml('In the research tab, you can jot down notes as you do research on the text for your '
                               'sermon.<br><br>You are able to use basic formatting as you record your notes: bold, '
                               'underline, italic, and bullet points.')
        research_text.setFont(self.gui.standard_font)
        research_text.setReadOnly(True)
        research_layout.addWidget(research_text)

        self.addTab(research_widget, 'Research Tab')

    def make_sermon(self):
        """
        Method to create the sermon widget for the help tabs.
        """
        sermon_widget = QWidget()
        sermon_layout = QVBoxLayout()
        sermon_widget.setLayout(sermon_layout)
        sermon_label = QLabel('Sermon Tab')
        sermon_label.setFont(self.gui.bold_font)
        sermon_layout.addWidget(sermon_label)

        sermon_text = QTextEdit()
        sermon_text.setHtml(
            'The sermon tab is for recording important details about your sermon, as well as the manuscript of the '
            'sermon itself.<br><br>At the top, you can enter such information as the date of the sermon, the location '
            'where the sermon will be given, as well as the title of your sermon. You can also record what call to '
            'worship you\'ll be using that day as well as a hymn of response, if either of these apply.<br><br> Just '
            'like in the research tab, you are able to format the text of your sermon manuscript with bold, italic, '
            'and underline fonts as you need to.'
        )
        sermon_text.setFont(self.gui.standard_font)
        sermon_text.setReadOnly(True)
        sermon_layout.addWidget(sermon_text)

        self.addTab(sermon_widget, 'Sermon Tab')
        
        
class FontFaceComboBox(QComboBox):
    """
    Creates a custom QComboBox that displays all fonts on the system in their own style.
    :param gui.GUI gui: The current instance of GUI
    """
    def __init__(self, gui):
        """
        :param gui.GUI gui: The current instance of GUI
        """
        super().__init__()
        self.gui = gui
        self.populate_widget()

    def populate_widget(self):
        try:
            row = 0
            model = self.model()
            families = QFontDatabase.families()
            for font in families:
                self.addItem(font)
                model.setData(model.index(row, 0), QFont(font, 14), Qt.ItemDataRole.FontRole)
                row += 1

        except Exception:
            self.gui.spd.error_log()


class PrintHandler(QWidget):
    def __init__(self, gui):
        super().__init__()
        self.gui = gui
        self.print_font = self.gui.standard_font
        self.print_font.setPointSize(self.gui.standard_font.pixelSize())
        self.line_height = 1.14
        self.html = self.get_all_data()
        self.printer = QPrinter()
        self.make_print_widget()
        self.show()

    def get_all_data(self):
        # get all text values from the GUI
        all_data = []
        for i in range(self.gui.scripture_frame_layout.count()):
            component = self.gui.scripture_frame_layout.itemAt(i).widget()

            if isinstance(component, QLineEdit):
                all_data.append(component.text())
            elif isinstance(component, QTextEdit):
                all_data.append(component.toSimplifiedHtml())

        for i in range(self.gui.exegesis_frame_layout.count()):
            component = self.gui.exegesis_frame_layout.itemAt(i).widget()

            if isinstance(component, QTextEdit) and not component.objectName() == 'textbox':
                all_data.append(component.toSimplifiedHtml())

        for i in range(self.gui.outline_frame_layout.count()):
            component = self.gui.outline_frame_layout.itemAt(i).widget()

            if isinstance(component, QTextEdit) and not component.objectName() == 'textbox':
                all_data.append(component.toSimplifiedHtml())

        for i in range(self.gui.research_frame_layout.count()):
            component = self.gui.research_frame_layout.itemAt(i).widget()

            if isinstance(component, QTextEdit) and not component.objectName() == 'textbox':
                all_data.append(component.toSimplifiedHtml())

        for i in range(self.gui.sermon_frame_layout.count()):
            component = self.gui.sermon_frame_layout.itemAt(i).widget()

            if isinstance(component, QLineEdit) or isinstance(component, QDateEdit):
                if isinstance(component, QLineEdit):
                    all_data.append(component.text())
                else:
                    all_data.append(component.date().toString('yyyy-MM-dd'))
            elif isinstance(component, QTextEdit) and not component.objectName() == 'textbox':
                all_data.append(component.toSimplifiedHtml())

        text_with_headers = []
        for i in range(len(all_data)):
            if '<p>' not in all_data[i]:
                all_data[i] = f'<p>{all_data[i]}</p>'

            has_contents = False
            this_string = re.sub('<.*?>', '', all_data[i]).strip()
            if len(this_string) > 0:
                has_contents = True

            if has_contents:
                text_with_headers.append(
                    f'<b><u>{self.gui.spd.user_settings["label" + str(i + 1)]}</u></b>'
                )
                text_with_headers.append(all_data[i])

        html = '\n'.join(text_with_headers)
        return html

    def make_document(self):
        printer_page_rect_inch = self.printer.pageRect(QPrinter.Unit.Inch)
        page_rect = QRectF(
            printer_page_rect_inch.x() * 96,
            printer_page_rect_inch.y() * 96,
            printer_page_rect_inch.width() * 96,
            printer_page_rect_inch.height() * 96
        )

        document = QTextDocument()
        text_option = QTextOption()
        text_option.setWrapMode(QTextOption.WrapMode.WordWrap)
        document.setDefaultTextOption(text_option)
        document.setDefaultFont(self.print_font)
        document.setPageSize(QSizeF(page_rect.width(), page_rect.height()))
        document.setDocumentMargin(page_rect.width() * (0.5 / 8.5))
        document.setDefaultStyleSheet(
            'p { '
            'font-family: "' + self.print_font.family() + '";'
            'font-size: ' + str(self.print_font.pointSize()) + 'pt;'
            'line-height: ' + str(self.line_height) + ';'
            '}'
        )
        document.setHtml(self.html)

        page_pixmaps = []
        current_y = 0
        painter = QPainter()
        painter.setFont(self.print_font)
        self.num_pages = document.pageCount()
        for i in range(document.pageCount()):
            pixmap = QPixmap(int(page_rect.width()), int(page_rect.height()))
            pixmap.fill(Qt.GlobalColor.white)

            painter.begin(pixmap)
            painter.translate(0, -current_y)
            document.drawContents(painter)
            painter.end()

            page_pixmaps.append(pixmap)
            current_y += int(page_rect.height())

        return document, page_pixmaps

    def make_print_widget(self):
        self.setWindowTitle('Print Record')
        self.setWindowFlag(Qt.WindowType.Window)

        print_layout = QHBoxLayout(self)

        preview_widget = QWidget()
        print_layout.addWidget(preview_widget)
        preview_layout = QVBoxLayout(preview_widget)

        self.preview_label = QLabel()
        preview_layout.addWidget(self.preview_label)

        page_widget = QWidget()
        preview_layout.addWidget(page_widget)
        page_layout = QHBoxLayout(page_widget)

        previous_button = QPushButton('<')
        previous_button.setObjectName('previous')
        previous_button.setFont(self.gui.bold_font)
        previous_button.pressed.connect(self.change_page)
        page_layout.addStretch()
        page_layout.addWidget(previous_button)

        self.page_label = QLabel()
        self.page_label.setFont(self.gui.bold_font)
        page_layout.addSpacing(20)
        page_layout.addWidget(self.page_label)
        page_layout.addSpacing(20)

        next_button = QPushButton('>')
        next_button.setObjectName('next')
        next_button.setFont(self.gui.bold_font)
        next_button.pressed.connect(self.change_page)
        page_layout.addWidget(next_button)
        page_layout.addStretch()

        options_box = QWidget()
        options_layout = QVBoxLayout(options_box)
        print_layout.addWidget(options_box)

        print_to_label = QLabel('Print to:')
        print_to_label.setFont(self.gui.bold_font)
        options_layout.addWidget(print_to_label)

        win_management = wmi.WMI()
        printers = win_management.Win32_Printer()
        printer_combobox = QComboBox()
        printer_combobox.setFont(self.gui.standard_font)
        default = ''
        for printer in printers:
            if not printer.Hidden:
                printer_combobox.addItem(printer.Name)
            if printer.Default:
                default = printer.Name
        printer_combobox.setCurrentText(default)
        printer_combobox.currentIndexChanged.connect(lambda: self.printer_change(printer_combobox.currentText()))
        options_layout.addWidget(printer_combobox)

        self.printer_change(default)

        text_options_widget = QWidget()
        options_layout.addWidget(text_options_widget)
        text_options_layout = QHBoxLayout(text_options_widget)

        font_size_label = QLabel('Font Size:')
        font_size_label.setFont(self.gui.bold_font)
        text_options_layout.addWidget(font_size_label)

        font_size_spinbox = QSpinBox()
        font_size_spinbox.setObjectName('font_size')
        font_size_spinbox.setSingleStep(2)
        font_size_spinbox.setFont(self.gui.standard_font)
        font_size_spinbox.setValue(self.print_font.pointSize())
        font_size_spinbox.valueChanged.connect(self.change_text_options)
        text_options_layout.addWidget(font_size_spinbox)

        point_label = QLabel('pt')
        point_label.setFont(self.gui.standard_font)
        text_options_layout.addWidget(point_label)
        text_options_layout.addSpacing(20)

        line_spacing_label = QLabel('Line spacing:')
        line_spacing_label.setFont(self.gui.bold_font)
        text_options_layout.addWidget(line_spacing_label)

        line_spacing_combobox = QComboBox()
        line_spacing_combobox.setObjectName('line_spacing')
        line_spacing_options = [
            'Single',
            '1.14',
            '1.5',
            'Double'
        ]
        line_spacing_combobox.addItems(line_spacing_options)
        line_spacing_combobox.setCurrentIndex(1)
        line_spacing_combobox.setFont(self.gui.standard_font)
        line_spacing_combobox.currentIndexChanged.connect(self.change_text_options)
        text_options_layout.addWidget(line_spacing_combobox)

        button_box = QWidget()
        button_layout = QHBoxLayout(button_box)
        options_layout.addWidget(button_box)

        ok_button = QPushButton('Ok')
        ok_button.pressed.connect(self.do_print)
        ok_button.setFont(self.gui.standard_font)
        button_layout.addWidget(ok_button)

        cancel_button = QPushButton('Cancel')
        cancel_button.pressed.connect(self.deleteLater)
        cancel_button.setFont(self.gui.standard_font)
        button_layout.addWidget(cancel_button)

        options_layout.addStretch()

    def printer_change(self, printer_name):
        self.printer = QPrinter(QPrinter.PrinterMode.HighResolution)
        self.printer.setPrinterName(printer_name)

        self.document, self.page_pixmaps = self.make_document()
        self.preview_label.setPixmap(
            self.page_pixmaps[0].scaledToWidth(400, Qt.TransformationMode.SmoothTransformation))
        self.current_page = 0
        self.page_label.setText(f'Page {self.current_page + 1} of {len(self.page_pixmaps)}')

    def change_text_options(self):
        if self.sender().objectName() == 'font_size':
            self.print_font.setPointSize(self.sender().value())
        else:
            if self.sender().currentText() == 'Single':
                self.line_height = 1
            elif self.sender().currentText() == 'Double':
                self.line_height = 2
            elif self.sender().currentText() == '1.14':
                self.line_height = 1.14
            elif self.sender().currentText() == '1.5':
                self.line_height = 1.5
        self.document, self.page_pixmaps = self.make_document()
        self.preview_label.setPixmap(
            self.page_pixmaps[0].scaledToWidth(400, Qt.TransformationMode.SmoothTransformation))
        self.current_page = 0
        self.page_label.setText(f'Page {self.current_page + 1} of {len(self.page_pixmaps)}')

    def change_page(self):
        if self.sender().objectName() == 'previous':
            self.current_page -= 1
            if self.current_page < 0:
                self.current_page = 0
        elif self.sender().objectName() == 'next':
            self.current_page += 1
            if self.current_page == len(self.page_pixmaps):
                self.current_page = len(self.page_pixmaps) - 1

        self.preview_label.setPixmap(
            self.page_pixmaps[self.current_page].scaledToWidth(400, Qt.TransformationMode.SmoothTransformation))
        self.page_label.setText(f'Page {self.current_page + 1} of {len(self.page_pixmaps)}')

    def do_print(self):
        self.document.print(self.printer)
        self.deleteLater()
