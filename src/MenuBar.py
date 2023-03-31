"""
@author Jeremy G. Wilson

Copyright 2023 Jeremy G. Wilson

This file is a part of the Sermon Prep Database program (v.3.4.0)

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
import os
import re
import shutil
import sys

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QStandardItemModel, QColor, QFontDatabase, QStandardItem, QPixmap
from PyQt5.QtWidgets import QFileDialog, QWidget, QVBoxLayout, QLabel, QTableView, QPushButton, QColorDialog, \
    QTabWidget, QHBoxLayout, QComboBox, QTextBrowser, QDialog, QLineEdit, QTextEdit, QDateEdit, QMessageBox
from pynput.keyboard import Key, Controller
from TopFrame import TopFrame

class MenuBar:
    def __init__(self, win, gui, spd):
        self.keyboard = Controller()
        self.win = win
        self.gui = gui
        self.spd = spd
        menu_bar = self.win.menuBar()

        menu_style = '''
            QMenu { background: white; } 
            QMenu:separator:hr { background-color: white; height: 0px; border-top: 1px solid black; margin: 5px } 
            QMenu::item:unselected { background-color: white; } 
            QMenu::item:selected { background: ''' + self.gui.background_color + '''; color: black; }'''

        file_menu = menu_bar.addMenu('File')
        file_menu.setStyleSheet(menu_style)
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
        edit_menu.setStyleSheet(menu_style)
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

        bg_color_action = config_menu.addAction('Change Accent Color')
        bg_color_action.setToolTip('Choose a different color for accents and borders')
        bg_color_action.triggered.connect(lambda: self.color_change('bg'))

        fg_color_action = config_menu.addAction('Change Background Color')
        fg_color_action.setToolTip('Choose a different color for the background')
        fg_color_action.triggered.connect(lambda: self.color_change('fg'))

        font_action = config_menu.addAction('Change Font')
        font_action.setToolTip('Change the font and font size used in the program')
        font_action.triggered.connect(self.change_font)

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
        record_menu.setStyleSheet(menu_style)
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
        help_menu.setStyleSheet(menu_style)

        help_action = help_menu.addAction('Help Topics')
        help_action.triggered.connect(self.show_help)

        about_action = help_menu.addAction('About')
        about_action.triggered.connect(self.show_about)

    def print_rec(self):
        from reportlab.lib.pagesizes import letter
        from reportlab.lib.styles import ParagraphStyle
        from reportlab.platypus import BaseDocTemplate, Frame, PageTemplate, Paragraph
        import os

        #get text from gui
        text = []
        for i in range(self.gui.scripture_frame_layout.count()):
            component = self.gui.scripture_frame_layout.itemAt(i).widget()

            if isinstance(component, QLineEdit):
                text.append(component.text())
            elif isinstance(component, QTextEdit):
                text.append(self.format_paragraph(component.toHtml()))

        for i in range(self.gui.exegesis_frame_layout.count()):
            component = self.gui.exegesis_frame_layout.itemAt(i).widget()

            if isinstance(component, QTextEdit) and not component.objectName() == 'textbox':
                text.append(self.format_paragraph(component.toHtml()))

        for i in range(self.gui.outline_frame_layout.count()):
            component = self.gui.outline_frame_layout.itemAt(i).widget()

            if isinstance(component, QTextEdit) and not component.objectName() == 'textbox':
                text.append(self.format_paragraph(component.toHtml()))

        for i in range(self.gui.research_frame_layout.count()):
            component = self.gui.research_frame_layout.itemAt(i).widget()

            if isinstance(component, QTextEdit) and not component.objectName() == 'textbox':
                text.append(self.format_paragraph(component.toHtml()))

        for i in range(self.gui.sermon_frame_layout.count()):
            component = self.gui.sermon_frame_layout.itemAt(i).widget()

            if isinstance(component, QLineEdit) or isinstance(component, QDateEdit):
                if isinstance(component, QLineEdit):
                    text.append(component.text())
                else:
                    text.append(component.date().toString('yyyy-MM-dd'))
            elif isinstance(component, QTextEdit) and not component.objectName() == 'textbox':
                text.append(self.format_paragraph(component.toHtml()))

        print_file_loc = self.spd.app_dir + '/print.pdf'

        normal = ParagraphStyle(
            name='Normal',
            fontName='Helvetica',
            fontSize=11,
            firstLineIndent=18,
            leading=16,
            spaceBefore=3,
            spaceAfter=6
        )

        bullet = ParagraphStyle(
            name='Normal',
            fontName='Helvetica',
            fontSize=11,
            leading=16,
            bulletIndent=20,
            leftIndent=30
        )

        heading = ParagraphStyle(
            name='Heading',
            fontName='Helvetica-Bold',
            fontSize=11,
            spaceBefore=6,
            spaceAfter=3
        )

        doc = BaseDocTemplate(print_file_loc, pagesize=letter)
        frame = Frame(
            doc.leftMargin,
            doc.bottomMargin,
            doc.width,
            doc.height,
            0, 0, 0, 0,
            id='normal')
        template = PageTemplate(id='test', frames=frame)
        doc.addPageTemplates([template])

        doc_text = []
        for i in range(0, len(text)):
            paragraph_text = text[i]
            if isinstance(paragraph_text, str)\
                    and any(c.isalpha() for c in paragraph_text)\
                    or 'date' in self.spd.user_settings[i + 5].lower():
                t = re.sub('<.*?>', '', paragraph_text)
                if any(c.isalpha for c in t):
                    doc_text.append(Paragraph(self.spd.user_settings[i + 5], heading))
                    if '<bullet>' in paragraph_text:
                        doc_text.append(Paragraph(paragraph_text, bullet))
                    else:
                        doc_text.append(Paragraph(paragraph_text, normal))
            else:
                has_contents = False
                for paragraph in paragraph_text:
                    t = re.sub('<.*?>', '', paragraph)
                    if any(c.isalpha() for c in t):
                        has_contents = True

                if has_contents:
                    doc_text.append(Paragraph(self.spd.user_settings[i + 5], heading))

                    for paragraph in paragraph_text:
                        if any(c.isalpha() for c in paragraph):
                            if '<bullet>' in paragraph:
                                doc_text.append(Paragraph(paragraph, bullet))
                            else:
                                doc_text.append(Paragraph(paragraph, normal))
        doc.build(doc_text)

        from subprocess import Popen, PIPE
        self.spd.write_to_log('Opening print subprocess')

        if sys.platform == 'win32':
            print('beginning win32 print process')
            CREATE_NO_WINDOW = 0x08000000
            p = Popen(
                [
                    self.spd.cwd + 'ghostscript/gsprint.exe',
                    print_file_loc,
                    '-ghostscript',
                    self.spd.cwd + 'ghostscript/gswin64.exe',
                    '-query'],
                creationflags=CREATE_NO_WINDOW,
                stdout=PIPE,
                stderr=PIPE)

            self.spd.write_to_log('Capturing print subprocess sdtout & stderr')
            stdout, stderr = p.communicate()
            self.spd.write_to_log('stdout:' + str(stdout))
            self.spd.write_to_log('stderr:' + str(stderr))

            os.remove(print_file_loc)

        elif sys.platform == 'linux':
            p = Popen(['/usr/bin/lpstat', '-a'], stdin = PIPE, stdout = PIPE, stderr = PIPE)
            stdout, stderr = p.communicate()
            stdout = str(stdout).replace('b\'', '')
            stdout = stdout.replace('\'', '')
            lines = stdout.split('\\n')

            printers = []
            for line in lines:
                printers.append(line.split(' ')[0].replace('_', ' '))

            print_dialog = QDialog(self.gui.main_widget)
            layout = QVBoxLayout()
            print_dialog.setLayout(layout)

            label = QLabel('Choose Printer:')
            layout.addWidget(label)

            printer_combobox = QComboBox()
            for printer in printers:
                printer_combobox.addItem(printer)
            layout.addWidget(printer_combobox)

            button_widget = QWidget()
            button_layout = QHBoxLayout()
            button_widget.setLayout(button_layout)

            ok_button = QPushButton('OK')
            ok_button.pressed.connect(lambda: self.linux_print(print_dialog, printer_combobox.currentText(), print_file_loc))
            button_layout.addWidget(ok_button)

            cancel_button = QPushButton('Cancel')
            cancel_button.pressed.connect(print_dialog.destroy)
            button_layout.addWidget(cancel_button)

            layout.addWidget(button_widget)

            print_dialog.show()

    def format_paragraph(self, comp_text):
        comp_text = re.sub('<p.*?>', '', comp_text, flags=re.DOTALL)
        comp_text = re.sub('</p>', '<br><br>', comp_text)

        slice = re.findall('<span.*?span>', comp_text)
        for item in slice:
            addition = ['', '']
            if 'font-weight' in item:
                addition[0] = '<strong>' + addition[0]
                addition[1] = addition[1] + '</strong>'
            if 'text-decoration' in item:
                addition[0] = '<u>' + addition[0]
                addition[1] = addition[1] + '</u>'
            if 'font-style' in item:
                addition[0] = '<i>' + addition[0]
                addition[1] = addition[1] + '</i>'

            new_string = re.sub('<span.*?>', addition[0], item)
            new_string = re.sub('</span>', addition[1], new_string)
            comp_text = comp_text.replace(item, new_string)

        slice = re.findall('<li.*?</li>', comp_text)
        for item in slice:
            new_string = re.sub('<li.*?>', '<bullet>&bull;</bullet>', item)
            new_string = re.sub('</li>', '<br><br>', new_string)
            comp_text = comp_text.replace(item, new_string)

        comp_text = re.sub('<span.*?>', '', comp_text)
        comp_text = re.sub('</span>', '', comp_text)
        comp_text = re.sub('<style?(.*?)style>', '', comp_text, flags=re.DOTALL)

        text_split = comp_text.split('<br><br>')

        return text_split

    def linux_print(self, print_dialog, item_text, printFileLoc):
        print_dialog.destroy()
        item_text = item_text.replace(' ', '_')

        from subprocess import Popen, PIPE
        import os

        try:
            p = Popen(
                [
                    'lpr',
                    '-P',
                    item_text,
                    printFileLoc
                ],
                stdin = PIPE,
                stdout = PIPE,
                stderr = PIPE
            )

            self.spd.write_to_log('Capturing print subprocess sdtout & stderr')
            stdout, stderr = p.communicate()
            self.spd.write_to_log('stdout:' + str(stdout))
            self.spd.write_to_log('stderr:' + str(stderr))

            os.remove(printFileLoc)
        except Exception as err:
            self.spd.write_to_log('MenuBar.linux_print: ' + str(err))

    def do_backup(self):
        import os
        user_dir = os.path.expanduser('~')

        fileName = QFileDialog.getSaveFileName(self.win, 'Create Backup',
                                               user_dir + '/sermon_prep_database_backup.db', 'Database File (*.db)');
        import shutil
        shutil.copy(self.spd.db_loc, fileName[0])
        self.spd.write_to_log('Created Backup as ' + fileName[0])

        QMessageBox.information(
            None,
            'Backup Created',
            'Backup successfully created as ' + fileName[0],
            QMessageBox.Ok
        )

    def restore_backup(self):
        dialog = QFileDialog()
        dialog.setWindowTitle('Restore from Backup')
        dialog.setNameFilter('Database File (*.db)')
        dialog.setFileMode(QFileDialog.ExistingFile)
        dialog.setDirectory(self.spd.app_dir)

        dialog.exec_()

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
                QMessageBox.Ok
            )

            try:
                self.spd.get_ids()
                self.spd.get_date_list()
                self.spd.get_scripture_list()
                self.spd.get_user_settings()
                self.gui.top_frame.dates_cb.addItems(self.spd.dates)
                for item in self.spd.references:
                    self.gui.top_frame.references_cb.addItem(item[0])
                self.spd.last_rec()
            except Exception as err:
                shutil.copy(self.spd.app_dir + '/active-database-backup.db', self.spd.db_loc)
                self.spd.write_to_log('MenuBar.restore_backup: ' + str(err))

                QMessageBox.critical(
                    None,
                    'Error Loading Database Data',
                    'There was a problem loading the data from your backup. Your prior database has not been changed.  '
                    'The error is as follows:\r\n' + str(err),
                    QMessageBox.Ok
                )

                self.spd.get_ids()
                self.spd.get_date_list()
                self.spd.get_scripture_list()
                self.spd.get_user_settings()
                self.gui.top_frame.dates_cb.addItems(self.spd.dates)
                for item in self.spd.references:
                    self.gui.references_cb.addItem(item[0])
                self.spd.last_rec()
            else:
                QMessageBox.information(
                    None,
                    'Backup Restored',
                    'Backup successfully restored.\n\nA copy of your prior database has been saved as ' + self.spd.app_dir
                    + '/active-database-backup.db',
                    QMessageBox.Ok
                )

    def import_from_files(self):
        QMessageBox.information(
            self.gui.win,
            'Import from Files',
            'For best results, the files you are importing should be named according to this syntax:\n\n'
            'YYYY-MM-DD.book.chapter.verse-verse\n\n'
            'For example, a sermon preached on May 20th, 2011 on Mark 3:1-12, saved as a Microsoft Word document,'
            'would be named:\n\n'
            '2011-05-11.mark.3.1-12.docx', QMessageBox.Ok)
        from GetFromDocx import GetFromDocx
        GetFromDocx(self.gui)

    def import_bible(self):
        file = QFileDialog.getOpenFileName(
            self.gui.win,
            'Choose Bible File',
            os.path.expanduser('~'),
            'XML Bible File (*.xml)'
        )
        try:
            if file[0]:
                shutil.copy(file[0], self.spd.app_dir + '/my_bible.xml')
                self.spd.bible_file = self.spd.app_dir + '/my_bible.xml'
        except Exception as ex:
            self.spd.write_to_log(str(ex))

    def rename_labels(self):
        self.rename_widget = QWidget()
        self.rename_widget.setWindowFlag(Qt.Window)
        self.rename_widget.resize(920, 400)
        rename_widget_layout = QVBoxLayout()
        self.rename_widget.setLayout(rename_widget_layout)
        self.rename_widget.setStyleSheet('''
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

        rename_label = QLabel('Use this table to set new names for any of the labels in this program.\n'
                             'Double-click a label under "New Label" to rename it')
        rename_widget_layout.addWidget(rename_label)

        model = QStandardItemModel(len(self.spd.user_settings), 2)
        for i in range(len(self.spd.user_settings)):
            if i > 4:
                item = QStandardItem(self.spd.user_settings[i])
                item.setEditable(False)
                model.setItem(i - 5, 0, item)
                item2 = QStandardItem(self.spd.user_settings[i])
                model.setItem(i - 5, 1, item2)
        model.setHeaderData(0, Qt.Horizontal, 'Current Label')
        model.setHeaderData(1, Qt.Horizontal, 'New Label')

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
        new_labels = []
        for i in range(model.rowCount()):
            new_labels.append(model.data(model.index(i, 1)))
        self.rename_widget.destroy()
        self.spd.write_label_changes(new_labels)
        for widget in self.gui.main_widget.children():
            if isinstance(widget, QWidget):
                self.gui.layout.removeWidget(widget)

        self.gui.menu_bar = self
        self.gui.top_frame = TopFrame(self.win, self.gui, self.spd)
        self.gui.layout.addWidget(self.gui.top_frame)
        self.gui.build_tabbed_frame()
        self.gui.build_scripture_tab()
        self.gui.build_exegesis_tab()
        self.gui.build_outline_tab()
        self.gui.build_research_tab()
        self.gui.build_sermon_tab()
        self.gui.set_style_sheets()

        self.spd.get_by_index(self.spd.current_rec_index)

    def color_change(self, type):
        if type == 'bg':
            color_chooser = QColorDialog()
            new_color = color_chooser.getColor(QColor(self.gui.accent_color))
            if not new_color == QColor():
                self.gui.accent_color = new_color.name()
        else:
            color_chooser = QColorDialog()
            new_color = color_chooser.getColor(QColor(self.gui.background_color))
            if not new_color == QColor():
                self.gui.background_color = new_color.name()
        self.gui.set_style_sheets()
        self.spd.write_color_changes()

    def disable_spell_check(self):
        if self.disable_spell_check_action.isChecked():
            self.spd.disable_spell_check = True
            self.spd.write_spell_check_changes()
        else:
            self.spd.disable_spell_check = False
            if not self.spd.sym_spell:
                from Dialogs import timed_popup
                timed_popup('Please wait while the dictionary is loaded...', 5000, self.gui.accent_color)
                self.spd.app.processEvents()
                self.spd.load_dictionary()
            self.spd.write_spell_check_changes()

    def show_help(self):
        global sh
        sh = ShowHelp(self.gui.background_color, self.gui.accent_color, self.gui.font_family, self.gui.font_size, self.spd)
        sh.show()

    def show_about(self):
        global about_win
        about_win = QWidget()
        about_win.setStyleSheet('background-color: ' + self.gui.background_color + ';')
        about_win.resize(600, 400)
        about_layout = QVBoxLayout()
        about_win.setLayout(about_layout)

        about_label = QLabel('Sermon Prep Database v.3.4.0')
        about_label.setStyleSheet('font-family: "Helvetica"; font-weight: bold; font-size: 16px;')
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
        about_text.setStyleSheet('font-family: "Helvetica"; font-size: 16px')
        about_text.setOpenExternalLinks(True)
        about_text.setReadOnly(True)
        about_layout.addWidget(about_text)

        about_win.show()

    def remove_words(self):
        from Dialogs import RemoveCustomWords
        self.rcm = RemoveCustomWords(self.spd)

    def change_font(self):
        fonts = QFontDatabase()
        current_font = self.gui.font_family
        current_size = self.gui.font_size

        global font_chooser
        font_chooser = QWidget()
        font_layout = QVBoxLayout()
        font_chooser.setLayout(font_layout)
        font_chooser.setStyleSheet('background: ' + self.gui.background_color + ';')

        top_panel = QWidget()
        top_layout = QHBoxLayout()
        top_panel.setLayout(top_layout)
        font_layout.addWidget(top_panel)

        family_combo_box = QComboBox()
        family_combo_box.addItems(fonts.families())
        family_combo_box.currentIndexChanged.connect(
            lambda: self.apply_font(font_chooser, family_combo_box.currentText(), self.gui.font_size, False))
        top_layout.addWidget(family_combo_box)

        size_combo_box = QComboBox()
        sizes = ['10', '12', '14', '16', '18', '20', '24', '28', '32']
        size_combo_box.addItems(sizes)
        size_combo_box.currentIndexChanged.connect(
            lambda: self.apply_font(font_chooser, self.gui.font_family, size_combo_box.currentText(), False))
        top_layout.addWidget(size_combo_box)

        family_combo_box.setCurrentText(self.gui.font_family)
        size_combo_box.setCurrentText(self.gui.font_size)

        bottom_panel = QWidget()
        bottom_layout = QHBoxLayout()
        bottom_panel.setLayout(bottom_layout)
        font_layout.addWidget(bottom_panel)

        ok_button = QPushButton('OK')
        ok_button.clicked.connect(
            lambda: self.apply_font(font_chooser, family_combo_box.currentText(), size_combo_box.currentText(), True))
        bottom_layout.addWidget(ok_button)

        cancel_button = QPushButton('Cancel')
        cancel_button.clicked.connect(lambda: self.apply_font(font_chooser, current_font, current_size, True))
        bottom_layout.addWidget(cancel_button)

        font_chooser.show()

    def apply_font(self, fontChooser, family, size, close):
        self.gui.font_family = family
        self.gui.font_size = size
        self.gui.set_style_sheets()
        if close:
            self.spd.write_font_changes(family, size)
            fontChooser.destroy()

    def press_ctrl_z(self):
        self.keyboard.press(Key.ctrl)
        self.keyboard.press('z')
        self.keyboard.release('z')
        self.keyboard.release(Key.ctrl)

    def press_ctrl_y(self):
        self.keyboard.press(Key.ctrl)
        self.keyboard.press('y')
        self.keyboard.release('y')
        self.keyboard.release(Key.ctrl)

    def press_ctrl_x(self):
        self.keyboard.press(Key.ctrl)
        self.keyboard.press('x')
        self.keyboard.release('x')
        self.keyboard.release(Key.ctrl)

    def press_ctrl_c(self):
        self.keyboard.press(Key.ctrl)
        self.keyboard.press('c')
        self.keyboard.release('c')
        self.keyboard.release(Key.ctrl)

    def press_ctrl_v(self):
        self.keyboard.press(Key.ctrl)
        self.keyboard.press('v')
        self.keyboard.release('v')
        self.keyboard.release(Key.ctrl)

    def do_exit(self):
        goon = True
        if self.gui.changes:
            goon = self.spd.ask_save()
        if goon:
            sys.exit(0)

class ShowHelp(QTabWidget):
    def __init__(self, background_color, accent_color, font_family, font_size, spd):
        super().__init__()
        self.background_color = background_color
        self.accent_color = accent_color
        self.font_family = font_family
        self.font_size = font_size
        self.spd = spd
        self.bold_font = 'font-family: "' + self.font_family + '"; font-weight: bold; font-size: ' + self.font_size + 'pt;'
        self.plain_font = 'font-family: "' + self.font_family + '"; font-size: ' + self.font_size + 'pt;'

        self.setStyleSheet('''
                    QTabBar::tab {
                        background-color: ''' + self.accent_color + ''';
                        color: white;
                        font-family: "''' + self.font_family + '''";
                        font-size: ''' + self.font_size + '''pt;
                        width: 140px;
                        height: 30px;
                        padding: 10px;}
                    QTabBar::tab:selected {
                        background-color: ''' + self.background_color + ''';
                        color: black;
                        font-family: "''' + self.font_family + '''";
                        font-size: ''' + str(self.font_size) + '''pt;
                        width: 140px;
                        height: 30px;
                        font-weight: bold;}
                    ''')

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
        intro_widget = QWidget()
        intro_widget.setStyleSheet('background-color: ' + self.background_color)
        intro_layout = QVBoxLayout()
        intro_widget.setLayout(intro_layout)

        intro_label = QLabel('Introduction')
        intro_label.setStyleSheet(self.bold_font)
        intro_layout.addWidget(intro_label)

        intro_text = QTextEdit(
            'Thank-you for trying out this Sermon Prep Database. Its purpose is to provide an easy-to-use program to '
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
            'assimilate your unique biological and technological distinctiveness, email pastorjeremywilson@gmail.com'
        )
        intro_text.setStyleSheet('background-color: ' + self.background_color + '; ' + self.plain_font)
        intro_text.setMinimumWidth(750)
        intro_text.setReadOnly(True)
        intro_layout.addWidget(intro_text)

        self.addTab(intro_widget, 'Introduction')

    def make_menu(self):
        menu_widget = QWidget()
        menu_widget.setStyleSheet('background-color: ' + self.background_color)
        menu_layout = QVBoxLayout()
        menu_widget.setLayout(menu_layout)

        menu_label = QLabel('Menu Bar')
        menu_label.setStyleSheet(self.bold_font)
        menu_layout.addWidget(menu_label)
        menu_pic = QPixmap(self.spd.cwd + 'resources/menuPic.png')
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
            '2023 on Mark 9:1-12, saved as a Microsoft Word document would have the file name <strong>2023-03-19.mark.'
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
            'program as well as what the labels say and the font that is used.<br><br>First, you\'ll see commands that '
            'change the accent and background colors of the program. Changing the accent color will affect the '
            'background color of the left-side tabs, while changing the background colors will change the color '
            'surrounding the different boxes where you enter information.<br><br>Next, the Edit menu contains a '
            'command to <strong>Rename Labels</strong>. Choosing this will open up a new window where you can '
            'customize the labels that appear above each data entry box. This new window will have two columns, the '
            'first showing the current label, and the second being where you can type in what you would rather have '
            'that label say. For example, if you like to use, "Big Idea of the Text," instead of, "Central Proposition '
            'of the Text," you would click inside the right-hand "Central Proposition of the Text" box and change it '
            'to read, "Big Idea of the Text."<br><br>Finally, there is a command to <strong>Change Font</strong>. '
            'This allows you to choose a custom font and font size for the labels and entry boxes throughout the '
            'program. You may have a particular font or size that you find easier to read, and you can change to that '
            'here.<br><br><strong>Record</strong><br>In the Record menu are the same navigation controls that appear '
            'in the upper bar of the window. In addition is a <strong>Delete Record</strong> option. Use this if you '
            'are currently viewing a record you would like to delete; if it\'s a duplicate or unneeded, for example. '
            'Deleting a record cannot be undone, so you will be prompted to make sure you really want to delete the '
            'current record.'
        )
        menu_text.setStyleSheet('background-color: ' + self.background_color + '; ' + self.plain_font)
        menu_text.setReadOnly(True)
        menu_text.setOpenExternalLinks(True)
        menu_layout.addWidget(menu_text)

        self.addTab(menu_widget, 'Menu Bar')

    def make_tool(self):
        tool_widget = QWidget()
        tool_layout = QVBoxLayout()
        tool_widget.setLayout(tool_layout)

        tool_label = QLabel('Toolbar')
        tool_label.setStyleSheet(self.bold_font)
        tool_layout.addWidget(tool_label)

        tool_text = QTextEdit()
        tool_text.setHtml(
            'Along the top of the window, you\'ll see a toolbar containing different formatting and navigation tools.'
            '<br><br><img src="' + self.spd.cwd + 'resources/undoPic.png"><br>'
            'First, you\'ll see Undo and Redo buttons that you can use to undo or redo any editing you perform. Of '
            'course, Ctrl-Z and Ctrl-Y will also perform these functions.<br><br><img src="' + self.spd.cwd +
            'resources/formatPic.png"><br>Beside those are the buttons you can use to format your text with bold, '
            'italic, underline, and bullet-point options.<br><br><img src="' + self.spd.cwd +
            'resources/showScripturePic.png"><br>Right beside these is an option to show the sermon text on all tabs. '
            'If you would like to keep your sermon text handy while you are entering, for example, your exegesis notes '
            'or research notes, click this icon. The sermon text will be shown to the right of the Exegesis, Outline, '
            'Research, and Sermon tabs.<br><br><img src="' + self.spd.cwd + 'resources/searchPic.png"><br>In the '
            'middle of this bar are two pull-down menus where you can bring up past sermons by the sermon date '
            'or the sermon\'s scripture. This becomes increasingly useful as you add more and more sermons to the '
            'database, allowing you to see what you\'ve preached on in the past or how you\'ve dealt with particular '
            'texts.<br><br>Next to these is a search box. By entering a passage or keyword into this box and pressing '
            'Enter, you can search for scripture passages or words you\'ve used in past sermons. Doing so will create '
            'a new tab showing any records where your search term(s) appear, and double-clicking any of the results '
            'will bring up that particular sermon\'s record. This new tab can be closed by pressing the "X" icon.<br>'
            '<br><img src="' + self.spd.cwd + 'resources/navPic.png"><br>To the right of this upper bar are the navigation '
            'buttons, as well as the save and print buttons. These '
            'buttons will allow you to navigate to the first or last sermons in your database or switch to the '
            'previous or next sermons. Just to the right of these navigation buttons is the "New Record" button. '
            'This will create a new, blank record to start recording information for your next sermon.<br><br>After '
            'this is the save and print buttons. The save button will save all your current info to the database, '
            'while the print button will print a basic hard-copy of the information you have for the currently '
            'showing sermon.'
        )
        tool_text.setStyleSheet('background-color: ' + self.background_color + '; ' + self.plain_font)
        tool_text.setReadOnly(True)
        tool_layout.addWidget(tool_text)

        self.addTab(tool_widget, 'Toolbar')

    def make_scripture(self):
        scripture_widget = QWidget()
        scripture_layout = QVBoxLayout()
        scripture_widget.setLayout(scripture_layout)
        scripture_label = QLabel('Scripture Tab')
        scripture_label.setStyleSheet(self.bold_font)
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
        scripture_text.setStyleSheet('background-color: ' + self.background_color + '; ' + self.plain_font)
        scripture_text.setReadOnly(True)
        scripture_layout.addWidget(scripture_text)

        self.addTab(scripture_widget, 'Scripture Tab')

    def make_exeg(self):
        exeg_widget = QWidget()
        exeg_layout = QVBoxLayout()
        exeg_widget.setLayout(exeg_layout)
        exeg_label = QLabel('Exegesis Tab')
        exeg_label.setStyleSheet(self.bold_font)
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
        exeg_text.setStyleSheet('background-color: ' + self.background_color + '; ' + self.plain_font)
        exeg_text.setReadOnly(True)
        exeg_layout.addWidget(exeg_text)

        self.addTab(exeg_widget, 'Exegesis Tab')

    def make_outline(self):
        outline_widget = QWidget()
        outline_layout = QVBoxLayout()
        outline_widget.setLayout(outline_layout)

        outline_label = QLabel('Outline Tab')
        outline_label.setStyleSheet(self.bold_font)
        outline_layout.addWidget(outline_label)

        outline_text = QTextEdit()
        outline_text.setHtml(
            'In the outline tab, you can record outlines for the sermon text as well as for your sermon. You can also '
            'save any illustration ideas that come up as you study'
        )
        outline_text.setStyleSheet('background-color: ' + self.background_color + '; ' + self.plain_font)
        outline_text.setReadOnly(True)
        outline_layout.addWidget(outline_text)

        self.addTab(outline_widget, 'Outline Tab')

    def make_research(self):
        research_widget = QWidget()
        research_layout = QVBoxLayout()
        research_widget.setLayout(research_layout)

        research_label = QLabel('Research Tab')
        research_label.setStyleSheet(self.bold_font)
        research_layout.addWidget(research_label)

        research_text = QTextEdit()
        research_text.setHtml('In the research tab, you can jot down notes as you do research on the text for your '
                               'sermon.<br><br>You are able to use basic formatting as you record your notes: bold, '
                               'underline, italic, and bullet points.')
        research_text.setStyleSheet('background-color: ' + self.background_color + '; ' + self.plain_font)
        research_text.setReadOnly(True)
        research_layout.addWidget(research_text)

        self.addTab(research_widget, 'Research Tab')

    def make_sermon(self):
        sermon_widget = QWidget()
        sermon_layout = QVBoxLayout()
        sermon_widget.setLayout(sermon_layout)
        sermon_label = QLabel('Sermon Tab')
        sermon_label.setStyleSheet(self.bold_font)
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
        sermon_text.setStyleSheet('background-color: ' + self.background_color + '; ' + self.plain_font)
        sermon_text.setReadOnly(True)
        sermon_layout.addWidget(sermon_text)

        self.addTab(sermon_widget, 'Sermon Tab')
