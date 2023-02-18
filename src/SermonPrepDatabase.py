'''
@author Jeremy G. Wilson

Copyright 2023 Jeremy G. Wilson

This file is a part of the Sermon Prep Database program (v.3.3.6)

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
'''

import logging
import os
import re
import sqlite3
import sys
import time
from datetime import datetime
from os.path import exists

from PyQt5.QtCore import QThread, Qt, QObject, pyqtSignal
from PyQt5.QtGui import QFont, QMovie, QColor
from PyQt5.QtWidgets import QApplication, QLineEdit, QTextEdit, QDateEdit, QWidget, QLabel, QGridLayout, QProgressBar, \
    QDialog
from symspellpy import SymSpell

from Dialogs import yes_no_cancel_box
from gui import GUI


# The main program class that handles startup functions such as checking for/creating a new database, instantiating
# the gui, and polling the database for data. Also handles any database reading and writing functions.
class SermonPrepDatabase(QThread):
    change_text = pyqtSignal(str)
    finished = pyqtSignal()

    gui = None
    ids = []
    dates = []
    references = []
    db_loc = None
    current_rec_index = 0
    user_settings = None
    app = None
    platform = ''
    cwd = ''
    sym_spell = None

    # On startup, initialize a QApplication, get the platform, set the app_dir and db_loc, instantiate the GUI
    def run(self):
        time.sleep(1.0)
        self.change_text.emit('Getting Platform')
        time.sleep(0.5)
        self.platform = sys.platform

        self.change_text.emit('Getting Directories')
        time.sleep(0.5)
        user_dir = os.path.expanduser('~')
        if self.platform == 'win32':
            self.app_dir = user_dir + '/AppData/Roaming/Sermon Prep Database'
        elif self.platform == 'linux':
            self.app_dir = user_dir + '/.sermonPrepDatabase'
            os.environ['QT_QPA_PLATFORM'] = 'offscreen'
        self.db_loc = self.app_dir + '/sermon_prep_database.db'

        if not exists(self.app_dir):
            os.mkdir(self.app_dir)

        self.write_to_log('platform is ' + self.platform)
        self.write_to_log('current working directory is ' + self.cwd)
        self.write_to_log('application directory is ' + self.app_dir)
        self.write_to_log('database location is ' + self.db_loc)

        self.change_text.emit('Loading Dictionaries')
        time.sleep(0.5)
        self.load_dictionary()

        self.change_text.emit('Creating GUI')
        time.sleep(0.5)

        self.finished.emit()

    def load_dictionary(self):
        self.sym_spell = SymSpell()
        self.sym_spell.create_dictionary(self.cwd + 'resources/dictionary_en_us_custom.txt')
        with open(self.cwd + 'resources/custom_words.txt', 'r') as file:
            custom_words = file.readlines()
        for entry in custom_words:
            self.sym_spell.create_dictionary_entry(entry.strip(), 1)

    def add_to_dictionary(self, widget, word):
        try:
            self.sym_spell.create_dictionary_entry(word, 1)
            with open(self.cwd + 'resources/custom_words.txt', 'a') as file:
                file.write(word + '\n')
            text = widget.toMarkdown()
            widget.setMarkdown(text)

        except Exception:
            logging.exception('')


    # retrieve the list of ID numbers from the database
    def get_ids(self):
        self.ids = []
        conn = sqlite3.connect(self.db_loc)
        cur = conn.cursor()
        results = cur.execute('SELECT ID FROM sermon_prep_database').fetchall()
        for item in results:
            self.ids.append(item[0])

    # retrieve the list of dates from the database
    def get_date_list(self):
        self.dates = []
        conn = sqlite3.connect(self.db_loc)
        cur = conn.cursor()
        results = cur.execute('SELECT date FROM sermon_prep_database').fetchall()
        for item in results:
            self.dates.append(item[0])

    # retrieve the list of scripture references from the database
    def get_scripture_list(self):
        self.references = []
        conn = sqlite3.connect(self.db_loc)
        cur = conn.cursor()
        self.references = cur.execute(
            'SELECT sermon_reference, ID FROM sermon_prep_database ORDER BY sermon_reference').fetchall()

    # retrieve the user's custom settings from the database
    def get_user_settings(self):
        conn = sqlite3.connect(self.db_loc)
        cur = conn.cursor()
        self.user_settings = cur.execute('SELECT * FROM user_settings').fetchall()[0]

    # copy the current database to a dated file as a backup
    def backup_db(self):
        from datetime import datetime
        now = datetime.now().strftime("%m-%d-%Y_%H-%M-%S")
        backup_file = self.app_dir + '/sermon_prep_database.backup-' + now + '.db'

        file_array = os.listdir(self.app_dir)
        backup_files = []
        for file in file_array:
            if 'backup-' in file:
                backup_files.append(file)

        if (len(backup_files) >= 5): # keep only 5 prior copies of the database
            os.remove(self.app_dir + '/' + backup_files[0])

        import shutil
        shutil.copy(self.db_loc, backup_file)
        self.write_to_log('New backup file created at ' + backup_file)

    # save user's color changes to the database
    def write_color_changes(self):
        sql = 'UPDATE user_settings SET bgcolor = "' + self.gui.accent_color + '", fgcolor = "' + self.gui.background_color + '" WHERE ID = 1;'
        conn = sqlite3.connect(self.db_loc)
        cur = conn.cursor()
        cur.execute(sql)
        conn.commit()
        self.get_user_settings()

    # save user's font changes to the database
    def write_font_changes(self, family, size):
        sql = 'UPDATE user_settings SET font_family = "' + family + '", font_size = "' + size + '" WHERE ID = 1;'
        conn = sqlite3.connect(self.db_loc)
        cur = conn.cursor()
        cur.execute(sql)
        conn.commit()
        self.get_user_settings()

    # save user's label changes to the database
    def write_label_changes(self, new_labels):
        sql = 'UPDATE user_settings SET '
        for i in range(1, 22):
            if i == 21:
                sql += 'l' + str(i) + ' = "' + new_labels[i - 1] + '" WHERE ID = 1;'
            else:
                sql += 'l' + str(i) + ' = "' + new_labels[i - 1] + '",'
        conn = sqlite3.connect(self.db_loc)
        cur = conn.cursor()
        cur.execute(sql)
        conn.commit()
        self.get_user_settings()

    def get_record_data(self):
        conn = sqlite3.connect(self.db_loc)
        cur = conn.cursor()
        results = cur.execute("SELECT * FROM sermon_prep_database WHERE ID = " + str(self.ids[self.current_rec_index]))
        record = results.fetchall()
        return record

    # retrieve the data for a specific record based on the index of the ids array
    def get_by_index(self, index):
        if len(self.ids) > 0:
            self.current_rec_index = index
            counter = 0
            id_to_find = self.ids[index]
            for item in self.references:
                if item[1] == id_to_find:
                    break
                counter += 1

            self.gui.top_frame.dates_cb.blockSignals(True)
            self.gui.top_frame.references_cb.blockSignals(True)
            self.gui.top_frame.dates_cb.setCurrentIndex(index)
            self.gui.top_frame.references_cb.setCurrentIndex(counter)
            self.gui.top_frame.dates_cb.blockSignals(False)
            self.gui.top_frame.references_cb.blockSignals(False)

            conn = sqlite3.connect(self.db_loc)
            cur = conn.cursor()
            results = cur.execute("SELECT * FROM sermon_prep_database WHERE ID = " + str(self.ids[index]))
            record = results.fetchall()
            self.gui.fill_values(record)

            if index == 0:
                self.gui.top_frame.first_rec_button.setEnabled(False)
                self.gui.top_frame.prev_rec_button.setEnabled(False)
                self.gui.top_frame.next_rec_button.setEnabled(True)
                self.gui.top_frame.last_rec_button.setEnabled(True)
            elif index == len(self.ids) - 1:
                self.gui.top_frame.first_rec_button.setEnabled(True)
                self.gui.top_frame.prev_rec_button.setEnabled(True)
                self.gui.top_frame.next_rec_button.setEnabled(False)
                self.gui.top_frame.last_rec_button.setEnabled(False)
            else:
                self.gui.top_frame.first_rec_button.setEnabled(True)
                self.gui.top_frame.prev_rec_button.setEnabled(True)
                self.gui.top_frame.next_rec_button.setEnabled(True)
                self.gui.top_frame.last_rec_button.setEnabled(True)
        else:
            self.new_rec()

    # save the user's data from all elements of the GUI
    def save_rec(self):
        conn = sqlite3.connect(self.db_loc)
        cur = conn.cursor()
        sql = "SELECT * FROM sermon_prep_database where 1=0;"
        cur.execute(sql)

        columns = [d[0] for d in cur.description]

        sql = 'UPDATE sermon_prep_database SET '
        rec_id = self.ids[self.current_rec_index]
        sql += '"' + columns[0] + '" = "' + str(rec_id) + '",'

        index = 1
        for i in range(self.gui.scripture_frame_layout.count()):
            component = self.gui.scripture_frame_layout.itemAt(i).widget()
            if isinstance(component, QLineEdit):
                sql += '"' + columns[index] + '" = "' + component.text().replace('"', '&quot') + '",'
                index += 1
            elif isinstance(component, QTextEdit):
                string = self.reformat_string_for_save(component.toMarkdown())
                sql += '"' + columns[index] + '" = "' + string + '",'
                index += 1
        for i in range(self.gui.exegesis_frame_layout.count()):
            component = self.gui.exegesis_frame_layout.itemAt(i).widget()
            if isinstance(component, QTextEdit) and not component.objectName() == 'textbox':
                string = self.reformat_string_for_save(component.toMarkdown())
                sql += '"' + columns[index] + '" = "' + string + '",'
                index += 1
        for i in range(self.gui.outline_frame_layout.count()):
            component = self.gui.outline_frame_layout.itemAt(i).widget()
            if isinstance(component, QTextEdit) and not component.objectName() == 'textbox':
                string = self.reformat_string_for_save(component.toMarkdown())
                sql += '"' + columns[index] + '" = "' + string + '",'
                index += 1
        for i in range(self.gui.research_frame_layout.count()):
            component = self.gui.research_frame_layout.itemAt(i).widget()
            if isinstance(component, QTextEdit) and not component.objectName() == 'textbox':
                string = self.reformat_string_for_save(component.toMarkdown())
                sql += '"' + columns[index] + '" = "' + string + '",'
                index += 1
        for i in range(self.gui.sermon_frame_layout.count()):
            component = self.gui.sermon_frame_layout.itemAt(i).widget()
            if isinstance(component, QLineEdit) or isinstance(component, QDateEdit):
                if isinstance(component, QLineEdit):
                    sql += '"' + columns[index] + '" = "' + component.text().replace('"', '&quot') + '",'
                else:
                    print(component.date().toString('yyyy-MM-dd'))
                    sql += '"' + columns[index] + '" = "' + component.date().toString('yyyy-MM-dd') + '",'
                index += 1
            elif isinstance(component, QTextEdit) and not component.objectName() == 'textbox':
                string = self.reformat_string_for_save(component.toMarkdown())
                sql += '"' + columns[index] + '" = "' + string + '" WHERE ID = "' + str(rec_id) + '"'

        conn = sqlite3.connect(self.db_loc)
        cur = conn.cursor()
        cur.execute(sql)
        conn.commit()

        from Dialogs import timed_popup
        timed_popup('Record Saved', 1000, self.gui.accent_color)
        self.write_to_log('Database saved - ' + self.db_loc)

        self.gui.changes = False

    # QTextEdit borks up the formatting when reloading markdown characters, so convert them to
    # HTML tags instead
    def reformat_string_for_save(self, string):
        string = string.replace('"', '&quot') # also handle quotes for the sake of SQL

        slice = re.findall('_.*?_', string)
        for item in slice:
            if '***' in item:
                new_string = item.replace('*', '')
                new_string = new_string.replace('_', '')
                new_string = '<u><i><strong>' + new_string + '</strong></i></u>'
                string = string.replace(item,  new_string)
            elif '**' in item:
                new_string = item.replace('*', '')
                new_string = new_string.replace('_', '')
                new_string = '<u><strong>' + new_string + '</strong></u>'
                string = string.replace(item, new_string)
            elif '*' in item:
                new_string = item.replace('*', '')
                new_string = new_string.replace('_', '')
                new_string = '<u><i>' + new_string + '</i></u>'
                string = string.replace(item, new_string)
            else:
                new_string = new_string.replace('_', '')
                new_string = '<u>' + new_string + '</u>'
                string = string.replace(item, new_string)

        slice = re.findall('\*\*\*.*?\*\*\*', string)
        for item in slice:
            new_string = item.replace('*', '')
            new_string = '<i><strong>' + new_string + '</strong></i>'
            string = string.replace(item, new_string)

        slice = re.findall('\*\*.*?\*\*', string)
        for item in slice:
            new_string = item.replace('*', '')
            new_string = new_string.replace('_', '')
            new_string = '<strong>' + new_string + '</strong>'
            string = string.replace(item, new_string)

        slice = re.findall('\*.*?\*', string)
        for item in slice:
            new_string = item.replace('*', '')
            new_string = new_string.replace('_', '')
            new_string = '<i>' + new_string + '</i>'
            string = string.replace(item, new_string)

        string = string.replace('*', '')
        return string

    # change the &quot string back to '"', which was changed to facilitate easier SQL commands
    def reformat_string_for_load(self, string):
        string = string.replace('&quot', '"')
        return string

    # function to search the text of all database entries to see if user's string is found
    def get_search_results(self, search_text):
        full_text_result_list = []
        conn = sqlite3.connect(self.db_loc)
        cur = conn.cursor()
        all_data = cur.execute('SELECT * FROM sermon_prep_database').fetchall()

        #search first for the full search text
        add_item = True
        for line in all_data:
            for item in line:
                    num_matches = str(item).lower().count(search_text.lower())
                    if num_matches > 0:
                        for a in full_text_result_list:
                            if a[0][0] == line[0]:
                                add_item = False
                        if add_item:
                            full_text_result_list.append([line, search_text, num_matches])
            add_item = True

        # then search for each individual word in the search text
        individual_word_result_list = []
        text_split = search_text.split(' ')
        words_found = []
        add_item = True
        for line in all_data:
            for item in line:
                try:
                    for word in text_split:
                        num_matches = str(item).lower().count(word.lower())
                        words_found.append(word)
                        add_word = True
                        cleaned_words_found = []

                        for word in words_found:
                            for new_word in cleaned_words_found:
                                if word == new_word:
                                    add_word = False
                            if add_word:
                                cleaned_words_found.append(word)

                    if num_matches > 0:
                        for a in full_text_result_list:
                            if a[0][0] == line[0]:
                                add_item = False
                        for a in individual_word_result_list:
                            if a[0][0] == line[0]:
                                add_item = False
                        if add_item:
                            individual_word_result_list.append([line, cleaned_words_found, num_matches])

                    words_found = []
                except Exception as ex:
                    print(ex)
            add_item = True

        # reorder the search results based on number of matches, full text first
        sorted_results = []
        while len(full_text_result_list) > 0:
            highest_num = -1
            highest_index = -1
            for i in range(0, len(full_text_result_list)):
                if full_text_result_list[i][2] > highest_num:
                    highest_num = full_text_result_list[i][2]
                    highest_index = i
            sorted_results.append(full_text_result_list[highest_index])
            full_text_result_list.pop(highest_index)

        while len(individual_word_result_list) > 0:
            highest_num = -1
            highest_index = -1
            for i in range(0, len(individual_word_result_list)):
                if individual_word_result_list[i][2] > highest_num:
                    highest_num = individual_word_result_list[i][2]
                    highest_index = i
            sorted_results.append(individual_word_result_list[highest_index])
            individual_word_result_list.pop(highest_index)

        return sorted_results

    # retrieve the first record of the database and set the current index to 0
    def first_rec(self):
        goon = True
        if self.gui.changes:
            goon = self.ask_save()
        if goon:
            self.current_rec_index = 0
            self.get_by_index(self.current_rec_index)

    # retrieve the previous record of the database and set the current index less by 1
    def prev_rec(self):
        goon = True
        if self.gui.changes:
            goon = self.ask_save()
        if goon:
            if self.current_rec_index != 0:
                self.current_rec_index = self.current_rec_index - 1
                self.get_by_index(self.current_rec_index)

    # retrieve the next record of the database and set the current index up by 1
    def next_rec(self):
        goon = True
        if self.gui.changes:
            goon = self.ask_save()
        if goon:
            if self.current_rec_index != len(self.ids) - 1:
                self.current_rec_index = self.current_rec_index + 1
                self.get_by_index(self.current_rec_index)

    # retrieve the last record of the database and set the current index to the length of the ids array
    def last_rec(self):
        goon = True
        if self.gui.changes:
            goon = self.ask_save()
        if goon:
            self.current_rec_index = len(self.ids) - 1
            self.get_by_index(self.current_rec_index)

    # check for changes, then create a new record by inserting an id # that is one higher than the largest
    def new_rec(self):
        goon = True
        if self.gui.changes:
            goon = self.ask_save()
        if goon:
            if len(self.ids) > 0:
                new_id = self.ids[len(self.ids) - 1] + 1
            else:
                new_id = 1

            sql = 'INSERT INTO "sermon_prep_database" ("ID", "date", "sermon_reference") VALUES(' + str(
                new_id) + ', "None", "None")'
            conn = sqlite3.connect(self.db_loc)
            cur = conn.cursor()
            cur.execute(sql)
            conn.commit()

            import time
            time.sleep(0.5) # prevent a database lock, just in case SQLite takes a bit to update
            self.get_ids()
            self.get_date_list()
            self.get_scripture_list()

            self.gui.top_frame.dates_cb.blockSignals(True)
            self.gui.top_frame.references_cb.blockSignals(True)
            self.gui.top_frame.references_cb.clear()
            for item in self.references:
                self.gui.top_frame.references_cb.addItem(item[0])
            self.gui.top_frame.dates_cb.clear()
            self.gui.top_frame.dates_cb.addItems(self.dates)
            self.gui.top_frame.dates_cb.blockSignals(False)
            self.gui.top_frame.references_cb.blockSignals(False)

            self.last_rec()

    # double-check that the user really wants to delete this record, then remove it from the database
    # finish by loading the last record into the GUI
    def del_rec(self):
        response = yes_no_cancel_box(
            'Really Delete?',
            'Really delete the current record?\nThis action cannot be undone',
            self.gui.background_color
        )

        if response == 0:
            sql = 'DELETE FROM sermon_prep_database WHERE ID = "' + str(self.ids[self.current_rec_index]) + '";'
            conn = sqlite3.connect(self.db_loc)
            cur = conn.cursor()
            cur.execute(sql)
            conn.commit()

            self.get_ids()
            self.get_date_list()
            self.get_scripture_list()

            self.gui.top_frame.dates_cb.blockSignals(True)
            self.gui.top_frame.references_cb.blockSignals(True)
            self.gui.top_frame.references_cb.clear()
            for item in self.references:
                self.gui.top_frame.references_cb.addItem(item[0])
            self.gui.top_frame.dates_cb.clear()
            self.gui.top_frame.dates_cb.addItems(self.dates)
            self.gui.top_frame.dates_cb.blockSignals(False)
            self.gui.top_frame.references_cb.blockSignals(False)

            self.last_rec()

    # function to ask the user if they would like to save their work
    def ask_save(self):
        response = yes_no_cancel_box(
            'Save Changes?',
            'Changes have been made. Save changes?',
            self.gui.background_color
        )

        self.write_to_log('askSave response: ' + str(response))
        if response == 0:
            self.save_rec()
            return True
        elif response == 1:
            return True
        else:
            return False

    # function to write various messages to a log file, takes a string message as an argument
    def write_to_log(self, string):
        log_file_loc = self.app_dir + '/log.txt'

        if not exists(log_file_loc):
            logfile = open(log_file_loc, 'w')
            logfile.write('')
            logfile.close()

        string = str(datetime.now()) + ': ' + string + '\r\n'
        logfile = open(log_file_loc, 'a')
        logfile.writelines(string)
        logfile.close()

class LoadingBox(QDialog):
    def __init__(self):
        super().__init__()
        self.spd = SermonPrepDatabase()

        self.spd.cwd = os.getcwd().replace('\\', '/')
        if str(self.spd.cwd).endswith('src'):
            self.spd.cwd = self.spd.cwd.replace('src', '')
        else:
            self.spd.cwd = self.spd.cwd + '/'

        self.spd.finished.connect(self.end)
        self.spd.change_text.connect(self.change_text)
        self.spd.start()

        self.setWindowFlag(Qt.FramelessWindowHint)
        self.setModal(True)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setStyleSheet('background-color: transparent')
        self.setMinimumWidth(300)

        layout = QGridLayout()
        self.setLayout(layout)

        self.working_label = QLabel()
        self.working_label.setAutoFillBackground(False)
        movie = QMovie(self.spd.cwd + 'resources/waitIcon.webp')
        self.working_label.setMovie(movie)
        layout.addWidget(self.working_label, 0, 0, Qt.AlignHCenter)
        movie.start()

        self.status_label = QLabel('Starting...')
        self.status_label.setFont(QFont('Helvetica', 16, QFont.Bold))
        self.status_label.setStyleSheet('color: #d7d7f4; text-align: center;')
        layout.addWidget(self.status_label, 1, 0, Qt.AlignHCenter)

        self.show()

    def change_text(self, text):
        self.status_label.setText(text)
        app.processEvents()

    def end(self):
        self.spd.gui = GUI(self.spd)
        self.spd.current_rec_index = len(self.spd.ids) - 1
        self.spd.get_by_index(self.spd.current_rec_index)
        self.close()

if __name__ == '__main__':
    try:
        app = QApplication(sys.argv)
        loading_box = LoadingBox()
        app.exec()
    except Exception:
        logging.exception('')
