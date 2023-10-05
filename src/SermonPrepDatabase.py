"""
Author: Jeremy G. Wilson

Copyright: 2023 Jeremy G. Wilson

This file is a part of the Sermon Prep Database program (v.4.0.1)

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
import os
import pickle
import re
import sqlite3
import sys
import time
from PyQt5.QtCore import QThread, Qt, pyqtSignal
from PyQt5.QtGui import QFont, QMovie
from PyQt5.QtWidgets import QApplication, QLineEdit, QTextEdit, QDateEdit, QLabel, QGridLayout, QDialog, QVBoxLayout, \
    QMessageBox, QWidget
from datetime import datetime
from os.path import exists
from sqlite3 import OperationalError
from symspellpy import SymSpell

from gui import GUI


class SermonPrepDatabase(QThread):
    """
    The main program class that handles startup functions such as checking for/creating a new database, instantiating
    the gui, and polling the database for data. Also handles any database reading and writing functions.
    """
    change_text = pyqtSignal(str)
    finished = pyqtSignal()

    gui = None
    ids = []
    dates = []
    references = []
    db_loc = None
    app_dir = None
    bible_file = None
    disable_spell_check = None
    auto_fill = None
    current_rec_index = 0
    user_settings = None
    app = None
    platform = ''
    cwd = ''
    sym_spell = None

    def run(self):
        """
        On startup, initialize a QApplication, get the platform, set the app_dir and db_loc, instantiate the GUI
        Use the change_text signal to alter the text on the splash screen
        """
        time.sleep(1.0)
        self.change_text.emit('Getting Platform')
        time.sleep(0.5)
        self.platform = sys.platform

        try:
            self.change_text.emit('Getting Directories')
            time.sleep(0.5)
            user_dir = os.path.expanduser('~')

            #set the location of the user files differently depending on if we're on windows or linux
            if self.platform == 'win32':
                self.app_dir = user_dir + '/AppData/Roaming/Sermon Prep Database'
            elif self.platform == 'linux':
                self.app_dir = user_dir + '/.sermonPrepDatabase'
                os.environ['QT_QPA_PLATFORM'] = 'offscreen'
            self.db_loc = self.app_dir + '/sermon_prep_database.db'

            # create the user files directory if it doesn't exist
            if not exists(self.app_dir):
                os.mkdir(self.app_dir)

            self.write_to_log('platform is ' + self.platform)
            self.write_to_log('current working directory is ' + self.cwd)
            self.write_to_log('application directory is ' + self.app_dir)
            self.write_to_log('database location is ' + self.db_loc)

            if not exists(self.app_dir + '/custom_words.txt'):
                with open(self.app_dir + '/custom_words.txt', 'w'):
                    pass

            if exists(self.app_dir + '/my_bible.xml'):
                self.bible_file = self.app_dir + '/my_bible.xml'

            if exists(self.db_loc):
                self.disable_spell_check = self.check_spell_check()
                self.auto_fill = self.check_auto_fill()

                if not self.disable_spell_check:
                    self.change_text.emit('Loading Dictionaries')
                    time.sleep(0.5)
                    self.load_dictionary()
            else:
                self.disable_spell_check = 1

            self.change_text.emit('Creating GUI')
            time.sleep(0.5)

            self.finished.emit()

        except Exception as ex:
            self.write_to_log(str(ex), True)

    def check_spell_check(self):
        """
        Attempt to get the disable_spell_check value from the user's database. Create a disable_spell_check
        column if OperationalError is thrown.
        """
        try:
            conn = sqlite3.connect(self.db_loc)
            cursor = conn.cursor()
            result = cursor.execute('SELECT disable_spell_check FROM user_settings').fetchone()
            conn.close()

            if int(result[0]) == 0:
                return False
            else:
                return True
        except OperationalError:
            cursor.execute('ALTER TABLE user_settings ADD disable_spell_check TEXT;')
            conn.commit()
            cursor.execute('UPDATE user_settings SET disable_spell_check=0 WHERE ID="1";')
            conn.commit()
            conn.close()
            return False

    def check_auto_fill(self):
        """
        Attempt to get the auto_fill value from the user's database. Create a auto_fill
        column if OperationalError is thrown.
        """
        try:
            conn = sqlite3.connect(self.db_loc)
            cursor = conn.cursor()
            result = cursor.execute('SELECT auto_fill FROM user_settings').fetchone()
            conn.close()

            if int(result[0]) == 0:
                return False
            else:
                return True
        except OperationalError:
            cursor.execute('ALTER TABLE user_settings ADD auto_fill TEXT;')
            conn.commit()
            cursor.execute('UPDATE user_settings SET auto_fill=0 WHERE ID="1";')
            conn.commit()
            conn.close()
            return False

    def write_spell_check_changes(self):
        """
        Function to set the disable_spell_check value upon user input.
        """
        conn = sqlite3.connect(self.db_loc)
        cursor = conn.cursor()
        if self.disable_spell_check:
            cursor.execute('UPDATE user_settings SET disable_spell_check=1 WHERE ID="1";')
        else:
            cursor.execute('UPDATE user_settings SET disable_spell_check=0 WHERE ID="1";')
        conn.commit()
        conn.close

    def write_auto_fill_changes(self):
        """
        Function to set the auto_fill value upon user input.
        """
        conn = sqlite3.connect(self.db_loc)
        cursor = conn.cursor()
        if self.auto_fill:
            cursor.execute('UPDATE user_settings SET auto_fill=1 WHERE ID="1";')
        else:
            cursor.execute('UPDATE user_settings SET auto_fill=0 WHERE ID="1";')
        conn.commit()
        conn.close

    def load_dictionary(self):
        """
        Function to create a SymSpell object based on the default dictionary and the user's custom words list.
        For SymSpellPy documentation, see https://symspellpy.readthedocs.io/en/latest/index.html
        """
        if not exists(self.cwd + 'resources/default_dictionary.pkl'):
            self.sym_spell = SymSpell()
            self.sym_spell.create_dictionary(self.cwd + 'resources/default_dictionary.txt')
            self.sym_spell.save_pickle(os.path.normpath(self.cwd + 'resources/default_dictionary.pkl'))
        else:
            self.sym_spell = SymSpell()
            self.sym_spell.load_pickle(os.path.normpath(self.cwd + 'resources/default_dictionary.pkl'))

        with open(self.app_dir + '/custom_words.txt', 'r') as file:
            custom_words = file.readlines()
        for entry in custom_words:
            self.sym_spell.create_dictionary_entry(entry.strip(), 1)

    def add_to_dictionary(self, widget, word):
        """
        Function to add a word to the user's custom words file upon user input.

        :param QWidget widget: the widget from which the word was added (so that it can be rechecked with the new word)
        :param str word: the word to be added
        """
        try:
            self.sym_spell.create_dictionary_entry(word, 1)
            with open(self.app_dir + '/custom_words.txt', 'a') as file:
                file.write(word + '\n')
            widget.check_whole_text()

        except Exception as ex:
            self.write_to_log(str(ex), True)

    def get_ids(self):
        """
        Function to retrieve all ID numbers from the user's database.
        """
        self.ids = []
        conn = sqlite3.connect(self.db_loc)
        cur = conn.cursor()
        results = cur.execute('SELECT ID FROM sermon_prep_database').fetchall()
        for item in results:
            self.ids.append(item[0])

    # retrieve the list of dates from the database
    def get_date_list(self):
        """
        Function to retrieve all dates from the user's database.
        """
        self.dates = []
        conn = sqlite3.connect(self.db_loc)
        cur = conn.cursor()
        results = cur.execute('SELECT date FROM sermon_prep_database').fetchall()
        for item in results:
            self.dates.append(item[0])

    def get_scripture_list(self):
        """
        Function to retrieve all scripture references from the user's database.
        """
        self.references = []
        conn = sqlite3.connect(self.db_loc)
        cur = conn.cursor()
        self.references = cur.execute(
            'SELECT sermon_reference, ID FROM sermon_prep_database ORDER BY sermon_reference').fetchall()

    def get_user_settings(self):
        """
        Function to retrieve all user settings from the user's database.
        """
        conn = sqlite3.connect(self.db_loc)
        cur = conn.cursor()
        self.user_settings = cur.execute('SELECT * FROM user_settings').fetchall()[0]

    def backup_db(self):
        """
        Make a copy of the user's database, appending the date and time to the file name. Removes the oldest
        file if the number of backups will be greater than 5.
        """
        from datetime import datetime
        now = datetime.now().strftime("%m-%d-%Y_%H-%M-%S")
        backup_file = self.app_dir + '/sermon_prep_database.backup-' + now + '.db'

        file_array = os.listdir(self.app_dir)
        backup_files = []
        for file in file_array:
            if 'backup-' in file:
                backup_files.append(file)

        if (len(backup_files) >= 5):  # keep only 5 prior copies of the database
            os.remove(self.app_dir + '/' + backup_files[0])

        import shutil
        shutil.copy(self.db_loc, backup_file)
        self.write_to_log('New backup file created at ' + backup_file)

    def write_color_changes(self):
        """
        Function to save the user's color changes to the database.
        """
        sql = ('UPDATE user_settings SET bgcolor = "'
               + self.gui.accent_color
               + '", fgcolor = "'
               + self.gui.background_color
               + '" WHERE ID = 1;')
        conn = sqlite3.connect(self.db_loc)
        cur = conn.cursor()
        cur.execute(sql)
        conn.commit()
        self.get_user_settings()

    def write_font_changes(self, family, size):
        """
        Function to save the user's font changes to the database.

        :param str family: Name of the font family.
        :param str size: Size of the font.
        """
        sql = 'UPDATE user_settings SET font_family = "' + family + '", font_size = "' + size + '" WHERE ID = 1;'
        conn = sqlite3.connect(self.db_loc)
        cur = conn.cursor()
        cur.execute(sql)
        conn.commit()
        self.get_user_settings()

    # save user's label changes to the database
    def write_label_changes(self, new_labels):
        """
        Save the header label changes based on user input.
        :param list of str new_labels: The list of new label names.
        """
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
        """
        Function to retrieve a record from the user's database by id stored in self.current_rec_index
        """
        conn = sqlite3.connect(self.db_loc)
        cur = conn.cursor()
        results = cur.execute("SELECT * FROM sermon_prep_database WHERE ID = " + str(self.ids[self.current_rec_index]))
        record = results.fetchall()
        return record

    def get_by_index(self, index):
        """
        Function to retrieve a record based on a given index of self.ids.

        :param int index: Index number of self.ids
        """
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

    def save_rec(self):
        """
        Function to retrieve all data from all elements of the GUI and save it to the user's database.
        """
        try:
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
        except Exception as ex:
            self.write_to_log(str(ex), True)

    # QTextEdit borks up the formatting when reloading markdown characters, so convert them to
    # HTML tags instead
    def reformat_string_for_save(self, string):
        """
        Function to handle the formatting of markdown characters, converting them to HTML for saving.

        :param str string: The string to reformat.
        """
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
                new_string = item.replace('_', '')
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
        """
        Function to handle the formatting of a database string for insertion into a QTextEdit. Only quotes need to
        be handled as HTML tags will be handled by the QTextEdit.

        :param str string: The string to reformat.
        """
        string = string.replace('&quot', '"')
        return string

    def get_search_results(self, search_text):
        """
        Function to search the text of all database entries to see if the user's string is found.

        :param str search_text: User's search term(s)
        """
        search_text = search_text.strip()
        full_text_result_list = []
        conn = sqlite3.connect(self.db_loc)
        cur = conn.cursor()
        all_data = cur.execute('SELECT * FROM sermon_prep_database').fetchall()
        found_ids = []

        #search first for the full search text
        add_item = True
        for line in all_data:
            for item in line:
                num_matches = str(item).lower().count(search_text.lower())
                if num_matches > 0:
                    for id in found_ids:
                        if line[0] == id:
                            add_item = False
                    if add_item:
                        full_text_result_list.append([line, search_text, num_matches])
                        found_ids.append(line[0])
            add_item = True

        # then search for each individual word in the search text
        individual_word_result_list = []

        search_terms = []
        quotes = re.findall('".*?"', search_text)
        for item in quotes:
            search_terms.append(item.replace('"', '').strip())
            search_text = search_text.replace(item, '')

        search_split = search_text.split(' ')
        for item in search_split:
            if len(item) > 0:
                search_terms.append(item.strip())

        for line in all_data:
            already_found = False
            for id in found_ids:
                if line[0] == id:
                    already_found = True

            if not already_found:
                words_found_in_line = [None] * len(search_terms)
                add_item = False

                for item in line:
                    for i in range(len(search_terms)):
                        search_word = search_terms[i]
                        num_matches = str(item).lower().count(search_word.lower())
                        if num_matches > 0:
                            words_found_in_line[i] = True
                            add_item = True

                if add_item:
                    words_found = []
                    num_matches = 0
                    for i in range(len(search_terms)):
                        if words_found_in_line[i]:
                            words_found.append(search_terms[i])
                            num_matches += 1
                    individual_word_result_list.append([line, words_found, num_matches])
                    found_ids.append(line[0])

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

    def first_rec(self):
        """
        Retrieve the first record of the database and set the current index to 0.
        """
        goon = True
        if self.gui.changes:
            goon = self.ask_save()
        if goon:
            self.current_rec_index = 0
            self.get_by_index(self.current_rec_index)

    def prev_rec(self):
        """
        Retrieve the previous record of the database and set the current index less by 1.
        """
        goon = True
        if self.gui.changes:
            goon = self.ask_save()
        if goon:
            if self.current_rec_index != 0:
                self.current_rec_index = self.current_rec_index - 1
                self.get_by_index(self.current_rec_index)

    def next_rec(self):
        """
        Retrieve the next record of the database and set the current index more by 1.
        """
        goon = True
        if self.gui.changes:
            goon = self.ask_save()
        if goon:
            if self.current_rec_index != len(self.ids) - 1:
                self.current_rec_index = self.current_rec_index + 1
                self.get_by_index(self.current_rec_index)

    def last_rec(self):
        """
        Retrieve the last record of the database and set the current index to the highest index of the ids array.
        """
        goon = True
        if self.gui.changes:
            goon = self.ask_save()
        if goon:
            self.current_rec_index = len(self.ids) - 1
            self.get_by_index(self.current_rec_index)

    def new_rec(self):
        """
        Check for changes, then create a new record by inserting an id # that is one higher than the largest.
        """
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
            time.sleep(0.5)  # prevent a database lock, just in case SQLite takes a bit to update
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


    def del_rec(self):
        """
        Double-check that the user really wants to delete this record, then remove it from the database.
        Finish by loading the last record into the GUI.
        """
        response = QMessageBox.question(
            self.gui.win,
            'Really Delete?',
            'Really delete the current record?\nThis action cannot be undone',
            QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel
        )

        if response == QMessageBox.Yes:
            self.gui.changes = False
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

    def ask_save(self):
        """
        Function to ask the user if they would like to save their work before continuing with their recent action.
        """
        response = QMessageBox.question(
            self.gui.win,
            'Save Changes?',
            'Changes have been made. Save changes?',
            QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel
        )

        self.write_to_log('askSave response: ' + str(response))
        if response == QMessageBox.Yes:
            self.save_rec()
            return True
        elif response == QMessageBox.No:
            return True
        else:
            return False

    def write_to_log(self, string, critical=False):
        """
        Function to write various messages to the log file.

        :param str string: The text of the log message.
        :param boolean critical: Designate this as a critical error, showing a QMessageBox to the user.
        """
        if critical:
            QMessageBox.critical(None, 'Exception Thrown', 'An error has occurred:\n\n' + string)
        log_file_loc = self.app_dir + '/log.txt'

        # Create the log file if it doesn't yet exist.
        if not exists(log_file_loc):
            logfile = open(log_file_loc, 'w')
            logfile.write('')
            logfile.close()

        string = str(datetime.now()) + ': ' + string + '\r\n'
        logfile = open(log_file_loc, 'a')
        logfile.writelines(string)
        logfile.close()

    def insert_imports(self, errors, sermons):
        """
        Function to add sermons, imported from .docx or .txt files, to the user's database.

        :param list of str errors: Any errors encountered during the file parsing function.
        :param list of str sermons: The sermons gathered from the parsed files.
        """
        try:
            conn = sqlite3.connect(self.db_loc)
            cursor = conn.cursor()

            highest_num = 0
            for num in self.ids:
                if num > highest_num:
                    highest_num = num

            for sermon in sermons:
                highest_num += 1
                date = sermon[0]
                reference = sermon[1]
                text = self.reformat_string_for_save(sermon[2])
                title = sermon[3]
                sql = 'INSERT INTO sermon_prep_database (ID, date, sermon_reference, manuscript, sermon_title) VALUES("'\
                    + str(highest_num) + '", "' + date + '", "' + reference + '", "' + text + '", "' + title + '");'
                cursor.execute(sql)
                conn.commit()

            conn.close()

            import time
            time.sleep(0.5)  # prevent a database lock, just in case SQLite takes a bit to update
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

            message = str(len(sermons)) + ' sermons have been imported.'
            if len(errors) > 0:
                message += ' Error(s) occurred while importing. Would you like to view them now?'
                result = QMessageBox.question(
                    self.gui.win,
                    'Import Complete',
                    message,
                    QMessageBox.Yes | QMessageBox.No)
                
                if result == QMessageBox.Yes:
                    error_text = ''
                    for error in errors:
                        error_text += error[0] + ': ' + error[1] + '\n'

                    dialog = QDialog()
                    dialog.setWindowTitle('Import Errors')
                    layout = QVBoxLayout()
                    dialog.setLayout(layout)

                    label = QLabel('Errors:')
                    label.setFont(QFont(self.user_settings[3], int(self.user_settings[4]), QFont.Bold))
                    layout.addWidget(label)

                    text_edit = QTextEdit()
                    text_edit.setReadOnly(True)
                    text_edit.setMinimumWidth(1000)
                    text_edit.setLineWrapMode(QTextEdit.NoWrap)
                    text_edit.setFont(QFont(self.user_settings[3], int(self.user_settings[4])))
                    text_edit.setText(error_text)
                    layout.addWidget(text_edit)
                    dialog.exec()
            else:
                QMessageBox.information(None, 'Import Complete', message, QMessageBox.Ok)
        except Exception as ex:
            QMessageBox.critical(
                self.gui.win,
                'Error Occurred', 'An error occurred while importing:\n\n' + str(ex),
                QMessageBox.Ok
            )
            self.write_to_log(
                'From SermonPrepDatabase.insert_imports: ' + str(ex) + '\nMost recent sql statement: ' + text)

    def import_splash(self):
        """
        Function to apprise user of work being done while importing sermons.
        """
        self.widget = QWidget()
        self.widget.setStyleSheet('border: 3px solid black; background-color: ' + self.gui.background_color + ';')
        layout = QVBoxLayout()
        self.widget.setLayout(layout)

        importing_label = QLabel('Importing...')
        importing_label.setStyleSheet('border: none;')
        importing_label.setFont(QFont(self.gui.font_family, int(self.gui.font_size), QFont.Bold))
        layout.addWidget(importing_label)
        layout.addSpacing(50)

        self.dir_label = QLabel('Looking in...')
        self.dir_label.setStyleSheet('border: none;')
        self.dir_label.setFont(QFont(self.gui.font_family, int(self.gui.font_size)))
        layout.addWidget(self.dir_label)

        self.file_label = QLabel('Examining...')
        self.file_label.setStyleSheet('border: none;')
        self.file_label.setFont(QFont(self.gui.font_family, int(self.gui.font_size)))
        layout.addWidget(self.file_label)

        self.widget.setWindowModality(Qt.WindowModal)
        self.widget.setWindowFlag(Qt.FramelessWindowHint)
        self.widget.show()

    def change_dir(self, text):
        """
        Function to change the directory shown on the import splash.
        :param str text: Directory to show.
        """
        self.dir_label.setText(text)
        app.processEvents()

    def change_file(self, text):
        """
        Function to change the file name shown on the import splash.

        :param str text: File name to show.
        """
        self.file_label.setText(text)
        app.processEvents()

    def close_splash(self):
        """
        Function to close the import splash widget.
        """
        self.widget.deleteLater()


class LoadingBox(QDialog):
    """
    QDialog class that shows the user the process of starting up the application.

    :param QApplication app: The main QApplication for this program.
    """
    def __init__(self, app):
        try:
            super().__init__()
            self.spd = SermonPrepDatabase()
            self.spd.app = app

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
        except Exception as ex:
            self.spd.write_to_log(str(ex), True)

    def change_text(self, text):
        """
        Function to change the text shown on the splash screen.

        :param str text: The text to display.
        """
        self.status_label.setText(text)
        app.processEvents()

    def end(self):
        """
        Function to instantiate the GUI after all other preload processes have finished, then close the splash screen.
        """
        self.spd.gui = GUI(self.spd)
        self.spd.gui.open_import_splash.connect(self.spd.import_splash)
        self.spd.gui.change_import_splash_dir.connect(self.spd.change_dir)
        self.spd.gui.change_import_splash_file.connect(self.spd.change_file)
        self.spd.gui.close_import_splash.connect(self.spd.close_splash)

        # set the GUI to display the most recent record
        self.spd.current_rec_index = len(self.spd.ids) - 1
        self.spd.get_by_index(self.spd.current_rec_index)
        self.close()


# main entry point for the program
if __name__ == '__main__':
    app = QApplication(sys.argv)
    loading_box = LoadingBox(app)
    app.exec()
