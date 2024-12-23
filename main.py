"""
Author: Jeremy G. Wilson

Copyright: 2024 Jeremy G. Wilson

This file, and the files contained in the distribution are parts of
the Sermon Prep Database program (v.5.0.3)

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
"""

import os
import re
import shutil
import sqlite3
import sys
import time
import traceback

from PyQt6.QtCore import Qt, QThreadPool
from PyQt6.QtGui import QFont, QPixmap
from PyQt6.QtWidgets import QLineEdit, QTextEdit, QDateEdit, QLabel, QDialog, QVBoxLayout, \
    QMessageBox, QWidget, QApplication
from datetime import datetime
from os.path import exists
from sqlite3 import OperationalError


class SermonPrepDatabase:
    """
    The main program class that handles startup methods such as checking for/creating a new database, instantiating
    the gui, and polling the database for data. Also handles any database reading and writing methods.
    """
    gui = None
    ids = []
    dates = []
    references = []
    db_loc = None
    app_dir = None
    bible_file = None
    disable_spell_check = None
    auto_fill = None
    line_spacing = None
    current_rec_index = 0
    user_settings = None
    app = None
    platform = ''
    cwd = ''
    sym_spell = None

    def __init__(self, gui):
        """
        On startup, initialize a QApplication, get the platform, set the app_dir and db_loc, instantiate the GUI
        Use the change_text signal to alter the text on the splash screen
        """
        os.chdir(os.path.dirname(__file__))
        self.gui = gui
        self.spell_check_thread_pool = QThreadPool()
        self.spell_check_thread_pool.setStackSize(256000000)
        self.load_dictionary_thread_pool = QThreadPool()
        self.app = QApplication(sys.argv)

    def get_system_info(self):
        self.platform = sys.platform

        try:
            # get the current data directory and the user's home directory
            user_dir = os.path.expanduser('~')

            # set the location of the user files differently depending on if we're on windows or linux
            if self.platform == 'win32':
                self.app_dir = user_dir + '/AppData/Roaming/Sermon Prep Database'
            elif self.platform == 'linux':
                self.app_dir = user_dir + '/.sermonPrepDatabase'
                os.environ['QT_QPA_PLATFORM'] = 'offscreen'
            self.db_loc = self.app_dir + '/sermon_prep_database.db'

            # create the user files directory if it doesn't exist
            if not exists(self.app_dir):
                os.mkdir(self.app_dir)

            self.write_to_log('application version is v.5.0.3')
            self.write_to_log('platform is ' + self.platform)
            self.write_to_log('current working directory is ' + os.path.dirname(__file__))
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
                self.line_spacing = self.check_line_spacing()
            else:
                self.disable_spell_check = 1

        except Exception as ex:
            self.write_to_log(str(ex), True)

    def check_for_db(self):
        """
        Check if the database file exists. Prompt to import an existing database or create a new database if not. If
        it exists, but does not include the line_spacing column in the user_settings table, then it is an old version.
        """
        if not exists(self.db_loc):
            response = QMessageBox.question(
                None,
                'Database Not Found',
                'It looks like this is the first time you\'ve run Sermon Prep Database v3.3.4.\n'
                'Would you like to import an old database?\n'
                '(Choose "No" to create a new database)',
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No | QMessageBox.StandardButton.Cancel
            )

            if response == QMessageBox.StandardButton.Yes:
                from convert_database import ConvertDatabase
                ConvertDatabase(self)
            elif response == QMessageBox.StandardButton.No:
                # Create a new database in the user's App Data directory by copying the existing database template
                shutil.copy('resources/database_template.db', self.db_loc)
                QMessageBox.information(None, 'Database Created', 'A new database has been created.',
                                        QMessageBox.StandardButton.Ok)
                self.gui.app.processEvents()
            else:
                quit(0)
        else:
            # check that the existing database is in the new format
            result = self.check_for_old_version()

            if result == -1:
                response = QMessageBox.question(
                    None,
                    'Old Database Found',
                    'It appears that you are upgrading from a previous version of Sermon Prep Database. Your database '
                    'file will need to be upgraded before you can continue. Upgrade now?',
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
                )

                if response == QMessageBox.StandardButton.Yes:
                    from convert_database import ConvertDatabase
                    ConvertDatabase(self, 'existing')
                else:
                    quit(0)
            else:
                # check that all columns since major update exist
                self.check_spell_check()
                self.check_auto_fill()
                self.check_line_spacing()
                self.check_font_color()
                self.check_text_background()

        self.write_to_log('checkForDB completed')

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

    def check_line_spacing(self):
        """
        Attempt to get the line_spacing value from the user's database. Create a line_spacing
        column if OperationalError is thrown.
        """
        try:
            conn = sqlite3.connect(self.db_loc)
            cursor = conn.cursor()
            result = cursor.execute('SELECT line_spacing FROM user_settings').fetchone()
            conn.close()

            return str(result[0])
        except OperationalError:
            cursor.execute('ALTER TABLE user_settings ADD line_spacing TEXT;')
            conn.commit()
            cursor.execute('UPDATE user_settings SET line_spacing=1.0 WHERE ID="1";')
            conn.commit()
            conn.close()
            return False

    def check_font_color(self):
        """
        Attempt to get the font_color value from the user's database. Create a font_color
        column if OperationalError is thrown.
        """
        try:
            conn = sqlite3.connect(self.db_loc)
            cursor = conn.cursor()
            result = cursor.execute('SELECT font_color FROM user_settings').fetchone()
            conn.close()

            return str(result[0])
        except OperationalError:
            cursor.execute('ALTER TABLE user_settings ADD font_color TEXT;')
            conn.commit()
            cursor.execute('UPDATE user_settings SET font_color="black" WHERE ID="1";')
            conn.commit()
            conn.close()
            return False

    def check_text_background(self):
        """
        Attempt to get the text_background value from the user's database. Create a text_background
        column if OperationalError is thrown.
        """
        try:
            conn = sqlite3.connect(self.db_loc)
            cursor = conn.cursor()
            result = cursor.execute('SELECT text_background FROM user_settings').fetchone()
            conn.close()

            return str(result[0])
        except OperationalError:
            cursor.execute('ALTER TABLE user_settings ADD text_background TEXT;')
            conn.commit()
            cursor.execute('UPDATE user_settings SET text_background="white" WHERE ID="1";')
            conn.commit()
            conn.close()
            return False

    def check_for_old_version(self):
        try:
            conn = sqlite3.connect(self.db_loc)
            cursor = conn.cursor()
            result = cursor.execute('SELECT bgcolor FROM user_settings').fetchone()
            conn.close()
        except OperationalError:
            conn.close()
            return -1

    def write_spell_check_changes(self):
        """
        Method to set the disable_spell_check value upon user input.
        """
        conn = sqlite3.connect(self.db_loc)
        cursor = conn.cursor()
        if str(self.gui.spell_check) == '1':
            cursor.execute('UPDATE user_settings SET disable_spell_check=1 WHERE ID="1";')
        else:
            cursor.execute('UPDATE user_settings SET disable_spell_check=0 WHERE ID="1";')
        conn.commit()
        conn.close()

    def write_auto_fill_changes(self):
        """
        Method to set the auto_fill value upon user input.
        """
        conn = sqlite3.connect(self.db_loc)
        cursor = conn.cursor()
        if self.auto_fill:
            cursor.execute('UPDATE user_settings SET auto_fill=1 WHERE ID="1";')
        else:
            cursor.execute('UPDATE user_settings SET auto_fill=0 WHERE ID="1";')
        conn.commit()
        conn.close()

    def write_line_spacing_changes(self):
        """
        Method to write user's line spacing choice to the database.
        """
        conn = sqlite3.connect(self.db_loc)
        cursor = conn.cursor()
        cursor.execute('UPDATE user_settings SET line_spacing=' + self.line_spacing + ' WHERE ID="1";')
        conn.commit()
        conn.close()

    def add_to_dictionary(self, widget, word):
        """
        Method to add a word to the user's custom words file upon user input.

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
        Method to retrieve all ID numbers from the user's database.
        """
        self.ids = []
        conn = sqlite3.connect(self.db_loc)
        cur = conn.cursor()
        results = cur.execute('SELECT ID FROM sermon_prep_database').fetchall()
        conn.close()

        for item in results:
            self.ids.append(item[0])

    def get_date_list(self):
        """
        Method to retrieve all dates from the user's database.
        """
        self.dates = []
        conn = sqlite3.connect(self.db_loc)
        cur = conn.cursor()
        results = cur.execute('SELECT date FROM sermon_prep_database').fetchall()
        conn.close()

        for item in results:
            self.dates.append(item[0])

    def get_scripture_list(self):
        """
        Method to retrieve all scripture references from the user's database.
        """
        self.references = []
        conn = sqlite3.connect(self.db_loc)
        cur = conn.cursor()
        self.references = cur.execute(
            'SELECT sermon_reference, ID FROM sermon_prep_database ORDER BY sermon_reference').fetchall()
        conn.close()

    def get_user_settings(self):
        """
        Method to retrieve all user settings from the user's database.
        """
        conn = sqlite3.connect(self.db_loc)
        cur = conn.cursor()
        self.user_settings = cur.execute('SELECT * FROM user_settings').fetchall()[0]
        conn.close()

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

    def write_color_changes(self, type):
        """
        Method to save the user's color changes to the database.
        """
        sql = ('UPDATE user_settings SET bgcolor = "' + type + '" WHERE ID = 1;')
        conn = sqlite3.connect(self.db_loc)
        cur = conn.cursor()
        cur.execute(sql)
        conn.commit()
        conn.close()

        self.get_user_settings()

    def write_font_changes(self, family, size):
        """
        Method to save the user's font changes to the database.

        :param str family: Name of the font family.
        :param str size: Size of the font.
        """
        sql = 'UPDATE user_settings SET font_family = "' + family + '", font_size = "' + size + '" WHERE ID = 1;'
        conn = sqlite3.connect(self.db_loc)
        cur = conn.cursor()
        cur.execute(sql)
        conn.commit()
        conn.close()

        self.get_user_settings()

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
        conn.close()

        self.get_user_settings()

    def get_record_data(self):
        """
        Method to retrieve a record from the user's database by id stored in self.current_rec_index
        """
        conn = sqlite3.connect(self.db_loc)
        cur = conn.cursor()
        results = cur.execute("SELECT * FROM sermon_prep_database WHERE ID = " + str(self.ids[self.current_rec_index]))
        record = results.fetchall()
        conn.close()

        return record

    def get_by_index(self, index):
        """
        Method to retrieve a record based on a given index of self.ids.

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

            self.gui.toolbar.dates_cb.blockSignals(True)
            self.gui.toolbar.references_cb.blockSignals(True)
            self.gui.toolbar.dates_cb.setCurrentIndex(index)
            self.gui.toolbar.references_cb.setCurrentIndex(counter)
            self.gui.toolbar.dates_cb.blockSignals(False)
            self.gui.toolbar.references_cb.blockSignals(False)

            conn = sqlite3.connect(self.db_loc)
            cur = conn.cursor()
            results = cur.execute("SELECT * FROM sermon_prep_database WHERE ID = " + str(self.ids[index]))
            record = results.fetchall()
            conn.close()

            if index == 0:
                self.gui.toolbar.first_rec_button.setEnabled(False)
                self.gui.toolbar.prev_rec_button.setEnabled(False)
                self.gui.toolbar.next_rec_button.setEnabled(True)
                self.gui.toolbar.last_rec_button.setEnabled(True)
            elif index == len(self.ids) - 1:
                self.gui.toolbar.first_rec_button.setEnabled(True)
                self.gui.toolbar.prev_rec_button.setEnabled(True)
                self.gui.toolbar.next_rec_button.setEnabled(False)
                self.gui.toolbar.last_rec_button.setEnabled(False)
            else:
                self.gui.toolbar.first_rec_button.setEnabled(True)
                self.gui.toolbar.prev_rec_button.setEnabled(True)
                self.gui.toolbar.next_rec_button.setEnabled(True)
                self.gui.toolbar.last_rec_button.setEnabled(True)

            self.gui.fill_values(record)
        else:
            self.new_rec()

    def save_rec(self):
        """
        Method to retrieve all data from all elements of the GUI and save it to the user's database.
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
                    sql += '"' + columns[index] + '" = "' + component.text().replace('"', '&quot;') + '",'
                    index += 1
                elif isinstance(component, QTextEdit):
                    string = component.toSimplifiedHtml()
                    sql += '"' + columns[index] + '" = "' + string + '",'
                    index += 1
            for i in range(self.gui.exegesis_frame_layout.count()):
                component = self.gui.exegesis_frame_layout.itemAt(i).widget()
                if isinstance(component, QTextEdit) and not component.objectName() == 'textbox':
                    string = component.toSimplifiedHtml()
                    sql += '"' + columns[index] + '" = "' + string + '",'
                    index += 1
            for i in range(self.gui.outline_frame_layout.count()):
                component = self.gui.outline_frame_layout.itemAt(i).widget()
                if isinstance(component, QTextEdit) and not component.objectName() == 'textbox':
                    string = component.toSimplifiedHtml()
                    sql += '"' + columns[index] + '" = "' + string + '",'
                    index += 1
            for i in range(self.gui.research_frame_layout.count()):
                component = self.gui.research_frame_layout.itemAt(i).widget()
                if isinstance(component, QTextEdit) and not component.objectName() == 'textbox':
                    string = component.toSimplifiedHtml()
                    sql += '"' + columns[index] + '" = "' + string + '",'
                    index += 1
            for i in range(self.gui.sermon_frame_layout.count()):
                component = self.gui.sermon_frame_layout.itemAt(i).widget()
                if isinstance(component, QLineEdit) or isinstance(component, QDateEdit):
                    if isinstance(component, QLineEdit):
                        sql += '"' + columns[index] + '" = "' + component.text().replace('"', '&quot;') + '",'
                    else:
                        sql += '"' + columns[index] + '" = "' + component.date().toString('yyyy-MM-dd') + '",'
                    index += 1
                elif isinstance(component, QTextEdit) and not component.objectName() == 'textbox':
                    string = component.toSimplifiedHtml()
                    sql += '"' + columns[index] + '" = "' + string + '" WHERE ID = "' + str(rec_id) + '"'

            conn = sqlite3.connect(self.db_loc)
            cur = conn.cursor()
            cur.execute(sql)
            conn.commit()
            conn.close()

            from dialogs import timed_popup
            timed_popup(self.gui, 'Record Saved', 1000)
            self.write_to_log('Database saved - ' + self.db_loc)

            self.gui.changes = False
        except Exception as ex:
            self.write_to_log(str(ex), True)

    def reformat_string_for_load(self, string):
        """
        Method to handle the formatting of an older-style database string for insertion into a QTextEdit. Only those
        strings missing paragraph markers as well as old-style &quots need to be handled.

        :param str string: The string to reformat.
        """
        string = string.strip()
        if '<p>' not in string:
            string_split = string.split('\n\n')
            string = '<p>' + '</p>\n<p>'.join(string_split) + '</p>'
        # replace any antiquated &quots without the semicolon
        string = re.sub(r'&amp;quot(?!;)', '"', string)
        string = re.sub(r'&quot(?!;)', '"', string)
        return string

    def get_search_results(self, search_text):
        """
        Method to search the text of all database entries to see if the user's string is found.

        :param str search_text: User's search term(s)
        """
        search_text = search_text.strip()
        full_text_result_list = []
        conn = sqlite3.connect(self.db_loc)
        cur = conn.cursor()
        all_data = cur.execute('SELECT * FROM sermon_prep_database').fetchall()
        conn.close()

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
            conn.close()

            import time
            time.sleep(0.5)  # prevent a database lock, just in case SQLite takes a bit to update
            self.get_ids()
            self.get_date_list()
            self.get_scripture_list()

            self.gui.toolbar.dates_cb.blockSignals(True)
            self.gui.toolbar.references_cb.blockSignals(True)
            self.gui.toolbar.references_cb.clear()
            for item in self.references:
                self.gui.toolbar.references_cb.addItem(item[0])
            self.gui.toolbar.dates_cb.clear()
            self.gui.toolbar.dates_cb.addItems(self.dates)
            self.gui.toolbar.dates_cb.blockSignals(False)
            self.gui.toolbar.references_cb.blockSignals(False)

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
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No | QMessageBox.StandardButton.Cancel
        )

        if response == QMessageBox.StandardButton.Yes:
            self.gui.changes = False
            sql = 'DELETE FROM sermon_prep_database WHERE ID = "' + str(self.ids[self.current_rec_index]) + '";'
            conn = sqlite3.connect(self.db_loc)
            cur = conn.cursor()
            cur.execute(sql)
            conn.commit()
            conn.close()

            self.get_ids()
            self.get_date_list()
            self.get_scripture_list()

            self.gui.toolbar.dates_cb.blockSignals(True)
            self.gui.toolbar.references_cb.blockSignals(True)
            self.gui.toolbar.references_cb.clear()
            for item in self.references:
                self.gui.toolbar.references_cb.addItem(item[0])
            self.gui.toolbar.dates_cb.clear()
            self.gui.toolbar.dates_cb.addItems(self.dates)
            self.gui.toolbar.dates_cb.blockSignals(False)
            self.gui.toolbar.references_cb.blockSignals(False)

            self.last_rec()

    def ask_save(self):
        """
        Method to ask the user if they would like to save their work before continuing with their recent action.
        """
        response = QMessageBox.question(
            self.gui.win,
            'Save Changes?',
            'Changes have been made. Save changes?',
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No | QMessageBox.StandardButton.Cancel
        )

        self.write_to_log('askSave response: ' + str(response))
        if response == QMessageBox.StandardButton.Yes:
            self.save_rec()
            return True
        elif response == QMessageBox.StandardButton.No:
            return True
        else:
            return False

    def write_to_log(self, string, critical=False):
        """
        Method to write various messages to the log file.

        :param str string: The text of the log message.
        :param boolean critical: Designate this as a critical error, showing a QMessageBox to the user.
        """
        if critical:
            QMessageBox.critical(None, 'Exception Thrown', 'An error has occurred:\n\n' + str(string))
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
        Method to add sermons, imported from .docx or .txt files, to the user's database.

        :param list of str errors: Any errors encountered during the file parsing method.
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
                text = sermon[2]
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

            self.gui.toolbar.dates_cb.blockSignals(True)
            self.gui.toolbar.references_cb.blockSignals(True)
            self.gui.toolbar.references_cb.clear()
            for item in self.references:
                self.gui.toolbar.references_cb.addItem(item[0])
            self.gui.toolbar.dates_cb.clear()
            self.gui.toolbar.dates_cb.addItems(self.dates)
            self.gui.toolbar.dates_cb.blockSignals(False)
            self.gui.toolbar.references_cb.blockSignals(False)

            self.last_rec()

            message = str(len(sermons)) + ' sermons have been imported.'
            if len(errors) > 0:
                message += ' Error(s) occurred while importing. Would you like to view them now?'
                result = QMessageBox.question(
                    self.gui.win,
                    'Import Complete',
                    message,
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                
                if result == QMessageBox.StandardButton.Yes:
                    error_text = ''
                    for error in errors:
                        error_text += error[0] + ': ' + error[1] + '\n'

                    dialog = QDialog()
                    dialog.setWindowTitle('Import Errors')
                    layout = QVBoxLayout()
                    dialog.setLayout(layout)

                    label = QLabel('Errors:')
                    label.setFont(QFont(self.user_settings[3], int(self.user_settings[4]), QFont.Weight.Bold))
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
                QMessageBox.information(None, 'Import Complete', message, QMessageBox.StandardButton.Ok)
        except Exception as ex:
            QMessageBox.critical(
                self.gui.win,
                'Error Occurred', 'An error occurred while importing:\n\n' + str(ex),
                QMessageBox.StandardButton.Ok
            )
            self.write_to_log(
                'From SermonPrepDatabase.insert_imports: ' + str(ex) + '\nMost recent sql statement: ' + text)

    def import_splash(self):
        """
        Method to apprise user of work being done while importing sermons.
        """
        self.widget = QWidget()
        layout = QVBoxLayout()
        self.widget.setLayout(layout)

        importing_label = QLabel('Importing...')
        importing_label.setFont(QFont(self.gui.font_family, int(self.gui.font_size), QFont.Weight.Bold))
        layout.addWidget(importing_label)
        layout.addSpacing(50)

        self.dir_label = QLabel('Looking in...')
        self.dir_label.setFont(QFont(self.gui.font_family, int(self.gui.font_size)))
        layout.addWidget(self.dir_label)

        self.file_label = QLabel('Examining...')
        self.file_label.setFont(QFont(self.gui.font_family, int(self.gui.font_size)))
        layout.addWidget(self.file_label)

        self.widget.setWindowModality(Qt.WindowModality.WindowModal)
        self.widget.setWindowFlag(Qt.WindowType.FramelessWindowHint)
        self.widget.show()

    def change_dir(self, text):
        """
        Method to change the directory shown on the import splash.
        :param str text: Directory to show.
        """
        self.dir_label.setText(text)
        self.gui.app.processEvents()

    def change_file(self, text):
        """
        Method to change the file name shown on the import splash.

        :param str text: File name to show.
        """
        self.file_label.setText(text)
        self.gui.app.processEvents()

    def close_splash(self):
        """
        Method to close the import splash widget.
        """
        self.widget.deleteLater()


def log_unhandled_exception(exc_type, exc_value, exc_traceback):
    """
    Provides a method for handling exceptions that aren't handled elsewhere in the program.
    :param exc_type:
    :param exc_value:
    :param exc_traceback:
    :return:
    """
    if issubclass(exc_type, KeyboardInterrupt):
        # Will call default excepthook
        sys.__excepthook__(exc_type, exc_value, exc_traceback)
        return

    exc_type = str(exc_type).replace('<class ', '')
    exc_type = exc_type.replace('>', '')
    full_traceback = str(traceback.StackSummary.extract(traceback.walk_tb(exc_traceback)))
    full_traceback = full_traceback.replace('[', '').replace(']', '')
    full_traceback = full_traceback.replace('<FrameSummary ', '')
    full_traceback = full_traceback.replace('>', '')
    full_traceback_split = full_traceback.split(',')
    formatted_traceback = ''
    for i in range(len(full_traceback_split)):
        if i == 0:
            formatted_traceback += full_traceback_split[i] + '\n'
        else:
            formatted_traceback += '    ' + full_traceback_split[i] + '\n'

    date_time = time.ctime(time.time())
    log_text = (f'\n{date_time}:\n'
                f'    UNHANDLED EXCEPTION\n'
                f'    {exc_type}\n'
                f'    {exc_value}\n'
                f'    {full_traceback}')
    with open(os.path.expanduser('~/AppData/Roaming/Sermon Prep Database/') + './error.log', 'a') as file:
        file.write(log_text)

    message_box = QMessageBox()
    message_box.setWindowTitle('Unhandled Exception')
    message_box.setIconPixmap(QPixmap('resources/svg/div-zero-bug.svg'))
    message_box.setText(
        '<strong>Well, that wasn\'t supposed to happen!</strong><br><br>An unhandled exception occurred:<br>'
        f'{exc_type}<br>'
        f'{exc_value}<br>'
        f'{full_traceback}')
    message_box.setStandardButtons(QMessageBox.StandardButton.Close)
    message_box.exec()


# main entry point for the program
if __name__ == '__main__':
    sys.excepthook = log_unhandled_exception
    from gui import GUI
    app = QApplication(sys.argv)
    gui = GUI()
    app.exec()
