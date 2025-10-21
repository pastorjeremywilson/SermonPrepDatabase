import os
import re
import shutil
import sqlite3
from os.path import exists

from PyQt6.QtWidgets import QFileDialog, QDialog, QGridLayout, QLabel, QProgressBar, QPushButton, QMessageBox


class ConvertDatabase(QDialog):
    """
    ConvertDatabase is a class to perform the necessary alterations to a database from an older version of Sermon
    Prep Database in order to make it work with this version.
    """
    def __init__(self, spd, type=None):
        super(ConvertDatabase, self).__init__()
        self.spd = spd

        if type == 'existing':
            shutil.copy(self.spd.db_loc, os.path.expanduser('~') + '/sermon_prep_database.db.old')
            result = self.ask_for_file(self.spd.db_loc)
        else:
            result = self.ask_for_file()

        if result == 1:
            self.setWindowTitle('Importing Records')
            self.setModal(True)

            self.progress_layout = QGridLayout()
            self.setLayout(self.progress_layout)

            self.progress_label = QLabel('Importing records')
            self.progress_layout.addWidget(self.progress_label, 0, 0)

            self.progress_bar = QProgressBar()
            self.progress_bar.setMinimum(0)
            self.progress_bar.setMaximum(100)
            self.progress_layout.addWidget(self.progress_bar, 1, 0)

            self.show()
            self.convert_database()

    def ask_for_file(self, file=None):
        """
        Method to create a QFileDialog to ask for the user's old database file
        """
        if not file:
            dialog = QFileDialog()
            dialog.setWindowTitle('Select Database')
            dialog.setNameFilter('Database File (*.db)')
            dialog.setFileMode(QFileDialog.FileMode.ExistingFile)
            dialog.setDirectory(os.path.expanduser('~'))

            dialog.exec()
            if not dialog.selectedFiles():
                file = None
            else:
                file = dialog.selectedFiles()[0]

        if not file:  # create a new database if user didn't choose a file or file wasn't provided
            self.spd.write_to_log('ConvertDatabase.__init__: No file selected')

            app_dir = os.path.expanduser('~') + '/AppData/Roaming/Sermon Prep Database'
            if not exists(app_dir):
                os.mkdir(app_dir)

            shutil.copy('resources/database_template.db', self.spd.db_loc)

            QMessageBox.information(
                None,
                'Database Created',
                'A new database has been created.',
                QMessageBox.StandardButton.Ok
            )
            return 0
        else:
            # if the database's user_settings table has the bgcolor column, this database doesn't need to be
            # converted
            if not self.spd.check_for_old_version() == -1:
                shutil.copy(file, self.spd.db_loc)
            else:
                self.spd.write_to_log('ConvertDatabase.__init__: Converting database from ' + file)

                # retrieve all the data from the user's previous database
                # do this first in case the old database is the same name and location as self.spd.db_loc
                conn = sqlite3.connect(file)
                cur = conn.cursor()
                sql = 'SELECT * FROM sermon_prep_database'
                results = cur.execute(sql)
                self.all_data = results.fetchall()

                if not exists(self.spd.app_dir):
                    os.mkdir(self.spd.app_dir)

                shutil.copy('resources/database_template.db', self.spd.db_loc)

                # remove the new user introduction record from the database template
                conn = sqlite3.connect(self.spd.db_loc)
                cur = conn.cursor()
                sql = 'DELETE FROM sermon_prep_database WHERE ID = 1'  # remove the initial entry from the database template
                cur.execute(sql)
                conn.commit()
                return 1

    def convert_database(self):
        try:
            id = []
            pericope_entry = []
            pericope_text = []
            reference_entry = []
            scripture_text = []
            fcft_entry = []
            gat_entry = []
            cpt_entry = []
            pb_entry = []
            fcfs_entry = []
            gas_entry = []
            cps_entry = []
            scripture_outline_text = []
            sermon_outline_text = []
            illustration_text = []
            research_text = []
            title_entry = []
            date_entry = []
            location_entry = []
            ctw_entry = []
            hor_entry = []
            sermon_text = []

            for row in self.all_data:
                self.progress_label.setText('Converting id # ' + str(row[0]))
                self.progress_bar.setValue(self.progress_bar.value() + 1)

                # add the data from the items in each row to the above declared lists
                id.append(row[0])
                pericope_entry.append(self.clean_and_strip(row[2]))
                pericope_text.append(self.clean_and_strip(row[3]))
                reference_entry.append(self.clean_and_strip(row[5]))
                scripture_text.append(self.clean_and_strip(row[4]))
                fcft_entry.append(self.clean_and_strip(row[9]))
                gat_entry.append(self.clean_and_strip(row[10]))
                cpt_entry.append(self.clean_and_strip(row[7]))
                pb_entry.append(self.clean_and_strip(row[8]))
                fcfs_entry.append(self.clean_and_strip(row[11]))
                gas_entry.append(self.clean_and_strip(row[12]))
                cps_entry.append(self.clean_and_strip(row[18]))
                scripture_outline_text.append(self.clean_and_strip(row[6]))
                sermon_outline_text.append(self.clean_and_strip(row[14]))
                illustration_text.append(self.clean_and_strip(row[15]))
                research_text.append(self.clean_and_strip(row[17]))
                title_entry.append(self.clean_and_strip(row[13]))
                date_entry.append(self.clean_and_strip(row[1]))
                location_entry.append(self.clean_and_strip(row[19]))
                ctw_entry.append(self.clean_and_strip(row[21]))
                hor_entry.append(self.clean_and_strip(row[20]))
                sermon_text.append(self.clean_and_strip(row[16]))

            for i in range(0, len(id)):
                self.progress_label.setText('Inserting id #' + str(id[i]) + ' into database')
                self.progress_bar.setValue(self.progress_bar.value() + 1)

                # build an SQL statement to add the data from the arrays to new records in the new database
                conn = sqlite3.connect(self.spd.db_loc)
                cur = conn.cursor()
                sql = 'INSERT INTO sermon_prep_database ("ID","pericope","pericope_texts","sermon_reference","sermon_scripture","fcft","gat","cpt","pb","fcfs","gas","cps","scripture_outline","sermon_outline","illustrations","research","sermon_title","date","location","call_to_worship","hymn_of_response","manuscript") VALUES ('
                sql += '"' + str(id[i]) + '",'
                sql += '"' + pericope_entry[i] + '",'
                sql += '"' + pericope_text[i] + '",'
                sql += '"' + reference_entry[i] + '",'
                sql += '"' + scripture_text[i] + '",'
                sql += '"' + fcft_entry[i] + '",'
                sql += '"' + gat_entry[i] + '",'
                sql += '"' + cpt_entry[i] + '",'
                sql += '"' + pb_entry[i] + '",'
                sql += '"' + fcfs_entry[i] + '",'
                sql += '"' + gas_entry[i] + '",'
                sql += '"' + cps_entry[i] + '",'
                sql += '"' + scripture_outline_text[i] + '",'
                sql += '"' + sermon_outline_text[i] + '",'
                sql += '"' + illustration_text[i] + '",'
                sql += '"' + research_text[i] + '",'
                sql += '"' + title_entry[i] + '",'
                sql += '"' + date_entry[i] + '",'
                sql += '"' + location_entry[i] + '",'
                sql += '"' + ctw_entry[i] + '",'
                sql += '"' + hor_entry[i] + '",'
                sql += '"' + sermon_text[i] + '");'
                cur.execute(sql)
                conn.commit()
        except sqlite3.OperationalError as err:  # catch problems with the opening of the old database
            self.spd.write_to_log('ConvertDatabase.convertDatabase: ' + str(err))
            QMessageBox.critical(
                None,
                'Invalid Database',
                'Import failed. Please restart the program and try again.\n\nError:\n' + str(err),
                QMessageBox.StandardButton.Ok
            )

            os.remove(self.spd.db_loc)

            quit()

        else:  # if no errors, confirm database conversion
            self.progress_label.setText('Database successfully imported')
            self.progress_bar.setValue(100)

            continue_button = QPushButton('Continue')
            continue_button.pressed.connect(self.close)
            self.progress_layout.addWidget(continue_button, 2, 0)

    def clean_and_strip(self, string_in):
        """
        Method to convert any characters that won't work properly with the new version
        """
        if string_in:
            string_out = string_in.strip()
            string_out = re.sub('"', '\'', string_out)
            string_out = re.sub(' +', ' ', string_out)
            string_out = re.sub('\n', '<br />', string_out)

            return string_out
        else:
            return ""
