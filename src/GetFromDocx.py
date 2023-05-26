"""
@author Jeremy G. Wilson

Copyright 2023 Jeremy G. Wilson

This file is a part of the Sermon Prep Database program (v.3.4.4)

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
import datetime
import os.path
import re
import tempfile
import zipfile
from os.path import exists
from xml.etree import ElementTree

from PyQt5.QtWidgets import QFileDialog, QMessageBox


class GetFromDocx:
    def __init__(self, gui):
        self.gui = gui
        self.books = [
            'genesis',
            'exodus',
            'leviticus',
            'numbers',
            'deuteronomy',
            'joshua',
            'judges',
            'ruth',
            'samuel',
            'kings',
            'chronicles',
            'ezra',
            'nehemiah',
            'esther',
            'job',
            'psalms',
            'psalm',
            'proverbs',
            'ecclesiastes',
            'song',
            'isaiah',
            'jeremiah',
            'lamentations',
            'ezekiel',
            'daniel',
            'hosea',
            'joel',
            'amos',
            'obadiah',
            'jonah',
            'micah',
            'nahum',
            'habakkuk',
            'zephaniah',
            'haggai',
            'zechariah',
            'malachi',
            'matthew',
            'mark',
            'luke',
            'john',
            'acts',
            'romans',
            'corinthians',
            'galatians',
            'ephesians',
            'philippians',
            'colossians',
            'thessalonians',
            'timothy',
            'titus',
            'philemon',
            'hebrews',
            'james',
            'peter',
            'john',
            'jude',
            'revelation'
        ]
        folder = self.get_folder()
        if folder:
            response = QMessageBox.question(
                self.gui.win,
                'Search Subdirectories',
                'Also search subdirectories of this folder?\n(Choose "No" to only import files from this folder)',
                QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel
            )

            if response == QMessageBox.Yes:
                self.parse_files(folder, True)
            elif response == QMessageBox.No:
                self.parse_files(folder, False)

    def get_folder(self):
        folder = QFileDialog.getExistingDirectory(
            None,
            'Choose Folder',
            os.path.expanduser('~/Documents'),
            QFileDialog.ShowDirsOnly
        )
        return folder

    def parse_files(self, folder, recurse):
        self.gui.open_import_splash.emit()
        file_list = []

        if recurse:
            walk = os.walk(folder, True, None, True)
            for item in walk:
                self.gui.change_import_splash_dir.emit('Looking in ' + item[0])
                if len(item[2]) > 0:
                    for file in item[2]:
                        if '.docx' in file and not '~' in file:
                            file_list.append(item[0] + '/' + file)
                            self.gui.change_import_splash_file.emit('Found ' + file)
                        elif '.txt' in file or '.odt' in file:
                            file_list.append(item[0] + '/' + file)
                            self.gui.change_import_splash_file.emit('Found ' + file)
        else:
            self.gui.change_import_splash_dir.emit('Looking in ' + folder)
            for file in os.listdir(folder):
                if '.docx' in file and not '~' in file:
                    file_list.append(folder + '/' + file)
                    self.gui.change_import_splash_file.emit('Found ' + file)
                elif '.txt' in file or '.odt' in file:
                    file_list.append(folder + '/' + file)
                    self.gui.change_import_splash_file.emit('Found ' + file)

        errors = []
        sermons = []
        self.gui.change_import_splash_dir.emit('Importing')
        for i in range(0, len(file_list)):
            self.gui.change_import_splash_file.emit(file_list[i])
            file_name_split = file_list[i].split('/')
            file_name = file_name_split[len(file_name_split) - 1]
            period_split = file_name.split('.')
            if len(period_split) > 2:
                split = period_split
            else:
                split = file_name.split(' ')

            index = 0
            date = ''
            reference = ''
            date_error = True
            reference_error = True
            for item in split:
                try:
                    date = datetime.datetime.strptime(item, '%Y-%m-%d')
                    date = item
                    date_error = False
                except ValueError:
                    pass

                for book in self.books:
                    if book in item.lower():
                        try:
                            item = item[0].upper() + item[1:len(item)]
                            reference = item + ' ' + period_split[index + 1] + ':' + period_split[index + 2]
                            reference_error = False
                        except IndexError:
                            pass
                index += 1

            if date_error and reference_error:
                errors.append([file_list[i], 'Unable to parse date or scripture reference from file name'])
            elif date_error:
                errors.append([file_list[i], 'Unable to parse date from file name'])
            elif reference_error:
                errors.append([file_list[i], 'Unable to parse scripture reference from file name'])

            if '.docx' in file_list[i].lower():
                file_loc = file_list[i]
                unzip_folder = tempfile.gettempdir() + '/spd_zip'
                if not exists(unzip_folder):
                    os.mkdir(unzip_folder)
                unzip_success = False
                try:
                    with zipfile.ZipFile(file_loc, 'r') as zipped:
                        zipped.extractall(unzip_folder)
                        unzip_success = True
                except zipfile.BadZipfile:
                    unzip_success = False

                if unzip_success:
                    document_file = unzip_folder + '/word/document.xml'

                    tree = ElementTree.parse(document_file)
                    root = tree.getroot()

                    sermon_text = ''
                    text_found = False
                    for elem in root.iter():
                        tag = re.sub('{.*?}', '', elem.tag)
                        if tag == 'p':
                            for p_elem in elem.iter():
                                tag = re.sub('{.*?}', '', p_elem.tag)
                                if tag == 'r':
                                    for r_elem in p_elem.iter():
                                        tag = re.sub('{.*?}', '', r_elem.tag)
                                        if tag == 't':
                                            for t_elem in r_elem.iter():
                                                if len(sermon_text.strip()) > 0:
                                                    text_found = True
                                                sermon_text += str(t_elem.text)
                            sermon_text += '\n'

                    if not text_found:
                        errors.append([file_list[i], 'Unable to find any text in file'])
                    else:
                        sermons.append([date, reference, sermon_text])

            elif '.odt' in file_list[i].lower():
                file_loc = file_list[i]
                unzip_folder = tempfile.gettempdir() + '/spd_zip'
                if not exists(unzip_folder):
                    os.mkdir(unzip_folder)
                unzip_success = False
                try:
                    with zipfile.ZipFile(file_loc, 'r') as zipped:
                        zipped.extractall(unzip_folder)
                        unzip_success = True
                except zipfile.BadZipfile:
                    unzip_success = False

                if unzip_success:
                    document_file = unzip_folder + '/content.xml'

                    tree = ElementTree.parse(document_file)
                    root = tree.getroot()

                    sermon_text = ''
                    for elem in root.iter():
                        tag = re.sub('{.*?}', '', elem.tag)
                        if tag == 'document-content':
                            for doc_con in elem.iter():
                                tag = re.sub('{.*?}', '', doc_con.tag)
                                if tag == 'body':
                                    for bod in doc_con.iter():
                                        tag = re.sub('{.*?}', '', bod.tag)
                                        if tag == 'p':
                                            for p in bod.iter():
                                                for item in p.iter():
                                                    if item.text:
                                                        if not item.text in sermon_text:
                                                            sermon_text += item.text
                                            sermon_text += '\n'
                    if len(sermon_text) > 0:
                        sermons.append([date, reference, sermon_text])
                    else:
                        errors.append([file_list[i], 'Unable to find any text in file'])

            elif '.txt' in file_list[i].lower():
                with open(file_list[i]) as file:
                    sermon_text = file.read()
                if len(sermon_text) > 0:
                    sermons.append([date, reference, sermon_text])
                else:
                    errors.append([file_list[i], 'Unable to find any text in file'])

        self.gui.spd.insert_imports(errors, sermons)
