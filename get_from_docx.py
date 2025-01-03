import datetime
import os.path
import re
import tempfile
import zipfile
from os.path import exists
from xml.etree import ElementTree

from PyQt6.QtWidgets import QFileDialog, QMessageBox


class GetFromDocx:
    """
    GetFromDocx is a class that will parse data from .docx, .odt, or .txt files so they can be imported as
    database entries
    """
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
            # give the option to also recurse subdirectories of the user's folder
            response = QMessageBox.question(
                self.gui,
                'Search Subdirectories',
                'Also search subdirectories of this folder?\n(Choose "No" to only import files from this folder)',
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No | QMessageBox.StandardButton.Cancel
            )

            if response == QMessageBox.StandardButton.Yes:
                self.parse_files(folder, True)
            elif response == QMessageBox.StandardButton.No:
                self.parse_files(folder, False)

    def get_folder(self):
        """
        Method that creates a QFileDialog for the user to choose a directory to import files from
        """
        folder = QFileDialog.getExistingDirectory(
            None,
            'Choose Folder',
            os.path.expanduser('~/Documents'),
            QFileDialog.Option.ShowDirsOnly
        )
        return folder

    def parse_files(self, folder, recurse):
        """
        Method to parse the files contained in the user's directory.

        :param str folder: The user's chosen directory.
        :param boolean recurse: Recurse subdirectories.
        """
        self.gui.open_import_splash.emit()
        file_list = []

        # get a list of file names located in the user's directory
        if recurse:
            walk = os.walk(folder, True, None, True)
            for item in walk:
                self.gui.change_import_splash_dir.emit('Looking in ' + item[0])
                if len(item[2]) > 0:
                    for file in item[2]:
                        if '.docx' in file and not '~' in file: # skip over MS temporary files
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
        self.gui.change_import_splash_dir.emit('Converting')
        for i in range(0, len(file_list)):
            # first, attempt to extract reference and date from file name
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
                # first, unzip the .docx file and extract the document.xml file
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
                    errors.append([file_list[i], 'Not a valid .docx file'])
                    unzip_success = False

                if unzip_success:
                    document_file = unzip_folder + '/word/document.xml'

                    tree = ElementTree.parse(document_file)
                    root = tree.getroot()

                    sermon_text = ''
                    text_found = False
                    # iterate through the tags in document.xml and extract the paragraphs therein
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
                            sermon_text += '\n\n'

                    if not text_found:
                        errors.append([file_list[i], 'Unable to find any text in file'])
                    else:
                        sermons.append([date, reference, sermon_text, file_name])

            elif '.odt' in file_list[i].lower():
                # first, unzip the .odt file and extract content.xml
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
                    errors.append([file_list[i], 'Not a valid .odt file'])
                    unzip_success = False

                if unzip_success:
                    document_file = unzip_folder + '/content.xml'

                    tree = ElementTree.parse(document_file)
                    root = tree.getroot()

                    sermon_text = ''
                    # iterate through the tags in content.xml and extract the paragraphs therein
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
                        sermons.append([date, reference, sermon_text, file_name])
                    else:
                        errors.append([file_list[i], 'Unable to find any text in file'])

            elif '.txt' in file_list[i].lower():
                with open(file_list[i]) as file:
                    sermon_text = file.read()
                if len(sermon_text) > 0:
                    sermons.append([date, reference, sermon_text, file_name])
                else:
                    errors.append([file_list[i], 'Unable to find any text in file'])

        self.gui.spd.insert_imports(errors, sermons)
        self.gui.close_import_splash.emit()

        self.gui.tabbed_frame.setCurrentWidget(self.gui.sermon_frame)
