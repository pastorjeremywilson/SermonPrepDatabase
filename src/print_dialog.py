"""
@author Jeremy G. Wilson

Copyright 2023 Jeremy G. Wilson

This file is a part of the Sermon Prep Database program (v.4.2.3)

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

import subprocess

import wmi as wmi
from PyQt5.QtGui import QImage, QPixmap, QIcon
from PyQt5.QtWidgets import QWidget, QGridLayout, QLabel, QPushButton, QHBoxLayout, QComboBox, QDialog
from fitz import fitz


class PrintDialog(QDialog):
    """
    PrintDialog is a class that will let the user preview their printout, choose a printer, and execute a print job.
    """
    def __init__(self, pdf_file, gui, landscape=False):
        """
        :param str pdf_file: The file location of the PDF created by reportLab
        :param GUI gui: The program's GUI object
        :param boolean landscape: Whether the print should be landscape, defaults to False
        """
        super().__init__()

        self.gui = gui
        self.landscape = landscape
        self.pdf = fitz.open(pdf_file)
        self.pdf_file = pdf_file
        self.num_pages = len(self.pdf)
        self.pages = []
        self.current_page = 0

        self.layout = QGridLayout()
        self.layout.setRowStretch(0, 1)
        self.layout.setRowStretch(1, 1)
        self.layout.setRowStretch(2, 1)
        self.layout.setRowStretch(3, 100)
        self.layout.setRowStretch(4, 1)

        self.setLayout(self.layout)
        self.setWindowTitle('Print')
        self.setGeometry(50, 50, 100, 100)
        self.get_pages()
        self.init_components()

    def init_components(self):
        """
        Creates the components that will be included in this QDialog.
        """
        self.preview_label = QLabel('Preview:')
        self.preview_label.setFont(self.gui.bold_font)
        self.layout.addWidget(self.preview_label, 0, 0)

        self.pdf_label = QLabel()
        self.pdf_label.setPixmap(self.pages[0])
        self.pdf_label.setFixedSize(self.pages[0].size())
        self.layout.addWidget(self.pdf_label, 1, 0, 3, 1)

        nav_button_widget = QWidget()
        nav_button_layout = QHBoxLayout()
        nav_button_widget.setLayout(nav_button_layout)

        previous_button = QPushButton()
        previous_button.setIcon(QIcon(self.gui.spd.cwd + '/resources/spPrevRecIcon'))
        previous_button.setAutoFillBackground(False)
        previous_button.setStyleSheet('border: none')
        previous_button.pressed.connect(self.previous_page)
        nav_button_layout.addStretch()
        nav_button_layout.addWidget(previous_button)
        nav_button_layout.addSpacing(20)

        self.page_label = QLabel('Page 1 of ' + str(self.num_pages))
        self.page_label.setFont(self.gui.bold_font)
        nav_button_layout.addWidget(self.page_label)
        nav_button_layout.addSpacing(20)

        next_button = QPushButton()
        next_button.setIcon(QIcon(self.gui.spd.cwd + '/resources/svg/spNextRecIcon.svg'))
        next_button.setAutoFillBackground(False)
        next_button.setStyleSheet('border: none')
        next_button.pressed.connect(self.next_page)
        nav_button_layout.addWidget(next_button)
        nav_button_layout.addStretch()

        self.layout.addWidget(nav_button_widget, 4, 0)

        printer_label = QLabel('Send to:')
        printer_label.setFont(self.gui.bold_font)
        self.layout.addWidget(printer_label, 0, 1)

        printer_combobox = self.get_printers()
        printer_combobox.setFont(self.gui.standard_font)
        self.layout.addWidget(printer_combobox, 1, 1)

        ok_cancel_buttons = QWidget()
        ok_cancel_layout = QHBoxLayout()
        ok_cancel_buttons.setLayout(ok_cancel_layout)

        ok_button = QPushButton('Print')
        ok_button.setFont(self.gui.standard_font)
        ok_button.pressed.connect(lambda: self.do_print(printer_combobox.currentText()))
        ok_cancel_layout.addStretch()
        ok_cancel_layout.addWidget(ok_button)
        ok_cancel_layout.addSpacing(20)

        cancel_button = QPushButton('Cancel')
        cancel_button.setFont(self.gui.standard_font)
        cancel_button.pressed.connect(self.cancel)
        ok_cancel_layout.addWidget(cancel_button)
        ok_cancel_layout.addStretch()

        self.layout.addWidget(ok_cancel_buttons, 2, 1)

    def get_printers(self):
        """
        Obtain a list of printers from the system.
        """
        win_management = wmi.WMI()
        printers = win_management.Win32_Printer()

        printer_combobox = QComboBox()
        default = ''
        for printer in printers:
            if not printer.Hidden:
                printer_combobox.addItem(printer.Name)
            if printer.Default:
                default = printer.Name

        printer_combobox.setCurrentText(default)
        return printer_combobox

    def get_pages(self):
        """
        Method to extract pages from the pdf and convert them to pixmaps
        """
        for page in self.pdf:
            pixmap = page.get_pixmap()

            fmt = QImage.Format_RGBA8888 if pixmap.alpha else QImage.Format_RGB888
            q_pixmap = QPixmap.fromImage(QImage(pixmap.samples_ptr, pixmap.width, pixmap.height, fmt))

            self.pages.append(q_pixmap)

    def previous_page(self):
        """
        Method to change the preview to the previous page
        """
        if not self.current_page == 0:
            self.current_page -= 1
            self.page_label.setText('Page ' + str(self.current_page + 1) + ' of ' + str(self.num_pages))
            self.pdf_label.setPixmap(self.pages[self.current_page])

    def next_page(self):
        """
        Method to change the preview to the next page
        """
        if not self.current_page == self.num_pages - 1:
            self.current_page += 1
            self.page_label.setText('Page ' + str(self.current_page + 1) + ' of ' + str(self.num_pages))
            self.pdf_label.setPixmap(self.pages[self.current_page])

    def do_print(self, printer):
        """
        Method to create the print job
        """
        print('Opening print subprocess')
        CREATE_NO_WINDOW = 0x08000000
        if self.landscape:
            command = [
                    'ghostscript/gsprint.exe',
                    '-sDEVICE=mswinpr2',
                    '-sOutputFile="%printer%' + printer + '"',
                    '-landscape',
                    self.pdf_file,
                    '-ghostscript',
                    'ghostscript/gswin64c.exe'
            ]
        else:
            command = [
                'ghostscript/gsprint.exe',
                '-sDEVICE=mswinpr2',
                '-sOutputFile="%printer%' + printer + '"',
                self.pdf_file,
                '-ghostscript',
                'ghostscript/gswin64c.exe'
            ]
        print(command)

        p = subprocess.Popen(
            command,
            creationflags=CREATE_NO_WINDOW,
            stdout=subprocess.PIPE, stderr=subprocess.PIPE)

        print('Capturing print subprocess sdtout & stderr')
        stdout, stderr = p.communicate()
        print(stdout, stderr)
        self.done(0)

    def cancel(self):
        self.done(0)

    def closeEvent(self, evt):
        self.done(0)
