'''
@author Jeremy G. Wilson

Copyright 2023 Jeremy G. Wilson

This file is a part of the Sermon Prep Database program (v.3.3.7)

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

from PyQt5.QtCore import QSize, Qt, QTimer
from PyQt5.QtWidgets import QDialog, QGridLayout, QLabel, QPushButton, QVBoxLayout, QWidget, QHBoxLayout, QListWidget


# function to show a simple OK dialog box
def message_box(title, message, bg):
    dialog = QDialog()
    dialog.setWindowTitle(title)
    dialog.setStyleSheet('background: ' + bg)
    dialog.setModal(True)

    layout = QGridLayout()
    dialog.setLayout(layout)

    label = QLabel(message)
    layout.addWidget(label, 0, 0, 1, 3)

    ok_button = QPushButton('OK')
    ok_button.setFixedSize(QSize(50, 30))
    ok_button.pressed.connect(lambda: dialog.done(0))
    layout.addWidget(ok_button, 1, 1, 1, 1)

    dialog.exec()

# function to show a simple Yes/No/Cancel dialog box+
def yes_no_cancel_box(*args):
    try:
        #args = title, message, bg, (yes, no)
        result = -1

        dialog = QDialog()
        dialog.setStyleSheet('background: ' + args[2])

        dialog.setWindowTitle(args[0])

        layout = QVBoxLayout()
        dialog.setLayout(layout)

        label = QLabel(args[1])
        layout.addWidget(label)

        button_container = QWidget()
        container_layout = QHBoxLayout()
        button_container.setLayout(container_layout)

        yes_button = QPushButton('Yes')
        yes_button.pressed.connect(lambda: dialog.done(0))
        container_layout.addWidget(yes_button)

        container_layout.addSpacing(10)

        no_button = QPushButton('No')
        no_button.pressed.connect(lambda: dialog.done(1))
        container_layout.addWidget(no_button)

        container_layout.addSpacing(10)

        cancel_button = QPushButton('Cancel')
        cancel_button.pressed.connect(lambda: dialog.done(2))
        container_layout.addWidget(cancel_button)

        if len(args) == 5:
            yes_button.setText(args[3])
            no_button.setText(args[4])

        layout.addWidget(button_container)

        response = dialog.exec()
    except Exception:
        logging.exception('')
    return response

# function to show a timed popup message, takes a message string and milliseconds to display as arguments
def timed_popup(message, millis, bg):
        dialog = QDialog()
        dialog.setWindowFlag(Qt.FramelessWindowHint)
        dialog.setWindowOpacity(0.75)
        dialog.setBaseSize(QSize(200, 75))
        dialog.setStyleSheet('background-color: ' + bg)
        dialog.setModal(True)

        layout = QVBoxLayout()
        dialog.setLayout(layout)

        label = QLabel(message)
        label.setStyleSheet('font-size: 18pt; color: white;')
        layout.addWidget(label)

        timer = QTimer()
        timer.singleShot(millis, lambda: dialog.done(0))

        dialog.show()

class RemoveCustomWords:
    def __init__(self, spd):
        self.changes = False

        self.spd = spd
        self.widget = QWidget()
        self.widget.setWindowTitle('Sermon Prep Database - Remove Custom Words')

        layout = QVBoxLayout()
        self.widget.setLayout(layout)

        with open(spd.cwd + 'resources/custom_words.txt') as file:
            custom_words = file.readlines()

        self.word_list = QListWidget()
        for word in custom_words:
            self.word_list.addItem(word.strip())
        layout.addWidget(self.word_list)

        button_widget = QWidget()
        button_layout = QHBoxLayout()
        button_widget.setLayout(button_layout)

        remove_button = QPushButton('Remove')
        remove_button.pressed.connect(self.remove_word)
        button_layout.addWidget(remove_button)

        close_button = QPushButton('Close')
        close_button.pressed.connect(self.close)
        button_layout.addWidget(close_button)

        layout.addWidget(button_widget)

        self.widget.show()

    def remove_word(self):
        self.changes = True
        self.word_list.takeItem(self.word_list.currentRow())

    def close(self):
        if self.changes:
            lines = []
            file = open(self.spd.cwd + 'resources/custom_words.txt', 'w')
            try:
                for r in range(len(self.word_list)):
                    lines.append(self.word_list.item(r).text() + '\n')
                file.writelines(lines)
                file.close()
            except Exception as ex:
                self.spd.write_to_log(ex)
        self.widget.close()
