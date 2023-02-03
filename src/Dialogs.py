'''
@author Jeremy G. Wilson

Copyright 2022 Jeremy G. Wilson

This file is a part of the Sermon Prep Database program (v.3.3.0)

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

from PyQt5.QtCore import QSize, Qt, QTimer
from PyQt5.QtWidgets import QDialog, QGridLayout, QLabel, QPushButton, QVBoxLayout, QWidget, QHBoxLayout


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

    if len(args) == 5:
        yes_button.setText(args[3])
        no_button.setText(args[4])

    container_layout.addSpacing(10)

    cancel_button = QPushButton('Cancel')
    cancel_button.pressed.connect(lambda: dialog.done(2))
    container_layout.addWidget(cancel_button)

    layout.addWidget(button_container)
    result = dialog.exec()
    return result

# function to show a timed popup message, takes a message string and milliseconds to display as arguments
def timed_popup(message, millis, bg):
    dialog = QDialog()
    dialog.setWindowFlag(Qt.FramelessWindowHint)
    dialog.setBaseSize(QSize(200, 75))
    dialog.setStyleSheet('background: ' + bg)
    dialog.setModal(True)

    layout = QVBoxLayout()
    dialog.setLayout(layout)

    label = QLabel(message)
    label.setStyleSheet('font-size: 18pt; color: white')
    layout.addWidget(label)

    timer = QTimer()
    timer.singleShot(millis, lambda: dialog.done(0))

    dialog.show()
