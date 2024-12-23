from PyQt6.QtCore import QSize, Qt, QTimer
from PyQt6.QtGui import QFont
from PyQt6.QtWidgets import QDialog, QGridLayout, QLabel, QPushButton, QVBoxLayout, QWidget, QHBoxLayout, QListWidget


def message_box(title, message, bg):
    """
    Function to show a simple OK dialog box

    :param str title: Window title
    :param str message: Dialog message
    :param str bg: Dialog box's background color
    """
    dialog = QDialog()
    dialog.setWindowTitle(title)
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


def yes_no_cancel_box(*args):
    """
    Function to show a simple Yes/No/Cancel dialog box

    :param any args: Expects str title, str message, str bg, optional str replacement for Yes, optional str replacement for No
    """
    #args = title, message, bg, (yes, no)
    result = -1

    dialog = QDialog()

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
    return response


def timed_popup(gui, message, millis):
    """
    Function to show a timed popup message.

    :param GUI gui: The current instance of GUI
    :param str message: The message of the popup
    :param int millis: The time, in milliseconds, to display the message
    """
    dialog = QDialog()
    dialog.setObjectName('timed_popup')
    dialog.setParent(gui)
    dialog.setWindowFlag(Qt.WindowType.FramelessWindowHint)
    dialog.setModal(True)

    layout = QVBoxLayout()
    dialog.setLayout(layout)

    label = QLabel(message)
    label.setObjectName('timed_popup')
    label.setAlignment(Qt.AlignmentFlag.AlignCenter)
    label.setFont(QFont(gui.font_family, 24))
    label.adjustSize()
    dialog.setFixedSize(label.width() + 40, label.height() + 20)

    layout.addWidget(label, Qt.AlignmentFlag.AlignCenter)

    timer = QTimer()
    timer.singleShot(millis, lambda: dialog.done(0))

    dialog.show()

class RemoveCustomWords:
    """
    Class to enable the user to remove words they have added to the dictionary.
    """
    def __init__(self, spd):
        """
        :param SermonPrepDatabase spd: The SermonPrepDatabase object
        """
        self.changes = False
        self.removed_words = []

        self.spd = spd
        self.widget = QWidget()
        self.widget.setWindowTitle('Sermon Prep Database - Remove Custom Words')

        layout = QVBoxLayout()
        self.widget.setLayout(layout)

        # retrieve the words from the custom word list
        with open(spd.app_dir + '/custom_words.txt') as file:
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
        """
        Method to acknowledge that changes were made and to remove the word from the word list
        """
        self.changes = True
        self.removed_words.append(self.word_list.currentItem().text())
        self.word_list.takeItem(self.word_list.currentRow())

    def close(self):
        """
        If changes were made, write to the file. Close the widget.
        """
        if self.changes:
            lines = []
            file = open(self.spd.app_dir + '/custom_words.txt', 'w')
            try:
                for r in range(len(self.word_list)):
                    lines.append(self.word_list.item(r).text() + '\n')
                file.writelines(lines)
                file.close()

                for word in self.removed_words:
                    self.spd.sym_spell.delete_dictionary_entry(word)

                from gui import CustomTextEdit
                for widget in self.spd.gui.tabbed_frame.currentWidget().findChildren(CustomTextEdit):
                    widget.check_whole_text()
            except Exception as ex:
                self.spd.write_to_log(str(ex), True)
        self.widget.close()
