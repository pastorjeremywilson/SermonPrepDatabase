import os
import re
from os.path import exists

from PyQt5.QtCore import QRunnable, Qt
from PyQt5.QtGui import QTextCursor
from symspellpy import Verbosity, SymSpell


class SpellCheck(QRunnable):
    finished = False

    def __init__(self, widget=None, type=None, gui=None):
        super().__init__()
        self.widget = widget
        self.type = type
        self.gui = gui

    def run(self):
        if self.type == 'whole':
            self.check_whole_text()
        elif self.type == 'previous':
            self.check_previous_word()

    def check_whole_text(self):
        """
        Method to check every word in this QTextEdit for spelling errors.
        """

        cursor = self.widget.textCursor()
        cursor.movePosition(QTextCursor.Start)

        while not self.gui.spd.disable_spell_check:
            last_word = False
            cursor.select(cursor.WordUnderCursor)
            word = cursor.selection().toPlainText()

            # if the position doesn't change after moving to the next character, this is the last word
            pos_before_move = cursor.position()
            cursor.movePosition(QTextCursor.NextCharacter, cursor.KeepAnchor)
            pos_after_move = cursor.position()
            if pos_before_move == pos_after_move:
                last_word = True

            # if there's an apostrophe, check the next two characters for contraction letters
            if cursor.selection().toPlainText().endswith('\''):
                cursor.movePosition(QTextCursor.NextCharacter, cursor.KeepAnchor)
                if re.search('[a-z]$', cursor.selection().toPlainText()):
                    word = cursor.selection().toPlainText()
                    cursor.movePosition(QTextCursor.NextCharacter, cursor.KeepAnchor)
                    if re.search('[a-z]$', cursor.selection().toPlainText()):
                        word = cursor.selection().toPlainText()

            cursor.movePosition(QTextCursor.PreviousCharacter, cursor.KeepAnchor)

            cleaned_word = self.clean_word(word)

            suggestions = None
            if len(cleaned_word) > 0 and not any(c.isnumeric() for c in cleaned_word):
                if any(h.isalpha() for h in cleaned_word):
                    suggestions = self.gui.spd.sym_spell.lookup(
                        cleaned_word, Verbosity.CLOSEST,
                        max_edit_distance=2,
                        include_unknown=True
                    )

                if suggestions:
                    char_format = cursor.charFormat()
                    if not suggestions[0].term == cleaned_word:
                        char_format.setForeground(Qt.red)
                        cursor.mergeCharFormat(char_format)

                    else:
                        if char_format.foreground() == Qt.red:
                            char_format.setForeground(Qt.black)
                            cursor.mergeCharFormat(char_format)

            cursor.clearSelection()
            cursor.movePosition(QTextCursor.NextWord)

            if last_word:
                break

        self.finished = True

    def check_previous_word(self):
        """
        Method to spell-check the word previous to the user's cursor. Skips over punctuation if that is the first
        previous "word" found.
        """
        self.widget.blockSignals(True) # we don't want changeevent fired while manipulating for spell check
        punctuations = [',', '.', '?', '!', ')', ';', ':', '-']

        cursor = self.widget.textCursor()
        cursor.movePosition(QTextCursor.PreviousWord)
        cursor.select(cursor.WordUnderCursor)

        word = cursor.selection().toPlainText()
        for punctuation in punctuations:
            if word == punctuation:
                cursor.clearSelection()
                cursor.movePosition(QTextCursor.PreviousWord)
                cursor.movePosition(QTextCursor.PreviousWord)
                cursor.select(cursor.WordUnderCursor)
                word = cursor.selection().toPlainText()
                break

        # if there's an apostrophe, check the next two characters for contraction letters
        if cursor.selection().toPlainText().endswith('\''):
            cursor.movePosition(QTextCursor.NextCharacter, cursor.KeepAnchor)
            if re.search('[a-z]$', cursor.selection().toPlainText()):
                word = cursor.selection().toPlainText()
                cursor.movePosition(QTextCursor.NextCharacter, cursor.KeepAnchor)
                if re.search('[a-z]$', cursor.selection().toPlainText()):
                    word = cursor.selection().toPlainText()

        cleaned_word = self.clean_word(word)

        suggestions = None
        if len(cleaned_word) > 0 and not any(c.isnumeric() for c in cleaned_word):
            if any(h.isalpha() for h in cleaned_word):
                suggestions = self.gui.spd.sym_spell.lookup(cleaned_word, Verbosity.CLOSEST, max_edit_distance=2,
                                                            include_unknown=True)

            if suggestions:
                char_format = cursor.charFormat()
                # if the first suggestion is the same as the word, then it's not spelled wrong
                if not suggestions[0].term == cleaned_word:
                    char_format.setForeground(Qt.red)
                    cursor.mergeCharFormat(char_format)
                else:
                    char_format.setForeground(Qt.black)
                    cursor.mergeCharFormat(char_format)

        cursor.clearSelection()
        cursor.movePosition(QTextCursor.NextWord)

        char_format = cursor.charFormat()
        char_format.setForeground(Qt.black)
        cursor.mergeCharFormat(char_format)
        self.widget.setTextCursor(cursor)

        self.widget.blockSignals(False)

    def check_single_word(self, word):
        cleaned_word = self.clean_word(word)

        suggestions = None
        if len(cleaned_word) > 0 and not any(c.isnumeric() for c in cleaned_word):
            if any(h.isalpha() for h in cleaned_word):
                suggestions = self.gui.spd.sym_spell.lookup(cleaned_word, Verbosity.CLOSEST, max_edit_distance=2,
                                                            include_unknown=True)

            if suggestions:
                return suggestions
            else:
                return -1

    def clean_word(self, word):
        """
        Method to strip any non-word characters as well as pluralizing apostrophes out of a word to be spell-checked.

        :param str word: Word to be cleaned.
        """
        chars = ['.', ',', ';', ':', '?', '!', '"', '...', '*', '-', '_',
                 '\n', '\u2026', '\u201c', '\u201d']
        single_quotes = ['\u2018', '\u2019']

        cleaned_word = word.lower().strip()

        for char in chars:
            cleaned_word = cleaned_word.replace(char, '')
        for single_quote in single_quotes:
            cleaned_word = cleaned_word.replace(single_quote, '\'')

        cleaned_word = cleaned_word.replace('\'s', '')
        cleaned_word = cleaned_word.replace('s\'', 's')
        cleaned_word = cleaned_word.replace("<[.?*]>", '')

        if cleaned_word.startswith('\''):
            cleaned_word = cleaned_word[1:len(cleaned_word)]
        if cleaned_word.endswith('\''):
            cleaned_word = cleaned_word[0:len(cleaned_word) - 1]

        # there's a chance that utf-8-sig artifacts will be attached to the word
        # encoding to utf-8 then decoding as ascii removes them
        cleaned_word = cleaned_word.encode('utf-8').decode('ascii', errors='ignore')

        return cleaned_word


class LoadDictionary(QRunnable):
    def __init__(self, main):
        super().__init__()
        self.main = main

    def run(self):
        """
        Method to create a SymSpell object based on the default dictionary and the user's custom words list.
        For SymSpellPy documentation, see https://symspellpy.readthedocs.io/en/latest/index.html
        """
        if not exists(self.main.cwd + '/resources/default_dictionary.pkl'):
            self.main.sym_spell = SymSpell()
            self.main.sym_spell.create_dictionary(self.main.cwd + '/resources/default_dictionary.txt')
            self.main.sym_spell.save_pickle(os.path.normpath(self.main.cwd + '/resources/default_dictionary.pkl'))
        else:
            self.main.sym_spell = SymSpell()
            self.main.sym_spell.load_pickle(os.path.normpath(self.main.cwd + '/resources/default_dictionary.pkl'))

        with open(self.main.app_dir + '/custom_words.txt', 'r') as file:
            custom_words = file.readlines()
        for entry in custom_words:
            self.main.sym_spell.create_dictionary_entry(entry.strip(), 1)