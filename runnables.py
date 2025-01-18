import os
import re
from os.path import exists

from PyQt6.QtCore import QRunnable, Qt
from PyQt6.QtGui import QTextCursor
from symspellpy import Verbosity, SymSpell


class SpellCheck(QRunnable):
    finished = False

    def __init__(self, widget=None, type=None, gui=None):
        super().__init__()
        self.widget = widget
        self.type = type
        self.gui = gui
        self.setAutoDelete(True)
        if self.widget:
            self.cursor = self.widget.textCursor()

    def run(self):
        # don't run a spell check if the widget is devoid of text
        if len(self.widget.toPlainText().strip()) == 0:
            return
        if self.type == 'whole':
            self.check_whole_text()
        elif self.type == 'previous':
            self.check_word('previous')
        elif self.type == 'current':
            self.check_word('current')

    def check_whole_text(self):
        """
        Method to check every word in this QTextEdit for spelling errors.
        """
        self.cursor = self.widget.textCursor()
        self.cursor.movePosition(QTextCursor.MoveOperation.Start)

        marked_word_indices = []
        while not self.gui.spd.disable_spell_check:
            self.cursor.select(QTextCursor.SelectionType.WordUnderCursor)
            word = self.cursor.selection().toPlainText()

            # if the position doesn't change after moving to the next character, this is the last word
            pos_before_move = self.cursor.position()
            self.cursor.movePosition(QTextCursor.MoveOperation.NextCharacter, QTextCursor.MoveMode.KeepAnchor)
            pos_after_move = self.cursor.position()
            if pos_before_move == pos_after_move:
                self.gui.clear_changes_signal.emit()
                break

            # if there's an apostrophe, check the next two characters for contraction letters
            if self.cursor.selection().toPlainText().endswith('\''):
                self.cursor.movePosition(QTextCursor.MoveOperation.NextCharacter, QTextCursor.MoveMode.KeepAnchor)
                if re.search('[a-z]$', self.cursor.selection().toPlainText()):
                    word = self.cursor.selection().toPlainText()
                    self.cursor.movePosition(QTextCursor.MoveOperation.NextCharacter, QTextCursor.MoveMode.KeepAnchor)
                    if re.search('[a-z]$', self.cursor.selection().toPlainText()):
                        word = self.cursor.selection().toPlainText()

            self.cursor.movePosition(QTextCursor.MoveOperation.PreviousCharacter, QTextCursor.MoveMode.KeepAnchor)

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
                    selection_start = self.cursor.selectionStart()
                    # if the first suggestion is the same as the word, then it's not spelled wrong
                    if not suggestions[0].term == cleaned_word:
                        marked_word_indices.append(selection_start)

            self.cursor.clearSelection()
            self.cursor.movePosition(QTextCursor.MoveOperation.NextWord)

        self.gui.set_text_color_signal.emit(self.widget, marked_word_indices, Qt.GlobalColor.red)
        self.gui.clear_changes_signal.emit()

    def check_word(self, word_type):
        """
        Method to spell-check a single word, either previous to the user's cursor or at the user's current cursor.
        Skips over punctuation if that is the first previous "word" found.
        """
        punctuations = [',', '.', '?', '!', ')', ';', ':', '-', '\n', ' ']

        if not self.widget:
            return
        self.cursor = self.widget.textCursor()
        if word_type == 'previous':
            self.cursor.movePosition(QTextCursor.MoveOperation.PreviousWord)
            self.cursor.select(QTextCursor.SelectionType.WordUnderCursor)
            while (self.cursor.selection().toPlainText() in punctuations
                   or len(self.cursor.selection().toPlainText()) == 0):
                self.cursor.clearSelection()
                self.cursor.movePosition(QTextCursor.MoveOperation.PreviousWord)
                self.cursor.movePosition(QTextCursor.MoveOperation.PreviousWord)
                self.cursor.select(QTextCursor.SelectionType.WordUnderCursor)
                if self.cursor.position() == 0:
                    break
        else:
            self.cursor.select(QTextCursor.SelectionType.WordUnderCursor)

        word = self.cursor.selection().toPlainText()
        for punctuation in punctuations:
            if word == punctuation:
                self.cursor.clearSelection()
                self.cursor.movePosition(QTextCursor.MoveOperation.PreviousWord)
                self.cursor.movePosition(QTextCursor.MoveOperation.PreviousWord)
                self.cursor.select(QTextCursor.SelectionType.WordUnderCursor)
                word = self.cursor.selection().toPlainText()
                break

        # if there's an apostrophe, check the next two characters for contraction letters
        if self.cursor.selection().toPlainText().endswith('\''):
            self.cursor.movePosition(QTextCursor.MoveOperation.NextCharacter, QTextCursor.MoveMode.KeepAnchor)
            if re.search('[a-z]$', self.cursor.selection().toPlainText()):
                word = self.cursor.selection().toPlainText()
                self.cursor.movePosition(QTextCursor.MoveOperation.NextCharacter, QTextCursor.MoveMode.KeepAnchor)
                if re.search('[a-z]$', self.cursor.selection().toPlainText()):
                    word = self.cursor.selection().toPlainText()

        cleaned_word = self.clean_word(word)

        suggestions = None
        if len(cleaned_word) > 0 and not any(c.isnumeric() for c in cleaned_word):
            suggestions = self.gui.spd.sym_spell.lookup(cleaned_word, Verbosity.CLOSEST, max_edit_distance=2,
                                                            include_unknown=True)
            word_index = [self.cursor.selectionStart()]
            if suggestions:
                # if the first suggestion is the same as the word, then it's not spelled wrong
                if suggestions[0].term == cleaned_word:
                    self.gui.set_text_color_signal.emit(self.widget, word_index, Qt.GlobalColor.black)
                else:
                    self.gui.set_text_color_signal.emit(self.widget, word_index, Qt.GlobalColor.red)
            else: # if no suggestions, word is obviously misspelled
                self.gui.set_text_color_signal.emit(self.widget, word_index, Qt.GlobalColor.red)

        self.cursor.clearSelection()
        if 'previous' in word_type:
            self.cursor.movePosition(QTextCursor.MoveOperation.End)
            self.gui.reset_cursor_color_signal.emit(self.widget)

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
        if not exists('resources/default_dictionary.pkl'):
            self.main.sym_spell = SymSpell()
            self.main.sym_spell.create_dictionary('resources/default_dictionary.txt')
            self.main.sym_spell.save_pickle(os.path.normpath('resources/default_dictionary.pkl'))
        else:
            self.main.sym_spell = SymSpell()
            self.main.sym_spell.load_pickle(os.path.normpath('resources/default_dictionary.pkl'))

        with open(self.main.app_dir + '/custom_words.txt', 'r') as file:
            custom_words = file.readlines()
        for entry in custom_words:
            self.main.sym_spell.create_dictionary_entry(entry.strip(), 1)
