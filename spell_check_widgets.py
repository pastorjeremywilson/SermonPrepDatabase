import re

from PyQt6.QtCore import Qt
from PyQt6.QtGui import QAction, QTextCursor, QSyntaxHighlighter, QTextCharFormat, QTextOption
from PyQt6.QtWidgets import QTextEdit
from symspellpy import Verbosity

class SpellCheckTextEdit(QTextEdit):
    """
    SpellCheckTextEdit is an implementation of QTextEdit that adds spell-checking capabilities.
    Also modifies the standard context menu to include spell check suggestions and an option to remove
    extra white space from the text.

    :param GUI gui: The GUI object
    """
    def __init__(self, gui):
        super().__init__()
        self.setObjectName('custom_text_edit')
        self.gui = gui
        self.full_spell_check_done = False
        self.document().setDefaultStyleSheet(
            'p {'
                'font-family: ' + self.gui.main.user_settings['font_family'] + ';'
                'font-size: ' + self.gui.main.user_settings['font_size'] + 'pt;'
                'line-height: ' + self.gui.main.user_settings['line_spacing'] + ';'
            '}'
        )

        self.spell_check_highlighter = SpellCheckHighlighter(self.document(), self.gui)

        self.textChanged.connect(self.text_changed)

    def contextMenuEvent(self, evt):
        """
        @override
        Alters the standard context menu of this QTextEdit to include a list of spell-check words as well as an option
        to remove extraneous whitespace from the text.
        """
        menu = self.createStandardContextMenu()

        clean_whitespace_action = QAction("Remove extra whitespace")
        clean_whitespace_action.triggered.connect(self.clean_whitespace)
        menu.insertAction(menu.actions()[0], clean_whitespace_action)
        menu.insertSeparator(menu.actions()[1])

        if not self.gui.main.disable_spell_check:
            cursor = self.cursorForPosition(evt.pos())
            cursor.select(QTextCursor.SelectionType.WordUnderCursor)
            word = cursor.selection().toPlainText()
            cleaned_word = word.replace('’', '\'')
            cleaned_word = re.sub('[^a-z\']', '', cleaned_word.lower())
            cleaned_word = cleaned_word.replace('\'s', '')
            if cleaned_word.endswith('\''):
                cleaned_word = cleaned_word[:-1]
            if cleaned_word.startswith('\''):
                cleaned_word = cleaned_word[1:]

            if len(cleaned_word) > 0:
                upper = False
                if word[0].isupper():
                    upper = True

                suggestions = self.gui.main.sym_spell.lookup(
                    cleaned_word, Verbosity.CLOSEST, max_edit_distance=2, include_unknown=False)

                next_menu_index = 0
                spell_actions = {}

                number_of_suggestions = len(suggestions)
                if number_of_suggestions > 10: number_of_suggestions = 11

                for i in range(number_of_suggestions):
                    term = suggestions[i].term
                    if upper:
                        term = term[0].upper() + term[1:]
                    spell_actions['action% s' % str(i)] = QAction(term)
                    spell_actions['action% s' % str(i)].setData((cursor, term))
                    spell_actions['action% s' % str(i)].triggered.connect(self.replace_word)
                    menu.insertAction(menu.actions()[i], spell_actions['action% s' % str(i)])

                    next_menu_index = i + 1

                menu.insertSeparator(menu.actions()[next_menu_index])
                action = QAction('Add to dictionary')
                action.triggered.connect(lambda: self.gui.main.add_to_dictionary(self, cleaned_word))
                menu.insertAction(menu.actions()[next_menu_index + 2], action)
                menu.insertSeparator(menu.actions()[next_menu_index + 3])

            menu.exec(evt.globalPos())
            menu.close()

    def toSimplifiedHtml(self):
        """
        Method to strip unneeded html tags and convert others to simpler tags
        """
        string = self.toHtml()
        string = re.split('<body.*?>', string)[1]

        # preserve any desired tags from getting removed during the wholesale <.*?> removal
        string = re.sub('<p.*?>', '{p}', string)
        string = re.sub('</p>', '{/p}', string)
        string = re.sub('<ul.*?>', '{ul}', string)
        string = re.sub('</ul>', '{/ul}', string)
        string = re.sub('<li.*?>', '{li}', string)
        string = re.sub('</li>', '{/li}', string)

        bold_texts = re.findall('<span.*?font-weight.*?</span>', string)
        for text in bold_texts:
            new_text = re.sub('<.*?>', '', text)
            # keep spaces on the outside of formatting marks
            if new_text.startswith(' '):
                new_text = ' {b}' + new_text[1:]
            else:
                new_text = '{b}' + new_text
            if new_text.endswith(' '):
                new_text = new_text[:-1] + '{/b} '
            else:
                new_text = new_text + '{/b}'
            string = string.replace(text, new_text)

        italic_texts = re.findall('<span.*?font-style.*?</span>', string)
        for text in italic_texts:
            new_text = re.sub('<.*?>', '', text)
            # keep spaces on the outside of formatting marks
            if new_text.startswith(' '):
                new_text = ' {i}' + new_text[1:]
            else:
                new_text = '{i}' + new_text
            if new_text.endswith(' '):
                new_text = new_text[:-1] + '{/i} '
            else:
                new_text = new_text + '{/i}'
            string = string.replace(text, new_text)

        underline_texts = re.findall('<span.*?text-decoration.*?</span>', string)
        for text in underline_texts:
            new_text = re.sub('<.*?>', '', text)
            # keep spaces on the outside of formatting marks
            if new_text.startswith(' '):
                new_text = ' {u}' + new_text[1:]
            else:
                new_text = '{u}' + new_text
            if new_text.endswith(' '):
                new_text = new_text[:-1] + '{/u} '
            else:
                new_text = new_text + '{/u}'
            string = string.replace(text, new_text)

        # convert preserved tags back to their original form
        string = re.sub('<.*?>', '', string)
        string = string.strip().replace('{p}', '<p>')
        string = string.strip().replace('{/p}', '</p>\n')
        string = string.replace('{ul}', '<ul>')
        string = string.replace('{/ul}', '</ul>')
        string = string.replace('{li}', '<li>')
        string = string.replace('{/li}', '</li>\n')
        string = string.replace('{b}', '<b>')
        string = string.replace('{/b}', '</b>')
        string = string.replace('{i}', '<i>')
        string = string.replace('{/i}', '</i>')
        string = string.replace('{u}', '<u>')
        string = string.replace('{/u}', '</u>')

        return string.strip()

    def text_changed(self):
        self.gui.changes = True

    def replace_word(self):
        """
        Method to replace a misspelled word if the user chooses a replacement from the context menu.
        """
        sender = self.sender()
        cursor = sender.data()[0]
        term = sender.data()[1]

        self.setTextCursor(cursor)
        selected_word = self.textCursor().selectedText().strip()

        # check if the original word includes punctuation so that it can be preserved
        punctuation = ''
        if not selected_word[len(selected_word) - 1].isalpha():
            punctuation = selected_word[len(selected_word) - 1]

        self.textCursor().removeSelectedText()
        cursor = self.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.NextCharacter, QTextCursor.MoveMode.KeepAnchor)
        self.setTextCursor(cursor)

        cursor = self.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.PreviousCharacter, QTextCursor.MoveMode.MoveAnchor)
        self.setTextCursor(cursor)

        self.textCursor().insertText(term)

        self.gui.changes = True

    def clean_whitespace(self):
        """
        Method to remove duplicate spaces and tabs from the document.
        """
        component = self.gui.focusWidget()
        if isinstance(component, QTextEdit):
            string = component.toHtml()
            string = re.sub(' +', ' ', string)
            string = re.sub('\t+', '\t', string)

            component.setHtml(string)
        self.gui.changes = True


class SpellCheckLineEdit(QTextEdit):
    """
    SpellCheckLineEdit is an implementation of QTextEdit that adds the same spell-checking capabilities as above, but
    formats the QTextEdit to look and behave like a QLineEdit.

    :param GUI gui: The GUI object
    """
    def __init__(self, gui):
        super().__init__()
        self.gui = gui
        self.setWordWrapMode(QTextOption.WrapMode.NoWrap)
        self.setFixedHeight(30)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.spell_check_highlighter = SpellCheckHighlighter(self.document(), self.gui)
        self.textChanged.connect(self.text_changed)

    def text_changed(self):
        self.gui.changes = True

    def contextMenuEvent(self, evt):
        """
        @override
        Alters the standard context menu of this QTextEdit to include a list of spell-check words as well as an option
        to remove extraneous whitespace from the text.
        """
        menu = self.createStandardContextMenu()

        if not self.gui.main.disable_spell_check:
            cursor = self.cursorForPosition(evt.pos())
            cursor.select(QTextCursor.SelectionType.WordUnderCursor)
            word = cursor.selection().toPlainText()
            cleaned_word = word.replace('’', '\'')
            cleaned_word = re.sub('[^a-z\']', '', cleaned_word.lower())
            cleaned_word = cleaned_word.replace('\'s', '')
            if cleaned_word.endswith('\''):
                cleaned_word = cleaned_word[:-1]
            if cleaned_word.startswith('\''):
                cleaned_word = cleaned_word[1:]

            if len(cleaned_word) > 0:
                upper = False
                if word[0].isupper():
                    upper = True

                suggestions = self.gui.main.sym_spell.lookup(
                    cleaned_word, Verbosity.CLOSEST, max_edit_distance=2, include_unknown=False)

                next_menu_index = 0
                spell_actions = {}

                number_of_suggestions = len(suggestions)
                if number_of_suggestions > 10: number_of_suggestions = 11

                for i in range(number_of_suggestions):
                    term = suggestions[i].term
                    if upper:
                        term = term[0].upper() + term[1:]
                    spell_actions['action% s' % str(i)] = QAction(term)
                    spell_actions['action% s' % str(i)].setData((cursor, term))
                    spell_actions['action% s' % str(i)].triggered.connect(self.replace_word)
                    menu.insertAction(menu.actions()[i], spell_actions['action% s' % str(i)])

                    next_menu_index = i + 1

                menu.insertSeparator(menu.actions()[next_menu_index])
                action = QAction('Add to dictionary')
                action.triggered.connect(lambda: self.gui.main.add_to_dictionary(self, cleaned_word))
                menu.insertAction(menu.actions()[next_menu_index + 2], action)
                menu.insertSeparator(menu.actions()[next_menu_index + 3])

            menu.exec(evt.globalPos())
            menu.close()

    def replace_word(self):
        """
        Method to replace a misspelled word if the user chooses a replacement from the context menu.
        """
        sender = self.sender()
        cursor = sender.data()[0]
        term = sender.data()[1]

        self.setTextCursor(cursor)
        selected_word = self.textCursor().selectedText().strip()

        # check if the original word includes punctuation so that it can be preserved
        punctuation = ''
        if not selected_word[len(selected_word) - 1].isalpha():
            punctuation = selected_word[len(selected_word) - 1]

        self.textCursor().removeSelectedText()
        cursor = self.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.NextCharacter, QTextCursor.MoveMode.KeepAnchor)
        self.setTextCursor(cursor)

        cursor = self.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.PreviousCharacter, QTextCursor.MoveMode.MoveAnchor)
        self.setTextCursor(cursor)

        self.textCursor().insertText(term)

        self.gui.changes = True

    def text(self):
        return self.toPlainText()

    def keyPressEvent(self, evt):
        if evt.key() == Qt.Key.Key_Tab:
            evt.ignore()
            self.focusNextPrevChild(False)
        elif evt.key() == Qt.Key.Key_Backtab:
            evt.ignore()
            self.focusNextPrevChild(True)
        else:
            super().keyPressEvent(evt)


class SpellCheckHighlighter(QSyntaxHighlighter):
    def __init__(self, parent, gui):
        super().__init__(parent)
        self.gui = gui
        self.misspell_format = QTextCharFormat()
        self.misspell_format.setUnderlineColor(Qt.GlobalColor.red)
        self.misspell_format.setUnderlineStyle(QTextCharFormat.UnderlineStyle.SpellCheckUnderline)
        self.position = 0
        self.length = 0

    def highlightBlock(self, text):
        if not self.gui.main.user_settings['disable_spell_check']:
            basic_split = text.split()
            words = []
            for item in basic_split:
                if '-' in item:
                    split_item = item.split('-')
                    for word in split_item:
                        words.append(word)
                else:
                    words.append(item)

            current_pos = 0
            for word in words:
                cleaned_word = word.replace('’', '\'')
                cleaned_word = re.sub('[^a-z\']', '', cleaned_word.lower())
                cleaned_word = cleaned_word.replace('\'s', '')
                if cleaned_word.endswith('\''):
                    cleaned_word = cleaned_word[:-1]
                if cleaned_word.startswith('\''):
                    cleaned_word = cleaned_word[1:]

                suggestions = self.gui.main.sym_spell.lookup(
                    cleaned_word, Verbosity.CLOSEST, max_edit_distance=2, include_unknown=False)
                # if the first suggestion isn't the same as the word, or there are no suggestions, the word is spelled wrong
                if len(suggestions) == 0 or not suggestions[0].term == cleaned_word:
                    start = text.find(word, current_pos)
                    if start != -1:
                        self.setFormat(start, len(cleaned_word), self.misspell_format)
                        current_pos = start + len(word)
                else:
                    current_pos += len(word) + 1