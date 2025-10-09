import os
from os.path import exists

from PyQt6.QtCore import QRunnable
from symspellpy import SymSpell


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
