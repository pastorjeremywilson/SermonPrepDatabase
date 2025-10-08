import os
from os.path import exists

from PyQt6.QtCore import QRunnable, QThreadPool
from symspellpy import SymSpell

from widgets import StartupSplash


class InitialStartup(QRunnable):
    """
    QDialog class that shows the user the process of starting up the application.

    :param GUI gui: The GUI class.
    """
    def __init__(self, gui):
        super().__init__()

        self.gui = gui
        self.startup_splash = StartupSplash(gui, 6)

    def run(self):
        self.gui.update_startup_splash_text.emit('Getting System Info')
        self.gui.main.get_system_info()

        self.gui.update_startup_splash_text.emit('Getting User Settings')
        self.gui.main.get_user_settings()

        self.gui.update_startup_splash_text.emit('Loading Dictionaries')
        self.gui.main.spell_check_thread_pool = QThreadPool()
        self.gui.main.spell_check_thread_pool.setStackSize(256000000)
        self.gui.main.load_dictionary_thread_pool = QThreadPool()
        ld = LoadDictionary(self.gui.main)
        self.gui.main.load_dictionary_thread_pool.start(ld)
        self.gui.main.load_dictionary_thread_pool.waitForDone()

        self.gui.update_startup_splash_text.emit('Getting Indices')
        self.gui.main.get_ids()
        self.gui.main.get_date_list()
        self.gui.main.get_scripture_list()
        self.gui.main.backup_db()

        self.gui.update_startup_splash_text.emit(' Finishing Up ')
        self.gui.create_main_gui.emit()

        self.gui.changes = False
        self.gui.startup_splash_end.emit()


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
