"""
@author Jeremy G. Wilson

Copyright 2023 Jeremy G. Wilson

This file is a part of the Sermon Prep Database program (v.3.4.1)

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

import logging
import re
import xml.etree.ElementTree as ET
from os.path import exists


class GetScripture:
    root = None

    def __init__(self, spd):
        if spd.bible_file and exists(spd.bible_file):
            tree = ET.parse(spd.bible_file)
            self.root = tree.getroot()

        self.books = [
            ['Genesis', 'gen', 'ge', 'gn'],
            ['Exodus', 'exod', 'exo', 'ex'],
            ['Leviticus', 'lev', 'le', 'lv'],
            ['Numbers', 'num', 'nu', 'nm', 'nb'],
            ['Deuteronomy', 'deut', 'de', 'dt'],
            ['Joshua', 'josh', 'jos', 'jsh'],
            ['Judges', 'judg', 'jg', 'jdgs'],
            ['Ruth', 'rth', 'ru'],
            ['1 Samuel', '1st samuel', '1 sa', '1sa', '1s', '1 sm', '1sm', '1st sam'],
            ['2 Samuel', '2nd samuel', '2 sa', '2sa', '2s', '2 sm', '2sm', '2nd sam'],
            ['1 Kings', '1st kings', '1 ki', '1ki', '1k', '1 kgs', '1kgs', '1st ki', '1st kgs'],
            ['2 Kings', '2nd kings', '2 ki', '2ki', '2k', '2 kgs', '2kgs', '2nd ki', '2nd kgs'],
            ['1 Chronicles', '1st chronicles', '1 ch', '1ch', '1 chron', '1chron', '1 chr', '1chr',
             '1st ch', '1st chron'],
            ['2 Chronicles', '2nd chronicles', '2 ch', '2ch', '2 chron', '2chron', '2 chr', '2chr',
             '2nd ch', '2nd chron'],
            ['Ezra', 'ezr', 'ez'],
            ['Nehemiah', 'neh', 'ne'],
            ['Esther', 'est', 'esth', 'es'],
            ['Job', 'jb'],
            ['Psalms', 'psalm', 'ps', 'psa', 'psm', 'pss'],
            ['Proverbs', 'pro', 'pr', 'prv'],
            ['Ecclesiastes', 'eccles', 'eccle', 'ec', 'qoh'],
            ['Song of Solomon', 'song', 'so', 'sos', 'canticle of canticles', 'canticles', 'cant'],
            ['Isaiah', 'isa', 'is'],
            ['Jeremiah', 'jer', 'je', 'jr'],
            ['Lamentations', 'lam', 'la'],
            ['Ezekiel', 'ezek', 'eze', 'ezk'],
            ['Daniel', 'dan', 'da', 'dn'],
            ['Hosea', 'hos', 'ho'],
            ['Joel', 'joe', 'jl'],
            ['Amos', 'am'],
            ['Obadiah', 'obad', 'ob'],
            ['Jonah', 'jnh', 'jon'],
            ['Micah', 'mic', 'mc'],
            ['Nahum', 'nah', 'na'],
            ['Habakkuk', 'hab', 'hb'],
            ['Zephaniah', 'zep', 'zp'],
            ['Haggai', 'hag', 'hg'],
            ['Zechariah', 'zech', 'zec', 'zc'],
            ['Malachi', 'mal', 'ml'],
            ['Matthew', 'matt', 'mat', 'mt'],
            ['Mark', 'mk', 'mar', 'mrk', 'mr'],
            ['Luke', 'luk', 'lk'],
            ['John', 'joh', 'jhn', 'jn'],
            ['Acts', 'act', 'ac'],
            ['Romans', 'rom', 'ro', 'rm'],
            ['1 Corinthians', '1st corinthians', '1 cor', '1cor', '1 co', '1co', '1corinthians', '1st cor', '1st co'],
            ['2 Corinthians', '2nd corinthians', '2 cor', '2cor', '2 co', '2co', '2corinthians', '2nd cor', '2nd co'],
            ['Galatians', 'gal', 'ga'],
            ['Ephesians', 'ephes', 'eph'],
            ['Philippians', 'phil', 'php', 'pp'],
            ['Colossians', 'col', 'co'],
            ['1 Thessalonians', '1st thessalonians', '1 thes', '1thes', '1 th', '1th', '1thessalonians',
             '1st thes', '1st th'],
            ['2 Thessalonians', '2nd thessalonians', '2 thes', '2thes', '2 th', '2th', '2thessalonians',
             '2nd thes', '2nd th'],
            ['1 Timothy', '1st timothy', '1 tim', '1tim', '1 ti', '1ti', '1timothy', '1st tim', '1st ti'],
            ['2 Timothy', '2nd timothy', '2 tim', '2tim', '2 ti', '2ti', '2timothy', '2nd tim', '2nd ti'],
            ['Titus', 'tit', 'ti'],
            ['Philemon', 'philem', 'phm', 'pm'],
            ['Hebrews', 'heb'],
            ['James', 'jas', 'jm'],
            ['1 Peter', '1st peter', '1 pet', '1pet', '1 pe', '1pe', '1 pt', '1pt', '1 p', '1p',
             '1st pet', '1st pe', '1st pt', '1st p'],
            ['2 Peter', '2nd peter', '2 pet', '2pet', '2 pe', '2pe', '2 pt', '2pt', '2 p', '2p',
             '2nd pet', '2nd pe', '2nd pt', '2nd p'],
            ['1 John', '1st john', '1 jn', '1jn', '1 jo', '1jo', '1 joh', '1joh', '1 jhn', '1jhn', '1 j', '1j',
             '1st jn', '1st jo', '1st joh', '1st jhn'],
            ['2 John', '2nd john', '2 jn', '2jn', '2 jo', '2jo', '2 joh', '2joh', '2 jhn', '2jhn', '2 j', '2j',
             '2nd jn', '2nd jo', '2nd joh', '2nd jhn'],
            ['3 John', '3rd john', '3 jn', '3jn', '3 jo', '3jo', '3 joh', '3joh', '3 jhn', '3jhn', '3 j', '3j',
             '3rd jn', '3rd jo', '3rd joh', '3rd jhn'],
            ['Jude', 'jud', 'jd'],
            ['Revelation', 'rev', 're', 'the revelation']
        ]

    def get_passage(self, reference):
        try:
            if self.root:
                reference_split = reference.split(' ')
                reference_ok = False
                if len(reference_split) > 1:
                    if any(a.isnumeric() for a in reference_split[0]) and all(b.isalpha() for b in reference_split[1]):
                        book = reference_split[0] + ' ' + reference_split[1]
                        passage_split = reference_split[2].split(':')
                    else:
                        if len(reference_split) > 2:
                            book = reference_split[0]
                            if not any(c.isnumeric for c in reference_split[1]):
                                book += ' ' + reference_split[1]
                                passage_split = reference_split[2].split(':')
                                if not any(d.isnumeric for d in reference_split[2]):
                                    book += ' ' + reference_split[2]
                                    passage_split = reference_split[3].split(':')
                        else:
                            book = reference_split[0]
                            passage_split = reference_split[1].split(':')

                    if len(passage_split) > 1:
                        chapter = passage_split[0]
                        verse_split = []
                        if '-' in passage_split[1]:
                            verse_split = passage_split[1].split('-')
                        elif '–' in passage_split[1]:
                            verse_split = passage_split[1].split('–')
                        if len(verse_split) > 1:
                            start_verse = verse_split[0]
                            end_verse = verse_split[1]
                            reference_ok = True
                        elif all(d.isnumeric for d in str(passage_split[1])):
                            start_verse = passage_split[1]
                            end_verse = passage_split[1]
                            reference_ok = True
                else:
                    return -1

                if reference_ok:
                    book = book.replace('.', '')
                    book = book.lower()
                    chapter = chapter.lower()
                    start_verse = start_verse.lower()
                    end_verse = end_verse.lower()

                    standard_book = None
                    book_number = None
                    for i in range(len(self.books)):
                        for j in range(len(self.books[i])):
                            if book == self.books[i][j].lower():
                                standard_book = self.books[i][0]
                                book_number = i + 1

                    if standard_book:
                        book_element = None
                        for child in self.root:
                            if child.get('bname'):
                                if child.get('bname') == standard_book:
                                    book_element = child
                            if not book_element:
                                if child.get('bnumber'):
                                    if child.get('bnumber') == str(book_number):
                                        book_element = child

                        if book_element:
                            for child in book_element:
                                if child.get('cnumber') == str(chapter):
                                    chapter_element = child

                            scripture_text = ""
                            for child in chapter_element:
                                try:
                                    if int(child.get('vnumber')) >= int(start_verse) and int(child.get('vnumber')) <= int(end_verse):
                                        scripture_text += child.get('vnumber') + ' ' + child.text + ' '
                                except ValueError:
                                    pass

                            scripture_text = re.sub('\s+', ' ', scripture_text).strip()

                            if len(scripture_text) > 0:
                                return scripture_text
                            else:
                                return -1
                    else:
                        return -1

                else:
                    return -1
        except Exception as ex:
            logging.exception(str(ex), True)
