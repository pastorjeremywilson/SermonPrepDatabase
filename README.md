[![Status](https://img.shields.io/badge/status-active-brightgreen.svg?style=plastic)](https://github.com/pastorjeremywilson/SermonPrepDatabase/pulse/monthly)
[![License](https://img.shields.io/badge/license-GPL-blue.svg?style=plastic)](https://www.gnu.org/licenses/gpl-3.0.en.html)

<img src='https://github.com/pastorjeremywilson/public/blob/main/spBanner.svg?raw=true' width='600px' />

# What Am I?

Sermon Prep Database is a program to organize - and store for easy retrieval - your thoughts and research when preparing a sermon.

# What's New in v.5.0.1?
Minor bug fix - Fixed crash when right-clicking inside an empty text edit area.

# Why Sermon Prep Database?

When preparing a sermon for your congregation, you probably find yourself in the habit of performing certain tasks: doing research from commentaries, jotting down outlines, following an exegesis plan, etc. This program gives you areas to record all of these things and then saves them into a database where you can retrieve them any time.

# Installation

Currently, Sermon Prep Database is available for the Microsoft Windows operating
system only. Download the current SPD installer (v.5.0.1) and run
it on your computer.

# Using Sermon Prep Database

*This information will get you started, but for more in-depth information*
*on using Sermon Prep Database, please click “Help” and “Help Topics” in the menu.*

### Screen Layout

When you first run Sermon Prep Database, you’ll be greeted with a screen consisting of tabs:

- Scripture

- Exegesis

- Outlines

- Research

- Sermon

In the Scripture tab, you can record the week's pericope and its readings, or the particular topic you're discussing and the readings to go along with it. You can also record the reference of the text you'll be preaching on, as well as the text itself.

In the Exegsis tab, you can record what you have discerned to be the law and gospel components of the text, it's "big picture" idea, the purpose of your sermon, and the way those things will be brought to your congregation.

The Outline tab is where you can store outlines for the scripture, our sermon, and any illustration(s) you may want to use.

The Research tab is where you can write down any important points you've discovered during your research.

Finally, the Sermon tab is where you can store the text of the sermon itself, as well as any important details regarding the sermon, such as the Date, Title, and Hymn of Response.

### Settings

There are various changes that can be made to how Sermon Prep Database works. By
choosing the "Edit" menu, and then "Configure," you will find that you can change such things as:

- The Program's Colors

- The Font being used

- The Line Spacing of the text boxes

- The Labels used in the program

- What words are stored in your spell-check dictionary

- Whether Spell Check is enabled or not

You might be particularly interested in changing the labels used in the program. For example, in the Exegesis tab, three of the text boxes are named, "Fallen Condition Focus of the Text," "Gospel Answer of the Text," and "Central Proposition of the Text," but you might prefer the wording, "Sin Problem," "Gospel Remedy," and "Big Idea." You can make those changes by opening the Edit menu, clicking, "Configure," and then choosing, "Rename Labels."

When you first use the program, you might decide that it would be convenient to import your old sermons from files you already have saved on your computer. In the "File" menu, select "Import Sermons From Files". You will be prompted to select a folder that contains your sermons. File formats the program can import from are Microsoft Word (.docx), Open Document Format (.odt), and Plain Text (.txt). It would be beneficial to do a little leg-work in advance of importing so that the program can automatically fill in some of the details of your sermon while it imports. To accomplish this, make sure your files are named in the following way:

`YYYY-MM-DD.book.chapter.verse-verse`

For example, a sermon preached on May 20th, 2011 on Mark 3:1-12, saved as a Microsoft word document, would be named:

`2011-05-11.mark.3.1-12.docx`

The import will work fine if you haven't named your files in this way, only the program will not be able to fill in such areas as the sermon date and text reference.

You also have the ability to import an **XML Bible**. Importing a bible will allow you to have the program auto-fill the Sermon Text area of the scripture tab simply by typing or copying a reference into the Sermon Text Reference box.

### Shortcut Keys

There are a few Shortcut Keys that can be used when using the program:

<table>
<thead>
	<tr>
		<th>Key</th>
		<th>Function</th>
		<th>Description</th>
	</tr>
</thead>
<tbody>
	<tr>
		<td>Ctrl-S</td>
		<td>Save</td>
		<td>Save your work</td>
	</tr>
	<tr>
		<td>Ctrl-P</td>
		<td>Print</td>
		<td>Creates a printout of this sermon's prep work</td>
	</tr>
	<tr>
		<td>Ctrl-B</td>
		<td>Bold</td>
		<td>Set the text to Bold (in text boxes only)</td>
	</tr>
	<tr>
		<td>Ctrl-I</td>
		<td>Italic</td>
		<td>Set the text to Italic (in text boxes only)</td>
	</tr>
	<tr>
		<td>Ctrl-U</td>
		<td>Underline</td>
		<td>Set the text to Underline (in text boxes only)</td>
	</tr>
	<tr>
		<td>Ctrl-Shift-B</td>
		<td>Bullets</td>
		<td>Create bullet points</td>
	</tr>
</tbody>
</table>

# Known Issues

# Technologies and Credits

<div>
    <img src='https://github.com/pastorjeremywilson/public/blob/main/python-logo-only.svg?raw=true' width=40px align='right' />
    Sermon Prep Database is written primarily in <a href="https://www.python.org" target="_blank">Python</a>, compiled through <a href="https://www.pyinstaller.org" target="_blank">PyInstaller</a>,
    and packaged into an installation executable with <a href="https://jrsoftware.org/isinfo.php" target="_blank">Inno Setup Compiler</a>.
</div>
<br>

<div>
    <img src='https://github.com/pastorjeremywilson/public/blob/main/sqlite370.jpg?raw=true' height=40px align='right' />
    Sermon Prep Database’s main database is <a href="https://www.sqlite.org" target="_blank">SQLite</a> and the remaining data files are
    stored in the <a href="https://www.json.org/json-en.html" target="_blank">JSON</a> format.
</div>
<br>

<div>
    <img src='https://github.com/pastorjeremywilson/public/blob/main/Qt-logo-neon.png?raw=true' height=40px align='right' />
    Sermon Prep Database uses <a href="https://www.qt.io/product/framework" target="_blank">Qt</a> (PyQt6) for the user interface.
</div>
<br>

<div>
    <img src='https://github.com/pastorjeremywilson/public/blob/main/ghostscript-plus-text.png?raw=true' height=40px align='right' />
    Printing made possible with the help of Artifex Software's <a href="https://www.ghostscript.com" target="_blank">Ghostscript</a>.
</div>
<br>
All trademarks (c) their respective owners.

# Licensing

<img src='https://github.com/pastorjeremywilson/public/blob/main/gnu-4.svg?raw=true' height=120px align='left' />
ProjectOn is licensed under the GNU General Public License (GNU GPL)
published by the Free Software Foundation, either version 3 of the
License, or (at your option) any later version. See
<a href="http://www.gnu.org/licenses/" target="_blank">http://www.gnu.org/licenses/</a>.