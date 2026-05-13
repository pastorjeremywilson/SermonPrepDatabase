#!/bin/bash
../.linux_venv/bin/pyinstaller --noconfirm --clean --windowed \
-i "../resources/svg/spIcon.svg" \
--add-data=../resources:resources/ \
--add-data=../README.html:. \
--add-data=../README.md:. \
--distpath=~/Desktop/output/dist/SermonPrepDatabase \
--workpath=~/Desktop/output/work/SermonPrepDatabase \
--name=sermon-prep-database \
../main.py
