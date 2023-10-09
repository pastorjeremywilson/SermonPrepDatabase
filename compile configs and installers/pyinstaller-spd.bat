@echo off
echo ====================
echo Options:
echo     noconfirm
echo     onedir
echo     windowed
echo     clean
echo Icon:
echo     C:/Users/pasto/Nextcloud/Documents/Python Workspace/Sermon Prep Database/resources/icons.ico
echo Add Data:
echo     C:/Users/pasto/Nextcloud/Documents/Python Workspace/Sermon Prep Database/gpl-3.0.rtf to ./
echo     C:/Users/pasto/Nextcloud/Documents/Python Workspace/Sermon Prep Database/src/ghostscript to ./ghostscript/
echo     C:/Users/pasto/Nextcloud/Documents/Python Workspace/Sermon Prep Database/resources to ./resources/ 
echo     C:/Users/pasto/Nextcloud/Documents/Python Workspace/Sermon Prep Database/src to ./src/
echo Distribution Path:
echo     C:\Users\pasto\Nextcloud\Documents\Python Workspace\Sermon Prep Database\compile configs and installers\output
echo Entry Point:
echo     C:/Users/pasto/Nextcloud/Documents/Python Workspace/Sermon Prep Database/src/SermonPrepDatabase.py
echo ====================
pyinstaller --noconfirm --onedir --windowed --clean --icon "C:/Users/pasto/Nextcloud/Documents/Python Workspace/Sermon Prep Database/resources/icons.ico" --add-data "C:/Users/pasto/Nextcloud/Documents/Python Workspace/Sermon Prep Database/gpl-3.0.rtf;." --add-data "C:/Users/pasto/Nextcloud/Documents/Python Workspace/Sermon Prep Database/src/ghostscript;ghostscript/" --add-data "C:/Users/pasto/Nextcloud/Documents/Python Workspace/Sermon Prep Database/resources;resources/" --add-data "C:/Users/pasto/Nextcloud/Documents/Python Workspace/Sermon Prep Database/src;src/" --distpath "C:\Users\pasto\Nextcloud\Documents\Python Workspace\Sermon Prep Database\compile configs and installers\output" "C:/Users/pasto/Nextcloud/Documents/Python Workspace/Sermon Prep Database/src/sermon_prep_database.py"