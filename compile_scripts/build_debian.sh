#!/bin/bash
PROGRAMNAME=sermon-prep-database
FRIENDLYNAME="Sermon Prep Database"
VERSION=5.1.2
ICON=resources/icon.png
VENV=.linux_venv
CONTROLSECTION=education
DEPENDENCIES="python3, libxcb-cursor0, libwayland-cursor0"
DESCRIPTION="Sermon Prep Database is a program to organize - and store for easy retrieval - your thoughts and research when preparing a sermon."
MENUSECTION=Applications/Office
TYPE=Application
CATEGORIES="Office"

if [[ $# -gt 0 ]] && [[ $1 = "-h" || $1 = "--help" ]]; then
    echo Use -k or --keep-build-dir to keep the build directory after completion. Otherwise use no arguments.
    exit
fi

if test -d $PROGRAMNAME.$VERSION; then
    echo Deleting Old Build Directory
    rm -r $PROGRAMNAME.$VERSION
fi

echo Creating Directories
mkdir -p $PROGRAMNAME.$VERSION/DEBIAN
mkdir -p $PROGRAMNAME.$VERSION/usr/bin
mkdir -p $PROGRAMNAME.$VERSION/usr/local/$PROGRAMNAME
mkdir -p $PROGRAMNAME.$VERSION/usr/share/applications

echo Creating Control File
cat > $PROGRAMNAME.$VERSION/DEBIAN/control <<EOF
Package: $PROGRAMNAME
Version: $VERSION
Section: $CONTROLSECTION
Priority: optional
Architecture: all
Replaces: $PROGRAMNAME (<< $VERSION)
Depends: $DEPENDENCIES
Maintainer: Jeremy Wilson pastorjeremywilson@gmail.com
Homepage: https://pastorjeremywilson.github.io
Description: $DESCRIPTION

EOF

echo Creating Menu File
cat > $PROGRAMNAME.$VERSION/DEBIAN/$PROGRAMNAME.menu <<EOF
Package($PROGRAMNAME): \
    Section="$MENUSECTION" \
    Title="$FRIENDLYNAME" \
    Command="/usr/bin/$PROGRAMNAME" \
    Icon="$ICON"
EOF

echo Creating Binary
cat > $PROGRAMNAME.$VERSION/usr/bin/$PROGRAMNAME <<EOF
cd /usr/local/$PROGRAMNAME
./$VENV/bin/python3 main.py
EOF
chmod +x $PROGRAMNAME.$VERSION/usr/bin/$PROGRAMNAME

echo Copying Program Data
cp ../*.py $PROGRAMNAME.$VERSION/usr/local/$PROGRAMNAME
cp ../README.* $PROGRAMNAME.$VERSION/usr/local/$PROGRAMNAME
cp -r ../resources $PROGRAMNAME.$VERSION/usr/local/$PROGRAMNAME
cp -r ../$VENV $PROGRAMNAME.$VERSION/usr/local/$PROGRAMNAME

echo Creating Desktop File
cat > "$PROGRAMNAME.$VERSION/usr/share/applications/$FRIENDLYNAME.desktop" <<EOF
[Desktop Entry]
Version=$VERSION
Exec=/usr/bin/$PROGRAMNAME
Comment=$DESCRIPTION
Terminal=false
PrefersNonDefaultGPU=false
Icon=/usr/local/$PROGRAMNAME/$ICON
Type=$TYPE
Name[en_US]=$FRIENDLYNAME
Categories=$CATEGORIES
EOF

echo Building .deb File
dpkg-deb --build $PROGRAMNAME.$VERSION

if [[ $# -eq 0 ]] || [[ $1 != "-k" && $1 != "--keep-build-dir" ]]; then
    echo Deleting Build Directory
    rm -r $PROGRAMNAME.$VERSION
fi
