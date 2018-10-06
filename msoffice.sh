#!/bin/bash
# Last revision : (2018-06-25)
# Tested : Xubuntu 18.04
# Author : Pmorch
# Script licence : GPLv3
#
# This script is designed for PlayOnLinux and PlayOnMac.

# CHANGELOG
# [Pmorch] (2011-08-26)
#   Update wine version to 3.11 to make it work in Xubuntu 18.04
#   Backported several features from similar MS Office 2010 source code
#   Installed fonts for bullets and smileys
#   Link to winehq issue about riched20
#   Added categories to POL_Shortcut (so the shortcuts get added to
#       .local/share/applications/*.desktop also)
# [Tinou] (2011-08-22 20-00)
#   Update for POL/POM 4
# [SuperPlumus] (2013-06-08 17-31)
#   gettext

[ "$PLAYONLINUX" = "" ] && exit 0
source "$PLAYONLINUX/lib/sources"

AUTHOR=Pmorch
TITLE="Microsoft Office 2007"
PREFIXNAME="Office2007"
WINE_VERSION="3.11"

POL_Debug_Init
POL_SetupWindow_Init

# This doesn't do anything in 4.2.12. See
# https://www.playonlinux.com/en/topic-15931.html
POL_SetupWindow_presentation "$TITLE" "Microsoft" "http://www.microsoft.com/" "$AUTHOR" "$PREFIXNAME"

POL_System_SetArch "x86"

# Preparation of Wine
POL_Wine_SelectPrefix "$PREFIXNAME"
POL_Wine_PrefixCreate "$WINE_VERSION"

cd "$POL_USER_ROOT/tmp"

# Making bulletslook right
POL_Wine_InstallFonts

# Making e.g. smileys look right
# Is this the right way to do it? See
# https://www.playonlinux.com/en/topic-15929.html
function copyExtraFonts() {
    FONTSURL=https://github.com/IamDH4/ttf-wps-fonts/archive/master.zip
    FONTSZIP=ttf-wps-fonts-master.zip
    FONTSDIR=ttf-wps-fonts-master

    if [ "$WINEPREFIX" = "" ] ; then
        POL_SetupWindow_message "How could there not be a WINEPREFIX: '$WINEPREFIX' - exiting" "$TITLE"
        exit 1;
    fi

    FONTSDESTINATION="$WINEPREFIX/drive_c/windows/Fonts"

    if [ ! -d $FONTSDESTINATION ] ; then
        POL_SetupWindow_message "How could there not be a Fonts directory: '$FONTSDESTINATION' - exiting" "$TITLE"
        exit 1;
    fi

    if ! command -v unzip > /dev/null || \
       ! command -v wget > /dev/null ; then
        POL_SetupWindow_message "unzip and/or wget are not installed - not installing extra fonts" "$TITLE"
        return
    fi

    if [ ! -d $PREFIXFONTSDIR ] ; then
        POL_SetupWindow_message "'$PREFIXFONTSDIR' didn't exist - not installing extra fonts" "$TITLE"
        return
    fi

    wget -O $FONTSZIP $FONTSURL
    unzip $FONTSZIP
    for f in symbol.ttf wingding.ttf WEBDINGS.TTF ; do
        echo Copying $f to $FONTSDESTINATION
        cp $FONTSDIR/$f $FONTSDESTINATION
    done

    rm -rf $FONTSZIP $FONTSDIR
}

copyExtraFonts

POL_SetupWindow_cdrom
POL_SetupWindow_check_cdrom "setup.exe"

POL_Wine_WaitBefore "$TITLE"
POL_Wine "$CDROM/setup.exe"

# See http://forum.winehq.org/viewtopic.php?f=8&t=23126&p=95555#p95555
POL_Wine_OverrideDLL native,builtin riched20

# Creation of shortcuts
POL_Shortcut "WINWORD.EXE" "Microsoft Word 2007" "" "" "Office;WordProcessor;"
POL_Shortcut "EXCEL.EXE" "Microsoft Excel 2007" "" "" "Office;Spreadsheet;"
POL_Shortcut "POWERPNT.EXE" "Microsoft Powerpoint 2007" "" "" "Office;Presentation;"

# These don't seem to do anything
# https://www.playonlinux.com/en/topic-15928.html
#
# POL_Extension_Write doc "Microsoft Word 2007"
# POL_Extension_Write docx "Microsoft Word 2007"
# POL_Extension_Write xls "Microsoft Excel 2007"
# POL_Extension_Write xlsx "Microsoft Excel 2007"
# POL_Extension_Write ppt "Microsoft Powerpoint 2007"
# POL_Extension_Write pptx "Microsoft Powerpoint 2007"

# I think this has something to do with the dialog box for entering initials
# only works for Word and Powerpoint, but I'm not sure. Excel doesn't start
# unless Word or Powerpoint has been started first.
POL_SetupWindow_message "MS Office 2017 has been installed. NOTE: start Word or Powerpoint first, not Excel" "$TITLE"

POL_SetupWindow_Close
exit
