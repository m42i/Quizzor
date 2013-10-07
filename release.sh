#!/bin/sh

# Package all necessary files and copy them along with the version.xml to the Dropbox folder
version=$1

zip -0 Quizzor.mmip Quizzor.vbs Install.ini Uninstall.ini license.txt 
cp Quizzor.mmip ~/Dropbox/Public/Quizzor/Quizzorv${version}.mmip
cp version.xml ~/Dropbox/Public/Quizzor/version.xml

# Create and copy german packages
./create_translated_packages.py -po -p Install_de.ini license.txt Uninstall.ini 
cp Quizzor_de.mmip ~/Dropbox/Public/Quizzor/de/Quizzorv${version}_de.mmip
cp version_de.xml ~/Dropbox/Public/Quizzor/de/version.xml
