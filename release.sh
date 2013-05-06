#!/bin/sh

# Package all necessary files and copy them along with the version.xml to the Dropbox folder

zip -0 Quizzor.mmip Quizzor.vbs Install.ini Uninstall.ini license.txt 
cp Quizzor.mmip version.xml ~/Dropbox/Public/Quizzor/

# Create and copy german packages
./create_translated_packages.py -po -p Install_de.ini license.txt Uninstall.ini 
cp Quizzor_de.mmip ~/Dropbox/Public/Quizzor/de/
cp version_de.xml ~/Dropbox/Public/Quizzor/de/version.xml
