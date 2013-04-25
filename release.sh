#!/bin/sh

# Package all necessary files and copy them along with the version.xml to the Dropbox folder

zip -0 Quizzor.mmip Quizzor.vbs Install.ini Uninstall.ini license.txt 
cp Quizzor.mmip version.xml ~/Dropbox/Public/
