#!/usr/bin/env python3

import sys
import os
import polib
import tempfile
import zipfile
import re
import argparse
from itertools import product
from pprint import pprint

def create_translation_dict(mopolib_object_list):
    """
    polib_object_list a list of either be a MOFile() or POFile()
    The list is processd in reverse order, so the first item has
    the highest priority, meaning overwriting any item with the same key.
    """
    result_dict = {}
    for mopolib_object in mopolib_object_list[::-1]:
        for entry in mopolib_object:
            result_dict[entry.msgid] = entry.msgstr
    return result_dict

def load_all_arguments():
    arg_parser = argparse.ArgumentParser(description=\
    '''
    Replace localized strings of a .vbs source file with localizations from
    a .po or .mo file.
    The .po or mo file has to be named like the project name followed by .po 
    or mo respectively.
    ''')
    arg_parser.add_argument('locales_dir', nargs='?', default='localization',
            help='Directory with language subdirectories, default=localization')
    arg_parser.add_argument('-t', '--test-mode', action='store_true',
            help='Enter test mode: write *_<locale>.vbs files instead\
                  of a .mmip package')
    arg_parser.add_argument('-e', '--encoding', default='latin-1',
            help='Defines the codec used for output and input files.\
                 default=latin-1')
    arg_parser.add_argument('-p', dest='package_files', nargs='+',
            default=['Install.ini', 'Uninstall.ini', 'license.txt'],
            help='A list of files, to be packaged in the .mmip file,\
                  default=Install.ini, Uninstall.ini and license.txt\n\
                  Ignored when in test mode.')
    arg_parser.add_argument('-s', dest='source_files', nargs='+',
            help='Define which source files shall be translated.\
                  default=project_name followed by .vbs, see -p')
    arg_parser.add_argument('-pn', dest='project_name', 
            default=os.path.split(os.getcwd())[1],
            help='Define the projects name. This is used for the .mmip\
                  filename default=last dirname of current path')
    mopo_group = arg_parser.add_mutually_exclusive_group()
    mopo_group.add_argument('-po', dest='locale_filetype',
            action='store_const', const='po', default='mo',
            help="Use .po files for translation.")
    mopo_group.add_argument('-mo', dest='locale_filetype', 
            action='store_const', const='mo', default='mo',
            help="Use .mo files for translation.")

    # Ensure that at least one source_file is provided
    arguments = arg_parser.parse_args()
    if not arguments.source_files:
        arguments.source_files = [arguments.project_name+'.vbs']
    globals().update(vars(arguments))

def delete_locale_indicator(filename, locale):
    '''Removes _<locale> from a given filename '''
    return filename.replace('_'+locale, '')

load_all_arguments()

# Define localization strings
locale_re = re.compile(r'''
    Localize(?:dFormat)?    # Methods begin with Localize or LocalizedFormat
    \s*\(\s*"               # Possible whitespaces after method name
    ((?:[^"]|"")*)          # Match anything but ", though do allow ""
    "                       # The search is over when a single " is found
''',re.VERBOSE|re.IGNORECASE)

for language in os.listdir(locales_dir):
    # Load localizations
    print('Processing language %s' % language)
    mopo_objects = []
    lang_dir = os.path.join(locales_dir,language)+os.path.sep
    if locale_filetype == 'mo':
        mopo_objects = [ polib.mofile(lang_dir+file) for file in \
                [f for f in os.listdir(lang_dir) if f[-3:]=='.mo']]
    elif locale_filetype == 'po':
        mopo_objects = [ polib.pofile(lang_dir+file) for file in \
                [f for f in os.listdir(lang_dir) if f[-3:]=='.po'] if file]
    # Move the project's locale file to the front, 
    # so it has the highest priority
    mopo_file = project_name+'.'+locale_filetype 
    if mopo_file in mopo_objects: 
        mopo_objects = mopo_objects.insert(0, 
                            mopo_objects.pop(mopo_objects.index(mopo_file)))
    
    print('Building translation dictionary')
    translations = create_translation_dict(mopo_objects)

    for source_file_name in source_files:
        print('Processing source file %s' % source_file_name)
        try:
            basename = os.path.splitext(os.path.basename(source_file_name))[0]
            with open(source_file_name, encoding=encoding) as source_file: 
                # Create a permanent file in test mode.
                temp_file = (open(basename+'_'+language+'.vbs', 
                                 encoding=encoding, mode='wt')
                             if test_mode
                             else tempfile.NamedTemporaryFile(prefix=basename,
                                suffix='.vbs', encoding=encoding, mode='wt'))

                # Process all lines
                for line in source_file:
                    for src_string in locale_re.findall(line):
                        if src_string in translations:
                            line = line.replace('"'+src_string+'"', \
                                    '"'+translations[src_string]+'"')
                        else:
                            print('NOT TRANSLATED: ' + src_string)
                    temp_file.write(line)
                
                # Write out to package
                temp_file.flush()
                if not test_mode:
                # if in test_mode the file has already been created
                    with zipfile.ZipFile(project_name+'_'+language+'.mmip', 
                                         mode='w') as package_file:
                        for filename in package_files:
                            package_file.write(filename, arcname=
                            delete_locale_indicator(filename, language)) 
                        package_file.write(temp_file.name, 
                                           arcname=source_file_name)

        except IOError as err:
            print("Error processing source file.", err)
            pass




