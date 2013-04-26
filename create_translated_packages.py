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

def create_translation_dict(polib_object):
  """
  polib_object can either be a MOFile() or POFile()
  """
  result_dict = {}
  for entry in polib_object:
    result_dict[entry.msgid] = entry.msgstr
  return result_dict

def load_all_arguments():
  arg_parser = argparse.ArgumentParser(description=\
  '''
  Replace localized strings of a .vbs source file with localizations from
  a .po or .mo file.
  The .po or mo file has to be named like the project name followed by .po or mo respectively.
  ''')
  arg_parser.add_argument('locales_dir', nargs='?', default='localization',\
      help='Directory with language subdirectories, default=localization')
  arg_parser.add_argument('-t', '--test-mode', action='store_true',\
      help='Enter test mode: write *_<locale>.vbs files instead\
        of a .mmip package')
  arg_parser.add_argument('-e', '--encoding', default='utf-8',\
      help='Defines the codec used for output and input files. \
        default=utf-8')
  arg_parser.add_argument('-p', dest='package_files', nargs='+',\
      default=['Install.ini', 'Uninstall.ini', 'license.txt'],
      help='A list of files, to be packaged in the .mmip file,\
        default=Install.ini, Uninstall.ini and license.txt\n\
        Ignored when in test mode.')
  arg_parser.add_argument('-s', dest='source_files', nargs='+',\
      help='Define which source files shall be translated.\
        default=project_name followed by .vbs, see -p')
  arg_parser.add_argument('-pn', dest='project_name', default=os.path.split(os.getcwd())[1],\
      help='Define the projects name. This is used for the .mmip filename\
        default=last dirname of current path')
  mopo_group = arg_parser.add_mutually_exclusive_group()
  mopo_group.add_argument('-po', dest='locale_filetype', action='store_const', const='po', default='mo',
      help="Use .po files for translation.")
  mopo_group.add_argument('-mo', dest='locale_filetype', action='store_const', const='mo', default='mo',
      help="Use .mo files for translation.")

  # Ensure that at least one source_file is provided
  arguments = arg_parser.parse_args()
  if not arguments.source_files:
    arguments.source_files = [arguments.project_name+'.vbs']
  globals().update(vars(arguments))

load_all_arguments()

# Define localization strings
locale_re = re.compile(\
    r'^[^\']*Localize(?:dFormat)?\s*\(\s*"((?:[^"]|"")*)"[^)]*\)')

for source_file_name, language in product(source_files, os.listdir(locales_dir)):
  try:
    with open(source_file_name, encoding=encoding) as source_file: 
      # Load localizations
      basename = os.path.splitext(os.path.basename(source_file_name))[0]
      ext = locale_filetype
      translations = create_translation_dict(\
          polib.mofile(os.path.join(locales_dir,language,basename+'.'+ext)) if locale_filetype == 'mo'\
          else polib.pofile(os.path.join(locales_dir,language,basename+'.'+ext)) )

      # Create a permanent file in test mode.
      temp_file = open(basename+'_'+language+'.vbs', encoding=encoding, mode='wt')\
          if test_mode else tempfile.NamedTemporaryFile(prefix=basename,\
            suffix='.vbs', encoding=encoding, mode='wt')

      # Process all lines
      for line in source_file:
        for src_string in locale_re.findall(line):
          if src_string in translations:
            line = line.replace('"'+src_string+'"', \
                '"'+translations[src_string]+'"')
          else:
            print('NOT TRANSLATED: ' + src_string)
          temp_file.write(line)
      
      # Write out to package, if in test_mode the file has already been created
      temp_file.flush()
      if not test_mode:
        with zipfile.ZipFile(project_name+'_'+language+'.mmip', mode='w') \
            as package_file:
          # First add additional files, if duplicates are present, the localized version is provided
          for filename in package_files:
            package_file.write(filename) 
          package_file.write(temp_file.name, arcname=source_file_name)

  except IOError as err:
    print("Error processing source file.", err)
    pass




