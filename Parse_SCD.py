#!/usr/bin/env python

###############################################################################
#
#    Parse_SCD.py is a script used to parse SCD information from Visio files
#    Copyright (C) 2023  Olav André Omdal
#
#    This program is free software: you can redistribute it and/or modify
#    it under the terms of the GNU Affero General Public License as
#    published by the Free Software Foundation, either version 3 of the
#    License, or (at your option) any later version.
#
#    This program is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU Affero General Public License for more details.
#
#    You should have received a copy of the GNU Affero General Public License
#    along with this program.  If not, see <https://www.gnu.org/licenses/>.
#
###############################################################################

###############################################################################
#
#   What it does:
#   -------------
#    
#   This script extracts all function blocks from a Visio SCD drawing and save
#   the result in a .csv file with the same name as the Visio SCD drawing.
#
#   The CSV file contens:
#   .--------------.-----------------.----------------.-----------------.
#   | Sheet        | Function-block  | Tag            | Description     |
#   |--------------|-----------------|----------------|-----------------|
#   | The name of  | The type of the | The Tag in the | The description |
#   | of the sheet | function block  | function block | field if used   |
#   '--------------'-----------------'----------------'-----------------'
#   
#   How to use:
#   -----------
#
#   This script can be used in multiple ways:
#
#   1. Run it in a folder with one or more Visio SCD drawings. It will generate
#      separate CSV files for each SCD drawing in the same folder
# 
#   2. Drag and drop a Visio SCD drawing on this file if your installation of
#      Python allow running Python-scripts as programs
#
#   3. Run the script with the filename as an argument:
#        python3 Parse_SCD.py MySCD
#               - or -
#        python3 Parse_SCD.py MySCD.vsdx
#      Both methods will parse the given SCD drawing.
#
#   NOTE:
#   ----
#   The visio-file is hard-parsed, newer versions of Microsoft Visio may cause
#   the script to fail.
#
###############################################################################

__author__ = "Olav André Omdal"
__copyright__ = "Copyright 2023, Olav André Omdal"
__credits__ = ["Olav André Omdal"]
__license__ = "agpl-3.0"
__version__ = "1.0.0"
__maintainer__ = "Olav André Omdal"
__email__ = "oaomdal@gmail.com"
__status__ = "Production"

import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

namespaces = {
    '' : 'http://schemas.microsoft.com/office/visio/2012/main',
    'r' : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}
filer = []


def parseFile(file: str):
    FB = []
    sheets = {}
    zf = zipfile.ZipFile(file, 'r')
    liste = 'sep=,\nSheet,Function-block,Tag,Description\n'

    # Finn alle versjoner av "Function block"
    tree = ET.parse(zf.open('visio/masters/masters.xml'))
    root = tree.getroot()
    for master in root.findall('.//Master[@Name]',namespaces):
        if 'Function block' in master.get('Name'):
            FB.append(master.get('ID'))

    # Finn sidenavn
    tree = ET.parse(zf.open('visio/pages/pages.xml'))
    root = tree.getroot()
    for sheet in root.findall('.//Page[@Name]',namespaces):
        PageName = sheet.get('Name')
        Rel = sheet.find('.//Rel',namespaces)
        filename = 'page' + re.findall(r'\d+',Rel.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'))[0] + '.xml'
        sheets['visio/pages/'+filename] = PageName

    # Søk deretter igjennom alle sidene for å finne funksjonsblokkene
    for name in zf.namelist():
        if 'page' in name:
            f = zf.open(name)

            tree = ET.parse(f)
            root = tree.getroot()

            for figur in root.findall(".//Shape[@Master]",namespaces):
                if figur.get('Master') in FB:
                    for section in figur.findall("Section[@N='Property']",namespaces):
                        blockType = ''
                        blockName = ''
                        blockDescription = ''
                        for row in section.findall("Row",namespaces):
                            if 'FB' in row.get('N'):
                                blockType = row.find('Cell',namespaces).get('V')
                            if 'Tag' in row.get('N'):
                                blockName = row.find('Cell',namespaces).get('V')
                            if 'Info' in row.get('N'):
                                blockDescription = row.find('Cell',namespaces).get('V')
                        liste += sheets[name] + ',' + blockType + ',' + blockName + ',' + blockDescription + '\n'
    csvfilename = file.replace('.vsdx','.csv')
    with open(csvfilename,'w') as csvfile:
        csvfile.write(liste)
    return True

# Skriv inn filnavn som argument
if len(sys.argv) == 2:
    #Se om fila finnes
    filename = sys.argv[1]
    if '.vsdx' not in filename:
        filename = filename + '.vsdx'
    if os.path.isfile(filename):
        print('Parsing ' + filename)
        parseFile(filename)
        exit()
    else:
        print(sys.argv[1] + ' is not a file.')
        print('Please type in the filename, or run this script without arguments')

if len(sys.argv) > 2:
    print('You have entered too many arguments.')
    print('You entered the following arguments: ' + ', '.join(sys.argv[1:]))

for fil in os.listdir():
    if '.vsdx' in fil:
        filer.append(fil)

for fil in filer:
    parseFile(fil)