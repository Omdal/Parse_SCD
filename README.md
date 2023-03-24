# Licence
Parse_SCD.py is a script used to parse SCD information from Visio files

Copyright (C) 2023  Olav André Omdal

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU Affero General Public License as
published by the Free Software Foundation, either version 3 of the
License, or (at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU Affero General Public License for more details.
You should have received a copy of the GNU Affero General Public License
along with this program.  If not, see <https://www.gnu.org/licenses/>.

# What is an SCD

An SCD is a System Control Diagram drawn according to the standard
IEC-63131. It shows the relation between function blocks in industrial
control systems.

# The goal of this project
The goal of this project is to simplify the next steps in the use of
SCDs. After drawing the SCD in Visio, this script extracts infromation
from the Visio drawing to simplify automatic generation of sourcecode
and documentation.

# What the script does
 
This script extracts all function blocks from a Visio SCD drawing and save
the result in a .csv file with the same name as the Visio SCD drawing.
The CSV file contens:

Sheet | Function-block | Tag | Description
-|-|-|-
The name of the sheet | The type of the function block | The Tag in the function block | The description if used

# How to use

This script can be used in multiple ways:
1. Run it in a folder with one or more Visio SCD drawings. It will generate
   separate CSV files for each SCD drawing in the same folder
2. Drag and drop a Visio SCD drawing on this file if your installation of
   Python allow running Python-scripts as programs
3. Run the script with the filename as an argument:  
    ``$ python3 Parse_SCD.py MySCD``  
        - or -  
    ``$ python3 Parse_SCD.py MySCD.vsdx``  
   Both methods will parse the given SCD drawing.

# NOTE
The visio-file is hard-parsed, newer versions of Microsoft Visio may cause
the script to fail.

# Version history
Date | Version | Description
-|-|-
24.03.2023 | 1.0.0 | Initial revision
