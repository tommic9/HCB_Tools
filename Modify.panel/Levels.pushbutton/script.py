# -*- coding: utf-8 -*-

__title__ = "Add Levels Elevation"
__doc__   = """This tool will add/update your level name to have its elevation"

How-to:
- Click on the button
- Change settings (optional)
- Rename Levels

Author: Tomasz Michalek"""

__min_revit_ver = 2019
__max_revit_ver = 2024
# __context__ = ('Walls','Floors', 'Roofs') # Make your button available only whn certain categories are selected


#  ___ __  __ ____   ___  ____ _____ ____
# |_ _|  \/  |  _ \ / _ \|  _ \_   _/ ___|
#  | || |\/| | |_) | | | | |_) || | \___ \
#  | || |  | |  __/| |_| |  _ < | |  ___) |
# |___|_|  |_|_|    \___/|_| \_\|_| |____/  IMPORTS

# Regular + Autodesk
import os, sys, math, datetime, time
from Autodesk.Revit.DB import * # Import everything from DB
from Autodesk.Revit.DB import  Transaction, Element, ElementId, FilteredElementCollector

#pyRevit
from pyrevit import revit, forms

# Custom imports
from Snippets._selection_ import get_selected_elements

# .NET Imports
import clr

from System.Collections.Generic import List # List <ElementType() <- it's special type of list that RevitAPI oftn requires.

# Custom imports
from Snippets._convert import convert_internal_to_m
# .NET Imports
import clr
clr.AddReference("System")

#  __     ___    ____  ___    _    ____  _     _____ ____
# \ \   / / \  |  _ \|_ _|  / \  | __ )| |   | ____/ ___|
#  \ \ / / _ \ | |_) || |  / _ \ |  _ \| |   |  _| \___ \
#   \ V / ___ \|  _ < | | / ___ \| |_) | |___| |___ ___) |
#    \_/_/   \_\_| \_\___/_/   \_\____/|_____|_____|____/  VARIABLES

doc     = __revit__.ActiveUIDocument.Document
uidoc   = __revit__.ActiveUIDocument
app     = __revit__.Application
PATH_SCRIPT = os.path.dirname(__file__)

# Symbols
symbol_start = "⌞"
symbol_end   = "⌝"

# from pyrevit.revit import uidoc, doc, app # Alternative

#  _____ _   _ _   _  ____ _____ ___ ___  _   _ ____
# |  ___| | | | \ | |/ ___|_   _|_ _/ _ \| \ | / ___|
# | |_  | | | |  \| | |     | |  | | | | |  \| \___ \
# |  _| | |_| | |\  | |___  | |  | | |_| | |\  |___) |
# |_|    \___/|_| \_|\____| |_| |___\___/|_| \_|____/
#                                                      FUNCTIONS



#   ____ _        _    ____ ____  _____ ____
#  / ___| |      / \  / ___/ ___|| ____/ ___|
# | |   | |     / _ \ \___ \___ \|  _| \___ \
# | |___| |___ / ___ \ ___) |__) | |___ ___) |
#  \____|_____/_/   \_\____/____/|_____|____/
#                                              CLASSES


#  __  __    _    ___ _   _
# |  \/  |  / \  |_ _| \ | |
# | |\/| | / _ \  | ||  \| |
# | |  | |/ ___ \ | || |\  |
# |_|  |_/_/   \_\___|_| \_|
#                            MAIN

# Get all Levels
all_levels = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()

# Get Levels Elevation
autoTransaction = False
if not doc.IsModifiable:
     autoTransaction = True

t = Transaction(doc,__title__)
t.Start()

for lvl in all_levels:
    lvl_elevation = lvl.Elevation
    # Convert to meters + rounding
    lvl_elevation_m = round(convert_internal_to_m(lvl_elevation),2)
    lvl_elevation_m_str = "+" + "{:.2f}".format(lvl_elevation_m) if lvl.Elevation > 0 else str(lvl_elevation_m)

    elevation_value = symbol_start + lvl_elevation_m_str + symbol_end
    # Check if elevation already exists
    if symbol_start in lvl.Name:
        # ELEVATIONS EXIST (Update)
        name_parts = lvl.Name.split(symbol_start)
        new_name   = name_parts[0] + " " + elevation_value
    else:
    # ELVATION DOES NOT EXIST (new)
         new_name = lvl.Name + " " + elevation_value

    # Add/Update Levels Elevations
    try:
        lvl.Name = new_name
        print ("Renamed: {) --> {)".format(lvl.Name, new_name))
    except:
        print("Could not change level's name")

if autoTransaction:
     t.Commit()


# go ahead and modify the database



# Report changes
