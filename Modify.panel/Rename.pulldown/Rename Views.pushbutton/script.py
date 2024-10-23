# -*- coding: utf-8 -*-
__title__ = "Rename Views"
__doc__ = """Date    =  14.11.2023
_____________________________________________________________________
Description:
Rename views with Replace, Prefix and Sufix
"""

# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝ IMPORTS
# ==================================================
# Regular + Autodesk
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import ObjectType, Selection, ISelectionFilter
# pyRevit
from pyrevit import forms

# .NET Imports
import clr
clr.AddReference("System")
from System.Collections.Generic import List

# Custom
from Snippets._selection import get_selected_elements

# ╦  ╦╔═╗╦═╗╦╔═╗╔╗ ╦  ╔═╗╔═╗
# ╚╗╔╝╠═╣╠╦╝║╠═╣╠╩╗║  ║╣ ╚═╗
#  ╚╝ ╩ ╩╩╚═╩╩ ╩╚═╝╩═╝╚═╝╚═╝ VARIABLES
# ==================================================
doc   = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app   = __revit__.Application

selection = uidoc.Selection # type: Selection

# Main

# 1. GET VIEW
selected_elements = get_selected_elements()
selected_views    = [el for el in selected_elements if issubclass(type(el), View)]

if not selected_views:
    selected_views = forms.select_views()

    if not selected_views:
        forms.alert('Nie wybrano widoku. \n'
                    'Spróbuj ponownie' , exitscript=True )

# 2. GET USER INPUT

# Window with field to insert a prefix, find, replace and suffix
from rpw.ui.forms import (FlexForm, Label, TextBox, Separator, Button)
components = [  Label('Prefix:'),   TextBox('prefix'),
                Label('Find:'),     TextBox('find'),
                Label('Replace:'),  TextBox('replace'),
                Label('Suffix:'),   TextBox('suffix'),
                Separator(),
                Button('Rename Views')]
form = FlexForm('Title', components)
form.show()

# value of inputs
user_inputs = form.values
prefix  = user_inputs['prefix']
find    = user_inputs['find']
replace = user_inputs['replace']
suffix  = user_inputs['suffix']

# iteration for every selected views
for view in selected_views:

    # 3. NEW NAME
    current_name = view.Name
    # method of rename views - like a manually
    new_name     = prefix + view.Name.replace(find, replace) + suffix


    # 4. RENAME VIEWS

    with Transaction(doc, __title__) as t:
        t.Start()

        for i in range(20):
            try:
                # printing info about old name compare to new name
                view.Name = new_name
                print ("{} -> {}".format(current_name, new_name))
                break
            except:
                new_name += "*"

        t.Commit()



