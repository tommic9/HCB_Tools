# -*- coding: utf-8 -*-
__title__ = "Rename Views"
__doc__ = """Date    = 15.11.2023 
_____________________________________________________________________
Description:
Rename views with replace, prefix and suffix
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

# 1 GET VIEWS
selected_elements = get_selected_elements()
selected_views    = [el for el in selected_elements if issubclass(type(el),View)]

if not selected_views:
    selected_views = forms.select_views()

    if not selected_views:
        forms.alert('Nie wybrano widku. \n'
                     'Spróbuj ponownie', exitscript=True)

# 2 GET USER INPUT
 from rpw.ui.forms import (FlexForm, Label, ComboBox, TextBox, TextBox,
                           Separator, Button, CheckBox)
 components = [Label('Prefix:'),    TextBox('prefix'),
               Label('Find:'),      TextBox('find'),
               Label('Replace:'),    TextBox('replace'),
               Label('Suffix:'),      TextBox('suffix'),
               Separator(), Button('Rename views')]
 form = FlexForm('Title', components)
 form.show()

 # value of inputs
 user_inputs = form.values
 prefix     = user_inputs['prefix']
 find       = user_inputs['find']
 replace    = user_inputs['replace']
 suffix     = user_inputs['suffix']


 for view in selected_views:

    # 3 NEW NAME
     current_name = view.Name
     new_name     = prefix + view.Name.replace(find, replace) + suffix


# 4 RENAME VIEWS
    with Transaction(doc, __title__) as t:
        t.Start()

        for i in range(20):
            try:
                view.Name = new_name
                print ("{} -> {}".format(current_name,new_name))
                break
            except:
                new_name += "*"

        t.Commit()