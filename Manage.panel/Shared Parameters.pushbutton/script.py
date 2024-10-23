
# -*- coding: utf-8 -*-
__title__ = "Shared parameters"
__doc__ = """Date    =  07.08.2024
_____________________________________________________________________
Description:
Checking Shared parameters
"""

# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝ IMPORTS
# ==================================================
# Regular + Autodesk
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import  Selection
# pyRevit


# .NET Imports
import clr
clr.AddReference("System")

# ╦  ╦╔═╗╦═╗╦╔═╗╔╗ ╦  ╔═╗╔═╗
# ╚╗╔╝╠═╣╠╦╝║╠═╣╠╩╗║  ║╣ ╚═╗
#  ╚╝ ╩ ╩╩╚═╩╩ ╩╚═╝╩═╝╚═╝╚═╝ VARIABLES
# ==================================================
doc   = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app   = __revit__.Application

selection = uidoc.Selection # type: Selection

# Main

# #Get Parameter Bindings Map
# bm = doc.ParameterBindings
#
# # Create a Forward Iterator
# itor = bm.ForwardIterator()
# itor.Reset()
#
# # Iterate over the map
# # loaded_parameters = []
# while itor.MoveNext():
#     try:
#         d = itor.Key #type: Definition
#         print(d.Name)
#     except:
#         pass
#

# def check_missing_params(list_p_names):
#     """Sprawdzenie czy wszystkie parametry wymagane w projekcie zostały wprowadzone
#     :param list_p_names:
#     :return:
#     """
#     #Get Parameter Bindings Map
#     bm = doc.ParameterBindings
#
#     # Create a Forward Iterator
#     itor = bm.ForwardIterator()
#     itor.Reset()
#
#     # Iterate over the map
#     loaded_parameters = []
#     while itor.MoveNext():
#         try:
#             d = itor.Key #type: Definition
#             loaded_parameters.append(d.Name)
#             # print(d.Name)
#         except:
#             pass
#
#     missing_params = [p_name for p_name in list_p_names if p_name not in loaded_parameters]
#
#     # missing_params = []
#     # for p_name in req_params:
#     #     if p_name not in loaded_parameters:
#     #         missing_params.append(p_name)
#     return missing_params

# if missing_params:
#     print("Missing Parameters:")
#     for p_name in missing_params:
#         print(p_name)

# # Acces Shared parameter file
# sp_file = app.OpenSharedParameterFile()
# #app.SharedParameterFilename = 'C:\Users\tomci\OneDrive - HellCold Sp. z o.o\Biblioteki BIM\010_Parametry współdzielone\Parametry współdzielone Hellcold.txt'
#
# # Find matching definition to missing parameter names
# missing_def = []
# if sp_file:
#     for group in sp_file.Groups:
#         #print('\nGroup Name: {}'.format(group.Name))
#         for p_def in group.Definitions:
#             if p_def.Name in missing_params:
#                 missing_def.append(p_def)
#
# # Select Category
# all_cats = doc.Settings.Categories
# cat_views = all_cats.get_Item(BuiltInCategory.OST_Views)
# cat_sheets = all_cats.get_Item(BuiltInCategory.OST_Sheets)
# cat_ductfitting = all_cats.get_Item(BuiltInCategory.OST_DuctFitting)
# cat_duct = all_cats.get_Item(BuiltInCategory.OST_DuctCurves)
#
# # Create
# cat_set = CategorySet()
# cat_set.Insert(cat_views)
# cat_set.Insert(cat_sheets)
# cat_set.Insert(cat_ductfitting)
# cat_set.Insert(cat_duct)
#
#
# # Create Binding
# binding = InstanceBinding(cat_set)
# #binding = TypeBinding(category_set)
#
# param_group = BuiltInParameterGroup.PG_IDENTITY_DATA
#
# # Add parameters
# t = Transaction(doc, 'AddSharedParameters')
# t.Start()
# for p_def in missing_def:
#     if not doc.ParameterBindings.Contains(p_def):
#         doc.ParameterBindings.Insert(p_def, binding, param_group)
#         print('Adding SharedParameter: {}'.format(p_def.Name))
#
# t.Commit()


# t = Transaction(doc, 'SetVariesByGroup')
# t.Start()
#
# binding_map = doc.ParameterBindings
# it = binding_map.ForwardIterator()
# it.Reset()
#
# while it.MoveNext():
#     p_def = it.Key
#
#     if req_params in p_def.Name:
#         try:
#             p_def.SetAllowCaryBetweenGroup(doc, True)
#             print('SetVaryByGroup for: {}'.format(p_def.Name))
#         except:
#             pass
#
# t.Commit()
from pyrevit import forms

# FUNCTIONS
def get_loaded_parameters_as_def():
    """Get Loaded Parameters in the project as Definitions"""
    definitions = []

    binding_map = doc.ParameterBindings
    it = binding_map.ForwardIterator()
    for i in it:
        definition = it.Key
        definitions.append(definition)
    return definitions

def check_missing_params(list_param_names):
    """Function to check Loaded Shared Parameters.
    :param list_param_names: List of Parameter names
    :return:                 List of Missing Parameter Names"""
    #1️⃣ Read Loaded Parameter Names
    loaded_p_definitions = get_loaded_parameters_as_def()
    loaded_param_names   = [d.Name for d in loaded_p_definitions]

    # Check if Parameters Missing
    missing_params = [p_name for p_name in list_param_names if p_name not in loaded_param_names]
    # 👇 It's the same as
    # missing_params = []
    # for p_name in list_param_names:
    #     if p_name not in loaded_param_names:
    #         missing_params.append(p_name)
    return missing_params

def load_params(param_category_map,
                bind_mode='instance',
                ):
    #type: (list, list, str, BuiltInParameterGroup) -> None
    """Function to check Loaded Shared Parameters.
    :param param_category_map:   Dictionary with Parameter names as keys and list of categories
                                'bic_cats' (list of categories) and 'p_group' (BuiltInParameterGroup)
    :param bind_mode:           Binding Mode: 'instance' / 'type' """

    # 📁 Ensure SharedParameterFile is available
    sp_file = app.OpenSharedParameterFile()
    if not sp_file:
        forms.alert('Could not find SharedParameter File. '
                    '\nPlease Set the File in Revit and Try Again', title=__title__, exitscript=True)
    p_filepath = r'C:\Users\tomci\OneDrive - HellCold Sp. z o.o\Biblioteki BIM\010_Parametry współdzielone\Parametry współdzielone Hellcold.txt'

    #🙋‍ Ask for User Confirmation
    missing_params = list(param_category_map.keys())
    if missing_params:
        confirmed = forms.alert("There are {n_missing} missing parameters for the script."
                                "\n{missing_params}"
                                "\n\nWould you like to try loading them from the following SharedParameterFile:"
                                "\n{p_filepath}".format(n_missing      = len(missing_params),
                                                        missing_params = '\n'.join(missing_params),
                                                        p_filepath     = sp_file.Filename),
                                yes=True, no=True)

        if confirmed:

            # Add Parameters (if possible)
            for d_group in sp_file.Groups:
                for p_def in d_group.Definitions:
                    if p_def.Name in param_category_map:
                        print("Attempting to load parameter: {}".format(p_def.Name))
                        data = param_category_map[p_def.Name]

                        # Prepare Categories
                        all_cats = doc.Settings.Categories
                        cats = [all_cats.get_Item(bic_cats) for bic_cats in data ['bic_cats']]

                        # Create Category Set
                        cat_set = CategorySet()
                        for cat in cats:
                            cat_set.Insert(cat)

                        # Create Binding
                        binding = TypeBinding(cat_set) if bind_mode == 'type' \
                            else InstanceBinding(cat_set)

                        # Insert Parameter Binding
                        doc.ParameterBindings.Insert(p_def, binding, data['p_group'])
                        missing_params.remove(p_def.Name)

            # SetAllowVaryBetweenGroups (If Possible)
            loaded_p_definitions = get_loaded_parameters_as_def()
            for definition in loaded_p_definitions:
                if definition.Name in param_category_map:
                    try:
                        definition.SetAllowVaryBetweenGroups(doc, True)
                    except:
                        pass

            #👀 Reported Not Loaded Parameters
            if missing_params:
                msg = "Couldn't Find following Parameters: \n{}".format('\n'.join(missing_params))
                forms.alert(msg, title=__title__)

# MAIN
required_params = {
    'HC_Area': {
        'bic-cats': [BuiltInCategory.OST_DuctFitting, BuiltInCategory.OST_DuctCurves],
        'p_group' : BuiltInParameterGroup.PG_CONSTRUCTION
    },
    'HC_System': {
        'bic-cats' : [BuiltInCategory.OST_DuctCurves,
                      BuiltInCategory.OST_FlexDuctCurves,
                      BuiltInCategory.OST_DuctFitting,
                      BuiltInCategory.OST_DuctTerminal,
                      BuiltInCategory.OST_DuctAccessory,
                      BuiltInCategory.OST_MechanicalEquipment,
                      BuiltInCategory.OST_PipeCurves,
                      BuiltInCategory.OST_PipeFitting,
                      BuiltInCategory.OST_PipeAccessory,
                      BuiltInCategory.OST_FlexPipeCurves,
                      BuiltInCategory.OST_PlumbingFixtures,
                      BuiltInCategory.OST_PlumbingEquipment],
        'p_group'  : BuiltInParameterGroup.PG_MECHANICAL
    },
    'HC_Duct_System':{
        'bic-cats' :[BuiltInCategory.OST_DuctCurves,
                     BuiltInCategory.OST_FlexDuctCurves,
                     BuiltInCategory.OST_DuctFitting,
                     BuiltInCategory.OST_DuctTerminal,
                     BuiltInCategory.OST_DuctAccessory,
                     BuiltInCategory.OST_MechanicalEquipment],
        'p_group'  : BuiltInParameterGroup.PG_MECHANICAL
    },
    'HC_Piping_System':{
        'bic-cats' : [BuiltInCategory.OST_PipeCurves,
                      BuiltInCategory.OST_PipeFitting,
                      BuiltInCategory.OST_PipeAccessory,
                      BuiltInCategory.OST_FlexPipeCurves,
                      BuiltInCategory.OST_PlumbingFixtures,
                      BuiltInCategory.OST_PlumbingEquipment,
                      BuiltInCategory.OST_MechanicalEquipment],
        'p_group'  : BuiltInParameterGroup.PG_MECHANICAL
    },
    'HC_Do_Zamówienia':{
        'bic-cats' : [BuiltInCategory.OST_DuctCurves,
                      BuiltInCategory.OST_FlexDuctCurves,
                      BuiltInCategory.OST_DuctFitting,
                      BuiltInCategory.OST_DuctTerminal,
                      BuiltInCategory.OST_DuctAccessory,
                      BuiltInCategory.OST_MechanicalEquipment,
                      BuiltInCategory.OST_PipeCurves,
                      BuiltInCategory.OST_PipeFitting,
                      BuiltInCategory.OST_PipeAccessory,
                      BuiltInCategory.OST_FlexPipeCurves,
                      BuiltInCategory.OST_PlumbingFixtures,
                      BuiltInCategory.OST_PlumbingEquipment,
                      BuiltInCategory.OST_ElectricalFixtures],
        'p_group'  : BuiltInParameterGroup.PG_IDENTITY_DATA
    },
    'HC_Zamówione':{
        'bic-cats' : [BuiltInCategory.OST_DuctCurves,
                      BuiltInCategory.OST_FlexDuctCurves,
                      BuiltInCategory.OST_DuctFitting,
                      BuiltInCategory.OST_DuctTerminal,
                      BuiltInCategory.OST_DuctAccessory,
                      BuiltInCategory.OST_MechanicalEquipment,
                      BuiltInCategory.OST_PipeCurves,
                      BuiltInCategory.OST_PipeFitting,
                      BuiltInCategory.OST_PipeAccessory,
                      BuiltInCategory.OST_FlexPipeCurves,
                      BuiltInCategory.OST_PlumbingFixtures,
                      BuiltInCategory.OST_PlumbingEquipment,
                      BuiltInCategory.OST_ElectricalFixtures],
        'p_group'  : BuiltInParameterGroup.PG_IDENTITY_DATA
    },
    'HC_Etap':{
        'bic-cats' : [BuiltInCategory.OST_ProjectInformation],
        'p_group'  : BuiltInParameterGroup.PG_TEXT
    },
    'HC_Branża':{
        'bic-cats' :[BuiltInCategory.OST_ProjectInformation],
        'p_group'  :BuiltInParameterGroup.PG_TEXT
    }
}


missing_params = check_missing_params(list(required_params.keys()))

if missing_params:
    # 🔓 Transaction Start
    t = Transaction(doc, 'Add SharedParameters')
    t.Start()

    # Filter the required_params dictionary to only include missing params
    params_to_load = {k: v for k, v in required_params.items() if k in missing_params}

    try:
        load_params(param_category_map=params_to_load,
                    bind_mode='instance')
    except Exception as e:
        print('Error while adding shared parameters: ', e)
    else:
        print('Parameters successfully added.')

    t.Commit()

# Log results
if not missing_params:
    print("All parameters were already loaded.")
else:
    print("{} parameters were missing and attempted to load.".format(missing_params))