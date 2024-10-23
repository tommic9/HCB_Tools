# -*- coding: utf-8 -*-

__title__ = "SpaceParamLinked"
__doc__ = """Narzędzie które pobiera nazwę i numer przestrzeni z podlinkowanego projektu i przypisuje do elementów w projekcie""

How-to:


Author: Tomasz Michalek"""

# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝ IMPORTS
# ==================================================
# Regular + Autodesk
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import *
from Autodesk.Revit.DB.Mechanical import *
# Importowanie niezbędnych modułów Revit API
import clr
import sys

# Dodanie ścieżek do Revit API
clr.AddReference('RevitServices')
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

clr.AddReference('RevitAPIUI')
from Autodesk.Revit.UI import *

# Dostęp do aktywnego dokumentu
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
doc = uiapp.ActiveUIDocument.Document
from pyrevit import forms


# .NET Imports
import clr
clr.AddReference("System")
import random
from System.Collections.Generic import List

# ╦  ╦╔═╗╦═╗╦╔═╗╔╗ ╦  ╔═╗╔═╗
# ╚╗╔╝╠═╣╠╦╝║╠═╣╠╩╗║  ║╣ ╚═╗
#  ╚╝ ╩ ╩╩╚═╩╩ ╩╚═╝╩═╝╚═╝╚═╝ VARIABLES
# ==================================================
doc   = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app   = __revit__.Application

# ╔╦╗╔═╗╦╔╗╔
# ║║║╠═╣║║║║
# ╩ ╩╩ ╩╩╝╚╝ MAIN
# ==================================================

def convert_internal_units(value, get_internal=True, units='mm'):
    # type: (float, bool, str) -> float
    """Funkcja do konwersji jednostek wewnętrznych na metry lub odwrotnie.
    :param value:        Wartość do konwersji
    :param get_internal: True, aby uzyskać jednostki wewnętrzne, False, aby uzyskać metry
    :param units:        Wybierz pożądane jednostki: ['m', 'm2', 'cm', 'mm']
    :return:             Długość w jednostkach wewnętrznych lub metrach."""

    from Autodesk.Revit.DB import UnitTypeId
    if units == 'm':
        units = UnitTypeId.Meters
    elif units == "m2":
        units = UnitTypeId.SquareMeters
    elif units == 'cm':
        units = UnitTypeId.Centimeters
    elif units == 'mm':
        units = UnitTypeId.Millimeters
    elif units == 'm3/h':
        units = UnitTypeId.CubicMetersPerHour

    if get_internal:
        return UnitUtils.ConvertToInternalUnits(value, units)
    return UnitUtils.ConvertFromInternalUnits(value, units)

# Function to get air terminals in the active view
# Funkcja do pobierania wszystkich elementów w kategorii "Mechanical Equipment" z danego dokumentu
def get_mechanical_equipment_elements(document):
    collector = FilteredElementCollector(document)
    category = BuiltInCategory.OST_MechanicalEquipment
    collector.OfCategory(category).WhereElementIsNotElementType()
    return collector.ToElements()

# from rpw.ui.forms import (FlexForm, Label, TextBox, Separator, Button)
# components = [  Label('Project Parameter Name for new value:'),   TextBox('param_name'),
#                 Separator(),
#                 Button('Set parameter name')]
# form = FlexForm('Title', components)
# form.show()
#
# user_inputs = form.values
# actual_flow_param_name = user_inputs['param_name']
# #actual_flow_param_name = "Actual Flow"
#
# air_terminals = get_air_terminals_in_active_view(doc)


# Funkcja do pobierania wartości parametrów przestrzeni powiązanej z elementem
def get_space_parameters(element, document):
    # Uzyskanie referencji przestrzeni
    space_id = element.get_Parameter(BuiltInParameter.RBS_ELEM_SPACE).AsElementId()
    if space_id == ElementId.InvalidElementId:
        return None, None  # Element nie jest przypisany do żadnej przestrzeni
    space = document.GetElement(space_id)
    if space:
        space_number = space.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsString()
        space_name = space.get_Parameter(BuiltInParameter.ROOM_NAME).AsString()
        return space_number, space_name
    return None, None


# Funkcja do aktualizacji parametrów współdzielonych
def update_shared_parameters(element, lin_room_number, lin_room_name):
    # Pobranie parametrów współdzielonych
    param_number = element.LookupParameter("LIN_ROOM_NUMBER")
    param_name = element.LookupParameter("LIN_ROOM_NAME")

    if param_number and param_name:
        if param_number.IsReadOnly or param_name.IsReadOnly:
            return  # Nie można modyfikować parametrów tylko do odczytu
        param_number.Set(lin_room_number if lin_room_number else "")
        param_name.Set(lin_room_name if lin_room_name else "")


# Funkcja główna
def main():
    # Start a transaction
    with Transaction(doc, __title__) as t:
        t.Start()

    # Lista wszystkich dokumentów (główny + podlinkowane)
    all_documents = [doc] + list(uiapp.Application.Documents.[RevitLinkInstance]())

    for linked in all_documents:
        if isinstance(linked, RevitLinkInstance):
            # Dostęp do dokumentu podlinkowanego
            linked_doc = linked.GetLinkDocument()
            if not linked_doc:
                continue
        else:
            linked_doc = linked  # Główny dokument

        # Pobranie elementów "Mechanical Equipment"
        elements = get_mechanical_equipment_elements(linked_doc)

        for elem in elements:
            space_number, space_name = get_space_parameters(elem, linked_doc)
            if space_number is not None or space_name is not None:
                # Aktualizacja parametrów współdzielonych w elemencie
                update_shared_parameters(elem, space_number, space_name)

    # Zakończenie transakcji
    t.Commit()

    # Powiadomienie użytkownika
    TaskDialog.Show("Sukces", "Parametry zostały zaktualizowane pomyślnie.")


# Uruchomienie funkcji głównej
if __name__ == "__main__":
    main()
