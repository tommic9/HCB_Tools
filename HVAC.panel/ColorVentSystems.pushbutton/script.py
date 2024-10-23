# -*- coding: utf-8 -*-
__title__ = "Color Vent Systems"
__doc__ = """Date  = 01.11.2023 
_____________________________________________________________________
Komentarz:
To jest skrypt który nadpisuje systemy kanałów wypełnieniem według ustalongo schematu Lineara
It is a script which coolorize ducts systems by Linear layout"""
# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝ IMPORTS
# ==================================================
# Regular + Autodesk
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import *
from Autodesk.Revit.DB.Mechanical import *
# pyRevit
from pyrevit import forms

# .NET Imports
import clr
clr.AddReference("System")
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

def create_view_filters_and_set_overrides(doc, view, system_names, color_map):
    """
    Tworzy filtry widoku dla systemów wentylacyjnych i stosuje nadpisania wypełnienia.
    :param doc: Dokument Revit
    :param view: Widok Revit, w którym mają być zastosowane filtry
    :param system_names: Lista nazw systemów do utworzenia filtrów
    :param color_map: Słownik mapujący nazwy systemów na kolory (RGB)
    """

    solid_fill_pattern_id = ElementId.InvalidElementId
    for fill_pattern in FilteredElementCollector(doc).OfClass(FillPatternElement):
        if fill_pattern.Name == "<Solid fill>":
            solid_fill_pattern_id = fill_pattern.Id
            break

    # Pobierz listę ID filtrów już zastosowanych do widoku
    applied_filter_ids = view.GetFilters()

    for system_name in system_names:
        color = color_map.get(system_name, None)
        if color:
            revit_color = Color(color[0], color[1], color[2])
            ogs = OverrideGraphicSettings()
            ogs.SetSurfaceForegroundPatternColor(revit_color)
            ogs.SetSurfaceForegroundPatternId(solid_fill_pattern_id)

            # Tworzenie reguły filtru dla nazwy systemu
            rule = ParameterFilterRuleFactory.CreateContainsRule(ElementId(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM, True))
            elementFilter = ElementParameterFilter(rule)

            categories = List[ElementId]([ElementId(BuiltInCategory.OST_DuctCurves),
                                          ElementId(BuiltInCategory.OST_DuctFitting),
                                          ElementId(BuiltInCategory.OST_DuctAccessory),
                                          ElementId(BuiltInCategory.OST_AirTerminal),
                                          ElementId(BuiltInCategory.OST_FlexDuct)])
            # Sprawdź, czy filtr o tej nazwie już istnieje
            existing_filter_id = None
            for filter in FilteredElementCollector(doc).OfClass(ParameterFilterElement).ToElements():
                if filter.Name == system_name:
                    existing_filter_id = filter.Id
                    break

            if existing_filter_id is None:
                # Utwórz nowy filtr, jeśli nie istnieje
                filter = ParameterFilterElement.Create(doc, system_name, categories, elementFilter)
            else:
                # Jeśli filtr już istnieje, użyj istniejącego ID
                filter = doc.GetElement(existing_filter_id)

            # Sprawdź, czy filtr jest już zastosowany do widoku
            if not filter.Id in applied_filter_ids:
                # Dodaj filtr do widoku, jeśli nie jest jeszcze zastosowany
                view.AddFilter(filter.Id)

            # Jeśli filtr jest już zastosowany, zaktualizuj nadpisania
            view.SetFilterOverrides(filter.Id, ogs)

# Przykładowe użycie
doc = __revit__.ActiveUIDocument.Document
view = doc.ActiveView
system_names = ['V_Exhaust', 'V_Supply', 'V_Extract', 'V_Outdoor']
color_map = {
    'V_Exhaust air (Wyrzut)': (192, 128, 0),  # Brązowy
    'V_Supply air (Nawiew)': (0, 128, 255),  # Błękitny
    'V_Extract air (Wywiew)': (255, 255, 0),  # Żółty
    'V_Outdoor air (Czerpny)': (82, 165, 82)  # Zielony
}
t = Transaction(doc, 'Tworzenie filtrów widoku i nadpisania')
t.Start()
create_view_filters_and_set_overrides(doc, view, system_names, color_map)
t.Commit()
