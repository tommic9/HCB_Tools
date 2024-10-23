# -*- coding: utf-8 -*-
__title__ = "DuctFittingsArea LINEAR"
__doc__ = """Date = 15.10.2024
_____________________________________________________________________
Comment:
Script that calculates the surface area of ventilation fittings according to DIN standard formulas for LINEAR"""

# Imports
import math
import clr
import time

# Autodesk i .NET Imports
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import *
from Autodesk.Revit.DB.Mechanical import *
from pyrevit import forms, script

clr.AddReference("System")
from System.Collections.Generic import List

# Mierzenie czasu rozpoczęcia
start_time = time.clock()

# def convert_internal_units(value, get_internal=True, unit_type='length', doc=None):
#     """
#     Function to convert Internal units to meters or square meters, or vice versa.
#     :param value:        Value to convert
#     :param get_internal: True to get internal units, False to get meters or square meters
#     :param unit_type:    'length' for meters, 'area' for square meters
#     :param doc:          Revit document object
#     :return:             Length or area in Internal units or meters/square meters
#     """
#     if doc is None:
#         doc = __revit__.ActiveUIDocument.Document
#
#     units = doc.GetUnits()
#
#     if unit_type == 'length':
#         if get_internal:
#             return UnitUtils.ConvertToInternalUnits(value, units.GetFormatOptions(SpecTypeId.Length).GetUnitTypeId())
#         else:
#             return UnitUtils.ConvertFromInternalUnits(value, units.GetFormatOptions(SpecTypeId.Length).GetUnitTypeId())
#     elif unit_type == 'area':
#         if get_internal:
#             return UnitUtils.ConvertToInternalUnits(value, units.GetFormatOptions(SpecTypeId.Area).GetUnitTypeId())
#         else:
#             return UnitUtils.ConvertFromInternalUnits(value, units.GetFormatOptions(SpecTypeId.Area).GetUnitTypeId())
#     else:
#         raise ValueError("Unsupported unit_type. Use 'length' or 'area'.")

#Funkcja konwersji jednostek
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

    if get_internal:
        return UnitUtils.ConvertToInternalUnits(value, units)
    return UnitUtils.ConvertFromInternalUnits(value, units)

# Funkcja do pobierania wartości parametru
def get_param_value(param):
    """Get a value from a Parameter based on its StorageType."""
    if not param:
        return None

    storage_type = param.StorageType
    if storage_type == StorageType.Double:
        return param.AsDouble()
    elif storage_type == StorageType.ElementId:
        return param.AsElementId()
    elif storage_type == StorageType.Integer:
        return param.AsInteger()
    elif storage_type == StorageType.String:
        return param.AsString()
    else:
        print('Parameter [{}] is not available for this element'.format(param.Definition.Name))
        return None

# Definicja wymaganych parametrów dla każdego typu kształtki
required_params = {
    'BA': ['A', 'B', 'D', 'E', 'F', 'R', 'LIN_VE_ANG_W'],
    'BO': ['A', 'B', 'E'],
    'BS': ['A', 'B', 'D', 'E', 'F', 'R', 'LIN_VE_ANG_W'],
    'ES': ['L', 'A', 'B', 'D', 'E', 'R'],
    'OA': ['L', 'A', 'B', 'C', 'D', 'E', 'F', 'M', 'T'],
    'OS': ['L', 'A', 'B', 'C', 'D', 'E', 'F', 'M', 'T'],
    'RA': ['L', 'A', 'B', 'C', 'D', 'E', 'F', 'M', 'T'],
    'RS': ['L', 'A', 'B', 'C', 'D', 'E', 'F', 'M', 'T'],
    'SU': ['L', 'A', 'B', 'D', 'R'],
    'TD': ['L', 'A', 'B', 'C', 'D', 'H', 'R'],
    'TG': ['L', 'A', 'B', 'D', 'H', 'M', 'N', 'R1', 'R2'],
    'UA': ['L', 'A', 'B', 'C', 'D', 'E', 'F'],
    'US': ['L', 'A', 'B', 'C', 'D', 'E', 'F'],
    'WA': ['A', 'B', 'D', 'E', 'F', 'R', 'LIN_VE_ANG_W'],
    'WS': ['A', 'B', 'D', 'E', 'F', 'R', 'LIN_VE_ANG_W'],
    'HS': ['A', 'B', 'D', 'E', 'M', 'L', 'H']
}

# Funkcja do obliczania obwodu (OB) i długości (L) na podstawie typu kształtki i parametrów
def calculate_L_Obw(dim_type, params):
    L, Obw = None, None

    if dim_type == 'BS':
        Obw = 2 * (params['A'] + params['B'])
        L = (params['LIN_VE_ANG_W'] * (params['R'] + params['B'])) + params['E'] + params['F']
    elif dim_type == 'BO':
        Obw = 1
        L1 = (2 * (params['A'] + params['B'])) * params['E']
        A = params['A'] * params['B']
        L = L1 + A
    elif dim_type == 'BA':
        if params['B'] >= params['D']:
            Obw = 2 * (params['A'] + params['B'])
            L = (params['LIN_VE_ANG_W'] * (params['R'] + params['B'])) + params['E'] + params['F']
        else:
            Obw = 2 * (params['A'] + params['D'])
            L = (params['LIN_VE_ANG_W'] * (params['R'] + params['B'])) + params['E'] + params['F']
    elif dim_type == 'WS':
        Obw = 2 * (params['A'] + params['B'])
        L = 2 * params['B'] + params['E'] + params['F']
    elif dim_type == 'WA':
        if params['B'] >= params['D']:
            Obw = 2 * (params['A'] + params['B'])
            L = params['B'] + params['D'] + params['E'] + params['F']
        else:
            Obw = 2 * (params['A'] + params['D'])
            L = params['B'] + params['D'] + params['E'] + params['F']
    elif dim_type == 'US':
        if params['A'] + params['B'] >= params['C'] + params['D']:
            Obw = 2 * (params['A'] + params['B'])
            L = math.sqrt(params['L'] ** 2 + (params['E'] ** 2))
        else:
            Obw = 2 * (params['C'] + params['D'])
            L = math.sqrt(params['L'] ** 2 + (params['F'] ** 2))
    elif dim_type == 'UA':
        if params['A'] + params['B'] >= params['C'] + params['D']:
            Obw = 2 * (params['A'] + params['B'])
            L = math.sqrt(params['L'] ** 2 + (params['B'] - params['D'] + params['E']) ** 2)
        else:
            Obw = 2 * (params['C'] + params['D'])
            L = math.sqrt(params['L'] ** 2 + (params['E'] ** 2))
    elif dim_type == 'OS':
        if params['A'] + params['B'] >= 2 * math.pi * math.sqrt((2 * params['D'] + 2 * params['C'])/2):
            Obw = 2 * (params['A'] + params['B'])
            L = math.sqrt(params['L'] ** 2 + (params['E'] ** 2))
        else:
            Obw = 2 * math.pi * math.sqrt((2 * params['D'] + 2 * params['C'])/2)
            L = math.sqrt(params['L'] ** 2 + (params['F'] ** 2))
    elif dim_type == 'OA':
        if params['A'] + params['B'] >= (2 * math.pi * math.sqrt((2 * params['D'] + 2 * params['C']) / 2)) / 2:
            Obw = 2 * (params['A'] + params['B'])
        else:
            Obw = 2 * math.pi * math.sqrt((2 * params['D'] + 2 * params['C']) / 2)
        if params['B'] - params['D'] + params['E'] >= params['E']:
            L = math.sqrt(params['L'] ** 2 + (params['B'] - params['D'] + params['E']) ** 2)
        elif params['A'] - params['D'] + params['F'] >= params['F']:
            L = math.sqrt(params['L'] ** 2 + (params['A'] - params['D'] + params['F']) ** 2)
        else:
            L = math.sqrt(params['L'] ** 2 + max(params['E'], params['F']) ** 2)
    elif dim_type == 'RS':
        if params['A'] + params['B'] >= (math.pi * params['D'])/2:
            Obw = 2 * (params['A'] + params['B'])
        else:
            Obw = math.pi * params['D']
        if params['E'] >= params['F']:
            L = math.sqrt(params['L'] ** 2 + (params['E'] ** 2))
        else:
            L = math.sqrt(params['L'] ** 2 + (params['F'] ** 2))
    elif dim_type == 'RA':
        if params['A'] + params['B'] >= (math.pi * params['D'])/2:
            Obw = 2 * (params['A'] + params['B'])
        else:
            Obw = math.pi * params['D']
        if params['B'] - params['D'] + params['E'] >= params['E']:
            L = math.sqrt(params['L'] ** 2 + (params['B'] - params['D'] + params['E']) ** 2)
        elif params['A'] - params['D'] + params['F'] >= params['F']:
            L = math.sqrt(params['L'] ** 2 + (params['A'] - params['D'] + params['F']) ** 2)
        else:
            L = math.sqrt(params['L'] ** 2 + max(params['E'], params['F']) ** 2)
    elif dim_type == 'ES':
        Obw = 2 * (params['A'] + params['B'])
        L = math.sqrt(params['L'] ** 2 + params['E'] ** 2)
    elif dim_type == 'EA':
        if params['B'] >= params['D']:
            Obw = 2 * (params['A'] + params['B'])
            L = math.sqrt(params['L'] ** 2 + (params['B'] - params['D'] + params['E']) ** 2)
        else:
            Obw = 2 * (params['C'] + params['D'])
            L = math.sqrt(params['L'] ** 2 + (params['E'] ** 2))
    elif dim_type == 'TG':
        if params['A'] + params['B'] >= params['A'] + params['D']:
            Obw1 = 2 * (params['A'] + params['B'])
        else:
            Obw1 = 2 * (params['A'] + params['D'])
        L1 = params['L']
        Obw2 = 2 * (params['A'] + params['H'])
        if params['D'] + params['M'] - params['B'] >= params['M']:
            L2 = params['D'] + params['M'] - params['B']
        else:
            L2 = params['M']
        L = L1 * Obw1 + L2 * Obw2
        Obw = 1
    elif dim_type == 'TA':
        if params['B'] >= params['D']:
            Obw1 = 2 * (params['A'] + params['B'])
        else:
            Obw1 = 2 * (params['C'] + params['D'])
        L1 = math.sqrt(params['L'] ** 2 + (params['E'] ** 2))
        Obw2 = 2 * (params['A'] + params['H'])
        if params['D'] + params['M'] - params['B'] - params['E'] >= params['M']:
            L2 = params['D'] + params['M'] - params['B'] - params['E']
        else:
            L2 = params['M']
        L = L1 * Obw1 + L2 * Obw2
        Obw = 1
    elif dim_type == 'HS':
        params['M'] = max(params['M'], 100)  # ensure M is at least 100
        if params['B'] >= params['D'] + params['M'] + params['H']:
            Obw = 2 * (params['A'] + params['B'])
            L = math.sqrt(params['L'] ** 2 + (params['B'] - params['H'] - params['M'] - params['D'] + params['E']) ** 2)
        else:
            Obw = 2 * (params['C'] + params['D'] + params['M'] - params['H'])
            L = math.sqrt(params['L'] ** 2 + (params['E'] ** 2))
    else:
        raise ValueError("Nieznany typ kształtki: {}".format(dim_type))

    return L, Obw

# Główna funkcja
def main():
    doc = __revit__.ActiveUIDocument.Document
    uidoc = __revit__.ActiveUIDocument
    app = __revit__.Application
    output = script.get_output()

    # Pobranie wszystkich kształtek wentylacyjnych
    all_duct_fittings = FilteredElementCollector(doc) \
        .OfCategory(BuiltInCategory.OST_DuctFitting) \
        .OfClass(FamilyInstance) \
        .ToElements()

    # Filtracja elementów z parametrem Size zawierającym znak "x"
    all_duct_fittings = [df for df in all_duct_fittings if df.LookupParameter("Size").AsString().Contains("x")]

    # Mapowanie nazw parametrów
    sparameter = ['A', 'B', 'C', 'D', 'E', 'F', 'H', 'L', 'M', 'M2', 'N', 'R', 'R1', 'R2', 'R3', 'R4', 'T']
    parameter_names = ['LIN_VE_DIM_' + p for p in sparameter]
    parameter_names.append('LIN_VE_ANG_W')

    # Przetwarzanie każdej kształtki wentylacyjnej
    for df in all_duct_fittings:
        # Pobieranie wartości parametrów
        param_values = {p: get_param_value(df.LookupParameter('LIN_VE_DIM_' + p)) for p in sparameter}
        param_values['LIN_VE_ANG_W'] = get_param_value(df.LookupParameter('LIN_VE_ANG_W'))

        HC_Area = df.LookupParameter("HC_Area")
        dim_type = get_param_value(df.LookupParameter("LIN_VE_DIM_TYP"))

        # Sprawdzanie brakujących parametrów przed obliczeniami
        if dim_type not in required_params:
            print("Nieznany typ kształtki: {}".format(dim_type))
            continue

        missing_params = [p for p in required_params[dim_type] if param_values.get(p) is None]
        if missing_params:
            print("Nie znaleziono parametru o ID {}: {}".format(df.Id, missing_params))
            continue

        # Sprawdzanie brakujących parametrów przed obliczeniami
        try:
            L, Obw = calculate_L_Obw(dim_type, param_values)
            L_m = convert_internal_units(L,True, 'm')
            Obw_m = convert_internal_units(Obw, True, 'm')
            area = L_m * Obw_m
            round_up_area = round(area, 2) / 10
            df_linkify = output.linkify(df.Id, df.Symbol.Family.Name)

            with Transaction(doc, __title__) as t:
                t.Start()

                if HC_Area:
                    if round_up_area<1:
                        HC_Area.Set(1)
                    else:
                        HC_Area.Set(round_up_area)
                else:
                    print("Parameter 'HC_Area' not found in element {}".format(df.Id))
                t.Commit()
        except ValueError as e:
            print(e)
            continue

        # Drukowanie szczegółów tylko jeśli którykolwiek parametr ma wartość inną niż None
        if any(value is not None for value in param_values.values()):
            print("TYP: {} , Fam_Name: {}, ID: {}".format(dim_type, df_linkify, df.Id))
            if HC_Area:
                print("HC_Area: {} m²".format(round_up_area/10 if round_up_area/10 >= 1 else 1))
            print('-' * 50)

if __name__ == "__main__":
    main()

# Mierzenie czasu zakończenia
end_time = time.clock()
execution_time = end_time - start_time
print('-' * 50)
print("Time of script:{} sec.".format(execution_time))