# -*- coding: utf-8 -*-

# Imports
from Autodesk.Revit.DB import *
# Variables
uidoc = __revit__.ActiveUIDocument
doc   = __revit__.ActiveUIDocument.Document
app   = __revit__.Application

rvt_year = int(app.VersionNumber)

# Functions

def convert_internal_units(length, get_internal = True, unit = 'm'):
    #type:  (float, bool) -> float
    """Function to convert Internal units to meter or vice versa
    :param length:          Value to convert
    :param get_internal:    Tru to get internal units, False to get
    :param unit:            Select desired units: ['m', 'm2','cm','cm2','mm','mm2']
    :return:                Length in Internal units or meters
    """

    if rvt_year >= 2021
        # New method
        from Autodesk.Revit.DB import UnitTypeId
        if unit =='m':       units = UnitTypeId.Meters
        elif unit =='m2':    units = UnitTypeId.SquareMeters
        elif unit =='mm':    units = UnitTypeId.Millimeters
        elif unit =='mm2':   units = UnitTypeId.SquareMillimeters
        elif unit =='cm':    units = UnitTypeId.Centimeters
        elif unit =='cm2':   units = UnitTypeId.SquareCentimeters
    else:
        from Autodesk.Revit.DB import DisplayUnitType
        if unit == 'm':     units = DisplayUnitType.DUT_METERS
        elif unit == 'm2':  units = DisplayUnitType.DUT_SQUARE_METERS

    if get_internal:
        return UnitUtils.ConvertToInternalUnits(length, units)

    return UnitUtils.ConvertFromInternalUnits(length, units)