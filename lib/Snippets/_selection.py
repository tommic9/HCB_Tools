# -*- coding: utf-8 -*-

# Imports
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import  ObjectType,PickBoxStyle, Selection, ISelectionFilter
# Variables
uidoc = __revit__.ActiveUIDocument
doc   = __revit__.ActiveUIDocument.Document
selection = uidoc.Selection # type: Selection

# Functions

def get_selected_elements():
    """Function to get selected elements in Revit UI"""
    return[doc.GetElement(e_id) for e_id in uidoc.Selection.GetElementIds()]

def get_selected_elements_(uidoc=uidoc, filter_types=None, filter_categories=None):
    """Function to get selected elements in Revit UI
    :param uidoc:       Needed to provide uidoc if you want to jump between projects!
    :param filter_types: Option to filter elements by Type
    :param filter_categories: Option to filter by BuiltInCategory
    :return: list of selected elements (Filtered optionally)

    e.g.
    walls_floors = get_selected_elements_(filter_types=[Wall, Floor])
    window_doors = get_selected_elements_(filter_categories=[BuiltInCategory.OST_Windows,
                                                             BuiltInCategory.OST_Doors])"""
    doc = uidoc.Document
    return  [doc.GetElement(e_id) for e_id in uidoc.Selection.GetElementIds()]

# ISelection Filter

class ISelectionFilter_Classes (ISelectionFilter):
    def __init__(self, allowed_types):
        """ISelectionFilter made to filter with types
        :param allowed_types: list of allowed Types"""
        self.allowed_types = allowed_types
    def AllowElement(self, element):
        if type(element) in self.allowed_types:
            return True

class ISelectionFilter_Categories (ISelectionFilter):
    def __init__(self, allowed_categories):
        """ISelectionFilter made to filter with categories
        :param allowed_categories: list of allowed Categories"""
        self.allowed_categories = allowed_categories
    def AllowElement(self, element):
        if element.Category.BuiltInCategory in self.allowed_categories:
            return True