"""
Quick Access to Project Links

Revit API 18.2
"""

__author__ = 'Kris Horn'

import os
import sys
import subprocess

from Autodesk.Revit.DB import (FilteredElementCollector, CADLinkType, RevitLinkType)
from Autodesk.Revit.DB.ModelPathUtils import ConvertModelPathToUserVisiblePath
from rpw.ui.forms import (Alert, FlexForm, Label, ComboBox,Separator, Button, CheckBox)
from pyrevit import forms

uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document

links = {}

# Filter for Revit Model Links
filteredRevit = FilteredElementCollector( doc )\
                .OfClass( RevitLinkType )\
                .ToElements()

# Filter for CAD Links and Imports
filteredCAD = FilteredElementCollector( doc )\
                .OfClass( CADLinkType )\
                .ToElements()

# Retrive Aboslute File Path for each Revit Model Link and add to Dictionary
for i in filteredRevit:
    file_path = ConvertModelPathToUserVisiblePath( i.GetExternalFileReference().GetAbsolutePath() )
    links.update({ os.path.basename(file_path): os.path.dirname(file_path)})

    # Update Central Dictionary Containing Links
    if os.path.exists(file_path):
        links.update({ os.path.basename(file_path): os.path.dirname(file_path)})
    else: # Indicate Broken Revit Links
        links.update({ '(FILE NOT FOUND) {}'.format( os.path.basename(file_path) ): os.path.dirname(file_path)})

# Retrieve Absolute File Path for each CAD Link and add to Dictionary
for i in filteredCAD:
    if i.IsExternalFileReference():# (Ignore CAD Imports)
        file_path = ConvertModelPathToUserVisiblePath( i.GetExternalFileReference().GetAbsolutePath() )

        # Update Central Dictionary Containing Links
        if os.path.exists(file_path): 
            links.update({ os.path.basename(file_path): os.path.dirname(file_path)})
        else: # Indicate Broken CAD Links
            links.update({ '(FILE NOT FOUND) {}'.format( os.path.basename(file_path) ): os.path.dirname(file_path)})


# If there are no Linked Documents
if not links.keys():
    	Alert(
		'There are no links loaded into the project', 
		title="No Links Available",
		exit=True
	)

# Create Form
selection = forms.SelectFromList.show( sorted(links.keys(), key=lambda x: (os.path.splitext(x)[1], os.path.splitext(x)[0])),
										button_name='Select Links to Open',
                                        multiselect=True
									)

# If User Cancels Script
if selection is None:
    sys.exit(1)

# Open the Directory in Windows Explorer for each Linked Document Selected
selectedPaths = list(set([ links[key] for key in selection ]))
for file_path in selectedPaths:
    if os.path.exists(file_path):
        subprocess.Popen('explorer {}'.format( file_path ))