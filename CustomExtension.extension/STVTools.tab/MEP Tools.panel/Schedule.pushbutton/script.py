"""
Push shared parameters from a specified schedule into the selected revit element(s)

Revit API 18.2
"""

__author__ = 'Kris Horn'
__context__ = 'Selection'

# TODO convert shared parameters to family parameters during purge

import sys

from Autodesk.Revit.DB import (FilteredElementCollector, BuiltInCategory, BuiltInParameterGroup, Transaction, ParameterType)
from rpw.ui.forms import (Alert, FlexForm, Label, ComboBox,Separator, Button, CheckBox)
from pyrevit import forms

from lib.functions import SharedParametersFileClass
from lib.FamilyLoadOptions import FamilyLoadOptions

uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document

# Orginal, User Specified Shared Parameters File
pathOriginal = uidoc.Application.Application.SharedParametersFilename

# Get Selected Families
elements = [ doc.GetElement(id) for id in uidoc.Selection.GetElementIds() ]
family_dict = { element.Symbol.FamilyName: element.Symbol.Family 
    for element in elements 
    }

# Get Available Schedules
collection = FilteredElementCollector( doc ).OfCategory( BuiltInCategory.OST_Schedules )
schedule_dict = { e.Name:e 
    for e in collection 
    if not e.Definition.IsKeySchedule and 
		e.Definition.CategoryId in [ family.FamilyCategoryId for family in family_dict.values() ]
    }

if len(family_dict) > 0 and len(schedule_dict) > 0:
	# Create FlexForm
	components = [
		ComboBox('schedule', schedule_dict, sort=True),
		CheckBox('purge', 'Purge Family Parameters (Recommended)', default=True),
		Separator(),
		Button('Schedule')
	]
	form = FlexForm('Select Schedule', components)
	form.show()

	# Get FlexForm Values
	scheduleSelection = form.values.get('schedule')
	toPurge = form.values.get('purge')

	# If User Cancels Script
	if scheduleSelection is None:
		sys.exit(1)

	# Get Shared Paramters from Schedule's Fields
	scheduleDefinition = scheduleSelection.Definition
	schedulableFields = [ scheduleDefinition.GetField( index ).GetSchedulableField() 
		for index in range(0, scheduleDefinition.GetFieldCount()) 
		if not scheduleDefinition.GetField( index ).IsCalculatedField and 
			not scheduleDefinition.GetField( index ).IsCombinedParameterField ]
		
	sharedParameterElements = [ doc.GetElement( field.ParameterId ) 
		for field in schedulableFields 
		if doc.GetElement( field.ParameterId ) is not None ]

	# TODO check for shared parameter
	sharedParameterNames = [ element.Name for element in sharedParameterElements ]
	instanceParameters = forms.SelectFromList.show( sharedParameterNames,
												button_name='Select Instance Parameters',
                                            	multiselect=True
											)
		 
	sharedParameters = { element.Name: 
								{
									'guid':element.GuidValue,
									'definition': element.GetDefinition(), 
									'isInstance': True if element.Name in instanceParameters else False
								}
		for element in sharedParameterElements }
	
	# Edit Families
	for family in family_dict.values():
		familyDoc = doc.EditFamily( family )
		familyManager = familyDoc.FamilyManager
		
		t = Transaction(familyDoc, "Schedule Element(s)")
		t.Start()
		
		# If User Selected to Purge Parameters
		if toPurge:
			ignore = [ # Geometry Specific Paramter Types
					ParameterType.Length,
					ParameterType.HVACDuctSize,
					ParameterType.PipeSize,
					ParameterType.WireSize,
					ParameterType.ElectricalCableTraySize,
					ParameterType.ElectricalConduitSize,
					ParameterType.Angle,
					ParameterType.Material,
					ParameterType.YesNo,
					ParameterType.Slope,
					ParameterType.MassDensity
				]
			# To Purge Family Parameters
			for p in familyManager.GetParameters():
				try:
					if p.Definition.ParameterType not in ignore or 'schedule' in p.Definition.Name.lower():
						# Fails if Parameter is Not Removable
						familyManager.RemoveParameter( p )
				except Exception as e:
					pass

		# Open Temporary Shared Parameters File
		with SharedParametersFileClass( pathOriginal, uidoc ) as spFile:
			for key in sharedParameters: 
				# Create New Shared Parameter as a Duplicate
				guid = sharedParameters.get( key )['guid']
				definition = sharedParameters.get( key )['definition']
				isInstance = sharedParameters.get( key )['isInstance']
				externalDefnition = spFile.copyExternalDefinition( guid, definition )
				try:
					# Fails if Parameter Already Exists
					familyManager.AddParameter( externalDefnition, BuiltInParameterGroup.PG_DATA, isInstance )
				except Exception as e:
					pass
			
		t.Commit()

		# Load Families and Overwrite
		familyDoc.LoadFamily( doc, FamilyLoadOptions() )
		familyDoc.Close()

else:
	Alert(
		'Selected Families: {}\nAvailable Schedules: {}'.format( len(family_dict), len(schedule_dict) ), 
		title="Error", 
		header="There was a Problem",
		exit=True
	)