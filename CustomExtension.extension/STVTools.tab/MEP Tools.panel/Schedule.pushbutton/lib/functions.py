import os
import clr

from Autodesk.Revit.DB import (ExternalDefinitionCreationOptions, ParameterType)
from rpw.ui.forms import Alert

clr.AddReference('System')
from System.IO import (StreamReader, StreamWriter)
from System.IO.File import Delete

class SharedParametersFileClass():

	def __init__(self, filePath, uidoc):
		self.filePath = filePath
		self.uidoc = uidoc
	
		try:
			self.filereader = StreamReader( filePath )
			self.uidoc.Application.Application.OpenSharedParameterFile()
		except:
			Alert(
				'Check Your Shared Parameters File Path', 
				title="Error", 
				header="Invalid Shared Parameters File",
				exit=True
			)

		self.temp = os.environ['USERPROFILE'] + "\\AppData\\Local\\temp.txt"


	def __enter__(self):
		self.file = StreamWriter( self.temp )
		while self.filereader.Peek() > -1:
			line = self.filereader.ReadLine()
			if line.startswith(('*', '#', 'META')):
				self.file.WriteLine( line )
		self.file.Close()

		self.uidoc.Application.Application.SharedParametersFilename = self.temp
		self.fileDefinition = self.uidoc.Application.Application.OpenSharedParameterFile()
		self.groupDefinition = self.fileDefinition.Groups.Create('temp_group')

		return self


	def __exit__(self, type, value, traceback):
		Delete( self.temp )
		self.uidoc.Application.Application.SharedParametersFilename = self.filePath


	def copyExternalDefinition(self, guid, definition):
		xDefOptions = ExternalDefinitionCreationOptions(definition.Name, definition.ParameterType)
		xDefOptions.Description = ''
		xDefOptions.GUID = guid

		return self.groupDefinition.Definitions.Create( xDefOptions )