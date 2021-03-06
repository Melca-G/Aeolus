from Autodesk.Revit.DB import IFamilyLoadOptions

# Override Existing Family and Parameters
class FamilyLoadOptions(IFamilyLoadOptions):
	def OnFamilyFound(self, familyInUse, overwriteParameterValues):
		return True
	
	def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
		return True