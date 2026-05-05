Option Explicit 

CONST VERSION = "dev"

Sub Build
	CreateAddinFrom "..\ExportExcelToTxt.xls"
	FileRenameAddSuffix "..\ExportExcelToTxt.xla" , "-" & VERSION
End Sub