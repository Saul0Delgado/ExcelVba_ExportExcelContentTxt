Option Explicit

DoBuild

CONST b_BUILDSUB = "Build"
Dim b_xl		'Referencia al Objeto Excel.
Dim b_Path 		'Ruta de ejecución del script.

Dim b_wsh
Dim b_fso

Sub DoBuild
	Set b_wsh = CreateObject("WScript.Shell")
	Set b_fso = CreateObject( "Scripting.FileSystemObject" )
	
	b_path = b_wsh.CurrentDirectory
	OpenExcel
	b_xl.DisplayAlerts = False

	'on error resume next
	Execute b_BUILDSUB
	if err.number<>0 then
	  println ErrorString : err.clear
	end if

	CloseExcel
end sub

Sub OpenExcel()
	if IsValidObject(b_xl) then exit sub

	set b_xl = Nothing ''Objetizamos' la variable.
	on error resume next : Set b_xl = GetObject(, "Excel.Application") : on error goto 0

	If b_xl Is Nothing Then _
		Set b_xl = CreateObject("Excel.Application")

End sub

Sub CloseExcel()
	If IsValidObject(b_xl) Then _
		b_xl.Quit

	Set b_xl = Nothing
End Sub

Sub PrintLn(Message)
  Wscript.Echo Message
End Sub

Function IsValidObject(ByRef obj)
	IsValidObject = False
	if IsObject(obj) Then _
		if not obj is nothing Then _ 
			IsValidObject=true
end function

Function ErrorString()
  ErrorString = "[ERROR] No. " & err.number & " - " & err.description & " - " & err.source
end function

function OpenWorkbook(relativePath)
  set OpenWorkbook = b_xl.Workbooks.Open(b_path & "\" & relativePath)
end function

function PathRemoveExtension(path)
  PathRemoveExtension = Left(path, (InStrRev(path, ".", -1, vbTextCompare) - 1))
end function

'--------------------------------------------------------------------------------
'Activity: Toma un .XLS y graba un .XLA con la propiedad IsAddin = True.
'Misma carpeta, mismo nombre, extensión diferente.

CONST ADDIN_EXTENSION = ".xla"

sub CreateAddInFrom(relativePath)
  dim wb: set wb = OpenWorkbook(relativePath)

  Dim FinalName
  FinalName = PathRemoveExtension(wb.FullName)
  FinalName = FinalName & ADDIN_EXTENSION

  wb.IsAddin = True
  wb.SaveAs FinalName 
  wb.Close
end sub

'--------------------------------------------------------------------------------
'Activity: Renombra un archivo, agregándole un sufijo. Deja la extensión intacta.

sub FileRenameAddSuffix(filename, suffix)
	Dim finalName
	finalName = b_fso.GetBaseName(filename) & suffix & "." & b_fso.GetExtensionName(filename)

	FileRename filename, finalName
end sub

sub FileRename(src, dest)
	dest = b_fso.buildpath( b_fso.getparentfoldername(src) , dest)
	if b_fso.fileExists(dest) then b_fso.deletefile(dest)
	b_fso.movefile src, dest
end sub