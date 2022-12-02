Option Explicit

''' You NEED Microsoft Word installed on the computer executing this vbs script
'''
''' VB script for converting all .docx documents in a folder to PDF
''' Usage: copy vbs to folder with word documents and execute.
''' Mandatory params are two positional args searchText and replaceText
'''
' debug with cscript //X ConvertAllDocxToPdf.vbs [/p:<folder-with-docx-files-path>] [/a|pdfa]

' see WdSaveFormat enumeration constants:
Const wdFormatPDF = 17  ' PDF format. 
Dim saveAsPdfA

Dim fso
Dim workPath
Dim workDir

Dim WordApp
Dim WordDoc

' Global variables
Dim arguments
Set arguments = WScript.Arguments

Set fso = CreateObject("Scripting.FileSystemObject")

' Parse args and set variables from args
Sub ParseArgs()
 
  If arguments.Named.Exists("p") Or arguments.Named.Exists("path") Then
    workPath = arguments.Named.Item("p")
	If workPath = Empty Then
	  workPath = arguments.Named.Item("path")
	End if
	
	If not fso.FolderExists(workPath) Then
	  WScript.Echo("Path " & workPath & " does not exist")
	  WScript.Quit 1
	End if
  Else 
    workPath = fso.GetParentFolderName(WScript.ScriptFullName)
  End If
  
  saveAsPdfA = arguments.Named.Exists("a") Or arguments.Named.Exists("pdfa")
  
End Sub

ParseArgs

'Create a Microsoft Word Object and make it invisible. 
Set WordApp = CreateObject("Word.Application")
WordApp.Visible = FALSE

'Find all files in current directory
Set workDir = fso.GetFolder(workPath)

Dim objFiles
Set objFiles = workDir.Files

Dim objFile
Dim pdfFileName

'Iterate through files of type docx and call replacement function text
For Each objFile In objFiles

  If UCase(fso.GetExtensionName(objFile.Name)) = "DOCX" Then

	' Instantiate word doc
    Set WordDoc = WordApp.Documents.Open(objFile.Path)
	
	pdfFileName = workDir.Path + "\" + fso.GetBaseName(objFile.Name) + ".pdf"
	
	' Focus to doc
	WordApp.Documents(objFile.Name).Activate 'switch to open document
    
	' Save as PDF
	WordDoc.ExportAsFixedFormat pdfFileName, wdFormatPDF, , 1, , , , , , , 1, , , saveAsPdfA
	
    WordDoc.Close 0

  End If

Next

WordApp.Quit