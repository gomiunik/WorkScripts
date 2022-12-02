Option Explicit

''' You NEED Microsoft Word installed on the computer executing this vbs script
'''
''' VB script for searching and replacing text in all Word documents
''' Usage: copy vbs to folder with word documents and execute.
''' Mandatory params are two positional args searchText and replaceText
'''
' debug with cscript //X FindAndReplaceAllDocs.vbs "searchText" "replaceText"


Const wdReplaceAll  = 2

Dim fso
Dim workPath
Dim workDir

Dim WordApp
Dim WordDoc

Dim sOldText
Dim sNewText
Dim bWholeWord
Dim bCaseInsensitive

' Global variables
Dim arguments
Set arguments = WScript.Arguments

Set fso = CreateObject("Scripting.FileSystemObject")

' Parse args and set variables from args
Sub ParseArgs()

  If arguments.Unnamed.Count < 2 Then
    WScript.Echo "Usage: " & WScript.ScriptName & " <oldText> <newText> [/p|path=<docFilesPath>] [/w|wholeword] [/ci|caseinsensitive]"
    WScript.Quit 1
  End If
  
  sOldText = arguments.Unnamed.Item(0)
  sNewText = arguments.Unnamed.Item(1)
  
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
  
  bWholeWord = arguments.Named.Exists("w") Or arguments.Named.Exists("wholeword")
  bCaseInsensitive = arguments.Named.Exists("w") Or arguments.Named.Exists("wholeword")
  
End Sub

'Procedure to edit word document add name and serial number. 
Sub SearchAndRep(searchTerm, replaceTerm, WordApp, WordDoc)
  
  Dim myRange
  Set myRange = WordDoc.Range(0, 0)
  myRange.Find.Execute searchTerm, True, bWholeWord, False,,,,,, replaceTerm,wdReplaceAll

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

'Iterate through files of type docx and call replacement function text
For Each objFile In objFiles

  If UCase(fso.GetExtensionName(objFile.Name)) = "DOCX" Then

	' Instantiate word doc
    Set WordDoc = WordApp.Documents.Open(objFile.Path)
	
	' Focus to doc
	WordApp.Documents(objFile.Name).Activate 'switch to open document
    
	' Do the searh and replace
	SearchAndRep sOldText, sNewText, WordApp, WordDoc
    
	' Save and close
	WordDoc.Save
    WordDoc.Close

  End If

Next

WordApp.Quit