################################################################################
#
# Script to generate a list of recent documents, including full path
# Written by Vaughn Miller
#
#################################################################################

# Change the below path to an appropriate path
objStartFolder = "c:\Documents and Settings\User\Recent"

set WshShell = WScript.CreateObject("Wscript.Shell")
set objFSO = CreateObject("Scripting.FileSystemObject")

set objFolder = objFSO.GetFolder(objStartFolder)

# Build the collection of files
Set colFiles = objFolder.Files

# Loop through the collection and echo out the target path for each .lnk file
For Each objFile in colFiles
  if UCase(objFSO.GetExtensionName(objFile.Name)) = "LNK" Then
		set oShellLink = WshShell.CreateShortcut(objStartFolder & "\" & objFile.Name)
		WScript.Echo oShellLink.TargetPath
	End If
Next

