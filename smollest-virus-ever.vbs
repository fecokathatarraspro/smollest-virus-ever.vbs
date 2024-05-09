Dim objFSO, objFile, strDesktopPath, strFileName, strText

' Create a FileSystemObject
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Retrieve the path of the Desktop folder using environment variables
strDesktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop")

' Define the file name
strFileName = "sussy_imposter.txt"

' Combine the desktop path with the file name to get the full file path
strFilePath = objFSO.BuildPath(strDesktopPath, strFileName)

' Define the text you want to write to the file
strText = "sussy"

' Create a new text file on the desktop
Set objFile = objFSO.CreateTextFile(strFilePath)

' Write the text to the file
objFile.Write strText

' Close the file
objFile.Close
