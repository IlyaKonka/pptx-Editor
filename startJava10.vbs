Set shell = CreateObject("WScript.Shell")
dim fso: set fso = CreateObject("Scripting.FileSystemObject")
dim CurrentDirectory
CurrentDirectory = fso.GetAbsolutePathName(".")
dim Directory
Directory = CurrentDirectory & "\jars"
shell.CurrentDirectory = Directory
shell.Run "starterJava10.bat" , 0, True