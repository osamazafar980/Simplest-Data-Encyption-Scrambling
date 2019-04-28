Option Explicit
Dim fso 
Dim folder , file , oFile , line, data,fName,decrypt,fext
Dim UName,DeskDir,chunks,ws
Set ws = CreateObject("Wscript.shell")
DeskDir = ws.SpecialFolders("Desktop")
chunks = Split(DeskDir,"\")
UName=chunks(2)

Set fso = CreateObject("Scripting.FileSystemObject")
Set oFile= fso.OpenTextFile(UName&".txt",1,True)
Do Until oFile.AtEndOfStream
line = oFile.ReadLine
data = Split(line,"$$")
file = data(0)
fName = Split(file,"\")
fext = Split(fName(UBound(fName)),".")
folder = data(1)
decrypt=Split(fName(UBound(fName)),".")
fso.MoveFile folder , Left(file,(Len(file)-Len(fName(UBound(fName)))))&decrypt(0)&"."&decrypt(1)
fso.DeleteFolder (Left(folder,(Len(folder)-(Len(decrypt(0))+6))))
Loop
oFile.Close
MsgBox "Done" , vbInformation