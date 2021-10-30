Sub getfilenameexcel()
Dim fso As Scripting.FilesystemObject
Dim fsofolder As Scripting.Folder
Dim fsofile As Scripting.File
Set fso = CreateObject("scripting.filesystemobject")
Set fsofolder = fso.GetFolder("G:\Phan_loai_Doan_vien_20-21\Data")
ce = 2
For Each fsofile In fsofolder.Files
Range("A" & ce).Value = fsofile.Path
ce = ce + 1
Next fsofile
End Sub
