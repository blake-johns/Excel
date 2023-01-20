

```vba
Function FindAllFiles(path, Optional FileType = "*", Optional SubFolders = True) As Variant
    'path: string path to folder you want to seach
    'FileType: regex pattern for type of files you want (ex. "\.xlsm$" for only files ending with .xlsm)
    'SubFolders: expecting true or false, sets recursion to search subfolders for files
    'note: requires FindFiles function to run
    
    'create arraylist to pass files between function runs
    Set FileStorage = CreateObject("System.Collections.ArrayList")
    
    'run main function passing in arraylist, path, filetype, and whether you want subfolders
    FilesFound = FindFiles(FileStorage, path, FileType, SubFolders)
    
    'create array from arraylist
    FindAllFiles = FileStorage.toarray()
    
End Function

Function FindFiles(arr, path, Optional FileType = "*", Optional SubFolders = True)
    'arr: outside arraylist passed to allow additional files each recursion pass
    'path: string path to folder you want to seach
    'FileType: regex pattern for type of files you want (ex. "\.xlsm$" for only files ending with .xlsm)
    'SubFolders: expecting true or false, sets recursion to search subfolders for files
    'note: needs to be ran from FindAllFiles function
    
    Dim re As Object
    Dim objFSO, objFile, objFolder, objSubFolder As Object
    Dim sFile, sFolder, SubFiles As Object
    
    'set up objects needed for function below
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set re = CreateObject("VBScript.RegExp")
    Set objFolder = objFSO.GetFolder(path)
    
    'set up regex pattern matching for FileType
    With re
        .IgnoreCase = True
        .Global = True
        .Pattern = FileType
    End With
    
    'search current directory and add matching files to ArrayList(arr)
    For Each objFile In objFolder.Files
        sFile = objFile.path
        If re.test(sFile) And (Not sFile Like "*~$*") Then
            arr.Add sFile
        End If
    Next
    
    'search sub directories, recursively running function for each subfolder
    If SubFolders Then
        For Each objSubFolder In objFolder.SubFolders
            sFolder = objSubFolder.path
            SubFiles = FindFiles(arr, sFolder, FileType, True)
        Next
    End If
        
End Function
```
