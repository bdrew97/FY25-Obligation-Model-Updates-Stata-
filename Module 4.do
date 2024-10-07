Option Explicit
Option Base 1

Function RemoveBackSlash(filename As String)
   RemoveBackSlash = filename
   If Len(filename) = InStrRev(filename, "\", -1) Then
      RemoveBackSlash = Left(filename, Len(filename) - 1)
      
   End If
   
End Function

Function FileOrDirExists(PathName As String) As Boolean
     'Macro Purpose: Function returns TRUE if the specified file or folder exists, false if not.
     'File usage   : Provide full file path and extension
     'Folder usage : Provide full folder path Accepts with/without trailing "\" (Windows)
    
    Dim iTemp As Integer
     
     'Ignore errors to allow for error evaluation
    On Error Resume Next
    iTemp = GetAttr(PathName)
     
     'Check if error exists and set response appropriately
    Select Case Err.Number
    Case Is = 0
        FileOrDirExists = True
    Case Else
        FileOrDirExists = False
    End Select
     'Resume error checking
    On Error GoTo 0
End Function
 
Sub ThreeExecutableExist()
    'Test if directory or file exists
    Dim FileNames(1 To 3) As String
    Dim Description(1 To 3) As String
    Dim count As Integer
     
    FileNames(1) = Range("Stata_Dofile_Path").Value
    FileNames(2) = Range("Stata_Executable").Value
    FileNames(3) = Range("CSC_Path").Value
    
    Description(1) = "Stata Dofile Path"
    Description(2) = "Stata Executable Path"
    Description(3) = "CSC Path"
    
    For count = 1 To 3
      If FileOrDirExists(FileNames(count)) Then
'          MsgBox FileNames(count) & " exists!"
      Else
          MsgBox Description(count) & " does not exist.  Please correct before running Model."
          End
      End If
    Next

End Sub

Sub ThreeExecutableExistParas(Path1 As String, Path2 As String, Path3 As String)
    'Test if directory or file exists
    
    Dim FileNames(1 To 3) As String
    Dim count As Integer
    
    FileNames(1) = Path1  ' Range("Stata_Dofile_Path").Value
    FileNames(2) = Path2  ' Range("Stata_Executable").Value
    FileNames(3) = Path3  ' Range("CSC_Path").Value
    
    For count = 1 To 3
      If FileOrDirExists(FileNames(count)) Then
          MsgBox FileNames(count) & " exists!"
      Else
          MsgBox FileNames(count) & " does not exist."
          End
      End If
    Next
End Sub


 

