' Inserted 07/2013, Kayla
Public Sub unprotectAll()
Dim Sh As Worksheet
Dim myPassword As String
    myPassword = "budgetrocks"

    For Each Sh In ActiveWorkbook.Worksheets
    Sh.Unprotect password:=myPassword
Next Sh
End Sub


Sub DFC_Risk_File()
  Dim filename As String
 
 filename = Application.GetOpenFilename("Excel Files (*.xls),*.xls", , "DFC Risk File")

 If filename <> "False" Then
    Range("FY2013_Risk_File").Select
    ActiveCell.FormulaR1C1 = filename
 End If
End Sub

Sub ICRAS_Risk()
    Dim filename As String

    filename = Application.GetOpenFilename("Excel Files (*.xls),*.xls", , "ICRAS Risk File")

    If filename <> "False" Then
    
        Range("ICRAS_Risk_File").Select
        ActiveCell.FormulaR1C1 = filename
 
 End If
End Sub

Sub dofile()
  Dim xRow&, vSF
  Dim Foldername, InitialFoldr$

  InitialFoldr$ = Range("Stata_Dofile_Path").Text '<<< Startup folder to begin searching from

  With Application.FileDialog(msoFileDialogFolderPicker)
      .InitialFileName = Application.DefaultFilePath & "\"
      .Title = "Please Select the DoFile Folder"
      .InitialFileName = InitialFoldr$
      .Show
      If .SelectedItems.count <> 0 Then
          Foldername = .SelectedItems(1) & "\"
      End If
  End With

 Range("Stata_Dofile_Path").Select
 ActiveCell.FormulaR1C1 = Foldername

End Sub



Sub Stata()
  Dim filename As String

 filename = Application.GetOpenFilename("Executable Files (*.exe),*.exe", , "Select Stata Executable")

 If filename <> "False" Then
    
     Range("Stata_Executable").Select
     ActiveCell.FormulaR1C1 = filename
     
 End If

End Sub


Sub Temporary_File_Path()
  Dim xRow&, vSF
  Dim Foldername, InitialFoldr$

  InitialFoldr$ = Range("Temp_Path").Text '<<< Startup folder to begin searching from

  With Application.FileDialog(msoFileDialogFolderPicker)
      .InitialFileName = Application.DefaultFilePath & "\"
      .Title = "Please select a folder to list Files from"
      .InitialFileName = InitialFoldr$
      .Show
      If .SelectedItems.count <> 0 Then
          Foldername = .SelectedItems(1) & "\"
      End If
  End With

 Range("Temp_Path").Select
 ActiveCell.FormulaR1C1 = Foldername

End Sub

Sub CSC_Executable()

  Dim filename As String

 filename = Application.GetOpenFilename("Executable Files (*.exe),*.exe", , "Select SubsidyCLI Executable")

 If filename <> "False" Then
    
     Range("CSC_Path").Select
     ActiveCell.FormulaR1C1 = filename
     
 End If


End Sub

Sub CreateOutputFolder()

 Dim LoanOfficer As String
 Dim OutputFolder As String
 Dim OutputFolder2 As String
 
 LoanOfficer = Trim(Range("Loan_Officer").Text)
 OutputFolder = Range("Stata_Dofile_Path") & "Outputs\" & LoanOfficer
 OutputFolder2 = OutputFolder & "\Subfolder - DI FX Zero-Fi"
 
 Folder = Dir(OutputFolder, vbDirectory)
 Folder2 = Dir(OutputFolder2, vbDirectory)
 
 If Folder = "" Then
    MkDir OutputFolder
 End If
 
 If Folder2 = "" Then
    MkDir OutputFolder2
 End If
 
End Sub

' Inserted 10/2013, Kayla
Sub DeleteOutputs()
    Dim myvar As Variant
    Dim myFolder As String
    Dim OutputFolder As String
    Dim i As Long
    
    'Delete Outputs from Previous Run in UI
    Application.ScreenUpdating = False
    Call unprotectAll
    Call CreateOutputFolder
    On Error Resume Next
    Range("CSC_Output").ClearContents
    Range("LifeTable").ClearContents
    Range("CSC_Input").ClearContents
    Range("Raw_CF").ClearContents
    Range("FX_CF").ClearContents
    Sheets("Cash Flow data").Cells.Clear
    Sheets("NFD").Cells.Clear
    Sheets("Combined").Cells.Clear
    Sheets("Messages").Cells.Clear
    Sheets("Subsidy Calculations").Cells.Clear
    Sheets("Discounts").Cells.Clear
    Sheets("Distributions").Cells.Clear
    Sheets("PV Factors").Cells.Clear
    Sheets("Subsidy").Cells.Clear
    Sheets("Source data").Cells.Clear
    On Error GoTo 0
    Application.ScreenUpdating = True
    'Delete Outputs from Previous Run in Output Folder
    OutputFolder = Range("Stata_Dofile_Path") & "Outputs\" & Range("Loan_Officer")
    myvar = FileList(OutputFolder, "*")
  
    For i = LBound(myvar) To UBound(myvar)
        On Error Resume Next
        Kill OutputFolder & myvar(i)
    Next
End Sub

'Inserted 10/2013, Kayla
 Sub TestFileOpen()
    Dim OutputFolder As String
    Dim Output(1 To 13) As String
    Dim count As Integer
    
    OutputFolder = Range("Stata_Dofile_Path") & "Outputs\" & Range("Loan_Officer")
    Output(1) = OutputFolder & "\cscoutput.xls"
    Output(2) = OutputFolder & "\cscinput.xls"
    Output(3) = OutputFolder & "\cscinput.xlsx"
    Output(4) = OutputFolder & "\CSC Cashflow.xlsx"
    Output(5) = OutputFolder & "\Life Table.xlsx"
    Output(6) = OutputFolder & "DFC Obligation UI Stata Input_DisbursementSchedule.xls"
    Output(7) = OutputFolder & "DFC Obligation UI Stata Input_InterestCalculations.xls"
    Output(8) = OutputFolder & "DFC Obligation UI Stata Input_PrincipalSchedule.xls"
    Output(9) = OutputFolder & "DFC Obligation UI Stata Input_StataInput.xls"
    Output(10) = OutputFolder & "DFC Obligation UI Stata Input_OtherFeesSchedule.xls"
    Output(11) = OutputFolder & "DFC Obligation UI Stata Input_FeeSharingSchedule.xls"
    Output(12) = OutputFolder & "DFC Obligation UI Stata Input_FXInputs.xls"
    Output(13) = OutputFolder & "\CSC Cashflow FX.xlsx"
    
    For count = 1 To 13
       ' Test to see if the file is open.
       If IsFileOpen(Output(count)) Then
           ' Display a message stating the file in use.
           MsgBox Output(count) & " " & "is in use. Please close this file before running the Model."
           End
       End If
       Next
   End Sub
   
   Function FileList(fldr As String, Optional fltr As String = "*.*") As Variant
    Dim sTemp As String, sHldr As String
    If Right$(fldr, 1) <> "\" Then fldr = fldr & "\"
    sTemp = Dir(fldr & fltr)
    If sTemp = "" Then
        FileList = Split("No files found", "|")  'ensures an  array is returned
        Exit Function
    End If
    Do
        sHldr = Dir
        If sHldr = "" Then Exit Do
        sTemp = sTemp & "|" & sHldr
     Loop
    FileList = Split(sTemp, "|")
End Function

   ' This function checks to see if a file is open or not. If the file is
   ' already open, it returns True. If the file is not open, it returns
   ' False. Otherwise, a run-time error will occur because there is
   ' some other problem accessing the file.

   Function IsFileOpen(filename As String)
       Dim filenum As Integer, errnum As Integer

       On Error Resume Next   ' Turn error checking off.
       filenum = FreeFile()   ' Get a free file number.
       ' Attempt to open the file and lock it.
       Open filename For Input Lock Read As #filenum
       Close filenum          ' Close the file.
       errnum = Err           ' Save the error number that occurred.
       On Error GoTo 0        ' Turn error checking back on.

       ' Check to see which error occurred.
       Select Case errnum

           ' No error occurred.
           ' File is NOT already open by another user.
           Case 0
               IsFileOpen = False

           ' Error number for "Permission Denied."
           ' File is already opened by another user.
           Case 70
               IsFileOpen = True

           ' Another error occurred.
   '        Case Else
   '            Error errnum
       End Select
   End Function


