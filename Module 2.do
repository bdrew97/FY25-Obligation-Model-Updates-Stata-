#If VBA7 Then
    Declare PtrSafe Function OpenProcess Lib "kernel32" _
                             (ByVal dwDesiredAccess As LongPtr, _
                              ByVal bInheritHandle As LongPtr, _
                              ByVal dwProcessId As Long) As LongPtr
    
    Declare PtrSafe Function GetExitCodeProcess Lib "kernel32" _
                                    (ByVal hProcess As LongPtr, _
                                     lpExitCode As Long) As LongPtr
#Else
    Declare Function OpenProcess Lib "kernel32" _
                                 (ByVal dwDesiredAccess As Long, _
                                  ByVal bInheritHandle As Long, _
                                  ByVal dwProcessId As Long) As Long
    
    Declare Function GetExitCodeProcess Lib "kernel32" _
                                        (ByVal hProcess As Long, _
                                         lpExitCode As Long) As Long

#End If

Public Const PROCESS_QUERY_INFORMATION = &H400
Public Const STILL_ACTIVE = &H103
Public ExitCode As Long

Option Base 1
Public AutoLoop As Integer
Public AutoLoop2 As Integer
Public FinSubsidy(1 To 20) As String
Public TotSubsidy(1 To 20) As String
Public RandInd As String
Dim FinSub As String
Dim TotSub As String

' Inserted 07/2013, Kayla

Public Sub ShellAndWait(ByVal PathName As String, Optional WindowState)
' Declare PtrSafe Function OpenProcess Lib "kernel64" (ByVal dwDesiredAcess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long

    Dim hProg As Long
    #If VBA7 Then
        Dim hProcess As LongPtr
    #Else
        Dim hProcess As Long
    #End If
    
    'fill in the missing parameter and execute the program
    If IsMissing(WindowState) Then WindowState = 1
    

    '*****************Runs STATA*********************
    hProg = Shell(PathName, WindowState)
    'hProg is a "process ID under Win32. To get the process handle:
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, hProg)
    
    Do
        'populate Exitcode variable
        GetExitCodeProcess hProcess, ExitCode
        DoEvents
    Loop While ExitCode = STILL_ACTIVE
    
End Sub
Sub Run()
Dim i As Integer
Dim Col As Integer
Dim Premium_cmd As String
Dim PRE As String
Dim dofile As String
Dim prem_out As String
Dim IRR_Path As String
Dim IRR_Adj_Path As String
Dim LastColumn As Integer



If Sheets("Inputs").Range("Run_Option") = "Single Run" Then
' Input Validation
     With Application.ActiveWorkbook.Sheets("Validation Checklist")
        Col = 3
    
        For i = 11 To 212
         If .Cells(i, Col).Text = "FALSE" Then
             MsgBox "Row " & .Cells(i, Col - 2) & " " & .Cells(i, Col + 1)
             Exit Sub
         End If
        Next
     End With

    If Sheets("Inputs").Range("Store_Option") = "Final Estimate" Then
        Call FinalPassword
    End If
     
'Ensure IR Factor is 1 if not zero-financed
Sheets("StataInput").Unprotect password:="budgetrocks"
Sheets("StataInput").Range("IRFactor") = 1
Sheets("StataInput").Protect password:="budgetrocks"
    
' Call Single Run
    Call automation
    
     ' If Range("FX_Denomination").Value = "Yes" Then
     '       Premium_cmd = Range("Stata_Executable").Text
     '       DoFilePath = RemoveBackSlash(Range("Stata_Dofile_Path").Text)
     '       dofile = RemoveBackSlash(Range("Stata_Dofile_Path").Text) & "\DoFiles\v1_2024_FinanceModel\Premium Calculation.do"
     '       Premium_cmd = """" & Premium_cmd & """"
     '       dofile = """" & dofile & """"
     '       DoFilePath = """" & DoFilePath & """"
     '       Premium_cmd = Premium_cmd & " do " & " " & dofile & " " & DoFilePath
     '       ShellAndWait (Premium_cmd)
     '       prem_out = RemoveBackSlash(Range("Stata_Dofile_Path").Text)
     '       prem_out = prem_out & "\Outputs\FX Premium.xls"
     '       Workbooks.Open prem_out
     '       Range("A2").Copy
     '       ActiveWorkbook.Close
     '       Call unprotectAll
     '       ActiveWorkbook.Sheets("Cash Flow - Raw").Select
     '       Range("A1").PasteSpecial
     '       Range("E37") = Range("E37") - Range("A1")
     '       Range("E37").Copy
     '       ActiveWorkbook.Sheets("Cash Flow - To CSC").Select
     '       Range("E37").PasteSpecial
     '       Call ProtectSheets
     '       Call run_csc
     '   End If

' Call Save Results
    Call get_csc_subsidytab
    Call ProtectSheets
    If Sheets("Inputs").Range("Store_Option") = "Final Estimate" Then
        Call SaveResults
    End If
    Sheets("Inputs").Select
    Sheets("Inputs").Range("CSC_Output").Select
    
ElseIf Sheets("Inputs").Range("Run_Option") = "Zero Financing (Interest Goal Seek)" Then
' Input Validation
     With Application.ActiveWorkbook.Sheets("Validation Checklist")
        Col = 3
    
        For i = 11 To 212
         If .Cells(i, Col).Text = "FALSE" Then
             MsgBox "Row " & .Cells(i, Col - 2) & " is FALSE: " & .Cells(i, Col + 1)
             Exit Sub
         End If
        Next
     End With
     
    If Sheets("Inputs").Range("Store_Option") = "Final Estimate" Then
        Call FinalPassword
    End If

' Call Zero Financing Loop
    Call IRHelper
' Call Save Results
    Application.GoTo Reference:="CSC_Output"
    FinSub = Range("CSC_Output").Cells(2).Value
    If Abs(FinSub) <= 0.00499999 Then
     '   If Range("FX_Denomination").Value = "Yes" Then
     '       Premium_cmd = Range("Stata_Executable").Text
     '       DoFilePath = RemoveBackSlash(Range("Stata_Dofile_Path").Text)
     '       dofile = RemoveBackSlash(Range("Stata_Dofile_Path").Text) & "\DoFiles\v1_2024_FinanceModel\Premium Calculation.do"
     '       Premium_cmd = """" & Premium_cmd & """"
     '       dofile = """" & dofile & """"
     '       DoFilePath = """" & DoFilePath & """"
     '       Premium_cmd = Premium_cmd & " do " & " " & dofile & " " & DoFilePath
     '       ShellAndWait (Premium_cmd)
     '       prem_out = RemoveBackSlash(Range("Stata_Dofile_Path").Text)
     '       prem_out = prem_out & "\Outputs\FX Premium.xls"
     '       Workbooks.Open prem_out
     '       Range("A2").Copy
     '       ActiveWorkbook.Close
     '       Call unprotectAll
     '       ActiveWorkbook.Sheets("Cash Flow - Raw").Select
     '       Range("A1").PasteSpecial
     '       Range("E34") = Range("E34") - Range("A1")
     '       Range("E34").Copy
     '       ActiveWorkbook.Sheets("Cash Flow - To CSC").Select
     '       Range("E34").PasteSpecial
     '       Call ProtectSheets
     '       Call run_csc
     '   End If
        
        Call get_csc_subsidytab
        Call ProtectSheets
        If Sheets("Inputs").Range("Store_Option") = "Final Estimate" Then
            Call SaveResults
        End If
        Sheets("Inputs").Select
        Sheets("Inputs").Range("CSC_Output").Select
    End If
    
ElseIf Sheets("Inputs").Range("Run_Option") = "Zero Total Subsidy (Fee Goal Seek)" Then
' Input Validation
    'Validation 1: Zero Sub Goal Seek cannot be used with All-In Single Rate
    If Application.ActiveWorkbook.Sheets("Inputs").Range("Int_Type") = "All-In Single Rate" And Range("Direct_Guar").Value = "Direct Loan" Then
        MsgBox "Zero Total Subsidy (Fee Goal Seek) cannot be applied to Direct Loans with an All-In Single Rate. Please run the Model using 'Single Run' or 'Zero Financing (Interest Goal Seek)'"
        Exit Sub
    ElseIf Application.ActiveWorkbook.Sheets("Inputs").Range("Int_Type") = "All-In Single Rate" And Range("Direct_Guar").Value = "DPA" Then
        MsgBox "Zero Total Subsidy (Fee Goal Seek) cannot be applied to Direct Loans with an All-In Single Rate. Please run the Model using 'Single Run' or 'Zero Financing (Interest Goal Seek)'"
        Exit Sub
    End If
    'Validation 2: Initial Test Value for Pre/Post Completion Fees cannot be 0.
    Const ZeroFeeRate As Double = 0.000999999
    If Application.ActiveWorkbook.Sheets("Inputs").Range("Direct_Guar").Value <> "LPG" And Application.ActiveWorkbook.Sheets("Inputs").Range("Direct_Guar").Value <> "Non-LPG" Then
        If Application.ActiveWorkbook.Sheets("Inputs").Range("fee_precomp").Value < ZeroFeeRate And Application.ActiveWorkbook.Sheets("Inputs").Range("fee_postcomp").Value < ZeroFeeRate Then
            MsgBox "Input a value for Pre & Post Completion Fee before running Zero Total Subsidy (Fee Goal Seek). The initial value for Pre & Post Completion Fee cannot be 0."
            Exit Sub
        ElseIf Application.ActiveWorkbook.Sheets("Inputs").Range("fee_postcomp").Value < ZeroFeeRate And Application.ActiveWorkbook.Sheets("Inputs").Range("Completion_Pt").Value = 1 Then
            MsgBox "Input a value for Post Completion Fee before running Zero Total Subsidy (Fee Goal Seek). The initial value for Post Completion Fee cannot be 0 if Completion Point is equal to 1."
            Exit Sub
        End If
    ElseIf Application.ActiveWorkbook.Sheets("Inputs").Range("Direct_Guar").Value = "LPG" Or Application.ActiveWorkbook.Sheets("Inputs").Range("Direct_Guar").Value = "Non-LPG" Then
        If Application.ActiveWorkbook.Sheets("Inputs").Range("Fee_DCA").Value < ZeroFeeRate Then
            MsgBox "Input a value for Annual Flat Utilization Fee before running Zero Total Subsidy (Fee Goal Seek). The initial value for Annual Flat Utilization Fee cannot be 0"
            Exit Sub
        End If
    End If
    
    'Validation 3: Zero Sub Goal Seek cannot be applied to Custom Fees (i.e. "Other Subsidy Fees")
    If Application.ActiveWorkbook.Sheets("Inputs").Range("OtherSubFees") = "Yes" Then
        MsgBox "Zero Total Subsidy (Fee Goal Seek) cannot be used for projects with 'Other Subsidy Fees'. Please run the Model using either 'Single Run' or 'Zero Financing (Interest Goal Seek)'"
        Exit Sub
    End If
    'Validation Checklist for all other inputs
     With Application.ActiveWorkbook.Sheets("Validation Checklist")
        Col = 3
    
        For i = 11 To 212
         If .Cells(i, Col).Text = "FALSE" Then
             MsgBox "Row " & .Cells(i, Col - 2) & " is FALSE: " & .Cells(i, Col + 1)
             Exit Sub
         End If
        Next
     End With
     
    If Sheets("Inputs").Range("Store_Option") = "Final Estimate" Then
        Call FinalPassword
    End If
'Call Zero Total Subsidy Loop
    Call ZeroTotSub
    
   ' If Range("FX_Denomination").Value = "Yes" Then
   '         Premium_cmd = Range("Stata_Executable").Text
   '         DoFilePath = RemoveBackSlash(Range("Stata_Dofile_Path").Text)
   '         dofile = RemoveBackSlash(Range("Stata_Dofile_Path").Text) & "\DoFiles\v1_2024_FinanceModel\Premium Calculation.do"
   '         Premium_cmd = """" & Premium_cmd & """"
   '         dofile = """" & dofile & """"
   '         DoFilePath = """" & DoFilePath & """"
   '         Premium_cmd = Premium_cmd & " do " & " " & dofile & " " & DoFilePath
   '         ShellAndWait (Premium_cmd)
   '         prem_out = RemoveBackSlash(Range("Stata_Dofile_Path").Text)
   '         prem_out = prem_out & "\Outputs\FX Premium.xls"
   '         Workbooks.Open prem_out
   '         Range("A2").Copy
   '         ActiveWorkbook.Close
   '         Call unprotectAll
   '         ActiveWorkbook.Sheets("Cash Flow - Raw").Select
   '         Range("A1").PasteSpecial
   '         Range("E37") = Range("E37") - Range("A1")
   '         Range("E37").Copy
   '         ActiveWorkbook.Sheets("Cash Flow - To CSC").Select
   '         Range("E37").PasteSpecial
   '         Call ProtectSheets
   '         Call run_csc
   '     End If
        
'Call Save Results
    Application.GoTo Reference:="CSC_Output"
    TotSub = Range("CSC_Output").Cells(1).Value
    If Abs(TotSub) <= 0.00499999 Then
        Call get_csc_subsidytab
        Call ProtectSheets
        If Sheets("Inputs").Range("Store_Option") = "Final Estimate" Then
            Call SaveResults
        End If
        Sheets("Inputs").Select
        Sheets("Inputs").Range("CSC_Output").Select
    End If
End If

'IRR Calculation - Commented out until we figure out better "guess" for the formula

Call unprotectAll

IRR_Path = RemoveBackSlash(Range("Stata_Dofile_Path").Text) & "\Outputs\" & Range("Loan_Officer").Text & "\IRR Calc.xlsx"
IRR_Adj_Path = RemoveBackSlash(Range("Stata_Dofile_Path").Text) & "\Outputs\" & Range("Loan_Officer").Text & "\Risk Adj IRR Calc.xlsx"

ActiveWorkbook.Sheets("IRR Calculations").Visible = True
ActiveWorkbook.Sheets("IRR Calculations").Select
Worksheets("IRR Calculations").Rows("5:6").ClearContents
Worksheets("IRR Calculations").Rows("9:10").ClearContents

Workbooks.Open IRR_Path
LastColumn = Cells(2, Columns.count).End(xlToLeft).Column
Range(Cells(1, 1), Cells(2, LastColumn)).Select
Selection.Copy
ActiveWorkbook.Close
ActiveWorkbook.Sheets("IRR Calculations").Select
Range("A5").PasteSpecial

Workbooks.Open IRR_Adj_Path
Range(Cells(1, 1), Cells(2, LastColumn)).Select
Selection.Copy
ActiveWorkbook.Close
ActiveWorkbook.Sheets("IRR Calculations").Select
Range("A9").PasteSpecial
Sheets("IRR Calculations").Visible = False
Sheets("Inputs").Select
Call ProtectSheets

End Sub

Sub RunStata()

  Dim rv As Variant
  Dim cmd_line As String
  Dim dtafile As String
  Dim dofile As String
  Dim i As Integer
  Dim TempSave As String
  Dim OutputPath As String
  Dim PPS As String
  Dim OSF As String
  Dim FSS As String
  Dim FXC As String
  Dim CurrencyType As String
  Dim LoanType As String
  Dim LoanOfficer As String
  
  ThreeExecutableExist
  ' Inserted 10/2013, Kayla
  DeleteOutputs
  TestFileOpen
  
  Calculate
  ActiveWorkbook.Save
  
  xlsmname = ActiveWorkbook.Name
  xlsmfullname = ActiveWorkbook.FullName
  
  Application.ScreenUpdating = False
  
' Export UI data to Stata and calculate necessary monthly figures
  cmd_line = Range("Stata_Executable").Text
    
'  FY2013RiskPath = RemoveBackSlash(Range("FY2013_Risk_File").Text)
'  ICRASPath = RemoveBackSlash(Range("ICRAS_Risk_File").Text)
  DoFilePath = RemoveBackSlash(Range("Stata_Dofile_Path").Text)

' Brandon edit - Loan Type variable
  LoanType = Trim(Range("Direct_Guar").Text)
  LoanOfficer = Trim(Range("Loan_Officer").Text)

  dofile = RemoveBackSlash(Range("Stata_Dofile_Path").Text) & "\DoFiles\v1_2024_FinanceModel\DFC Obligation Master Dofile.do"
  'dofile = RemoveBackSlash(Range("Stata_Dofile_Path").Text) & "\DoFiles\v4_Merge_FX\DFC Obligation Master Dofile.do"

  PPS = RemoveBackSlash(Range("Prin_Payment_Structure").Text)
  OSF = RemoveBackSlash(Range("OtherSubFees").Text)
  
  'If Range("Direct_Guar") <> "LPG" And Range("Direct_Guar") <> "Non-LPG" And Range("Direct_Guar") <> "Direct Loan" Then
  If Range("Direct_Guar") = "Investment Guaranty" Or Range("Direct_Guar") = "Direct Loan" Or Range("Direct_Guar") = "DPA" Then
    FSS = RemoveBackSlash(Range("FeeSharing").Text)
  Else
    FSS = "No"
  End If
  
  FXC = RemoveBackSlash(Range("FX_Denomination").Text)
  
  If Range("Currency") <> "" Then
    CurrencyType = RemoveBackSlash(Range("Currency").Text)
  End If
  
  If FXC = "No" Then
    CurrencyType = "USD"
  End If

  cmd_line = """" & cmd_line & """"
  FY2013RiskPath = """" & FY2013RiskPath & """"
  ICRASPath = """" & ICRASPath & """"
  dofile = """" & dofile & """"
    
' Saves the workbook as a macro
'
' Divide up Sheets that may be needed and individually save to .xls format
    Application.DisplayAlerts = False
    Call unprotectAll

  ' 1) StataInput Sheet
    Sheets("StataInput").Visible = True
    Sheets("StataInput").Copy
    Range("A1:CJ3").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    TempSave = DoFilePath & "\Outputs\" & LoanOfficer & "\DFC Obligation UI Stata Input_StataInput.xls"
    Application.CutCopyMode = False
    ActiveWorkbook.SaveAs filename:=TempSave, FileFormat:= _
        xlExcel8, CreateBackup:=False
    ActiveWorkbook.Close
    
  ' 2) Disbursement Schedule Sheet
    Application.ScreenUpdating = True
    
    Sheets("StataInput").Visible = False
    Set wb = ThisWorkbook
    Set NewWorkbook = Workbooks.Add
    wb.Worksheets("Disbursement Schedule").Range("A11:G111").Copy
    NewWorkbook.Sheets(1).Range("A11").PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    TempSave = DoFilePath & "\Outputs\" & LoanOfficer & "\DFC Obligation UI Stata Input_DisbursementSchedule.xls"
    Application.CutCopyMode = False
        NewWorkbook.SaveAs filename:=TempSave, FileFormat:= _
        xlExcel8, CreateBackup:=False
    NewWorkbook.Close
    
  ' 3) Interest Calculation Sheet
    
    Set wb = ThisWorkbook
    Set NewWorkbook = Workbooks.Add
    wb.Worksheets("Interest Calculations").Range("A1:X400").Copy
    NewWorkbook.Sheets(1).Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    TempSave = DoFilePath & "\Outputs\" & LoanOfficer & "\DFC Obligation UI Stata Input_InterestCalculations.xls"
    Application.CutCopyMode = False
        NewWorkbook.SaveAs filename:=TempSave, FileFormat:= _
        xlExcel8, CreateBackup:=False
    NewWorkbook.Close

    'Windows(xlsmname).Activate
    'Sheets("Interest Calculations").Visible = False
    
    Application.ScreenUpdating = False
    
  ' 4) Principal Schedule Sheet
    If Range("Prin_Payment_Structure").Value = "Custom" Then
        Sheets("Principal Schedule").Select
        Range("A4:AF372").Copy
        Set NewWorkbook = Workbooks.Add
        Range("A4:AF372").Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        TempSave = DoFilePath & "\Outputs\" & LoanOfficer & "\DFC Obligation UI Stata Input_PrincipalSchedule.xls"
        Application.CutCopyMode = False
        ActiveWorkbook.SaveAs filename:=TempSave, FileFormat:= _
            xlExcel8, CreateBackup:=False
        ActiveWorkbook.Close
    End If
    
' 5) Other Fees Schedule Sheet
    If Range("OtherSubFees").Value = "Yes" Then
        Sheets("Other Fees Schedule").Select
        Range("A4:AF394").Copy
        Set NewWorkbook = Workbooks.Add
        Range("A4:AF394").Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        TempSave = DoFilePath & "\Outputs\" & LoanOfficer & "\DFC Obligation UI Stata Input_OtherFeesSchedule.xls"
        Application.CutCopyMode = False
        ActiveWorkbook.SaveAs filename:=TempSave, FileFormat:= _
            xlExcel8, CreateBackup:=False
        ActiveWorkbook.Close
    End If
    
' 6) Fee Sharing Schedule Sheet
    If Range("FeeSharing").Value = "Yes" Then
        Sheets("Fee Sharing Schedule").Visible = True
        Sheets("Fee Sharing Schedule").Select
        Range("A4:AF394").Copy
        Set NewWorkbook = Workbooks.Add
        Range("A4:AF394").Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        TempSave = DoFilePath & "\Outputs\" & LoanOfficer & "\DFC Obligation UI Stata Input_FeeSharingSchedule.xls"
        Application.CutCopyMode = False
        ActiveWorkbook.SaveAs filename:=TempSave, FileFormat:= _
            xlExcel8, CreateBackup:=False
        ActiveWorkbook.Close
        Sheets("Fee Sharing Schedule").Visible = False
    End If
    
' 7) FX Inputs Sheet
    If Range("FX_Denomination").Value = "Yes" Then
        Sheets("FX Inputs").Select
        Range("A4:AF394").Copy
        Set NewWorkbook = Workbooks.Add
        Range("A4:AF394").Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        TempSave = DoFilePath & "\Outputs\" & LoanOfficer & "\DFC Obligation UI Stata Input_FXInputs.xls"
        Application.CutCopyMode = False
        ActiveWorkbook.SaveAs filename:=TempSave, FileFormat:= _
            xlExcel8, CreateBackup:=False
        ActiveWorkbook.Close
    End If
        
    Application.DisplayAlerts = True
   
' Command line to open Stata and run the dofile
  DoFilePath = """" & DoFilePath & """"
  'Brandon edit
  LoanOfficer = """" & LoanOfficer & """"
  cmd_line = cmd_line & " do " & " " & dofile & " " & DoFilePath & " " & PPS & " " & "v1_2024_FinanceModel" & " " & OSF & " " & FSS & " " & FXC & " " & CurrencyType & " " & LoanOfficer
  'cmd_line = cmd_line & " do " & " " & dofile & " " & DoFilePath & " " & PPS & " " & "v1_2024_FinanceModel" & " " & OSF & " " & FSS & " " & FXC & " " & CurrencyType

  'rv = Shell(cmd_line)
  ShellAndWait (cmd_line)
    
' Save as original File - xlsx format

   Application.DisplayAlerts = False
      ActiveWorkbook.SaveAs filename:=xlsmfullname, FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
   Application.DisplayAlerts = Truexlsmfullname = ActiveWorkbook.FullName
   
End Sub
Sub SaveResults()

  Dim rv As Variant
  Dim cmd_line As String
  Dim dtafile As String
  Dim dofile As String
  Dim i As Integer
  Dim TempSave As String
  Dim OutputPath As String
  Dim xlsmfullname2 As String
  Dim modProjID As String
    'Variables for creating Index #
  Dim Low As Double
  Dim High As Double
  
  Calculate
  ActiveWorkbook.Save
  
  On Error GoTo ErrorHandler
  
  xlsmname = ActiveWorkbook.Name
  xlsmname2 = """" & xlsmname & """"
  xlsmfullname = ActiveWorkbook.FullName
  xlsmfullname2 = """" & xlsmfullname & """"
    
  Application.ScreenUpdating = False
  
' Export UI data to Stata and calculate necessary monthly figures
  cmd_line = Range("Stata_Executable").Text
    
'  FY2013RiskPath = RemoveBackSlash(Range("FY2013_Risk_File").Text)
'  ICRASPath = RemoveBackSlash(Range("ICRAS_Risk_File").Text)
  DoFilePath = RemoveBackSlash(Range("Stata_Dofile_Path").Text)
' Brandon edit - Loan Type variable
  LoanType = Range("Direct_Guar").Text
  LoanOfficer = Range("Loan_Officer").Text
  
  dofile = RemoveBackSlash(Range("Stata_Dofile_Path").Text) & "\DoFiles\v1_2024_FinanceModel\Save Results.do"
'dofile = RemoveBackSlash(Range("Stata_Dofile_Path").Text) & "\DoFiles\v4_Merge_FX\Save Results.do"

  cmd_line = """" & cmd_line & """"
  FY2013RiskPath = """" & FY2013RiskPath & """"
  ICRASPath = """" & ICRASPath & """"
  dofile = """" & dofile & """"

' Saves the workbook as a macro
'
' Before dividing up sheets to save, generate a Random Index # to identify the run
    Low = 1
    High = 1000000
    Randomize Timer
    R = Int((High - Low + 1) * Rnd() + Low)
'    MsgBox R
    modProjID = Replace(Range("desc_idproject").Value, " ", "-")
    RandInd = modProjID + "_" + CStr(R)
'    MsgBox "The RunID is " & RandInd
    
' Divide up & save inputs Sheets in .xls format used by Stata
    Application.DisplayAlerts = False
    Call unprotectAll
    
  ' 1) StataInput Sheet
    Sheets("StataInput").Visible = True
    Sheets("StataInput").Copy
    Range("A1:CL3").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    TempSave = DoFilePath & "\Outputs\" & LoanOfficer & "\DFC Obligation UI Stata Input_StataInput.xls"
    Kill TempSave
    Application.CutCopyMode = False
    ActiveWorkbook.SaveAs filename:=TempSave, FileFormat:= _
        xlExcel8, CreateBackup:=False
    ActiveWorkbook.Close
        'Also Save in an Batch Obligation Inputs Folder
        Application.ScreenUpdating = True
        Windows(xlsmname).Activate
        Application.ScreenUpdating = False
        Sheets("StataInput").Copy
        Range("A1:CH3").Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
'        TempSave = "S:\Summit Consulting\Obligation Model Batch Run Inputs\" & RandInd & " " & "DFC Obligation UI Stata Input_StataInput.xls"
        TempSave = "C:\temp\Obligation Model Batch Run Inputs\" & RandInd & " " & "DFC Obligation UI Stata Input_StataInput.xls"
        Application.CutCopyMode = False
            ActiveWorkbook.SaveAs filename:=TempSave, FileFormat:= _
            xlExcel8, CreateBackup:=False
        ActiveWorkbook.Close
        Sheets("StataInput").Visible = False
        
  ' 2) Disbursement Schedule Sheet, Save in Batch Obligation Inputs Folder
        Application.ScreenUpdating = True
        Windows(xlsmname).Activate
        Application.ScreenUpdating = False
        Sheets("Disbursement Schedule").Select
        Range("A11:H111").Copy
        Set NewWorkbook = Workbooks.Add
        Range("A11:H111").Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
'        TempSave = "S:\Summit Consulting\Obligation Model Batch Run Inputs\" & RandInd & " " & "DFC Obligation UI Stata Input_DisbursementSchedule.xls"
        TempSave = "C:\temp\Obligation Model Batch Run Inputs\" & RandInd & " " & "DFC Obligation UI Stata Input_DisbursementSchedule.xls"
        Application.CutCopyMode = False
            ActiveWorkbook.SaveAs filename:=TempSave, FileFormat:= _
            xlExcel8, CreateBackup:=False
        ActiveWorkbook.Close
        
  ' 3) Interest Calculation Sheet, Save in Batch Obligation Inputs Folder
        Application.ScreenUpdating = True
        Windows(xlsmname).Activate
        Application.ScreenUpdating = False
        Sheets("Interest Calculations").Visible = True
        Sheets("Interest Calculations").Select
        Range("A1:W400").Copy
        Set NewWorkbook = Workbooks.Add
        Range("A1:W400").Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
'        TempSave = "S:\Summit Consulting\Obligation Model Batch Run Inputs\" & RandInd & " " & "DFC Obligation UI Stata Input_InterestCalculations.xls"
        TempSave = "C:\temp\Obligation Model Batch Run Inputs\" & RandInd & " " & "DFC Obligation UI Stata Input_InterestCalculations.xls"
        
        Application.CutCopyMode = False
            ActiveWorkbook.SaveAs filename:=TempSave, FileFormat:= _
            xlExcel8, CreateBackup:=False
        ActiveWorkbook.Close
        
        Application.ScreenUpdating = True
        Windows(xlsmname).Activate
        Application.ScreenUpdating = False
        Sheets("Interest Calculations").Visible = False
        
  ' 4) Principal Schedule Sheet, Save in Batch Obligation Inputs Folder
    If Range("Prin_Payment_Structure").Value = "Custom" Then
            Application.ScreenUpdating = True
            Windows(xlsmname).Activate
            Application.ScreenUpdating = False
            Sheets("Principal Schedule").Select
            Range("A4:AF372").Copy
            Set NewWorkbook = Workbooks.Add
            Range("A4:AF372").Select
            Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
'            TempSave = "S:\Summit Consulting\Obligation Model Batch Run Inputs\" & RandInd & " " & "DFC Obligation UI Stata Input_PrincipalSchedule.xls"
            TempSave = "C:\temp\Obligation Model Batch Run Inputs\" & RandInd & " " & "DFC Obligation UI Stata Input_PrincipalSchedule.xls"
            Application.CutCopyMode = False
            ActiveWorkbook.SaveAs filename:=TempSave, FileFormat:= _
                xlExcel8, CreateBackup:=False
            ActiveWorkbook.Close
            Application.ScreenUpdating = True
            Windows(xlsmname).Activate
            Application.ScreenUpdating = False
    End If
' 5) Other Fees Schedule Sheet, Save in Batch Obligation Inputs Folder
    If Range("OtherSubFees").Value = "Yes" Then
            Application.ScreenUpdating = True
            Windows(xlsmname).Activate
            Application.ScreenUpdating = False
            Sheets("Other Fees Schedule").Select
            Range("A4:AF394").Copy
            Set NewWorkbook = Workbooks.Add
            Range("A4:AF394").Select
            Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
'            TempSave = "S:\Summit Consulting\Obligation Model Batch Run Inputs\" & RandInd & " " & "DFC Obligation UI Stata Input_PrincipalSchedule.xls"
            TempSave = "C:\temp\Obligation Model Batch Run Inputs\" & RandInd & " " & "DFC Obligation UI Stata Input_OtherFeesSchedule.xls"
            Application.CutCopyMode = False
            ActiveWorkbook.SaveAs filename:=TempSave, FileFormat:= _
                xlExcel8, CreateBackup:=False
            ActiveWorkbook.Close
            Application.ScreenUpdating = True
            Windows(xlsmname).Activate
            Application.ScreenUpdating = False
    End If
    
' 6) Fee Sharing Schedule Sheet, Save in Batch Obligation Inputs Folder
    If Range("FeeSharing").Value = "Yes" Then
            Application.ScreenUpdating = True
            Windows(xlsmname).Activate
            Application.ScreenUpdating = False
            Sheets("Fee Sharing Schedule").Select
            Range("A4:AF394").Copy
            Set NewWorkbook = Workbooks.Add
            Range("A4:AF394").Select
            Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
'            TempSave = "S:\Summit Consulting\Obligation Model Batch Run Inputs\" & RandInd & " " & "DFC Obligation UI Stata Input_PrincipalSchedule.xls"
            TempSave = "C:\temp\Obligation Model Batch Run Inputs\" & RandInd & " " & "DFC Obligation UI Stata Input_FeeSharingSchedule.xls"
            Application.CutCopyMode = False
            ActiveWorkbook.SaveAs filename:=TempSave, FileFormat:= _
                xlExcel8, CreateBackup:=False
            ActiveWorkbook.Close
            Application.ScreenUpdating = True
            Windows(xlsmname).Activate
            Application.ScreenUpdating = False
    End If
    
' 7) FX Inputs Sheet, Save in Batch Obligation Inputs Folder
    If Range("FX_Denomination").Value = "Yes" Then
        Application.ScreenUpdating = True
        Windows(xlsmname).Activate
        Application.ScreenUpdating = False
        Sheets("FX Inputs").Select
        Range("A4:AF394").Copy
        Set NewWorkbook = Workbooks.Add
        Range("A4:AF394").Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        TempSave = "C:\temp\Obligation Model Batch Run Inputs\" & RandInd & " " & "DFC Obligation UI Stata Input_FXInputs.xls"
        Application.CutCopyMode = False
        ActiveWorkbook.SaveAs filename:=TempSave, FileFormat:= _
            xlExcel8, CreateBackup:=False
        ActiveWorkbook.Close
        Application.ScreenUpdating = True
        Windows(xlsmname).Activate
        Application.ScreenUpdating = False
    End If
    
    Application.DisplayAlerts = True
    
' Error Handler
ErrorHandler:
 Select Case Err.Number
    Case Is = 1004
        ' File Path for saving inputs does not exist. This is not essential for subsidy estimation, so ignore Error and Resume execution.
        Resume Next
    Case Is = 0
        ' Do Nothing - No error
    Case Else
        MsgBox Err.Number & ": " & Err.Description
        End
 End Select
 'Resume Error Checking
 On Error GoTo 0
    
' Command line to open Stata and run the Save Results dofile
  DoFilePath = """" & DoFilePath & """"
  LoanOfficer = """" & LoanOfficer & """"
  cmd_line = cmd_line & " do " & " " & dofile & " " & DoFilePath & " " & xlsmfullname2 & " " & xlsmname2 & " " & RandInd & " " & LoanOfficer
  'rv = Shell(cmd_line)
  ShellAndWait (cmd_line)
    
' Save as original File - xlsx format

   Application.DisplayAlerts = False
      ActiveWorkbook.SaveAs filename:=xlsmfullname, FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
   Application.DisplayAlerts = Truexlsmfullname = ActiveWorkbook.FullName
   
   Call ProtectSheets
   
End Sub


Sub automation()
 Dim count As Integer
 
 Application.ScreenUpdating = False
 For count = 1 To 1
  AutoLoop = count
  'Run STATA
  RunStata
  ' Application.Wait Now + #12:00:03 AM#
  
  'Run GetStataOutput
  GetStataOutput
  
  'Run CSC calculator and save the result in the TestRangeRateStart range
  run_csc
 
 Next

Application.ScreenUpdating = True

End Sub


Sub IRHelper()
 Dim IRFactor(1 To 20) As Double, JStep(1 To 20) As Double
 
 Const SecondFactorP As Double = 1.1 ' or 1.25
 Const SecondFactorN As Double = 0.9
 Const ZeroFinancingRate As Double = 0.00499999
 
 AutoLoop = 1
 AutoLoop2 = 1
 Application.ScreenUpdating = False
  
 Do
    ' calculate the value of next Factor
   If AutoLoop = 1 Then
        IRFactor(1) = 1
        JStep(1) = 1
   Else
     If AutoLoop = 2 Then
        ' The second factor is decided as following. If the financial subsidy is
        ' greater than 0, the IRFactor will be assigned as either 1.1 or 1.25
        ' Otherwise, it will be assigned as 0.9
        If FinSubsidy(1) > 0 Then
            IRFactor(2) = SecondFactorP
        Else
            IRFactor(2) = SecondFactorN
        End If
        'IRFactor(2) = 1.5
     Else
        ' The new IRFactor is determined by the IRFactor, JStep and FinSubsidy in the previous run
        IRFactor(AutoLoop) = IRFactor(AutoLoop - 1) - JStep(AutoLoop - 1) * FinSubsidy(AutoLoop - 1)
     End If
   End If
   
   ' Set the factor in the hidden input worksheet
   ChangeIRFactor (CStr(IRFactor(AutoLoop)))
   
   'Run STATA
   RunStata
   'Application.Wait Now + #12:00:03 AM#
  
   'Run GetStataOutput
   GetStataOutput
  
   'Run CSC calculator and save the result in the Range CSC_Output
   run_csc
   
   'After the CSC generates the output, the new JStep will be determined.
    If IRFactor(AutoLoop) <= 0 Then
        MsgBox "No Positive Interest Rate exists to produce a Zero Financing Subsidy Rate, given the loan parameters. The Subsidy Rate given a Base Interest Rate of Zero is returned."
        Exit Sub
    End If
   If AutoLoop > 1 And IRFactor(AutoLoop) <> 0 Then
      JStep(AutoLoop) = (IRFactor(AutoLoop) - IRFactor(AutoLoop - 1)) / (FinSubsidy(AutoLoop) - FinSubsidy(AutoLoop - 1))
   End If
   
   ' increase the counter and subscript of array by 1
   AutoLoop = AutoLoop + 1
 
 Loop While Abs(FinSubsidy(AutoLoop - 1) - 0) >= ZeroFinancingRate
  
 'Final run to include capitalized spread
 If Abs(FinSubsidy(AutoLoop - 1) - 0) < ZeroFinancingRate Then
 'Run STATA
   RunStata
   'Application.Wait Now + #12:00:03 AM#
  
   'Run GetStataOutput
   GetStataOutput
  
   'Run CSC calculator and save the result in the Range CSC_Output
   run_csc
 End If
 
   Application.ScreenUpdating = True
   
End Sub

Sub ChangeIRFactor(factor As String)
    If factor < 0 Then
        ActiveWorkbook.Sheets("StataInput").Unprotect password:="budgetrocks"
        ActiveWorkbook.Sheets("StataInput").Range("IRFactor").Value = 0
        ActiveWorkbook.Save
    Else
        ActiveWorkbook.Sheets("StataInput").Unprotect password:="budgetrocks"
        ActiveWorkbook.Sheets("StataInput").Range("IRFactor").Value = factor
        ActiveWorkbook.Save
    End If
End Sub

Sub ZeroTotSub()
 AutoLoop2 = 1
 'Run IRHelper
 IRHelper
 AutoLoop2 = 2
 
 Dim feefactor(1 To 20) As Double, KStep(1 To 20) As Double
 Dim prevfeefactor As Double
 Const SecondFactorP As Double = 1.1 ' or 1.25
 Const SecondFactorN As Double = 0.9
 Const ZeroSubRate As Double = 0.00499999
 
 Application.ScreenUpdating = False
  
 Do
    ' calculate the value of next Factor
   If AutoLoop2 = 1 Then
        feefactor(1) = 1
        KStep(1) = 1
   Else
     If AutoLoop2 = 2 Then
        ' The second factor is decided as following. If the total subsidy is
        ' greater than 0, the FeeFactor will be assigned as either 1.1 or 1.25
        ' Otherwise, it will be assigned as 0.9
        If TotSubsidy(1) > 0 Then
            feefactor(1) = 1
            KStep(1) = 1
            feefactor(2) = SecondFactorP
        Else
            feefactor(1) = 1
            KStep(1) = 1
            feefactor(2) = SecondFactorN
        End If
     Else
        ' The new FeeFactor is determined by the FeeFactor, KStep and TotSubsidy in the previous run
        feefactor(AutoLoop2) = feefactor(AutoLoop2 - 1) - KStep(AutoLoop2 - 1) * TotSubsidy(AutoLoop2 - 1)
     End If
   End If
   
   ' Change the Pre & Post Completion Fees on Inputs tab to equal intial run value times Feefactor
   Call ChangeFeeFactor(CStr(feefactor(AutoLoop2)), CStr(feefactor(AutoLoop2 - 1)))
   
   'Run STATA
   RunStata
   'Application.Wait Now + #12:00:03 AM#
  
   'Run GetStataOutput
   GetStataOutput
  
   'Run CSC calculator and save the result in the Range CSC_Output
   run_csc
   
   'After the CSC generates the output, the new KStep will be determined.
   If feefactor(AutoLoop2) < ZeroFeeRate Then
        MsgBox "No Positive Fee Rate exists to produce a Zero Total Subsidy Rate given the loan parameters. The Subsidy Rate given a Zero Pre & Post Completion Fee is returned."
        Exit Sub
    End If
   If AutoLoop2 > 1 And feefactor(AutoLoop2) <> 0 Then
      KStep(AutoLoop2) = (feefactor(AutoLoop2) - feefactor(AutoLoop2 - 1)) / (TotSubsidy(AutoLoop2) - TotSubsidy(AutoLoop2 - 1))
    End If
   
   ' increase the counter and subscript of array by 1
   AutoLoop2 = AutoLoop2 + 1
 
 Loop While Abs(TotSubsidy(AutoLoop2 - 1) - 0) >= ZeroSubRate
 
  
   Application.ScreenUpdating = True
   
End Sub

Sub ChangeFeeFactor(factor2 As String, prevfactor As String)
'    ActiveWorkbook.Sheets("StataInput").Range("FeeFactor").Value = factor2
    ActiveWorkbook.Sheets("Inputs").Unprotect password:="budgetrocks"
    
    If ActiveWorkbook.Sheets("Inputs").Range("Direct_Guar").Value = "LPG" Or ActiveWorkbook.Sheets("Inputs").Range("Direct_Guar").Value = "Non-LPG" Then
        If factor2 < 0 Then
            ActiveWorkbook.Sheets("Inputs").Range("Fee_DCA").Value = 0
            ActiveWorkbook.Save
        Else
            ActiveWorkbook.Sheets("Inputs").Range("Fee_DCA").Value = factor2 * (Range("Fee_DCA").Value) / prevfactor
            ActiveWorkbook.Save
        End If
    
    Else
        If factor2 < 0 Then
            ActiveWorkbook.Sheets("Inputs").Range("fee_precomp").Value = 0
            ActiveWorkbook.Sheets("Inputs").Range("fee_postcomp").Value = 0
            ActiveWorkbook.Save
        Else
            ActiveWorkbook.Sheets("Inputs").Range("fee_postcomp").Value = factor2 * (Range("fee_postcomp").Value) / prevfactor
            ActiveWorkbook.Sheets("Inputs").Range("fee_precomp").Value = factor2 * (Range("fee_precomp").Value) / prevfactor
            ActiveWorkbook.Save
        End If
    End If
End Sub


