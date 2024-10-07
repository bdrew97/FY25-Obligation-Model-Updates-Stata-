Option Explicit

Dim CSC_location As Range
Dim CashFlowFile As Range
Dim dest As Range
Dim csc_full_path As String
Dim csc_output_fn As String
Dim csc_input_path As String
Dim csc_input_fn As String
Dim csc_output_path As String
Dim default_workbook_name As String
Dim cohort_year As Date
Dim range_name As String
Dim FXC As String
Dim TotOb As Double
Dim FXCap As Double
Dim SubDollars As Double
Dim TotSub As Double
Dim FinSub As Double
Dim DefSub As Double
Dim FeeSub As Double
Dim OthSub As Double
Dim AuthAmt As Double
Dim IGPercent As Double
Dim GuarCeil As Double
Dim First_Loss As Double
Dim FirstLossAdj As Double
Dim TotSubAdj As Double
Dim FinSubAdj As Double
Dim DefSubAdj As Double
Dim FeeSubAdj As Double
Dim OthSubAdj As Double
Dim DiscRate As Double
Dim SubsidyRangeAdj As Variant
Dim SubsidyRange As Variant


Sub run_csc()

   Application.ScreenUpdating = False
   
   Call unprotectAll
   csc_input_path = ActiveWorkbook.Path
   Sheets("Inputs").Select
   csc_output_path = RemoveBackSlash(Range("Stata_Dofile_Path").Text) & "\Outputs\" & Range("Loan_Officer").Text
   default_workbook_name = ActiveWorkbook.Name
   
   cohort_year = Range("date_obligation").Value
   
    ' Create the cash flow input file first
    Call prepare_csc_input

    Dim csc As String
    Dim arg2 As String
    Dim arg3 As String
    Dim arg4 As String
    Dim cmd_line As String
    Dim rv As Variant
    
    Dim fs
    Set fs = CreateObject("Scripting.FileSystemObject")

'   ---------------------------------------------------------------------
'   The full-path file names may have embedded spaces.  If so, the DOS
'   command handler will not parse the line correctly.  The solution is
'   to wrap everything with double quotes.  When that's done, the command
'   line can be assembled with a single space between each field.
'   ---------------------------------------------------------------------

    csc = """" & Range("CSC_Path").Value & """"
    
    arg2 = """" & csc_output_fn & """"
    
    arg3 = """" & csc_input_fn & """"
    arg4 = """" & "CSC_Input" & """"
    
'    cmd_line = csc & " " & "xls" & " " & arg2 & " " & arg3 & " " & arg4

'   CSC command
   cmd_line = csc & " -i " & arg3 & " -o " & arg2 & " -n " & arg4
    
    ShellAndWait (cmd_line)
'    Call get_csc_subsidytab
    
    Application.Wait Now + #12:00:01 AM#
    Call get_csc_output
    Call get_csc_subsidytab
    
    Application.ScreenUpdating = True
    Call ProtectSheets
    'If Sheets("Inputs").Range("Store_Option") = "Final Estimate" Then
    '    Call SaveResults
    'End If
    Sheets("Inputs").Select
    Sheets("Inputs").Range("CSC_Output").Select
    
End Sub


Sub prepare_csc_input()

   csc_input_fn = csc_output_path & "\cscinput.xlsx"
   csc_output_fn = csc_output_path & "\cscoutput.xlsx"
      
   
   Dim Wk As Workbook
   Set Wk = Workbooks.Add
   
   Application.DisplayAlerts = False
   
   Wk.SaveAs filename:=csc_input_fn
   
   Application.ScreenUpdating = True
   'Windows(default_workbook_name).Activate
   ThisWorkbook.Activate
   
    Worksheets("Cash Flow - To CSC").Range("A9:SF600").Copy _
    Destination:=Workbooks("cscinput.xlsx").Worksheets("Sheet1").Range("A1")
    
    Windows("cscinput.xlsx").Activate
    Application.ScreenUpdating = False
    
    'To remove the extra Total Column in the cash input
    Columns("C:C").Select
    Selection.Delete Shift:=xlToLeft
    Application.CutCopyMode = False
    
    Dim LastRow As Long
    Dim LastCol As Long
    Dim LastCell As String
    
    If WorksheetFunction.CountA(Cells) > 0 Then
        LastRow = Cells.Find(What:="Input Files:", After:=[A1], SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
        LastCol = Cells.Find(What:="*", After:=[A1], SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    End If
    
    LastCell = Split(Cells(1, LastCol).Address, "$")(1) & (LastRow - 1)
    ActiveSheet.Names.Add Name:="CSC_Input", RefersTo:=ActiveSheet.Range("C2", LastCell), Visible:=True
    
    Wk.Save
    
    '''  10/5 add extra paragraph here to fix the glitchs
    Dim i As Integer
    Dim lRow As Integer
    For i = 1 To 100
     If Wk.ActiveSheet.Cells(i, 2).Text = "160" Then
         lRow = i
         Exit For
     End If
    Next

   For i = 4 To 150
    If InStr(Wk.ActiveSheet.Cells(lRow, i), "=") Then
      Wk.ActiveSheet.Cells(lRow, i).Formula = "="""""
   End If
   Next
   ' End of paragraph
   Wk.Save
   Dim tempFile As String
   
   tempFile = Left(csc_input_fn, Len(csc_input_fn) - 1)
   
   Wk.SaveAs filename:= _
        tempFile, FileFormat _
        :=xlExcel8, password:="", WriteResPassword:="", ReadOnlyRecommended:= _
        False, CreateBackup:=False
 
   Wk.Close
   csc_input_fn = tempFile

End Sub


Sub get_csc_output()  'rates information start from cell F167 in the Inputs tab

   ' check the csc output file exists or not
   
   Do While FileOrDirExists(csc_output_fn) = False
     Application.Wait Now + #12:00:01 AM#
   Loop

    FXC = Workbooks(default_workbook_name).Sheets("Inputs").Range("FX_Denomination").Value
   
   Workbooks.Open filename:=csc_output_fn

   Sheets("Subsidy").Select
    
   Dim fyear As Integer
   fyear = Year(cohort_year)
   
   Dim mth As Integer
   mth = Month(cohort_year)
   
   If mth > 9 Then
      fyear = fyear + 1
   End If
   
   Dim LastRow As Long
   
   
   Dim i As Integer
 
   For i = 1 To 100
      If Cells(i, 1).Value = fyear Then
         LastRow = i
         Exit For
      End If
   Next
   
      
   Dim ranges As String
   Dim FX_Subsidy As Variant
   ranges = "B" & LastRow & ":" & "G" & LastRow
   
   
'Start
   
   TotSub = Sheets("Subsidy").Range(ranges).Cells(1, 1).Value
   FinSub = Sheets("Subsidy").Range(ranges).Cells(1, 2).Value
   DefSub = Sheets("Subsidy").Range(ranges).Cells(1, 3).Value
   FeeSub = Sheets("Subsidy").Range(ranges).Cells(1, 4).Value
   OthSub = Sheets("Subsidy").Range(ranges).Cells(1, 5).Value
   DiscRate = Sheets("Subsidy").Range(ranges).Cells(1, 6).Value
   
   
   If FXC = "Yes" Then
   Application.ScreenUpdating = True
   'Windows(default_workbook_name).Activate
   ThisWorkbook.Activate
   Application.ScreenUpdating = False
   
   SubsidyRange = Array(TotSub, FinSub, DefSub, FeeSub, OthSub, DiscRate)
   Sheets("Inputs").Range("CSC_Output_Orig") = WorksheetFunction.Transpose(SubsidyRange)
   Sheets("Inputs").Range("Fin_Subsidy") = FinSub
   
   TotOb = Sheets("Cash Flow - To CSC").Range("E20").Value
   FXCap = Sheets("Inputs").Range("FXCap").Value
   
   SubDollars = (Round(TotSub, 2) / 100) * TotOb
   
   TotSub = (SubDollars / FXCap) * 100
   FinSub = (TotOb / FXCap) * FinSub
   DefSub = (TotOb / FXCap) * DefSub
   FeeSub = (TotOb / FXCap) * FeeSub
   OthSub = (TotOb / FXCap) * OthSub
   
   Else
    Application.ScreenUpdating = True
    'Windows(default_workbook_name).Activate
    ThisWorkbook.Activate
    Application.ScreenUpdating = False
   End If
      
   SubsidyRange = Array(TotSub, FinSub, DefSub, FeeSub, OthSub, DiscRate)
   
   If Sheets("Inputs").Range("Direct_Guar") = "LPG" Or Sheets("Inputs").Range("Direct_Guar") = "Non-LPG" Then

   AuthAmt = Sheets("Inputs").Range("_3Terms")
   IGPercent = Sheets("Inputs").Range("perc_guaranteed")
   First_Loss = Sheets("Inputs").Range("First_Loss")
   GuarCeil = AuthAmt * IGPercent
   'FirstLossAdj = GuarCeil - First_Loss
   
   SubDollars = (Round(TotSub, 2) / 100) * AuthAmt
   
   TotSubAdj = (SubDollars / ((AuthAmt - First_Loss) * IGPercent)) * 100
   FinSubAdj = (AuthAmt / ((AuthAmt - First_Loss) * IGPercent)) * FinSub
   DefSubAdj = (AuthAmt / ((AuthAmt - First_Loss) * IGPercent)) * DefSub
   FeeSubAdj = (AuthAmt / ((AuthAmt - First_Loss) * IGPercent)) * FeeSub
   OthSubAdj = (AuthAmt / ((AuthAmt - First_Loss) * IGPercent)) * OthSub
   
   SubsidyRangeAdj = Array(TotSubAdj, FinSubAdj, DefSubAdj, FeeSubAdj, OthSubAdj, DiscRate)
   Sheets("Inputs").Range("CSC_Output_Orig") = WorksheetFunction.Transpose(SubsidyRangeAdj)
   
   End If
   
   'Guarantee ceiling adjustments for non-MTU guarantees not denominated in a local currency. This information is displayed in results section of Inputs tab
   If FXC = "No" And (Sheets("Inputs").Range("Direct_Guar") = "Investment Guaranty" Or Sheets("Inputs").Range("Direct_Guar") = "Investment Guaranty - NHSG" Or Sheets("Inputs").Range("Direct_Guar") = "Breach of Contract (AAD)") Then

   AuthAmt = Sheets("Inputs").Range("_3Terms")
   IGPercent = Sheets("Inputs").Range("perc_guaranteed")
   GuarCeil = AuthAmt * IGPercent
   
   SubDollars = (Round(TotSub, 2) / 100) * AuthAmt
   
   TotSubAdj = (SubDollars / (AuthAmt * IGPercent)) * 100
   FinSubAdj = (AuthAmt / GuarCeil) * FinSub
   DefSubAdj = (AuthAmt / GuarCeil) * DefSub
   FeeSubAdj = (AuthAmt / GuarCeil) * FeeSub
   OthSubAdj = (AuthAmt / GuarCeil) * OthSub
   
   SubsidyRangeAdj = Array(TotSubAdj, FinSubAdj, DefSubAdj, FeeSubAdj, OthSubAdj, DiscRate)
   Sheets("Inputs").Range("CSC_Output_Orig") = WorksheetFunction.Transpose(SubsidyRangeAdj)
   
   End If
   
   Sheets("Inputs").Range("CSC_Output") = WorksheetFunction.Transpose(SubsidyRange)
   Sheets("Inputs").Range("Fin_Subsidy") = FinSub
   'End
    ActiveWorkbook.Save
    
    'Save the financial subsidy rate & total subsidy rate of each run in a global array
    If AutoLoop2 >= 1 Then
    TotSubsidy(AutoLoop2) = Range("CSC_Output").Cells(1).Value
    End If
    If AutoLoop >= 1 Then
    FinSubsidy(AutoLoop) = Range("CSC_Output").Cells(2).Value
    End If
    Application.ScreenUpdating = True
    Windows(Dir(csc_output_fn)).Activate
    Application.ScreenUpdating = False
    ActiveWindow.Close
       
    ' Remove the temporary CSC input cash flow xlsx file
    'If Dir(csc_input_fn) <> "" Then _
    '       Kill csc_input_fn

End Sub


Sub get_csc_subsidytab()
 Dim range1 As Range
 Dim ws As Worksheet
 
 Application.ScreenUpdating = True
    'Windows(default_workbook_name).Activate
    ThisWorkbook.Activate
 Application.ScreenUpdating = False
    
  Application.DisplayAlerts = False
    On Error Resume Next
    Sheets("Subsidy").Cells.Clear
    Sheets("Source data").Cells.Clear
    Sheets("Distributions").Cells.Clear
    Sheets("Discounts").Cells.Clear
    Sheets("PV Factors").Cells.Clear
    Sheets("Subsidy Calculations").Cells.Clear
    Sheets("Messages").Cells.Clear
    Sheets("Combined").Cells.Clear
    Sheets("NFD").Cells.Clear
    Application.DisplayAlerts = True
    On Error GoTo 0
    
   ' check the csc output file exists or not
   Do While FileOrDirExists(csc_output_fn) = False
     Application.Wait Now + #12:00:01 AM#
   Loop
   
   Workbooks.Open filename:=csc_output_fn
' The following line triggers a VBA message if CSC does not save output (likely due to settings). Add Error Handling Here.
   On Error Resume Next
   Sheets("Subsidy").Select
       Select Case Err.Number
    Case Is = 0
    '
    Case Else
        MsgBox "An error has occured when attempting to retrieve CSC output. Please ensure that the CSC Output Preferences are consistent with Model Requirements. Refer to 'FY 2014 Obligation Model Setup Instructions' for the appropriate CSC settings"
        End
    End Select
     'Resume error checking
    On Error GoTo 0

' Move Sheets into Obligation Model
    Application.ScreenUpdating = True
    Workbooks(default_workbook_name).Activate
    Call unprotectAll
   
    Windows(Dir(csc_output_fn)).Activate
    For Each ws In ActiveWorkbook.Worksheets
        If (ws.Name <> "CSC") And (ws.Name <> "Summary") Then
            ws.Cells.Copy
            Workbooks(default_workbook_name).Activate
            Workbooks(default_workbook_name).Sheets(ws.Name).Select
            Sheets(ws.Name).Cells.PasteSpecial
            Application.CutCopyMode = False
            ActiveSheet.Tab.Color = RGB(153, 204, 255)
            Windows(Dir(csc_output_fn)).Activate
        End If
    Next ws

    Windows(Dir(csc_output_fn)).Activate
    ActiveWindow.Close False
    
    Workbooks(default_workbook_name).Activate
    Call ProtectSheets
    'Application.ScreenUpdating = True
  End Sub

Sub ProtectSheets()
Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        ws.Protect password:="budgetrocks"
    Next ws
End Sub



