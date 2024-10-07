Sub GetStataOutput()
    
Dim OutputPath As String
Dim OutputPath_Risk As String
Dim OutputPath_RiskICRAS As String
Dim CurrentFile As String
Dim SelectionStore As String
Dim CSCRangeString As String
Dim RawCSCRangeString As String
Dim FXCSCRangeString As String
Dim LastRow As Integer
Dim LastCol As Integer
Dim LastCell As String
Dim LastRowUse As String
Dim i  As Integer
Dim default_workbook_name As String
Dim shape As Excel.shape



'*******************************************************************************************
' STEP 1 - Get *Risk* Table, FOR FUTURE UPDATE                      ************************
'*******************************************************************************************

'    Application.ScreenUpdating = False
'
'' Remove Risk Table Sheet - if it exists
'    Application.DisplayAlerts = False
'     On Error Resume Next
'     Sheets("Risk Table").Delete
'     Sheets("Risk Table ICRAS").Delete
'     On Error GoTo 0
'    Application.DisplayAlerts = True
'
'' Store Current File Name
'  CurrentFile = ActiveWorkbook.Name
'
''Import Risk Table File
'    OutputPath_Risk = Range("Stata_Dofile_Path").Text & "Outputs\Risk Table.xlsx"
'    OutputPath_RiskICRAS = Range("Stata_Dofile_Path").Text & "Outputs\Risk Table ICRAS.xlsx"
'
'    If Workbooks(CurrentFile).Sheets("Inputs").Range("Default_Method") = "Moody's Risk Methodology" Then
'        Workbooks.Open filename:=OutputPath_Risk
'    ElseIf Workbooks(CurrentFile).Sheets("Inputs").Range("Default_Method") = "ICRAS" Then
'        Workbooks.Open filename:=OutputPath_RiskICRAS
'    End If
'
'' Move into Current File (Obligation Model)
'    Sheets(1).Select
'    Sheets(1).Move After:=Workbooks(CurrentFile).Sheets("Subsidy Term Sheet")
'    On Error Resume Next
'    Sheets("Risk Table").Tab.Color = RGB(255, 155, 139)
'    Sheets("Risk Table ICRAS").Tab.Color = RGB(255, 155, 139)
'    On Error GoTo 0
'  ' Put cursor in correct place
'    Application.ScreenUpdating = True
'    Range("A1").Select
    
'*******************************************************************************************
' STEP 2 - Get Life Table              *****************************************************
'*******************************************************************************************

    Application.ScreenUpdating = False
    
    default_workbook_name = ActiveWorkbook.Name

' Remove Life Table Sheet - if it exists
    Application.DisplayAlerts = False
     On Error Resume Next
     Sheets("Life Table").Cells.Clear
     Sheets("Life Table Copy").Delete
     On Error GoTo 0
    Application.DisplayAlerts = True
  
' Store Current File Name
  CurrentFile = ActiveWorkbook.Name
  
'Import Life Table File
    OutputPath = Range("Stata_Dofile_Path").Text & "Outputs\" & Range("Loan_Officer").Text & "\Life Table.xlsx"
    Workbooks.Open filename:=OutputPath
    
'Format Life Table
    
  ' Replace "_" with " "
    Rows("1:1").Select
    Selection.Replace What:="_", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With

    
  ' Column Width and 1st Row Height
    ActiveWindow.Zoom = 90
    Cells.Select
    Cells.EntireColumn.AutoFit
    Columns("C:C").ColumnWidth = 10.71
    Columns("D:D").ColumnWidth = 13.71
    Columns("E:E").ColumnWidth = 12.86
    Columns("G:G").ColumnWidth = 11
    Columns("J:J").ColumnWidth = 8.71
    Columns("L:L").ColumnWidth = 12.29
    Rows("1:1").EntireRow.AutoFit
    Rows("1:1").Select
  
  

  ' Format First Row
  
    Rows("1:1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Font.Bold = True

  ' Number Formatting
    Columns("D:M").Select
    Selection.Style = "Comma"
    Columns("N:S").Select
    Selection.NumberFormat = "0.000%"
    Columns("T:AC").Select
    Selection.Style = "Comma"
    Columns("AD:AF").Select
    Selection.NumberFormat = "0.00"
    Columns("AG:AG").Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.000%"
    Columns("AH:AI").Select
    Selection.Style = "Comma"
    
' Move into Current File (Obligation Model)

    Sheets(1).Cells.Copy
    Application.ScreenUpdating = True
    Workbooks(default_workbook_name).Activate
    Workbooks(default_workbook_name).Sheets("Life Table").Select
    Sheets("Life Table").Cells.PasteSpecial
    Workbooks("Life Table.xlsx").Activate
    Application.ScreenUpdating = False
    Application.CutCopyMode = False
    ActiveWorkbook.Close False
  ' Copy Header from Inputs sheet
  
    For Each shape In ActiveSheet.Shapes
        shape.Delete
    Next
  
    Sheets("Inputs").Select
    Rows("1:8").Select
    Selection.Copy
    Sheets("Life Table").Select
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown
    Range("D6").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "Life Table"
    Sheets("Inputs").Select
    ActiveSheet.Shapes.Range(Array("Picture 10")).Select
    Selection.Copy
    Sheets("Life Table").Select
    Range("A1").Select
    ActiveSheet.Paste
    Columns("A:B").Select
    Selection.ColumnWidth = 14.3
    Rows("8").Select
    Selection.RowHeight = 12#
    Range("A1:B5").Select
    Application.ScreenUpdating = True
    Range("B1").Activate
    Application.ScreenUpdating = False
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = -0.249946592608417
        .PatternTintAndShade = 0
    End With
    ActiveWindow.DisplayGridlines = False
    Sheets("Life Table").Tab.Color = RGB(255, 155, 139)
    
  ' Display as Table
  
    Range("A9").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    SelectionStore = Selection.Address
    ActiveWorkbook.Names.Add Name:="LT_Table", RefersToR1C1:=SelectionStore
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$9:$AI$5000"), , xlYes).Name = _
        "LifeTable"
    Range("LifeTable[#All]").Select
    ActiveSheet.ListObjects("LifeTable").TableStyle = "TableStyleLight9"
    Rows("9:9").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
  ' Ensure column size works
    Cells.Select
    Cells.EntireColumn.AutoFit
    Columns("D:D").ColumnWidth = 18
    
  '  Freeze Panes
    
    ActiveWindow.FreezePanes = False
    Range("C10").Select
    ActiveWindow.FreezePanes = True
    
  ' FOR FUTURE MODEL UPDATE:
'    ' Create Copy to Contain Excel formulas
'      Sheets("Life Table").Copy After:=Sheets("Life Table")
'      Sheets("Life Table (2)").Name = "Life Table Copy"
'       With ActiveSheet
'          .ListObjects(1).Name = "LTCopy"
'      End With
'
'    ' Insert Life Table Formulas (eventually move so that this only occurs at last iteration)
'    '  Call LTFormulas
  
  ' Put cursor in correct place
    Application.ScreenUpdating = True
    Range("A10").Select
    Range("A9").Select

'*******************************************************************************************
' STEP 3 - Get Cash Flow               *****************************************************
'*******************************************************************************************

    Application.ScreenUpdating = False
    
' Remove Cash Flow Raw Sheet - if it exists
    Application.DisplayAlerts = False
     On Error Resume Next
     Sheets("Cash Flow - Raw").Cells.Clear
     Sheets("Cash Flow - To CSC").Cells.Clear
     On Error GoTo 0
    Application.DisplayAlerts = True

'Import Cash Flow File
    OutputPath = Range("Stata_Dofile_Path").Text & "Outputs\" & Range("Loan_Officer").Text & "\CSC Cashflow.xlsx"
    Workbooks.Open filename:=OutputPath

' Format Cash Flow Raw Sheet
    ChangeNumberFormat
    Columns("B:B").Select
    Selection.NumberFormat = "0"
    ActiveWindow.Zoom = 90
    Cells.Select
    Cells.EntireColumn.AutoFit
    
' Move into Current File (Obligation Model)
    
    Application.ScreenUpdating = True
    Workbooks("CSC Cashflow.xlsx").Activate
    Sheets(1).Select
    Sheets(1).Cells.Copy
    Workbooks(default_workbook_name).Activate
    Workbooks(default_workbook_name).Sheets("Cash Flow - Raw").Select
    Sheets("Cash Flow - Raw").Cells.PasteSpecial
    Sheets("Cash Flow - Raw").Tab.Color = RGB(255, 155, 139)
    Workbooks("CSC Cashflow.xlsx").Activate
    Application.ScreenUpdating = False
    Application.CutCopyMode = False
    ActiveWorkbook.Close False
    

' Copy Header from Inputs sheet
  
    For Each shape In ActiveSheet.Shapes
        shape.Delete
    Next
  
    Sheets("Inputs").Select
    Rows("1:8").Select
    Selection.Copy
    Sheets("Cash Flow - Raw").Select
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown
    Range("D6").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "Raw Cash Flow (From Stata)"

    Sheets("Inputs").Select
    ActiveSheet.Shapes.Range(Array("Picture 10")).Select
    Selection.Copy
    Sheets("Cash Flow - Raw").Select
    Range("A1").Select
    ActiveSheet.Paste
    Columns("A:B").Select
    Selection.ColumnWidth = 14
    Rows("8").Select
    Selection.RowHeight = 12#
    Range("A1:B5").Select
    Application.ScreenUpdating = True
    Range("B1").Activate
    Application.ScreenUpdating = False
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = -0.249946592608417
        .PatternTintAndShade = 0
    End With
    
  ' Move the Heading to the Left and Indent
    Range("D2:D7").Select
    Selection.Copy
    Range("C2").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    With Selection
        .HorizontalAlignment = xlLeft
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 5
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("D3:D7").Select
    Selection.ClearContents
    Range("D10").Select
    
  ' Filter First Row
    Rows("9:9").Select
    Selection.AutoFilter
    Range("D10").Select
    
  ' Erase empty cells
    Range("E10:OJ20").ClearContents
    
  ' Add Summation Column and Hide
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("C9").Select
    ActiveCell.FormulaR1C1 = "Totals"
    Range("C10").Select
    ActiveCell.FormulaR1C1 = "=+SUM(RC[2]:RC[16381])"
    Range("C10").Select
    Selection.AutoFill Destination:=Range("C10:C1000")
    Range("C10:C1000").Select
    Columns("C:C").Select
    Selection.EntireColumn.Hidden = True
    
    Columns("A:A").Select
    Selection.NumberFormat = "General"

  ' Freeze Panes
    ActiveWindow.FreezePanes = False
    Range("D:D").ColumnWidth = 70
    Range("E10").Select
    ActiveWindow.FreezePanes = True
     
' Define CSC Input Range
  ' Find Last Row
    
    'LastRow = ActiveSheet.UsedRange.Rows.count
    
    'For i = 1 To LastRow
        'With Cells(i, 1)
            'If .Value = "Input Files:" Then
                'LastRowUse = i
            'End If
        'End With
        'Next i
        
    LastRow = Cells.Find(What:="Input Files:", After:=[A1], SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    LastCol = Cells.Find(What:="*", After:=[A1], SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
        
    LastCell = "R" & (LastRow - 2) & "C" & LastCol

    RawCSCRangeString = "='Cash Flow - Raw'!R10C4:" & LastCell
    RawCSCRangeString = "" & RawCSCRangeString & ""
    
    'Inserted 10/2013, Kayla
    Range("E10").Select
    ActiveWorkbook.Names.Add Name:="Raw_CF", RefersToR1C1:=RawCSCRangeString
    
 ' Copy Sheet for User Manipulation
    Sheets("Cash Flow - To CSC").Select
    Sheets("Cash Flow - To CSC").Cells.Clear
    
    For Each shape In ActiveSheet.Shapes
        shape.Delete
    Next
    
    Sheets("Cash Flow - Raw").Select
    Sheets("Cash Flow - Raw").Cells.Copy
    Sheets("Cash Flow - To CSC").Select
    Sheets("Cash Flow - To CSC").Cells.PasteSpecial
    Application.CutCopyMode = False
    Range("D6").Select
    ActiveCell.FormulaR1C1 = _
        "Cash Flow - to CSC"
    Range("E10").Select
    
 ' Set CSC Range
    LastRow = Cells.Find(What:="Input Files:", After:=[A1], SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    LastCol = Cells.Find(What:="*", After:=[A1], SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
        
    LastCell = "R" & (LastRow - 2) & "C" & LastCol
 
    CSCRangeString = "='Cash Flow - To CSC'!R10C4:" & LastCell
    CSCRangeString = "" & CSCRangeString & ""
    
    Range("D10").Select
    ActiveWorkbook.Names.Add Name:="CSC_Input", RefersToR1C1:=CSCRangeString
    
  ' Link Results on input sheet to Life Table
  
    Sheets("Inputs").Select
    Range("Nominal_Sums").Select
    Selection.ClearContents
        
    If Range("Direct_Guar").Value = "Direct Loan" Or Range("Direct_Guar").Value = "DPA" Then
        Range("USD_Commitment").Select
        ActiveCell.FormulaR1C1 = _
            "=+VLOOKUP(""Obligation"",'Cash Flow - To CSC'!R10C4:R40C5,2,0)"
    ElseIf Range("Direct_Guar").Value = "Investment Guaranty" Then
        Range("USD_Commitment").Select
        ActiveCell.FormulaR1C1 = _
            "=+VLOOKUP(""Commitment"",'Cash Flow - To CSC'!R10C4:R40C5,2,0)"
    End If
     
    Range("WAL_Lookup").Select
    ActiveCell.FormulaR1C1 = _
        "=+HLOOKUP(""Total WAL"",'Life Table'!R9C1:R10C35,2,0)"
        
    Range("term_interestrate").Select
    ActiveCell.FormulaR1C1 = _
        "=+HLOOKUP(""Base Interest WA"",'Life Table'!R9C1:R10C35,2,0)"
        
    Range("Nominal_Interest").Select
     ActiveCell.FormulaR1C1 = _
         "=+SUMIF('Cash Flow - To CSC'!R10C2:R2448C2,170,'Cash Flow - To CSC'!R10C3:R2448C3)"

    Range("Nominal_Principal").Select
    ActiveCell.FormulaR1C1 = _
        "=+SUMIF('Cash Flow - To CSC'!R10C2:R2448C2,160,'Cash Flow - To CSC'!R10C3:R2448C3)"

    Range("Nominal_Fees").Select
    ActiveCell.FormulaR1C1 = _
        "=+SUMIF('Cash Flow - To CSC'!R10C2:R1000C2,150,'Cash Flow - To CSC'!R10C3:R2448C3)+SUMIF('Cash Flow - To CSC'!R10C2:R1000C2,180,'Cash Flow - To CSC'!R10C3:R2448C3)+SUMIF('Cash Flow - To CSC'!R10C2:R1000C2,330,'Cash Flow - To CSC'!R10C3:R2448C3)"
        
    If Range("Direct_Guar").Value = "Direct Loan" Or Range("Direct_Guar").Value = "DPA" Then
        Range("Nominal_Default").Select
        ActiveCell.FormulaR1C1 = _
            "=+SUMIF('Cash Flow - To CSC'!R10C2:R1000C2,200,'Cash Flow - To CSC'!R10C3:R2448C3)"

    Else
        
        Range("Nominal_Default").Select
        ActiveCell.FormulaR1C1 = _
            "=+SUMIF('Cash Flow - To CSC'!R10C2:R1000C2,190,'Cash Flow - To CSC'!R10C3:R2448C3)- SUMIF('Cash Flow - To CSC'!R10C2:R1000C2,200,'Cash Flow - To CSC'!R10C3:R2448C3)"
    End If

' Add Warning Text Box
    Sheets("Cash Flow - Raw").Select
    AddTextBox
    
' Add CSC Button
    Sheets("Cash Flow - To CSC").Select
    ActiveSheet.Buttons.Add(590.25, 44.25, 144.75, 33.75).Select
    Selection.Characters.Text = "Run CSC"
    Selection.OnAction = "run_csc"
   ' ActiveSheet.Shapes("Button 1").Select
    Selection.Name = "Run_CSC_Button"
    ActiveSheet.Shapes("Run_CSC_Button").Select
    With Selection.Characters(Start:=1, Length:=8).Font
        .Name = "Calibri"
        .FontStyle = "Bold"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
    End With
    
    Sheets("Inputs").Select
    ActiveSheet.Shapes.Range(Array("Picture 10")).Select
    Selection.Copy
    Sheets("Cash Flow - To CSC").Select
    Range("A1").Select
    ActiveSheet.Paste
    
    Range("E10").Select

    Application.ScreenUpdating = True
 '   Sheets("Cash Flow - Raw").Select
 '   Range("E10").Select
 
'*******************************************************************************************
' STEP 4 - Get FX Cash Flow               *****************************************************
'*******************************************************************************************
'Inserted 10/2018, Christine

    Application.ScreenUpdating = False
    If Sheets("Inputs").Range("FX_Denomination") = "No" Then
        Sheets("Cash Flow - FX").Visible = False
    ElseIf Sheets("Inputs").Range("FX_Denomination") = "Yes" Then
        Sheets("Cash Flow - FX").Visible = True
    End If

' Remove Cash Flow Raw Sheet - if it exists
    Application.DisplayAlerts = False
     On Error Resume Next
     Sheets("Cash Flow - FX").Cells.Clear
     On Error GoTo 0
    Application.DisplayAlerts = True

'Import Cash Flow File
    OutputPath = Range("Stata_Dofile_Path").Text & "Outputs\" & Range("Loan_Officer") & "\CSC Cashflow FX.xlsx"
    If Dir$(OutputPath) = "" Then
        Exit Sub
    End If
    Workbooks.Open filename:=OutputPath

' Format Cash Flow Raw Sheet
    ChangeNumberFormat
    Columns("B:B").Select
    Selection.NumberFormat = "0"
    ActiveWindow.Zoom = 90
    Cells.Select
    Cells.EntireColumn.AutoFit
    
' Move into Current File (Obligation Model)

    Sheets(1).Select
    Sheets(1).Name = "Cash Flow - FX"
    Sheets(1).Cells.Copy
    Application.ScreenUpdating = True
    Workbooks(default_workbook_name).Activate
    Workbooks(default_workbook_name).Sheets("Cash Flow - FX").Select
    Sheets("Cash Flow - FX").Cells.PasteSpecial
    Application.CutCopyMode = False
    Sheets("Cash Flow - FX").Tab.Color = RGB(255, 155, 139)
    Workbooks("CSC Cashflow FX.xlsx").Activate
    Application.ScreenUpdating = False
    ActiveWorkbook.Close False

' Copy Header from Inputs sheet
  
    For Each shape In ActiveSheet.Shapes
        shape.Delete
    Next
  
    Sheets("Inputs").Select
    Rows("1:8").Select
    Selection.Copy
    Sheets("Cash Flow - FX").Select
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown
    Range("D6").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "FX Cash Flow (From Stata)"

    Sheets("Inputs").Select
    ActiveSheet.Shapes.Range(Array("Picture 10")).Select
    Selection.Copy
    Sheets("Cash Flow - FX").Select
    Range("A1").Select
    ActiveSheet.Paste
    Columns("A:B").Select
    Selection.ColumnWidth = 14
    Rows("8").Select
    Selection.RowHeight = 12#
    Range("A1:B5").Select
    Application.ScreenUpdating = True
    Range("B1").Activate
    Application.ScreenUpdating = False
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = -0.249946592608417
        .PatternTintAndShade = 0
    End With
    
  ' Move the Heading to the Left and Indent
    Range("D2:D7").Select
    Selection.Copy
    Range("C2").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    With Selection
        .HorizontalAlignment = xlLeft
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 5
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("D3:D7").Select
    Selection.ClearContents
    Range("D10").Select
    
 
    
  ' Filter First Row
    Rows("9:9").Select
    Selection.AutoFilter
    Range("D10").Select
    
  ' Erase empty cells
    Range("E10:OJ20").ClearContents
    
  ' Add Summation Column and Hide
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("C9").Select
    ActiveCell.FormulaR1C1 = "Totals"
    Range("C10").Select
    ActiveCell.FormulaR1C1 = "=+SUM(RC[2]:RC[16381])"
    Range("C10").Select
    Selection.AutoFill Destination:=Range("C10:C1000")
    Range("C10:C1000").Select
    Columns("C:C").Select
    Selection.EntireColumn.Hidden = True
    
    Columns("A:A").Select
    Selection.NumberFormat = "General"
    
  ' Freeze Panes
    ActiveWindow.FreezePanes = False
    Range("D:D").ColumnWidth = 70
    Range("E10").Select
    ActiveWindow.FreezePanes = True
    
    ' Define Range
    LastRow = Cells.Find(What:="Input Files:", After:=[A1], SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    LastCol = Cells.Find(What:="*", After:=[A1], SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
        
    LastCell = "R" & (LastRow - 2) & "C" & LastCol
    FXCSCRangeString = "='Cash Flow - FX'!R10C4:" & LastCell
    FXCSCRangeString = "" & FXCSCRangeString & ""
 
    Range("E10").Select
    ActiveWorkbook.Names.Add Name:="FX_CF", RefersToR1C1:=FXCSCRangeString
    
    ' Add Warning Text Box
    Sheets("Cash Flow - FX").Select
    AddTextBoxFX
    
End Sub

Sub ChangeNumberFormat()
    Dim i  As Integer
    Dim j As Integer
    Dim MaxI As Integer
    Dim MaxJ As Integer
    Dim FormatType As String
    
    ' i = row
    ' j = column
    
    MaxI = ActiveSheet.UsedRange.Rows.count
    MaxJ = ActiveSheet.UsedRange.Columns.count

    For i = 1 To MaxI
        For j = 1 To MaxJ
            
             With Cells(i, j)
             
               If j = 3 And .Value = "*FY" Or .Value = "*Quarter" Or .Value = "*Year" Or .Value = "*Month" _
                Or .Value = "*Period" Or .Value = "Budget Year" Or .Value = "Cohort" Or .Value = "*Period (Semiannual)" _
                Then
                  FormatType = "General"
                ElseIf j = 3 And .Value = "*Cumulative Default Rate" Or .Value = "*Marginal Default Rate" Or _
                  .Value = "*Recovery Rate" Or .Value = "*Default_Rate" Then
                  FormatType = "Rate"
                ElseIf j = 3 And .Value <> "*FY" And .Value <> "*Quarter" And .Value <> "*Year" And .Value <> "*Month" And _
                  .Value <> "*Period" And .Value <> "*Cumulative Default Rate" And .Value <> "*Marginal Default Rate" And _
                  .Value <> "Budget Year" And .Value <> "Cohort" And .Value <> "*Recovery Rate" And .Value <> "*Default_Rate" _
                  And .Value <> "*Period (Semiannual)" Then
                  FormatType = ""
               End If
               If .Value <> "" And IsNumeric(.Value) And FormatType = "" Then
                .Value = .Value
                .NumberFormat = "_(#,##0.00_);_((#,##0.00);_(""-""??_);_(@_)"
                ElseIf .Value <> "" And IsNumeric(.Value) And FormatType = "General" Then
                   .Value = .Value
                   .NumberFormat = "0"
                ElseIf .Value <> "" And IsNumeric(.Value) And FormatType = "Rate" Then
                   .Value = .Value
                   .NumberFormat = "0.0000"
                End If
             End With
      Next j
    Next i
End Sub


Sub AddTextBox()
   Range("G11:N18").Select
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    Selection.Font.Bold = True
    With Selection.Font
        .Name = "Calibri"
        .Size = 18
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .Color = -16776961
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    Range("G11").Select
    ActiveCell.FormulaR1C1 = _
        "This Sheet contains the raw Cash Flow as produced by Stata.  All hand manipulations should be done on the ""Cash Flow - To CSC"" sheet which is a copy of this sheet.  The CSC named range is the ""Cash Flow - To CSC"" sheet."

 
End Sub


Sub AddTextBoxFX()
   Range("G11:N18").Select
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    Selection.Font.Bold = True
    With Selection.Font
        .Name = "Calibri"
        .Size = 18
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .Color = -16776961
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    Range("G11").Select
    ActiveCell.FormulaR1C1 = _
        "This Sheet contains the raw FX Denominated Cash Flows as produced by Stata.  All hand manipulations should be done on the ""Cash Flow - To CSC"" sheet which contains the USD denominated cash flows.  The CSC named range is the ""Cash Flow - To CSC"" sheet."

 
End Sub

' For Future Model Updates, Include formulas in LT tab
'Sub LTFormulas()
''UPB SOP
'    ThisWorkbook.Worksheets("Life Table Copy").Range("LTCopy[UPB SOP]").Select
'    ThisWorkbook.Worksheets("Life Table Copy").Range("LTCopy[UPB SOP]").Value = "=round(SUMIFS([Disbursement],[Disbursement Number],[@[Disbursement Number]],[Repayment Date],""<""&[@[Repayment Date]])-SUMIFS([Principal],[Disbursement Number],[@[Disbursement Number]],[Repayment Date],""<""&[@[Repayment Date]])+SUMIFS([Cap Interest],[Disbursement Number],[@[Disbursement Number]],[Repayment Date],""<""&[@[Repayment Date]]),2)"
''UPB EOP
'    ThisWorkbook.Worksheets("Life Table Copy").Range("LTCopy[UPB EOP]").Select
'    ThisWorkbook.Worksheets("Life Table Copy").Range("LTCopy[UPB EOP]").Value = "=round([@[UPB SOP]]+[@Disbursement]+[@[Cap Interest]]-[@Principal],2)"
'
'
'
'End Sub


