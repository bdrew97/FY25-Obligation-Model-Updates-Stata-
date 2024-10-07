Sub Clear_Inputs()

    'Clears Inputs for the first three sections
    'Leaves references & links to extrenal files
    Application.ScreenUpdating = False
    
    Dim Sh As Worksheet
    Dim myPassword As String
        myPassword = "budgetrocks"

        For Each Sh In ActiveWorkbook.Worksheets
        Sh.Unprotect password:=myPassword
    Next Sh

    Sheets("Inputs").Select
    
    'Clear Description
    Worksheets("Inputs").Range("F12:F24").ClearContents
    Worksheets("Inputs").Range("F27:F31").ClearContents
    Worksheets("Inputs").Range("F35:F36").ClearContents
    Worksheets("Inputs").Range("F40").ClearContents
    Worksheets("Inputs").Range("F42").ClearContents
    Worksheets("Inputs").Range("F47:F55").ClearContents
    
    'Clear Inputs
    Worksheets("Inputs").Range("F57").ClearContents
    Worksheets("Inputs").Range("F61").ClearContents
    Worksheets("Inputs").Range("F65:F67").ClearContents
    Worksheets("Inputs").Range("F71:F72").ClearContents
    Worksheets("Inputs").Range("F74:F75").ClearContents
    Worksheets("Inputs").Range("F78").ClearContents
    Worksheets("Inputs").Range("F79").ClearContents
    Worksheets("Inputs").Range("F81").ClearContents
    Worksheets("Inputs").Range("H75").ClearContents
    Worksheets("Inputs").Range("F84:F86").ClearContents
    Worksheets("Inputs").Range("F88").ClearContents
    Worksheets("Inputs").Range("F91:F94").ClearContents
    Worksheets("Inputs").Range("F97").ClearContents
    Worksheets("Inputs").Range("F99").ClearContents
    
    'Clear Risk Ratings
    Worksheets("Inputs").Range("F115").ClearContents
    Worksheets("Inputs").Range("F119:F120").ClearContents
    Worksheets("Inputs").Range("F123").ClearContents
    Worksheets("Inputs").Range("F127").ClearContents
    
    'Clear Disbursement Schedule
    Sheets("Disbursement Schedule").Select
    Worksheets("Disbursement Schedule").Range("B13:B111").ClearContents
    Worksheets("Disbursement Schedule").Range("H12:H111").ClearContents
    Worksheets("Disbursement Schedule").Range("I12:I111").ClearContents
    Worksheets("Disbursement Schedule").Range("G12:G111").ClearContents
    
    'Clear Principal Schedule
    Sheets("Principal Schedule").Select
    Worksheets("Principal Schedule").Range("H13:AF372").ClearContents
    
    'Clear Fee Schedule
    Sheets("Other Fees Schedule").Select
    Worksheets("Other Fees Schedule").Range("H12:AF391").ClearContents
    
    'Clear Fee Sharing Schedule
    'Sheets("Fee Sharing Schedule").Select
    'Worksheets("Fee Sharing Schedule").Range("H12:AF391").ClearContents
    
    'Clear Foreign Currency Schedule
    'Sheets("FX Inputs").Select
    'Worksheets("FX Inputs").Range("H9:AF391").ClearContents
    
    Sheets("Inputs").Select
    Application.ScreenUpdating = True
End Sub

Sub FinalPassword()

    Dim password As String
    password = InputBox("Enter Password to Execute a Final Estimate")
    If password <> "budgetrocks" Then
        MsgBox ("Invalid Password for Final Run")
        End
    End If

End Sub

