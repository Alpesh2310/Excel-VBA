# Excel-VBA

Sub stu()

    Sheets("Entry").Select
    Range("B1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Database").Select
    Range("A1").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlUp).Select
    ActiveCell.Offset(1, 0).Select
    
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
    Sheets("Entry").Select
    Range("B1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("B6").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-3]C:R[-1]C)"
    Range("B1").Select
    
End Sub
