Attribute VB_Name = "Module5"
Sub FEEDER()
'
' FeederLoadinglist Macro
'
' Keyboard Shortcut: Ctrl+Shift+H
'
    Selection.AutoFilter
    Columns("E:E").Select
    Selection.Replace What:="1", Replacement:="2", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="0", Replacement:="1", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Range("L3").Select
    Selection.Copy
    Range("A3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A4").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[3]&""-""&RC[11]"
    Range("A4").Select
    Selection.AutoFill Destination:=Range("A4:A83"), Type:=xlFillDefault
    Range("A4:A83").Select
    Range("A84").Select
    ActiveCell.FormulaR1C1 = "=RC[1]&""-""&RC[4]&""-""&RC[11]"
    Range("A84").Select
    Selection.AutoFill Destination:=Range("A84:A351"), Type:=xlFillDefault
    Range("A84:A351").Select
    Range("A351").Select
    ActiveWindow.SmallScroll Down:=-351
    Range("A4:A351").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("B:H").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    Columns("G:G").Select
    Selection.Delete Shift:=xlToLeft
    Columns("L:O").Select
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    Selection.Cut
    Columns("D:D").Select
    Selection.Insert Shift:=xlToRight
    Range("D3").Select
    ActiveCell.FormulaR1C1 = "Type"
    Range("E3").Select
    ActiveCell.FormulaR1C1 = "Size"
    Range("G3").Select
    ActiveCell.FormulaR1C1 = "Part Height"
    Columns("P:P").Select
    Selection.Cut
    Columns("I:I").Select
    Selection.Insert Shift:=xlToRight
    Columns("J:K").Select
    Selection.Delete Shift:=xlToLeft
    Columns("N:N").Select
    Selection.Delete Shift:=xlToLeft
    Columns("O:O").Select
    Selection.Delete Shift:=xlToLeft
    Range("J3").Select
    ActiveCell.FormulaR1C1 = "Tray Dir"
    Range("L3").Select
    Selection.Copy
    Range("K3").Select
    ActiveSheet.Paste
    Columns("L:L").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    Range("L3").Select
    ActiveCell.FormulaR1C1 = "Barcode Label"
    Range("M3").Select
    ActiveCell.FormulaR1C1 = "Reference"
    Range("J4:L352").Select
    Selection.ClearContents
    ActiveWindow.SmallScroll Down:=-21
    Rows("1:2").Select
    Range("B2").Activate
    Selection.Delete Shift:=xlUp
    Columns("M:M").Select
    Selection.TextToColumns Destination:=Range("M1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(1000, 1)), _
        TrailingMinusNumbers:=True
    ActiveSheet.Range("$A$1:$M$1048576").AutoFilter Field:=2, Criteria1:="<>"
    ActiveWindow.SmallScroll Down:=-24
    Range("B1").Select
    ActiveWindow.ScrollColumn = 1
    Range("A1:XFD1048576").Select
    ActiveWindow.SmallScroll Down:=-39
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Paste
    ActiveWindow.Zoom = 85
    ActiveWindow.Zoom = 70
    ActiveWindow.Zoom = 55
    ActiveWindow.Zoom = 40
    ActiveWindow.Zoom = 25
    ActiveWindow.Zoom = 40
    ActiveWindow.Zoom = 55
    ActiveWindow.Zoom = 70
    ActiveWindow.Zoom = 85
    Rows("1:1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$Z$55").AutoFilter Field:=9, Criteria1:="0"
    Rows("2:10000").Select
    Selection.Delete Shift:=xlUp
    ActiveSheet.Range("$A$1:$Z$38").AutoFilter Field:=9
    Selection.AutoFilter
    Range("T13").Select
       
   Dim FilePath As String

    FilePath = Application.GetSaveAsFilename

    ActiveWorkbook.SaveAs Filename:=FilePath & "_NXT.xlsx", FileFormat:=xlOpenXMLWorkbook
    
    Sheets("blank").Delete
    
    ActiveWorkbook.Save
    
    ActiveWorkbook.Close
    

End Sub

