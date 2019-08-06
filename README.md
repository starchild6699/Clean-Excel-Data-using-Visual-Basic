# Clean-Excel-Data-using-Visual-Basic

The purpose of this project is to remove rows based on column value. Currently, inventory data is exported from a custom app into Excel.  In Excel, the data is manually filtered by searching aand deleting, running a recorded Macro and then more manual searching and deleting.  The goal is to write Visual Basic code that can combine the exising manual searches and deletions with the existing recorded Macro.

Requirements:
1.	When the exported data has opened as an Excel spreadsheet, manually sort the OS Column from Z to A and delete all rows with an OS.
4.	Run the recorded Macro
5.	Delete all rows with value $0 in colmun 
6.	Delete all rows with empty values in column 
7.	Delete all rows with value COA in column
8.	Delete all rows with value FBA in column
9.  Delete all rows with value Retail Sales in column
10.	Delete all rows with value Ebay Sales in column
11.	Delete all rows with value IT Staff in column
12. Delete all rows with value Windows 10 Pro in column
13. Delete all rows with value Windows 10 Home in column
14. After deleting rows above, delete columns labeled Reserved By, Reserved For and COA

When exported correctly, the Excel Columns will have the following labels:

' EXPORTED COLUMN VALUES
' ==
' A	  Unique #		        1	
' B	  Prod. Code		      2
' C	  Serial #		        3
' D	  Case			          4
' E	  OEM		        	    5
' F	  Model			          6
' G	  Wholesale Price	    7
' H	  Condition		        8
' I	  Warehouse (Current)	9	
' J	  Damages			        10
' K	  Configuration		    11
' L	  COA			            12	
' M	  Installed OS		    13	
' N	  Reserved By		      14
' O	  Reserved For		    15

Existing Macro which was recorded manually is below

Sub Lists()
'
' Lists Macro
'
' Keyboard Shortcut: Ctrl+l
'
    Cells.Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Sheet").Select
    Application.CutCopyMode = False
    ActiveWindow.SelectedSheets.Delete
    Columns("M:O").Select
    Selection.Delete Shift:=xlToLeft
    Columns("A:A").Select
    Columns("A:A").EntireColumn.AutoFit
    Columns("B:B").Select
    Columns("B:B").EntireColumn.AutoFit
    Columns("C:C").Select
    Columns("C:C").EntireColumn.AutoFit
    Columns("D:D").Select
    Columns("D:D").EntireColumn.AutoFit
    Columns("E:E").Select
    Columns("E:E").EntireColumn.AutoFit
    Columns("F:F").Select
    Columns("F:F").EntireColumn.AutoFit
    Columns("G:G").Select
    Columns("G:G").EntireColumn.AutoFit
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "Price"
    Columns("G:G").Select
    Columns("G:G").EntireColumn.AutoFit
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "Warehouse"
    Columns("I:I").Select
    Columns("I:I").EntireColumn.AutoFit
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    Columns("L:L").Select
    Selection.Cut
    Columns("H:H").Select
    Selection.Insert Shift:=xlToRight
    Columns("L:L").Select
    Selection.TextToColumns Destination:=Range("L1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=True, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), _
        Array(7, 1), Array(8, 1)), TrailingMinusNumbers:=True
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "Configuration"
    Range("M1").Select
    ActiveCell.FormulaR1C1 = "Hard Drive"
    Range("N1").Select
    ActiveCell.FormulaR1C1 = "RAM"
    Range("O1").Select
    ActiveCell.FormulaR1C1 = "Optical"
    Range("P1").Select
    ActiveCell.FormulaR1C1 = "Screen Size"
    Columns("Q:T").Select
    Selection.Delete Shift:=xlToLeft
    Columns("L:P").Select
    Range("P1").Activate
    Selection.Cut
    Columns("H:H").Select
    Selection.Insert Shift:=xlToRight
    Columns("H:H").Select
    Columns("F:F").ColumnWidth = 29.86
    Columns("F:F").ColumnWidth = 33.57
    Columns("H:H").Select
    Columns("H:H").EntireColumn.AutoFit
    Selection.ColumnWidth = 37.57
    Columns("I:I").Select
    Columns("I:I").EntireColumn.AutoFit
    Selection.ColumnWidth = 17.71
    Columns("J:J").Select
    Columns("J:J").EntireColumn.AutoFit
    Selection.ColumnWidth = 13.86
    Columns("K:K").Select
    Columns("K:K").EntireColumn.AutoFit
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 9
    Columns("M:M").Select
    Columns("M:M").EntireColumn.AutoFit
    Selection.ColumnWidth = 29.14
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Columns("G:G").Select
    Selection.NumberFormat = "$#,##0.00"
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Rows("1:1").Select
    With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
End 

End Sub

Three Macros that I have written and validated as working are below

Sub Delete_Ebay_Sales()

    Dim lRow As Long
    Dim iCntr As Long
    lRow = 30000
    For iCntr = lRow To 1 Step -1
        If Cells(iCntr, 15).Value = "Ebay Sales" Then
            Rows(iCntr).Delete
        End If
    Next
    End
 End Sub

Sub Delete_FBA()

    Dim lRow As Long
    Dim iCntr As Long
    lRow = 30000
    For iCntr = lRow To 1 Step -1
        If Cells(iCntr, 15).Value = "FBA" Then
            Rows(iCntr).Delete
        End If
    Next
    End
 End Sub

Sub Delete_Retail_Sales()

    Dim lRow As Long
    Dim iCntr As Long
    lRow = 30000
    For iCntr = lRow To 1 Step -1
        If Cells(iCntr, 15).Value = "Ebay Sales" Then
            Rows(iCntr).Delete
        End If
    Next
    End
End Sub







