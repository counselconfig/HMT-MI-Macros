Attribute VB_Name = "Module1"
Option Explicit


Sub Macro2()
'
' Macro2 Macro
'



'running labels data
  Windows("enhanced.xlsx").Activate
    Sheets(Array("FOI", "Parli")).Select
    Sheets("Parli").Activate
    ActiveWindow.SelectedSheets.Delete
    Selection.AutoFilter
    Range("M1").Select
    ActiveSheet.Range("A:BB").AutoFilter Field:=2, Criteria1:= _
        "Ministerial Correspondence"
    Range("R1").Select
    ActiveSheet.Range("A:BB").AutoFilter Field:=25, Criteria1:=Array( _
        "Open", "Open (Reopened - case data update)", _
        "Open (Reopened - case processing restarted)"), Operator:=xlFilterValues
    Range("K1").Select
    ActiveSheet.Range("A:BB").AutoFilter Field:=24, Criteria1:= _
        "Print / Prepare"
'    ActiveWindow.SmallScroll ToRight:=-10
'    Columns("B:J").Select
'    Selection.EntireColumn.Hidden = True
'    Columns("L:AC").Select
'    Selection.EntireColumn.Hidden = True
    
    Columns("A:A").Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Paste
    Sheets("Correspondence").Select
    Columns("AB:AB").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Sheet1").Select
    Range("B1").Select
    ActiveSheet.Paste
    Sheets("Correspondence").Select
    Columns("n:n").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Sheet1").Select
    Columns("c:c").Select
    ActiveSheet.Paste
     Columns("A:A").EntireColumn.AutoFit
    Columns("B:B").EntireColumn.AutoFit
    Columns("C:C").EntireColumn.AutoFit
    Sheets("correspondence").Select
    Application.CutCopyMode = False
    ActiveWindow.SelectedSheets.Delete
 
    
    
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp

    Sheets.Add After:=ActiveSheet
    Sheets("Sheet2").Select
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "=INT(COUNTA(Sheet1!C)/2)"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "=OFFSET(Sheet1!R[-2]C,R[-2]C,0)"
    

    
      
      
      Range("A3").Select
    Selection.Copy
 
    
    Range("A4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
      

    
    Cells.Find(What:=Range("a4"), After:=ActiveCell, LookIn:=xlValues, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    Cells.FindNext(After:=ActiveCell).Activate
    Sheets("Sheet1").Select
    Cells.FindNext(After:=ActiveCell).Activate
    
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    
    Range("D1").Select

    ActiveSheet.Paste
    ActiveWindow.View = xlNormalView
    ActiveWindow.View = xlPageLayoutView
    
    Selection.ColumnWidth = 13.4
    ActiveWindow.View = xlPageBreakPreview
    Selection.RowHeight = 175.1
    Sheets("Sheet2").Select
    ActiveWindow.SelectedSheets.Delete
    
    Cells.Select
    ActiveWindow.View = xlNormalView
    ActiveWindow.View = xlPageLayoutView
    Selection.ColumnWidth = 13.4
    
    With Selection
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    Columns("C:C").ColumnWidth = 17.33
    Cells.Select
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = ""
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 100
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    Application.PrintCommunication = True
    ActiveWindow.SmallScroll Down:=0
    With Selection
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With

'

'
  
  Columns("C:C").ColumnWidth = 22.13
    Columns("A:A").ColumnWidth = 15.33
    Columns("B:B").Select
    Selection.ColumnWidth = 21.87
    Columns("C:C").Select
    Selection.ColumnWidth = 12.2

        End Sub



