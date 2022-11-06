Attribute VB_Name = "Module1"
Option Explicit
Sub stocks()

Windows("cdr stocks.xlsm").Activate
    Sheets("data").Visible = True

Windows("ecase data.xlsx").Activate
    Cells.Select
  
    Selection.Copy
  
    Windows("cdr stocks.xlsm").Activate
    Sheets("data").Select
    Cells.Select
    
    ActiveSheet.Paste
    

    
    
    ActiveWorkbook.Names.Add Name:="stocks", RefersToR1C1:= _
        "=OFFSET(data!R1C1,,,COUNTA(data!C1),53)"
    ActiveWorkbook.Names("stocks").Comment = ""
    
     ActiveWorkbook.RefreshAll
     
     Sheets("data").Select
    ActiveWindow.SelectedSheets.Visible = False

End Sub
