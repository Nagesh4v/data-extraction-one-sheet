Sub CC_Summary_Extraction()
'
'
' Prepare the workbook

tabs = ActiveWorkbook.Worksheets.Count
workbookName = ActiveWorkbook.Name
Application.DisplayAlerts = False
'
Do While tabs > 6
    Sheets(tabs).Delete
    tabs = tabs - 1
Loop
'
Application.DisplayAlerts = True
'
   Sheets("Result").Select
    Range("A2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    ActiveWorkbook.Save
'


    Application.ScreenUpdating = False
    Application.AskToUpdateLinks = False
    'countries = Array("Germany", "Austria", "Europe", "France", "Italy", "Spain", "Netherlands", "Belgium", "UK", "Switzerland", "USA")
    'countries = Array("Spain", "UK", "France", "Netherlands", "Belgium", "Germany")
     countries = Array("Austria", "USA", "Switzerland", "Europe", "Italy", "Spain", "UK", "France", "Netherlands", "Belgium", "Germany")

    'Iterate over all CC-Sheets
    For Each country In countries
    'Open next CC Sheet
        Filename = "Channel Controlling 2018 " & country & ".xlsx"
        Application.DisplayAlerts = False
        Workbooks.Open "Z:\800-Management\830-Controlling\833-Marketing\Channel Controlling 2018\" & Filename, UpdateLinks:=3
        Application.DisplayAlerts = True

     'Copy/Paste all the data from sheet Summary in CC sheet to sheet 1 in current workbook
        Sheets("Summary ").Select
        Cells.Select
        Selection.Copy
        Windows(workbookName).Activate
        Sheets("1").Select
        Cells.Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Application.CutCopyMode = False

    'Save and close CC sheet
        Windows(Filename).Activate
        Range("A1").Select
        'ActiveWorkbook.Save
        'ActiveWindow.Close
    
    'Close without saving
        ActiveWorkbook.Saved = True
        ActiveWorkbook.Close savechanges:=False

      'Run the first step
        Windows(workbookName).Activate
        Call Schritt1


    If country = "Switzerland" Then
     'Find the cell which contains "Other Sales Channels" in a particular column
        Columns("A:A").Select 'column to look in
        Set cell = Selection.Find(What:="Other Sales Channels", After:=ActiveCell, LookIn:=xlFormulas, _
            LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
            MatchCase:=False, SearchFormat:=False)

        'CHF Find the cell which contains "Other Sales Channels" in a particular column
        Set cell2 = Selection.Find(What:="Other Sales Channels", After:=ActiveCell, LookIn:=xlFormulas, _
            LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, _
            MatchCase:=False, SearchFormat:=False)

         'CHD part
            Range("A7").Select
            Range(Selection, Selection.End(xlToRight)).Select
            Range(Selection, cell.Offset(-2, 0)).Select

        Call Schritt2

          'CHF part
            lastTab = ActiveWorkbook.Worksheets.Count
            Sheets("2").Select
            Selection.Clear
            Worksheets(lastTab).Select
            cell.Offset(5, 0).Select
            Range(Selection, Selection.End(xlToRight)).Select
            Range(Selection, cell2.Offset(-3, 0)).Select

        Call Schritt2


        Else
        Columns("A:A").Select
    Set cell = Selection.Find(What:="Other Sales Channels", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)

        Range("A7").Select
        Range(Selection, Selection.End(xlToRight)).Select
        Range(Selection, cell.Offset(-3, 0)).Select
        Call Schritt2
     End If


    'Find the last empty cell in column D and select range from it untill the last cell with data in clumn C
        n1 = Range("D" & Rows.Count).End(xlUp).Row
        n1 = n1 + 1
        n2 = Range("C" & Rows.Count).End(xlUp).Row
        Range("D" & n1 & ":D" & n2).Select

     'Paste the country name in the previously selected range
        If country = "UK" Then
        Selection.Value = "England"
        Else
        Selection.Value = country
        End If

    Next country

    ActiveWorkbook.Save
        Application.ScreenUpdating = True
    Application.AskToUpdateLinks = True

    MsgBox ("Macro finished successfully")
End Sub

Sub Schritt1()
    Sheets("2").Select
    Range("A2:Y1000").Select
    Selection.ClearContents
  
  Sheets("1").Select
  Range( _
        "B:B,CE:CE,CK:CK,CQ:CQ,CW:CW,DC:DC,DI:DI,DO:DO,DU:DU,EA:EA,EG:EG,EM:EM,ES:ES" _
        ).Select

    Selection.Copy
    Sheets.Add After:=Sheets(Sheets.Count)
    ActiveSheet.Paste
    Range("A7").Select
End Sub


Sub Schritt2()

    Application.CutCopyMode = False
    Selection.Copy
    Sheets("2").Select
    Range("A2").Select
    ActiveSheet.Paste
    Sheets("3").Select
    ActiveSheet.PivotTables("PivotTable2").PivotCache.Refresh
    Range("A6:C6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Result").Select

    Range("A2").Select
    Selection.End(xlDown).Select
    If (ActiveCell.Value = "") Then
        Range("A2").Select
    Else: ActiveCell.Offset(1, 0).Range("A1").Select
    End If
    ActiveSheet.Paste
End Sub
