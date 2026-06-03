Attribute VB_Name = "Módulo1"
Sub ChangeSeriesFormulaAllCharts()
    ''' Do all charts in sheet
    Dim oChart As ChartObject
    Dim OldString As String, NewString As String
    Dim mySrs As Series

    OldString = InputBox("Enter the string to be replaced:", "Enter old string")

    If Len(OldString) > 1 Then
        NewString = InputBox("Enter the string to replace " & """" _
            & OldString & """:", "Enter new string")
        For Each oChart In ActiveSheet.ChartObjects
            For Each mySrs In oChart.Chart.SeriesCollection
                mySrs.Formula = WorksheetFunction.Substitute(mySrs.Formula, _
                    OldString, NewString)
            Next
        Next
    Else
        MsgBox "Nothing to be replaced.", vbInformation, "Nothing Entered"
    End If
End Sub

Sub ChangeSeriesFormula()
    ''' Just do active chart
    If ActiveChart Is Nothing Then
        '' There is no active chart
        MsgBox "Please select a chart and try again.", vbExclamation, _
            "No Chart Selected"
        Exit Sub
    End If

    Dim OldString As String, NewString As String, strTemp As String
    Dim mySrs As Series

    OldString = InputBox("Enter the string to be replaced:", "Enter old string")

    If Len(OldString) > 1 Then
        NewString = InputBox("Enter the string to replace " & """" _
            & OldString & """:", "Enter new string")
        '' Loop through all series
        For Each mySrs In ActiveChart.SeriesCollection
            strTemp = WorksheetFunction.Substitute(mySrs.Formula, _
                OldString, NewString)
            mySrs.Formula = strTemp
        Next
    Else
        MsgBox "Nothing to be replaced.", vbInformation, "Nothing Entered"
    End If
End Sub


Sub DuplicateCurrentSheetMultipleTimes()
    Dim wsTemplate As Worksheet
    Dim wsNew As Worksheet
    Dim namesSheet As Worksheet
    Dim lastRow As Long
    Dim i As Long

    ' Set the current active sheet as the template
    Set wsTemplate = ActiveWorkbook.ActiveSheet

    ' Set the sheet where the names are stored and the column with names (change as needed)
    Set namesSheet = ActiveWorkbook.Sheets("SheetNames")
    lastRow = namesSheet.Cells(namesSheet.Rows.Count, "A").End(xlUp).Row ' Assumes names are in column A

    ' Loop through the list of names and create duplicates
    Application.ScreenUpdating = False
    For i = 1 To lastRow
        wsTemplate.Copy After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count)
        Set wsNew = ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count)
        wsNew.Name = namesSheet.Cells(i, 1).Value ' Assumes names are in column A
    Next i
    Application.ScreenUpdating = True

    MsgBox lastRow & " sheets have been created based on the current sheet.", vbInformation
End Sub


Sub Ajustar_caja_graph()
    'ActiveChart.Axes(xlValue).AxisTitle.Select
   '  ActiveChart.Selection.InsideTop = 31
    ActiveChart.PlotArea.Select
    ActiveChart.PlotArea.InsideTop = 42 ' 42
    ActiveChart.PlotArea.InsideHeight = 179 ' 179 195
End Sub

