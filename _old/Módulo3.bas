Attribute VB_Name = "Módulo3"

Sub Copy_paste_absolute()
    '
    ' Macro6 Macro
    '
    Dim origSelection As Range
    Dim destCell As Range
    Dim destRange As Range

    ' Store the original selection
    Set origSelection = Selection

    ' Replace "=" with "@@@" in the original selection
    origSelection.Replace What:="=", Replacement:="@@@", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2

    ' Prompt user to select the destination cell
    Set destCell = Application.InputBox("Please select the destination cell:", "Destination Cell", Type:=8)
    If destCell Is Nothing Then Exit Sub  ' Handle Cancel button

    ' Copy the modified selection to the destination cell
    origSelection.Copy Destination:=destCell

    ' Define the destination range, same size as the original selection
    Set destRange = destCell.Resize(origSelection.Rows.Count, origSelection.Columns.Count)

    ' Replace "@@@" back to "=" in the original selection
    origSelection.Replace What:="@@@", Replacement:="=", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2

    ' Replace "@@@" back to "=" in the destination range
    destRange.Replace What:="@@@", Replacement:="=", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
End Sub




