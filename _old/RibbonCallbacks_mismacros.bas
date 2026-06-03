Attribute VB_Name = "RibbonCallbacks"
Option Explicit

'=============================================================================
' RIBBON CALLBACKS for mismacros.xlam
'=============================================================================
' These Subs are called by the custom ribbon buttons.
' Each wraps an existing macro and accepts IRibbonControl as required.
'
' INSTALLATION:
'   1. Open the .xlam in Excel (Alt+F11 to open VBA Editor)
'   2. File > Import File > select this .bas file
'   3. Save the .xlam
'   4. Inject customUI14_mismacros.xml with Office RibbonX Editor
'
' NOTE: The custom functions TASAPCT and EvalFormula are not included here
'       because they are worksheet functions, not macro buttons.
'=============================================================================

Public Sub RibbonChangeSeriesFormula(control As IRibbonControl)
    ChangeSeriesFormula
End Sub

Public Sub RibbonChangeSeriesFormulaAllCharts(control As IRibbonControl)
    ChangeSeriesFormulaAllCharts
End Sub

Public Sub RibbonDuplicateCurrentSheetMultipleTimes(control As IRibbonControl)
    DuplicateCurrentSheetMultipleTimes
End Sub

Public Sub RibbonAjustar_caja_graph(control As IRibbonControl)
    Ajustar_caja_graph
End Sub

Public Sub RibbonCopy_paste_absolute(control As IRibbonControl)
    Copy_paste_absolute
End Sub
