Attribute VB_Name = "RibbonCallbacks"
Option Explicit

'=============================================================================
' RIBBON CALLBACKS for colores_tramas_oficiales_actualizacion.xlam
'=============================================================================
' These Subs are called by the custom ribbon buttons.
' Each wraps an existing macro and accepts IRibbonControl as required.
'
' INSTALLATION:
'   1. Open the .xlam in Excel (Alt+F11 to open VBA Editor)
'   2. File > Import File > select this .bas file
'   3. DELETE the old Módulo2 (which had GetCustomUI/AbrirMacro/CerrarMacro)
'   4. Save the .xlam
'   5. Inject customUI14_colores_tramas.xml with Office RibbonX Editor
'=============================================================================

Public Sub RibbonAplicadorDeTemas(control As IRibbonControl)
    AplicadorDeTemas
End Sub

Public Sub RibbonTema_FondoBlanco_AlGraficoActivo(control As IRibbonControl)
    Tema_FondoBlanco_AlGraficoActivo
End Sub

Public Sub RibbonTema_FondoAzul_AlGraficoActivo(control As IRibbonControl)
    Tema_FondoAzul_AlGraficoActivo
End Sub

Public Sub RibbonTema_FondoBlanco_A_Todos_LaHoja(control As IRibbonControl)
    Tema_FondoBlanco_A_Todos_LaHoja
End Sub

Public Sub RibbonTema_FondoAzul_A_Todos_LaHoja(control As IRibbonControl)
    Tema_FondoAzul_A_Todos_LaHoja
End Sub
