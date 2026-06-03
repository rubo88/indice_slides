Attribute VB_Name = "SectionNavBar"
Option Explicit

'=============================================================================
' SECTION NAVIGATION BAR - POWERPOINT ADD-IN VERSION
'=============================================================================
' Creates a navigation bar at the top of every slide using arrow/chevron
' shapes that match each defined section in the presentation.
'
' This version is designed to be packaged as a .ppam add-in with a
' custom ribbon tab. The ribbon callbacks accept IRibbonControl.
'=============================================================================

' Tag prefix used to identify nav bar shapes (for cleanup/re-run)
Private Const NAV_TAG As String = "SectionNav_"

' Constants (defined explicitly so the macro works without type library references)
Private Const SHAPE_HOME_PLATE As Long = 15   ' msoShapeHomePlate
Private Const SHAPE_CHEVRON As Long = 52      ' msoShapeChevron
Private Const msoTrue As Long = -1
Private Const msoFalse As Long = 0
Private Const msoAnchorMiddle As Long = 3
Private Const msoAutoSizeNone As Long = 0
Private Const msoAlignCenter As Long = 2
Private Const ppMouseClick As Long = 1

'=============================================================================
' RIBBON CALLBACKS
'=============================================================================
' These Subs are called by the custom ribbon buttons.
' Each accepts a single IRibbonControl parameter as required by the ribbon.
'=============================================================================

Public Sub RibbonCreateNavBar(control As IRibbonControl)
    CreateSectionNavBar
End Sub

Public Sub RibbonRemoveNavBar(control As IRibbonControl)
    RemoveSectionNavBar
End Sub

'=============================================================================
' CreateSectionNavBar - Main entry point
'=============================================================================
Sub CreateSectionNavBar()

    Dim pres As Presentation
    Dim sld As Slide
    Dim secProps As SectionProperties
    Dim i As Long

    Set pres = ActivePresentation
    Set secProps = pres.SectionProperties

    If secProps.Count = 0 Then
        MsgBox "This presentation has no sections defined." & vbCrLf & _
               "Add sections first (View > Slide Sorter, right-click > Add Section).", _
               vbExclamation, "No Sections Found"
        Exit Sub
    End If

    '=== Collect section data ===
    ' Skip empty sections, "Portada", and "Apéndice"/"Apendice".
    ' Portada slides get NO nav bar at all.
    ' Apéndice slides get the nav bar but with NO section highlighted.
    Dim secNames() As String
    Dim secFirstSlides() As Long
    Dim numSections As Long
    numSections = 0

    Dim portadaFirstSlide As Long: portadaFirstSlide = 0
    Dim portadaLastSlide As Long: portadaLastSlide = 0

    ' Track Apéndice slide ranges (nav bar shown, nothing highlighted)
    Dim apendiceFirstSlide As Long: apendiceFirstSlide = 0
    Dim apendiceLastSlide As Long: apendiceLastSlide = 0

    ReDim secNames(1 To secProps.Count)
    ReDim secFirstSlides(1 To secProps.Count)

    For i = 1 To secProps.Count
        If secProps.SlidesCount(i) > 0 Then
            Dim secNameLower As String
            secNameLower = LCase(secProps.Name(i))

            If secNameLower = "portada" Then
                portadaFirstSlide = secProps.FirstSlide(i)
                portadaLastSlide = portadaFirstSlide + secProps.SlidesCount(i) - 1

            ElseIf IsApendice(secNameLower) Then
                ' Track the range; if multiple Apéndice sections exist, expand range
                Dim aFirst As Long, aLast As Long
                aFirst = secProps.FirstSlide(i)
                aLast = aFirst + secProps.SlidesCount(i) - 1
                If apendiceFirstSlide = 0 Then
                    apendiceFirstSlide = aFirst
                    apendiceLastSlide = aLast
                Else
                    If aFirst < apendiceFirstSlide Then apendiceFirstSlide = aFirst
                    If aLast > apendiceLastSlide Then apendiceLastSlide = aLast
                End If

            Else
                numSections = numSections + 1
                secNames(numSections) = secProps.Name(i)
                secFirstSlides(numSections) = secProps.FirstSlide(i)
            End If
        End If
    Next i

    If numSections = 0 Then
        MsgBox "No displayable sections found.", vbExclamation, "No Sections"
        Exit Sub
    End If

    ReDim Preserve secNames(1 To numSections)
    ReDim Preserve secFirstSlides(1 To numSections)

    '=== Layout Configuration (in points) ===
    Dim slideW As Single: slideW = pres.PageSetup.SlideWidth

    Const NAV_TOP As Single = 11
    Const NAV_HEIGHT As Single = 15.5
    Const NAV_LEFT As Single = 15.5
    Const NAV_RIGHT As Single = 15.5
    Const BORDER_WEIGHT As Single = 1
    Const FONT_SIZE As Single = 10
    Const TEXT_MARGIN As Single = 2.835

    Dim chevronIndent As Single
    chevronIndent = NAV_HEIGHT

    ' Default palette (colored background slides)
    Dim BORDER_CLR As Long
    Dim FILL_CLR As Long
    Dim FONT_CLR As Long
    Dim INACTIVE_FILL_VISIBLE As Long
    Dim ACTIVE_FONT_CLR As Long

    '=== Calculate shape dimensions ===
    Dim totalWidth As Single
    totalWidth = slideW - NAV_LEFT - NAV_RIGHT

    Dim shapeW As Single
    If numSections > 1 Then
        shapeW = (totalWidth + chevronIndent * (numSections - 1)) / numSections
    Else
        shapeW = totalWidth
    End If

    '=== Process every slide ===
    For Each sld In pres.Slides

        DeleteNavShapes sld

        If portadaFirstSlide > 0 Then
            If sld.SlideIndex >= portadaFirstSlide And sld.SlideIndex <= portadaLastSlide Then
                GoTo NextSlide
            End If
        End If

        ' Determine which section this slide belongs to.
        ' Apéndice slides: show bar but highlight nothing (currentSec stays 0).
        Dim currentSec As Long: currentSec = 0
        Dim isApendiceSlide As Boolean: isApendiceSlide = False

        If apendiceFirstSlide > 0 Then
            If sld.SlideIndex >= apendiceFirstSlide And sld.SlideIndex <= apendiceLastSlide Then
                isApendiceSlide = True
            End If
        End If

        If Not isApendiceSlide Then
            For i = numSections To 1 Step -1
                If sld.SlideIndex >= secFirstSlides(i) Then
                    currentSec = i
                    Exit For
                End If
            Next i
        End If

        ' === Choose color palette based on slide background ===
        Dim slideIsWhite As Boolean
        slideIsWhite = IsWhiteBackground(sld)

        If slideIsWhite Then
            ' White-background palette: black borders/text, black highlight
            BORDER_CLR = RGB(0, 0, 0)
            FILL_CLR = RGB(0, 0, 0)           ' Active section: black fill
            FONT_CLR = RGB(0, 0, 0)            ' Inactive text: black
            ACTIVE_FONT_CLR = RGB(255, 255, 255) ' Active text: white
            INACTIVE_FILL_VISIBLE = msoTrue    ' Inactive: white fill visible
        Else
            ' Colored-background palette (original)
            BORDER_CLR = RGB(141, 164, 200)
            FILL_CLR = RGB(255, 255, 255)      ' Active section: white fill
            FONT_CLR = RGB(141, 164, 200)      ' Inactive text: blue
            ACTIVE_FONT_CLR = RGB(141, 164, 200) ' Active text: blue
            INACTIVE_FILL_VISIBLE = msoFalse   ' Inactive: transparent
        End If

        Dim shp As Shape
        Dim xPos As Single

        For i = 1 To numSections
            xPos = NAV_LEFT + (i - 1) * (shapeW - chevronIndent)

            Set shp = sld.Shapes.AddShape(SHAPE_CHEVRON, xPos, NAV_TOP, shapeW, NAV_HEIGHT)
            On Error Resume Next
            shp.Adjustments.Item(1) = 1 - (chevronIndent / shapeW)
            On Error GoTo 0

            shp.Name = NAV_TAG & i

            If i = currentSec Then
                ' Active section
                shp.Fill.Visible = msoTrue
                shp.Fill.Solid
                shp.Fill.ForeColor.RGB = FILL_CLR
            Else
                ' Inactive section
                If slideIsWhite Then
                    shp.Fill.Visible = msoTrue
                    shp.Fill.Solid
                    shp.Fill.ForeColor.RGB = RGB(255, 255, 255) ' White fill
                Else
                    shp.Fill.Visible = msoFalse  ' Transparent
                End If
            End If

            shp.Line.Visible = msoTrue
            shp.Line.ForeColor.RGB = BORDER_CLR
            shp.Line.Weight = BORDER_WEIGHT

            With shp.TextFrame2
                .MarginLeft = TEXT_MARGIN
                .MarginRight = TEXT_MARGIN
                .MarginTop = TEXT_MARGIN
                .MarginBottom = TEXT_MARGIN
                .VerticalAnchor = msoAnchorMiddle
                .AutoSize = msoAutoSizeNone
                .WordWrap = msoTrue
            End With

            shp.TextFrame2.TextRange.Text = i & ". " & secNames(i)

            With shp.TextFrame2.TextRange.ParagraphFormat
                .Alignment = msoAlignCenter
                .SpaceBefore = 0
                .SpaceAfter = 0
            End With

            With shp.TextFrame2.TextRange.Font
                .Size = FONT_SIZE
                .Bold = msoTrue
                If i = currentSec Then
                    .Fill.ForeColor.RGB = ACTIVE_FONT_CLR
                Else
                    .Fill.ForeColor.RGB = FONT_CLR
                End If
            End With

            Dim targetSlide As Slide
            Set targetSlide = pres.Slides(secFirstSlides(i))

            shp.ActionSettings(ppMouseClick).Hyperlink.SubAddress = _
                targetSlide.SlideID & "," & targetSlide.SlideIndex & ","

        Next i
NextSlide:
    Next sld

    MsgBox "Navigation bar created!" & vbCrLf & _
           numSections & " sections across " & pres.Slides.Count & " slides.", _
           vbInformation, "Section Nav Bar"

End Sub

'=============================================================================
' RemoveSectionNavBar
'=============================================================================
Sub RemoveSectionNavBar()

    Dim pres As Presentation
    Dim sld As Slide
    Dim totalRemoved As Long

    Set pres = ActivePresentation
    totalRemoved = 0

    For Each sld In pres.Slides
        totalRemoved = totalRemoved + DeleteNavShapes(sld)
    Next sld

    MsgBox totalRemoved & " navigation shapes removed from " & _
           pres.Slides.Count & " slides.", vbInformation, "Cleanup Complete"

End Sub

'=============================================================================
' HELPER FUNCTIONS
'=============================================================================

'-----------------------------------------------------------------------------
' IsWhiteBackground - Returns True if the slide background is white or
' very close to white (RGB each >= 250). Handles solid fills, theme
' backgrounds, and defaults to True if background is not filled
' (since PowerPoint's default is white).
'-----------------------------------------------------------------------------
Private Function IsWhiteBackground(sld As Slide) As Boolean
    Dim bgColor As Long
    Dim r As Long, g As Long, b As Long

    On Error GoTo Assume_White

    With sld.Background.Fill
        If .Visible = msoFalse Then
            ' No visible fill = default white
            IsWhiteBackground = True
            Exit Function
        End If
        bgColor = .ForeColor.RGB
    End With

    ' Extract RGB components
    r = bgColor Mod 256
    g = (bgColor \ 256) Mod 256
    b = (bgColor \ 65536) Mod 256

    ' Consider "white" if all channels >= 250 (allows near-white #FAFAFA etc.)
    IsWhiteBackground = (r >= 250 And g >= 250 And b >= 250)
    Exit Function

Assume_White:
    ' If any error reading background, assume white
    IsWhiteBackground = True
End Function

'-----------------------------------------------------------------------------
' IsApendice - Returns True if name matches "apéndice" or "apendice"
' (case-insensitive, expects input already lowered)
'-----------------------------------------------------------------------------
Private Function IsApendice(ByVal nameLower As String) As Boolean
    IsApendice = (nameLower = "apendice" Or nameLower = Chr(97) & "p" & Chr(233) & "ndice")
    ' Chr(233) = "é"  — ensures the accent is matched regardless of code page
End Function

Private Function DeleteNavShapes(sld As Slide) As Long

    Dim s As Shape
    Dim toDelete As New Collection
    Dim count As Long: count = 0

    For Each s In sld.Shapes
        If Left(s.Name, Len(NAV_TAG)) = NAV_TAG Then
            toDelete.Add s
            count = count + 1
        End If
    Next s

    Dim item As Variant
    For Each item In toDelete
        item.Delete
    Next item

    Set toDelete = Nothing
    DeleteNavShapes = count

End Function
