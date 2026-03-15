Attribute VB_Name = "SectionNavBar"
Option Explicit

'=============================================================================
' SECTION NAVIGATION BAR MACRO
'=============================================================================
' Creates a navigation bar at the top of every slide using arrow/chevron
' shapes that match each defined section in the presentation.
'
' - First shape: Pentagon arrow (homePlate - flat left, arrow right)
' - Subsequent shapes: Chevron (notch left, arrow right)
' - Current section is highlighted with white fill
' - All shapes are clickable hyperlinks to the first slide of each section
' - Text shows "N. SectionName" in bold
'
' Requirements: The presentation must have sections defined.
' Re-running the macro will delete previous nav bars and recreate them.
'
' USAGE:
'   1. Open the VBA Editor (Alt+F11)
'   2. Import this file (File > Import File)
'   3. Run "CreateSectionNavBar" from the Macros dialog (Alt+F8)
'   4. To remove: Run "RemoveSectionNavBar"
'
' CUSTOMIZATION:
'   Adjust the constants in the "Layout Configuration" section below
'   to change colors, sizes, margins, and font.
'=============================================================================

' Tag prefix used to identify nav bar shapes (for cleanup/re-run)
Private Const NAV_TAG As String = "SectionNav_"

' Constants (defined explicitly so the macro works without type library references)
' Shape types
Private Const SHAPE_HOME_PLATE As Long = 15   ' msoShapeHomePlate
Private Const SHAPE_CHEVRON As Long = 52      ' msoShapeChevron
' Tri-state / visibility
Private Const msoTrue As Long = -1
Private Const msoFalse As Long = 0
' Text anchoring
Private Const msoAnchorMiddle As Long = 3
Private Const msoAutoSizeNone As Long = 0
' Text alignment
Private Const msoAlignCenter As Long = 2
' Mouse click action
Private Const ppMouseClick As Long = 1

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

    '=== Collect section data (skip empty sections and "Portada") ===
    Dim secNames() As String
    Dim secFirstSlides() As Long
    Dim numSections As Long
    numSections = 0

    ' Track Portada slide ranges so we can skip them later
    Dim portadaFirstSlide As Long: portadaFirstSlide = 0
    Dim portadaLastSlide As Long: portadaLastSlide = 0

    ReDim secNames(1 To secProps.Count)
    ReDim secFirstSlides(1 To secProps.Count)

    For i = 1 To secProps.Count
        If secProps.SlidesCount(i) > 0 Then
            If LCase(secProps.Name(i)) = "portada" Then
                ' Remember Portada slide range but don't add to nav sections
                portadaFirstSlide = secProps.FirstSlide(i)
                portadaLastSlide = portadaFirstSlide + secProps.SlidesCount(i) - 1
            Else
                numSections = numSections + 1
                secNames(numSections) = secProps.Name(i)
                secFirstSlides(numSections) = secProps.FirstSlide(i)
            End If
        End If
    Next i

    If numSections = 0 Then
        MsgBox "No non-Portada sections found.", vbExclamation, "No Sections"
        Exit Sub
    End If

    ReDim Preserve secNames(1 To numSections)
    ReDim Preserve secFirstSlides(1 To numSections)

    '=== Layout Configuration (in points) ===========================
    ' Adjust these constants to customize the navigation bar appearance

    Dim slideW As Single: slideW = pres.PageSetup.SlideWidth

    Const NAV_TOP As Single = 11           ' Distance from top of slide
    Const NAV_HEIGHT As Single = 15.5      ' Height of each nav shape
    Const NAV_LEFT As Single = 15.5        ' Left margin
    Const NAV_RIGHT As Single = 15.5       ' Right margin
    Const BORDER_WEIGHT As Single = 1      ' Border line weight
    Const FONT_SIZE As Single = 10         ' Text font size
    Const TEXT_MARGIN As Single = 2.835    ' Internal text margin (~1mm)

    ' Chevron indent depth (notch/arrow size on chevron shapes)
    ' Original uses ~0.61 x height. Full height gives a clear ~45-degree point.
    Dim chevronIndent As Single
    chevronIndent = NAV_HEIGHT          ' = full bar height

    ' Colors (change these to match your theme)
    Dim BORDER_CLR As Long: BORDER_CLR = RGB(141, 164, 200)  ' #8DA4C8
    Dim FILL_CLR As Long: FILL_CLR = RGB(255, 255, 255)      ' White highlight

    '=== Calculate shape dimensions ===
    Dim totalWidth As Single
    totalWidth = slideW - NAV_LEFT - NAV_RIGHT

    ' Shapes overlap by chevronIndent so each notch interlocks with the previous arrow
    Dim shapeW As Single
    If numSections > 1 Then
        shapeW = (totalWidth + chevronIndent * (numSections - 1)) / numSections
    Else
        shapeW = totalWidth
    End If

    '=== Process every slide ===
    For Each sld In pres.Slides

        ' --- Always clean up existing nav shapes (in case sections changed) ---
        DeleteNavShapes sld

        ' --- Skip Portada slides: no nav bar on those ---
        If portadaFirstSlide > 0 Then
            If sld.SlideIndex >= portadaFirstSlide And sld.SlideIndex <= portadaLastSlide Then
                GoTo NextSlide
            End If
        End If

        ' --- Determine which section this slide belongs to ---
        Dim currentSec As Long: currentSec = 0
        For i = numSections To 1 Step -1
            If sld.SlideIndex >= secFirstSlides(i) Then
                currentSec = i
                Exit For
            End If
        Next i

        ' --- Create one shape per section ---
        Dim shp As Shape
        Dim xPos As Single

        For i = 1 To numSections
            xPos = NAV_LEFT + (i - 1) * (shapeW - chevronIndent)

            ' Use chevron preset for all shapes (consistent look, proper text frames)
            Set shp = sld.Shapes.AddShape(SHAPE_CHEVRON, xPos, NAV_TOP, shapeW, NAV_HEIGHT)
            ' Adj = 1 - indent/width: higher value = smaller notch/arrow
            On Error Resume Next
            shp.Adjustments.Item(1) = 1 - (chevronIndent / shapeW)
            On Error GoTo 0

            shp.Name = NAV_TAG & i

            ' --- Fill: white for current section, none for others ---
            If i = currentSec Then
                shp.Fill.Visible = msoTrue
                shp.Fill.Solid
                shp.Fill.ForeColor.RGB = FILL_CLR
            Else
                shp.Fill.Visible = msoFalse
            End If

            ' --- Border ---
            shp.Line.Visible = msoTrue
            shp.Line.ForeColor.RGB = BORDER_CLR
            shp.Line.Weight = BORDER_WEIGHT

            ' --- Text frame setup ---
            With shp.TextFrame2
                .MarginLeft = TEXT_MARGIN
                .MarginRight = TEXT_MARGIN
                .MarginTop = TEXT_MARGIN
                .MarginBottom = TEXT_MARGIN
                .VerticalAnchor = msoAnchorMiddle
                .AutoSize = msoAutoSizeNone
                .WordWrap = msoTrue
            End With

            ' --- Text content ---
            shp.TextFrame2.TextRange.Text = i & ". " & secNames(i)

            With shp.TextFrame2.TextRange.ParagraphFormat
                .Alignment = msoAlignCenter
                .SpaceBefore = 0
                .SpaceAfter = 0
            End With

            With shp.TextFrame2.TextRange.Font
                .Size = FONT_SIZE
                .Bold = msoTrue
                .Fill.ForeColor.RGB = BORDER_CLR
            End With

            ' --- Hyperlink to first slide of this section ---
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
' RemoveSectionNavBar - Removes all navigation bar shapes from all slides
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
' DeleteNavShapes - Removes all nav bar shapes from a single slide
' Returns the count of shapes deleted
'-----------------------------------------------------------------------------
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
