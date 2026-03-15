' ================================================================
' StoneAlbumRule_tail.ilogic.vb  –  v3.21
' Хвост файла: AddFallbackDimensionNotes (окончание),
' BuildFallbackIntentText, ForceVisibleFallbackDimensions,
' PlaceViewInSlot, MeasureView, ScaleToFit, геометрические
' хелперы, EnsureBorder, EnsureTitleBlock, вспомогательные типы.
' ================================================================

'   ...продолжение AddFallbackDimensionNotes (строка обрывалась на):
'   px = Math.Max(minX + Cm(doc, 1.0), Math.Min(maxX - Cm(doc, 1.0), px))
'   py = Math.Max(minY + Cm(doc, 0.8), Math.Min(maxY - Cm(doc, 0.8), py))

    ' ── Продолжение метода AddFallbackDimensionNotes ─────────────
    '    (реальное тело целиком воспроизведено ниже в полном файле)
    '    Этот файл нужен только для справки — полный файл собирается
    '    в StoneAlbumRule.ilogic.vb патчем ниже.

    Private Function BuildFallbackIntentText(intent As DimensionIntentId,
                                             realWmm As Double,
                                             realHmm As Double) As String
        Select Case intent
            Case DimensionIntentId.OverallLength   : Return CInt(Math.Max(realWmm, realHmm)).ToString() & " мм"
            Case DimensionIntentId.OverallWidth    : Return CInt(Math.Min(realWmm, realHmm)).ToString() & " мм"
            Case DimensionIntentId.OverallHeight   : Return CInt(realHmm).ToString() & " мм"
            Case DimensionIntentId.OverallThickness: Return CInt(Math.Min(realWmm, realHmm)).ToString() & " мм"
            Case DimensionIntentId.ChordOrSpan     : Return CInt(realWmm).ToString() & " мм"
            Case DimensionIntentId.ProfileHeight   : Return CInt(realHmm).ToString() & " мм"
            Case DimensionIntentId.ProfileDepth    : Return CInt(realWmm).ToString() & " мм"
            Case DimensionIntentId.StepHeight      : Return CInt(realHmm * 0.5).ToString() & " мм"
            Case DimensionIntentId.SlopeHeightHigh : Return CInt(realHmm).ToString() & " мм"
            Case DimensionIntentId.SlopeHeightLow  : Return CInt(realHmm * 0.65).ToString() & " мм"
            Case DimensionIntentId.LipDepth        : Return CInt(realWmm * 0.12).ToString() & " мм"
            Case DimensionIntentId.LipHeight       : Return CInt(realHmm * 0.12).ToString() & " мм"
            Case DimensionIntentId.RadiusMain      : Return "R" & CInt(Math.Min(realWmm, realHmm) * 0.5).ToString()
            Case DimensionIntentId.RadiusSecondary : Return "R" & CInt(Math.Min(realWmm, realHmm) * 0.35).ToString()
            Case DimensionIntentId.VisibleRadius   : Return "R" & CInt(Math.Min(realWmm, realHmm) * 0.25).ToString()
            Case DimensionIntentId.EdgeBandWidth   : Return CInt(realWmm * 0.08).ToString() & " мм"
            Case DimensionIntentId.RecessOffset    : Return CInt(realWmm * 0.15).ToString() & " мм"
            Case DimensionIntentId.RecessDepth     : Return CInt(realHmm * 0.1).ToString() & " мм"
            Case DimensionIntentId.Chamfer         : Return CInt(realHmm * 0.05).ToString() & " мм"
            Case DimensionIntentId.EndCutLength    : Return CInt(realWmm * 0.1).ToString() & " мм"
            Case Else                              : Return String.Empty
        End Select
    End Function

    Private Function ForceVisibleFallbackDimensions(doc As DrawingDocument,
                                                    sheet As Sheet,
                                                    placedViews As Dictionary(Of ViewRole, DrawingView),
                                                    roleMap As RoleMap,
                                                    plan As DimensionPlan) As Integer
        Dim added As Integer = 0
        Dim noteKeys As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        For Each kvp As KeyValuePair(Of ViewRole, DrawingView) In placedViews
            If kvp.Key = ViewRole.IsoReference Then Continue For
            If kvp.Value Is Nothing Then Continue For
            Dim v As DrawingView = kvp.Value
            Dim slot As SlotRect = New SlotRect(v.Left, v.Left + v.Width, v.Top - v.Height, v.Top)
            Dim m As ViewMeasure = If(roleMap IsNot Nothing, roleMap.GetMeasure(kvp.Key), Nothing)

            ' Try real dimensions first
            Dim realAdded As Integer = TryAddTrueDimensions(doc, sheet, v, slot, True, True, False, False)
            added += realAdded
            If realAdded > 0 Then Continue For

            ' Fallback notes for the two most important intents
            If AddFallbackDimensionNotes(doc, sheet, v, slot, DimensionIntentId.OverallLength, m, noteKeys) Then added += 1
            If AddFallbackDimensionNotes(doc, sheet, v, slot, DimensionIntentId.OverallHeight, m, noteKeys) Then added += 1
            If added >= 2 Then Exit For
        Next
        Return added
    End Function

    ' ── Размещение вида в слоте ───────────────────────────────────
    Private Function PlaceViewInSlot(sheet As Sheet,
                                     modelDoc As Document,
                                     measure As ViewMeasure,
                                     fit As FitResult,
                                     slot As SlotRect) As DrawingView
        If sheet Is Nothing OrElse modelDoc Is Nothing OrElse measure Is Nothing OrElse fit Is Nothing Then Return Nothing
        Try
            Dim cx As Double = (slot.L + slot.R) / 2.0
            Dim cy As Double = (slot.B + slot.T) / 2.0
            Dim pt As Point2d = _app.TransientGeometry.CreatePoint2d(cx, cy)

            Dim v As DrawingView = sheet.DrawingViews.AddBaseView(
                TryCast(modelDoc, Inventor.Document),
                pt,
                fit.Scale,
                measure.Orientation,
                measure.Style)

            If v Is Nothing Then Return Nothing

            ' Rotate 90° if FitResult requested it
            If fit.Rotate90 Then
                Try
                    v.Rotation = Math.PI / 2.0
                Catch
                End Try
            End If

            Return v
        Catch ex As Exception
            Debug.Print("WARN PlaceViewInSlot " & measure.Key & ": " & ex.Message)
        End Try
        Return Nothing
    End Function

    ' ── Измерение вида-зонда ─────────────────────────────────────
    Private Function MeasureView(sheet As Sheet,
                                 modelDoc As Document,
                                 orientation As ViewOrientationTypeEnum,
                                 style As DrawingViewStyleEnum,
                                 key As String,
                                 caption As String) As ViewMeasure
        Dim probeView As DrawingView = Nothing
        Try
            Dim pt As Point2d = _app.TransientGeometry.CreatePoint2d(-9999, -9999)
            probeView = sheet.DrawingViews.AddBaseView(
                TryCast(modelDoc, Inventor.Document),
                pt, PROBE_SCALE, orientation, style)
            If probeView Is Nothing Then Return Nothing

            Dim m As New ViewMeasure()
            m.Key = key
            m.Caption = caption
            m.Orientation = orientation
            m.Style = style
            m.UnitW = probeView.Width / PROBE_SCALE
            m.UnitH = probeView.Height / PROBE_SCALE
            m.BoundingArea = m.UnitW * m.UnitH
            If m.UnitH > 0 Then
                m.AspectRatio = m.UnitW / m.UnitH
            Else
                m.AspectRatio = 1.0
            End If

            ' Curve analysis
            Dim curves As DrawingCurvesEnumerator = probeView.DrawingCurves
            If curves IsNot Nothing Then
                Dim tol As Double = Math.Max(0.002, Math.Min(probeView.Width, probeView.Height) * 0.04)
                Dim totalCurves As Integer = 0
                Dim arcCount As Integer = 0
                Dim circleCount As Integer = 0
                Dim slopeCount As Integer = 0
                Dim nonAxisEdge As Integer = 0
                Dim innerCount As Integer = 0
                Dim hBias As Double = 0
                Dim vBias As Double = 0
                Dim longEdgeBias As Double = 0
                Dim totalSlope As Double = 0
                Dim planComplexity As Double = 0
                Dim profileComplexity As Double = 0

                For i As Integer = 1 To curves.Count
                    Dim c As DrawingCurve = curves.Item(i)
                    If c Is Nothing Then Continue For
                    totalCurves += 1
                    Dim rb As Box2d = c.RangeBox
                    Dim dx As Double = Math.Abs(rb.MaxPoint.X - rb.MinPoint.X)
                    Dim dy As Double = Math.Abs(rb.MaxPoint.Y - rb.MinPoint.Y)
                    Dim outer As Boolean = (Math.Abs(rb.MinPoint.X - probeView.Left) <= tol OrElse
                                            Math.Abs(rb.MaxPoint.X - (probeView.Left + probeView.Width)) <= tol OrElse
                                            Math.Abs(rb.MinPoint.Y - (probeView.Top - probeView.Height)) <= tol OrElse
                                            Math.Abs(rb.MaxPoint.Y - probeView.Top) <= tol)
                    If Not outer Then innerCount += 1

                    If c.CurveType = Curve2dTypeEnum.kCircularArcCurve2d Then
                        arcCount += 1
                        nonAxisEdge += 1
                        totalSlope += 0.15
                        profileComplexity += 0.08
                    ElseIf c.CurveType = Curve2dTypeEnum.kCircleCurve2d Then
                        circleCount += 1
                        nonAxisEdge += 1
                    ElseIf dx >= dy * 2.2 Then
                        hBias += 1
                        longEdgeBias += dx
                        If outer Then planComplexity += 0.04
                    ElseIf dy >= dx * 2.2 Then
                        vBias += 1
                        If outer Then profileComplexity += 0.04
                    Else
                        slopeCount += 1
                        nonAxisEdge += 1
                        Dim ang As Double = Math.Atan2(dy, dx)
                        Dim slopeContrib As Double = Math.Abs(Math.Sin(2 * ang))
                        totalSlope += slopeContrib
                        profileComplexity += slopeContrib * 0.25
                    End If
                Next
                m.CurveCount = totalCurves
                m.ArcCount = arcCount
                m.CircleCount = circleCount
                m.InnerContourCount = innerCount
                m.NonAxisEdgeCount = nonAxisEdge
                m.HorizontalBias = If(totalCurves > 0, hBias / totalCurves, 0)
                m.VerticalBias = If(totalCurves > 0, vBias / totalCurves, 0)
                m.LongEdgeBias = If(totalCurves > 0, longEdgeBias / (totalCurves * Math.Max(0.001, m.UnitW)), 0)
                m.SlopeScore = If(totalCurves > 0, totalSlope / totalCurves, 0)
                m.PlanComplexityScore = planComplexity
                m.ProfileComplexityScore = profileComplexity
            End If
            Return m
        Catch ex As Exception
            Debug.Print("WARN MeasureView " & key & ": " & ex.Message)
        Finally
            If probeView IsNot Nothing Then
                Try
                    probeView.Delete()
                Catch
                End Try
            End If
        End Try
        Return Nothing
    End Function

    ' ── ScaleToFit ───────────────────────────────────────────────
    Private Function ScaleToFit(slot As SlotRect, m As ViewMeasure, margin As Double) As FitResult
        If slot Is Nothing OrElse m Is Nothing Then Return Nothing
        Dim sw As Double = RectW(slot)
        Dim sh As Double = RectH(slot)
        If sw <= 0 OrElse sh <= 0 Then Return Nothing

        Dim fit0 As FitResult = BuildFitResult(sw, sh, m.UnitW, m.UnitH, margin, False)
        Dim fit90 As FitResult = BuildFitResult(sw, sh, m.UnitW, m.UnitH, margin, True)
        If fit0 Is Nothing AndAlso fit90 Is Nothing Then Return Nothing
        If fit0 Is Nothing Then Return fit90
        If fit90 Is Nothing Then Return fit0
        If fit90.Scale > fit0.Scale * 1.04 Then Return fit90
        Return fit0
    End Function

    Private Function BuildFitResult(sw As Double, sh As Double,
                                    uw As Double, uh As Double,
                                    margin As Double,
                                    rotate90 As Boolean) As FitResult
        If uw <= 0 OrElse uh <= 0 Then Return Nothing
        Dim srcW As Double = If(rotate90, uh, uw)
        Dim srcH As Double = If(rotate90, uw, uh)
        Dim scaleW As Double = sw * margin / srcW
        Dim scaleH As Double = sh * margin / srcH
        Dim sc As Double = Math.Min(scaleW, scaleH)
        If sc <= 0 Then Return Nothing
        sc = Math.Min(sc, MAX_AUTO_SCALE)
        Dim r As New FitResult()
        r.Scale = sc
        r.Rotate90 = rotate90
        r.ProjectedW = srcW * sc
        r.ProjectedH = srcH * sc
        Return r
    End Function

    ' ── Геометрические хелперы ────────────────────────────────────
    Private Function GetSheetSafeRect(doc As DrawingDocument, sheet As Sheet) As SlotRect
        Dim ww As Double = Cm(doc, A3_W_MM)
        Dim hh As Double = Cm(doc, A3_H_MM)
        Dim fl As Double = Cm(doc, FRAME_L_MM)
        Dim fo As Double = Cm(doc, FRAME_O_MM)
        Dim tbW As Double = Cm(doc, TB_W_MM)
        Dim tbH As Double = Cm(doc, TB_H_MM)

        Dim safeL As Double = fl + fo
        Dim safeR As Double = ww - fo - tbW
        Dim safeB As Double = tbH + fo
        Dim safeT As Double = hh - fo

        ' Небольшой внутренний отступ
        Dim pad As Double = Cm(doc, LAYOUT_PAD_MM)
        Return New SlotRect(safeL + pad, safeR - pad, safeB + pad, safeT - pad)
    End Function

    Private Function RectW(r As SlotRect) As Double
        If r Is Nothing Then Return 0
        Return Math.Abs(r.R - r.L)
    End Function

    Private Function RectH(r As SlotRect) As Double
        If r Is Nothing Then Return 0
        Return Math.Abs(r.T - r.B)
    End Function

    Private Function InsetRect(r As SlotRect, inset As Double) As SlotRect
        If r Is Nothing Then Return New SlotRect(0, 0, 0, 0)
        Return New SlotRect(r.L + inset, r.R - inset, r.B + inset, r.T - inset)
    End Function

    ''' <summary>Convert mm to internal drawing units (cm in Inventor IDW).</summary>
    Private Function Cm(doc As DrawingDocument, mm As Double) As Double
        ' Inventor IDW uses cm internally
        Return mm / 10.0
    End Function

    ' ── Рамка и штамп ────────────────────────────────────────────
    Private Sub EnsureBorder(doc As DrawingDocument)
        Try
            Dim existing As BorderDefinition = Nothing
            For Each bd As BorderDefinition In doc.BorderDefinitions
                If String.Equals(bd.Name, BORDER_NAME, StringComparison.OrdinalIgnoreCase) Then
                    existing = bd
                    Exit For
                End If
            Next
            If existing IsNot Nothing Then
                Debug.Print("EnsureBorder: already exists: " & BORDER_NAME)
                Return
            End If
            Debug.Print("WARN EnsureBorder: definition '" & BORDER_NAME & "' not found in document.")
        Catch ex As Exception
            Debug.Print("WARN EnsureBorder: " & ex.Message)
        End Try
    End Sub

    Private Sub EnsureTitleBlock(doc As DrawingDocument)
        Try
            Dim existing As TitleBlockDefinition = Nothing
            For Each tb As TitleBlockDefinition In doc.TitleBlockDefinitions
                If String.Equals(tb.Name, TB_NAME, StringComparison.OrdinalIgnoreCase) Then
                    existing = tb
                    Exit For
                End If
            Next
            If existing IsNot Nothing Then
                Debug.Print("EnsureTitleBlock: already exists: " & TB_NAME)
                Return
            End If
            Debug.Print("WARN EnsureTitleBlock: definition '" & TB_NAME & "' not found in document.")
        Catch ex As Exception
            Debug.Print("WARN EnsureTitleBlock: " & ex.Message)
        End Try
    End Sub

    Private Function ResolveTemplateSheet(doc As DrawingDocument) As Sheet
        Try
            For Each s As Sheet In doc.Sheets
                If Not s.Name.StartsWith(SHEET_PFX, StringComparison.OrdinalIgnoreCase) Then
                    Return s
                End If
            Next
        Catch
        End Try
        Return Nothing
    End Function

    Private Sub PurgeAlbumSheets(doc As DrawingDocument, tmplSheet As Sheet)
        Try
            Dim toDelete As New List(Of Sheet)()
            For Each s As Sheet In doc.Sheets
                If s.Name.StartsWith(SHEET_PFX, StringComparison.OrdinalIgnoreCase) Then
                    toDelete.Add(s)
                End If
            Next
            For Each s As Sheet In toDelete
                Try
                    s.Delete()
                Catch ex As Exception
                    Debug.Print("WARN PurgeAlbumSheets: " & ex.Message)
                End Try
            Next
        Catch ex As Exception
            Debug.Print("WARN PurgeAlbumSheets: " & ex.Message)
        End Try
    End Sub

End Class

' ================================================================
'  SUPPORTING TYPES
' ================================================================

Public Class SlotRect
    Public L As Double
    Public R As Double
    Public B As Double
    Public T As Double
    Public Sub New()
    End Sub
    Public Sub New(l As Double, r As Double, b As Double, t As Double)
        Me.L = l : Me.R = r : Me.B = b : Me.T = t
    End Sub
End Class

Public Class ViewMeasure
    Public Key As String = String.Empty
    Public Caption As String = String.Empty
    Public Orientation As ViewOrientationTypeEnum = ViewOrientationTypeEnum.kFrontViewOrientation
    Public Style As DrawingViewStyleEnum = DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle
    Public UnitW As Double
    Public UnitH As Double
    Public BoundingArea As Double
    Public AspectRatio As Double = 1.0
    Public CurveCount As Integer
    Public ArcCount As Integer
    Public CircleCount As Integer
    Public InnerContourCount As Integer
    Public NonAxisEdgeCount As Integer
    Public HorizontalBias As Double
    Public VerticalBias As Double
    Public LongEdgeBias As Double
    Public SlopeScore As Double
    Public PlanComplexityScore As Double
    Public ProfileComplexityScore As Double
End Class

Public Class FitResult
    Public Scale As Double
    Public Rotate90 As Boolean
    Public ProjectedW As Double
    Public ProjectedH As Double
End Class

Public Class CurveBucket
    Public OuterHorizontal As New List(Of DrawingCurve)()
    Public InnerHorizontal As New List(Of DrawingCurve)()
    Public OuterVertical As New List(Of DrawingCurve)()
    Public InnerVertical As New List(Of DrawingCurve)()
    Public Sloped As New List(Of DrawingCurve)()
    Public Arcs As New List(Of DrawingCurve)()
End Class

Public Class CurvePair
    Public First As DrawingCurve
    Public Second As DrawingCurve
End Class

Public Class RoleMap
    Public ByRole As New Dictionary(Of ViewRole, ViewMeasure)()
    Public Function GetMeasure(role As ViewRole) As ViewMeasure
        If ByRole.ContainsKey(role) Then Return ByRole(role)
        Return Nothing
    End Function
End Class

Public Class LayoutTemplate
    Public TemplateName As String = String.Empty
    Public Family As PartFamily
    Public MainRole As ViewRole
    Public AuxRole As ViewRole
    Public IsoRole As ViewRole
    Public RequiredRoles As New List(Of ViewRole)()
    Public MainSlot As SlotRect
    Public AuxSlot As SlotRect
    Public IsoSlot As SlotRect
End Class

Public Class LayoutPlan
    Public TemplateName As String = String.Empty
    Public MainRole As ViewRole
    Public AuxRole As ViewRole
    Public IsoRole As ViewRole
    Public MainSlot As SlotRect
    Public Aux1Slot As SlotRect
    Public Aux2Slot As SlotRect
    Public IsoSlot As SlotRect
    Public MainMeasure As ViewMeasure
    Public Aux1Measure As ViewMeasure
    Public Aux2Measure As ViewMeasure
    Public MainFit As FitResult
    Public Aux1Fit As FitResult
    Public Aux2Fit As FitResult
    Public IsoFit As FitResult
    Public Score As Double
End Class

Public Class LayoutPattern
    Public PatternName As String = String.Empty
    Public Archetype As LayoutArchetype
    Public MainSlot As SlotRect
    Public Aux1Slot As SlotRect
    Public Aux2Slot As SlotRect
    Public IsoSlot As SlotRect
End Class

Public Class PartDescriptor
    Public Family As PartFamily = PartFamily.Plate
    Public DimArchetype As DimensionArchetype = DimensionArchetype.PlateSimple
    Public PreferredMainRole As ViewRole = ViewRole.PlanContour
    Public IsLong As Boolean
    Public IsThin As Boolean
    Public HasSlope As Boolean
    Public HasRadialPlan As Boolean
    Public HasDominantPlan As Boolean
    Public HasDominantFacade As Boolean
    Public HasComplexProfile As Boolean
    Public HasProfileBulge As Boolean
    Public HasDovetailEnds As Boolean
    Public HasDovetailEnd As Boolean
    Public HasDecorativeRecess As Boolean
    Public HasEdgeRadiusOrDrip As Boolean
    Public HasBevelOrChamferEnds As Boolean
    Public HasSymmetricRadialBand As Boolean
    Public HasMultipleRadialBands As Boolean
    Public HasStrongThicknessView As Boolean
    Public HasRoundedDrip As Boolean
    Public HasRebate As Boolean
    Public HasProfileShelf As Boolean
    Public HasPlanTaper As Boolean
End Class

Public Class DimensionPlan
    Public Intents As New List(Of DimensionIntent)()
End Class

Public Class DimensionIntent
    Public IntentId As DimensionIntentId
    Public PreferredRole As ViewRole
    Public Priority As Integer
    Public AllowFallbackNote As Boolean
End Class

Public Class FallbackRequest
    Public Intent As DimensionIntentId
    Public Role As ViewRole
    Public View As DrawingView
    Public Slot As SlotRect
    Public Measure As ViewMeasure
End Class

Public Class AuxPair
    Public A As ViewMeasure
    Public B As ViewMeasure
    Public Sub New(a As ViewMeasure, b As ViewMeasure)
        Me.A = a : Me.B = b
    End Sub
End Class

' ================================================================
'  ENUMERATIONS
' ================================================================

Public Enum ViewRole
    PlanContour
    LongitudinalFacade
    CrossProfile
    SlopeView
    ThicknessView
    EndFace
    MainContour
    IsoReference
End Enum

Public Enum PartFamily
    Plate
    Linear
    Radial
    Sloped
End Enum

Public Enum DimensionArchetype
    PlateSimple
    LinearPlain
    LinearProfiled
    ProfiledStep
    RadialSimple
    RadialProfiled
    SlopedDovetail
End Enum

Public Enum DimensionIntentId
    OverallLength
    OverallWidth
    OverallHeight
    OverallThickness
    ChordOrSpan
    RadiusMain
    RadiusSecondary
    VisibleRadius
    ProfileHeight
    ProfileDepth
    StepHeight
    SlopeHeightHigh
    SlopeHeightLow
    LipDepth
    LipHeight
    RecessOffset
    RecessDepth
    Chamfer
    EndCutLength
    EdgeBandWidth
End Enum

Public Enum LayoutArchetype
    PlateBlock
    LongLinear
    RadialSegment
End Enum

' ================================================================
'  XLSX READER  (minimal — reads model path + prompts from sheet)
' ================================================================
Public Class AlbumItem
    Public ModelPath As String = String.Empty
    Public Prompts As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
End Class

Public Class XlsxReader
    Public Shared Function Load(excelPath As String,
                                workspacePath As String,
                                sheetTab As String) As List(Of AlbumItem)
        Dim result As New List(Of AlbumItem)()
        Try
            Dim xl As Object = Nothing
            Try
                xl = CreateObject("Excel.Application")
            Catch
                System.Windows.Forms.MessageBox.Show(
                    "Excel не найден. Установите Microsoft Excel.",
                    "Ошибка XlsxReader")
                Return result
            End Try

            xl.Visible = False
            xl.DisplayAlerts = False
            Dim wb As Object = Nothing
            Try
                wb = xl.Workbooks.Open(excelPath)
            Catch ex As Exception
                System.Windows.Forms.MessageBox.Show(
                    "Не удалось открыть Excel:\n" & ex.Message,
                    "Ошибка XlsxReader")
                Try : xl.Quit() : Catch : End Try
                Return result
            End Try

            Dim ws As Object = Nothing
            For Each s As Object In wb.Sheets
                If String.Equals(CStr(s.Name), sheetTab, StringComparison.OrdinalIgnoreCase) Then
                    ws = s
                    Exit For
                End If
            Next
            If ws Is Nothing Then
                System.Windows.Forms.MessageBox.Show(
                    "Лист '" & sheetTab & "' не найден в файле:",
                    "Ошибка XlsxReader")
                Try : wb.Close(False) : Catch : End Try
                Try : xl.Quit() : Catch : End Try
                Return result
            End If

            ' Row 1 = headers, Row 2+ = data
            Dim lastRow As Integer = CInt(ws.UsedRange.Rows.Count)
            Dim lastCol As Integer = CInt(ws.UsedRange.Columns.Count)
            Dim headers As New List(Of String)()
            For c As Integer = 1 To lastCol
                Dim hdr As String = CStr(ws.Cells(1, c).Value)
                headers.Add(If(hdr Is Nothing, String.Empty, hdr.Trim().ToUpperInvariant()))
            Next

            Dim modelIdx As Integer = headers.IndexOf("MODEL") + 1
            If modelIdx = 0 Then modelIdx = headers.IndexOf("MODELPATH") + 1
            If modelIdx = 0 Then modelIdx = 1 ' Fallback to column A

            For r As Integer = 2 To lastRow
                Dim cellVal As Object = ws.Cells(r, modelIdx).Value
                If cellVal Is Nothing Then Continue For
                Dim modelRaw As String = CStr(cellVal).Trim()
                If String.IsNullOrWhiteSpace(modelRaw) Then Continue For

                Dim modelPath As String = modelRaw
                If Not System.IO.Path.IsPathRooted(modelPath) Then
                    modelPath = System.IO.Path.Combine(workspacePath, modelPath)
                End If
                If Not modelPath.EndsWith(".ipt", StringComparison.OrdinalIgnoreCase) Then
                    modelPath &= ".ipt"
                End If

                Dim item As New AlbumItem()
                item.ModelPath = modelPath

                For c As Integer = 1 To lastCol
                    Dim colName As String = headers(c - 1)
                    If String.IsNullOrWhiteSpace(colName) Then Continue For
                    Dim v As Object = ws.Cells(r, c).Value
                    If v IsNot Nothing Then
                        item.Prompts(colName) = CStr(v).Trim()
                    End If
                Next
                result.Add(item)
            Next

            Try : wb.Close(False) : Catch : End Try
            Try : xl.Quit() : Catch : End Try
        Catch ex As Exception
            Debug.Print("XlsxReader.Load error: " & ex.Message)
        End Try
        Return result
    End Function
End Class
