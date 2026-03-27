' IsoDimensionsFromIptRule.ilogic.vb
' Отдельное iLogic-правило:
' после расстановки видов на листах добавляет/обновляет подписи
' с габаритами модели из исходного .ipt/.iam рядом с изометрическими видами.

Option Explicit On

Imports Inventor
Imports System
Imports System.Collections.Generic

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    Dim doc As DrawingDocument = TryCast(app.ActiveDocument, DrawingDocument)
    If doc Is Nothing Then
        System.Windows.Forms.MessageBox.Show("Откройте .idw документ.", "IsoDimensionsFromIptRule")
        Return
    End If

    Dim processed As Integer = 0
    Dim skipped As Integer = 0
    Dim issueLog As New List(Of String)()
    Dim techAllowancePerSideMm As Double = ResolveTechAllowancePerSideMm()

    For Each sh As Sheet In doc.Sheets
        For Each v As DrawingView In sh.DrawingViews
            If Not IsIsometric(v) Then Continue For

            Dim openedHere As Boolean = False
            Dim modelDoc As Document = ResolveModelDocument(app, v, openedHere, issueLog)
            If modelDoc Is Nothing Then
                skipped += 1
                Continue For
            End If

            Try
                Dim mm As ModelOverallExtents = GetModelOverallExtentsMm(modelDoc)
                Dim noteText As String = BuildIsoDimNoteText(v, modelDoc, mm, techAllowancePerSideMm)
                UpsertIsoDimNote(doc, sh, v, noteText)
                processed += 1
            Catch ex As Exception
                skipped += 1
                issueLog.Add("Ошибка обработки вида '" & SafeViewName(v) & "': " & ex.Message)
            Finally
                If openedHere Then
                    Try
                        modelDoc.Close(True)
                    Catch exClose As Exception
                        issueLog.Add("Не удалось закрыть модель: " & SafeModelPath(modelDoc) & " | " & exClose.Message)
                    End Try
                End If
            End Try
        Next
    Next

    doc.Update2(True)

    Dim msg As String = "Изометрические виды обработаны: " & processed.ToString() & vbCrLf &
                        "Пропущено: " & skipped.ToString() & vbCrLf & vbCrLf &
                        "Примечание: правило добавляет чистовые/заготовительные L×W×H и массу." & vbCrLf &
                        "Габариты рассчитываются через orientation-agnostic minimum bounding rectangle (XY)."

    If issueLog.Count > 0 Then
        msg &= vbCrLf & vbCrLf & "Проблемные модели/виды:" & vbCrLf
        For i As Integer = 0 To Math.Min(issueLog.Count - 1, 20)
            msg &= "• " & issueLog(i) & vbCrLf
        Next
        If issueLog.Count > 20 Then
            msg &= "... и ещё " & (issueLog.Count - 20).ToString() & " записей."
        End If
    End If

    System.Windows.Forms.MessageBox.Show(msg, "IsoDimensionsFromIptRule")
End Sub

Private Function IsIsometric(v As DrawingView) As Boolean
    If v Is Nothing Then Return False
    ' Исправлено: свойство должно называться ViewOrientationType
    Select Case v.ViewOrientationType
        Case ViewOrientationTypeEnum.kIsoTopLeftViewOrientation,
             ViewOrientationTypeEnum.kIsoTopRightViewOrientation,
             ViewOrientationTypeEnum.kIsoBottomLeftViewOrientation,
             ViewOrientationTypeEnum.kIsoBottomRightViewOrientation
            Return True
        Case Else
            Return False
    End Select
End Function

Private Function ResolveModelDocument(app As Inventor.Application,
                                      v As DrawingView,
                                      ByRef openedHere As Boolean,
                                      issueLog As List(Of String)) As Document
    openedHere = False
    If v Is Nothing Then Return Nothing

    Dim desc As DocumentDescriptor = Nothing
    Try
        desc = v.ReferencedDocumentDescriptor
    Catch ex As Exception
        issueLog.Add("Вид '" & SafeViewName(v) & "': нет ReferencedDocumentDescriptor | " & ex.Message)
    End Try
    If desc Is Nothing Then Return Nothing

    Try
        If desc.ReferencedDocument IsNot Nothing Then Return desc.ReferencedDocument
    Catch ex As Exception
        issueLog.Add("Вид '" & SafeViewName(v) & "': ReferencedDocument недоступен | " & ex.Message)
    End Try

    Dim fullPath As String = String.Empty
    Try
        fullPath = desc.FullDocumentName
    Catch ex As Exception
        issueLog.Add("Вид '" & SafeViewName(v) & "': невозможно получить путь модели | " & ex.Message)
    End Try

    If String.IsNullOrWhiteSpace(fullPath) Then
        issueLog.Add("Вид '" & SafeViewName(v) & "': пустой путь к модели.")
        Return Nothing
    End If

    If Not System.IO.File.Exists(fullPath) Then
        issueLog.Add("Файл модели не найден: " & fullPath)
        Return Nothing
    End If

    Dim openAttempts As Integer = 3
    For attempt As Integer = 1 To openAttempts
        Try
            Dim wasSilent As Boolean = app.SilentOperation
            Try
                app.SilentOperation = True
                Dim opened As Document = app.Documents.Open(fullPath, False)
                openedHere = (opened IsNot Nothing)
                If opened IsNot Nothing Then Return opened
            Finally
                app.SilentOperation = wasSilent
            End Try
        Catch ex As Exception
            If attempt < openAttempts Then
                issueLog.Add("Повтор открытия " & attempt.ToString() & "/" & openAttempts.ToString() & ": " & fullPath)
                System.Windows.Forms.Application.DoEvents()
                System.Threading.Thread.Sleep(200)
            Else
                issueLog.Add("Не удалось открыть модель после " & openAttempts.ToString() & " попыток: " & fullPath & " | " & ex.Message)
            End If
        End Try
    Next

    Return Nothing
End Function

Private Sub UpsertIsoDimNote(doc As DrawingDocument,
                             sh As Sheet,
                             v As DrawingView,
                             txt As String)
    If sh Is Nothing OrElse v Is Nothing Then Return

    Dim tag As String = BuildViewTag(v)
    Dim existing As GeneralNote = FindTaggedIsoNote(sh, tag)
    If existing IsNot Nothing Then
        Try
            existing.Text = txt
            Return
        Catch
            Try
                existing.Delete()
            Catch
            End Try
        End Try
    End If

    Dim tg As TransientGeometry = ThisApplication.TransientGeometry
    Dim defaultX As Double = v.Left + v.Width + MmToCm(6.0)
    Dim defaultY As Double = v.Top - MmToCm(6.0)
    Dim margin As Double = MmToCm(6.0)

    Dim safeX As Double = defaultX
    Dim safeY As Double = defaultY

    Try
        Dim maxX As Double = Math.Max(margin, sh.Width - margin)
        Dim minX As Double = margin
        If safeX > maxX Then safeX = maxX
        If safeX < minX Then safeX = minX

        Dim maxY As Double = Math.Max(margin, sh.Height - margin)
        Dim minY As Double = margin
        If safeY > maxY Then safeY = maxY
        If safeY < minY Then safeY = minY
    Catch
    End Try

    Dim p As Point2d = tg.CreatePoint2d(safeX, safeY)
    Dim note As GeneralNote = sh.DrawingNotes.GeneralNotes.AddFitted(p, txt)
    TagIsoNote(note, tag)

    Try
        Dim style As TextStyle = Nothing
        Try
            style = doc.StylesManager.TextStyles.Item("RKM_SPDS_Title_1_4mm")
        Catch
        End Try
        If style IsNot Nothing Then note.TextStyle = style
    Catch
    End Try
End Sub

Private Function BuildViewTag(v As DrawingView) As String
    Dim key As String = "ISO|"
    Try
        key &= v.Name
    Catch
    End Try
    Return key
End Function

Private Sub TagIsoNote(n As GeneralNote, tag As String)
    If n Is Nothing Then Return
    Try
        Dim sets As AttributeSets = n.AttributeSets
        Dim set1 As AttributeSet = Nothing
        If sets.NameIsUsed("RKM_ISO_DIMS") Then
            set1 = sets.Item("RKM_ISO_DIMS")
        Else
            set1 = sets.Add("RKM_ISO_DIMS")
        End If

        If set1.NameIsUsed("ViewTag") Then
            set1.Item("ViewTag").Value = tag
        Else
            set1.Add("ViewTag", ValueTypeEnum.kStringType, tag)
        End If
    Catch
    End Try
End Sub

Private Function FindTaggedIsoNote(sh As Sheet, tag As String) As GeneralNote
    If sh Is Nothing Then Return Nothing

    For Each n As GeneralNote In sh.DrawingNotes.GeneralNotes
        Try
            If n.AttributeSets.NameIsUsed("RKM_ISO_DIMS") Then
                Dim set1 As AttributeSet = n.AttributeSets.Item("RKM_ISO_DIMS")
                If set1.NameIsUsed("ViewTag") Then
                    Dim v As String = ""
                    Try
                        v = CStr(set1.Item("ViewTag").Value)
                    Catch
                    End Try
                    If String.Equals(v, tag, StringComparison.OrdinalIgnoreCase) Then
                        Return n
                    End If
                End If
            End If
        Catch
        End Try
    Next

    Return Nothing
End Function

Private Function BuildIsoDimNoteText(v As DrawingView,
                                     modelDoc As Document,
                                     m As ModelOverallExtents,
                                     techAllowancePerSideMm As Double) As String
    Dim modelName As String = ""
    Try
        modelName = System.IO.Path.GetFileNameWithoutExtension(modelDoc.FullFileName)
    Catch
    End Try
    If String.IsNullOrWhiteSpace(modelName) Then modelName = "Модель"

    Dim viewName As String = "Изометрия"
    Try
        If Not String.IsNullOrWhiteSpace(v.Name) Then viewName = v.Name
    Catch
    End Try

    Dim allowanceDelta As Double = Math.Max(0.0, techAllowancePerSideMm) * 2.0

    Return viewName & " (" & modelName & ")" & vbCrLf &
           "Чистовые, мм: L=" & FormatMm(m.LengthMm) & "  W=" & FormatMm(m.WidthMm) & "  H=" & FormatMm(m.HeightMm) & vbCrLf &
           "Заготовка (+" & FormatMm(techAllowancePerSideMm) & " мм/сторона): L=" & FormatMm(m.LengthMm + allowanceDelta) &
           "  W=" & FormatMm(m.WidthMm + allowanceDelta) & "  H=" & FormatMm(m.HeightMm) & vbCrLf &
           "Масса (" & FormatMm(m.DensityKgM3) & " кг/м³): " & FormatMm(m.MassKg) & " кг"
End Function

Private Function GetModelOverallExtentsMm(modelDoc As Document) As ModelOverallExtents
    Dim res As New ModelOverallExtents()
    If modelDoc Is Nothing Then Return res

    Dim points As List(Of Point2dData) = GetModelPlanPointsMm(modelDoc)
    If points.Count >= 3 Then
        Dim mmr As MinAreaRect2d = ComputeMinimumBoundingRectangle(points)
        res.LengthMm = Math.Max(mmr.WidthMm, mmr.HeightMm)
        res.WidthMm = Math.Min(mmr.WidthMm, mmr.HeightMm)
    End If

    Dim rb As Box = Nothing
    If TypeOf modelDoc Is PartDocument Then
        rb = DirectCast(modelDoc, PartDocument).ComponentDefinition.RangeBox
    ElseIf TypeOf modelDoc Is AssemblyDocument Then
        rb = DirectCast(modelDoc, AssemblyDocument).ComponentDefinition.RangeBox
    Else
        Return res
    End If

    If rb Is Nothing Then Return res

    Dim dz As Double = Math.Abs(rb.MaxPoint.Z - rb.MinPoint.Z) * 10.0
    res.HeightMm = dz

    If res.LengthMm <= 0 OrElse res.WidthMm <= 0 Then
        Dim dx As Double = Math.Abs(rb.MaxPoint.X - rb.MinPoint.X) * 10.0
        Dim dy As Double = Math.Abs(rb.MaxPoint.Y - rb.MinPoint.Y) * 10.0
        Dim vals() As Double = New Double() {dx, dy}
        Array.Sort(vals)
        Array.Reverse(vals)
        res.LengthMm = vals(0)
        res.WidthMm = vals(1)
    End If

    res.DensityKgM3 = ResolveStoneDensityKgM3()
    res.MassKg = ComputeMassKg(modelDoc, res.DensityKgM3)
    Return res
End Function

Private Function ComputeMassKg(modelDoc As Document, densityKgM3 As Double) As Double
    Try
        Dim mp As MassProperties = Nothing
        If TypeOf modelDoc Is PartDocument Then
            mp = DirectCast(modelDoc, PartDocument).ComponentDefinition.MassProperties
        ElseIf TypeOf modelDoc Is AssemblyDocument Then
            mp = DirectCast(modelDoc, AssemblyDocument).ComponentDefinition.MassProperties
        End If

        If mp Is Nothing Then Return 0.0

        Dim volumeCm3 As Double = mp.Volume
        If volumeCm3 <= 0.0 Then Return 0.0

        Dim volumeM3 As Double = volumeCm3 / 1000000.0
        Return volumeM3 * Math.Max(0.0, densityKgM3)
    Catch
        Return 0.0
    End Try
End Function

Private Function ResolveStoneDensityKgM3() As Double
    Dim fallbackDensity As Double = 2800.0
    Dim valueText As String = ""
    Try
        valueText = CStr(iProperties.Value("Custom", "StoneDensityKgM3"))
    Catch
    End Try

    Dim parsed As Double = 0.0
    If Double.TryParse(valueText, parsed) Then
        If parsed >= 1500.0 AndAlso parsed <= 5000.0 Then Return parsed
    End If

    Return fallbackDensity
End Function

Private Function ResolveTechAllowancePerSideMm() As Double
    Dim fallbackMm As Double = 5.0

    Dim valueText As String = ""
    Try
        valueText = CStr(iProperties.Value("Custom", "TechAllowancePerSideMm"))
    Catch
    End Try

    Dim parsed As Double = 0.0
    If Double.TryParse(valueText, parsed) AndAlso parsed >= 0 Then
        Return parsed
    End If

    Return fallbackMm
End Function

Private Function GetModelPlanPointsMm(modelDoc As Document) As List(Of Point2dData)
    Dim result As New List(Of Point2dData)()
    Try
        If TypeOf modelDoc Is PartDocument Then
            Dim pDoc As PartDocument = DirectCast(modelDoc, PartDocument)
            For Each b As SurfaceBody In pDoc.ComponentDefinition.SurfaceBodies
                For Each vtx As Vertex In b.Vertices
                    result.Add(New Point2dData(vtx.Point.X * 10.0, vtx.Point.Y * 10.0))
                Next
            Next
        ElseIf TypeOf modelDoc Is AssemblyDocument Then
            Dim aDoc As AssemblyDocument = DirectCast(modelDoc, AssemblyDocument)
            For Each occ As ComponentOccurrence In aDoc.ComponentDefinition.Occurrences
                Try
                    Dim orb As Box = occ.RangeBox
                    AddRangeBoxPlanCorners(result, orb)
                Catch
                End Try
            Next
        End If
    Catch
    End Try
    Return result
End Function

Private Sub AddRangeBoxPlanCorners(target As List(Of Point2dData), rb As Box)
    If target Is Nothing OrElse rb Is Nothing Then Return
    target.Add(New Point2dData(rb.MinPoint.X * 10.0, rb.MinPoint.Y * 10.0))
    target.Add(New Point2dData(rb.MaxPoint.X * 10.0, rb.MinPoint.Y * 10.0))
    target.Add(New Point2dData(rb.MaxPoint.X * 10.0, rb.MaxPoint.Y * 10.0))
    target.Add(New Point2dData(rb.MinPoint.X * 10.0, rb.MaxPoint.Y * 10.0))
End Sub

Private Function ComputeMinimumBoundingRectangle(points As List(Of Point2dData)) As MinAreaRect2d
    Dim hull As List(Of Point2dData) = BuildConvexHull(points)
    If hull.Count < 3 Then Return New MinAreaRect2d(0, 0)

    Dim bestArea As Double = Double.MaxValue
    Dim bestW As Double = 0.0
    Dim bestH As Double = 0.0

    For i As Integer = 0 To hull.Count - 1
        Dim p0 As Point2dData = hull(i)
        Dim p1 As Point2dData = hull((i + 1) Mod hull.Count)

        Dim ex As Double = p1.X - p0.X
        Dim ey As Double = p1.Y - p0.Y
        Dim eLen As Double = Math.Sqrt(ex * ex + ey * ey)
        If eLen < 0.000001 Then Continue For

        Dim ux As Double = ex / eLen
        Dim uy As Double = ey / eLen
        Dim vx As Double = -uy
        Dim vy As Double = ux

        Dim minU As Double = Double.MaxValue
        Dim maxU As Double = Double.MinValue
        Dim minV As Double = Double.MaxValue
        Dim maxV As Double = Double.MinValue

        For Each p As Point2dData In hull
            Dim pu As Double = p.X * ux + p.Y * uy
            Dim pv As Double = p.X * vx + p.Y * vy
            If pu < minU Then minU = pu
            If pu > maxU Then maxU = pu
            If pv < minV Then minV = pv
            If pv > maxV Then maxV = pv
        Next

        Dim w As Double = Math.Max(0.0, maxU - minU)
        Dim h As Double = Math.Max(0.0, maxV - minV)
        Dim area As Double = w * h
        If area < bestArea Then
            bestArea = area
            bestW = w
            bestH = h
        End If
    Next

    Return New MinAreaRect2d(bestW, bestH)
End Function

Private Function BuildConvexHull(points As List(Of Point2dData)) As List(Of Point2dData)
    Dim pts As New List(Of Point2dData)()
    For Each p As Point2dData In points
        pts.Add(p)
    Next

    pts.Sort(Function(a As Point2dData, b As Point2dData)
                 If a.X <> b.X Then Return a.X.CompareTo(b.X)
                 Return a.Y.CompareTo(b.Y)
             End Function)

    Dim unique As New List(Of Point2dData)()
    Dim lastX As Double = Double.NaN
    Dim lastY As Double = Double.NaN
    For Each p As Point2dData In pts
        If unique.Count = 0 OrElse Math.Abs(p.X - lastX) > 0.000001 OrElse Math.Abs(p.Y - lastY) > 0.000001 Then
            unique.Add(p)
            lastX = p.X
            lastY = p.Y
        End If
    Next

    If unique.Count <= 2 Then Return unique

    Dim lower As New List(Of Point2dData)()
    For Each p As Point2dData In unique
        While lower.Count >= 2 AndAlso Cross(lower(lower.Count - 2), lower(lower.Count - 1), p) <= 0
            lower.RemoveAt(lower.Count - 1)
        End While
        lower.Add(p)
    Next

    Dim upper As New List(Of Point2dData)()
    For i As Integer = unique.Count - 1 To 0 Step -1
        Dim p As Point2dData = unique(i)
        While upper.Count >= 2 AndAlso Cross(upper(upper.Count - 2), upper(upper.Count - 1), p) <= 0
            upper.RemoveAt(upper.Count - 1)
        End While
        upper.Add(p)
    Next

    lower.RemoveAt(lower.Count - 1)
    upper.RemoveAt(upper.Count - 1)
    lower.AddRange(upper)
    Return lower
End Function

Private Function Cross(a As Point2dData, b As Point2dData, c As Point2dData) As Double
    Return (b.X - a.X) * (c.Y - a.Y) - (b.Y - a.Y) * (c.X - a.X)
End Function

Private Function SafeViewName(v As DrawingView) As String
    Try
        Return v.Name
    Catch
        Return "<unnamed view>"
    End Try
End Function

Private Function SafeModelPath(d As Document) As String
    Try
        Return d.FullFileName
    Catch
        Return "<unknown model>"
    End Try
End Function

Private Function MmToCm(mm As Double) As Double
    Return mm / 10.0
End Function

Private Function FormatMm(mm As Double) As String
    Return Math.Round(mm, 1).ToString("0.#")
End Function

Private Class ModelOverallExtents
    Public LengthMm As Double
    Public WidthMm As Double
    Public HeightMm As Double
    Public DensityKgM3 As Double
    Public MassKg As Double
End Class

Private Structure Point2dData
    Public X As Double
    Public Y As Double
    Public Sub New(xValue As Double, yValue As Double)
        X = xValue
        Y = yValue
    End Sub
End Structure

Private Structure MinAreaRect2d
    Public WidthMm As Double
    Public HeightMm As Double
    Public Sub New(w As Double, h As Double)
        WidthMm = w
        HeightMm = h
    End Sub
End Structure
