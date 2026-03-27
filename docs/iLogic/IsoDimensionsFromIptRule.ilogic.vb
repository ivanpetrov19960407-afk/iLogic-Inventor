' IsoDimensionsFromIptRule.ilogic.vb
' Отдельное iLogic-правило:
' после расстановки видов на листах добавляет/обновляет подписи
' с габаритами модели из исходного .ipt/.iam рядом с изометрическими видами.

Option Explicit On

Imports Inventor
Imports System

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    Dim doc As DrawingDocument = TryCast(app.ActiveDocument, DrawingDocument)
    If doc Is Nothing Then
        System.Windows.Forms.MessageBox.Show("Откройте .idw документ.", "IsoDimensionsFromIptRule")
        Return
    End If

    Dim processed As Integer = 0
    Dim skipped As Integer = 0

    For Each sh As Sheet In doc.Sheets
        For Each v As DrawingView In sh.DrawingViews
            If Not IsIsometric(v) Then Continue For

            Dim openedHere As Boolean = False
            Dim modelDoc As Document = ResolveModelDocument(app, v, openedHere)
            If modelDoc Is Nothing Then
                skipped += 1
                Continue For
            End If

            Try
                Dim mm As ModelOverallExtents = GetModelOverallExtentsMm(modelDoc)
                Dim noteText As String = BuildIsoDimNoteText(v, modelDoc, mm)
                UpsertIsoDimNote(doc, sh, v, noteText)
                processed += 1
            Catch
                skipped += 1
            Finally
                If openedHere Then
                    Try
                        modelDoc.Close(True)
                    Catch
                    End Try
                End If
            End Try
        Next
    Next

    doc.Update2(True)

    Dim msg As String = "Изометрические виды обработаны: " & processed.ToString() & vbCrLf &
                        "Пропущено: " & skipped.ToString() & vbCrLf & vbCrLf &
                        "Примечание: правило добавляет габариты L×W×H из RangeBox модели (.ipt/.iam)."
    System.Windows.Forms.MessageBox.Show(msg, "IsoDimensionsFromIptRule")
End Sub

Private Function IsIsometric(v As DrawingView) As Boolean
    If v Is Nothing Then Return False
    Select Case v.ViewOrientation
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
                                      ByRef openedHere As Boolean) As Document
    openedHere = False
    If v Is Nothing Then Return Nothing

    Dim desc As DocumentDescriptor = Nothing
    Try
        desc = v.ReferencedDocumentDescriptor
    Catch
    End Try
    If desc Is Nothing Then Return Nothing

    Try
        If desc.ReferencedDocument IsNot Nothing Then Return desc.ReferencedDocument
    Catch
    End Try

    Dim fullPath As String = String.Empty
    Try
        fullPath = desc.FullDocumentName
    Catch
    End Try
    If String.IsNullOrWhiteSpace(fullPath) Then Return Nothing
    If Not System.IO.File.Exists(fullPath) Then Return Nothing

    Try
        Dim wasSilent As Boolean = app.SilentOperation
        Try
            app.SilentOperation = True
            Dim opened As Document = app.Documents.Open(fullPath, False)
            openedHere = (opened IsNot Nothing)
            Return opened
        Finally
            app.SilentOperation = wasSilent
        End Try
    Catch
        Return Nothing
    End Try
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
    Dim p As Point2d = tg.CreatePoint2d(v.Left + v.Width + MmToCm(6.0), v.Top - MmToCm(6.0))
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
                                     m As ModelOverallExtents) As String
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

    Return viewName & " (" & modelName & ")" & vbCrLf &
           "Габариты из .ipt/.iam, мм:" & vbCrLf &
           "L=" & FormatMm(m.LengthMm) & "  W=" & FormatMm(m.WidthMm) & "  H=" & FormatMm(m.HeightMm)
End Function

Private Function GetModelOverallExtentsMm(modelDoc As Document) As ModelOverallExtents
    Dim res As New ModelOverallExtents()

    If modelDoc Is Nothing Then Return res

    Dim rb As Box = Nothing
    If TypeOf modelDoc Is PartDocument Then
        rb = DirectCast(modelDoc, PartDocument).ComponentDefinition.RangeBox
    ElseIf TypeOf modelDoc Is AssemblyDocument Then
        rb = DirectCast(modelDoc, AssemblyDocument).ComponentDefinition.RangeBox
    Else
        Return res
    End If

    If rb Is Nothing Then Return res

    Dim dx As Double = Math.Abs(rb.MaxPoint.X - rb.MinPoint.X)
    Dim dy As Double = Math.Abs(rb.MaxPoint.Y - rb.MinPoint.Y)
    Dim dz As Double = Math.Abs(rb.MaxPoint.Z - rb.MinPoint.Z)

    Dim vals() As Double = New Double() {dx * 10.0, dy * 10.0, dz * 10.0}
    Array.Sort(vals)
    Array.Reverse(vals)

    res.LengthMm = vals(0)
    res.WidthMm = vals(1)
    res.HeightMm = vals(2)
    Return res
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
End Class
