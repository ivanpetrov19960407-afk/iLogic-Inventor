' ================================================================
' StoneAlbumRule.ilogic.vb  –  v3.14
' Архитектура точно повторяет рабочий VBA RKM_IdwAlbum.bas
' Источник: vba-inventor / RKM_IdwAlbum.bas, RKM_FrameBorder.bas,
'           RKM_TitleBlockPrompted.bas, RKM_Excel.bas
' v3.14: smart auxiliary view selection from Front/Top/Right,
'        disable view notes by default, remove shrink loop from PlaceViewInSlot
' v3.13: restore border re-edit (removed early return in EnsureBorder),
'        cap view scale at 2.0 (was 100), hide probe labels in MeasureView
' v3.12: absolute slot geometry from A3 SPDS frame, no dynamic safeRect
'        removed GetSheetSafeRect/InsetRect/ContentSlot calls from PlaceViewsSlotBased
'        ScaleToFit cap raised from 20 to 100
' v3.11: fix PlaceViewInSlot — remove shrink loop, views now fill slots
'        (ViewFitsSlot always returned False due to annotation bounding box)
' v3.10: don't re-edit existing border/titleblock defs (fixes text overflow),
'        smart view-slot matching (permutation-based best fit),
'        dim notes via DrawingNotes (real mm from view scale)
' v3.9: FIX — ps array 0-based (New String(6){}), ps(k) вместо ps(k+1)
'       AddTitleBlock — убран Nothing (второй аргумент пустой)
'       4-view layout: LARGE=Front, ISO=IsoTopRight(Shaded), SMALL=Top, WIDE=Right
' v3.8: ГЛАВНЫЙ ФИКС — AddCustomBorder → AddBorder (правильный метод API)
'       AddCustomBorder не существует в iLogic — ошибка во всех прошлых версиях
'       sheet.Border/TitleBlock ReadOnly — убраны fallback-присвоения
'       AddCustomBorder не существует в iLogic — это была главная ошибка с самого начала
'       Fallback: sheet.Border = BORDER_NAME (iLogic-way)
' v3.7: ФИКС — SilentOperation=True вокруг AddCustomBorder и AddTitleBlock
'       (точно как в VBA ApplyRkmBorderToSheet / ApplyRkmTitleBlockToSheetWithPrompts)
' v3.6: ФИКС borderDef/tbDef — получаем в каждом BuildOneSheet заново
'       из doc.BorderDefinitions.Item() чтобы не протухали после doc.Update2
'       Убрали doc.Update2(True) после AddTitleBlock (мешал RCW)
' v3.5: ФИКС ViewFitsSlot — через v.Left/v.Top (как VBA DoesViewFitRect)
'       вместо ненадёжного v.Center.X±Width/2
'       ФИКС PlaceViewInSlot — вид создаётся сразу в центре слота,
'       лишнее v.Center=... убрано
'       v3.4: глобальный SilentOperation убран.
'       v3.3: MAX_AUTO_SCALE=20, TITLEBLOCK_GAP=0.05.
'       v3.2: полный порт slot-based layout.
' ================================================================

Option Explicit On

Imports Inventor
Imports System
Imports System.Collections.Generic

Sub Main()
    Dim excelPath     As String = String.Empty
    Dim workspacePath As String = String.Empty
    Dim sheetTabName  As String = "ALBUM"

    Try
        excelPath = iProperties.Value("Custom", "AlbumExcel")
    Catch
    End Try
    Try
        workspacePath = iProperties.Value("Custom", "AlbumWorkspace")
    Catch
    End Try
    Try
        sheetTabName = iProperties.Value("Custom", "AlbumSheet")
    Catch
    End Try

    Dim newExcel As String = InputBox(
        "Путь к Excel-файлу альбома (.xlsx):" & vbCrLf &
        "(оставьте как есть или введите новый)",
        "Шаг 1 из 3 — Excel", excelPath)
    If newExcel IsNot Nothing Then
        If Not String.IsNullOrWhiteSpace(newExcel) Then excelPath = newExcel.Trim()
    End If
    If String.IsNullOrWhiteSpace(excelPath) Then
        System.Windows.Forms.MessageBox.Show("Путь к Excel не указан.", "Отмена")
        Return
    End If

    If String.IsNullOrWhiteSpace(workspacePath) Then
        Try
            Dim proj As Object = ThisApplication.DesignProjectManager.ActiveDesignProject
            If proj IsNot Nothing Then workspacePath = CStr(proj.WorkspacePath)
        Catch
        End Try
    End If
    If String.IsNullOrWhiteSpace(workspacePath) Then
        workspacePath = System.IO.Path.GetDirectoryName(excelPath)
    End If

    Dim newWs As String = InputBox(
        "Папка с 3D-моделями (.ipt):",
        "Шаг 2 из 3 — Папка моделей", workspacePath)
    If newWs IsNot Nothing Then
        If Not String.IsNullOrWhiteSpace(newWs) Then workspacePath = newWs.Trim()
    End If

    Dim newSheet As String = InputBox(
        "Имя листа в Excel:", "Шаг 3 из 3 — Лист Excel", sheetTabName)
    If newSheet IsNot Nothing Then
        If Not String.IsNullOrWhiteSpace(newSheet) Then sheetTabName = newSheet.Trim()
    End If
    If String.IsNullOrWhiteSpace(sheetTabName) Then
        System.Windows.Forms.MessageBox.Show("Имя листа Excel не указано.", "Отмена")
        Return
    End If

    Try
        iProperties.Value("Custom", "AlbumExcel") = excelPath
    Catch
    End Try
    Try
        iProperties.Value("Custom", "AlbumWorkspace") = workspacePath
    Catch
    End Try
    Try
        iProperties.Value("Custom", "AlbumSheet") = sheetTabName
    Catch
    End Try

    Dim doc As DrawingDocument = TryCast(ThisApplication.ActiveDocument, DrawingDocument)
    If doc Is Nothing Then
        System.Windows.Forms.MessageBox.Show("Откройте .idw документ.", "Ошибка")
        Return
    End If

    Dim rule As New AlbumBuilder(ThisApplication)
    rule.Build(doc, excelPath, workspacePath, sheetTabName)
End Sub

' ================================================================
'  BUILDER
' ================================================================
Public Class AlbumBuilder

    Private ReadOnly _app As Inventor.Application

    ' Геометрия А3 СПДС
    Private Const A3_W_MM       As Double = 420.0
    Private Const A3_H_MM       As Double = 297.0
    Private Const FRAME_L_MM    As Double = 20.0
    Private Const FRAME_O_MM    As Double = 5.0
    Private Const TB_W_MM       As Double = 185.0
    Private Const TB_H_MM       As Double = 55.0
    Private Const BORDER_NAME   As String = "RKM_SPDS_A3_BORDER_V12"
    Private Const TB_NAME       As String = "RKM_SPDS_A3_FORM3_V17"
    Private Const SHEET_PFX     As String = "ALB_"
    Private Const ALBUM_MODE_VISUAL As Boolean = True
    Private Const ADD_VIEW_NOTES As Boolean = False

    ' Параметры layout (точно из VBA RKM_IdwAlbum.bas)
    Private Const GAP_MM              As Double = 8.0
    Private Const LAYOUT_PAD_MM       As Double = 6.0
    Private Const SAFE_LEFT_RATIO     As Double = 0.05
    Private Const SAFE_RIGHT_RATIO    As Double = 0.03
    Private Const SAFE_TOP_RATIO      As Double = 0.04
    Private Const SAFE_BOTTOM_RATIO   As Double = 0.03
    Private Const TITLEBLOCK_GAP_RATIO As Double = 0.05
    Private Const TECH_TOP_BAND_RATIO  As Double = 0.31
    Private Const TECH_RIGHT_COL_RATIO As Double = 0.34
    Private Const TECH_SMALL_SLOT_RATIO As Double = 0.38
    Private Const ORTHO_SCALE_MARGIN   As Double = 0.95
    Private Const ISO_SCALE_MARGIN     As Double = 0.9
    Private Const MAX_AUTO_SCALE       As Double = 20.0
    Private Const PROBE_SCALE          As Double = 0.1
    Private Const SLOT_CONTENT_PAD_MM  As Double = 2.0
    Private Const CAPTION_CLEAR_TOP_MM As Double = 7.0

    Public Sub New(app As Inventor.Application)
        _app = app
    End Sub

    Public Sub Build(doc As DrawingDocument, excelPath As String, workspacePath As String, sheetTab As String)
        Dim items As List(Of AlbumItem) = XlsxReader.Load(excelPath, workspacePath, sheetTab)
        If items.Count = 0 Then
            System.Windows.Forms.MessageBox.Show(
                "Excel не содержит строк." & vbCrLf & "Файл: " & excelPath & vbCrLf & "Лист: " & sheetTab,
                "Пустой список")
            Return
        End If

        Dim okCount   As Integer = 0
        Dim failCount As Integer = 0

        ' SilentOperation НЕ включаем глобально — иначе виды не обновляются!
        ' Точечно включается только вокруг Documents.Open внутри BuildOneSheet.
        ' v3.6: определения создаём один раз, но в каждом BuildOneSheet
        '       получаем свежую ссылку через .Item() чтобы RCW не протухал.
        Try
            ' Инициализируем определения рамки и штампа один раз
            EnsureBorder(doc)
            EnsureTitleBlock(doc)
            Dim tmplSheet As Sheet = ResolveTemplateSheet(doc)
            PurgeAlbumSheets(doc, tmplSheet)

            For i As Integer = 0 To items.Count - 1
                Dim item As AlbumItem = items(i)
                Dim promptSheet As String = String.Empty
                Dim promptSheets As String = String.Empty
                item.Prompts.TryGetValue("SHEET", promptSheet)
                item.Prompts.TryGetValue("SHEETS", promptSheets)

                If String.IsNullOrWhiteSpace(promptSheet) Then item.Prompts("SHEET") = (i + 1).ToString()
                If String.IsNullOrWhiteSpace(promptSheets) Then item.Prompts("SHEETS") = items.Count.ToString()

                Dim ok As Boolean = BuildOneSheet(doc, item, i + 1)
                If ok Then
                    okCount += 1
                Else
                    failCount += 1
                    Debug.Print("WARN: лист не собран, строка Excel=" & (i + 1).ToString() & ", модель=" & item.ModelPath)
                End If
            Next

            If tmplSheet IsNot Nothing Then
                Try
                    tmplSheet.Activate()
                Catch
                End Try
            End If

        Catch ex As Exception
            Debug.Print("Build error: " & ex.Message)
        End Try

        Dim msg As String = "Альбом собран: " & okCount & " листов."
        If failCount > 0 Then msg &= vbCrLf & "Не собрано: " & failCount & " (модели не найдены или ошибка видов)."
        System.Windows.Forms.MessageBox.Show(msg, "Готово",
            System.Windows.Forms.MessageBoxButtons.OK,
            System.Windows.Forms.MessageBoxIcon.Information)
    End Sub

    ' ── Один лист ──
    ' v3.6: borderDef/tbDef получаем заново внутри каждого вызова
    '       (COM RCW протухает после doc.Update2 если ссылка хранится снаружи)
    Private Function BuildOneSheet(doc As DrawingDocument, item As AlbumItem, rowIndex As Integer) As Boolean
        If Not System.IO.File.Exists(item.ModelPath) Then
            Debug.Print("WARN: модель не найдена, строка Excel=" & rowIndex.ToString() & ", путь=" & item.ModelPath)
            Return False
        End If

        Dim sheet As Sheet = Nothing
        Dim modelDoc As Document = Nothing
        Dim openedHere As Boolean = False

        Try
            Dim sheetName As String = SHEET_PFX & System.IO.Path.GetFileNameWithoutExtension(item.ModelPath)
            sheet = doc.Sheets.Add(
                DrawingSheetSizeEnum.kA3DrawingSheetSize,
                PageOrientationTypeEnum.kLandscapePageOrientation)
            sheet.Name = sheetName
            sheet.Activate()

            ' Убираем старые виды
            For vi As Integer = sheet.DrawingViews.Count To 1 Step -1
                Try
                    sheet.DrawingViews.Item(vi).Delete()
                Catch
                End Try
            Next

            ' Рамка СПДС — v3.6: получаем свежую ссылку каждый раз
            Dim borderDef As BorderDefinition = Nothing
            Try
                borderDef = doc.BorderDefinitions.Item(BORDER_NAME)
            Catch ex As Exception
                Debug.Print("WARN BorderDef.Item: " & ex.Message)
            End Try
            ' v3.8: AddBorder (не AddCustomBorder!) + fallback sheet.Border = name
            Dim borderOk As Boolean = False
            Try
                _app.SilentOperation = True
                Try
                    If sheet.Border IsNot Nothing Then
                        sheet.Border.Delete()
                    End If
                Catch
                End Try

                If borderDef IsNot Nothing Then
                    Try
                        sheet.AddBorder(borderDef)
                        borderOk = True
                        Debug.Print("AddBorder OK на листе: " & sheet.Name)
                    Catch ex As Exception
                        Debug.Print("WARN AddBorder: " & ex.Message)
                    End Try
                End If
            Finally
                _app.SilentOperation = False
            End Try
            If Not borderOk Then
                Debug.Print("WARN: рамка НЕ применилась на листе: " & sheet.Name)
            End If

            ' Штамп Форма 3 — v3.6: тоже свежая ссылка
            Dim tbDef As TitleBlockDefinition = Nothing
            Try
                tbDef = doc.TitleBlockDefinitions.Item(TB_NAME)
            Catch ex As Exception
                Debug.Print("WARN TBDef.Item: " & ex.Message)
            End Try
            Dim ps() As String = New String(6) {}
            Dim order As String() = {"CODE","PROJECT_NAME","DRAWING_NAME","ORG_NAME","STAGE","SHEET","SHEETS"}
            For k As Integer = 0 To order.Length - 1
                Dim v As String = String.Empty
                item.Prompts.TryGetValue(order(k), v)
                ps(k) = If(String.IsNullOrEmpty(v), "", v)
            Next
            ' v3.8: AddTitleBlock + fallback sheet.TitleBlock = name
            Dim tbOk As Boolean = False
            Try
                _app.SilentOperation = True
                Try
                    If sheet.TitleBlock IsNot Nothing Then
                        sheet.TitleBlock.Delete()
                    End If
                Catch
                End Try

                If tbDef IsNot Nothing Then
                    Try
                        sheet.AddTitleBlock(tbDef, , ps)
                        tbOk = True
                        Debug.Print("AddTitleBlock OK на листе: " & sheet.Name)
                    Catch ex As Exception
                        Debug.Print("WARN AddTitleBlock: " & ex.Message)
                    End Try
                End If
            Finally
                _app.SilentOperation = False
            End Try
            If Not tbOk Then
                Debug.Print("WARN: штамп НЕ применился на листе: " & sheet.Name)
            End If

            ' v3.6: НЕ вызываем doc.Update2 здесь — это инвалидирует COM-ссылки
            ' TitleBlock.RangeBox читаем без принудительного обновления
            ' (достаточно того что sheet.Activate() уже обновил геометрию листа)

            ' Открываем модель
            For Each ed As Document In _app.Documents
                If String.Equals(ed.FullFileName, item.ModelPath, StringComparison.OrdinalIgnoreCase) Then
                    modelDoc = ed
                    Exit For
                End If
            Next
            If modelDoc Is Nothing Then
                ' SilentOperation=True только вокруг Open — без окна прогресса
                _app.SilentOperation = True
                Try
                    modelDoc = _app.Documents.Open(item.ModelPath, False)
                    openedHere = (modelDoc IsNot Nothing)
                Catch ex As Exception
                    Debug.Print("WARN Open, строка Excel=" & rowIndex.ToString() & ", модель=" & item.ModelPath & ": " & ex.Message)
                Finally
                    _app.SilentOperation = False
                End Try
            End If
            If modelDoc Is Nothing Then
                Debug.Print("WARN: не удалось открыть, строка Excel=" & rowIndex.ToString() & ", модель=" & item.ModelPath)
                Try
                    sheet.Delete()
                Catch
                End Try
                Return False
            End If

            ' Виды через slot-based layout
            Dim viewsOk As Boolean = PlaceViewsSlotBased(doc, sheet, modelDoc)
            If Not viewsOk Then
                Debug.Print("WARN: не удалось построить виды, строка Excel=" & rowIndex.ToString() & ", лист=" & sheetName)
            End If
            doc.Update2(True)

            If sheet.DrawingViews.Count < 3 Then
                Debug.Print("WARN: менее 3 видов на листе: " & sheetName)
                Try
                    sheet.Delete()
                Catch
                End Try
                Return False
            End If

            Return True

        Catch ex As Exception
            Debug.Print("ERROR: BuildOneSheet, строка Excel=" & rowIndex.ToString() & ", модель=" & item.ModelPath & ": " & ex.Message)
            If sheet IsNot Nothing Then
                Try
                    sheet.Delete()
                Catch
                End Try
            End If
            Return False
        Finally
            If modelDoc IsNot Nothing AndAlso openedHere Then
                Try
                    modelDoc.Close(True)
                Catch
                End Try
            End If
        End Try
    End Function

    ' ================================================================
    '  ADAPTIVE TECH LAYOUT
    ' ================================================================
    Private Function PlaceViewsSlotBased(doc As DrawingDocument, sheet As Sheet, modelDoc As Document) As Boolean
        Dim safe As SlotRect = GetSheetSafeRect(doc, sheet)
        Dim gap As Double = Cm(doc, GAP_MM)

        Dim mFront As ViewMeasure = MeasureView(sheet, modelDoc, ViewOrientationTypeEnum.kFrontViewOrientation, DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle)
        Dim mTop As ViewMeasure = MeasureView(sheet, modelDoc, ViewOrientationTypeEnum.kTopViewOrientation, DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle)
        Dim mRight As ViewMeasure = MeasureView(sheet, modelDoc, ViewOrientationTypeEnum.kRightViewOrientation, DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle)
        Dim mIso As ViewMeasure = MeasureView(sheet, modelDoc, ViewOrientationTypeEnum.kIsoTopRightViewOrientation, DrawingViewStyleEnum.kShadedDrawingViewStyle)

        If ALBUM_MODE_VISUAL AndAlso mIso Is Nothing Then
            Debug.Print("WARN: не удалось измерить изометрический shaded вид")
            Return False
        End If

        Dim auxMeasures As New List(Of ViewMeasure)()
        If mFront IsNot Nothing Then auxMeasures.Add(mFront)
        If mTop IsNot Nothing Then auxMeasures.Add(mTop)
        If mRight IsNot Nothing Then auxMeasures.Add(mRight)

        If auxMeasures.Count < 2 Then
            Debug.Print("WARN: не удалось измерить два вспомогательных вида")
            Return False
        End If

        Dim pairs As List(Of AuxPair) = BuildAuxPairs(auxMeasures)
        Dim patterns As List(Of LayoutPattern) = BuildLayoutPatterns(doc, safe, gap)
        Dim best As LayoutPlan = Nothing

        For Each ptn As LayoutPattern In patterns
            For Each pair As AuxPair In pairs
                Dim candidate As LayoutPlan = EvaluatePlanFlexible(ptn, mIso, pair.A, pair.B)
                If candidate Is Nothing Then Continue For

                If best Is Nothing OrElse candidate.Score > best.Score Then
                    best = candidate
                End If
            Next
        Next

        If best Is Nothing Then
            Debug.Print("WARN: не найден корректный layout визуализации")
            Return False
        End If

        Dim vMain As DrawingView = PlaceViewInSlot(sheet, modelDoc, best.MainMeasure, best.MainFit, best.MainSlot)
        If vMain Is Nothing Then Return False

        Dim vAux1 As DrawingView = PlaceViewInSlot(sheet, modelDoc, best.Aux1Measure, best.Aux1Fit, best.Aux1Slot)
        Dim vAux2 As DrawingView = PlaceViewInSlot(sheet, modelDoc, best.Aux2Measure, best.Aux2Fit, best.Aux2Slot)

        If vAux1 Is Nothing OrElse vAux2 Is Nothing Then
            Debug.Print("WARN: не удалось разместить все вспомогательные виды")
            Return False
        End If

        If ADD_VIEW_NOTES Then
            AddDimNotes(doc, sheet, vMain, best.MainSlot, True)
            AddDimNotes(doc, sheet, vAux1, best.Aux1Slot, False)
            AddDimNotes(doc, sheet, vAux2, best.Aux2Slot, False)
        End If

        Return True
    End Function

    Private Function BuildLayoutPatterns(doc As DrawingDocument, safe As SlotRect, gap As Double) As List(Of LayoutPattern)
        Dim result As New List(Of LayoutPattern)()
        Dim w As Double = RectW(safe)
        Dim h As Double = RectH(safe)
        If w <= gap * 4 OrElse h <= gap * 4 Then Return result

        Dim auxBand As Double = Math.Max(h * 0.26, Cm(doc, 34.0))
        auxBand = Math.Min(auxBand, h * 0.34)
        Dim auxW As Double = Math.Max((w - gap * 3.0) / 2.0, Cm(doc, 45.0))

        ' Вариант B (низкий риск): большой изометрический вид + 2 маленьких ортогональных сверху.
        Dim pA As New LayoutPattern()
        pA.MainSlot = New SlotRect(safe.L, safe.R, safe.B, safe.T - auxBand - gap)
        pA.Aux1Slot = New SlotRect(safe.L, safe.L + auxW, safe.T - auxBand, safe.T)
        pA.Aux2Slot = New SlotRect(safe.R - auxW, safe.R, safe.T - auxBand, safe.T)
        pA.IsoSlot = New SlotRect(0, 0, 0, 0)
        result.Add(pA)

        ' Запасной вариант: справа вертикальный столбец с двумя малыми видами.
        Dim sideCol As Double = Math.Min(Math.Max(w * 0.26, Cm(doc, 45.0)), w * 0.34)
        Dim pB As New LayoutPattern()
        pB.MainSlot = New SlotRect(safe.L, safe.R - sideCol - gap, safe.B, safe.T)
        pB.Aux1Slot = New SlotRect(safe.R - sideCol, safe.R, safe.B + (h + gap) / 2.0, safe.T)
        pB.Aux2Slot = New SlotRect(safe.R - sideCol, safe.R, safe.B, safe.B + (h - gap) / 2.0)
        pB.IsoSlot = New SlotRect(0, 0, 0, 0)
        result.Add(pB)

        For Each p As LayoutPattern In result
            p.MainSlot = InsetRect(p.MainSlot, gap * 0.25)
            p.Aux1Slot = InsetRect(p.Aux1Slot, gap * 0.25)
            p.Aux2Slot = InsetRect(p.Aux2Slot, gap * 0.25)
        Next
        Return result
    End Function

    Private Function BuildAuxPairs(measures As List(Of ViewMeasure)) As List(Of AuxPair)
        Dim result As New List(Of AuxPair)()

        For i As Integer = 0 To measures.Count - 2
            For j As Integer = i + 1 To measures.Count - 1
                result.Add(New AuxPair(measures(i), measures(j)))
            Next
        Next

        Return result
    End Function

    Private Function EvaluatePlanFlexible(ptn As LayoutPattern,
                                          mainM As ViewMeasure,
                                          auxA As ViewMeasure,
                                          auxB As ViewMeasure) As LayoutPlan
        Dim best As LayoutPlan = Nothing

        For pass As Integer = 0 To 1
            Dim a1 As ViewMeasure = If(pass = 0, auxA, auxB)
            Dim a2 As ViewMeasure = If(pass = 0, auxB, auxA)

            Dim plan As New LayoutPlan()
            plan.MainSlot = ptn.MainSlot
            plan.Aux1Slot = ptn.Aux1Slot
            plan.Aux2Slot = ptn.Aux2Slot
            plan.IsoSlot = ptn.IsoSlot

            plan.MainMeasure = mainM
            plan.Aux1Measure = a1
            plan.Aux2Measure = a2

            plan.MainFit = ScaleToFit(plan.MainSlot, mainM, ISO_SCALE_MARGIN)
            plan.Aux1Fit = ScaleToFit(plan.Aux1Slot, a1, ORTHO_SCALE_MARGIN)
            plan.Aux2Fit = ScaleToFit(plan.Aux2Slot, a2, ORTHO_SCALE_MARGIN)

            If plan.MainFit Is Nothing OrElse plan.MainFit.Scale <= 0 Then Continue For
            If plan.Aux1Fit Is Nothing OrElse plan.Aux1Fit.Scale <= 0 Then Continue For
            If plan.Aux2Fit Is Nothing OrElse plan.Aux2Fit.Scale <= 0 Then Continue For

            Dim mainArea As Double = plan.MainFit.ProjectedW * plan.MainFit.ProjectedH
            Dim aux1Area As Double = plan.Aux1Fit.ProjectedW * plan.Aux1Fit.ProjectedH
            Dim aux2Area As Double = plan.Aux2Fit.ProjectedW * plan.Aux2Fit.ProjectedH
            Dim auxArea As Double = aux1Area + aux2Area

            Dim workArea As Double = RectW(plan.MainSlot) * RectH(plan.MainSlot) +
                                     RectW(plan.Aux1Slot) * RectH(plan.Aux1Slot) +
                                     RectW(plan.Aux2Slot) * RectH(plan.Aux2Slot)
            If workArea <= 0 Then Continue For

            Dim fill As Double = (mainArea + auxArea) / workArea
            Dim mainDominance As Double = mainArea / Math.Max(0.0001, mainArea + auxArea)
            Dim auxBalance As Double = Math.Min(aux1Area, aux2Area) / Math.Max(0.0001, Math.Max(aux1Area, aux2Area))

            plan.Score = fill * 0.35 + mainDominance * 0.45 + auxBalance * 0.2
            If mainDominance < 0.45 Then plan.Score -= 0.3

            If best Is Nothing OrElse plan.Score > best.Score Then
                best = plan
            End If
        Next

        Return best
    End Function

    Private Function EvaluatePlan(ptn As LayoutPattern,
                                  mainM As ViewMeasure,
                                  aux1M As ViewMeasure,
                                  aux2M As ViewMeasure,
                                  isoM As ViewMeasure) As LayoutPlan
        Dim plan As New LayoutPlan()
        plan.MainSlot = ptn.MainSlot
        plan.Aux1Slot = ptn.Aux1Slot
        plan.Aux2Slot = ptn.Aux2Slot
        plan.IsoSlot = ptn.IsoSlot

        plan.MainFit = ScaleToFit(plan.MainSlot, mainM, ISO_SCALE_MARGIN)
        plan.Aux1Fit = ScaleToFit(plan.Aux1Slot, aux1M, ORTHO_SCALE_MARGIN)
        plan.Aux2Fit = ScaleToFit(plan.Aux2Slot, aux2M, ORTHO_SCALE_MARGIN)
        If isoM IsNot Nothing Then plan.IsoFit = ScaleToFit(plan.IsoSlot, isoM, ORTHO_SCALE_MARGIN)

        If plan.MainFit Is Nothing OrElse plan.MainFit.Scale <= 0 Then Return Nothing
        If plan.Aux1Fit Is Nothing OrElse plan.Aux1Fit.Scale <= 0 Then Return Nothing
        If plan.Aux2Fit Is Nothing OrElse plan.Aux2Fit.Scale <= 0 Then Return Nothing

        Dim mainArea As Double = plan.MainFit.ProjectedW * plan.MainFit.ProjectedH
        Dim auxArea As Double = plan.Aux1Fit.ProjectedW * plan.Aux1Fit.ProjectedH +
                                plan.Aux2Fit.ProjectedW * plan.Aux2Fit.ProjectedH
        Dim isoArea As Double = 0
        If plan.IsoFit IsNot Nothing Then
            isoArea = plan.IsoFit.ProjectedW * plan.IsoFit.ProjectedH
        End If

        Dim workArea As Double = RectW(plan.MainSlot) * RectH(plan.MainSlot) +
                                 RectW(plan.Aux1Slot) * RectH(plan.Aux1Slot) +
                                 RectW(plan.Aux2Slot) * RectH(plan.Aux2Slot)
        If workArea <= 0 Then Return Nothing

        Dim fill As Double = (mainArea + auxArea + isoArea * 0.2) / workArea
        Dim mainDominance As Double = mainArea / Math.Max(0.0001, (mainArea + auxArea))
        Dim compactness As Double = Math.Min(fill, 0.9)

        plan.Score = compactness * 0.4 + mainDominance * 0.6
        If mainDominance < 0.55 Then plan.Score -= 0.4
        Return plan
    End Function

    Private Sub AddDimNotes(doc As DrawingDocument, sheet As Sheet,
                            v As DrawingView, slot As SlotRect,
                            isMain As Boolean)
        If v Is Nothing Then Return

        Dim added As Integer = 0
        If isMain Then
            Try
                Dim px As Double = Math.Max(slot.L + Cm(doc, 3.0), Math.Min(slot.R - Cm(doc, 3.0), v.Left + v.Width / 2.0))
                Dim py As Double = Math.Max(slot.B + Cm(doc, 2.0), Math.Min(slot.T - Cm(doc, 2.0), v.Top - v.Height / 2.0 - Cm(doc, 2.0)))
                sheet.DrawingNotes.Add(_app.TransientGeometry.CreatePoint2d(px, py), "Главный вид (изометрия, shaded)")
            Catch ex As Exception
                Debug.Print("WARN AddDimNotes main caption: " & ex.Message)
            End Try
            Return
        End If

        added += TryAddTrueDimensions(doc, sheet, v, slot, True, False)
        If added < 1 Then AddFallbackDimensionNotes(doc, sheet, v, slot, True, False)
    End Sub

    Private Function TryAddTrueDimensions(doc As DrawingDocument, sheet As Sheet,
                                          v As DrawingView, slot As SlotRect,
                                          addHorizontal As Boolean,
                                          addVertical As Boolean) As Integer
        Dim count As Integer = 0
        Try
            Dim curves As DrawingCurvesEnumerator = v.DrawingCurves
            If curves Is Nothing OrElse curves.Count < 2 Then Return 0

            Dim minXCurve As DrawingCurve = Nothing
            Dim maxXCurve As DrawingCurve = Nothing
            Dim minYCurve As DrawingCurve = Nothing
            Dim maxYCurve As DrawingCurve = Nothing
            Dim minX As Double = Double.MaxValue
            Dim maxX As Double = Double.MinValue
            Dim minY As Double = Double.MaxValue
            Dim maxY As Double = Double.MinValue

            For i As Integer = 1 To curves.Count
                Dim c As DrawingCurve = curves.Item(i)
                Dim rb As Box2d = c.RangeBox
                Dim cx As Double = (rb.MinPoint.X + rb.MaxPoint.X) / 2.0
                Dim cy As Double = (rb.MinPoint.Y + rb.MaxPoint.Y) / 2.0
                If cx < minX Then minX = cx : minXCurve = c
                If cx > maxX Then maxX = cx : maxXCurve = c
                If cy < minY Then minY = cy : minYCurve = c
                If cy > maxY Then maxY = cy : maxYCurve = c
            Next

            If addHorizontal AndAlso minXCurve IsNot Nothing AndAlso maxXCurve IsNot Nothing Then
                Try
                    Dim i1 As GeometryIntent = sheet.CreateGeometryIntent(minXCurve, PointIntentEnum.kMidPointIntent)
                    Dim i2 As GeometryIntent = sheet.CreateGeometryIntent(maxXCurve, PointIntentEnum.kMidPointIntent)
                    Dim p As Point2d = _app.TransientGeometry.CreatePoint2d((slot.L + slot.R) / 2.0, Math.Min(slot.T - Cm(doc, 2.0), v.Top + Cm(doc, 4.0)))
                    sheet.DrawingDimensions.GeneralDimensions.AddLinear(p, i1, i2, DimensionTypeEnum.kHorizontalDimensionType)
                    count += 1
                Catch
                End Try
            End If

            If addVertical AndAlso minYCurve IsNot Nothing AndAlso maxYCurve IsNot Nothing Then
                Try
                    Dim j1 As GeometryIntent = sheet.CreateGeometryIntent(minYCurve, PointIntentEnum.kMidPointIntent)
                    Dim j2 As GeometryIntent = sheet.CreateGeometryIntent(maxYCurve, PointIntentEnum.kMidPointIntent)
                    Dim p2 As Point2d = _app.TransientGeometry.CreatePoint2d(Math.Min(slot.R - Cm(doc, 1.0), v.Left + v.Width + Cm(doc, 3.0)), (slot.B + slot.T) / 2.0)
                    sheet.DrawingDimensions.GeneralDimensions.AddLinear(p2, j1, j2, DimensionTypeEnum.kVerticalDimensionType)
                    count += 1
                Catch
                End Try
            End If
        Catch ex As Exception
            Debug.Print("WARN TryAddTrueDimensions: " & ex.Message)
        End Try
        Return count
    End Function

    Private Sub AddFallbackDimensionNotes(doc As DrawingDocument, sheet As Sheet,
                                          v As DrawingView, slot As SlotRect,
                                          addHorizontal As Boolean,
                                          addVertical As Boolean)
        Try
            Dim sc As Double = v.Scale
            If sc <= 0.0001 Then Return

            Dim realWmm As Double = Math.Round(v.Width / sc * 10.0)
            Dim realHmm As Double = Math.Round(v.Height / sc * 10.0)

            If addHorizontal AndAlso realWmm > 1 Then
                Dim px As Double = Math.Max(slot.L + Cm(doc, 3.0), Math.Min(slot.R - Cm(doc, 3.0), v.Left + v.Width / 2.0))
                Dim py As Double = Math.Min(slot.T - Cm(doc, 1.0), v.Top + Cm(doc, 5.0))
                sheet.DrawingNotes.Add(_app.TransientGeometry.CreatePoint2d(px, py),
                                       "↔ " & String.Format("{0:F0} мм", realWmm))
            End If

            If addVertical AndAlso realHmm > 1 Then
                Dim rightSpace As Double = slot.R - (v.Left + v.Width)
                Dim leftSpace As Double = v.Left - slot.L
                Dim px2 As Double
                If rightSpace >= leftSpace Then
                    px2 = Math.Min(slot.R - Cm(doc, 1.0), v.Left + v.Width + Cm(doc, 2.5))
                Else
                    px2 = Math.Max(slot.L + Cm(doc, 1.0), v.Left - Cm(doc, 2.5))
                End If
                Dim py2 As Double = Math.Max(slot.B + Cm(doc, 2.0), Math.Min(slot.T - Cm(doc, 2.0), v.Top - v.Height / 2.0))
                sheet.DrawingNotes.Add(_app.TransientGeometry.CreatePoint2d(px2, py2),
                                       "↕ " & String.Format("{0:F0} мм", realHmm))
            End If
        Catch ex As Exception
            Debug.Print("WARN AddFallbackDimensionNotes: " & ex.Message)
        End Try
    End Sub

    Private Function MeasureView(sheet As Sheet, modelDoc As Document,
                                 orient As ViewOrientationTypeEnum,
                                 style As DrawingViewStyleEnum) As ViewMeasure
        Dim probe As DrawingView = Nothing
        Try
            probe = sheet.DrawingViews.AddBaseView(
                modelDoc,
                _app.TransientGeometry.CreatePoint2d(sheet.Width / 2, sheet.Height / 2),
                PROBE_SCALE, orient, style)
            If probe Is Nothing Then Return Nothing
            Try
                probe.ShowLabel = False
            Catch
            End Try

            Dim dd As DrawingDocument = TryCast(sheet.Parent, DrawingDocument)
            If dd IsNot Nothing Then dd.Update2(True)

            Dim m As New ViewMeasure()
            m.UnitW = probe.Width / PROBE_SCALE
            m.UnitH = probe.Height / PROBE_SCALE
            m.Orient = orient
            m.Style = style
            If m.UnitW < 0.0001 AndAlso m.UnitH < 0.0001 Then Return Nothing
            Return m
        Catch ex As Exception
            Debug.Print("WARN: MeasureView failed: " & ex.Message)
            Return Nothing
        Finally
            If probe IsNot Nothing Then
                Try
                    probe.Delete()
                Catch
                End Try
            End If
        End Try
    End Function

    Private Function ScaleToFit(slot As SlotRect, m As ViewMeasure, margin As Double) As FitResult
        If m Is Nothing Then Return Nothing
        Dim sw As Double = RectW(slot)
        Dim sh As Double = RectH(slot)
        If sw <= 0 OrElse sh <= 0 Then Return Nothing

        Dim res0 As FitResult = BuildFitResult(sw, sh, m.UnitW, m.UnitH, margin, False)
        Dim res90 As FitResult = BuildFitResult(sw, sh, m.UnitW, m.UnitH, margin, True)

        Dim best As FitResult = res0
        If best Is Nothing OrElse (res90 IsNot Nothing AndAlso res90.Scale > best.Scale) Then
            best = res90
        End If
        Return best
    End Function

    Private Function BuildFitResult(slotW As Double, slotH As Double,
                                    unitW As Double, unitH As Double,
                                    margin As Double, rotate90 As Boolean) As FitResult
        Dim w As Double = unitW
        Dim h As Double = unitH
        If rotate90 Then
            w = unitH
            h = unitW
        End If
        If w <= 0.0001 OrElse h <= 0.0001 Then Return Nothing

        Dim sc As Double = Math.Min(slotW / w, slotH / h) * margin
        sc = Math.Min(sc, 2.0)
        If sc < 0.02 Then Return Nothing

        Dim r As New FitResult()
        r.Scale = sc
        r.Rotate90 = rotate90
        r.ProjectedW = w * sc
        r.ProjectedH = h * sc
        Return r
    End Function

    Private Function PlaceViewInSlot(sheet As Sheet, modelDoc As Document,
                                 m As ViewMeasure, fit As FitResult,
                                 slot As SlotRect) As DrawingView
        If m Is Nothing OrElse fit Is Nothing Then Return Nothing
        Dim cx As Double = (slot.L + slot.R) / 2.0
        Dim cy As Double = (slot.B + slot.T) / 2.0
        Dim v As DrawingView = Nothing

        Try
            v = sheet.DrawingViews.AddBaseView(
                modelDoc,
                _app.TransientGeometry.CreatePoint2d(cx, cy),
                fit.Scale, m.Orient, m.Style)
            If v Is Nothing Then Return Nothing

            Try
                v.ShowLabel = False
            Catch
            End Try

            If fit.Rotate90 Then
                Try
                    v.Rotation = v.Rotation + Math.PI / 2.0
                Catch exRot As Exception
                    Debug.Print("WARN Rotate90: " & exRot.Message)
                End Try
            End If

            Dim dd As DrawingDocument = TryCast(sheet.Parent, DrawingDocument)
            If dd IsNot Nothing Then dd.Update2(True)

            If Not ViewFitsSlot(v, slot) Then
                Debug.Print("WARN: view does not fully fit slot after placement")
            End If

            Return v

        Catch ex As Exception
            Debug.Print("WARN PlaceViewInSlot: " & ex.Message)
            If v IsNot Nothing Then
                Try
                    v.Delete()
                Catch
                End Try
            End If
            Return Nothing
        End Try
    End Function

    Private Function ViewFitsSlot(v As DrawingView, slot As SlotRect) As Boolean
        If v Is Nothing Then Return False
        Dim vL As Double = v.Left
        Dim vT As Double = v.Top
        Dim vR As Double = vL + v.Width
        Dim vB As Double = vT - v.Height
        Dim eps As Double = 0.001
        Return (vL >= slot.L - eps AndAlso vR <= slot.R + eps AndAlso
                vB >= slot.B - eps AndAlso vT <= slot.T + eps)
    End Function

    Private Function GetSheetSafeRect(doc As DrawingDocument, sheet As Sheet) As SlotRect
        Dim frameL As Double = Cm(doc, FRAME_L_MM)
        Dim frameR As Double = Cm(doc, A3_W_MM - FRAME_O_MM)
        Dim frameB As Double = Cm(doc, FRAME_O_MM)
        Dim frameT As Double = Cm(doc, A3_H_MM - FRAME_O_MM)

        Dim l As Double = frameL + Cm(doc, 4.0)
        Dim r As Double = frameR - Cm(doc, 3.0)
        Dim b As Double = frameB + Cm(doc, 3.0)
        Dim t As Double = frameT - Cm(doc, 3.0)

        Try
            If sheet.TitleBlock IsNot Nothing Then
                Dim rb As Box2d = sheet.TitleBlock.RangeBox
                If rb IsNot Nothing Then
                    Dim tbTop As Double = Math.Max(rb.MinPoint.Y, rb.MaxPoint.Y)
                    b = Math.Max(b, tbTop + Cm(doc, 4.0))
                End If
            End If
        Catch ex As Exception
            Debug.Print("SafeRect TB err: " & ex.Message)
        End Try

        If t <= b + Cm(doc, 25.0) Then
            b = Cm(doc, 65.0)
            t = frameT - Cm(doc, 2.0)
        End If
        Return New SlotRect(l, r, b, t)
    End Function

    Private Function InsetRect(r As SlotRect, d As Double) As SlotRect
        Return New SlotRect(r.L + d, r.R - d, r.B + d, r.T - d)
    End Function
    Private Function RectW(r As SlotRect) As Double
        Return r.R - r.L
    End Function
    Private Function RectH(r As SlotRect) As Double
        Return r.T - r.B
    End Function
    ' ================================================================
    '  РАМКА СПДС А3
    ' ================================================================
    Public Function EnsureBorder(doc As DrawingDocument) As BorderDefinition
        Dim def As BorderDefinition = Nothing
        Try
            def = doc.BorderDefinitions.Item(BORDER_NAME)
        Catch
        End Try

        If def Is Nothing Then
            Try
                def = doc.BorderDefinitions.Add(BORDER_NAME)
            Catch ex As Exception
                Debug.Print("WARN BorderDef.Add: " & ex.Message)
                Try
                    def = doc.BorderDefinitions.Item(BORDER_NAME)
                Catch ex2 As Exception
                    Debug.Print("WARN BorderDef.Item fallback: " & ex2.Message)
                    Return Nothing
                End Try
            End Try
        End If

        Dim sk As DrawingSketch = Nothing
        def.Edit(sk)
        Try
            For i As Integer = sk.SketchLines.Count To 1 Step -1
                Try
                    sk.SketchLines.Item(i).Delete()
                Catch
                End Try
            Next
            sk.SketchLines.AddByTwoPoints(P(0,0), P(0.0001, 0.0001))
            sk.SketchLines.AddByTwoPoints(
                P(Cm(doc, A3_W_MM), Cm(doc, A3_H_MM)),
                P(Cm(doc, A3_W_MM) - 0.0001, Cm(doc, A3_H_MM) - 0.0001))
            sk.SketchLines.AddAsTwoPointRectangle(
                P(Cm(doc, FRAME_L_MM), Cm(doc, FRAME_O_MM)),
                P(Cm(doc, A3_W_MM - FRAME_O_MM), Cm(doc, A3_H_MM - FRAME_O_MM)))
        Finally
            def.ExitEdit(True)
        End Try
        Return def
    End Function

    ' ================================================================
    '  ШТАМП ФОРМА 3
    ' ================================================================
    Public Function EnsureTitleBlock(doc As DrawingDocument) As TitleBlockDefinition
        Dim def As TitleBlockDefinition = Nothing
        Try
            def = doc.TitleBlockDefinitions.Item(TB_NAME)
        Catch
        End Try
        ' Если определение уже есть, не перерисовываем:
        ' это сохраняет пользовательские правки и гарантирует стабильную привязку Prompt-ов.
        If def IsNot Nothing Then Return def

        Try
            def = doc.TitleBlockDefinitions.Add(TB_NAME)
        Catch ex As Exception
            Debug.Print("WARN TBDef.Add: " & ex.Message)
            Return Nothing
        End Try

        Dim sk As DrawingSketch = Nothing
        def.Edit(sk)
        Try
            For i As Integer = sk.TextBoxes.Count To 1 Step -1
                Try
                    sk.TextBoxes.Item(i).Delete()
                Catch
                End Try
            Next
            For i As Integer = sk.SketchLines.Count To 1 Step -1
                Try
                    sk.SketchLines.Item(i).Delete()
                Catch
                End Try
            Next
            DrawTbGeometry(doc, sk)
            DrawTbLabels(doc, sk)
        Finally
            def.ExitEdit(True)
        End Try
        Return def
    End Function

    Private Sub DrawTbGeometry(doc As DrawingDocument, sk As DrawingSketch)
        Dim x2 As Double = -Cm(doc, FRAME_O_MM)
        Dim y1 As Double =  Cm(doc, FRAME_O_MM)
        Dim x1 As Double = x2 - Cm(doc, TB_W_MM)
        Dim y2 As Double = y1 + Cm(doc, TB_H_MM)

        sk.SketchLines.AddByTwoPoints(P(0,0), P(-0.0001, 0.0001))
        sk.SketchLines.AddAsTwoPointRectangle(P(x1,y1), P(x2,y2))

        VL(doc, sk, x1, y1,   7,  0, 55) : VL(doc, sk, x1, y1,  17,  0, 55)
        VL(doc, sk, x1, y1,  27,  0, 55) : VL(doc, sk, x1, y1,  42,  0, 55)
        VL(doc, sk, x1, y1,  57,  0, 55) : VL(doc, sk, x1, y1,  67,  0, 55)
        VL(doc, sk, x1, y1, 137,  0, 40) : VL(doc, sk, x1, y1, 152, 15, 40)
        VL(doc, sk, x1, y1, 167, 15, 40)

        Dim y As Double
        For y = 5.0 To 30.0 Step 5.0
            HL(doc, sk, x1, y1, 0, 67, y)
        Next
        HL(doc, sk, x1, y1,   0, 185, 15) : HL(doc, sk, x1, y1,   0,  67, 35)
        HL(doc, sk, x1, y1, 137, 185, 35) : HL(doc, sk, x1, y1,   0, 185, 40)
        HL(doc, sk, x1, y1,   0,  67, 45) : HL(doc, sk, x1, y1,   0,  67, 50)
    End Sub

    Private Sub DrawTbLabels(doc As DrawingDocument, sk As DrawingSketch)
        Dim x2 As Double = -Cm(doc, FRAME_O_MM)
        Dim y1 As Double =  Cm(doc, FRAME_O_MM)
        Dim x1 As Double = x2 - Cm(doc, TB_W_MM)

        Lbl(doc, sk, x1, y1,   0, 35,   7, 40, "Изм.")
        Lbl(doc, sk, x1, y1,   7, 35,  17, 40, "Кол.уч")
        Lbl(doc, sk, x1, y1,  17, 35,  27, 40, "Лист")
        Lbl(doc, sk, x1, y1,  27, 35,  42, 40, "№ doc.")
        Lbl(doc, sk, x1, y1,  42, 35,  57, 40, "Подп.")
        Lbl(doc, sk, x1, y1,  57, 35,  67, 40, "Дата")
        Lbl(doc, sk, x1, y1, 137, 35, 152, 40, "Стадия")
        Lbl(doc, sk, x1, y1, 152, 35, 167, 40, "Лист")
        Lbl(doc, sk, x1, y1, 167, 35, 185, 40, "Листов")

        Prm(doc, sk, x1, y1,  67, 40, 185, 55, "CODE")
        Prm(doc, sk, x1, y1,  67, 15, 137, 40, "PROJECT_NAME")
        Prm(doc, sk, x1, y1,  67,  0, 137, 15, "DRAWING_NAME")
        Prm(doc, sk, x1, y1, 137,  0, 185, 15, "ORG_NAME")
        Prm(doc, sk, x1, y1, 137, 15, 152, 35, "STAGE")
        Prm(doc, sk, x1, y1, 152, 15, 167, 35, "SHEET")
        Prm(doc, sk, x1, y1, 167, 15, 185, 35, "SHEETS")
    End Sub

    ' ================================================================
    '  ВСПОМОГАТЕЛЬНЫЕ
    ' ================================================================
    Private Function ResolveTemplateSheet(doc As DrawingDocument) As Sheet
        For Each s As Sheet In doc.Sheets
            If Not s.Name.StartsWith(SHEET_PFX, StringComparison.OrdinalIgnoreCase) Then Return s
        Next
        Return If(doc.Sheets.Count > 0, doc.Sheets.Item(1), Nothing)
    End Function

    Private Sub PurgeAlbumSheets(doc As DrawingDocument, tmpl As Sheet)
        Dim toDelete As New List(Of Sheet)()
        For Each s As Sheet In doc.Sheets
            If s.Name.StartsWith(SHEET_PFX, StringComparison.OrdinalIgnoreCase) Then
                If tmpl Is Nothing OrElse Not Object.ReferenceEquals(s, tmpl) Then
                    toDelete.Add(s)
                End If
            End If
        Next
        For Each s As Sheet In toDelete
            Try
                s.Delete()
            Catch
            End Try
        Next
    End Sub

    Private Sub VL(doc As DrawingDocument, sk As DrawingSketch, x0 As Double, y0 As Double, atMm As Double, yFr As Double, yTo As Double)
        sk.SketchLines.AddByTwoPoints(P(x0+Cm(doc,atMm), y0+Cm(doc,yFr)), P(x0+Cm(doc,atMm), y0+Cm(doc,yTo)))
    End Sub
    Private Sub HL(doc As DrawingDocument, sk As DrawingSketch, x0 As Double, y0 As Double, xFr As Double, xTo As Double, atMm As Double)
        sk.SketchLines.AddByTwoPoints(P(x0+Cm(doc,xFr), y0+Cm(doc,atMm)), P(x0+Cm(doc,xTo), y0+Cm(doc,atMm)))
    End Sub
    Private Sub Lbl(doc As DrawingDocument, sk As DrawingSketch, x0 As Double, y0 As Double, l As Double, b As Double, r As Double, t As Double, txt As String)
        Dim tb As Inventor.TextBox = sk.TextBoxes.AddByRectangle(P(x0+Cm(doc,l), y0+Cm(doc,b)), P(x0+Cm(doc,r), y0+Cm(doc,t)), txt)
        tb.HorizontalJustification = HorizontalTextAlignmentEnum.kAlignTextCenter
        tb.VerticalJustification   = VerticalTextAlignmentEnum.kAlignTextMiddle
    End Sub
    Private Sub Prm(doc As DrawingDocument, sk As DrawingSketch, x0 As Double, y0 As Double, l As Double, b As Double, r As Double, t As Double, nm As String)
        Dim tb As Inventor.TextBox = sk.TextBoxes.AddByRectangle(P(x0+Cm(doc,l), y0+Cm(doc,b)), P(x0+Cm(doc,r), y0+Cm(doc,t)), "<Prompt>" & nm & "</Prompt>")
        tb.HorizontalJustification = HorizontalTextAlignmentEnum.kAlignTextCenter
        tb.VerticalJustification   = VerticalTextAlignmentEnum.kAlignTextMiddle
    End Sub
    Private Function Cm(doc As DrawingDocument, mm As Double) As Double
        If doc Is Nothing Then Return mm * 0.1
        Return doc.UnitsOfMeasure.ConvertUnits(mm, UnitsTypeEnum.kMillimeterLengthUnits, UnitsTypeEnum.kCentimeterLengthUnits)
    End Function
    Private Function P(x As Double, y As Double) As Point2d
        Return _app.TransientGeometry.CreatePoint2d(x, y)
    End Function

    ' ================================================================
    '  ВСПОМОГАТЕЛЬНЫЕ КЛАССЫ
    ' ================================================================
    Public Class SlotRect
        Public L As Double
        Public R As Double
        Public B As Double
        Public T As Double
        Public Sub New(l As Double, r As Double, b As Double, t As Double)
            Me.L = l : Me.R = r : Me.B = b : Me.T = t
        End Sub
    End Class

    Public Class ViewMeasure
        Public UnitW  As Double
        Public UnitH  As Double
        Public Orient As ViewOrientationTypeEnum
        Public Style  As DrawingViewStyleEnum
    End Class

    Public Class AuxPair
        Public A As ViewMeasure
        Public B As ViewMeasure

        Public Sub New(a As ViewMeasure, b As ViewMeasure)
            Me.A = a
            Me.B = b
        End Sub
    End Class

    Public Class FitResult
        Public Scale As Double
        Public Rotate90 As Boolean
        Public ProjectedW As Double
        Public ProjectedH As Double
    End Class

    Public Class LayoutPattern
        Public MainSlot As SlotRect
        Public Aux1Slot As SlotRect
        Public Aux2Slot As SlotRect
        Public IsoSlot As SlotRect
    End Class

    Public Class LayoutPlan
        Public MainSlot As SlotRect
        Public Aux1Slot As SlotRect
        Public Aux2Slot As SlotRect
        Public IsoSlot As SlotRect

        Public MainFit As FitResult
        Public Aux1Fit As FitResult
        Public Aux2Fit As FitResult
        Public IsoFit As FitResult

        Public MainMeasure As ViewMeasure
        Public Aux1Measure As ViewMeasure
        Public Aux2Measure As ViewMeasure

        Public Score As Double
    End Class

    Public Class AlbumItem
        Public ModelPath As String = String.Empty
        Public Prompts As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
    End Class

End Class

' ================================================================
'  XLSX READER  (ZIP+XML через Reflection, без COM Excel)
' ================================================================
Public NotInheritable Class XlsxReader

    Public Shared Function Load(excelPath As String, workspacePath As String, sheetTab As String) As List(Of AlbumBuilder.AlbumItem)
        Dim result As New List(Of AlbumBuilder.AlbumItem)()
        Try
            Dim asmComp As System.Reflection.Assembly = Nothing
            Try
                asmComp = System.Reflection.Assembly.Load("System.IO.Compression, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
            Catch
            End Try
            If asmComp Is Nothing Then
                Try
                    asmComp = System.Reflection.Assembly.LoadWithPartialName("System.IO.Compression")
                Catch
                End Try
            End If

            Dim asmFS As System.Reflection.Assembly = Nothing
            Try
                asmFS = System.Reflection.Assembly.Load("System.IO.Compression.FileSystem, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
            Catch
            End Try
            If asmFS Is Nothing Then
                Try
                    asmFS = System.Reflection.Assembly.LoadWithPartialName("System.IO.Compression.FileSystem")
                Catch
                End Try
            End If

            Dim zipFileType As System.Type = Nothing
            If asmFS IsNot Nothing Then zipFileType = asmFS.GetType("System.IO.Compression.ZipFile")
            If zipFileType Is Nothing AndAlso asmComp IsNot Nothing Then zipFileType = asmComp.GetType("System.IO.Compression.ZipFile")
            If zipFileType Is Nothing Then zipFileType = System.Type.GetType("System.IO.Compression.ZipFile, System.IO.Compression.FileSystem")
            If zipFileType Is Nothing Then Throw New Exception("ZipFile тип не найден. Проверьте .NET Framework.")

            Dim zip As Object = Nothing
            Try
                zip = zipFileType.InvokeMember("OpenRead",
                    System.Reflection.BindingFlags.Static Or System.Reflection.BindingFlags.Public Or System.Reflection.BindingFlags.InvokeMethod,
                    Nothing, Nothing, New Object() {excelPath})
            Catch ix As System.Reflection.TargetInvocationException
                Dim inner As Exception = If(ix.InnerException, CType(ix, Exception))
                Throw New Exception("ZipFile.OpenRead: " & inner.Message)
            End Try
            If zip Is Nothing Then Throw New Exception("ZipFile.OpenRead вернул Nothing.")

            Try
                Dim sst As List(Of String) = ReadSharedStrings(zip)
                Dim sheetPath As String = FindSheetPath(zip, sheetTab)
                If String.IsNullOrEmpty(sheetPath) Then Throw New Exception("Лист '" & sheetTab & "' не найден.")

                Dim rows As List(Of List(Of String)) = ReadSheet(zip, sheetPath, sst)
                If rows.Count < 2 Then Throw New Exception("Лист пустой.")

                Dim hIdx As Integer = DetectHeader(rows)
                Dim hMap As Dictionary(Of String, Integer) = BuildMap(rows(hIdx))
                If Not hMap.ContainsKey("MODEL_PATH") Then
                    Dim h As New System.Text.StringBuilder()
                    For Each s As String In rows(hIdx)
                        If Not String.IsNullOrWhiteSpace(s) Then h.Append("[" & s & "] ")
                    Next
                    Throw New Exception("MODEL_PATH не найден. Заголовки: " & h.ToString())
                End If

                Dim mc As Integer = hMap("MODEL_PATH")
                Dim keys As String() = {"CODE","PROJECT_NAME","DRAWING_NAME","ORG_NAME","STAGE","SHEET","SHEETS"}

                For r As Integer = hIdx + 1 To rows.Count - 1
                    Dim row As List(Of String) = rows(r)
                    If row.Count <= mc Then Continue For
                    Dim raw As String = If(row(mc) IsNot Nothing, row(mc).Trim(), "")
                    If String.IsNullOrWhiteSpace(raw) Then Continue For

                    Dim resolved As String = ResolvePath(raw, workspacePath, excelPath)
                    If String.IsNullOrWhiteSpace(resolved) Then resolved = raw

                    Dim item As New AlbumBuilder.AlbumItem()
                    item.ModelPath = resolved
                    For Each key As String In keys
                        If hMap.ContainsKey(key) Then
                            Dim c As Integer = hMap(key)
                            If c < row.Count AndAlso row(c) IsNot Nothing Then item.Prompts(key) = row(c).Trim()
                        End If
                    Next
                    result.Add(item)
                Next
            Finally
                Try
                    Dim d As System.Reflection.MethodInfo = zip.GetType().GetMethod("Dispose")
                    If d IsNot Nothing Then d.Invoke(zip, Nothing)
                Catch
                End Try
            End Try

        Catch ex As Exception
            Dim real As Exception = ex
            Do While real.InnerException IsNot Nothing
                real = real.InnerException
            Loop
            System.Windows.Forms.MessageBox.Show("Ошибка чтения Excel:" & vbCrLf & real.Message & vbCrLf & "[" & real.GetType().Name & "]", "Ошибка")
        End Try
        Return result
    End Function

    Private Shared Function GetEntry(zip As Object, entryName As String) As String
        Dim ge As System.Reflection.MethodInfo = zip.GetType().GetMethod("GetEntry")
        Dim entry As Object = ge.Invoke(zip, New Object() {entryName})
        If entry Is Nothing Then Return Nothing
        Dim om As System.Reflection.MethodInfo = entry.GetType().GetMethod("Open")
        Dim stream As System.IO.Stream = CType(om.Invoke(entry, Nothing), System.IO.Stream)
        Using sr As New System.IO.StreamReader(stream, System.Text.Encoding.UTF8)
            Return sr.ReadToEnd()
        End Using
    End Function

    Private Shared Function FindSheetPath(zip As Object, tabName As String) As String
        Dim wb As String = GetEntry(zip, "xl/workbook.xml")
        If wb Is Nothing Then Return "xl/worksheets/sheet1.xml"
        Dim rId As String = ""
        Dim first As String = ""
        For Each m As System.Text.RegularExpressions.Match In
                System.Text.RegularExpressions.Regex.Matches(wb, "<sheet[^>]+name=""([^""]+)""[^>]+r:id=""([^""]+)""", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            If first = "" Then first = m.Groups(2).Value
            If String.Equals(m.Groups(1).Value, tabName, StringComparison.OrdinalIgnoreCase) Then
                rId = m.Groups(2).Value
                Exit For
            End If
        Next
        If rId = "" Then rId = first
        If rId = "" Then Return "xl/worksheets/sheet1.xml"
        Dim rels As String = GetEntry(zip, "xl/_rels/workbook.xml.rels")
        If rels IsNot Nothing Then
            Dim rm As System.Text.RegularExpressions.Match = System.Text.RegularExpressions.Regex.Match(rels, "Id=""" & rId & """[^>]+Target=""([^""]+)""")
            If rm.Success Then
                Dim t As String = rm.Groups(1).Value
                If Not t.StartsWith("xl/") Then t = "xl/" & t
                Return t
            End If
        End If
        Return "xl/worksheets/sheet1.xml"
    End Function

    Private Shared Function ReadSharedStrings(zip As Object) As List(Of String)
        Dim r As New List(Of String)()
        Dim xml As String = GetEntry(zip, "xl/sharedStrings.xml")
        If xml Is Nothing Then Return r
        For Each m As System.Text.RegularExpressions.Match In
                System.Text.RegularExpressions.Regex.Matches(xml, "<si>(.*?)</si>", System.Text.RegularExpressions.RegexOptions.Singleline)
            Dim sb As New System.Text.StringBuilder()
            For Each tm As System.Text.RegularExpressions.Match In
                    System.Text.RegularExpressions.Regex.Matches(m.Groups(1).Value, "<t(?:[^>]*)>(.*?)</t>", System.Text.RegularExpressions.RegexOptions.Singleline)
                sb.Append(XmlDecode(tm.Groups(1).Value))
            Next
            r.Add(sb.ToString())
        Next
        Return r
    End Function

    Private Shared Function ReadSheet(zip As Object, path As String, sst As List(Of String)) As List(Of List(Of String))
        Dim result As New List(Of List(Of String))()
        Dim xml As String = GetEntry(zip, path)
        If xml Is Nothing Then Return result
        For Each rowM As System.Text.RegularExpressions.Match In
                System.Text.RegularExpressions.Regex.Matches(xml, "<row[^>]*>(.*?)</row>", System.Text.RegularExpressions.RegexOptions.Singleline)
            Dim maxC As Integer = -1
            Dim cd As New Dictionary(Of Integer, String)()
            For Each cm As System.Text.RegularExpressions.Match In
                    System.Text.RegularExpressions.Regex.Matches(rowM.Groups(1).Value, "<c\s+r=""([A-Z]+)\d+""([^>]*)>(.*?)</c>", System.Text.RegularExpressions.RegexOptions.Singleline)
                Dim ci As Integer = ColIdx(cm.Groups(1).Value)
                Dim attrs As String = cm.Groups(2).Value
                Dim inner As String = cm.Groups(3).Value
                Dim ct As String = ""
                Dim tm As System.Text.RegularExpressions.Match = System.Text.RegularExpressions.Regex.Match(attrs, "\bt=""([^""]+)""")
                If tm.Success Then ct = tm.Groups(1).Value
                Dim val As String = ""
                Dim vm As System.Text.RegularExpressions.Match = System.Text.RegularExpressions.Regex.Match(inner, "<v>(.*?)</v>", System.Text.RegularExpressions.RegexOptions.Singleline)
                If vm.Success Then
                    Dim rv As String = XmlDecode(vm.Groups(1).Value)
                    If ct = "s" Then
                        Dim idx As Integer = 0
                        If Integer.TryParse(rv, idx) AndAlso idx < sst.Count Then
                            val = sst(idx)
                        End If
                    Else
                        val = rv
                    End If
                End If
                Dim ism As System.Text.RegularExpressions.Match = System.Text.RegularExpressions.Regex.Match(inner, "<is>.*?<t>(.*?)</t>.*?</is>", System.Text.RegularExpressions.RegexOptions.Singleline)
                If ism.Success Then val = XmlDecode(ism.Groups(1).Value)
                If ci > maxC Then maxC = ci
                cd(ci) = val
            Next
            If maxC >= 0 Then
                Dim row As New List(Of String)()
                For ci As Integer = 0 To maxC
                    If cd.ContainsKey(ci) Then
                        row.Add(cd(ci))
                    Else
                        row.Add("")
                    End If
                Next
                result.Add(row)
            End If
        Next
        Return result
    End Function

    Private Shared Function DetectHeader(rows As List(Of List(Of String))) As Integer
        For r As Integer = 0 To Math.Min(19, rows.Count - 1)
            If BuildMap(rows(r)).ContainsKey("MODEL_PATH") Then Return r
        Next
        Dim best As Integer = 0
        Dim bestN As Integer = 0
        For r As Integer = 0 To Math.Min(19, rows.Count - 1)
            Dim n As Integer = 0
            For Each v As String In rows(r)
                If Not String.IsNullOrWhiteSpace(v) Then n += 1
            Next
            If n > bestN Then
                bestN = n
                best = r
            End If
        Next
        Return best
    End Function

    Private Shared Function BuildMap(row As List(Of String)) As Dictionary(Of String, Integer)
        Dim m As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
        For c As Integer = 0 To row.Count - 1
            Dim k As String = MapAlias(row(c))
            If Not String.IsNullOrEmpty(k) AndAlso Not m.ContainsKey(k) Then m(k) = c
        Next
        Return m
    End Function

    Private Shared Function MapAlias(raw As String) As String
        If raw Is Nothing Then Return ""
        Dim n As String = raw.Trim().ToUpperInvariant()
        Select Case n
            Case "MODEL_PATH","MODEL","P","ПУТЬ","ФАЙЛ","МОДЕЛЬ","PATH","FILEPATH","FILE_PATH" : Return "MODEL_PATH"
            Case "CODE","ШИФР","АРТИКУЛ","ОБОЗНАЧЕНИЕ"                : Return "CODE"
            Case "PROJECT_NAME","PROJECT","ОБЪЕКТ","ПРОЕКТ"           : Return "PROJECT_NAME"
            Case "DRAWING_NAME","TITLE","НАИМЕНОВАНИЕ","ИМЯ ЧЕРТЕЖА" : Return "DRAWING_NAME"
            Case "ORG_NAME","ОРГАНИЗАЦИЯ","КОМПАНИЯ"                  : Return "ORG_NAME"
            Case "STAGE","СТАДИЯ"                                     : Return "STAGE"
            Case "SHEET","ЛИСТ"                                       : Return "SHEET"
            Case "SHEETS","ЛИСТОВ"                                    : Return "SHEETS"
        End Select
        Dim src As String() = {"А","Б","В","Г","Д","Е","Ё","Ж","З","И","Й","К","Л","М","Н","О","П","Р","С","Т","У","Ф","Х","Ц","Ч","Ш","Щ","Ъ","Ы","Ь","Э","Ю","Я"}
        Dim dst As String() = {"A","B","V","G","D","E","E","ZH","Z","I","Y","K","L","M","N","O","P","R","S","T","U","F","H","C","CH","SH","SCH","","Y","","E","YU","YA"}
        For i As Integer = 0 To src.Length - 1
            n = n.Replace(src(i), dst(i))
        Next
        n = System.Text.RegularExpressions.Regex.Replace(n, "[^A-Z0-9]+", "_").Trim("_"c)
        Select Case n
            Case "MODEL_PATH","MODEL","P" : Return "MODEL_PATH"
            Case "CODE","SHIFR"           : Return "CODE"
            Case "PROJECT_NAME","OBEKT"   : Return "PROJECT_NAME"
            Case "DRAWING_NAME","NAIMENOVANIE" : Return "DRAWING_NAME"
            Case "ORG_NAME","ORGANIZACIYA" : Return "ORG_NAME"
            Case "STAGE","STADIYA"        : Return "STAGE"
            Case "SHEET","LIST"           : Return "SHEET"
            Case "SHEETS","LISTOV"        : Return "SHEETS"
        End Select
        Return n
    End Function

    Private Shared Function ResolvePath(inp As String, ws As String, xlPath As String) As String
        If System.IO.File.Exists(inp) Then Return System.IO.Path.GetFullPath(inp)
        If Not String.IsNullOrWhiteSpace(ws) Then
            Dim c As String = System.IO.Path.Combine(ws, inp)
            If System.IO.File.Exists(c) Then Return System.IO.Path.GetFullPath(c)
        End If
        Dim xd As String = System.IO.Path.GetDirectoryName(xlPath)
        If Not String.IsNullOrWhiteSpace(xd) Then
            Dim c As String = System.IO.Path.Combine(xd, inp)
            If System.IO.File.Exists(c) Then Return System.IO.Path.GetFullPath(c)
        End If
        Dim fn As String = System.IO.Path.GetFileName(inp)
        If Not fn.EndsWith(".ipt", StringComparison.OrdinalIgnoreCase) Then fn &= ".ipt"
        For Each root As String In {ws, xd}
            If String.IsNullOrWhiteSpace(root) Then Continue For
            Dim d As String = System.IO.Path.Combine(root, fn)
            If System.IO.File.Exists(d) Then Return System.IO.Path.GetFullPath(d)
            Try
                For Each sub1 As String In System.IO.Directory.GetDirectories(root)
                    Dim s As String = System.IO.Path.Combine(sub1, fn)
                    If System.IO.File.Exists(s) Then Return System.IO.Path.GetFullPath(s)
                Next
            Catch
            End Try
        Next
        Return String.Empty
    End Function

    Private Shared Function ColIdx(col As String) As Integer
        Dim idx As Integer = 0
        For Each ch As Char In col.ToUpper()
            idx = idx * 26 + (AscW(ch) - AscW("A"c) + 1)
        Next
        Return idx - 1
    End Function

    Private Shared Function XmlDecode(s As String) As String
        If s Is Nothing Then Return ""
        Return s.Replace("&amp;","&").Replace("&lt;","<").Replace("&gt;",">").Replace("&quot;","""").Replace("&apos;","'")
    End Function

End Class

