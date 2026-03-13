' ================================================================
' StoneAlbumRule.ilogic.vb  –  v3.11
' Архитектура точно повторяет рабочий VBA RKM_IdwAlbum.bas
' Источник: vba-inventor / RKM_IdwAlbum.bas, RKM_FrameBorder.bas,
'           RKM_TitleBlockPrompted.bas, RKM_Excel.bas
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
    If newExcel Is Nothing Then Return
    If Not String.IsNullOrWhiteSpace(newExcel) Then excelPath = newExcel.Trim()
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
    If newWs Is Nothing Then Return
    If Not String.IsNullOrWhiteSpace(newWs) Then workspacePath = newWs.Trim()

    Dim newSheet As String = InputBox(
        "Имя листа в Excel:", "Шаг 3 из 3 — Лист Excel", sheetTabName)
    If newSheet Is Nothing Then Return
    If Not String.IsNullOrWhiteSpace(newSheet) Then sheetTabName = newSheet.Trim()

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
                If String.IsNullOrWhiteSpace(item.Prompts("SHEET"))  Then item.Prompts("SHEET")  = (i + 1).ToString()
                If String.IsNullOrWhiteSpace(item.Prompts("SHEETS")) Then item.Prompts("SHEETS") = items.Count.ToString()

                Dim ok As Boolean = BuildOneSheet(doc, item)
                If ok Then okCount += 1 Else failCount += 1
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
    Private Function BuildOneSheet(doc As DrawingDocument, item As AlbumItem) As Boolean
        If Not System.IO.File.Exists(item.ModelPath) Then
            Debug.Print("WARN: модель не найдена: " & item.ModelPath)
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
            _app.SilentOperation = True
            Try
                If sheet.Border IsNot Nothing Then sheet.Border.Delete()
            Catch
            End Try
            Dim borderOk As Boolean = False
            If borderDef IsNot Nothing Then
                Try
                    sheet.AddBorder(borderDef)
                    borderOk = True
                    Debug.Print("AddBorder OK на листе: " & sheet.Name)
                Catch ex As Exception
                    Debug.Print("WARN AddBorder: " & ex.Message)
                End Try
            End If
            If Not borderOk Then
                Debug.Print("WARN: рамка НЕ применилась на листе: " & sheet.Name)
            End If
            _app.SilentOperation = False

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
            _app.SilentOperation = True
            Try
                If sheet.TitleBlock IsNot Nothing Then sheet.TitleBlock.Delete()
            Catch
            End Try
            Dim tbOk As Boolean = False
            If tbDef IsNot Nothing Then
                Try
                    sheet.AddTitleBlock(tbDef, , ps)
                    tbOk = True
                    Debug.Print("AddTitleBlock OK на листе: " & sheet.Name)
                Catch ex As Exception
                    Debug.Print("WARN AddTitleBlock: " & ex.Message)
                End Try
            End If
            If Not tbOk Then
                Debug.Print("WARN: штамп НЕ применился на листе: " & sheet.Name)
            End If
            _app.SilentOperation = False

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
                    Debug.Print("WARN Open: " & ex.Message)
                Finally
                    _app.SilentOperation = False
                End Try
            End If
            If modelDoc Is Nothing Then
                Debug.Print("WARN: не удалось открыть: " & item.ModelPath)
                Try
                    sheet.Delete()
                Catch
                End Try
                Return False
            End If

            ' Виды через slot-based layout
            Dim viewsOk As Boolean = PlaceViewsSlotBased(doc, sheet, modelDoc)
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
            Debug.Print("ERROR: BuildOneSheet: " & ex.Message)
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
    '  SLOT-BASED LAYOUT  (точный порт GetTechSlotRects / PlaceTechViewLayout)
    ' ================================================================
    Private Function PlaceViewsSlotBased(doc As DrawingDocument, sheet As Sheet, modelDoc As Document) As Boolean
        ' 1. Безопасная зона
        Dim safeRect As SlotRect = GetSheetSafeRect(sheet)

        ' 2. Сжимаем на LAYOUT_PAD_MM
        Dim pad   As Double = Cm(doc, LAYOUT_PAD_MM)
        Dim pRect As SlotRect = InsetRect(safeRect, pad)

        ' 3. Делим на 4 слота (как VBA GetTechSlotRects)
        Dim gap   As Double = Cm(doc, GAP_MM)
        Dim splitX As Double = pRect.R - RectW(pRect) * TECH_RIGHT_COL_RATIO
        Dim splitY As Double = pRect.T - RectH(pRect) * TECH_TOP_BAND_RATIO
        Dim smallRight As Double = pRect.L + (splitX - pRect.L) * TECH_SMALL_SLOT_RATIO

        ' rawSlots
        Dim slotSmall As SlotRect = New SlotRect(pRect.L,        smallRight - gap, splitY + gap, pRect.T)
        Dim slotWide  As SlotRect = New SlotRect(smallRight + gap, pRect.R,        splitY + gap, pRect.T)
        Dim slotLarge As SlotRect = New SlotRect(pRect.L,        splitX - gap,    pRect.B,       splitY - gap)
        Dim slotIso   As SlotRect = New SlotRect(splitX + gap,   pRect.R,         pRect.B,       splitY - gap)

        ' contentSlots (с паддингом и отступом на подпись)
        Dim cPad    As Double = Cm(doc, SLOT_CONTENT_PAD_MM)
        Dim capClear As Double = Cm(doc, CAPTION_CLEAR_TOP_MM)

        Dim csSmall As SlotRect = ContentSlot(slotSmall, cPad, capClear)
        Dim csWide  As SlotRect = ContentSlot(slotWide,  cPad, capClear)
        Dim csLarge As SlotRect = ContentSlot(slotLarge, cPad, capClear)
        Dim csIso   As SlotRect = ContentSlot(slotIso,   cPad, cPad)    ' ISO — без подписи

        ' 4. Измеряем 4 ориентации через probe-виды
        Dim mFront As ViewMeasure = MeasureView(sheet, modelDoc, ViewOrientationTypeEnum.kFrontViewOrientation, DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle)
        Dim mTop   As ViewMeasure = MeasureView(sheet, modelDoc, ViewOrientationTypeEnum.kTopViewOrientation,   DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle)
        Dim mRight As ViewMeasure = MeasureView(sheet, modelDoc, ViewOrientationTypeEnum.kRightViewOrientation, DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle)
        Dim mIso   As ViewMeasure = MeasureView(sheet, modelDoc, ViewOrientationTypeEnum.kIsoTopRightViewOrientation, DrawingViewStyleEnum.kShadedDrawingViewStyle)

        If mFront Is Nothing OrElse mTop Is Nothing OrElse mRight Is Nothing Then
            Debug.Print("WARN: не удалось измерить виды")
            Return False
        End If

        ' 5. v3.10: Умный матчинг видов → слотам (перебор 6 перестановок)
        Dim measures() As ViewMeasure = {mFront, mTop, mRight}
        Dim slots() As SlotRect = {csLarge, csSmall, csWide}
        Dim perms()() As Integer = {
            New Integer(){0,1,2}, New Integer(){0,2,1},
            New Integer(){1,0,2}, New Integer(){1,2,0},
            New Integer(){2,0,1}, New Integer(){2,1,0}
        }
        Dim bestPerm() As Integer = perms(0)
        Dim bestScore As Double = -1
        For Each perm As Integer() In perms
            Dim score As Double = 0
            Dim valid As Boolean = True
            For i As Integer = 0 To 2
                Dim sc As Double = ScaleToFit(slots(i), measures(perm(i)), ORTHO_SCALE_MARGIN)
                If sc <= 0 Then
                    valid = False
                    Exit For
                End If
                score += sc
            Next
            If valid AndAlso score > bestScore Then
                bestScore = score
                bestPerm = perm
            End If
        Next

        Dim largeScale As Double = ScaleToFit(csLarge, measures(bestPerm(0)), ORTHO_SCALE_MARGIN)
        Dim smallScale As Double = ScaleToFit(csSmall, measures(bestPerm(1)), ORTHO_SCALE_MARGIN)
        Dim wideScale  As Double = ScaleToFit(csWide,  measures(bestPerm(2)), ORTHO_SCALE_MARGIN)
        Dim isoScale   As Double = 0

        If largeScale <= 0 OrElse smallScale <= 0 OrElse wideScale <= 0 Then
            Debug.Print("WARN: не удалось подобрать масштаб")
            Return False
        End If

        If mIso IsNot Nothing Then
            isoScale = ScaleToFit(csIso, mIso, ISO_SCALE_MARGIN)
        End If
        Debug.Print("Scales: LARGE=" & largeScale & " SMALL=" & smallScale & " WIDE=" & wideScale & " ISO=" & isoScale & " perm=" & bestPerm(0) & bestPerm(1) & bestPerm(2))

        ' 6. Размещаем виды в слоты (с учётом найденной перестановки)
        Dim v1 As DrawingView = PlaceViewInSlot(sheet, modelDoc, measures(bestPerm(0)), largeScale, csLarge)
        AddDimNotes(doc, sheet, v1, csLarge)
        Dim v2 As DrawingView = PlaceViewInSlot(sheet, modelDoc, measures(bestPerm(1)), smallScale, csSmall)
        AddDimNotes(doc, sheet, v2, csSmall)
        Dim v3 As DrawingView = PlaceViewInSlot(sheet, modelDoc, measures(bestPerm(2)), wideScale,  csWide)
        AddDimNotes(doc, sheet, v3, csWide)

        If v1 Is Nothing OrElse v2 Is Nothing OrElse v3 Is Nothing Then
            Debug.Print("WARN: не все основные виды размещены")
            If v1 Is Nothing AndAlso v2 Is Nothing AndAlso v3 Is Nothing Then Return False
        End If

        ' ISO
        If isoScale > 0 AndAlso mIso IsNot Nothing Then
            Try
                Dim isoV As DrawingView = PlaceViewInSlot(sheet, modelDoc, mIso, isoScale, csIso)
                AddDimNotes(doc, sheet, isoV, csIso)
            Catch
            End Try
        End If

        Return True
    End Function

    ' v3.10: Размеры через DrawingNotes — реальные мм из масштаба вида
    Private Sub AddDimNotes(doc As DrawingDocument, sheet As Sheet,
                             v As DrawingView, slot As SlotRect)
        If v Is Nothing Then Return
        Try
            Dim sc As Double = v.Scale
            If sc <= 0.0001 Then Return
            Dim realW As Double = Math.Round(v.Width  / sc * 10.0)
            Dim realH As Double = Math.Round(v.Height / sc * 10.0)
            Dim cx As Double = (slot.L + slot.R) / 2.0
            Dim cy As Double = (slot.B + slot.T) / 2.0
            ' Горизонтальный размер — НАД видом
            sheet.DrawingNotes.Add(
                _app.TransientGeometry.CreatePoint2d(cx, slot.T + Cm(doc, 5.0)),
                String.Format("{0:F0}", realW))
            ' Вертикальный размер — СПРАВА от вида
            If realH > 1 Then
                sheet.DrawingNotes.Add(
                    _app.TransientGeometry.CreatePoint2d(slot.R + Cm(doc, 6.0), cy),
                    String.Format("{0:F0}", realH))
            End If
        Catch ex As Exception
            Debug.Print("WARN AddDimNotes: " & ex.Message)
        End Try
    End Sub

    ' Создаёт probe-вид, замеряет размеры при масштабе 1:1, удаляет
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
            _app.ActiveDocument.Update2(True)
            Dim m As New ViewMeasure()
            m.UnitW = probe.Width  / PROBE_SCALE
            m.UnitH = probe.Height / PROBE_SCALE
            m.Orient = orient
            m.Style  = style
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

    ' Подбирает масштаб вписания вида в слот (учитывает поворот на 90°)
    Private Function ScaleToFit(slot As SlotRect, m As ViewMeasure, margin As Double) As Double
        If m Is Nothing Then Return 0
        Dim sw As Double = slot.R - slot.L
        Dim sh As Double = slot.T - slot.B
        If sw <= 0 OrElse sh <= 0 Then Return 0
        Dim sc0 As Double = 0
        Dim sc90 As Double = 0
        If m.UnitW > 0.0001 AndAlso m.UnitH > 0.0001 Then
            sc0 = Math.Min(sw / m.UnitW, sh / m.UnitH) * margin
        End If
        ' Попробуем поворот 90°
        If m.UnitH > 0.0001 AndAlso m.UnitW > 0.0001 Then
            sc90 = Math.Min(sw / m.UnitH, sh / m.UnitW) * margin
        End If
        Dim best As Double = Math.Max(sc0, sc90)
        If best > MAX_AUTO_SCALE Then best = MAX_AUTO_SCALE
        If best < 0.02 Then Return 0
        Return best
    End Function

    ' Размещает вид в центр слота с заданным масштабом (без shrink-цикла)
    Private Function PlaceViewInSlot(sheet As Sheet, modelDoc As Document,
                                     m As ViewMeasure, scaleVal As Double,
                                     slot As SlotRect) As DrawingView
        If m Is Nothing Then Return Nothing
        Dim cx As Double = (slot.L + slot.R) / 2.0
        Dim cy As Double = (slot.B + slot.T) / 2.0
        Dim v As DrawingView = Nothing
        Try
            v = sheet.DrawingViews.AddBaseView(
                modelDoc,
                _app.TransientGeometry.CreatePoint2d(cx, cy),
                scaleVal, m.Orient, m.Style)
            If v Is Nothing Then Return Nothing
            Try
                v.ShowLabel = False
            Catch
            End Try
            Return v
        Catch ex As Exception
            Debug.Print("WARN PlaceViewInSlot sc=" & scaleVal & ": " & ex.Message)
            If v IsNot Nothing Then
                Try
                    v.Delete()
                Catch
                End Try
            End If
            Return Nothing
        End Try
    End Function

    ' Проверяет вписание вида в слот через Left/Top (VBA-стиль)
    ' v.Left = левый край, v.Top = верхний край (Y снизу вверх)
    ' bounds: Left=v.Left, Right=v.Left+v.Width, Bottom=v.Top-v.Height, Top=v.Top
    Private Function ViewFitsSlot(v As DrawingView, slot As SlotRect) As Boolean
        Dim vL As Double = v.Left
        Dim vT As Double = v.Top
        Dim vR As Double = vL + v.Width
        Dim vB As Double = vT - v.Height
        Dim eps As Double = 0.001
        Return (vL >= slot.L - eps AndAlso vR <= slot.R + eps AndAlso
                vB >= slot.B - eps AndAlso vT <= slot.T + eps)
    End Function

    ' SafeArea по реальному TitleBlock.RangeBox (как VBA GetSheetSafeRectCm)
    Private Function GetSheetSafeRect(sheet As Sheet) As SlotRect
        Dim w As Double = sheet.Width
        Dim h As Double = sheet.Height

        Dim wL As Double = w * SAFE_LEFT_RATIO
        Dim wR As Double = w * (1.0 - SAFE_RIGHT_RATIO)
        Dim wB As Double = h * SAFE_BOTTOM_RATIO
        Dim wT As Double = h * (1.0 - SAFE_TOP_RATIO)

        ' Учитываем реальный TitleBlock
        Try
            If sheet.TitleBlock IsNot Nothing Then
                Dim rb As Box2d = sheet.TitleBlock.RangeBox
                If rb IsNot Nothing Then
                    ' Штамп в нижнем правом углу — MaxPoint.Y = верхний край штампа (в см)
                    Dim tbTop As Double = Math.Max(rb.MinPoint.Y, rb.MaxPoint.Y)
                    Dim tbGapY As Double = h * TITLEBLOCK_GAP_RATIO
                    Debug.Print("SafeArea: sheet h=" & h & " tbTop=" & tbTop & " gap=" & tbGapY & " wB before=" & wB)
                    If tbTop + tbGapY > wB Then
                        wB = tbTop + tbGapY
                    End If
                    Debug.Print("SafeArea: wB after=" & wB)
                End If
            End If
        Catch ex As Exception
            Debug.Print("SafeArea TB err: " & ex.Message)
        End Try

        ' Fallback если зона вырождена (высота < 20% листа)
        If wT <= wB + h * 0.2 Then
            wB = h * 0.22   ' 22% от низа = ~65мм — больше чем штамп 55мм
            wT = h * 0.96
            Debug.Print("SafeArea: FALLBACK wB=" & wB & " wT=" & wT)
        End If

        Debug.Print("SafeArea final: L=" & wL & " R=" & wR & " B=" & wB & " T=" & wT & " W=" & (wR-wL) & " H=" & (wT-wB))
        Return New SlotRect(wL, wR, wB, wT)
    End Function

    ' Слот с учётом отступа на подпись сверху
    Private Function ContentSlot(raw As SlotRect, padCm As Double, topClear As Double) As SlotRect
        Dim t As Double = raw.T - topClear
        Dim b As Double = raw.B + padCm
        If t <= b + padCm Then
            t = raw.T - padCm
            b = raw.B + padCm
        End If
        Return New SlotRect(raw.L + padCm, raw.R - padCm, b, t)
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
        If def IsNot Nothing Then Return def   ' v3.10: существует — не перерисовываем

        Try
            def = doc.BorderDefinitions.Add(BORDER_NAME)
        Catch ex As Exception
            Debug.Print("WARN BorderDef.Add: " & ex.Message)
            Return Nothing
        End Try

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
        If def IsNot Nothing Then Return def   ' v3.10: существует — не перерисовываем

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


