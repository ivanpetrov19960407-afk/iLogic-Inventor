' ================================================================
' StoneAlbumRule.ilogic.vb  –  v2.0
' Автоматическое создание/обновление IDW-альбома из Excel
' Изделия из натурального камня: ступени, плиты, элементы фасада
'
' Источник логики: vba-inventor / RKM_IdwAlbum.bas (PR #51)
'                  vba-inventor / RKM_Excel.bas
'                  vba-inventor / RKM_FrameBorder.bas
'                  vba-inventor / RKM_TitleBlockPrompted.bas
' ================================================================

Option Explicit On

Imports Inventor
Imports System
Imports System.Collections.Generic
' Imports System.IO.Compression -- не используется, сборка загружается через Assembly.LoadWithPartialName
' Imports System.IO  -- убрано: конфликт Path/File с Inventor (используем полные пути System.IO.Path / System.IO.File)

' ================================================================
'  ТОЧКА ВХОДА
' ================================================================
Sub Main()
    ' ── 1. Читаем сохранённые пути из кастомных свойств документа ──
    Dim excelPath     As String = String.Empty
    Dim workspacePath As String = String.Empty
    Dim sheetTabName  As String = "ALBUM"

    Try : excelPath     = iProperties.Value("Custom", "AlbumExcel")     : Catch : End Try
    Try : workspacePath = iProperties.Value("Custom", "AlbumWorkspace") : Catch : End Try
    Try : sheetTabName  = iProperties.Value("Custom", "AlbumSheet")     : Catch : End Try

    ' ── 2. Диалог настроек — показываем всегда, подставляем текущие значения ──
    Dim newExcel As String = InputBox(
        "Путь к Excel-файлу альбома (.xlsx):" & vbCrLf &
        "(оставьте как есть или введите новый)",
        "Шаг 1 из 3 — Excel",
        excelPath)
    If newExcel Is Nothing Then Return          ' нажали Отмена
    If Not String.IsNullOrWhiteSpace(newExcel) Then excelPath = newExcel.Trim()
    If String.IsNullOrWhiteSpace(excelPath) Then
        System.Windows.Forms.MessageBox.Show("Путь к Excel не указан. Отмена.", "Ошибка")
        Return
    End If

    ' Дефолт папки — рядом с Excel
    If String.IsNullOrWhiteSpace(workspacePath) Then
        workspacePath = System.IO.Path.GetDirectoryName(excelPath)
    End If

    Dim newWorkspace As String = InputBox(
        "Папка с 3D-моделями (.ipt):" & vbCrLf &
        "(оставьте как есть или введите новый путь)",
        "Шаг 2 из 3 — Папка моделей",
        workspacePath)
    If newWorkspace Is Nothing Then Return
    If Not String.IsNullOrWhiteSpace(newWorkspace) Then workspacePath = newWorkspace.Trim()

    Dim newSheet As String = InputBox(
        "Имя листа в Excel с данными альбома:",
        "Шаг 3 из 3 — Лист Excel",
        sheetTabName)
    If newSheet Is Nothing Then Return
    If Not String.IsNullOrWhiteSpace(newSheet) Then sheetTabName = newSheet.Trim()

    ' ── 3. Сохраняем пути обратно в свойства документа (на следующий запуск) ──
    Try : iProperties.Value("Custom", "AlbumExcel")     = excelPath     : Catch : End Try
    Try : iProperties.Value("Custom", "AlbumWorkspace") = workspacePath : Catch : End Try
    Try : iProperties.Value("Custom", "AlbumSheet")     = sheetTabName  : Catch : End Try

    ' ── 4. Запуск ──
    Dim doc As DrawingDocument = TryCast(ThisApplication.ActiveDocument, DrawingDocument)
    If doc Is Nothing Then
        System.Windows.Forms.MessageBox.Show("Откройте .idw документ перед запуском правила.", "Ошибка")
        Return
    End If

    Dim rule As New StoneAlbumRule(ThisApplication)
    rule.BuildAlbumFromExcel(doc, excelPath, workspacePath, sheetTabName)
End Sub

' ================================================================
'  ГЛАВНЫЙ КЛАСС
' ================================================================
Public Class StoneAlbumRule

    Private ReadOnly _app As Inventor.Application

    ' --- Геометрия А3 (мм) ---
    Private Const A3_W_MM     As Double = 420.0
    Private Const A3_H_MM     As Double = 297.0
    Private Const FRAME_L_MM  As Double = 20.0    ' левое поле (под подшивку)
    Private Const FRAME_O_MM  As Double = 5.0     ' остальные поля
    Private Const TB_W_MM     As Double = 185.0   ' ширина штампа
    Private Const TB_H_MM     As Double = 55.0    ' высота штампа

    ' --- Имена ресурсов ---
    Private Const BORDER_NAME     As String = "RKM_SPDS_A3_BORDER_V12"
    Private Const TITLEBLOCK_NAME As String = "RKM_SPDS_A3_FORM3_V17"
    Private Const SHEET_PREFIX    As String = "ALB_"

    ' --- Компоновка ---
    Private Const GAP_MM         As Double = 8.0
    Private Const PAD_MM         As Double = 6.0
    Private Const TOP_ROW_RATIO  As Double = 0.36
    Private Const SIDE_COL_RATIO As Double = 0.34
    Private Const LABEL_OFFSET   As Double = 4.0   ' мм, отступ подписи вида вниз
    Private Const LABEL_MARGIN   As Double = 2.0   ' мм, отступ подписи влево

    ' --- Масштабный ряд (от крупного к мелкому) ---
    Private Shared ReadOnly SCALE_CANDIDATES As Double() = {
        5.0, 4.0, 3.0, 2.0, 1.5, 1.25, 1.0,
        0.9, 0.8, 0.75, 0.67, 0.5, 0.4, 0.33, 0.25, 0.2, 0.1
    }
    Private Const SCALE_FIT_MARGIN As Double = 0.95
    Private Const MIN_SCALE        As Double = 0.05

    Public Sub New(app As Inventor.Application)
        _app = app
    End Sub

    ' ================================================================
    '  ВНЕШНЯЯ ТОЧКА ВХОДА
    ' ================================================================
    Public Sub BuildAlbumFromExcel(
            doc As DrawingDocument,
            excelPath As String,
            workspacePath As String,
            sheetTabName As String)

        Dim items As List(Of AlbumItem) = ExcelLoader.Load(excelPath, workspacePath, sheetTabName)
        If items.Count = 0 Then
            System.Windows.Forms.MessageBox.Show(
                "Excel не содержит строк с MODEL_PATH." & vbCrLf & vbCrLf &
                "Файл: " & excelPath & vbCrLf &
                "Лист: " & sheetTabName & vbCrLf & vbCrLf &
                "Убедитесь что:" & vbCrLf &
                "  1. Лист называется '" & sheetTabName & "'" & vbCrLf &
                "  2. Первая строка содержит заголовки (MODEL_PATH, CODE, ...)" & vbCrLf &
                "  3. Ниже есть строки с данными",
                "Ошибка — пустой список",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Warning)
            Return
        End If

        Try
            _app.SilentOperation = True

            ' Убедиться что ресурсы (рамка+штамп) есть в документе
            Dim borderDef    As BorderDefinition    = EnsureBorderDefinition(doc)
            Dim titleDef     As TitleBlockDefinition = EnsureTitleBlockDefinition(doc)

            If borderDef Is Nothing OrElse titleDef Is Nothing Then
                Throw New Exception("Не удалось создать ресурсы рамки/штампа.")
            End If

            ' Пересобрать листы
            For i As Integer = 0 To items.Count - 1
                Dim item As AlbumItem = items(i)
                ' Дополняем авто-номерацию листов если не задана
                If String.IsNullOrWhiteSpace((If(item.Prompts.ContainsKey("SHEET"), item.Prompts("SHEET"), String.Empty))) Then
                    item.Prompts("SHEET") = (i + 1).ToString()
                End If
                If String.IsNullOrWhiteSpace((If(item.Prompts.ContainsKey("SHEETS"), item.Prompts("SHEETS"), String.Empty))) Then
                    item.Prompts("SHEETS") = items.Count.ToString()
                End If
                BuildSingleSheet(doc, item, borderDef, titleDef)
            Next

            ' Удалить устаревшие ALB_* листы которых нет в текущем списке
            RemoveStaleSheets(doc, items)

            doc.Update2(True)

        Finally
            _app.SilentOperation = False
        End Try

        ' Подсчёт: сколько листов с видами
        Dim sheetsWithViews As Integer = 0
        Dim sheetsNoModel   As Integer = 0
        For Each it As AlbumItem In items
            If System.IO.File.Exists(it.ModelPath) Then
                sheetsWithViews += 1
            Else
                sheetsNoModel += 1
            End If
        Next
        Dim summary As String = "Альбом собран: " & items.Count & " листов."
        If sheetsNoModel > 0 Then
            summary &= vbCrLf & vbCrLf &
                "⚠ " & sheetsNoModel & " листов без видов — модели не найдены." & vbCrLf &
                "Укажите правильную папку с .ipt файлами при следующем запуске."
        End If
        System.Windows.Forms.MessageBox.Show(summary, "Готово",
            System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information)
    End Sub

    ' ================================================================
    '  ОДИН ЛИСТ
    ' ================================================================
    Private Sub BuildSingleSheet(
            doc As DrawingDocument,
            item As AlbumItem,
            borderDef As BorderDefinition,
            titleDef As TitleBlockDefinition)

        Dim modelDoc As Document = Nothing
        Try
            ' ── Создаём / находим лист ──
            Dim sheetName As String = SHEET_PREFIX & System.IO.Path.GetFileNameWithoutExtension(item.ModelPath)
            Dim sheet As Sheet = EnsureSheet(doc, sheetName)
            sheet.Activate()

            ' Формат А3 альбомная
            Try
                sheet.Size        = DrawingSheetSizeEnum.kA3DrawingSheetSize
                sheet.Orientation = PageOrientationTypeEnum.kLandscapePageOrientation
            Catch : End Try

            ' Убрать старые виды
            RemoveAllViews(sheet)

            ' ── Рамка и штамп — всегда ──
            ApplyBorder(sheet, borderDef)
            ApplyTitleBlock(sheet, titleDef, item.Prompts)

            ' ── Виды — только если модель найдена ──
            If Not System.IO.File.Exists(item.ModelPath) Then
                Logger.Warn("Модель не найдена, лист создан без видов: " & item.ModelPath)
                Return
            End If

            modelDoc = _app.Documents.Open(item.ModelPath, False)
            BuildViews(doc, sheet, modelDoc)

        Catch ex As Exception
            Logger.Error("Лист не собран для " & item.ModelPath & ": " & ex.Message)
        Finally
            If modelDoc IsNot Nothing Then
                Try
                    modelDoc.Close(True)
                Catch
                End Try
            End If
        End Try
    End Sub

    ' ================================================================
    '  УДАЛЕНИЕ УСТАРЕВШИХ ЛИСТОВ
    ' ================================================================
    Private Sub RemoveStaleSheets(doc As DrawingDocument, items As List(Of AlbumItem))
        Dim activeNames As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        For Each item As AlbumItem In items
            activeNames.Add(SHEET_PREFIX & System.IO.Path.GetFileNameWithoutExtension(item.ModelPath))
        Next

        Dim toDelete As New List(Of Sheet)()
        For Each s As Sheet In doc.Sheets
            If s.Name.StartsWith(SHEET_PREFIX, StringComparison.OrdinalIgnoreCase) Then
                If Not activeNames.Contains(s.Name) Then toDelete.Add(s)
            End If
        Next

        For Each s As Sheet In toDelete
            Try
                s.Delete()
            Catch
            End Try
        Next
    End Sub

    ' ================================================================
    '  РАЗМЕЩЕНИЕ ВИДОВ  (портированный алгоритм из RKM_IdwAlbum VBA)
    ' ================================================================
    Private Sub BuildViews(doc As DrawingDocument, sheet As Sheet, modelDoc As Document)
        Dim firstAngle As Boolean = GetProjectionStandard(doc)

        Dim blockedRect As ViewRect = GetTitleBlockRect(doc)
        Dim frontRect   As ViewRect = GetFrontViewRect(doc, firstAngle)
        Dim topRect     As ViewRect = GetTopViewRect(doc, firstAngle)
        Dim sideRect    As ViewRect = GetSideViewRect(doc, firstAngle, frontRect)

        ' Выбираем лучший масштаб из кандидатов
        Dim selectedScale As Double = MIN_SCALE
        For Each sc As Double In SCALE_CANDIDATES
            If ProbeLayout(sheet, modelDoc, sc, frontRect, topRect, sideRect, blockedRect, firstAngle) Then
                selectedScale = sc
                Exit For
            End If
        Next

        Logger.Info("Выбран масштаб: " & selectedScale)

        Dim ok As Boolean = PlaceViews(doc, sheet, modelDoc, selectedScale, frontRect, topRect, sideRect, blockedRect, firstAngle)
        If Not ok Then
            ' Fallback на минимальный масштаб
            PlaceViews(doc, sheet, modelDoc, MIN_SCALE, frontRect, topRect, sideRect, blockedRect, firstAngle)
        End If
    End Sub

    ' --- Зондирование: создаём виды, проверяем, сразу удаляем ---
    Private Function ProbeLayout(
            sheet As Sheet,
            modelDoc As Document,
            scale As Double,
            frontRect As ViewRect, topRect As ViewRect, sideRect As ViewRect,
            blockedRect As ViewRect,
            firstAngle As Boolean) As Boolean

        Dim baseView As DrawingView = Nothing
        Dim topView  As DrawingView = Nothing
        Dim sideView As DrawingView = Nothing

        Try
            baseView = TryCreateBaseView(sheet, modelDoc, scale)
            If baseView Is Nothing Then Return False

            ' Подбираем точную позицию базового вида
            baseView.Center = Pt2d(RectCenterX(frontRect), RectCenterY(frontRect))
            Dim recalcScale As Double = RecalculateScale(baseView, scale, frontRect)
            If recalcScale < scale Then
                ' Пересоздаём с пересчитанным масштабом
                baseView.Delete()
                baseView = TryCreateBaseView(sheet, modelDoc, recalcScale)
                If baseView Is Nothing Then Return False
                baseView.Center = Pt2d(RectCenterX(frontRect), RectCenterY(frontRect))
            End If

            If Not ViewFitsRect(baseView, frontRect) OrElse ViewIntersectsRect(baseView, blockedRect) Then
                Return False
            End If

            topView  = TryAddProjectedView(sheet, baseView, BuildTopTarget(baseView, topRect, firstAngle))
            sideView = TryAddProjectedView(sheet, baseView, BuildSideTarget(baseView, sideRect))

            Return (topView IsNot Nothing AndAlso
                    sideView IsNot Nothing AndAlso
                    ViewFitsRect(topView,  topRect)  AndAlso Not ViewIntersectsRect(topView,  blockedRect) AndAlso
                    ViewFitsRect(sideView, sideRect) AndAlso Not ViewIntersectsRect(sideView, blockedRect))
        Finally
            SafeDelete(sideView)
            SafeDelete(topView)
            SafeDelete(baseView)
        End Try
    End Function

    ' --- Финальное размещение ---
    Private Function PlaceViews(
            doc As DrawingDocument,
            sheet As Sheet,
            modelDoc As Document,
            scale As Double,
            frontRect As ViewRect, topRect As ViewRect, sideRect As ViewRect,
            blockedRect As ViewRect,
            firstAngle As Boolean) As Boolean

        Dim baseView As DrawingView = TryCreateBaseView(sheet, modelDoc, scale)
        If baseView Is Nothing Then Return False

        ' Пересчёт масштаба чтобы вид точно влез
        baseView.Center = Pt2d(RectCenterX(frontRect), RectCenterY(frontRect))
        Dim recalcScale As Double = RecalculateScale(baseView, scale, frontRect)
        If recalcScale < scale Then
            baseView.Delete()
            baseView = TryCreateBaseView(sheet, modelDoc, recalcScale)
            If baseView Is Nothing Then Return False
            baseView.Center = Pt2d(RectCenterX(frontRect), RectCenterY(frontRect))
        End If

        Dim topView  As DrawingView = TryAddProjectedView(sheet, baseView, BuildTopTarget(baseView, topRect, firstAngle))
        Dim sideView As DrawingView = TryAddProjectedView(sheet, baseView, BuildSideTarget(baseView, sideRect))

        If topView Is Nothing OrElse sideView Is Nothing Then
            SafeDelete(sideView)
            SafeDelete(topView)
            SafeDelete(baseView)
            Return False
        End If

        ' Изометрический вид — тонированный (как в образце PDF)
        Dim isoRect As ViewRect = GetIsoViewRect(doc, blockedRect)
        Dim isoScale As Double  = Math.Max(scale * 0.7, MIN_SCALE)
        Try
            Dim isoView As DrawingView = sheet.DrawingViews.AddBaseView(
                modelDoc,
                Pt2d(RectCenterX(isoRect), RectCenterY(isoRect)),
                isoScale,
                ViewOrientationTypeEnum.kIsoTopRightViewOrientation,
                DrawingViewStyleEnum.kShadedDrawingViewStyle)
        Catch ex As Exception
            Logger.Warn("Изометрический вид не добавлен: " & ex.Message)
        End Try

        ' Подписи видов
        ApplyViewLabel(doc, baseView, "Вид спереди")
        ApplyViewLabel(doc, topView,  "Вид сверху")
        ApplyViewLabel(doc, sideView, If(firstAngle, "Вид слева", "Вид справа"))

        Return True
    End Function

    ' ================================================================
    '  ВСПОМОГАТЕЛЬНЫЕ МЕТОДЫ ДЛЯ ВИДОВ
    ' ================================================================

    Private Function TryCreateBaseView(sheet As Sheet, modelDoc As Document, scale As Double) As DrawingView
        Try
            Return sheet.DrawingViews.AddBaseView(
                modelDoc,
                Pt2d(sheet.Width / 2.0, sheet.Height / 2.0),
                scale,
                ViewOrientationTypeEnum.kFrontViewOrientation,
                DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle)
        Catch ex As Exception
            Logger.Warn("AddBaseView failed scale=" & scale & ": " & ex.Message)
            Return Nothing
        End Try
    End Function

    Private Function TryAddProjectedView(sheet As Sheet, baseView As DrawingView, targetPt As Point2d) As DrawingView
        Try
            Return sheet.DrawingViews.AddProjectedView(
                baseView,
                targetPt,
                DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle)
        Catch ex As Exception
            Logger.Warn("AddProjectedView failed: " & ex.Message)
            Return Nothing
        End Try
    End Function

    ' Пересчёт масштаба чтобы вид влез в зону с запасом SCALE_FIT_MARGIN
    Private Function RecalculateScale(view As DrawingView, currentScale As Double, targetRect As ViewRect) As Double
        If view Is Nothing Then Return currentScale
        If view.Width <= 0.0 OrElse view.Height <= 0.0 Then Return currentScale

        Dim fitX As Double = currentScale * (RectWidth(targetRect)  / view.Width)  * SCALE_FIT_MARGIN
        Dim fitY As Double = currentScale * (RectHeight(targetRect) / view.Height) * SCALE_FIT_MARGIN
        Dim recalc As Double = Math.Min(fitX, fitY)

        Return If(recalc < currentScale, recalc, currentScale)
    End Function

    Private Function BuildTopTarget(baseView As DrawingView, topRect As ViewRect, firstAngle As Boolean) As Point2d
        Dim cx As Double = baseView.Center.X
        Dim cy As Double = RectCenterY(topRect)

        ' Первый угол: план снизу, третий — сверху
        If firstAngle Then
            If cy >= baseView.Center.Y Then cy = baseView.Center.Y - 0.01
        Else
            If cy <= baseView.Center.Y Then cy = baseView.Center.Y + 0.01
        End If
        Return Pt2d(cx, cy)
    End Function

    Private Function BuildSideTarget(baseView As DrawingView, sideRect As ViewRect) As Point2d
        Dim cx As Double = RectCenterX(sideRect)
        If cx <= baseView.Center.X Then cx = baseView.Center.X + 0.01
        Return Pt2d(cx, baseView.Center.Y)
    End Function

    Private Sub ApplyViewLabel(doc As DrawingDocument, view As DrawingView, caption As String)
        If view Is Nothing OrElse view.Label Is Nothing Then Return
        Try
            view.ShowLabel = True
            view.Label.FormattedText = caption & vbCrLf & view.Label.FormattedText
            Dim posX As Double = view.Left + MmToCm(doc, LABEL_MARGIN)
            Dim posY As Double = (view.Top - view.Height) - MmToCm(doc, LABEL_OFFSET)
            view.Label.Position = Pt2d(posX, posY)
        Catch : End Try
    End Sub

    Private Function GetProjectionStandard(doc As DrawingDocument) As Boolean
        Try
            Return doc.StylesManager.ActiveStandardStyle.FirstAngleProjection
        Catch
            Return True  ' По умолчанию первый угол (Европейский)
        End Try
    End Function

    Private Sub SafeDelete(view As DrawingView)
        If view IsNot Nothing Then
            Try
                view.Delete()
            Catch
            End Try
        End If
    End Sub

    ' ================================================================
    '  ЗОНЫ КОМПОНОВКИ (портированы из VBA GetFrontViewRectCm и др.)
    ' ================================================================

    ' Безопасная рабочая зона листа (внутри рамки)
    Private Function GetSafeRect(doc As DrawingDocument) As ViewRect
        Dim s As Sheet = doc.ActiveSheet
        Return New ViewRect(
            MmToCm(doc, FRAME_L_MM),
            s.Width - MmToCm(doc, FRAME_O_MM),
            MmToCm(doc, FRAME_O_MM),
            s.Height - MmToCm(doc, FRAME_O_MM))
    End Function

    ' Зона занятая штампом (справа снизу)
    Private Function GetTitleBlockRect(doc As DrawingDocument) As ViewRect
        Dim safe As ViewRect = GetSafeRect(doc)
        Return New ViewRect(
            safe.Right - MmToCm(doc, TB_W_MM),
            safe.Right,
            safe.Bottom,
            safe.Bottom + MmToCm(doc, TB_H_MM))
    End Function

    ' Зона фронтального вида
    Private Function GetFrontViewRect(doc As DrawingDocument, firstAngle As Boolean) As ViewRect
        Dim safe    As ViewRect = InsetRect(GetSafeRect(doc), MmToCm(doc, PAD_MM))
        Dim padCm   As Double   = MmToCm(doc, GAP_MM)
        Dim splitX  As Double   = safe.Right  - RectWidth(safe)  * SIDE_COL_RATIO
        Dim splitY  As Double   = safe.Top    - RectHeight(safe) * TOP_ROW_RATIO

        If firstAngle Then
            Return New ViewRect(safe.Left, splitX - padCm, splitY + padCm, safe.Top - padCm)
        Else
            Return New ViewRect(safe.Left, splitX - padCm, safe.Bottom, splitY - padCm)
        End If
    End Function

    ' Зона вида сверху
    Private Function GetTopViewRect(doc As DrawingDocument, firstAngle As Boolean) As ViewRect
        Dim safe    As ViewRect = InsetRect(GetSafeRect(doc), MmToCm(doc, PAD_MM))
        Dim padCm   As Double   = MmToCm(doc, GAP_MM)
        Dim splitX  As Double   = safe.Right  - RectWidth(safe)  * SIDE_COL_RATIO
        Dim splitY  As Double   = safe.Top    - RectHeight(safe) * TOP_ROW_RATIO

        If firstAngle Then
            Return New ViewRect(safe.Left, splitX - padCm, safe.Bottom, splitY - padCm)
        Else
            Return New ViewRect(safe.Left, splitX - padCm, splitY + padCm, safe.Top - padCm)
        End If
    End Function

    ' Зона бокового вида (справа от фронта)
    Private Function GetSideViewRect(doc As DrawingDocument, firstAngle As Boolean, frontRect As ViewRect) As ViewRect
        Dim safe    As ViewRect = InsetRect(GetSafeRect(doc), MmToCm(doc, PAD_MM))
        Dim padCm   As Double   = MmToCm(doc, GAP_MM)
        Dim splitX  As Double   = safe.Right - RectWidth(safe) * SIDE_COL_RATIO

        Dim yBot As Double = Math.Max(frontRect.Bottom, safe.Bottom)
        Dim yTop As Double = Math.Min(frontRect.Top,    safe.Top)

        Return New ViewRect(splitX + padCm, safe.Right - padCm, yBot, yTop)
    End Function

    ' Зона изометрического вида (справа, над штампом)
    Private Function GetIsoViewRect(doc As DrawingDocument, blockedRect As ViewRect) As ViewRect
        Dim safe   As ViewRect = InsetRect(GetSafeRect(doc), MmToCm(doc, PAD_MM))
        Return New ViewRect(
            blockedRect.Left,
            safe.Right - MmToCm(doc, PAD_MM),
            blockedRect.Top + MmToCm(doc, PAD_MM),
            safe.Top - MmToCm(doc, PAD_MM))
    End Function

    Private Function InsetRect(r As ViewRect, delta As Double) As ViewRect
        Return New ViewRect(r.Left + delta, r.Right - delta, r.Bottom + delta, r.Top - delta)
    End Function

    Private Function ViewFitsRect(view As DrawingView, rect As ViewRect) As Boolean
        If view Is Nothing Then Return False
        Dim vr As ViewRect = ViewToRect(view)
        Return vr.Left >= rect.Left AndAlso vr.Right <= rect.Right AndAlso
               vr.Bottom >= rect.Bottom AndAlso vr.Top <= rect.Top
    End Function

    Private Function ViewIntersectsRect(view As DrawingView, rect As ViewRect) As Boolean
        If view Is Nothing Then Return False
        Dim vr As ViewRect = ViewToRect(view)
        Return Not (vr.Right <= rect.Left OrElse rect.Right <= vr.Left OrElse
                    vr.Top <= rect.Bottom OrElse rect.Top <= vr.Bottom)
    End Function

    Private Function ViewToRect(view As DrawingView) As ViewRect
        Return New ViewRect(view.Left, view.Left + view.Width, view.Top - view.Height, view.Top)
    End Function

    Private Function RectWidth(r As ViewRect) As Double
        Return r.Right - r.Left
    End Function

    Private Function RectHeight(r As ViewRect) As Double
        Return r.Top - r.Bottom
    End Function

    Private Function RectCenterX(r As ViewRect) As Double
        Return r.Left + RectWidth(r) / 2.0
    End Function

    Private Function RectCenterY(r As ViewRect) As Double
        Return r.Bottom + RectHeight(r) / 2.0
    End Function

    ' ================================================================
    '  РАМКА СПДС А3
    ' ================================================================
    Private Function EnsureBorderDefinition(doc As DrawingDocument) As BorderDefinition
        Dim def As BorderDefinition = Nothing
        Try
            def = doc.BorderDefinitions.Item(BORDER_NAME)
        Catch
        End Try

        If def Is Nothing Then
            _app.SilentOperation = True
            def = doc.BorderDefinitions.Add(BORDER_NAME)
            _app.SilentOperation = False
        End If

        Dim sk As DrawingSketch = Nothing
        def.Edit(sk)
        Try
            ClearSketchLines(sk)

            ' Микро-якоря — фиксируют размер листа, Inventor не сместит рамку
            sk.SketchLines.AddByTwoPoints(Pt2d(0, 0), Pt2d(0.0001, 0.0001))
            sk.SketchLines.AddByTwoPoints(
                Pt2d(MmToCm(doc, A3_W_MM), MmToCm(doc, A3_H_MM)),
                Pt2d(MmToCm(doc, A3_W_MM) - 0.0001, MmToCm(doc, A3_H_MM) - 0.0001))

            ' Внутренняя рамка СПДС
            sk.SketchLines.AddAsTwoPointRectangle(
                Pt2d(MmToCm(doc, FRAME_L_MM), MmToCm(doc, FRAME_O_MM)),
                Pt2d(MmToCm(doc, A3_W_MM - FRAME_O_MM), MmToCm(doc, A3_H_MM - FRAME_O_MM)))
        Finally
            def.ExitEdit(True)
        End Try

        Return def
    End Function

    Private Sub ApplyBorder(sheet As Sheet, def As BorderDefinition)
        Try
            If sheet.Border IsNot Nothing Then sheet.Border.Delete()
        Catch : End Try
        _app.SilentOperation = True
        ' ИСПРАВЛЕНО: правильный метод AddBorder (не AddCustomBorder)
        sheet.AddBorder(def, Nothing)
        _app.SilentOperation = False
    End Sub

    ' ================================================================
    '  ШТАМП СПДС ФОРМА 3
    ' ================================================================
    Private Function EnsureTitleBlockDefinition(doc As DrawingDocument) As TitleBlockDefinition
        Dim def As TitleBlockDefinition = Nothing
        Try
            def = doc.TitleBlockDefinitions.Item(TITLEBLOCK_NAME)
        Catch
        End Try

        If def Is Nothing Then
            _app.SilentOperation = True
            def = doc.TitleBlockDefinitions.Add(TITLEBLOCK_NAME)
            _app.SilentOperation = False
        End If

        Dim sk As DrawingSketch = Nothing
        def.Edit(sk)
        Try
            ClearSketchFull(sk)
            DrawTitleBlockGeometry(doc, sk)
            AddTitleBlockLabels(doc, sk)
        Finally
            def.ExitEdit(True)
        End Try

        Return def
    End Function

    Private Sub DrawTitleBlockGeometry(doc As DrawingDocument, sk As DrawingSketch)
        ' Штамп в правом нижнем углу: x2 = правый край рамки, y1 = нижний + 5мм
        Dim x2 As Double = -MmToCm(doc, FRAME_O_MM)
        Dim y1 As Double =  MmToCm(doc, FRAME_O_MM)
        Dim x1 As Double = x2 - MmToCm(doc, TB_W_MM)
        Dim y2 As Double = y1 + MmToCm(doc, TB_H_MM)

        sk.SketchLines.AddByTwoPoints(Pt2d(0, 0), Pt2d(-0.0001, 0.0001))
        sk.SketchLines.AddAsTwoPointRectangle(Pt2d(x1, y1), Pt2d(x2, y2))

        ' Вертикальные разделители (колонки левой части: Изм, Кол.уч, Лист, №dok, Подп, Дата)
        DrawVLine(doc, sk, x1, y1,   7,  0, 55)
        DrawVLine(doc, sk, x1, y1,  17,  0, 55)
        DrawVLine(doc, sk, x1, y1,  27,  0, 55)
        DrawVLine(doc, sk, x1, y1,  42,  0, 55)
        DrawVLine(doc, sk, x1, y1,  57,  0, 55)
        DrawVLine(doc, sk, x1, y1,  67,  0, 55)

        ' Вертикальные разделители правой части: Стадия, Лист, Листов
        DrawVLine(doc, sk, x1, y1, 137,  0, 40)
        DrawVLine(doc, sk, x1, y1, 152, 15, 40)
        DrawVLine(doc, sk, x1, y1, 167, 15, 40)

        ' Горизонтальные строки левой части
        Dim y As Double
        For y = 5.0 To 30.0 Step 5.0
            DrawHLine(doc, sk, x1, y1,   0,  67, y)
        Next
        DrawHLine(doc, sk, x1, y1,   0, 185, 15)
        DrawHLine(doc, sk, x1, y1,   0,  67, 35)
        DrawHLine(doc, sk, x1, y1, 137, 185, 35)
        DrawHLine(doc, sk, x1, y1,   0, 185, 40)
        DrawHLine(doc, sk, x1, y1,   0,  67, 45)
        DrawHLine(doc, sk, x1, y1,   0,  67, 50)
    End Sub

    Private Sub AddTitleBlockLabels(doc As DrawingDocument, sk As DrawingSketch)
        Dim x2 As Double = -MmToCm(doc, FRAME_O_MM)
        Dim y1 As Double =  MmToCm(doc, FRAME_O_MM)
        Dim x1 As Double = x2 - MmToCm(doc, TB_W_MM)

        ' Заголовки столбцов (статические)
        AddLabel(doc, sk, x1, y1,   0, 35,   7, 40, "Изм.")
        AddLabel(doc, sk, x1, y1,   7, 35,  17, 40, "Кол.уч")
        AddLabel(doc, sk, x1, y1,  17, 35,  27, 40, "Лист")
        AddLabel(doc, sk, x1, y1,  27, 35,  42, 40, "№ dok.")
        AddLabel(doc, sk, x1, y1,  42, 35,  57, 40, "Подп.")
        AddLabel(doc, sk, x1, y1,  57, 35,  67, 40, "Дата")
        AddLabel(doc, sk, x1, y1, 137, 35, 152, 40, "Стадия")
        AddLabel(doc, sk, x1, y1, 152, 35, 167, 40, "Лист")
        AddLabel(doc, sk, x1, y1, 167, 35, 185, 40, "Листов")

        ' Промт-поля (заполняются данными из Excel при добавлении штампа на лист)
        AddPrompt(doc, sk, x1, y1,  67, 40, 185, 55, "CODE")
        AddPrompt(doc, sk, x1, y1,  67, 15, 137, 40, "PROJECT_NAME")
        AddPrompt(doc, sk, x1, y1,  67,  0, 137, 15, "DRAWING_NAME")
        AddPrompt(doc, sk, x1, y1, 137,  0, 185, 15, "ORG_NAME")
        AddPrompt(doc, sk, x1, y1, 137, 15, 152, 35, "STAGE")
        AddPrompt(doc, sk, x1, y1, 152, 15, 167, 35, "SHEET")
        AddPrompt(doc, sk, x1, y1, 167, 15, 185, 35, "SHEETS")
    End Sub

    Private Sub ApplyTitleBlock(sheet As Sheet, def As TitleBlockDefinition, prompts As Dictionary(Of String, String))
        Try
            If sheet.TitleBlock IsNot Nothing Then sheet.TitleBlock.Delete()
        Catch : End Try

        ' ИСПРАВЛЕНО: массив 1-based, 7 промт-полей (CODE, PROJECT_NAME, DRAWING_NAME, ORG_NAME, STAGE, SHEET, SHEETS)
        Dim order As String() = {"CODE", "PROJECT_NAME", "DRAWING_NAME", "ORG_NAME", "STAGE", "SHEET", "SHEETS"}
        Dim ps(8) As String   ' индексы 1..7, 0 и 8 пусты
        For i As Integer = 0 To order.Length - 1
            Dim v As String = String.Empty
            If prompts IsNot Nothing Then prompts.TryGetValue(order(i), v)
            ps(i + 1) = If(String.IsNullOrEmpty(v), String.Empty, v)
        Next

        Try
            _app.SilentOperation = True
            ' ИСПРАВЛЕНО: передаём Nothing как второй аргумент (точка привязки — Nothing означает дефолт)
            sheet.AddTitleBlock(def, Nothing, ps)
        Catch ex As Exception
            Logger.Warn("AddTitleBlock failed sheet=" & sheet.Name & ": " & ex.Message)
        Finally
            _app.SilentOperation = False
        End Try
    End Sub

    ' ================================================================
    '  ГЕОМЕТРИЯ ШТАМПА — вспомогательные
    ' ================================================================
    Private Sub DrawVLine(doc As DrawingDocument, sk As DrawingSketch,
                          x0 As Double, y0 As Double,
                          atMm As Double, yFromMm As Double, yToMm As Double)
        sk.SketchLines.AddByTwoPoints(
            Pt2d(x0 + MmToCm(doc, atMm), y0 + MmToCm(doc, yFromMm)),
            Pt2d(x0 + MmToCm(doc, atMm), y0 + MmToCm(doc, yToMm)))
    End Sub

    Private Sub DrawHLine(doc As DrawingDocument, sk As DrawingSketch,
                          x0 As Double, y0 As Double,
                          xFromMm As Double, xToMm As Double, atMm As Double)
        sk.SketchLines.AddByTwoPoints(
            Pt2d(x0 + MmToCm(doc, xFromMm), y0 + MmToCm(doc, atMm)),
            Pt2d(x0 + MmToCm(doc, xToMm),   y0 + MmToCm(doc, atMm)))
    End Sub

    Private Sub AddLabel(doc As DrawingDocument, sk As DrawingSketch,
                         x0 As Double, y0 As Double,
                         lMm As Double, bMm As Double, rMm As Double, tMm As Double,
                         text As String)
        Dim tb As Inventor.TextBox = sk.TextBoxes.AddByRectangle(
            Pt2d(x0 + MmToCm(doc, lMm), y0 + MmToCm(doc, bMm)),
            Pt2d(x0 + MmToCm(doc, rMm), y0 + MmToCm(doc, tMm)),
            text)
        tb.HorizontalJustification = HorizontalTextAlignmentEnum.kAlignTextCenter
        tb.VerticalJustification   = VerticalTextAlignmentEnum.kAlignTextMiddle
    End Sub

    Private Sub AddPrompt(doc As DrawingDocument, sk As DrawingSketch,
                          x0 As Double, y0 As Double,
                          lMm As Double, bMm As Double, rMm As Double, tMm As Double,
                          promptName As String)
        Dim tb As Inventor.TextBox = sk.TextBoxes.AddByRectangle(
            Pt2d(x0 + MmToCm(doc, lMm), y0 + MmToCm(doc, bMm)),
            Pt2d(x0 + MmToCm(doc, rMm), y0 + MmToCm(doc, tMm)),
            "<Prompt>" & promptName & "</Prompt>")
        tb.HorizontalJustification = HorizontalTextAlignmentEnum.kAlignTextCenter
        tb.VerticalJustification   = VerticalTextAlignmentEnum.kAlignTextMiddle
    End Sub

    Private Sub ClearSketchLines(sk As DrawingSketch)
        If sk Is Nothing Then Return
        For i As Integer = sk.SketchLines.Count To 1 Step -1
            Try
                sk.SketchLines.Item(i).Delete()
            Catch
            End Try
        Next
    End Sub

    Private Sub ClearSketchFull(sk As DrawingSketch)
        If sk Is Nothing Then Return
        For i As Integer = sk.TextBoxes.Count To 1 Step -1
            Try
                sk.TextBoxes.Item(i).Delete()
            Catch
            End Try
        Next
        ClearSketchLines(sk)
    End Sub

    ' ================================================================
    '  ЛИСТЫ
    ' ================================================================
    Private Function EnsureSheet(doc As DrawingDocument, name As String) As Sheet
        For Each s As Sheet In doc.Sheets
            If String.Equals(s.Name.Split(":"c)(0), name, StringComparison.OrdinalIgnoreCase) Then
                Return s
            End If
        Next
        Dim newSheet As Sheet = doc.Sheets.Add(
            DrawingSheetSizeEnum.kA3DrawingSheetSize,
            PageOrientationTypeEnum.kLandscapePageOrientation)
        newSheet.Name = name
        Return newSheet
    End Function

    Private Sub RemoveAllViews(sheet As Sheet)
        If sheet Is Nothing Then Return
        For i As Integer = sheet.DrawingViews.Count To 1 Step -1
            Try
                sheet.DrawingViews.Item(i).Delete()
            Catch
            End Try
        Next
    End Sub

    ' ================================================================
    '  ЕДИНИЦЫ ИЗМЕРЕНИЯ
    ' ================================================================
    Private Function MmToCm(doc As DrawingDocument, mm As Double) As Double
        If doc Is Nothing Then Return mm * 0.1
        Return doc.UnitsOfMeasure.ConvertUnits(mm,
            UnitsTypeEnum.kMillimeterLengthUnits,
            UnitsTypeEnum.kCentimeterLengthUnits)
    End Function

    Private Function Pt2d(x As Double, y As Double) As Point2d
        Return _app.TransientGeometry.CreatePoint2d(x, y)
    End Function

    ' ================================================================
    '  DATA TRANSFER OBJECTS
    ' ================================================================
    Public Class AlbumItem
        Public Property ModelPath As String
        Public Property Prompts   As Dictionary(Of String, String)
        Public Sub New()
            Prompts = New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        End Sub
    End Class

    Public Class ViewRect
        Public ReadOnly Left   As Double
        Public ReadOnly Right  As Double
        Public ReadOnly Bottom As Double
        Public ReadOnly Top    As Double
        Public Sub New(l As Double, r As Double, b As Double, t As Double)
            Left = l : Right = r : Bottom = b : Top = t
        End Sub
    End Class

    ' ================================================================
    '  ЛОГГЕР
    ' ================================================================
    Private NotInheritable Class Logger
        Public Shared Sub Info(msg As String)
            Debug.Print("INFO:  " & msg)
        End Sub
        Public Shared Sub Warn(msg As String)
            Debug.Print("WARN:  " & msg)
        End Sub
        Public Shared Sub [Error](msg As String)
            Debug.Print("ERROR: " & msg)
        End Sub
    End Class

End Class

' ================================================================
'  ЗАГРУЗЧИК EXCEL (портирован из RKM_Excel.bas vba-inventor)
' ================================================================
Public NotInheritable Class ExcelLoader

    Private Const DEFAULT_SHEET As String = "ALBUM"

    ' ================================================================
    '  Чтение .xlsx через Reflection + ZipArchive (без прямой ссылки на сборку)
    '  ZipFile/ZipArchive загружаются через Assembly.Load во время выполнения
    ' ================================================================
    Public Shared Function Load(
            excelPath As String,
            workspacePath As String,
            sheetTabName As String) As List(Of StoneAlbumRule.AlbumItem)

        Dim result As New List(Of StoneAlbumRule.AlbumItem)()
        Try
            ' Загружаем System.IO.Compression и System.IO.Compression.FileSystem через Reflection
            Dim asmComp     As System.Reflection.Assembly = Nothing
            Dim asmCompFS   As System.Reflection.Assembly = Nothing
            Try
                asmComp   = System.Reflection.Assembly.Load("System.IO.Compression, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
            Catch
                Try
                    asmComp = System.Reflection.Assembly.LoadWithPartialName("System.IO.Compression")
                Catch
                End Try
            End Try
            Try
                asmCompFS = System.Reflection.Assembly.Load("System.IO.Compression.FileSystem, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
            Catch
                Try
                    asmCompFS = System.Reflection.Assembly.LoadWithPartialName("System.IO.Compression.FileSystem")
                Catch
                End Try
            End Try

            If asmComp Is Nothing Then
                Throw New Exception("Не удалось загрузить System.IO.Compression." & vbCrLf &
                                    "Проверьте версию .NET Framework в Inventor.")
            End If

            ' Получаем тип ZipFile (он в FileSystem сборке, или в основной)
            Dim zipFileType As System.Type = Nothing
            If asmCompFS IsNot Nothing Then zipFileType = asmCompFS.GetType("System.IO.Compression.ZipFile")
            If zipFileType Is Nothing Then zipFileType = asmComp.GetType("System.IO.Compression.ZipFile")
            If zipFileType Is Nothing Then
                ' .NET 4.5+ ZipFile может быть в основной сборке
                zipFileType = System.Type.GetType("System.IO.Compression.ZipFile, System.IO.Compression.FileSystem")
            End If
            If zipFileType Is Nothing Then
                Throw New Exception("Тип ZipFile не найден в загруженных сборках.")
            End If

            ' ZipFile.OpenRead(path) → ZipArchive
            Dim zipArchive As Object = Nothing
            Try
                zipArchive = zipFileType.InvokeMember(
                    "OpenRead",
                    System.Reflection.BindingFlags.Static Or System.Reflection.BindingFlags.Public Or System.Reflection.BindingFlags.InvokeMethod,
                    Nothing, Nothing, New Object() {excelPath})
            Catch invEx As System.Reflection.TargetInvocationException
                Dim inner As Exception = If(invEx.InnerException, CType(invEx, Exception))
                Throw New Exception("ZipFile.OpenRead не удался: " & inner.Message & " [" & inner.GetType().Name & "]")
            End Try

            If zipArchive Is Nothing Then
                Throw New Exception("ZipFile.OpenRead вернул Nothing для: " & excelPath)
            End If

            Try
                Dim sharedStrings As List(Of String) = ReadSharedStrings(zipArchive)
                Dim sheetXmlName  As String           = FindSheetXmlName(zipArchive, sheetTabName)
                If String.IsNullOrEmpty(sheetXmlName) Then
                    Throw New Exception("Лист '" & sheetTabName & "' не найден в " & excelPath)
                End If

                Dim rows As List(Of List(Of String)) = ReadSheet(zipArchive, sheetXmlName, sharedStrings)
                If rows.Count < 2 Then
                    Throw New Exception("Лист '" & sheetTabName & "' пустой или содержит только заголовок.")
                End If

                Dim headerRowIdx As Integer = DetectHeaderRow(rows)
                Dim headerMap    As Dictionary(Of String, Integer) = BuildHeaderMap(rows(headerRowIdx))

                If Not headerMap.ContainsKey("MODEL_PATH") Then
                    Dim found As New System.Text.StringBuilder()
                    For Each h As String In rows(headerRowIdx)
                        If Not String.IsNullOrWhiteSpace(h) Then found.Append("[" & h & "] ")
                    Next
                    Throw New Exception("Колонка MODEL_PATH не найдена." & vbCrLf &
                        "Допустимые: MODEL_PATH, MODEL, ПУТЬ, ФАЙЛ, МОДЕЛЬ" & vbCrLf &
                        "Найдено: " & found.ToString())
                End If

                Dim modelCol   As Integer   = headerMap("MODEL_PATH")
                Dim promptKeys As String()  = {"CODE","PROJECT_NAME","DRAWING_NAME","ORG_NAME","STAGE","SHEET","SHEETS"}

                For r As Integer = headerRowIdx + 1 To rows.Count - 1
                    Dim row As List(Of String) = rows(r)
                    If row.Count <= modelCol Then Continue For
                    Dim rawPath As String = If(row(modelCol) IsNot Nothing, row(modelCol).Trim(), "")
                    If String.IsNullOrWhiteSpace(rawPath) Then Continue For

                    Dim resolvedPath As String = ResolvePath(rawPath, workspacePath, excelPath)
                    If String.IsNullOrWhiteSpace(resolvedPath) Then resolvedPath = rawPath

                    Dim item As New StoneAlbumRule.AlbumItem()
                    item.ModelPath = resolvedPath
                    For Each key As String In promptKeys
                        If headerMap.ContainsKey(key) Then
                            Dim col As Integer = headerMap(key)
                            If col < row.Count AndAlso row(col) IsNot Nothing Then
                                item.Prompts(key) = row(col).Trim()
                            End If
                        End If
                    Next
                    result.Add(item)
                Next
            Finally
                ' Закрываем ZipArchive (Dispose)
                Try
                    Dim dispMethod As System.Reflection.MethodInfo = zipArchive.GetType().GetMethod("Dispose")
                    If dispMethod IsNot Nothing Then dispMethod.Invoke(zipArchive, Nothing)
                Catch
                End Try
            End Try

        Catch ex As Exception
            ' Разворачиваем InnerException от Reflection
            Dim realEx As Exception = ex
            Do While realEx.InnerException IsNot Nothing
                realEx = realEx.InnerException
            Loop
            Dim msg As String = realEx.Message
            If TypeOf realEx Is System.IO.FileNotFoundException Then
                msg = "Файл не найден: " & realEx.Message
            ElseIf TypeOf realEx Is System.IO.InvalidDataException Then
                msg = "Файл повреждён или не является .xlsx: " & realEx.Message
            End If
            System.Windows.Forms.MessageBox.Show(
                "Ошибка чтения Excel:" & vbCrLf & msg & vbCrLf & vbCrLf &
                "[" & realEx.GetType().Name & "]",
                "Ошибка ExcelLoader",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Error)
        End Try

        Return result
    End Function

    ' ── Получить содержимое файла внутри ZIP по имени ──
    Private Shared Function GetZipEntryText(zipArchive As Object, entryName As String) As String
        ' zipArchive.GetEntry(entryName) → ZipArchiveEntry
        Dim getEntry As System.Reflection.MethodInfo = zipArchive.GetType().GetMethod("GetEntry")
        Dim entry As Object = getEntry.Invoke(zipArchive, New Object() {entryName})
        If entry Is Nothing Then Return Nothing

        ' entry.Open() → Stream
        Dim openMethod As System.Reflection.MethodInfo = entry.GetType().GetMethod("Open")
        Dim stream As System.IO.Stream = CType(openMethod.Invoke(entry, Nothing), System.IO.Stream)
        Using sr As New System.IO.StreamReader(stream, System.Text.Encoding.UTF8)
            Return sr.ReadToEnd()
        End Using
    End Function

    ' ── Найти имя файла листа внутри ZIP ──
    Private Shared Function FindSheetXmlName(zipArchive As Object, sheetTabName As String) As String
        Dim wbXml As String = GetZipEntryText(zipArchive, "xl/workbook.xml")
        If wbXml Is Nothing Then Return "xl/worksheets/sheet1.xml"

        Dim rId As String    = ""
        Dim firstRid As String = ""
        Dim mc As System.Text.RegularExpressions.MatchCollection =
            System.Text.RegularExpressions.Regex.Matches(wbXml,
                "<sheet[^>]+name=""([^""]+)""[^>]+r:id=""([^""]+)""",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        For Each m As System.Text.RegularExpressions.Match In mc
            If firstRid = "" Then firstRid = m.Groups(2).Value
            If String.Equals(m.Groups(1).Value, sheetTabName, StringComparison.OrdinalIgnoreCase) Then
                rId = m.Groups(2).Value : Exit For
            End If
        Next
        If rId = "" Then rId = firstRid
        If rId = "" Then Return "xl/worksheets/sheet1.xml"

        Dim relsXml As String = GetZipEntryText(zipArchive, "xl/_rels/workbook.xml.rels")
        If relsXml IsNot Nothing Then
            Dim rm As System.Text.RegularExpressions.Match =
                System.Text.RegularExpressions.Regex.Match(relsXml, "Id=""" & rId & """[^>]+Target=""([^""]+)""")
            If rm.Success Then
                Dim target As String = rm.Groups(1).Value
                If Not target.StartsWith("xl/") Then target = "xl/" & target
                Return target
            End If
        End If
        Return "xl/worksheets/sheet1.xml"
    End Function

    ' ── Читаем sharedStrings.xml ──
    Private Shared Function ReadSharedStrings(zipArchive As Object) As List(Of String)
        Dim result As New List(Of String)()
        Dim xml As String = GetZipEntryText(zipArchive, "xl/sharedStrings.xml")
        If xml Is Nothing Then Return result

        Dim siMatches As System.Text.RegularExpressions.MatchCollection =
            System.Text.RegularExpressions.Regex.Matches(xml, "<si>(.*?)</si>",
                System.Text.RegularExpressions.RegexOptions.Singleline)
        For Each m As System.Text.RegularExpressions.Match In siMatches
            Dim tMatches As System.Text.RegularExpressions.MatchCollection =
                System.Text.RegularExpressions.Regex.Matches(m.Groups(1).Value, "<t(?:[^>]*)>(.*?)</t>",
                    System.Text.RegularExpressions.RegexOptions.Singleline)
            Dim sb As New System.Text.StringBuilder()
            For Each tm As System.Text.RegularExpressions.Match In tMatches
                sb.Append(XmlDecode(tm.Groups(1).Value))
            Next
            result.Add(sb.ToString())
        Next
        Return result
    End Function

    ' ── Читаем лист sheet.xml → List(Of List(Of String)) ──
    Private Shared Function ReadSheet(zipArchive As Object, sheetPath As String,
                                      sharedStrings As List(Of String)) As List(Of List(Of String))
        Dim result As New List(Of List(Of String))()
        Dim xml As String = GetZipEntryText(zipArchive, sheetPath)
        If xml Is Nothing Then Return result

        Dim rowMatches As System.Text.RegularExpressions.MatchCollection =
            System.Text.RegularExpressions.Regex.Matches(xml, "<row[^>]*>(.*?)</row>",
                System.Text.RegularExpressions.RegexOptions.Singleline)
        For Each rowM As System.Text.RegularExpressions.Match In rowMatches
            Dim cellMatches As System.Text.RegularExpressions.MatchCollection =
                System.Text.RegularExpressions.Regex.Matches(rowM.Groups(1).Value,
                    "<c\s+r=""([A-Z]+)\d+""([^>]*)>(.*?)</c>",
                    System.Text.RegularExpressions.RegexOptions.Singleline)

            Dim maxCol As Integer = -1
            Dim cellData As New Dictionary(Of Integer, String)()
            For Each cellM As System.Text.RegularExpressions.Match In cellMatches
                Dim colIdx  As Integer = ColLetterToIndex(cellM.Groups(1).Value)
                Dim attrs   As String  = cellM.Groups(2).Value
                Dim inner   As String  = cellM.Groups(3).Value
                Dim cellType As String = ""
                Dim typeM As System.Text.RegularExpressions.Match =
                    System.Text.RegularExpressions.Regex.Match(attrs, "\bt=""([^""]+)""")
                If typeM.Success Then cellType = typeM.Groups(1).Value

                Dim cellVal As String = ""
                Dim vM As System.Text.RegularExpressions.Match =
                    System.Text.RegularExpressions.Regex.Match(inner, "<v>(.*?)</v>",
                        System.Text.RegularExpressions.RegexOptions.Singleline)
                If vM.Success Then
                    Dim raw As String = XmlDecode(vM.Groups(1).Value)
                    If cellType = "s" Then
                        Dim idx As Integer = 0
                        If Integer.TryParse(raw, idx) AndAlso idx < sharedStrings.Count Then
                            cellVal = sharedStrings(idx)
                        End If
                    Else
                        cellVal = raw
                    End If
                End If
                Dim isM As System.Text.RegularExpressions.Match =
                    System.Text.RegularExpressions.Regex.Match(inner, "<is>.*?<t>(.*?)</t>.*?</is>",
                        System.Text.RegularExpressions.RegexOptions.Singleline)
                If isM.Success Then cellVal = XmlDecode(isM.Groups(1).Value)

                If colIdx > maxCol Then maxCol = colIdx
                cellData(colIdx) = cellVal
            Next

            If maxCol >= 0 Then
                Dim rowList As New List(Of String)()
                For ci As Integer = 0 To maxCol
                    rowList.Add(If(cellData.ContainsKey(ci), cellData(ci), ""))
                Next
                result.Add(rowList)
            End If
        Next
        Return result
    End Function

    ' ── Определить строку заголовка ──
    Private Shared Function DetectHeaderRow(rows As List(Of List(Of String))) As Integer
        For r As Integer = 0 To Math.Min(19, rows.Count - 1)
            If BuildHeaderMap(rows(r)).ContainsKey("MODEL_PATH") Then Return r
        Next
        Dim bestRow As Integer = 0 : Dim bestCnt As Integer = 0
        For r As Integer = 0 To Math.Min(19, rows.Count - 1)
            Dim cnt As Integer = 0
            For Each v As String In rows(r)
                If Not String.IsNullOrWhiteSpace(v) Then cnt += 1
            Next
            If cnt > bestCnt Then bestCnt = cnt : bestRow = r
        Next
        Return bestRow
    End Function

    Private Shared Function BuildHeaderMap(headerRow As List(Of String)) As Dictionary(Of String, Integer)
        Dim map As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
        For c As Integer = 0 To headerRow.Count - 1
            Dim canonical As String = NormalizeAndAlias(headerRow(c))
            If Not String.IsNullOrEmpty(canonical) AndAlso Not map.ContainsKey(canonical) Then
                map(canonical) = c
            End If
        Next
        Return map
    End Function

    Private Shared Function NormalizeAndAlias(raw As String) As String
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
        Dim t As String = Transliterate(n)
        t = System.Text.RegularExpressions.Regex.Replace(t, "[^A-Z0-9]+", "_").Trim("_"c)
        Select Case t
            Case "MODEL_PATH","MODEL","P"      : Return "MODEL_PATH"
            Case "CODE","SHIFR"                : Return "CODE"
            Case "PROJECT_NAME","OBEKT"        : Return "PROJECT_NAME"
            Case "DRAWING_NAME","NAIMENOVANIE" : Return "DRAWING_NAME"
            Case "ORG_NAME","ORGANIZACIYA"     : Return "ORG_NAME"
            Case "STAGE","STADIYA"             : Return "STAGE"
            Case "SHEET","LIST"                : Return "SHEET"
            Case "SHEETS","LISTOV"             : Return "SHEETS"
        End Select
        Return t
    End Function

    Private Shared Function Transliterate(s As String) As String
        Dim src As String() = {"А","Б","В","Г","Д","Е","Ё","Ж","З","И","Й","К","Л","М","Н","О","П","Р","С","Т","У","Ф","Х","Ц","Ч","Ш","Щ","Ъ","Ы","Ь","Э","Ю","Я"}
        Dim dst As String() = {"A","B","V","G","D","E","E","ZH","Z","I","Y","K","L","M","N","O","P","R","S","T","U","F","H","C","CH","SH","SCH","","Y","","E","YU","YA"}
        Dim r As String = s
        For i As Integer = 0 To src.Length - 1 : r = r.Replace(src(i), dst(i)) : Next
        Return r
    End Function

    Private Shared Function ResolvePath(input As String, workspace As String, excelFile As String) As String
        If System.IO.File.Exists(input) Then Return System.IO.Path.GetFullPath(input)
        If Not String.IsNullOrWhiteSpace(workspace) Then
            Dim c As String = System.IO.Path.Combine(workspace, input)
            If System.IO.File.Exists(c) Then Return System.IO.Path.GetFullPath(c)
        End If
        Dim excelDir As String = System.IO.Path.GetDirectoryName(excelFile)
        If Not String.IsNullOrWhiteSpace(excelDir) Then
            Dim c As String = System.IO.Path.Combine(excelDir, input)
            If System.IO.File.Exists(c) Then Return System.IO.Path.GetFullPath(c)
        End If
        Dim fileName As String = System.IO.Path.GetFileName(input)
        If Not fileName.EndsWith(".ipt", StringComparison.OrdinalIgnoreCase) Then fileName &= ".ipt"
        For Each root As String In {workspace, excelDir}
            If String.IsNullOrWhiteSpace(root) Then Continue For
            Dim direct As String = System.IO.Path.Combine(root, fileName)
            If System.IO.File.Exists(direct) Then Return System.IO.Path.GetFullPath(direct)
            Try
                For Each subDir As String In System.IO.Directory.GetDirectories(root)
                    Dim sub1 As String = System.IO.Path.Combine(subDir, fileName)
                    If System.IO.File.Exists(sub1) Then Return System.IO.Path.GetFullPath(sub1)
                Next
            Catch
            End Try
        Next
        Return String.Empty
    End Function

    Private Shared Function ColLetterToIndex(col As String) As Integer
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

    Private Shared Function SafeCell(value As Object) As String
        If value Is Nothing Then Return String.Empty
        Return Convert.ToString(value).Trim()
    End Function

End Class
