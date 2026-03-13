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
' Imports System.IO  -- убрано: конфликт Path/File с Inventor (используем полные пути System.IO.Path / System.IO.File)

' ================================================================
'  ТОЧКА ВХОДА
' ================================================================
Sub Main()
    ' Пути: сначала читаем кастомные свойства документа (если заданы),
    ' иначе спрашиваем через InputBox
    Dim excelPath As String     = String.Empty
    Dim workspacePath As String = String.Empty
    Dim sheetTabName As String  = "ALBUM"

    Try : excelPath     = iProperties.Value("Custom", "AlbumExcel")     : Catch : End Try
    Try : workspacePath = iProperties.Value("Custom", "AlbumWorkspace") : Catch : End Try
    Try : sheetTabName  = iProperties.Value("Custom", "AlbumSheet")     : Catch : End Try

    ' Fallback: запрос пути через InputBox если свойство не задано
    If String.IsNullOrWhiteSpace(excelPath) Then
        excelPath = InputBox("Укажите полный путь к Excel-файлу альбома (.xlsx):", "Путь к Excel", "")
        If String.IsNullOrWhiteSpace(excelPath) Then Return
    End If

    If String.IsNullOrWhiteSpace(workspacePath) Then
        workspacePath = System.IO.Path.GetDirectoryName(excelPath)
    End If

    Dim doc As DrawingDocument = TryCast(ThisApplication.ActiveDocument, DrawingDocument)
    If doc Is Nothing Then
        System.Windows.Forms.MessageBox.Show("Откройте .idw документ перед запуском правила.", "Ошибка", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
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
            System.Windows.Forms.MessageBox.Show("Excel не содержит строк с MODEL_PATH.", "Ошибка", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning)
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

        System.Windows.Forms.MessageBox.Show("Альбом собран: " & items.Count & " листов.", "Готово", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information)
    End Sub

    ' ================================================================
    '  ОДИН ЛИСТ
    ' ================================================================
    Private Sub BuildSingleSheet(
            doc As DrawingDocument,
            item As AlbumItem,
            borderDef As BorderDefinition,
            titleDef As TitleBlockDefinition)

        If Not System.IO.File.Exists(item.ModelPath) Then
            Logger.Warn("Файл модели не найден: " & item.ModelPath)
            Return
        End If

        Dim modelDoc As Document = Nothing
        Try
            ' Лист
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

            ' Рамка
            ApplyBorder(sheet, borderDef)

            ' Штамп с данными из Excel
            ApplyTitleBlock(sheet, titleDef, item.Prompts)

            ' Открыть 3D модель
            modelDoc = _app.Documents.Open(item.ModelPath, False)

            ' Расставить виды (алгоритм из VBA PR #51)
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
    Private Const HEADER_SCAN   As Integer = 10

    Public Shared Function Load(
            excelPath As String,
            workspacePath As String,
            sheetTabName As String) As List(Of StoneAlbumRule.AlbumItem)

        Dim result As New List(Of StoneAlbumRule.AlbumItem)()

        Dim xlApp  As Object = Nothing
        Dim xlBook As Object = Nothing

        Try
            xlApp = System.Activator.CreateInstance(System.Type.GetTypeFromProgID("Excel.Application"))
            xlApp.Visible      = False
            xlApp.DisplayAlerts = False

            xlBook = xlApp.Workbooks.Open(excelPath)

            ' Ищем лист ALBUM или берём первый
            Dim xlSheet As Object = ResolveSheet(xlBook, sheetTabName)
            If xlSheet Is Nothing Then
                Throw New Exception("Лист '" & sheetTabName & "' не найден в " & excelPath)
            End If

            Dim headerRow As Integer = DetectHeaderRow(xlSheet)
            Dim headerMap As Dictionary(Of String, Integer) = ReadHeaders(xlSheet, headerRow)

            If Not headerMap.ContainsKey("MODEL_PATH") Then
                Throw New Exception("Колонка MODEL_PATH (или алиас) не найдена.")
            End If

            Dim modelCol As Integer = headerMap("MODEL_PATH")
            Dim lastRow  As Integer = CInt(xlSheet.Cells(xlSheet.Rows.Count, modelCol).End(-4162).Row)

            For row As Integer = headerRow + 1 To lastRow
                Dim rawPath As String = SafeCell(xlSheet.Cells(row, modelCol).Value)
                If String.IsNullOrWhiteSpace(rawPath) Then Continue For

                Dim resolvedPath As String = ResolvePath(rawPath, workspacePath, excelPath)
                If String.IsNullOrWhiteSpace(resolvedPath) Then
                    Debug.Print("WARN:  MODEL_PATH не найден, строка пропущена: " & rawPath)
                    Continue For
                End If

                Dim item As New StoneAlbumRule.AlbumItem()
                item.ModelPath = resolvedPath

                ' Заполнение полей штампа из колонок Excel
                For Each key As String In {"CODE", "PROJECT_NAME", "DRAWING_NAME", "ORG_NAME", "STAGE", "SHEET", "SHEETS"}
                    If headerMap.ContainsKey(key) Then
                        item.Prompts(key) = SafeCell(xlSheet.Cells(row, headerMap(key)).Value)
                    End If
                Next

                result.Add(item)
            Next

        Catch ex As Exception
            Debug.Print("ERROR: Excel load failed: " & ex.Message)
        Finally
            Try
                If xlBook IsNot Nothing Then xlBook.Close(False)
            Catch
            End Try
            Try
                If xlApp  IsNot Nothing Then xlApp.Quit()
            Catch
            End Try
        End Try

        Return result
    End Function

    Private Shared Function ResolveSheet(xlBook As Object, name As String) As Object
        If Not String.IsNullOrWhiteSpace(name) Then
            For Each ws As Object In xlBook.Worksheets
                If String.Equals(CStr(ws.Name), name, StringComparison.OrdinalIgnoreCase) Then Return ws
            Next
        End If
        Return If(xlBook.Worksheets.Count > 0, xlBook.Worksheets(1), Nothing)
    End Function

    Private Shared Function DetectHeaderRow(xlSheet As Object) As Integer
        For r As Integer = 1 To HEADER_SCAN
            Dim map As Dictionary(Of String, Integer) = ReadHeaders(xlSheet, r)
            If map.ContainsKey("MODEL_PATH") Then Return r
        Next
        Return 1
    End Function

    ' Чтение заголовков с полной транслитерацией и алиасами (как в VBA RKM_Excel.bas)
    Private Shared Function ReadHeaders(xlSheet As Object, headerRow As Integer) As Dictionary(Of String, Integer)
        Dim map As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
        Dim lastCol As Integer = CInt(xlSheet.Cells(headerRow, xlSheet.Columns.Count).End(-4159).Column)

        For c As Integer = 1 To lastCol
            Dim raw As String = SafeCell(xlSheet.Cells(headerRow, c).Value)
            Dim canonical As String = NormalizeAndAlias(raw)
            If Not String.IsNullOrEmpty(canonical) AndAlso Not map.ContainsKey(canonical) Then
                map(canonical) = c
            End If
        Next

        map("__HEADER_ROW") = headerRow
        Return map
    End Function

    Private Shared Function NormalizeAndAlias(raw As String) As String
        ' 1) Алиас по точному совпадению (с поддержкой кириллицы)
        Dim n As String = raw.Trim().ToUpperInvariant()
        Select Case n
            Case "MODEL_PATH", "MODEL", "P", "ПУТЬ", "ФАЙЛ", "МОДЕЛЬ"    : Return "MODEL_PATH"
            Case "CODE", "ШИФР", "АРТИКУЛ", "ОБОЗНАЧЕНИЕ"                : Return "CODE"
            Case "PROJECT_NAME", "PROJECT", "ОБЪЕКТ", "ПРОЕКТ"           : Return "PROJECT_NAME"
            Case "DRAWING_NAME", "TITLE", "НАИМЕНОВАНИЕ", "ИМЯ ЧЕРТЕЖА" : Return "DRAWING_NAME"
            Case "ORG_NAME", "ОРГАНИЗАЦИЯ", "КОМПАНИЯ"                   : Return "ORG_NAME"
            Case "STAGE", "СТАДИЯ"                                        : Return "STAGE"
            Case "SHEET", "ЛИСТ"                                          : Return "SHEET"
            Case "SHEETS", "ЛИСТОВ"                                       : Return "SHEETS"
        End Select

        ' 2) Транслитерация + нормализация (для нестандартных русских заголовков)
        Dim t As String = Transliterate(n)
        t = System.Text.RegularExpressions.Regex.Replace(t, "[^A-Z0-9]+", "_").Trim("_"c)
        Select Case t
            Case "MODEL_PATH", "MODEL", "P"     : Return "MODEL_PATH"
            Case "CODE", "SHIFR"                : Return "CODE"
            Case "PROJECT_NAME", "OBEKT"        : Return "PROJECT_NAME"
            Case "DRAWING_NAME", "NAIMENOVANIE" : Return "DRAWING_NAME"
            Case "ORG_NAME", "ORGANIZACIYA"     : Return "ORG_NAME"
            Case "STAGE", "STADIYA"             : Return "STAGE"
            Case "SHEET", "LIST"                : Return "SHEET"
            Case "SHEETS", "LISTOV"             : Return "SHEETS"
        End Select

        Return t
    End Function

    Private Shared Function Transliterate(s As String) As String
        Dim src As String() = {"А","Б","В","Г","Д","Е","Ё","Ж","З","И","Й","К","Л","М","Н","О","П","Р","С","Т","У","Ф","Х","Ц","Ч","Ш","Щ","Ъ","Ы","Ь","Э","Ю","Я"}
        Dim dst As String() = {"A","B","V","G","D","E","E","ZH","Z","I","Y","K","L","M","N","O","P","R","S","T","U","F","H","C","CH","SH","SCH","","Y","","E","YU","YA"}
        Dim r As String = s
        For i As Integer = 0 To src.Length - 1
            r = r.Replace(src(i), dst(i))
        Next
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

        ' Поиск по имени файла рекурсивно в папках проекта
        Dim fileName As String = System.IO.Path.GetFileName(input)
        If Not fileName.EndsWith(".ipt", StringComparison.OrdinalIgnoreCase) Then fileName &= ".ipt"

        ' Рекурсивный поиск: проверяем только прямые вложенные папки (без Directory.EnumerateFiles)
        For Each root As String In {workspace, excelDir}
            If String.IsNullOrWhiteSpace(root) Then Continue For
            Dim direct As String = System.IO.Path.Combine(root, fileName)
            If System.IO.File.Exists(direct) Then Return System.IO.Path.GetFullPath(direct)
            ' Один уровень вложенности
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

    Private Shared Function SafeCell(value As Object) As String
        If value Is Nothing Then Return String.Empty
        Return Convert.ToString(value).Trim()
    End Function

End Class
