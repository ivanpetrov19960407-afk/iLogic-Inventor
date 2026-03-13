' ================================================================
' iLogic правило: Автоматическое создание альбома чертежей .idw
' Контекст: изделия из натурального камня (гранит, мрамор)
' ================================================================

Option Explicit On

Imports Inventor
Imports System
Imports System.Collections
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    Dim builder As New StoneAlbumRule(app)

    ' Пример запуска. Пути можно передать через параметры iLogic.
    Dim excelPath As String = "C:\RKM\orders.xlsx"
    Dim workspacePath As String = "C:\RKM\Projects"
    Dim sheetName As String = "ALBUM"

    builder.BuildAlbumFromExcel(excelPath, workspacePath, sheetName)
End Sub

Public Class StoneAlbumRule
    Private ReadOnly _app As Inventor.Application

    ' ---- Константы единиц/геометрии ----
    Private Const MM_TO_CM As Double = 0.1
    Private Const A3_WIDTH_MM As Double = 420.0
    Private Const A3_HEIGHT_MM As Double = 297.0
    Private Const FRAME_LEFT_MM As Double = 20.0
    Private Const FRAME_OTHER_MM As Double = 5.0
    Private Const TITLE_W_MM As Double = 185.0
    Private Const TITLE_H_MM As Double = 55.0

    ' ---- Константы компоновки ----
    Private Const BORDER_NAME As String = "RKM_SPDS_A3_BORDER"
    Private Const TITLEBLOCK_NAME As String = "RKM_SPDS_A3_FORM3"
    Private Const ALBUM_SHEET_PREFIX As String = "ALB_"
    Private Const GAP_MM As Double = 8.0
    Private Const LAYOUT_PAD_MM As Double = 6.0
    Private Const SCALE_STEP As Double = 0.97
    Private Const SCALE_MARGIN As Double = 0.95
    Private Const MAX_AUTO_SCALE As Double = 8.0
    Private Const MIN_AUTO_SCALE As Double = 0.05
    Private Const DIM_OFFSET_MM As Double = 7.0

    Public Sub New(app As Inventor.Application)
        _app = app
    End Sub

    ' Точка входа: чтение Excel -> построение листов -> сохранение IDW
    Public Sub BuildAlbumFromExcel(excelPath As String, workspacePath As String, sheetName As String)
        If String.IsNullOrWhiteSpace(excelPath) Then
            Throw New ArgumentException("Не указан путь к Excel.")
        End If

        Dim doc As DrawingDocument = TryCast(_app.ActiveDocument, DrawingDocument)
        If doc Is Nothing Then
            Throw New InvalidOperationException("Откройте документ .idw перед запуском правила.")
        End If

        Dim items As List(Of AlbumItem) = LoadAlbumItemsFromExcel(excelPath, workspacePath, sheetName)
        If items.Count = 0 Then
            Throw New InvalidOperationException("В Excel не найдено ни одной валидной строки с MODEL_PATH.")
        End If

        Try
            _app.SilentOperation = True

            For Each item As AlbumItem In items
                BuildSingleSheet(doc, item)
            Next

            doc.Update()
            doc.Save2(True)
        Catch ex As Exception
            Logger.Error("Ошибка сборки альбома: " & ex.Message)
            Throw
        Finally
            _app.SilentOperation = False
        End Try
    End Sub

    Private Sub BuildSingleSheet(doc As DrawingDocument, item As AlbumItem)
        If Not File.Exists(item.ModelPath) Then
            Logger.Warn("Модель не найдена: " & item.ModelPath)
            Return
        End If

        Dim modelDoc As Document = Nothing

        Try
            modelDoc = _app.Documents.Open(item.ModelPath, False)
            Dim sheet As Sheet = EnsureA3LandscapeSheet(doc, ALBUM_SHEET_PREFIX & Path.GetFileNameWithoutExtension(item.ModelPath))

            Dim borderDef As BorderDefinition = EnsureSpdsA3BorderDefinition(doc)
            ApplySpdsBorderToSheet(sheet, borderDef)
            ApplyTitleBlockWithPrompts(sheet, item.Prompts)

            BuildViewsWithAutoScale(doc, sheet, modelDoc)
            AutoDimensionOrthographicView(doc, sheet, sheet.DrawingViews.Item(1), "FRONT")

            doc.Update2(True)
        Catch ex As Exception
            Logger.Error("Лист не создан для " & item.ModelPath & ": " & ex.Message)
        Finally
            If modelDoc IsNot Nothing Then
                modelDoc.Close(True)
            End If
        End Try
    End Sub

    ' ========================= Excel & Metadata =========================
    Public Function LoadAlbumItemsFromExcel(excelPath As String, workspacePath As String, sheetName As String) As List(Of AlbumItem)
        Dim result As New List(Of AlbumItem)()

        Dim xlApp As Object = Nothing
        Dim xlBook As Object = Nothing
        Dim xlSheet As Object = Nothing

        Try
            xlApp = CreateObject("Excel.Application")
            xlApp.Visible = False
            xlApp.DisplayAlerts = False

            xlBook = xlApp.Workbooks.Open(excelPath)
            xlSheet = ResolveAlbumWorksheet(xlBook, sheetName)
            If xlSheet Is Nothing Then
                Throw New InvalidOperationException("Лист Excel не найден: " & sheetName)
            End If

            Dim headerRow As Integer = DetectHeaderRowIndex(xlSheet)
            Dim headerMap As Dictionary(Of String, Integer) = ReadHeaderMap(xlSheet, headerRow)
            If Not headerMap.ContainsKey("MODEL_PATH") Then
                Throw New InvalidOperationException("Обязательная колонка MODEL_PATH (или алиас) отсутствует.")
            End If

            Dim modelCol As Integer = headerMap("MODEL_PATH")
            Dim rowCount As Integer = CInt(xlSheet.Cells(xlSheet.Rows.Count, modelCol).End(-4162).Row) ' xlUp

            For row As Integer = headerRow + 1 To rowCount
                Dim rawPath As String = SafeCellText(xlSheet.Cells(row, modelCol).Value)
                If String.IsNullOrWhiteSpace(rawPath) Then Continue For

                Dim resolvedPath As String = ResolveModelPath(rawPath, workspacePath, excelPath)
                If String.IsNullOrWhiteSpace(resolvedPath) Then
                    Logger.Warn("Не удалось разрешить путь модели: " & rawPath)
                    Continue For
                End If

                Dim prompts As Dictionary(Of String, String) = CreatePromptMapFromRow(xlSheet, headerMap, row)
                result.Add(New AlbumItem With {
                    .ModelPath = resolvedPath,
                    .Prompts = prompts
                })
            Next

        Catch ex As Exception
            Logger.Error("Ошибка чтения Excel: " & ex.Message)
            Throw
        Finally
            Try
                If xlBook IsNot Nothing Then xlBook.Close(False)
            Catch
            End Try
            Try
                If xlApp IsNot Nothing Then xlApp.Quit()
            Catch
            End Try
        End Try

        Return result
    End Function

    Private Function ResolveAlbumWorksheet(xlBook As Object, requestedSheetName As String) As Object
        If xlBook Is Nothing Then Return Nothing

        If Not String.IsNullOrWhiteSpace(requestedSheetName) Then
            For Each ws As Object In xlBook.Worksheets
                If String.Equals(CStr(ws.Name), requestedSheetName, StringComparison.OrdinalIgnoreCase) Then
                    Return ws
                End If
            Next
        End If

        If xlBook.Worksheets.Count > 0 Then
            Return xlBook.Worksheets(1)
        End If

        Return Nothing
    End Function

    Private Function DetectHeaderRowIndex(xlSheet As Object) As Integer
        For r As Integer = 1 To 10
            Dim value As String = SafeCellText(xlSheet.Cells(r, 1).Value)
            If Not String.IsNullOrWhiteSpace(value) Then Return r
        Next
        Return 1
    End Function

    Private Function ReadHeaderMap(xlSheet As Object, headerRow As Integer) As Dictionary(Of String, Integer)
        Dim map As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
        Dim lastCol As Integer = CInt(xlSheet.Cells(headerRow, xlSheet.Columns.Count).End(-4159).Column) ' xlToLeft

        For c As Integer = 1 To lastCol
            Dim rawHeader As String = SafeCellText(xlSheet.Cells(headerRow, c).Value)
            Dim canonical As String = ResolveHeaderAlias(rawHeader)
            If Not String.IsNullOrWhiteSpace(canonical) AndAlso Not map.ContainsKey(canonical) Then
                map.Add(canonical, c)
            End If
        Next

        Return map
    End Function

    ' Алиасы заголовков: "CODE" = "Шифр", "PROJECT_NAME" = "Объект" и т.д.
    Private Function ResolveHeaderAlias(rawHeader As String) As String
        Dim n As String = rawHeader.Trim().ToUpperInvariant()

        Select Case n
            Case "MODEL_PATH", "MODEL", "P", "ПУТЬ", "ФАЙЛ", "МОДЕЛЬ"
                Return "MODEL_PATH"
            Case "CODE", "ШИФР", "АРТИКУЛ", "ОБОЗНАЧЕНИЕ"
                Return "CODE"
            Case "PROJECT_NAME", "PROJECT", "ОБЪЕКТ", "ПРОЕКТ"
                Return "PROJECT_NAME"
            Case "DRAWING_NAME", "TITLE", "НАИМЕНОВАНИЕ", "ИМЯ ЧЕРТЕЖА"
                Return "DRAWING_NAME"
            Case "ORG_NAME", "ОРГАНИЗАЦИЯ", "КОМПАНИЯ"
                Return "ORG_NAME"
            Case "STAGE", "СТАДИЯ"
                Return "STAGE"
            Case "SHEET", "ЛИСТ"
                Return "SHEET"
            Case "SHEETS", "ЛИСТОВ"
                Return "SHEETS"
        End Select

        Return String.Empty
    End Function

    ' ResolveModelPath: абсолютный путь -> workspace -> папка Excel -> рекурсивный поиск по имени
    Private Function ResolveModelPath(inputPath As String, workspacePath As String, excelPath As String) As String
        If String.IsNullOrWhiteSpace(inputPath) Then Return String.Empty

        Dim normalized As String = inputPath.Trim()

        If File.Exists(normalized) Then Return Path.GetFullPath(normalized)

        If Not String.IsNullOrWhiteSpace(workspacePath) Then
            Dim candidate As String = Path.Combine(workspacePath, normalized)
            If File.Exists(candidate) Then Return Path.GetFullPath(candidate)
        End If

        Dim excelDir As String = Path.GetDirectoryName(excelPath)
        If Not String.IsNullOrWhiteSpace(excelDir) Then
            Dim candidate As String = Path.Combine(excelDir, normalized)
            If File.Exists(candidate) Then Return Path.GetFullPath(candidate)
        End If

        ' Поиск .ipt по имени в папке проекта (если передано только имя/код)
        Dim fileName As String = Path.GetFileName(normalized)
        If Not fileName.EndsWith(".ipt", StringComparison.OrdinalIgnoreCase) Then
            fileName &= ".ipt"
        End If

        For Each root In New String() {workspacePath, excelDir}
            If String.IsNullOrWhiteSpace(root) OrElse Not Directory.Exists(root) Then Continue For

            Dim found As String = Directory.EnumerateFiles(root, fileName, SearchOption.AllDirectories).FirstOrDefault()
            If Not String.IsNullOrWhiteSpace(found) Then Return Path.GetFullPath(found)
        Next

        Return String.Empty
    End Function

    Private Function CreatePromptMapFromRow(xlSheet As Object, headerMap As Dictionary(Of String, Integer), row As Integer) As Dictionary(Of String, String)
        Dim result As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

        For Each key As String In GetPromptOrder()
            If headerMap.ContainsKey(key) Then
                result(key) = SafeCellText(xlSheet.Cells(row, headerMap(key)).Value)
            Else
                result(key) = String.Empty
            End If
        Next

        Return result
    End Function

    Private Function SafeCellText(value As Object) As String
        If value Is Nothing Then Return String.Empty
        Return Convert.ToString(value).Trim()
    End Function

    Private Function GetPromptOrder() As String()
        Return New String() {"CODE", "PROJECT_NAME", "DRAWING_NAME", "ORG_NAME", "STAGE", "SHEET", "SHEETS"}
    End Function

    ' ========================= Математика и единицы =========================
    Private Function MmToCm(valueMm As Double, doc As DrawingDocument) As Double
        If doc Is Nothing Then Return valueMm * MM_TO_CM
        Return doc.UnitsOfMeasure.ConvertUnits(valueMm, UnitsTypeEnum.kMillimeterLengthUnits, UnitsTypeEnum.kCentimeterLengthUnits)
    End Function

    Private Function CmToMm(valueCm As Double, doc As DrawingDocument) As Double
        If doc Is Nothing Then Return valueCm / MM_TO_CM
        Return doc.UnitsOfMeasure.ConvertUnits(valueCm, UnitsTypeEnum.kCentimeterLengthUnits, UnitsTypeEnum.kMillimeterLengthUnits)
    End Function

    Private Function Pt2d(x As Double, y As Double) As Point2d
        Return _app.TransientGeometry.CreatePoint2d(x, y)
    End Function

    ' ========================= СПДС: лист, рамка, штамп =========================
    Private Function EnsureA3LandscapeSheet(doc As DrawingDocument, sheetName As String) As Sheet
        For Each s As Sheet In doc.Sheets
            If String.Equals(s.Name, sheetName, StringComparison.OrdinalIgnoreCase) Then
                Return s
            End If
        Next

        Dim sheet As Sheet = doc.Sheets.Add(DrawingSheetSizeEnum.kA3DrawingSheetSize, PageOrientationTypeEnum.kLandscapePageOrientation, sheetName)
        Return sheet
    End Function

    Private Function EnsureSpdsA3BorderDefinition(doc As DrawingDocument) As BorderDefinition
        Dim def As BorderDefinition = Nothing

        For Each bd As BorderDefinition In doc.BorderDefinitions
            If String.Equals(bd.Name, BORDER_NAME, StringComparison.OrdinalIgnoreCase) Then
                def = bd
                Exit For
            End If
        Next

        If def Is Nothing Then
            def = doc.BorderDefinitions.Add(BORDER_NAME)
        End If

        Dim sk As DrawingSketch = Nothing
        def.Edit(sk)
        Try
            ClearSketch(sk)
            DrawSpdsBorderGeometry(sk, doc)
        Finally
            def.ExitEdit(True)
        End Try

        Return def
    End Function

    Private Sub ClearSketch(sketch As DrawingSketch)
        If sketch Is Nothing Then Return

        For i As Integer = sketch.SketchLines.Count To 1 Step -1
            sketch.SketchLines.Item(i).Delete()
        Next
    End Sub

    Private Sub DrawSpdsBorderGeometry(sk As DrawingSketch, doc As DrawingDocument)
        Dim xMax As Double = MmToCm(A3_WIDTH_MM, doc)
        Dim yMax As Double = MmToCm(A3_HEIGHT_MM, doc)

        Dim ix1 As Double = MmToCm(FRAME_LEFT_MM, doc)
        Dim iy1 As Double = MmToCm(FRAME_OTHER_MM, doc)
        Dim ix2 As Double = MmToCm(A3_WIDTH_MM - FRAME_OTHER_MM, doc)
        Dim iy2 As Double = MmToCm(A3_HEIGHT_MM - FRAME_OTHER_MM, doc)

        sk.SketchLines.AddAsTwoPointRectangle(Pt2d(0, 0), Pt2d(xMax, yMax))
        sk.SketchLines.AddAsTwoPointRectangle(Pt2d(ix1, iy1), Pt2d(ix2, iy2))
    End Sub

    Private Sub ApplySpdsBorderToSheet(sheet As Sheet, def As BorderDefinition)
        If sheet Is Nothing OrElse def Is Nothing Then Return

        If sheet.Border IsNot Nothing Then
            sheet.Border.Delete()
        End If

        sheet.AddBorder(def)
    End Sub

    Private Sub ApplyTitleBlockWithPrompts(sheet As Sheet, prompts As Dictionary(Of String, String))
        If sheet Is Nothing Then Return

        Dim def As TitleBlockDefinition = Nothing
        For Each td As TitleBlockDefinition In sheet.Parent.TitleBlockDefinitions
            If String.Equals(td.Name, TITLEBLOCK_NAME, StringComparison.OrdinalIgnoreCase) Then
                def = td
                Exit For
            End If
        Next

        If def Is Nothing Then
            Logger.Warn("Определение штампа не найдено: " & TITLEBLOCK_NAME)
            Return
        End If

        Dim promptStrings(7) As String ' 1..7
        Dim order As String() = GetPromptOrder()

        For i As Integer = 0 To order.Length - 1
            Dim key As String = order(i)
            If prompts IsNot Nothing AndAlso prompts.ContainsKey(key) Then
                promptStrings(i + 1) = prompts(key)
            Else
                promptStrings(i + 1) = String.Empty
            End If
        Next

        Try
            _app.SilentOperation = True
            If sheet.TitleBlock IsNot Nothing Then
                sheet.TitleBlock.Delete()
            End If
            sheet.AddTitleBlock(def, Nothing, promptStrings)
        Catch ex As Exception
            Logger.Warn("Не удалось заполнить штамп на листе " & sheet.Name & ": " & ex.Message)
        Finally
            _app.SilentOperation = False
        End Try
    End Sub

    ' ========================= Размещение и масштаб =========================
    Private Sub BuildViewsWithAutoScale(doc As DrawingDocument, sheet As Sheet, modelDoc As Document)
        Dim modelDef As ComponentDefinition = TryCast(modelDoc.ComponentDefinition, ComponentDefinition)
        If modelDef Is Nothing Then
            Throw New InvalidOperationException("Не удалось получить ComponentDefinition у модели: " & modelDoc.DisplayName)
        End If

        Dim modelSize As ModelSize = MeasureModelSizeCm(modelDef)
        Dim scale As Double = FindBestScaleFor3Views(doc, modelSize)

        Dim layout As LayoutData = BuildViewLayout(doc, scale, modelSize)

        Dim frontView As DrawingView = sheet.DrawingViews.AddBaseView(
            modelDoc,
            Pt2d(layout.FrontX, layout.FrontY),
            scale,
            ViewOrientationTypeEnum.kFrontViewOrientation,
            DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle)

        sheet.DrawingViews.AddProjectedView(frontView, Pt2d(layout.TopX, layout.TopY), DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle)
        sheet.DrawingViews.AddProjectedView(frontView, Pt2d(layout.LeftX, layout.LeftY), DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle)

        sheet.DrawingViews.AddBaseView(
            modelDoc,
            Pt2d(layout.IsoX, layout.IsoY),
            Math.Max(scale * 0.7, MIN_AUTO_SCALE),
            ViewOrientationTypeEnum.kIsoTopRightViewOrientation,
            DrawingViewStyleEnum.kShadedDrawingViewStyle)
    End Sub

    ' Итерационный алгоритм: уменьшаем масштаб, пока 3 вида не влезут в рабочую область и не наедут на штамп
    Private Function FindBestScaleFor3Views(doc As DrawingDocument, modelSize As ModelSize) As Double
        Dim scale As Double = MAX_AUTO_SCALE

        Do While scale > MIN_AUTO_SCALE
            If FitsAtScale(doc, modelSize, scale) Then
                Return scale * SCALE_MARGIN
            End If
            scale *= SCALE_STEP
        Loop

        Return MIN_AUTO_SCALE
    End Function

    Private Function FitsAtScale(doc As DrawingDocument, modelSize As ModelSize, scale As Double) As Boolean
        Dim gap As Double = MmToCm(GAP_MM, doc)
        Dim pad As Double = MmToCm(LAYOUT_PAD_MM, doc)

        Dim innerW As Double = MmToCm(A3_WIDTH_MM - FRAME_LEFT_MM - FRAME_OTHER_MM, doc)
        Dim innerH As Double = MmToCm(A3_HEIGHT_MM - 2 * FRAME_OTHER_MM, doc)

        ' Зона под ортогональные виды: рамка минус область штампа (185x55 мм) справа снизу
        Dim workW As Double = innerW - (2 * pad)
        Dim workH As Double = innerH - MmToCm(TITLE_H_MM, doc) - (2 * pad)

        Dim frontW As Double = modelSize.X * scale
        Dim frontH As Double = modelSize.Z * scale
        Dim topW As Double = modelSize.X * scale
        Dim topH As Double = modelSize.Y * scale
        Dim leftW As Double = modelSize.Y * scale
        Dim leftH As Double = modelSize.Z * scale

        Dim needW As Double = leftW + gap + frontW
        Dim needH As Double = frontH + gap + topH

        Return needW <= workW AndAlso needH <= workH AndAlso leftH <= workH AndAlso topW <= workW
    End Function

    Private Function BuildViewLayout(doc As DrawingDocument, scale As Double, modelSize As ModelSize) As LayoutData
        Dim gap As Double = MmToCm(GAP_MM, doc)
        Dim pad As Double = MmToCm(LAYOUT_PAD_MM, doc)

        Dim xMin As Double = MmToCm(FRAME_LEFT_MM, doc) + pad
        Dim yMin As Double = MmToCm(FRAME_OTHER_MM, doc) + MmToCm(TITLE_H_MM, doc) + pad
        Dim xMax As Double = MmToCm(A3_WIDTH_MM - FRAME_OTHER_MM, doc) - pad
        Dim yMax As Double = MmToCm(A3_HEIGHT_MM - FRAME_OTHER_MM, doc) - pad

        Dim frontW As Double = modelSize.X * scale
        Dim frontH As Double = modelSize.Z * scale
        Dim topH As Double = modelSize.Y * scale
        Dim leftW As Double = modelSize.Y * scale

        Dim centerX As Double = xMin + (xMax - xMin) * 0.52
        Dim centerY As Double = yMin + (yMax - yMin) * 0.45

        Dim frontX As Double = centerX
        Dim frontY As Double = centerY

        Dim topX As Double = frontX
        Dim topY As Double = frontY + (frontH / 2) + gap + (topH / 2)

        Dim leftX As Double = frontX - (frontW / 2) - gap - (leftW / 2)
        Dim leftY As Double = frontY

        Dim isoX As Double = xMax - MmToCm(TITLE_W_MM, doc) * 0.45
        Dim isoY As Double = yMin + MmToCm(TITLE_H_MM, doc) * 0.75

        Return New LayoutData With {
            .FrontX = frontX,
            .FrontY = frontY,
            .TopX = topX,
            .TopY = topY,
            .LeftX = leftX,
            .LeftY = leftY,
            .IsoX = isoX,
            .IsoY = isoY
        }
    End Function

    Private Function MeasureModelSizeCm(modelDef As ComponentDefinition) As ModelSize
        Dim box As Box = modelDef.RangeBox
        Return New ModelSize With {
            .X = Math.Abs(box.MaxPoint.X - box.MinPoint.X),
            .Y = Math.Abs(box.MaxPoint.Y - box.MinPoint.Y),
            .Z = Math.Abs(box.MaxPoint.Z - box.MinPoint.Z)
        }
    End Function

    ' ========================= AutoDimension =========================
    Private Sub AutoDimensionOrthographicView(doc As DrawingDocument, sheet As Sheet, view As DrawingView, viewKey As String)
        If view Is Nothing Then Return

        Dim ext As ViewExtents = GetViewExtents(view)
        If ext Is Nothing Then
            Logger.Warn("Не удалось определить габариты вида для AutoDimension: " & viewKey)
            Return
        End If

        AddHorizontalOverallDimension(doc, sheet, view, ext)
        AddVerticalOverallDimension(doc, sheet, view, ext)
    End Sub

    Private Sub AddHorizontalOverallDimension(doc As DrawingDocument, sheet As Sheet, view As DrawingView, ext As ViewExtents)
        If ext.LeftCurve Is Nothing OrElse ext.RightCurve Is Nothing Then Return

        Dim intentLeft As GeometryIntent = sheet.CreateGeometryIntent(ext.LeftCurve)
        Dim intentRight As GeometryIntent = sheet.CreateGeometryIntent(ext.RightCurve)
        Dim textPt As Point2d = Pt2d((ext.MinX + ext.MaxX) / 2.0, ext.MinY - MmToCm(DIM_OFFSET_MM, doc))

        Try
            sheet.DrawingDimensions.GeneralDimensions.AddLinear(textPt, intentLeft, intentRight, DimensionTypeEnum.kHorizontalDimensionType)
        Catch ex As Exception
            Logger.Warn("Горизонтальный габарит не поставлен: " & ex.Message)
        End Try
    End Sub

    Private Sub AddVerticalOverallDimension(doc As DrawingDocument, sheet As Sheet, view As DrawingView, ext As ViewExtents)
        If ext.BottomCurve Is Nothing OrElse ext.TopCurve Is Nothing Then Return

        Dim intentBottom As GeometryIntent = sheet.CreateGeometryIntent(ext.BottomCurve)
        Dim intentTop As GeometryIntent = sheet.CreateGeometryIntent(ext.TopCurve)
        Dim textPt As Point2d = Pt2d(ext.MaxX + MmToCm(DIM_OFFSET_MM, doc), (ext.MinY + ext.MaxY) / 2.0)

        Try
            sheet.DrawingDimensions.GeneralDimensions.AddLinear(textPt, intentBottom, intentTop, DimensionTypeEnum.kVerticalDimensionType)
        Catch ex As Exception
            Logger.Warn("Вертикальный габарит не поставлен: " & ex.Message)
        End Try
    End Sub

    Private Function GetViewExtents(view As DrawingView) As ViewExtents
        Dim hasAny As Boolean = False
        Dim minX As Double = Double.MaxValue
        Dim maxX As Double = Double.MinValue
        Dim minY As Double = Double.MaxValue
        Dim maxY As Double = Double.MinValue

        Dim leftCurve As DrawingCurve = Nothing
        Dim rightCurve As DrawingCurve = Nothing
        Dim topCurve As DrawingCurve = Nothing
        Dim bottomCurve As DrawingCurve = Nothing

        For Each curve As DrawingCurve In view.DrawingCurves
            Dim s As Point2d = curve.StartPoint
            Dim e As Point2d = curve.EndPoint
            If s Is Nothing OrElse e Is Nothing Then Continue For

            hasAny = True

            Dim cMinX As Double = Math.Min(s.X, e.X)
            Dim cMaxX As Double = Math.Max(s.X, e.X)
            Dim cMinY As Double = Math.Min(s.Y, e.Y)
            Dim cMaxY As Double = Math.Max(s.Y, e.Y)

            If cMinX < minX Then
                minX = cMinX
                leftCurve = curve
            End If
            If cMaxX > maxX Then
                maxX = cMaxX
                rightCurve = curve
            End If
            If cMinY < minY Then
                minY = cMinY
                bottomCurve = curve
            End If
            If cMaxY > maxY Then
                maxY = cMaxY
                topCurve = curve
            End If
        Next

        If Not hasAny Then Return Nothing

        Return New ViewExtents With {
            .MinX = minX,
            .MaxX = maxX,
            .MinY = minY,
            .MaxY = maxY,
            .LeftCurve = leftCurve,
            .RightCurve = rightCurve,
            .TopCurve = topCurve,
            .BottomCurve = bottomCurve
        }
    End Function

    ' ========================= DTO =========================
    Private Class AlbumItem
        Public Property ModelPath As String
        Public Property Prompts As Dictionary(Of String, String)
    End Class

    Private Class ModelSize
        Public Property X As Double
        Public Property Y As Double
        Public Property Z As Double
    End Class

    Private Class LayoutData
        Public Property FrontX As Double
        Public Property FrontY As Double
        Public Property TopX As Double
        Public Property TopY As Double
        Public Property LeftX As Double
        Public Property LeftY As Double
        Public Property IsoX As Double
        Public Property IsoY As Double
    End Class

    Private Class ViewExtents
        Public Property MinX As Double
        Public Property MaxX As Double
        Public Property MinY As Double
        Public Property MaxY As Double
        Public Property LeftCurve As DrawingCurve
        Public Property RightCurve As DrawingCurve
        Public Property TopCurve As DrawingCurve
        Public Property BottomCurve As DrawingCurve
    End Class

    Private NotInheritable Class Logger
        Public Shared Sub Warn(msg As String)
            Debug.Print("WARN: " & msg)
        End Sub

        Public Shared Sub [Error](msg As String)
            Debug.Print("ERROR: " & msg)
        End Sub
    End Class
End Class
