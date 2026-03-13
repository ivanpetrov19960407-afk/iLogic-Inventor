' ================================================================
' StoneAlbumRule.ilogic.vb  –  v3.1
' Архитектура точно повторяет рабочий VBA RKM_IdwAlbum.bas
' Источник: vba-inventor / RKM_IdwAlbum.bas, RKM_FrameBorder.bas,
'           RKM_TitleBlockPrompted.bas, RKM_Excel.bas
' v3.1: все однострочные Try/Catch развёрнуты (iLogic не поддерживает
'       Try : код : Catch : End Try внутри классов);
'       "Alias" → "MapAlias" (зарезервировано);
'       параметр "shared" → "sst" (зарезервировано).
' ================================================================

Option Explicit On

Imports Inventor
Imports System
Imports System.Collections.Generic

Sub Main()
    ' ── 1. Читаем сохранённые пути ──
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

    ' ── 2. Диалог (всегда, с текущими значениями как дефолт) ──
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

    ' workspace: из проекта Inventor, затем рядом с Excel
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

    ' ── 3. Сохраняем пути обратно ──
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

    ' ── 4. Запуск ──
    Dim doc As DrawingDocument = TryCast(ThisApplication.ActiveDocument, DrawingDocument)
    If doc Is Nothing Then
        System.Windows.Forms.MessageBox.Show("Откройте .idw документ.", "Ошибка")
        Return
    End If

    Dim rule As New AlbumBuilder(ThisApplication)
    rule.Build(doc, excelPath, workspacePath, sheetTabName)
End Sub

' ================================================================
'  BUILDER  (аналог BuildOrUpdateAlbumCore из VBA)
' ================================================================
Public Class AlbumBuilder

    Private ReadOnly _app As Inventor.Application

    ' Константы геометрии А3 СПДС
    Private Const A3_W_MM       As Double = 420.0
    Private Const A3_H_MM       As Double = 297.0
    Private Const FRAME_L_MM    As Double = 20.0
    Private Const FRAME_O_MM    As Double = 5.0
    Private Const TB_W_MM       As Double = 185.0
    Private Const TB_H_MM       As Double = 55.0
    Private Const BORDER_NAME   As String = "RKM_SPDS_A3_BORDER_V12"
    Private Const TB_NAME       As String = "RKM_SPDS_A3_FORM3_V17"
    Private Const SHEET_PFX     As String = "ALB_"

    ' Масштабный ряд (от крупного к мелкому)
    Private ReadOnly SCALES As Double() = {
        5.0, 4.0, 3.0, 2.5, 2.0, 1.5, 1.25, 1.0,
        0.75, 0.5, 0.4, 0.25, 0.2, 0.1, 0.05}

    Public Sub New(app As Inventor.Application)
        _app = app
    End Sub

    ' ── Главная точка входа ──
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

        _app.SilentOperation = True
        Try
            ' Рамка и штамп — один раз для всего документа
            Dim borderDef As BorderDefinition    = EnsureBorder(doc)
            Dim tbDef     As TitleBlockDefinition = EnsureTitleBlock(doc)

            ' Шаблонный лист (первый не-альбомный)
            Dim tmplSheet As Sheet = ResolveTemplateSheet(doc)

            ' Удаляем старые ALB_ листы
            PurgeAlbumSheets(doc, tmplSheet)

            ' Строим по одному листу
            For i As Integer = 0 To items.Count - 1
                Dim item As AlbumItem = items(i)
                If String.IsNullOrWhiteSpace(item.Prompts("SHEET"))  Then item.Prompts("SHEET")  = (i + 1).ToString()
                If String.IsNullOrWhiteSpace(item.Prompts("SHEETS")) Then item.Prompts("SHEETS") = items.Count.ToString()

                Dim ok As Boolean = BuildOneSheet(doc, item, borderDef, tbDef)
                If ok Then okCount += 1 Else failCount += 1
            Next

            ' Активируем шаблонный лист обратно
            If tmplSheet IsNot Nothing Then
                Try
                    tmplSheet.Activate()
                Catch
                End Try
            End If

        Finally
            _app.SilentOperation = False
        End Try

        Dim msg As String = "Альбом собран: " & okCount & " листов."
        If failCount > 0 Then msg &= vbCrLf & "Не собрано: " & failCount & " (модели не найдены или ошибка видов)."
        System.Windows.Forms.MessageBox.Show(msg, "Готово",
            System.Windows.Forms.MessageBoxButtons.OK,
            System.Windows.Forms.MessageBoxIcon.Information)
    End Sub

    ' ── Один лист ──
    Private Function BuildOneSheet(doc As DrawingDocument, item As AlbumItem,
                                   borderDef As BorderDefinition,
                                   tbDef As TitleBlockDefinition) As Boolean
        If Not System.IO.File.Exists(item.ModelPath) Then
            Debug.Print("WARN: модель не найдена: " & item.ModelPath)
            Return False
        End If

        Dim sheet As Sheet = Nothing
        Dim modelDoc As Document = Nothing
        Dim openedHere As Boolean = False

        Try
            ' Создаём лист
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

            ' Рамка СПДС
            Try
                If sheet.Border IsNot Nothing Then sheet.Border.Delete()
            Catch
            End Try
            Try
                sheet.AddCustomBorder(borderDef)
            Catch ex As Exception
                Debug.Print("WARN: AddCustomBorder: " & ex.Message)
            End Try

            ' Штамп Форма 3
            Try
                If sheet.TitleBlock IsNot Nothing Then sheet.TitleBlock.Delete()
            Catch
            End Try
            Dim ps(8) As String
            Dim order As String() = {"CODE","PROJECT_NAME","DRAWING_NAME","ORG_NAME","STAGE","SHEET","SHEETS"}
            For k As Integer = 0 To order.Length - 1
                Dim v As String = String.Empty
                item.Prompts.TryGetValue(order(k), v)
                ps(k + 1) = If(String.IsNullOrEmpty(v), "", v)
            Next
            Try
                sheet.AddTitleBlock(tbDef, Nothing, ps)
            Catch ex As Exception
                Debug.Print("WARN: AddTitleBlock: " & ex.Message)
            End Try

            ' Открываем модель
            For Each ed As Document In _app.Documents
                If String.Equals(ed.FullFileName, item.ModelPath, StringComparison.OrdinalIgnoreCase) Then
                    modelDoc = ed
                    Exit For
                End If
            Next
            If modelDoc Is Nothing Then
                modelDoc = _app.Documents.Open(item.ModelPath, False)
                openedHere = (modelDoc IsNot Nothing)
            End If
            If modelDoc Is Nothing Then
                Debug.Print("WARN: не удалось открыть: " & item.ModelPath)
                Try
                    sheet.Delete()
                Catch
                End Try
                Return False
            End If

            ' Виды
            Dim viewsOk As Boolean = PlaceViews(doc, sheet, modelDoc)
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
    '  ВИДЫ  — измеряем через probe-вид, подбираем масштаб
    ' ================================================================
    Private Function PlaceViews(doc As DrawingDocument, sheet As Sheet, modelDoc As Document) As Boolean
        ' Зона для видов = лист минус штамп и поля
        Dim shW As Double = sheet.Width
        Dim shH As Double = sheet.Height

        Dim padL As Double = MmToCm(doc, FRAME_L_MM + 5)
        Dim padO As Double = MmToCm(doc, FRAME_O_MM + 5)
        Dim tbW  As Double = MmToCm(doc, TB_W_MM)
        Dim tbH  As Double = MmToCm(doc, TB_H_MM)

        ' Доступная зона
        Dim zoneX1 As Double = padL
        Dim zoneY1 As Double = padO + tbH
        Dim zoneX2 As Double = shW - padO
        Dim zoneY2 As Double = shH - padO

        Dim zoneW As Double = zoneX2 - zoneX1
        Dim zoneH As Double = zoneY2 - zoneY1

        ' Измеряем размеры вида при масштабе 0.1 через probe
        Dim probeScale As Double = 0.1
        Dim probeView  As DrawingView = Nothing
        Dim natW As Double = 0
        Dim natH As Double = 0
        Try
            probeView = sheet.DrawingViews.AddBaseView(
                modelDoc,
                _app.TransientGeometry.CreatePoint2d(shW / 2, shH / 2),
                probeScale,
                ViewOrientationTypeEnum.kFrontViewOrientation,
                DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle)
            natW = probeView.Width  / probeScale
            natH = probeView.Height / probeScale
        Catch ex As Exception
            Debug.Print("WARN: probe view failed: " & ex.Message)
            Return False
        Finally
            If probeView IsNot Nothing Then
                Try
                    probeView.Delete()
                Catch
                End Try
            End If
        End Try

        If natW <= 0 OrElse natH <= 0 Then Return False

        ' Подбираем масштаб
        Dim gap As Double = MmToCm(doc, 8)
        Dim selectedScale As Double = SCALES(SCALES.Length - 1)

        For Each sc As Double In SCALES
            Dim frontW As Double = natW * sc
            Dim frontH As Double = natH * sc
            If frontW + natH * sc + gap <= zoneW * 1.05 AndAlso
               frontH + natH * sc * 0.5 + gap <= zoneH * 1.05 Then
                selectedScale = sc
                Exit For
            End If
        Next

        ' Размещаем: Front, Top, Side, Iso
        Dim frontCX As Double = zoneX1 + (zoneW * 0.28)
        Dim frontCY As Double = zoneY1 + (zoneH * 0.38)

        Dim baseView  As DrawingView = Nothing
        Dim topView   As DrawingView = Nothing
        Dim sideView  As DrawingView = Nothing
        Dim isoView   As DrawingView = Nothing

        Try
            ' Front view
            baseView = sheet.DrawingViews.AddBaseView(
                modelDoc,
                _app.TransientGeometry.CreatePoint2d(frontCX, frontCY),
                selectedScale,
                ViewOrientationTypeEnum.kFrontViewOrientation,
                DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle)
            Try
                baseView.ShowLabel = False
            Catch
            End Try

            ' Top — проецируем вверх
            topView = sheet.DrawingViews.AddProjectedView(
                baseView,
                _app.TransientGeometry.CreatePoint2d(frontCX, frontCY + natH * selectedScale + gap),
                DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle)
            Try
                topView.ShowLabel = False
            Catch
            End Try

            ' Right/Left side — проецируем вправо
            sideView = sheet.DrawingViews.AddProjectedView(
                baseView,
                _app.TransientGeometry.CreatePoint2d(frontCX + natW * selectedScale + gap, frontCY),
                DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle)
            Try
                sideView.ShowLabel = False
            Catch
            End Try

            ' Iso — тонированный, правый нижний угол зоны
            Dim isoCX As Double = zoneX1 + zoneW * 0.78
            Dim isoCY As Double = zoneY1 + zoneH * 0.30
            Try
                isoView = sheet.DrawingViews.AddBaseView(
                    modelDoc,
                    _app.TransientGeometry.CreatePoint2d(isoCX, isoCY),
                    selectedScale * 0.75,
                    ViewOrientationTypeEnum.kIsoTopRightViewOrientation,
                    DrawingViewStyleEnum.kShadedDrawingViewStyle)
                Try
                    isoView.ShowLabel = False
                Catch
                End Try
            Catch
            End Try

            Return True

        Catch ex As Exception
            Debug.Print("ERROR: PlaceViews: " & ex.Message)
            If sideView IsNot Nothing Then
                Try
                    sideView.Delete()
                Catch
                End Try
            End If
            If topView IsNot Nothing Then
                Try
                    topView.Delete()
                Catch
                End Try
            End If
            If baseView IsNot Nothing Then
                Try
                    baseView.Delete()
                Catch
                End Try
            End If
            Return False
        End Try
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
            _app.SilentOperation = True
            def = doc.BorderDefinitions.Add(BORDER_NAME)
            _app.SilentOperation = False
        End If

        Dim sk As DrawingSketch = Nothing
        def.Edit(sk)
        Try
            ' Очистка
            For i As Integer = sk.SketchLines.Count To 1 Step -1
                Try
                    sk.SketchLines.Item(i).Delete()
                Catch
                End Try
            Next

            ' Микро-якоря
            sk.SketchLines.AddByTwoPoints(P(0,0), P(0.0001, 0.0001))
            sk.SketchLines.AddByTwoPoints(
                P(Cm(doc, A3_W_MM), Cm(doc, A3_H_MM)),
                P(Cm(doc, A3_W_MM) - 0.0001, Cm(doc, A3_H_MM) - 0.0001))

            ' Внутренняя рамка
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
        If def Is Nothing Then
            _app.SilentOperation = True
            def = doc.TitleBlockDefinitions.Add(TB_NAME)
            _app.SilentOperation = False
        End If

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
    Private Function MmToCm(doc As DrawingDocument, mm As Double) As Double
        Return Cm(doc, mm)
    End Function
    Private Function P(x As Double, y As Double) As Point2d
        Return _app.TransientGeometry.CreatePoint2d(x, y)
    End Function

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
            ' Загружаем System.IO.Compression через Reflection
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
        ' Транслитерация
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
