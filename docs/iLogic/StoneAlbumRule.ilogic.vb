' ================================================================
' StoneAlbumRule.ilogic.vb  –  v3.20
' Архитектура точно повторяет рабочий VBA RKM_IdwAlbum.bas
' Источник: vba-inventor / RKM_IdwAlbum.bas, RKM_FrameBorder.bas,
'           RKM_TitleBlockPrompted.bas, RKM_Excel.bas
' v3.20: FIX dimensions — PlaceViewsByTemplate now adds semantic role aliases
'        into placedViews so ApplyDimensionPlan can find PlanContour/ThicknessView etc.
'        ResolveExistingRole extended to cover all possible role mismatches.
' v3.19: add facade-left long-linear layout subtype,
'        support left tall shaded facade + top profile + bottom-right iso,
'        improve long profiled step placement and scoring
' v3.18: keep all opposite views as candidates until final scoring,
'        fix wrong-side selection for steps/profiled parts,
'        add second plate/block layout variant,
'        improve sample-driven layout scoring
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
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices

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
    Private _lastBuildFailReason As SheetBuildFailReason = SheetBuildFailReason.None

    Private Enum SheetBuildFailReason
        None
        ModelPathUnresolved
        FileNotFound
        DocumentOpenFailed
        ProbeMeasureFailed
        LayoutSelectionFailed
        ViewPlacementFailed
        LessThan2FinalViews
        Unknown
    End Enum

    ' Геометрия А3 СПДС
    Private Const A3_W_MM       As Double = 420.0
    Private Const A3_H_MM       As Double = 297.0
    Private Const FRAME_L_MM    As Double = 20.0
    Private Const FRAME_O_MM    As Double = 5.0
    Private Const TB_W_MM       As Double = 185.0
    Private Const TB_H_MM       As Double = 55.0
    Private Const TITLE_TEXT_HEIGHT_MM As Double = 1.4
    Private Const BORDER_NAME   As String = "RKM_SPDS_A3_BORDER_V12"
    Private Const TB_NAME       As String = "RKM_SPDS_A3_FORM3_V17"
    Private Const SHEET_PFX     As String = "ALB_"
    Private Const ALBUM_MODE_VISUAL As Boolean = True
    Private Const ADD_VIEW_NOTES As Boolean = True
    Private Const ADD_VIEW_DIMENSIONS As Boolean = True
    Private Const DIMENSIONS_ON_MAIN_ISO As Boolean = False
    Private Const FORCE_ISOMETRIC_VIEWS As Boolean = True
    Private Const VIEW_CAPTION_GAP_MM As Double = 5.0
    Private Const VIEW_DIM_MIN_MM As Double = 10.0

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
        Dim modelPathUnresolvedCount As Integer = 0
        Dim fileNotFoundCount As Integer = 0
        Dim documentOpenFailedCount As Integer = 0
        Dim probeMeasureFailedCount As Integer = 0
        Dim layoutSelectionFailedCount As Integer = 0
        Dim viewPlacementFailedCount As Integer = 0
        Dim lessThan2FinalViewsCount As Integer = 0

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

                Dim excelRow As Integer = If(item.ExcelRow > 0, item.ExcelRow, i + 1)
                Dim ok As Boolean = BuildOneSheet(doc, item, excelRow)
                If ok Then
                    okCount += 1
                Else
                    failCount += 1
                    Select Case _lastBuildFailReason
                        Case SheetBuildFailReason.ModelPathUnresolved
                            modelPathUnresolvedCount += 1
                        Case SheetBuildFailReason.FileNotFound
                            fileNotFoundCount += 1
                        Case SheetBuildFailReason.DocumentOpenFailed
                            documentOpenFailedCount += 1
                        Case SheetBuildFailReason.ProbeMeasureFailed
                            probeMeasureFailedCount += 1
                        Case SheetBuildFailReason.LayoutSelectionFailed
                            layoutSelectionFailedCount += 1
                        Case SheetBuildFailReason.ViewPlacementFailed
                            viewPlacementFailedCount += 1
                        Case SheetBuildFailReason.LessThan2FinalViews
                            lessThan2FinalViewsCount += 1
                    End Select
                    Debug.Print("WARN: лист не собран, строка Excel=" & excelRow.ToString() & ", причина=" & _lastBuildFailReason.ToString() & ", модель=" & item.ModelPath)
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
        If failCount > 0 Then
            msg &= vbCrLf & "Не собрано: " & failCount
            msg &= vbCrLf & "Пути не найдены: " & modelPathUnresolvedCount
            msg &= vbCrLf & "Файлы не найдены: " & fileNotFoundCount
            msg &= vbCrLf & "Не открылись модели: " & documentOpenFailedCount
            msg &= vbCrLf & "Ошибки пробных видов: " & probeMeasureFailedCount
            msg &= vbCrLf & "Ошибка подбора layout: " & layoutSelectionFailedCount
            msg &= vbCrLf & "Ошибка размещения видов: " & viewPlacementFailedCount
            msg &= vbCrLf & "Менее 2 итоговых видов: " & lessThan2FinalViewsCount
        End If
        System.Windows.Forms.MessageBox.Show(msg, "Готово",
            System.Windows.Forms.MessageBoxButtons.OK,
            System.Windows.Forms.MessageBoxIcon.Information)
    End Sub

    ' ── Один лист ──
    ' v3.6: borderDef/tbDef получаем заново внутри каждого вызова
    '       (COM RCW протухает после doc.Update2 если ссылка хранится снаружи)
    Private Function BuildOneSheet(doc As DrawingDocument, item As AlbumItem, rowIndex As Integer) As Boolean
        _lastBuildFailReason = SheetBuildFailReason.None
        If String.IsNullOrWhiteSpace(item.ModelPath) Then
            _lastBuildFailReason = SheetBuildFailReason.ModelPathUnresolved
            Debug.Print("WARN: путь модели не разрешён, строка Excel=" & rowIndex.ToString() & ", исходное значение='" & item.SourceModelRaw & "'")
            Return False
        End If
        If Not System.IO.File.Exists(item.ModelPath) Then
            _lastBuildFailReason = SheetBuildFailReason.FileNotFound
            Debug.Print("WARN: модель не найдена, строка Excel=" & rowIndex.ToString() & ", путь=" & item.ModelPath)
            Return False
        End If

        Dim sheet As Sheet = Nothing
        Dim modelDoc As Document = Nothing
        Dim openedHere As Boolean = False

        Try
            Dim baseSheetName As String = SHEET_PFX & System.IO.Path.GetFileNameWithoutExtension(item.ModelPath)
            Dim sheetName As String = MakeUniqueSheetName(doc, baseSheetName)
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
                _lastBuildFailReason = SheetBuildFailReason.DocumentOpenFailed
                Debug.Print("WARN: не удалось открыть, строка Excel=" & rowIndex.ToString() & ", модель=" & item.ModelPath)
                Try
                    sheet.Delete()
                Catch
                End Try
                Return False
            End If

            Dim modelSize As ModelOverallExtents = GetModelOverallExtentsMm(modelDoc)

            ' Виды через slot-based layout
            Dim viewsFailReason As SheetBuildFailReason = SheetBuildFailReason.None
            Dim viewsOk As Boolean = PlaceViewsSlotBased(doc, sheet, modelDoc, modelSize, viewsFailReason)
            If Not viewsOk Then
                _lastBuildFailReason = viewsFailReason
                Debug.Print("WARN: не удалось построить виды, строка Excel=" & rowIndex.ToString() & ", лист=" & sheetName)
                Try
                    sheet.Delete()
                Catch
                End Try
                Return False
            End If
            doc.Update2(True)

            Dim finalViewCount As Integer = sheet.DrawingViews.Count
            Debug.Print("INFO final view count on sheet '" & sheetName & "': " & finalViewCount.ToString())
            If finalViewCount < 2 Then
                _lastBuildFailReason = SheetBuildFailReason.LessThan2FinalViews
                Debug.Print("WARN: менее 2 видов на листе: " & sheetName)
                Try
                    sheet.Delete()
                Catch
                End Try
                Return False
            End If

            _lastBuildFailReason = SheetBuildFailReason.None
            Return True

        Catch ex As Exception
            _lastBuildFailReason = SheetBuildFailReason.Unknown
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
    Private Function PlaceViewsSlotBased(doc As DrawingDocument, sheet As Sheet, modelDoc As Document, modelSize As ModelOverallExtents, ByRef failReason As SheetBuildFailReason) As Boolean
        failReason = SheetBuildFailReason.None
        Dim safe As SlotRect = GetSheetSafeRect(doc, sheet)
        Dim gap As Double = Cm(doc, GAP_MM)

        Dim mFront As ViewMeasure = MeasureView(sheet, modelDoc, ViewOrientationTypeEnum.kFrontViewOrientation, DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle, "FRONT", "Вид спереди")
        Dim mBack As ViewMeasure = MeasureView(sheet, modelDoc, ViewOrientationTypeEnum.kBackViewOrientation, DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle, "BACK", "Вид сзади")
        Dim mTop As ViewMeasure = MeasureView(sheet, modelDoc, ViewOrientationTypeEnum.kTopViewOrientation, DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle, "TOP", "Вид сверху")
        Dim mLeft As ViewMeasure = MeasureView(sheet, modelDoc, ViewOrientationTypeEnum.kLeftViewOrientation, DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle, "LEFT", "Вид слева")
        Dim mRight As ViewMeasure = MeasureView(sheet, modelDoc, ViewOrientationTypeEnum.kRightViewOrientation, DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle, "RIGHT", "Вид справа")
        Dim mIso As ViewMeasure = MeasureView(sheet, modelDoc, ViewOrientationTypeEnum.kIsoTopRightViewOrientation, DrawingViewStyleEnum.kShadedDrawingViewStyle, "ISO", "Изометрия")
        If mIso Is Nothing Then
            Debug.Print("WARN: iso completely unavailable, continue without iso")
        End If
        Dim all2D As List(Of ViewMeasure) = BuildAll2DCandidates(mFront, mBack, mTop, mLeft, mRight)
        If all2D Is Nothing OrElse all2D.Count < 2 Then
            failReason = SheetBuildFailReason.ProbeMeasureFailed
            Debug.Print("WARN: недостаточно 2D-кандидатов для подбора layout")
            Return False
        End If

        Dim descriptor As PartDescriptor = ClassifyPart(all2D)
        DebugPrintDescriptor(descriptor)

        Dim roleMap As RoleMap = ResolveRoles(descriptor, all2D, mIso)
        DebugPrintRoleMap(roleMap)

        Dim templates As List(Of LayoutTemplate) = BuildLayoutTemplates(doc, safe, gap, descriptor)
        If templates Is Nothing OrElse templates.Count = 0 Then
            Debug.Print("WARN: не удалось построить шаблоны layout")
            Return False
        End If

        Dim best As LayoutPlan = EvaluateBestTemplate(templates, descriptor, roleMap)
        If best Is Nothing Then
            failReason = SheetBuildFailReason.LayoutSelectionFailed
            Debug.Print("WARN: не найден корректный layout визуализации")
            Return False
        End If

        If FORCE_ISOMETRIC_VIEWS Then
            best.MainMeasure = ForceIsometricMeasure(best.MainMeasure, mIso)
            best.Aux1Measure = ForceIsometricMeasure(best.Aux1Measure, mIso)
            best.Aux2Measure = ForceIsometricMeasure(best.Aux2Measure, mIso)

            best.MainFit = ScaleToFit(best.MainSlot, best.MainMeasure, ISO_SCALE_MARGIN)
            best.Aux1Fit = ScaleToFit(best.Aux1Slot, best.Aux1Measure, ISO_SCALE_MARGIN)
            If best.Aux2Measure IsNot Nothing Then
                best.Aux2Fit = ScaleToFit(best.Aux2Slot, best.Aux2Measure, ISO_SCALE_MARGIN)
            End If
            If best.MainFit Is Nothing OrElse best.Aux1Fit Is Nothing Then
                failReason = SheetBuildFailReason.LayoutSelectionFailed
                Debug.Print("WARN: изометрический layout не рассчитан")
                Return False
            End If
        End If

        Debug.Print("Layout template=" & best.TemplateName & ", score=" & String.Format("{0:F3}", best.Score))

        Dim placedViews As Dictionary(Of ViewRole, DrawingView) = PlaceViewsByTemplate(sheet, modelDoc, best, roleMap)
        If placedViews Is Nothing OrElse placedViews.Count = 0 Then
            failReason = SheetBuildFailReason.ViewPlacementFailed
            Debug.Print("WARN: не удалось разместить виды по template")
            Return False
        End If

        If Not FORCE_ISOMETRIC_VIEWS Then
            Dim finalOrthogonal As Integer = 0
            If placedViews.ContainsKey(best.MainRole) Then finalOrthogonal += 1
            If placedViews.ContainsKey(best.AuxRole) Then finalOrthogonal += 1
            If finalOrthogonal < 2 Then
                failReason = SheetBuildFailReason.LessThan2FinalViews
                Debug.Print("WARN: после размещения менее 2 ортогональных видов")
                Return False
            End If
        End If

        If ADD_VIEW_NOTES Then
            AddViewRoleCaptions(doc, sheet, best, placedViews)
        End If

        If ADD_VIEW_DIMENSIONS Then
            Dim dimPlan As DimensionPlan = BuildDimensionPlan(descriptor, roleMap)
            DebugPrintDimensionPlan(dimPlan)
            Try
                ApplyDimensionPlan(doc, sheet, placedViews, roleMap, descriptor, dimPlan, modelSize)
            Catch ex As Exception
                Debug.Print("WARN dimensioning failed: " & ex.Message & "; views preserved")
            End Try
        End If

        Debug.Print("INFO final view count on sheet '" & sheet.Name & "': " & sheet.DrawingViews.Count.ToString())

        Return True
    End Function

    Private Sub DebugPrintDescriptor(d As PartDescriptor)
        If d Is Nothing Then Return
        Debug.Print("Part family=" & d.Family.ToString() &
                    ", archetype=" & d.DimArchetype.ToString() &
                    ", long=" & d.IsLong.ToString() &
                    ", thin=" & d.IsThin.ToString() &
                    ", domPlan=" & d.HasDominantPlan.ToString() &
                    ", domFacade=" & d.HasDominantFacade.ToString() &
                    ", radial=" & d.HasRadialPlan.ToString() &
                    ", slope=" & d.HasSlope.ToString() &
                    ", profileComplex=" & d.HasComplexProfile.ToString())
    End Sub

    Private Sub DebugPrintRoleMap(roleMap As RoleMap)
        If roleMap Is Nothing Then Return
        For Each kv As KeyValuePair(Of ViewRole, ViewMeasure) In roleMap.ByRole
            If kv.Value Is Nothing Then Continue For
            Debug.Print("Role " & kv.Key.ToString() & " => " & kv.Value.Key)
        Next
    End Sub

    Private Sub DebugPrintDimensionPlan(plan As DimensionPlan)
        If plan Is Nothing Then Return
        Dim ids As New List(Of String)()
        For Each i As DimensionIntent In plan.Intents
            ids.Add(i.IntentId.ToString() & "@" & i.PreferredRole.ToString() & "[P" & i.Priority.ToString() & "]")
        Next
        Debug.Print("DimensionPlan: " & String.Join(", ", ids.ToArray()))
    End Sub

    Private Function ClassifyPart(measures As List(Of ViewMeasure)) As PartDescriptor
        Dim d As New PartDescriptor()
        If measures Is Nothing OrElse measures.Count = 0 Then
            d.Family = PartFamily.Plate
            d.DimArchetype = DimensionArchetype.PlateSimple
            d.PreferredMainRole = ViewRole.PlanContour
            Return d
        End If

        Dim maxAr As Double = 0
        Dim maxSlope As Double = 0
        Dim maxArcRatio As Double = 0
        Dim maxProfile As Double = 0
        Dim maxPlan As Double = 0
        Dim strongSlopeCount As Integer = 0
        Dim strongArcCount As Integer = 0

        Dim largest As ViewMeasure = Nothing
        Dim second As ViewMeasure = Nothing

        For Each m As ViewMeasure In measures
            If m Is Nothing Then Continue For
            If largest Is Nothing OrElse m.BoundingArea > largest.BoundingArea Then
                second = largest
                largest = m
            ElseIf second Is Nothing OrElse m.BoundingArea > second.BoundingArea Then
                second = m
            End If

            maxAr = Math.Max(maxAr, m.AspectRatio)
            maxSlope = Math.Max(maxSlope, m.SlopeScore)
            maxProfile = Math.Max(maxProfile, m.ProfileComplexityScore)
            maxPlan = Math.Max(maxPlan, m.PlanComplexityScore)
            If m.CurveCount > 0 Then
                Dim arcRatio As Double = CDbl(m.ArcCount + m.CircleCount) / CDbl(Math.Max(1, m.CurveCount))
                maxArcRatio = Math.Max(maxArcRatio, arcRatio)
                If arcRatio >= 0.24 Then strongArcCount += 1
            End If
            If m.SlopeScore >= 0.24 Then strongSlopeCount += 1
        Next

        Dim domDiff As Double = 0
        If largest IsNot Nothing AndAlso second IsNot Nothing Then
            Dim a As Double = Math.Max(0.0001, largest.BoundingArea)
            Dim b As Double = Math.Max(0.0001, second.BoundingArea)
            domDiff = Math.Abs(a - b) / Math.Max(a, b)
        End If

        d.IsLong = (maxAr >= 3.9)
        d.IsThin = (maxAr >= 6.2)
        d.HasComplexProfile = (maxProfile >= 1.2)
        d.HasPlanTaper = (maxSlope >= 0.2 AndAlso maxPlan >= 0.95)
        d.HasDominantPlan = HasDominantByKey(measures, "TOP")
        d.HasDominantFacade = HasDominantByKey(measures, "FRONT") OrElse HasDominantByKey(measures, "BACK")
        d.HasDovetailEnds = (maxProfile >= 1.55 AndAlso maxSlope >= 0.2)
        d.HasDecorativeRecess = (maxPlan >= 1.25 AndAlso maxProfile >= 1.2)
        d.HasEdgeRadiusOrDrip = (maxArcRatio >= 0.16)
        d.HasBevelOrChamferEnds = (maxSlope >= 0.18 AndAlso maxProfile >= 1.08)
        d.HasSymmetricRadialBand = (maxArcRatio >= 0.28 AndAlso strongArcCount >= 2)
        d.HasProfileBulge = (maxProfile >= 1.5 AndAlso maxPlan >= 0.95)
        d.HasStrongThicknessView = (maxAr >= 4.6 OrElse (second IsNot Nothing AndAlso second.AspectRatio >= 4.0))
        d.HasRoundedDrip = d.HasEdgeRadiusOrDrip
        d.HasRebate = d.HasDecorativeRecess
        d.HasDovetailEnd = d.HasDovetailEnds OrElse d.HasBevelOrChamferEnds
        d.HasProfileShelf = d.HasComplexProfile OrElse d.HasProfileBulge
        d.HasMultipleRadialBands = d.HasSymmetricRadialBand

        Dim slopedCandidate As Boolean = (maxSlope >= 0.32 AndAlso domDiff >= 0.28 AndAlso strongSlopeCount >= 1)
        Dim radialCandidate As Boolean = (maxArcRatio >= 0.24 AndAlso (d.HasDominantPlan OrElse strongArcCount >= 2))

        d.HasSlope = slopedCandidate
        d.HasRadialPlan = radialCandidate

        If slopedCandidate Then
            d.Family = PartFamily.Sloped
        ElseIf radialCandidate Then
            d.Family = PartFamily.Radial
        ElseIf d.IsLong Then
            d.Family = PartFamily.Linear
        Else
            d.Family = PartFamily.Plate
        End If

        If d.Family = PartFamily.Sloped AndAlso d.HasDovetailEnd Then
            d.DimArchetype = DimensionArchetype.SlopedDovetail
        ElseIf d.Family = PartFamily.Radial Then
            If d.HasMultipleRadialBands Then
                d.DimArchetype = DimensionArchetype.RadialProfiled
            Else
                d.DimArchetype = DimensionArchetype.RadialSimple
            End If
        ElseIf d.Family = PartFamily.Linear Then
            If d.HasComplexProfile AndAlso (d.HasProfileShelf OrElse d.HasRoundedDrip) Then
                d.DimArchetype = DimensionArchetype.ProfiledStep
            ElseIf d.HasComplexProfile Then
                d.DimArchetype = DimensionArchetype.LinearProfiled
            Else
                d.DimArchetype = DimensionArchetype.LinearPlain
            End If
        Else
            d.DimArchetype = DimensionArchetype.PlateSimple
        End If

        d.PreferredMainRole = ResolvePreferredMainRole(d)
        Return d
    End Function

    Private Function ResolvePreferredMainRole(d As PartDescriptor) As ViewRole
        If d Is Nothing Then Return ViewRole.PlanContour
        Select Case d.Family
            Case PartFamily.Plate
                Return ViewRole.PlanContour
            Case PartFamily.Linear
                If d.HasProfileBulge AndAlso Not d.HasDominantFacade Then Return ViewRole.CrossProfile
                Return ViewRole.LongitudinalFacade
            Case PartFamily.Radial
                If d.HasComplexProfile OrElse d.HasProfileBulge Then Return ViewRole.CrossProfile
                Return ViewRole.PlanContour
            Case PartFamily.Sloped
                If d.HasSlope AndAlso (d.HasPlanTaper OrElse d.HasBevelOrChamferEnds) Then Return ViewRole.SlopeView
                Return ViewRole.PlanContour
            Case Else
                Return ViewRole.PlanContour
        End Select
    End Function

    Private Function HasDominantByKey(measures As List(Of ViewMeasure), key As String) As Boolean
        If measures Is Nothing Then Return False
        Dim maxArea As Double = 0
        Dim targetArea As Double = 0
        For Each m As ViewMeasure In measures
            If m Is Nothing Then Continue For
            maxArea = Math.Max(maxArea, m.BoundingArea)
            If String.Equals(m.Key, key, StringComparison.OrdinalIgnoreCase) Then
                targetArea = Math.Max(targetArea, m.BoundingArea)
            End If
        Next
        Return targetArea > 0 AndAlso targetArea >= maxArea * 0.82
    End Function

    Private Function ResolveRoles(descriptor As PartDescriptor,
                                  measures As List(Of ViewMeasure),
                                  isoMeasure As ViewMeasure) As RoleMap
        Dim map As New RoleMap()
        If measures Is Nothing Then Return map

        Dim orderedPairs As List(Of AuxPair) = BuildOrderedViewPairs(measures)

        Dim planCandidate As ViewMeasure = PickBestForRole(ViewRole.PlanContour, descriptor, measures, Nothing)
        Dim facadeCandidate As ViewMeasure = PickBestForRole(ViewRole.LongitudinalFacade, descriptor, measures, Nothing)
        Dim profileCandidate As ViewMeasure = PickBestForRole(ViewRole.CrossProfile, descriptor, measures, Nothing)
        Dim slopeCandidate As ViewMeasure = PickBestForRole(ViewRole.SlopeView, descriptor, measures, Nothing)
        Dim thickCandidate As ViewMeasure = PickBestForRole(ViewRole.ThicknessView, descriptor, measures, Nothing)
        Dim endCandidate As ViewMeasure = PickBestForRole(ViewRole.EndFace, descriptor, measures, Nothing)

        If slopeCandidate Is Nothing AndAlso descriptor IsNot Nothing AndAlso descriptor.Family = PartFamily.Sloped AndAlso orderedPairs.Count > 0 Then
            slopeCandidate = orderedPairs(0).A
        End If

        If thickCandidate Is Nothing Then
            thickCandidate = ChooseFallbackMeasure(profileCandidate, endCandidate)
        End If

        If descriptor IsNot Nothing Then
            If descriptor.Family = PartFamily.Sloped Then
                If planCandidate Is Nothing Then planCandidate = ChooseFallbackMeasure(facadeCandidate, profileCandidate)
                If slopeCandidate Is Nothing Then slopeCandidate = facadeCandidate
            ElseIf descriptor.Family = PartFamily.Radial Then
                If planCandidate Is Nothing Then planCandidate = ChooseFallbackMeasure(profileCandidate, facadeCandidate)
                If profileCandidate Is Nothing Then profileCandidate = ChooseFallbackMeasure(endCandidate, facadeCandidate)
            End If
        End If

        AssignRole(map, ViewRole.PlanContour, planCandidate)
        AssignRole(map, ViewRole.LongitudinalFacade, facadeCandidate)
        AssignRole(map, ViewRole.CrossProfile, profileCandidate)
        AssignRole(map, ViewRole.SlopeView, slopeCandidate)
        AssignRole(map, ViewRole.ThicknessView, thickCandidate)
        AssignRole(map, ViewRole.EndFace, endCandidate)

        Dim mainCandidate As ViewMeasure = PickMainContour(descriptor, map, measures, Nothing)
        Dim altMain As ViewMeasure = ChooseAlternativeMain(descriptor, map, mainCandidate)
        If mainCandidate IsNot Nothing AndAlso altMain IsNot Nothing Then
            Dim mainScore As Double = ScoreViewForRole(ViewRole.MainContour, descriptor, mainCandidate)
            Dim altScore As Double = ScoreViewForRole(ViewRole.MainContour, descriptor, altMain)
            If String.Equals(mainCandidate.Key, GetMeasureKeyByRole(map, descriptor), StringComparison.OrdinalIgnoreCase) AndAlso altScore >= mainScore * 0.95 Then
                mainCandidate = altMain
            End If
        End If

        AssignRole(map, ViewRole.MainContour, mainCandidate)
        AssignRole(map, ViewRole.IsoReference, isoMeasure)

        If map.GetMeasure(ViewRole.ThicknessView) Is Nothing Then
            AssignRole(map, ViewRole.ThicknessView, ChooseFallbackMeasure(map.GetMeasure(ViewRole.CrossProfile), map.GetMeasure(ViewRole.EndFace)))
        End If
        If descriptor IsNot Nothing AndAlso descriptor.Family = PartFamily.Sloped AndAlso map.GetMeasure(ViewRole.SlopeView) Is Nothing Then
            AssignRole(map, ViewRole.SlopeView, map.GetMeasure(ViewRole.LongitudinalFacade))
        End If

        Return map
    End Function

    Private Function ChooseFallbackMeasure(primary As ViewMeasure, secondary As ViewMeasure) As ViewMeasure
        If primary IsNot Nothing Then Return primary
        Return secondary
    End Function

    Private Function ForceIsometricMeasure(source As ViewMeasure, isoMeasure As ViewMeasure) As ViewMeasure
        If source Is Nothing Then Return Nothing
        Dim result As New ViewMeasure()
        result.Key = source.Key
        result.Caption = "Изометрия"
        result.Orientation = ViewOrientationTypeEnum.kIsoTopRightViewOrientation
        result.Style = DrawingViewStyleEnum.kShadedDrawingViewStyle

        If isoMeasure IsNot Nothing Then
            result.UnitW = isoMeasure.UnitW
            result.UnitH = isoMeasure.UnitH
            result.BoundingArea = isoMeasure.BoundingArea
            result.AspectRatio = isoMeasure.AspectRatio
            result.HorizontalBias = isoMeasure.HorizontalBias
            result.VerticalBias = isoMeasure.VerticalBias
            result.LongEdgeBias = isoMeasure.LongEdgeBias
            result.SlopeScore = isoMeasure.SlopeScore
            result.ProfileComplexityScore = isoMeasure.ProfileComplexityScore
            result.PlanComplexityScore = isoMeasure.PlanComplexityScore
            result.CurveCount = isoMeasure.CurveCount
            result.ArcCount = isoMeasure.ArcCount
            result.CircleCount = isoMeasure.CircleCount
            result.InnerContourCount = isoMeasure.InnerContourCount
            result.NonAxisEdgeCount = isoMeasure.NonAxisEdgeCount
        Else
            result.UnitW = source.UnitW
            result.UnitH = source.UnitH
            result.BoundingArea = source.BoundingArea
            result.AspectRatio = source.AspectRatio
            result.HorizontalBias = source.HorizontalBias
            result.VerticalBias = source.VerticalBias
            result.LongEdgeBias = source.LongEdgeBias
            result.SlopeScore = source.SlopeScore
            result.ProfileComplexityScore = source.ProfileComplexityScore
            result.PlanComplexityScore = source.PlanComplexityScore
            result.CurveCount = source.CurveCount
            result.ArcCount = source.ArcCount
            result.CircleCount = source.CircleCount
            result.InnerContourCount = source.InnerContourCount
            result.NonAxisEdgeCount = source.NonAxisEdgeCount
        End If

        Return result
    End Function

    Private Function GetMeasureKeyByRole(map As RoleMap, descriptor As PartDescriptor) As String
        If map Is Nothing OrElse descriptor Is Nothing Then Return String.Empty
        Dim r As ViewRole = descriptor.PreferredMainRole
        Dim m As ViewMeasure = map.GetMeasure(r)
        If m Is Nothing Then Return String.Empty
        Return m.Key
    End Function

    Private Function ChooseAlternativeMain(descriptor As PartDescriptor, map As RoleMap, currentMain As ViewMeasure) As ViewMeasure
        If map Is Nothing Then Return Nothing
        Dim pref As ViewRole = ViewRole.PlanContour
        If descriptor IsNot Nothing Then pref = descriptor.PreferredMainRole
        Dim candidate As ViewMeasure = map.GetMeasure(pref)
        If candidate Is Nothing OrElse currentMain Is Nothing Then Return candidate
        If String.Equals(candidate.Key, currentMain.Key, StringComparison.OrdinalIgnoreCase) Then
            If pref = ViewRole.PlanContour Then Return map.GetMeasure(ViewRole.CrossProfile)
            Return map.GetMeasure(ViewRole.PlanContour)
        End If
        Return candidate
    End Function

    Private Sub AssignRole(map As RoleMap, role As ViewRole, m As ViewMeasure)
        If map Is Nothing OrElse m Is Nothing Then Return
        map.ByRole(role) = m
    End Sub

    Private Function PickMainContour(descriptor As PartDescriptor,
                                     map As RoleMap,
                                     measures As List(Of ViewMeasure),
                                     used As HashSet(Of String)) As ViewMeasure
        If descriptor IsNot Nothing Then
            Dim preferred As ViewMeasure = map.GetMeasure(descriptor.PreferredMainRole)
            If preferred IsNot Nothing Then Return preferred

            If descriptor.Family = PartFamily.Sloped AndAlso map.ByRole.ContainsKey(ViewRole.SlopeView) Then Return map.ByRole(ViewRole.SlopeView)
            If (descriptor.Family = PartFamily.Radial OrElse descriptor.Family = PartFamily.Plate) AndAlso map.ByRole.ContainsKey(ViewRole.PlanContour) Then Return map.ByRole(ViewRole.PlanContour)
            If descriptor.Family = PartFamily.Linear AndAlso map.ByRole.ContainsKey(ViewRole.LongitudinalFacade) Then Return map.ByRole(ViewRole.LongitudinalFacade)
        End If
        Return PickBestForRole(ViewRole.MainContour, descriptor, measures, used)
    End Function

    Private Function PickBestForRole(role As ViewRole,
                                     descriptor As PartDescriptor,
                                     measures As List(Of ViewMeasure),
                                     used As HashSet(Of String)) As ViewMeasure
        Dim best As ViewMeasure = Nothing
        Dim bestScore As Double = Double.MinValue
        For Each m As ViewMeasure In measures
            If m Is Nothing Then Continue For
            Dim score As Double = ScoreViewForRole(role, descriptor, m)
            If used IsNot Nothing AndAlso used.Contains(m.Key) Then score -= 0.22
            If best Is Nothing OrElse score > bestScore Then
                best = m
                bestScore = score
            End If
        Next
        If best IsNot Nothing AndAlso used IsNot Nothing Then used.Add(best.Key)
        Return best
    End Function

    Private Function ScoreViewForRole(role As ViewRole, descriptor As PartDescriptor, m As ViewMeasure) As Double
        Dim s As Double = m.BoundingArea * 0.08 + m.CurveCount * 0.02
        Select Case role
            Case ViewRole.PlanContour
                If String.Equals(m.Key, "TOP", StringComparison.OrdinalIgnoreCase) Then s += 1.1
                s += m.PlanComplexityScore * 0.8
                s -= m.SlopeScore * 0.25
            Case ViewRole.LongitudinalFacade
                If String.Equals(m.Key, "FRONT", StringComparison.OrdinalIgnoreCase) OrElse String.Equals(m.Key, "BACK", StringComparison.OrdinalIgnoreCase) Then s += 0.9
                s += m.VerticalBias * 0.55
                s += Math.Min(2.0, m.AspectRatio) * 0.35
            Case ViewRole.CrossProfile
                If String.Equals(m.Key, "LEFT", StringComparison.OrdinalIgnoreCase) OrElse String.Equals(m.Key, "RIGHT", StringComparison.OrdinalIgnoreCase) Then s += 0.9
                s += m.ProfileComplexityScore * 0.75
                s += m.InnerContourCount * 0.06
            Case ViewRole.SlopeView
                s += m.SlopeScore * 1.25
                s += m.NonAxisEdgeCount * 0.06
            Case ViewRole.ThicknessView
                s += m.LongEdgeBias * 0.45
                If m.AspectRatio >= 4.5 Then s += 0.35
            Case ViewRole.EndFace
                s += m.ProfileComplexityScore * 0.45
                s += m.CircleCount * 0.1
            Case Else
                s += m.ProfileComplexityScore * 0.2
        End Select

        If descriptor IsNot Nothing Then
            If descriptor.Family = PartFamily.Radial Then s += m.ArcCount * 0.04
            If descriptor.Family = PartFamily.Sloped Then s += m.SlopeScore * 0.4
            If descriptor.Family = PartFamily.Linear AndAlso role = ViewRole.LongitudinalFacade Then s += m.AspectRatio * 0.12
        End If
        Return s
    End Function

    Private Function BuildLayoutTemplates(doc As DrawingDocument,
                                          safe As SlotRect,
                                          gap As Double,
                                          descriptor As PartDescriptor) As List(Of LayoutTemplate)
        Dim result As New List(Of LayoutTemplate)()
        Dim w As Double = RectW(safe)
        Dim h As Double = RectH(safe)
        Dim leftW As Double = safe.L + w * 0.32
        Dim rightL As Double = safe.L + w * 0.36

        Dim baseMain As SlotRect = InsetRect(New SlotRect(safe.L, leftW, safe.B, safe.T), gap * 0.18)
        Dim baseAuxTop As SlotRect = InsetRect(New SlotRect(rightL, safe.R, safe.B + h * 0.54, safe.T), gap * 0.18)
        Dim baseIso As SlotRect = InsetRect(New SlotRect(rightL + w * 0.08, safe.R, safe.B, safe.B + h * 0.46), gap * 0.18)

        result.Add(NewLayoutTemplate("LINEAR_FACADE_LEFT_PROFILE_TOP_ISO", PartFamily.Linear, ViewRole.LongitudinalFacade, ViewRole.CrossProfile, ViewRole.IsoReference, baseMain, baseAuxTop, baseIso))
        result.Add(NewLayoutTemplate("LINEAR_PROFILE_MAIN_FACADE_LEFT_ISO", PartFamily.Linear, ViewRole.CrossProfile, ViewRole.LongitudinalFacade, ViewRole.IsoReference,
                                     InsetRect(New SlotRect(safe.L, safe.L + w * 0.55, safe.B + h * 0.18, safe.T), gap * 0.18),
                                     InsetRect(New SlotRect(safe.L + w * 0.58, safe.R, safe.B + h * 0.52, safe.T), gap * 0.18),
                                     InsetRect(New SlotRect(safe.L + w * 0.62, safe.R, safe.B, safe.B + h * 0.46), gap * 0.18)))
        result.Add(NewLayoutTemplate("LINEAR_PLAN_MAIN_SIDE_LEFT_ISO", PartFamily.Linear, ViewRole.PlanContour, ViewRole.ThicknessView, ViewRole.IsoReference,
                                     InsetRect(New SlotRect(safe.L, safe.L + w * 0.62, safe.B + h * 0.34, safe.T), gap * 0.18),
                                     InsetRect(New SlotRect(safe.L + w * 0.65, safe.R, safe.B + h * 0.34, safe.T), gap * 0.18),
                                     InsetRect(New SlotRect(safe.L + w * 0.52, safe.R, safe.B, safe.B + h * 0.30), gap * 0.18)))

        result.Add(NewLayoutTemplate("PLATE_PLAN_MAIN_EDGE_LEFT_ISO", PartFamily.Plate, ViewRole.PlanContour, ViewRole.ThicknessView, ViewRole.IsoReference,
                                     InsetRect(New SlotRect(safe.L, safe.L + w * 0.64, safe.B + h * 0.32, safe.T), gap * 0.18),
                                     InsetRect(New SlotRect(safe.L + w * 0.66, safe.R, safe.B + h * 0.32, safe.T), gap * 0.18),
                                     InsetRect(New SlotRect(safe.L + w * 0.58, safe.R, safe.B, safe.B + h * 0.28), gap * 0.18)))
        result.Add(NewLayoutTemplate("PLATE_PLAN_MAIN_PROFILE_LEFT_ISO", PartFamily.Plate, ViewRole.PlanContour, ViewRole.CrossProfile, ViewRole.IsoReference,
                                     InsetRect(New SlotRect(safe.L, safe.L + w * 0.58, safe.B + h * 0.30, safe.T), gap * 0.18),
                                     InsetRect(New SlotRect(safe.L + w * 0.60, safe.R, safe.B + h * 0.50, safe.T), gap * 0.18),
                                     InsetRect(New SlotRect(safe.L + w * 0.64, safe.R, safe.B, safe.B + h * 0.45), gap * 0.18)))

        result.Add(NewLayoutTemplate("RADIAL_PLAN_MAIN_PROFILE_SIDE_ISO", PartFamily.Radial, ViewRole.PlanContour, ViewRole.CrossProfile, ViewRole.IsoReference,
                                     InsetRect(New SlotRect(safe.L, safe.L + w * 0.52, safe.B + h * 0.26, safe.T), gap * 0.18),
                                     InsetRect(New SlotRect(safe.L + w * 0.56, safe.R, safe.B + h * 0.46, safe.T), gap * 0.18),
                                     InsetRect(New SlotRect(safe.L + w * 0.60, safe.R, safe.B, safe.B + h * 0.42), gap * 0.18)))
        result.Add(NewLayoutTemplate("RADIAL_PROFILE_HEAVY", PartFamily.Radial, ViewRole.CrossProfile, ViewRole.PlanContour, ViewRole.IsoReference,
                                     InsetRect(New SlotRect(safe.L, safe.L + w * 0.57, safe.B + h * 0.18, safe.T), gap * 0.18),
                                     InsetRect(New SlotRect(safe.L + w * 0.60, safe.R, safe.B + h * 0.56, safe.T), gap * 0.18),
                                     InsetRect(New SlotRect(safe.L + w * 0.64, safe.R, safe.B, safe.B + h * 0.50), gap * 0.18)))
        result.Add(NewLayoutTemplate("RADIAL_PLAN_TOP_PROFILE_BOTTOMLEFT", PartFamily.Radial, ViewRole.PlanContour, ViewRole.CrossProfile, ViewRole.IsoReference,
                                     InsetRect(New SlotRect(safe.L + w * 0.35, safe.R, safe.B + h * 0.35, safe.T), gap * 0.18),
                                     InsetRect(New SlotRect(safe.L, safe.L + w * 0.62, safe.B, safe.B + h * 0.34), gap * 0.18),
                                     InsetRect(New SlotRect(safe.L, safe.L + w * 0.30, safe.B + h * 0.58, safe.T), gap * 0.18)))

        result.Add(NewLayoutTemplate("SLOPED_PLAN_COMPLEX", PartFamily.Sloped, ViewRole.PlanContour, ViewRole.SlopeView, ViewRole.IsoReference,
                                     InsetRect(New SlotRect(safe.L, safe.L + w * 0.60, safe.B + h * 0.28, safe.T), gap * 0.18),
                                     InsetRect(New SlotRect(safe.L + w * 0.63, safe.R, safe.B + h * 0.52, safe.T), gap * 0.18),
                                     InsetRect(New SlotRect(safe.L + w * 0.60, safe.R, safe.B, safe.B + h * 0.48), gap * 0.18)))
        result.Add(NewLayoutTemplate("SLOPED_FACADE_MAIN", PartFamily.Sloped, ViewRole.SlopeView, ViewRole.PlanContour, ViewRole.IsoReference,
                                     InsetRect(New SlotRect(safe.L, safe.L + w * 0.38, safe.B, safe.T), gap * 0.18),
                                     InsetRect(New SlotRect(safe.L + w * 0.42, safe.R, safe.B + h * 0.56, safe.T), gap * 0.18),
                                     InsetRect(New SlotRect(safe.L + w * 0.52, safe.R, safe.B, safe.B + h * 0.50), gap * 0.18)))
        result.Add(NewLayoutTemplate("SLOPED_WEDGE_LEFT", PartFamily.Sloped, ViewRole.SlopeView, ViewRole.PlanContour, ViewRole.IsoReference,
                                     InsetRect(New SlotRect(safe.L, safe.L + w * 0.42, safe.B, safe.T), gap * 0.16),
                                     InsetRect(New SlotRect(safe.L + w * 0.45, safe.R, safe.B + h * 0.52, safe.T), gap * 0.16),
                                     InsetRect(New SlotRect(safe.L + w * 0.58, safe.R, safe.B, safe.B + h * 0.48), gap * 0.16)))
        result.Add(NewLayoutTemplate("SLOPED_PLAN_WIDE", PartFamily.Sloped, ViewRole.PlanContour, ViewRole.SlopeView, ViewRole.IsoReference,
                                     InsetRect(New SlotRect(safe.L, safe.L + w * 0.68, safe.B + h * 0.38, safe.T), gap * 0.16),
                                     InsetRect(New SlotRect(safe.L + w * 0.70, safe.R, safe.B + h * 0.22, safe.T), gap * 0.16),
                                     InsetRect(New SlotRect(safe.L + w * 0.66, safe.R, safe.B, safe.B + h * 0.20), gap * 0.16)))

        If descriptor Is Nothing Then Return result
        Dim filtered As New List(Of LayoutTemplate)()
        For Each t As LayoutTemplate In result
            If t.Family = descriptor.Family Then filtered.Add(t)
        Next
        If filtered.Count > 0 Then Return filtered
        Return result
    End Function

    Private Function NewLayoutTemplate(name As String, fam As PartFamily,
                                       mainRole As ViewRole, auxRole As ViewRole, isoRole As ViewRole,
                                       mainSlot As SlotRect, auxSlot As SlotRect, isoSlot As SlotRect) As LayoutTemplate
        Dim t As New LayoutTemplate()
        t.TemplateName = name
        t.Family = fam
        t.MainRole = mainRole
        t.AuxRole = auxRole
        t.IsoRole = isoRole
        t.RequiredRoles.Add(mainRole)
        t.RequiredRoles.Add(auxRole)
        t.RequiredRoles.Add(isoRole)
        t.MainSlot = mainSlot
        t.AuxSlot = auxSlot
        t.IsoSlot = isoSlot
        Return t
    End Function

    Private Function EvaluateBestTemplate(templates As List(Of LayoutTemplate),
                                          descriptor As PartDescriptor,
                                          roleMap As RoleMap) As LayoutPlan
        Dim best As LayoutPlan = Nothing
        For Each t As LayoutTemplate In templates
            Dim p As LayoutPlan = EvaluateTemplate(t, descriptor, roleMap)
            If p Is Nothing Then Continue For
            If best Is Nothing OrElse p.Score > best.Score Then best = p
        Next
        Return best
    End Function

    Private Function EvaluateTemplate(t As LayoutTemplate,
                                      descriptor As PartDescriptor,
                                      roleMap As RoleMap) As LayoutPlan
        If t Is Nothing OrElse roleMap Is Nothing Then Return Nothing
        Dim mainM As ViewMeasure = roleMap.GetMeasure(t.MainRole)
        Dim auxM As ViewMeasure = roleMap.GetMeasure(t.AuxRole)
        Dim isoM As ViewMeasure = roleMap.GetMeasure(t.IsoRole)
        If mainM Is Nothing OrElse auxM Is Nothing Then Return Nothing

        Dim p As New LayoutPlan()
        p.TemplateName = t.TemplateName
        p.MainRole = t.MainRole
        p.AuxRole = t.AuxRole
        p.IsoRole = t.IsoRole
        p.MainSlot = t.MainSlot
        p.Aux1Slot = t.AuxSlot
        p.Aux2Slot = t.IsoSlot
        p.MainMeasure = mainM
        p.Aux1Measure = auxM
        p.Aux2Measure = isoM

        p.MainFit = ScaleToFit(p.MainSlot, mainM, ORTHO_SCALE_MARGIN)
        p.Aux1Fit = ScaleToFit(p.Aux1Slot, auxM, ORTHO_SCALE_MARGIN)
        If isoM IsNot Nothing Then
            p.Aux2Fit = ScaleToFit(p.Aux2Slot, isoM, ISO_SCALE_MARGIN)
        End If
        If p.MainFit Is Nothing OrElse p.Aux1Fit Is Nothing Then Return Nothing
        If isoM IsNot Nothing AndAlso p.Aux2Fit Is Nothing Then Return Nothing

        Dim mainFill As Double = (p.MainFit.ProjectedW * p.MainFit.ProjectedH) / Math.Max(0.0001, RectW(p.MainSlot) * RectH(p.MainSlot))
        Dim auxFill As Double = (p.Aux1Fit.ProjectedW * p.Aux1Fit.ProjectedH) / Math.Max(0.0001, RectW(p.Aux1Slot) * RectH(p.Aux1Slot))
        Dim complement As Double = Math.Abs(mainM.AspectRatio - auxM.AspectRatio)
        p.Score = mainFill * 0.55 + auxFill * 0.25 + Math.Min(1.0, complement / 3.0) * 0.2

        If descriptor IsNot Nothing Then
            If descriptor.Family = PartFamily.Sloped Then
                If t.MainRole = ViewRole.SlopeView Then p.Score += 0.18
                If t.AuxRole = ViewRole.PlanContour AndAlso t.MainRole = ViewRole.SlopeView Then p.Score += 0.12
                If Math.Abs(mainM.AspectRatio - auxM.AspectRatio) < 0.22 Then p.Score -= 0.14
            End If
            If descriptor.Family = PartFamily.Linear AndAlso t.MainRole = ViewRole.LongitudinalFacade Then p.Score += 0.12
            If descriptor.Family = PartFamily.Radial Then
                If String.Equals(t.TemplateName, "RADIAL_PROFILE_HEAVY", StringComparison.OrdinalIgnoreCase) AndAlso descriptor.HasComplexProfile Then p.Score += 0.24
                If String.Equals(t.TemplateName, "RADIAL_PLAN_TOP_PROFILE_BOTTOMLEFT", StringComparison.OrdinalIgnoreCase) AndAlso descriptor.HasProfileBulge Then p.Score += 0.18
                If t.AuxRole = ViewRole.CrossProfile AndAlso RectW(t.AuxSlot) >= RectW(t.MainSlot) * 0.55 Then p.Score += 0.1
            End If
            If descriptor.Family = PartFamily.Plate Then
                Dim orthogonalClarity As Double = (mainM.HorizontalBias + auxM.VerticalBias) * 0.22
                p.Score += orthogonalClarity
            End If
        End If
        Return p
    End Function

    ' ================================================================
    ' v3.20 FIX: PlaceViewsByTemplate now adds semantic role aliases
    ' so that ApplyDimensionPlan can find PlanContour, ThicknessView, etc.
    ' ================================================================
    Private Function PlaceViewsByTemplate(sheet As Sheet, modelDoc As Document,
                                          plan As LayoutPlan,
                                          roleMap As RoleMap) As Dictionary(Of ViewRole, DrawingView)
        Dim placed As New Dictionary(Of ViewRole, DrawingView)()
        If plan Is Nothing Then Return placed

        Dim mainV As DrawingView = PlaceViewInSlot(sheet, modelDoc, plan.MainMeasure, plan.MainFit, plan.MainSlot)
        Dim auxV As DrawingView = PlaceViewInSlot(sheet, modelDoc, plan.Aux1Measure, plan.Aux1Fit, plan.Aux1Slot)
        Dim isoV As DrawingView = Nothing
        If plan.Aux2Measure IsNot Nothing AndAlso plan.Aux2Fit IsNot Nothing Then
            isoV = PlaceViewInSlot(sheet, modelDoc, plan.Aux2Measure, plan.Aux2Fit, plan.Aux2Slot)
        End If

        If mainV Is Nothing OrElse auxV Is Nothing Then Return Nothing

        ' Store by template role
        placed(plan.MainRole) = mainV
        placed(plan.AuxRole) = auxV
        If isoV IsNot Nothing Then
            placed(plan.IsoRole) = isoV
        End If

        ' v3.20: Also register all semantic aliases from roleMap so that
        ' ApplyDimensionPlan can find e.g. PlanContour, ThicknessView, SlopeView
        ' even when the template used LongitudinalFacade/CrossProfile/IsoReference keys.
        If roleMap IsNot Nothing Then
            For Each kv As KeyValuePair(Of ViewRole, ViewMeasure) In roleMap.ByRole
                If kv.Value Is Nothing Then Continue For
                If kv.Key = plan.MainRole OrElse kv.Key = plan.AuxRole OrElse kv.Key = plan.IsoRole Then Continue For
                If kv.Key = ViewRole.IsoReference Then Continue For
                ' Find which placed view has the same measure (by Key)
                Dim targetView As DrawingView = Nothing
                If kv.Value.Key = plan.MainMeasure.Key Then
                    targetView = mainV
                ElseIf kv.Value.Key = plan.Aux1Measure.Key Then
                    targetView = auxV
                ElseIf plan.Aux2Measure IsNot Nothing AndAlso isoV IsNot Nothing AndAlso kv.Value.Key = plan.Aux2Measure.Key Then
                    targetView = isoV
                End If
                If targetView IsNot Nothing AndAlso Not placed.ContainsKey(kv.Key) Then
                    placed(kv.Key) = targetView
                    Debug.Print("v3.20 alias: " & kv.Key.ToString() & " => " & kv.Value.Key)
                End If
            Next
        End If

        Return placed
    End Function

    Private Sub AddViewRoleCaptions(doc As DrawingDocument,
                                    sheet As Sheet,
                                    plan As LayoutPlan,
                                    placed As Dictionary(Of ViewRole, DrawingView))
        If plan Is Nothing OrElse placed Is Nothing Then Return
        Try
            If placed.ContainsKey(plan.MainRole) Then AddViewAnnotations(doc, sheet, placed(plan.MainRole), plan.MainSlot, plan.MainMeasure, False)
            If placed.ContainsKey(plan.AuxRole) Then AddViewAnnotations(doc, sheet, placed(plan.AuxRole), plan.Aux1Slot, plan.Aux1Measure, False)
            If placed.ContainsKey(plan.IsoRole) Then AddViewAnnotations(doc, sheet, placed(plan.IsoRole), plan.Aux2Slot, plan.Aux2Measure, False)
        Catch ex As Exception
            Debug.Print("WARN AddViewRoleCaptions: " & ex.Message)
        End Try
    End Sub

    Private Function BuildDimensionPlan(descriptor As PartDescriptor,
                                        roleMap As RoleMap) As DimensionPlan
        Dim plan As New DimensionPlan()
        If descriptor Is Nothing Then Return plan

        Select Case descriptor.DimArchetype
            Case DimensionArchetype.PlateSimple
                plan.Intents.Add(NewIntent(DimensionIntentId.OverallLength, ViewRole.PlanContour, 0, True))
                plan.Intents.Add(NewIntent(DimensionIntentId.OverallWidth, ViewRole.PlanContour, 0, True))
                plan.Intents.Add(NewIntent(DimensionIntentId.OverallThickness, ViewRole.ThicknessView, 0, True))
                If descriptor.HasBevelOrChamferEnds Then plan.Intents.Add(NewIntent(DimensionIntentId.Chamfer, ViewRole.ThicknessView, 2, True))

            Case DimensionArchetype.LinearPlain
                plan.Intents.Add(NewIntent(DimensionIntentId.OverallLength, ViewRole.LongitudinalFacade, 0, True))
                plan.Intents.Add(NewIntent(DimensionIntentId.OverallHeight, ViewRole.LongitudinalFacade, 0, True))
                plan.Intents.Add(NewIntent(DimensionIntentId.OverallThickness, ViewRole.CrossProfile, 0, True))

            Case DimensionArchetype.LinearProfiled
                plan.Intents.Add(NewIntent(DimensionIntentId.OverallLength, ViewRole.LongitudinalFacade, 0, True))
                plan.Intents.Add(NewIntent(DimensionIntentId.OverallHeight, ViewRole.LongitudinalFacade, 0, True))
                plan.Intents.Add(NewIntent(DimensionIntentId.OverallThickness, ViewRole.CrossProfile, 0, True))
                plan.Intents.Add(NewIntent(DimensionIntentId.ProfileHeight, ViewRole.CrossProfile, 1, True))
                plan.Intents.Add(NewIntent(DimensionIntentId.ProfileDepth, ViewRole.CrossProfile, 1, True))
                plan.Intents.Add(NewIntent(DimensionIntentId.LipDepth, ViewRole.CrossProfile, 2, True))
                plan.Intents.Add(NewIntent(DimensionIntentId.LipHeight, ViewRole.CrossProfile, 2, True))

            Case DimensionArchetype.ProfiledStep
                plan.Intents.Add(NewIntent(DimensionIntentId.OverallLength, ViewRole.LongitudinalFacade, 0, True))
                plan.Intents.Add(NewIntent(DimensionIntentId.OverallWidth, ViewRole.PlanContour, 0, True))
                plan.Intents.Add(NewIntent(DimensionIntentId.OverallHeight, ViewRole.LongitudinalFacade, 0, True))
                plan.Intents.Add(NewIntent(DimensionIntentId.ProfileHeight, ViewRole.CrossProfile, 1, True))
                plan.Intents.Add(NewIntent(DimensionIntentId.ProfileDepth, ViewRole.CrossProfile, 1, True))
                plan.Intents.Add(NewIntent(DimensionIntentId.StepHeight, ViewRole.CrossProfile, 1, True))
                plan.Intents.Add(NewIntent(DimensionIntentId.LipDepth, ViewRole.CrossProfile, 2, True))
                plan.Intents.Add(NewIntent(DimensionIntentId.LipHeight, ViewRole.CrossProfile, 2, True))
                If descriptor.HasRoundedDrip Then
                    plan.Intents.Add(NewIntent(DimensionIntentId.VisibleRadius, ViewRole.CrossProfile, 2, True))
                ElseIf descriptor.HasBevelOrChamferEnds Then
                    plan.Intents.Add(NewIntent(DimensionIntentId.Chamfer, ViewRole.CrossProfile, 2, True))
                End If

            Case DimensionArchetype.RadialSimple
                plan.Intents.Add(NewIntent(DimensionIntentId.ChordOrSpan, ViewRole.PlanContour, 0, True))
                plan.Intents.Add(NewIntent(DimensionIntentId.RadiusMain, ViewRole.PlanContour, 0, True))
                plan.Intents.Add(NewIntent(DimensionIntentId.OverallThickness, ViewRole.CrossProfile, 0, True))

            Case DimensionArchetype.RadialProfiled
                plan.Intents.Add(NewIntent(DimensionIntentId.ChordOrSpan, ViewRole.PlanContour, 0, True))
                plan.Intents.Add(NewIntent(DimensionIntentId.RadiusMain, ViewRole.PlanContour, 0, True))
                If descriptor.HasMultipleRadialBands Then plan.Intents.Add(NewIntent(DimensionIntentId.RadiusSecondary, ViewRole.PlanContour, 1, True))
                plan.Intents.Add(NewIntent(DimensionIntentId.EdgeBandWidth, ViewRole.PlanContour, 1, True))
                plan.Intents.Add(NewIntent(DimensionIntentId.OverallThickness, ViewRole.CrossProfile, 0, True))
                plan.Intents.Add(NewIntent(DimensionIntentId.ProfileHeight, ViewRole.CrossProfile, 1, True))
                plan.Intents.Add(NewIntent(DimensionIntentId.ProfileDepth, ViewRole.CrossProfile, 1, True))
                plan.Intents.Add(NewIntent(DimensionIntentId.VisibleRadius, ViewRole.CrossProfile, 2, True))

            Case DimensionArchetype.SlopedDovetail
                plan.Intents.Add(NewIntent(DimensionIntentId.OverallLength, ViewRole.SlopeView, 0, True))
                plan.Intents.Add(NewIntent(DimensionIntentId.OverallWidth, ViewRole.PlanContour, 0, True))
                plan.Intents.Add(NewIntent(DimensionIntentId.OverallThickness, ViewRole.ThicknessView, 1, True))
                plan.Intents.Add(NewIntent(DimensionIntentId.SlopeHeightHigh, ViewRole.SlopeView, 0, True))
                plan.Intents.Add(NewIntent(DimensionIntentId.SlopeHeightLow, ViewRole.SlopeView, 0, True))
                plan.Intents.Add(NewIntent(DimensionIntentId.EndCutLength, ViewRole.SlopeView, 2, True))
                plan.Intents.Add(NewIntent(DimensionIntentId.Chamfer, ViewRole.SlopeView, 2, True))
        End Select
        Return plan
    End Function

    Private Function NewIntent(id As DimensionIntentId,
                               role As ViewRole,
                               priority As Integer,
                               allowFallback As Boolean) As DimensionIntent
        Dim i As New DimensionIntent()
        i.IntentId = id
        i.PreferredRole = role
        i.Priority = priority
        i.AllowFallbackNote = allowFallback
        Return i
    End Function

    Private Sub ApplyDimensionPlan(doc As DrawingDocument,
                                   sheet As Sheet,
                                   placedViews As Dictionary(Of ViewRole, DrawingView),
                                   roleMap As RoleMap,
                                   descriptor As PartDescriptor,
                                   plan As DimensionPlan,
                                   modelSize As ModelOverallExtents)
        If plan Is Nothing OrElse placedViews Is Nothing Then Return
        Dim placedIntents As New HashSet(Of DimensionIntentId)()
        Dim viewOverallCount As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
        Dim viewFeatureCount As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
        Dim viewIntentKeys As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim fallbackNoteKeys As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim globalDimensionKeys As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim realDimCount As Integer = 0
        Dim totalAdded As Integer = 0
        Dim pendingFallback As New List(Of FallbackRequest)()

        plan.Intents.Sort(Function(a As DimensionIntent, b As DimensionIntent) a.Priority.CompareTo(b.Priority))
        For Each intent As DimensionIntent In plan.Intents
            If placedIntents.Contains(intent.IntentId) Then Continue For
            If intent.PreferredRole = ViewRole.IsoReference Then Continue For

            Dim targetRole As ViewRole = ResolveExistingRole(intent.PreferredRole, placedViews, intent.IntentId)
            Dim v As DrawingView = Nothing
            If placedViews.ContainsKey(targetRole) Then v = placedViews(targetRole)
            If v Is Nothing Then
                Debug.Print("WARN ApplyDimensionPlan: view not found for role=" & targetRole.ToString() & " (intent=" & intent.IntentId.ToString() & ")")
                If intent.AllowFallbackNote Then
                    ' Try any available non-iso view as last resort
                    For Each kvp As KeyValuePair(Of ViewRole, DrawingView) In placedViews
                        If kvp.Key <> ViewRole.IsoReference AndAlso kvp.Value IsNot Nothing Then
                            v = kvp.Value
                            targetRole = kvp.Key
                            Debug.Print("  -> fallback to role=" & targetRole.ToString())
                            Exit For
                        End If
                    Next
                End If
                If v Is Nothing Then Continue For
            End If

            Dim slot As SlotRect = New SlotRect(v.Left, v.Left + v.Width, v.Top - v.Height, v.Top)
            Dim viewKey As String = targetRole.ToString()
            If IsOverallIntent(intent.IntentId) Then
                If GetCounter(viewOverallCount, viewKey) >= 3 Then Continue For
            Else
                If GetCounter(viewFeatureCount, viewKey) >= 3 Then Continue For
            End If

            Dim added As Integer = 0
            If IsOverallIntent(intent.IntentId) Then
                added = TryAddOverallDimensions(doc, sheet, v, slot, intent.IntentId, targetRole, roleMap.GetMeasure(targetRole), modelSize, viewIntentKeys, fallbackNoteKeys, viewOverallCount, viewKey, globalDimensionKeys)
            ElseIf IsProfileIntent(intent.IntentId) Then
                added = TryAddProfileDimensions(doc, sheet, v, slot, intent.IntentId, roleMap.GetMeasure(targetRole), viewIntentKeys, fallbackNoteKeys, viewFeatureCount, viewKey, globalDimensionKeys)
            Else
                added = TryAddFeatureDimensions(doc, sheet, v, slot, intent.IntentId, roleMap.GetMeasure(targetRole), viewIntentKeys, fallbackNoteKeys, viewFeatureCount, viewKey, globalDimensionKeys)
            End If

            If added > 0 Then
                placedIntents.Add(intent.IntentId)
                totalAdded += added
                If IsOverallIntent(intent.IntentId) AndAlso modelSize IsNot Nothing AndAlso modelSize.IsValid Then
                    Debug.Print("Dimension intent " & intent.IntentId.ToString() & " on " & targetRole.ToString() & ": model-driven overall note added (" & added.ToString() & ")")
                Else
                    realDimCount += added
                    Debug.Print("Dimension intent " & intent.IntentId.ToString() & " on " & targetRole.ToString() & ": real dimension added (" & added.ToString() & ")")
                End If
            ElseIf intent.AllowFallbackNote Then
                Dim req As New FallbackRequest()
                req.Intent = intent.IntentId
                req.Role = targetRole
                req.View = v
                req.Slot = slot
                req.Measure = roleMap.GetMeasure(targetRole)
                pendingFallback.Add(req)
                Debug.Print("Dimension intent " & intent.IntentId.ToString() & " on " & targetRole.ToString() & ": deferred fallback")
            Else
                Debug.Print("Dimension intent " & intent.IntentId.ToString() & " on " & targetRole.ToString() & ": failed completely")
            End If
        Next

        Dim keyViews As New List(Of DrawingView)()
        If placedViews.ContainsKey(ViewRole.CrossProfile) Then keyViews.Add(placedViews(ViewRole.CrossProfile))
        If placedViews.ContainsKey(ViewRole.SlopeView) Then keyViews.Add(placedViews(ViewRole.SlopeView))
        If keyViews.Count = 1 AndAlso placedViews.ContainsKey(ViewRole.LongitudinalFacade) Then keyViews.Add(placedViews(ViewRole.LongitudinalFacade))

        Dim requireFive As Boolean = (descriptor IsNot Nothing AndAlso (descriptor.DimArchetype = DimensionArchetype.ProfiledStep OrElse descriptor.DimArchetype = DimensionArchetype.RadialProfiled))
        If requireFive Then
            For Each keyView As DrawingView In keyViews
                If realDimCount >= 5 Then Exit For
                Dim slot As SlotRect = New SlotRect(keyView.Left, keyView.Left + keyView.Width, keyView.Top - keyView.Height, keyView.Top)
                Dim supplement As Integer = TryAddTrueDimensions(doc, sheet, keyView, slot, True, True, True, True, 5 - realDimCount, Nothing, "", False, False)
                realDimCount += supplement
            Next
        End If

        For Each req As FallbackRequest In pendingFallback
            If placedIntents.Contains(req.Intent) Then Continue For
            If AddFallbackDimensionNotes(doc, sheet, req.View, req.Slot, req.Intent, req.Role, req.Measure, modelSize, fallbackNoteKeys) Then
                placedIntents.Add(req.Intent)
                totalAdded += 1
                Debug.Print("Dimension intent " & req.Intent.ToString() & " on " & req.Role.ToString() & ": fallback note added")
            Else
                Debug.Print("Dimension intent " & req.Intent.ToString() & " on " & req.Role.ToString() & ": failed completely")
            End If
        Next

        If totalAdded = 0 Then
            Dim forced As Integer = ForceVisibleFallbackDimensions(doc, sheet, placedViews, roleMap, plan, modelSize)
            totalAdded += forced
            Debug.Print("ForceVisibleFallbackDimensions: added=" & forced.ToString())
        End If
    End Sub

    ' ================================================================
    ' v3.20: Extended ResolveExistingRole — comprehensive role fallback
    ' ================================================================
    Private Function ResolveExistingRole(preferred As ViewRole, placedViews As Dictionary(Of ViewRole, DrawingView), Optional intent As DimensionIntentId = DimensionIntentId.OverallLength) As ViewRole
        If placedViews Is Nothing Then Return preferred
        If placedViews.ContainsKey(preferred) Then Return preferred

        ' Thickness role must stay isolated from facade/plan to avoid duplicate overall size values
        If preferred = ViewRole.ThicknessView Then
            If placedViews.ContainsKey(ViewRole.CrossProfile) Then Return ViewRole.CrossProfile
            If placedViews.ContainsKey(ViewRole.EndFace) Then Return ViewRole.EndFace
            Return preferred
        End If

        ' SlopeView can fall back to LongitudinalFacade or PlanContour
        If preferred = ViewRole.SlopeView Then
            If placedViews.ContainsKey(ViewRole.LongitudinalFacade) Then Return ViewRole.LongitudinalFacade
            If placedViews.ContainsKey(ViewRole.PlanContour) Then Return ViewRole.PlanContour
        End If

        ' PlanContour can fall back to LongitudinalFacade
        If preferred = ViewRole.PlanContour Then
            If placedViews.ContainsKey(ViewRole.LongitudinalFacade) Then Return ViewRole.LongitudinalFacade
            If placedViews.ContainsKey(ViewRole.CrossProfile) Then Return ViewRole.CrossProfile
        End If

        ' CrossProfile can fall back to EndFace or ThicknessView
        If preferred = ViewRole.CrossProfile Then
            If placedViews.ContainsKey(ViewRole.ThicknessView) Then Return ViewRole.ThicknessView
            If placedViews.ContainsKey(ViewRole.EndFace) Then Return ViewRole.EndFace
            If placedViews.ContainsKey(ViewRole.LongitudinalFacade) Then Return ViewRole.LongitudinalFacade
        End If

        ' LongitudinalFacade can fall back to PlanContour or SlopeView
        If preferred = ViewRole.LongitudinalFacade Then
            If placedViews.ContainsKey(ViewRole.PlanContour) Then Return ViewRole.PlanContour
            If placedViews.ContainsKey(ViewRole.SlopeView) Then Return ViewRole.SlopeView
        End If

        ' EndFace can fall back to CrossProfile
        If preferred = ViewRole.EndFace Then
            If placedViews.ContainsKey(ViewRole.CrossProfile) Then Return ViewRole.CrossProfile
        End If

        ' MainContour — try all semantic roles in priority order
        If preferred = ViewRole.MainContour Then
            Dim priorities As ViewRole() = {ViewRole.PlanContour, ViewRole.LongitudinalFacade, ViewRole.CrossProfile, ViewRole.SlopeView, ViewRole.ThicknessView, ViewRole.EndFace}
            For Each r As ViewRole In priorities
                If placedViews.ContainsKey(r) Then Return r
            Next
        End If

        ' Last resort: return first non-iso key in placed dict
        For Each kvp As KeyValuePair(Of ViewRole, DrawingView) In placedViews
            If kvp.Key <> ViewRole.IsoReference AndAlso kvp.Value IsNot Nothing Then
                Return kvp.Key
            End If
        Next

        Return preferred
    End Function

    Private Enum DimensionAxis
        Horizontal
        Vertical
    End Enum

    Private Function BuildGlobalDimensionScope(viewKey As String, v As DrawingView) As String
        Dim viewId As String = If(v Is Nothing, "nil", RuntimeHelpers.GetHashCode(v).ToString())
        Return viewKey & "|" & viewId
    End Function

    Private Function ResolveOverallIntentAxis(intent As DimensionIntentId,
                                              role As ViewRole,
                                              v As DrawingView,
                                              measure As ViewMeasure) As DimensionAxis
        Dim majorAxis As DimensionAxis = ResolveMajorAxis(v, measure)
        Dim minorAxis As DimensionAxis = OppositeAxis(majorAxis)

        Select Case intent
            Case DimensionIntentId.OverallLength, DimensionIntentId.ChordOrSpan
                If role = ViewRole.PlanContour OrElse role = ViewRole.LongitudinalFacade Then Return majorAxis
                Return majorAxis
            Case DimensionIntentId.OverallWidth
                If role = ViewRole.PlanContour Then Return minorAxis
                Return minorAxis
            Case DimensionIntentId.OverallHeight
                If role = ViewRole.LongitudinalFacade Then Return minorAxis
                Return DimensionAxis.Vertical
            Case DimensionIntentId.OverallThickness
                If role = ViewRole.CrossProfile OrElse role = ViewRole.ThicknessView OrElse role = ViewRole.EndFace Then
                    Return minorAxis
                End If
                Return DimensionAxis.Vertical
            Case Else
                If IsHorizontalIntent(intent) Then Return DimensionAxis.Horizontal
                Return DimensionAxis.Vertical
        End Select
    End Function

    Private Function ResolveMajorAxis(v As DrawingView, measure As ViewMeasure) As DimensionAxis
        Dim w As Double = 0
        Dim h As Double = 0
        If v IsNot Nothing Then
            w = Math.Abs(v.Width)
            h = Math.Abs(v.Height)
        ElseIf measure IsNot Nothing Then
            w = Math.Abs(measure.UnitW)
            h = Math.Abs(measure.UnitH)
        End If
        Dim axis As DimensionAxis = If(w >= h, DimensionAxis.Horizontal, DimensionAxis.Vertical)
        If IsViewRotated90(v) Then
            axis = OppositeAxis(axis)
        End If
        Return axis
    End Function

    Private Function IsViewRotated90(v As DrawingView) As Boolean
        If v Is Nothing Then Return False
        Try
            Dim rightAngle As Double = Math.PI / 2.0
            Dim rot As Double = v.Rotation
            Dim steps As Integer = CInt(Math.Round(rot / rightAngle))
            Return (Math.Abs(steps) Mod 2) = 1
        Catch
        End Try
        Return False
    End Function

    Private Function OppositeAxis(axis As DimensionAxis) As DimensionAxis
        If axis = DimensionAxis.Horizontal Then Return DimensionAxis.Vertical
        Return DimensionAxis.Horizontal
    End Function

    Private Function IsOverallIntent(id As DimensionIntentId) As Boolean
        Return (id = DimensionIntentId.OverallLength OrElse id = DimensionIntentId.OverallWidth OrElse id = DimensionIntentId.OverallHeight OrElse
                id = DimensionIntentId.OverallThickness OrElse id = DimensionIntentId.ChordOrSpan)
    End Function

    Private Function IsProfileIntent(id As DimensionIntentId) As Boolean
        Return (id = DimensionIntentId.ProfileDepth OrElse id = DimensionIntentId.ProfileHeight OrElse id = DimensionIntentId.StepHeight OrElse
                id = DimensionIntentId.SlopeHeightHigh OrElse id = DimensionIntentId.SlopeHeightLow OrElse id = DimensionIntentId.VisibleRadius)
    End Function

    Private Function IsHorizontalIntent(id As DimensionIntentId) As Boolean
        Return (id = DimensionIntentId.OverallLength OrElse id = DimensionIntentId.OverallWidth OrElse id = DimensionIntentId.ChordOrSpan OrElse
                id = DimensionIntentId.ProfileDepth OrElse id = DimensionIntentId.LipDepth OrElse id = DimensionIntentId.RecessOffset OrElse
                id = DimensionIntentId.EndCutLength OrElse id = DimensionIntentId.EdgeBandWidth)
    End Function

    Private Function IsVerticalIntent(id As DimensionIntentId) As Boolean
        Return (id = DimensionIntentId.OverallHeight OrElse id = DimensionIntentId.OverallThickness OrElse id = DimensionIntentId.ProfileHeight OrElse
                id = DimensionIntentId.StepHeight OrElse id = DimensionIntentId.SlopeHeightHigh OrElse id = DimensionIntentId.SlopeHeightLow OrElse
                id = DimensionIntentId.LipHeight OrElse id = DimensionIntentId.RecessDepth OrElse id = DimensionIntentId.Chamfer)
    End Function

    Private Function IsRadiusIntent(id As DimensionIntentId) As Boolean
        Return (id = DimensionIntentId.RadiusMain OrElse id = DimensionIntentId.RadiusSecondary OrElse id = DimensionIntentId.VisibleRadius)
    End Function

    Private Function GetCounter(dict As Dictionary(Of String, Integer), key As String) As Integer
        If dict.ContainsKey(key) Then Return dict(key)
        Return 0
    End Function

    Private Function TryAddOverallDimensions(doc As DrawingDocument, sheet As Sheet, v As DrawingView, slot As SlotRect,
                                             intent As DimensionIntentId,
                                             role As ViewRole,
                                             measure As ViewMeasure,
                                             modelSize As ModelOverallExtents,
                                             usedKeys As HashSet(Of String),
                                             noteKeys As HashSet(Of String),
                                             counters As Dictionary(Of String, Integer),
                                             viewKey As String,
                                             globalDimensionKeys As HashSet(Of String)) As Integer
        Dim dedupeKey As String = viewKey & "|" & intent.ToString()
        If usedKeys.Contains(dedupeKey) Then Return 0

        ' overall-intents must be fully model-driven when real 3D extents are available
        If IsOverallIntent(intent) AndAlso modelSize IsNot Nothing AndAlso modelSize.IsValid Then
            If AddFallbackDimensionNotes(doc, sheet, v, slot, intent, role, measure, modelSize, noteKeys) Then
                usedKeys.Add(dedupeKey)
                counters(viewKey) = GetCounter(counters, viewKey) + 1
                Return 1
            End If
            Return 0
        End If

        Dim axis As DimensionAxis = ResolveOverallIntentAxis(intent, role, v, measure)
        Dim addH As Boolean = (axis = DimensionAxis.Horizontal)
        Dim addV As Boolean = (axis = DimensionAxis.Vertical)
        Dim dedupeScope As String = BuildGlobalDimensionScope(viewKey, v)

        Dim added As Integer = TryAddTrueDimensions(doc, sheet, v, slot, addH, addV, False, False, Integer.MaxValue, globalDimensionKeys, dedupeScope)
        If added = 0 Then
            If AddFallbackDimensionNotes(doc, sheet, v, slot, intent, role, measure, modelSize, noteKeys) Then added = 1
        End If

        If added > 0 Then
            usedKeys.Add(dedupeKey)
            counters(viewKey) = GetCounter(counters, viewKey) + added
        End If
        Return added
    End Function

    Private Function TryAddProfileDimensions(doc As DrawingDocument, sheet As Sheet, v As DrawingView, slot As SlotRect,
                                             intent As DimensionIntentId,
                                             measure As ViewMeasure,
                                             usedKeys As HashSet(Of String),
                                             noteKeys As HashSet(Of String),
                                             counters As Dictionary(Of String, Integer),
                                             viewKey As String,
                                             globalDimensionKeys As HashSet(Of String)) As Integer
        Dim addH As Boolean = (intent = DimensionIntentId.ProfileDepth)
        Dim addV As Boolean = (intent = DimensionIntentId.ProfileHeight OrElse intent = DimensionIntentId.StepHeight OrElse intent = DimensionIntentId.SlopeHeightHigh OrElse intent = DimensionIntentId.SlopeHeightLow)
        If intent = DimensionIntentId.VisibleRadius Then
            Dim radial As Integer = TryAddRadiusDimension(sheet, v)
            If radial > 0 Then Return radial
            If AddFallbackDimensionNotes(doc, sheet, v, slot, intent, ViewRole.MainContour, measure, Nothing, noteKeys) Then Return 1
            Return 0
        End If

        Dim dedupeKey As String = viewKey & "|" & intent.ToString()
        If usedKeys.Contains(dedupeKey) Then Return 0
        Dim dedupeScope As String = BuildGlobalDimensionScope(viewKey, v)
        Dim added As Integer = TryAddTrueDimensions(doc, sheet, v, slot, addH, addV, True, True, Integer.MaxValue, globalDimensionKeys, dedupeScope, False, False)
        If added = 0 Then
            added += TryAddOuterInnerDimension(doc, sheet, v, slot, addH, globalDimensionKeys, dedupeScope)
            If intent = DimensionIntentId.SlopeHeightHigh OrElse intent = DimensionIntentId.SlopeHeightLow Then
                added += TryAddSlopeHeightByExtremums(doc, sheet, v, slot, globalDimensionKeys, dedupeScope)
            End If
        End If
        If added = 0 Then
            If AddFallbackDimensionNotes(doc, sheet, v, slot, intent, ViewRole.MainContour, measure, Nothing, noteKeys) Then added = 1
        End If
        If added > 0 Then
            usedKeys.Add(dedupeKey)
            counters(viewKey) = GetCounter(counters, viewKey) + added
        End If
        Return added
    End Function

    Private Function TryAddFeatureDimensions(doc As DrawingDocument, sheet As Sheet, v As DrawingView, slot As SlotRect,
                                             intent As DimensionIntentId,
                                             measure As ViewMeasure,
                                             usedKeys As HashSet(Of String),
                                             noteKeys As HashSet(Of String),
                                             counters As Dictionary(Of String, Integer),
                                             viewKey As String,
                                             globalDimensionKeys As HashSet(Of String)) As Integer
        If intent = DimensionIntentId.RadiusMain OrElse intent = DimensionIntentId.RadiusSecondary Then
            Dim byRadius As Integer = TryAddRadiusDimension(sheet, v)
            If byRadius > 0 Then Return byRadius
            If AddFallbackDimensionNotes(doc, sheet, v, slot, intent, ViewRole.MainContour, measure, Nothing, noteKeys) Then Return 1
            Return 0
        End If

        Dim addH As Boolean = (intent = DimensionIntentId.LipDepth OrElse intent = DimensionIntentId.RecessOffset OrElse intent = DimensionIntentId.EndCutLength OrElse intent = DimensionIntentId.EdgeBandWidth)
        Dim addV As Boolean = (intent = DimensionIntentId.LipHeight OrElse intent = DimensionIntentId.RecessDepth OrElse intent = DimensionIntentId.Chamfer)

        Dim dedupeKey As String = viewKey & "|" & intent.ToString()
        If usedKeys.Contains(dedupeKey) Then Return 0
        Dim dedupeScope As String = BuildGlobalDimensionScope(viewKey, v)
        Dim added As Integer = TryAddTrueDimensions(doc, sheet, v, slot, addH, addV, True, True, Integer.MaxValue, globalDimensionKeys, dedupeScope, False, False)
        If added = 0 Then
            added += TryAddOuterInnerDimension(doc, sheet, v, slot, addH, globalDimensionKeys, dedupeScope)
            If intent = DimensionIntentId.EndCutLength OrElse intent = DimensionIntentId.Chamfer Then
                added += TryAddSlopedProjectionDimension(doc, sheet, v, slot, globalDimensionKeys, dedupeScope)
            End If
        End If
        If added = 0 Then
            If AddFallbackDimensionNotes(doc, sheet, v, slot, intent, ViewRole.MainContour, measure, Nothing, noteKeys) Then added = 1
        End If
        If added > 0 Then
            usedKeys.Add(dedupeKey)
            counters(viewKey) = GetCounter(counters, viewKey) + added
        End If
        Return added
    End Function

    Private Function TryAddRadiusDimension(sheet As Sheet, v As DrawingView) As Integer
        If sheet Is Nothing OrElse v Is Nothing Then Return 0
        Try
            Dim bucket As CurveBucket = CollectViewCurves(v)
            For Each c As DrawingCurve In bucket.Arcs
                Try
                    Dim gi As GeometryIntent = sheet.CreateGeometryIntent(c)
                    Dim rb As Box2d = c.RangeBox
                    Dim px As Double = Math.Min(v.Left + v.Width + 0.7, rb.MaxPoint.X + 0.8)
                    Dim py As Double = (rb.MinPoint.Y + rb.MaxPoint.Y) / 2.0
                    Dim pt As Point2d = _app.TransientGeometry.CreatePoint2d(px, py)
                    sheet.DrawingDimensions.GeneralDimensions.AddRadius(pt, gi)
                    Return 1
                Catch exArc As Exception
                    Debug.Print("WARN TryAddRadiusDimension arc: " & exArc.Message)
                End Try
            Next
        Catch ex As Exception
            Debug.Print("WARN TryAddRadiusDimension: " & ex.Message)
        End Try
        Debug.Print("TryAddRadiusDimension: no radius dimension added")
        Return 0
    End Function

    Private Function BuildLayoutPatternsByArchetype(doc As DrawingDocument, safe As SlotRect, gap As Double, archetype As LayoutArchetype) As List(Of LayoutPattern)
        Dim result As New List(Of LayoutPattern)()
        Dim w As Double = RectW(safe)
        Dim h As Double = RectH(safe)
        If w <= gap * 4 OrElse h <= gap * 4 Then Return result

        If archetype = LayoutArchetype.RadialSegment Then
            Dim leftW As Double = Math.Max(Cm(doc, 82.0), w * 0.36)
            Dim topSplit As Double = safe.B + h * 0.52

            Dim pA As New LayoutPattern()
            pA.PatternName = "RADIAL_A"
            pA.Archetype = LayoutArchetype.RadialSegment
            pA.MainSlot = New SlotRect(safe.L + leftW + gap, safe.R, safe.B + h * 0.34, safe.T)
            pA.Aux1Slot = New SlotRect(safe.L, safe.L + leftW, safe.B, topSplit - gap)
            pA.Aux2Slot = New SlotRect(safe.L, safe.L + leftW, topSplit, safe.T)
            result.Add(pA)

            Dim pB As New LayoutPattern()
            pB.PatternName = "RADIAL_B"
            pB.Archetype = LayoutArchetype.RadialSegment
            pB.MainSlot = New SlotRect(safe.L, safe.L + leftW, safe.B + h * 0.34, safe.T)
            pB.Aux1Slot = New SlotRect(safe.L + leftW + gap, safe.R, safe.B + h * 0.34, safe.T)
            pB.Aux2Slot = New SlotRect(safe.L, safe.L + leftW, safe.B, safe.B + h * 0.30)
            result.Add(pB)

        ElseIf archetype = LayoutArchetype.LongLinear Then
            Dim leftW As Double = Math.Max(Cm(doc, 75.0), w * 0.56)
            Dim rightL As Double = safe.L + leftW + gap

            Dim pA As New LayoutPattern()
            pA.PatternName = "LONG_A"
            pA.Archetype = LayoutArchetype.LongLinear
            pA.MainSlot = New SlotRect(safe.L, safe.L + leftW, safe.B, safe.T)
            pA.Aux1Slot = New SlotRect(rightL, safe.R, safe.B + h * 0.50, safe.T)
            pA.Aux2Slot = New SlotRect(rightL, safe.R, safe.B, safe.B + h * 0.46)
            result.Add(pA)

            Dim pB As New LayoutPattern()
            pB.PatternName = "LONG_B"
            pB.Archetype = LayoutArchetype.LongLinear
            pB.MainSlot = New SlotRect(safe.L, safe.L + leftW, safe.B + h * 0.06, safe.T)
            pB.Aux1Slot = New SlotRect(rightL, safe.R, safe.B + h * 0.54, safe.T)
            pB.Aux2Slot = New SlotRect(rightL, safe.R, safe.B, safe.B + h * 0.50)
            result.Add(pB)

            Dim facadeMainR As Double = safe.L + w * 0.27
            Dim facadeAux1L As Double = safe.L + w * 0.34
            Dim facadeAux2L As Double = safe.L + w * 0.45
            Dim rightProtect As Double = Math.Max(gap * 0.5, Cm(doc, 2.0))

            Dim pC As New LayoutPattern()
            pC.PatternName = "LONG_FACADE_LEFT"
            pC.Archetype = LayoutArchetype.LongLinear
            pC.MainSlot = New SlotRect(safe.L, facadeMainR, safe.B, safe.T)
            pC.Aux1Slot = New SlotRect(facadeAux1L, safe.R - rightProtect, safe.B + h * 0.56, safe.T)
            pC.Aux2Slot = New SlotRect(facadeAux2L, safe.R, safe.B, safe.B + h * 0.48)
            result.Add(pC)

        Else
            Dim leftW As Double = Math.Max(Cm(doc, 80.0), w * 0.34)

            Dim pA As New LayoutPattern()
            pA.PatternName = "PLATE_A"
            pA.Archetype = LayoutArchetype.PlateBlock
            pA.MainSlot = New SlotRect(safe.L + leftW + gap, safe.R, safe.B + h * 0.38, safe.T)
            pA.Aux1Slot = New SlotRect(safe.L, safe.L + leftW, safe.B, safe.T)
            pA.Aux2Slot = New SlotRect(safe.L + leftW + gap, safe.R, safe.B, safe.B + h * 0.34)
            result.Add(pA)

            Dim pB As New LayoutPattern()
            pB.PatternName = "PLATE_B"
            pB.Archetype = LayoutArchetype.PlateBlock
            pB.MainSlot = New SlotRect(safe.L, safe.L + w * 0.62, safe.B + h * 0.35, safe.T)
            pB.Aux1Slot = New SlotRect(safe.L + w * 0.62 + gap, safe.R, safe.B + h * 0.35, safe.T)
            pB.Aux2Slot = New SlotRect(safe.L, safe.L + w * 0.62, safe.B, safe.B + h * 0.30)
            result.Add(pB)
        End If

        For Each p As LayoutPattern In result
            p.MainSlot = InsetRect(p.MainSlot, gap * 0.22)
            p.Aux1Slot = InsetRect(p.Aux1Slot, gap * 0.22)
            p.Aux2Slot = InsetRect(p.Aux2Slot, gap * 0.22)
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

    Private Function BuildOrderedViewPairs(measures As List(Of ViewMeasure)) As List(Of AuxPair)
        Dim result As New List(Of AuxPair)()
        If measures Is Nothing Then Return result

        For i As Integer = 0 To measures.Count - 1
            Dim a As ViewMeasure = measures(i)
            If a Is Nothing Then Continue For
            For j As Integer = 0 To measures.Count - 1
                If i = j Then Continue For
                Dim b As ViewMeasure = measures(j)
                If b Is Nothing Then Continue For
                If String.Equals(a.Key, b.Key, StringComparison.OrdinalIgnoreCase) Then Continue For
                result.Add(New AuxPair(a, b))
            Next
        Next

        Return result
    End Function

    Private Function PickPreferredOppositeView(preferred As ViewMeasure, alternate As ViewMeasure) As ViewMeasure
        If preferred Is Nothing Then Return alternate
        If alternate Is Nothing Then Return preferred

        Dim d As Integer = preferred.CurveCount - alternate.CurveCount
        If Math.Abs(d) > 1 Then
            If d > 0 Then Return preferred
            Return alternate
        End If

        If String.Equals(preferred.Key, "FRONT", StringComparison.OrdinalIgnoreCase) AndAlso
           String.Equals(alternate.Key, "BACK", StringComparison.OrdinalIgnoreCase) Then
            Return preferred
        End If
        If String.Equals(preferred.Key, "LEFT", StringComparison.OrdinalIgnoreCase) AndAlso
           String.Equals(alternate.Key, "RIGHT", StringComparison.OrdinalIgnoreCase) Then
            Return preferred
        End If

        Return preferred
    End Function

    Private Function BuildAll2DCandidates(mFront As ViewMeasure,
                                          mBack As ViewMeasure,
                                          mTop As ViewMeasure,
                                          mLeft As ViewMeasure,
                                          mRight As ViewMeasure) As List(Of ViewMeasure)
        Dim result As New List(Of ViewMeasure)()

        Dim ordered As New List(Of ViewMeasure)()
        If mFront IsNot Nothing Then ordered.Add(mFront)
        If mBack IsNot Nothing Then ordered.Add(mBack)
        If mTop IsNot Nothing Then ordered.Add(mTop)
        If mLeft IsNot Nothing Then ordered.Add(mLeft)
        If mRight IsNot Nothing Then ordered.Add(mRight)

        For Each m As ViewMeasure In ordered
            If m Is Nothing Then Continue For
            If String.Equals(m.Key, "ISO", StringComparison.OrdinalIgnoreCase) Then Continue For
            Dim exists As Boolean = False
            For Each x As ViewMeasure In result
                If x Is m Then exists = True : Exit For
                If Not String.IsNullOrWhiteSpace(x.Key) AndAlso String.Equals(x.Key, m.Key, StringComparison.OrdinalIgnoreCase) Then
                    exists = True
                    Exit For
                End If
            Next
            If Not exists Then result.Add(m)
        Next

        Return result
    End Function

    Private Function BuildPreferredAuxMeasures(mFront As ViewMeasure,
                                               mBack As ViewMeasure,
                                               mTop As ViewMeasure,
                                               mLeft As ViewMeasure,
                                               mRight As ViewMeasure) As List(Of ViewMeasure)
        Dim result As New List(Of ViewMeasure)()

        Dim frontBack As ViewMeasure = PickPreferredOppositeView(mFront, mBack)
        Dim leftRight As ViewMeasure = PickPreferredOppositeView(mLeft, mRight)

        If frontBack IsNot Nothing Then result.Add(frontBack)
        If mTop IsNot Nothing Then result.Add(mTop)
        If leftRight IsNot Nothing Then result.Add(leftRight)

        Return result
    End Function

    Private Function IsProfileLikeMeasure(m As ViewMeasure) As Boolean
        If m Is Nothing Then Return False
        If String.Equals(m.Key, "TOP", StringComparison.OrdinalIgnoreCase) Then Return False
        If m.CurveCount >= 12 Then Return True
        If m.CurveCount >= 8 AndAlso m.AspectRatio <= 6.0 Then Return True
        If m.AspectRatio <= 3.4 AndAlso m.CurveCount >= 5 Then Return True
        Return False
    End Function

    Private Function IsLongitudinalMeasure(m As ViewMeasure) As Boolean
        If m Is Nothing Then Return False
        Return m.AspectRatio >= 4.2
    End Function

    Private Function IsFacadeLikeLongMeasure(m As ViewMeasure) As Boolean
        If m Is Nothing Then Return False
        If String.Equals(m.Key, "ISO", StringComparison.OrdinalIgnoreCase) Then Return False
        If m.CurveCount <= 2 Then Return False
        If Not IsLongitudinalMeasure(m) Then Return False
        If m.CurveCount <= 11 AndAlso m.AspectRatio >= 4.4 Then Return True
        If m.CurveCount <= 14 AndAlso m.AspectRatio >= 5.2 Then Return True
        Return False
    End Function

    Private Function IsDetailedProfileMeasure(m As ViewMeasure) As Boolean
        If m Is Nothing Then Return False
        If String.Equals(m.Key, "ISO", StringComparison.OrdinalIgnoreCase) Then Return False
        If String.Equals(m.Key, "TOP", StringComparison.OrdinalIgnoreCase) AndAlso m.CurveCount <= 6 Then Return False
        If IsFacadeLikeLongMeasure(m) AndAlso m.CurveCount <= 10 Then Return False
        If m.CurveCount >= 12 Then Return True
        If m.CurveCount >= 8 AndAlso m.AspectRatio <= 6.0 Then Return True
        If m.AspectRatio <= 4.6 AndAlso m.CurveCount >= 6 Then Return True
        Return False
    End Function

    Private Function WouldPreferVerticalPlacement(slot As SlotRect, m As ViewMeasure) As Boolean
        If slot Is Nothing OrElse m Is Nothing Then Return False
        Dim sw As Double = RectW(slot)
        Dim sh As Double = RectH(slot)
        If sw <= 0 OrElse sh <= 0 Then Return False

        Dim fit0 As FitResult = BuildFitResult(sw, sh, m.UnitW, m.UnitH, ORTHO_SCALE_MARGIN, False)
        Dim fit90 As FitResult = BuildFitResult(sw, sh, m.UnitW, m.UnitH, ORTHO_SCALE_MARGIN, True)
        If fit90 Is Nothing Then Return False
        If fit0 Is Nothing Then Return True

        Dim slotTall As Boolean = sh > sw * 1.08
        Dim fit90Tall As Boolean = fit90.ProjectedH >= fit90.ProjectedW * 1.02
        If fit90.Scale >= fit0.Scale * 1.04 AndAlso slotTall Then Return True
        If fit90.Scale > fit0.Scale AndAlso fit90Tall AndAlso slotTall Then Return True
        Return False
    End Function

    Private Function IsPlanLikeMeasure(m As ViewMeasure) As Boolean
        If m Is Nothing Then Return False
        If String.Equals(m.Key, "TOP", StringComparison.OrdinalIgnoreCase) Then Return True
        Dim area As Double = m.UnitW * m.UnitH
        If area <= 0 Then Return False
        If m.AspectRatio <= 3.8 AndAlso m.CurveCount <= 18 Then Return True
        If area >= 1200 AndAlso m.AspectRatio <= 5.5 Then Return True
        Return False
    End Function

    Private Function DetectLayoutArchetype(mFront As ViewMeasure,
                                           mBack As ViewMeasure,
                                           mTop As ViewMeasure,
                                           mLeft As ViewMeasure,
                                           mRight As ViewMeasure) As LayoutArchetype
        Dim allM As List(Of ViewMeasure) = BuildAll2DCandidates(mFront, mBack, mTop, mLeft, mRight)
        If allM Is Nothing OrElse allM.Count = 0 Then Return LayoutArchetype.PlateBlock

        Dim hasVeryLong As Boolean = False
        Dim hasRadial As Boolean = False
        Dim planCount As Integer = 0

        For Each m As ViewMeasure In allM
            If m Is Nothing Then Continue For
            If IsPlanLikeMeasure(m) Then planCount += 1
            If IsLongitudinalMeasure(m) AndAlso m.AspectRatio >= 5.0 Then hasVeryLong = True
            If m.CurveCount >= 15 AndAlso m.AspectRatio <= 4.8 Then hasRadial = True
            If IsPlanLikeMeasure(m) AndAlso m.CurveCount >= 13 AndAlso m.AspectRatio <= 4.2 Then hasRadial = True
        Next

        If hasRadial AndAlso (Not hasVeryLong) Then Return LayoutArchetype.RadialSegment
        If hasVeryLong Then Return LayoutArchetype.LongLinear
        If planCount >= 1 Then Return LayoutArchetype.PlateBlock
        Return LayoutArchetype.PlateBlock
    End Function

    Private Function SelectMainMeasureForArchetype(archetype As LayoutArchetype,
                                                    measures As List(Of ViewMeasure),
                                                    preferredFrontBack As ViewMeasure,
                                                    preferredLeftRight As ViewMeasure,
                                                    mTop As ViewMeasure) As ViewMeasure
        Dim best As ViewMeasure = Nothing
        Dim bestScore As Double = -1000000.0

        For Each m As ViewMeasure In measures
            If m Is Nothing Then Continue For
            Dim sc As Double = 0
            sc += m.CurveCount * 0.06

            If archetype = LayoutArchetype.RadialSegment Then
                If IsPlanLikeMeasure(m) Then sc += 2.2
                If m.CurveCount >= 14 Then sc += 1.2
                If IsLongitudinalMeasure(m) Then sc -= 1.4
                If IsProfileLikeMeasure(m) Then sc -= 0.35
            ElseIf archetype = LayoutArchetype.LongLinear Then
                If IsLongitudinalMeasure(m) Then sc += 2.6
                If IsProfileLikeMeasure(m) Then sc -= 0.6
                If String.Equals(m.Key, "TOP", StringComparison.OrdinalIgnoreCase) Then sc -= 0.3
            Else
                If IsPlanLikeMeasure(m) Then sc += 2.4
                If IsLongitudinalMeasure(m) Then sc -= 0.5
            End If

            If preferredFrontBack Is m Then sc += 0.2
            If preferredLeftRight Is m Then sc += 0.2
            If mTop Is m Then sc += 0.25

            If best Is Nothing OrElse sc > bestScore Then
                best = m
                bestScore = sc
            End If
        Next

        Return best
    End Function

    Private Function EvaluatePlanFlexible(ptn As LayoutPattern,
                                          archetype As LayoutArchetype,
                                          mainM As ViewMeasure,
                                          auxA As ViewMeasure,
                                          isoM As ViewMeasure) As LayoutPlan
        If ptn Is Nothing OrElse mainM Is Nothing OrElse auxA Is Nothing OrElse isoM Is Nothing Then Return Nothing
        If String.Equals(mainM.Key, auxA.Key, StringComparison.OrdinalIgnoreCase) Then Return Nothing

        Dim plan As New LayoutPlan()
        plan.MainSlot = ptn.MainSlot
        plan.Aux1Slot = ptn.Aux1Slot
        plan.Aux2Slot = ptn.Aux2Slot
        plan.MainMeasure = mainM
        plan.Aux1Measure = auxA
        plan.Aux2Measure = isoM

        plan.MainFit = ScaleToFit(plan.MainSlot, mainM, ORTHO_SCALE_MARGIN)
        plan.Aux1Fit = ScaleToFit(plan.Aux1Slot, auxA, ORTHO_SCALE_MARGIN)
        plan.Aux2Fit = ScaleToFit(plan.Aux2Slot, isoM, ISO_SCALE_MARGIN)

        If plan.MainFit Is Nothing OrElse plan.MainFit.Scale <= 0 Then Return Nothing
        If plan.Aux1Fit Is Nothing OrElse plan.Aux1Fit.Scale <= 0 Then Return Nothing
        If plan.Aux2Fit Is Nothing OrElse plan.Aux2Fit.Scale <= 0 Then Return Nothing

        Dim mainArea As Double = plan.MainFit.ProjectedW * plan.MainFit.ProjectedH
        Dim aux1Area As Double = plan.Aux1Fit.ProjectedW * plan.Aux1Fit.ProjectedH
        Dim aux2Area As Double = plan.Aux2Fit.ProjectedW * plan.Aux2Fit.ProjectedH

        Dim workArea As Double = RectW(plan.MainSlot) * RectH(plan.MainSlot) +
                                 RectW(plan.Aux1Slot) * RectH(plan.Aux1Slot) +
                                 RectW(plan.Aux2Slot) * RectH(plan.Aux2Slot)
        If workArea <= 0 Then Return Nothing

        Dim fill As Double = (mainArea + aux1Area + aux2Area) / workArea
        Dim mainDominance As Double = mainArea / Math.Max(0.0001, mainArea + aux1Area + aux2Area)
        Dim auxBalance As Double = Math.Min(aux1Area, aux2Area) / Math.Max(0.0001, Math.Max(aux1Area, aux2Area))
        plan.Score = fill * 0.34 + mainDominance * 0.38 + auxBalance * 0.28

        Dim mainProfile As Boolean = IsProfileLikeMeasure(mainM)
        Dim mainLong As Boolean = IsLongitudinalMeasure(mainM)
        Dim mainPlan As Boolean = IsPlanLikeMeasure(mainM)
        Dim auxProfile As Boolean = IsProfileLikeMeasure(auxA)
        Dim auxLong As Boolean = IsLongitudinalMeasure(auxA)
        Dim auxPlan As Boolean = IsPlanLikeMeasure(auxA)

        If mainArea < aux1Area * 0.72 Then plan.Score -= 0.26
        If Math.Abs(mainM.AspectRatio - auxA.AspectRatio) < 0.24 AndAlso Math.Abs(mainM.CurveCount - auxA.CurveCount) <= 1 Then plan.Score -= 0.18
        If mainM.CurveCount <= 2 AndAlso auxA.CurveCount <= 2 Then plan.Score -= 0.2
        If mainM.CurveCount >= 6 AndAlso auxA.CurveCount >= 4 AndAlso Math.Abs(mainM.AspectRatio - auxA.AspectRatio) > 0.7 Then plan.Score += 0.08

        If mainLong AndAlso auxProfile Then plan.Score += 0.07
        If mainProfile AndAlso auxLong Then plan.Score -= 0.08

        If String.Equals(mainM.Key, "ISO", StringComparison.OrdinalIgnoreCase) Then plan.Score -= 0.5
        If String.Equals(auxA.Key, "ISO", StringComparison.OrdinalIgnoreCase) Then plan.Score -= 0.5
        If String.Equals(isoM.Key, "ISO", StringComparison.OrdinalIgnoreCase) Then plan.Score += 0.08

        If RectW(plan.Aux2Slot) > RectW(plan.MainSlot) OrElse RectH(plan.Aux2Slot) > RectH(plan.MainSlot) Then plan.Score -= 0.12

        If archetype = LayoutArchetype.LongLinear Then
            If String.Equals(ptn.PatternName, "LONG_FACADE_LEFT", StringComparison.OrdinalIgnoreCase) Then
                Dim facadeMain As Boolean = IsFacadeLikeLongMeasure(mainM)
                Dim facadeAux As Boolean = IsFacadeLikeLongMeasure(auxA)
                Dim profileMain As Boolean = IsDetailedProfileMeasure(mainM)
                Dim profileAux As Boolean = IsDetailedProfileMeasure(auxA)
                Dim mainVertical As Boolean = plan.MainFit.Rotate90
                Dim preferVertical As Boolean = WouldPreferVerticalPlacement(plan.MainSlot, mainM)
                Dim similar2D As Boolean = Math.Abs(mainM.AspectRatio - auxA.AspectRatio) < 0.35 AndAlso Math.Abs(mainM.CurveCount - auxA.CurveCount) <= 2

                If facadeMain Then plan.Score += 0.42 Else plan.Score -= 0.34
                If profileAux Then plan.Score += 0.36 Else plan.Score -= 0.26
                If profileMain Then plan.Score -= 0.24
                If facadeAux Then plan.Score -= 0.22
                If String.Equals(mainM.Key, "FRONT", StringComparison.OrdinalIgnoreCase) Then plan.Score += 0.01
                If String.Equals(mainM.Key, "LEFT", StringComparison.OrdinalIgnoreCase) Then plan.Score += 0.008
                If String.Equals(auxA.Key, "TOP", StringComparison.OrdinalIgnoreCase) Then plan.Score += 0.01

                If auxA IsNot Nothing AndAlso Not String.Equals(auxA.Key, "ISO", StringComparison.OrdinalIgnoreCase) Then
                    plan.Score += 0.08
                End If
                If String.Equals(isoM.Key, "ISO", StringComparison.OrdinalIgnoreCase) Then plan.Score += 0.1

                If mainVertical Then plan.Score += 0.11
                If preferVertical AndAlso plan.MainFit.Rotate90 Then plan.Score += 0.15
                If preferVertical AndAlso (Not plan.MainFit.Rotate90) Then plan.Score -= 0.14

                If RectH(plan.MainSlot) > RectW(plan.MainSlot) * 1.4 Then plan.Score += 0.1
                If RectW(plan.Aux1Slot) > RectW(plan.MainSlot) * 1.2 Then plan.Score += 0.08
                If RectH(plan.Aux1Slot) >= RectH(plan.Aux2Slot) Then plan.Score += 0.05

                If similar2D Then plan.Score -= 0.16
                If mainLong AndAlso auxLong Then plan.Score -= 0.12
                If auxA.CurveCount <= 5 Then plan.Score -= 0.1
            Else
                If mainLong Then plan.Score += 0.28 Else plan.Score -= 0.18
                If auxProfile Then plan.Score += 0.18
                If mainLong AndAlso auxLong Then plan.Score -= 0.12
                If RectW(plan.MainSlot) > RectW(plan.Aux1Slot) Then plan.Score += 0.09
                If String.Equals(ptn.PatternName, "LONG_A", StringComparison.OrdinalIgnoreCase) Then
                    If RectW(plan.Aux2Slot) < RectW(plan.Aux1Slot) Then plan.Score += 0.03
                End If
                If mainProfile AndAlso (Not mainLong) Then plan.Score -= 0.2
                If mainM.AspectRatio < 2.8 AndAlso auxA.AspectRatio < 2.8 Then plan.Score -= 0.12
            End If

        ElseIf archetype = LayoutArchetype.PlateBlock Then
            Dim thinMain As Boolean = mainM.AspectRatio >= 6.0 AndAlso mainM.CurveCount <= 10
            Dim thinAux As Boolean = auxA.AspectRatio >= 6.0 AndAlso auxA.CurveCount <= 10

            If String.Equals(ptn.PatternName, "PLATE_A", StringComparison.OrdinalIgnoreCase) Then
                If mainPlan Then plan.Score += 0.24 Else plan.Score -= 0.18
                If thinMain Then plan.Score -= 0.22
                If thinAux Then plan.Score += 0.16
                If auxPlan Then plan.Score -= 0.08
                If RectW(plan.MainSlot) > RectW(plan.Aux1Slot) Then plan.Score += 0.08
                If RectH(plan.Aux2Slot) < RectH(plan.MainSlot) Then plan.Score += 0.05
            ElseIf String.Equals(ptn.PatternName, "PLATE_B", StringComparison.OrdinalIgnoreCase) Then
                If mainPlan Then plan.Score += 0.26
                If auxPlan OrElse auxProfile Then plan.Score += 0.12
                If thinMain Then plan.Score -= 0.22
                If RectW(plan.MainSlot) > RectW(plan.Aux1Slot) Then plan.Score += 0.08
            End If
            If String.Equals(mainM.Key, "LEFT", StringComparison.OrdinalIgnoreCase) OrElse String.Equals(mainM.Key, "RIGHT", StringComparison.OrdinalIgnoreCase) Then
                If thinMain AndAlso (Not String.Equals(ptn.PatternName, "PLATE_A", StringComparison.OrdinalIgnoreCase)) Then
                    plan.Score -= 0.1
                End If
            End If

        ElseIf archetype = LayoutArchetype.RadialSegment Then
            If mainPlan OrElse mainM.CurveCount >= 14 Then plan.Score += 0.24 Else plan.Score -= 0.12
            If auxProfile Then plan.Score += 0.14
            If mainProfile AndAlso (Not mainPlan) Then plan.Score -= 0.18
            If mainM.AspectRatio >= 6.0 AndAlso mainM.CurveCount < 12 Then plan.Score -= 0.22
            If auxA.AspectRatio >= 6.2 AndAlso auxA.CurveCount < 8 Then plan.Score -= 0.12
            If String.Equals(ptn.PatternName, "RADIAL_B", StringComparison.OrdinalIgnoreCase) Then
                If mainPlan Then plan.Score += 0.08
            End If
        End If

        If String.Equals(mainM.Key, "FRONT", StringComparison.OrdinalIgnoreCase) Then plan.Score += 0.0006
        If String.Equals(mainM.Key, "LEFT", StringComparison.OrdinalIgnoreCase) Then plan.Score += 0.0004
        If String.Equals(auxA.Key, "FRONT", StringComparison.OrdinalIgnoreCase) Then plan.Score += 0.0002
        If String.Equals(auxA.Key, "LEFT", StringComparison.OrdinalIgnoreCase) Then plan.Score += 0.0001

        Return plan
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
        If isoM IsNot Nothing Then plan.IsoFit = ScaleToFit(plan.IsoSlot, isoM, ISO_SCALE_MARGIN)

        If plan.MainFit Is Nothing OrElse plan.Aux1Fit Is Nothing OrElse plan.Aux2Fit Is Nothing Then Return Nothing

        plan.MainMeasure = mainM
        plan.Aux1Measure = aux1M
        plan.Aux2Measure = aux2M

        Dim mainArea As Double = plan.MainFit.ProjectedW * plan.MainFit.ProjectedH
        Dim auxArea As Double = plan.Aux1Fit.ProjectedW * plan.Aux1Fit.ProjectedH +
                                plan.Aux2Fit.ProjectedW * plan.Aux2Fit.ProjectedH
        Dim isoArea As Double = 0
        If plan.IsoFit IsNot Nothing Then isoArea = plan.IsoFit.ProjectedW * plan.IsoFit.ProjectedH

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

    Private Sub AddViewAnnotations(doc As DrawingDocument, sheet As Sheet,
                                   v As DrawingView, slot As SlotRect,
                                   measure As ViewMeasure,
                                   Optional includeDimensions As Boolean = True)
        If v Is Nothing Then Return

        Dim caption As String = "Вид"
        If measure IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(measure.Caption) Then
            caption = measure.Caption
        End If

        Try
            Dim px As Double = Math.Max(slot.L + Cm(doc, 2.0), Math.Min(slot.R - Cm(doc, 2.0), v.Left + v.Width / 2.0))
            Dim belowY As Double = v.Top - v.Height - Cm(doc, VIEW_CAPTION_GAP_MM)
            Dim topSafe As Double = slot.T - Cm(doc, 2.0)
            Dim botSafe As Double = slot.B + Cm(doc, 2.0)
            Dim capY As Double = belowY

            If capY < botSafe Then
                capY = Math.Min(topSafe, v.Top + Cm(doc, VIEW_CAPTION_GAP_MM))
            End If
            capY = Math.Max(botSafe, Math.Min(topSafe, capY))
            sheet.DrawingNotes.GeneralNotes.AddFitted(_app.TransientGeometry.CreatePoint2d(px, capY), caption)
        Catch ex As Exception
            Debug.Print("WARN AddViewAnnotations caption: " & ex.Message)
        End Try

        If Not includeDimensions Then Return
        If Not ADD_VIEW_DIMENSIONS Then Return
        If measure Is Nothing Then Return

        If String.Equals(measure.Key, "ISO", StringComparison.OrdinalIgnoreCase) AndAlso (Not DIMENSIONS_ON_MAIN_ISO) Then Return

        Dim facadeLike As Boolean = IsFacadeLikeLongMeasure(measure)
        Dim detailedProfile As Boolean = IsDetailedProfileMeasure(measure)
        Dim lightDim As Boolean = ShouldPreferLightDimensioning(measure)

        Dim addHorizontal As Boolean = True
        Dim addVertical As Boolean = True
        If lightDim Then
            addHorizontal = False
            addVertical = True
        ElseIf detailedProfile Then
            addHorizontal = True
            addVertical = True
        End If

        Dim added As Integer = 0
        Dim localNoteKeys As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        added += TryAddTrueDimensions(doc, sheet, v, slot, addHorizontal, addVertical)

        If detailedProfile Then
            If added < 2 Then AddFallbackDimensionNotes(doc, sheet, v, slot, DimensionIntentId.ProfileHeight, ViewRole.MainContour, measure, Nothing, localNoteKeys)
        ElseIf facadeLike Then
            If lightDim Then
                If added < 1 Then AddFallbackDimensionNotes(doc, sheet, v, slot, DimensionIntentId.OverallHeight, ViewRole.MainContour, measure, Nothing, localNoteKeys)
            Else
                If added < 1 Then AddFallbackDimensionNotes(doc, sheet, v, slot, DimensionIntentId.OverallHeight, ViewRole.MainContour, measure, Nothing, localNoteKeys)
                If added < 2 Then AddFallbackDimensionNotes(doc, sheet, v, slot, DimensionIntentId.OverallLength, ViewRole.MainContour, measure, Nothing, localNoteKeys)
            End If
        Else
            If added < 2 Then AddFallbackDimensionNotes(doc, sheet, v, slot, DimensionIntentId.ProfileHeight, ViewRole.MainContour, measure, Nothing, localNoteKeys)
        End If
    End Sub

    Private Function ShouldPreferLightDimensioning(m As ViewMeasure) As Boolean
        If m Is Nothing Then Return False
        If IsFacadeLikeLongMeasure(m) AndAlso (Not IsDetailedProfileMeasure(m)) Then Return True
        Return False
    End Function

    Private Function TryAddTrueDimensions(doc As DrawingDocument, sheet As Sheet,
                                          v As DrawingView, slot As SlotRect,
                                          addHorizontal As Boolean,
                                          addVertical As Boolean,
                                          Optional preferLocal As Boolean = False,
                                          Optional includeFeatureOffsets As Boolean = True,
                                          Optional maxToAdd As Integer = Integer.MaxValue,
                                          Optional globalDimensionKeys As HashSet(Of String) = Nothing,
                                          Optional dedupeScope As String = "",
                                          Optional includeOverallExtremes As Boolean = True,
                                          Optional allowOuterFeatureOffsets As Boolean = True) As Integer
        Dim count As Integer = 0
        Try
            Dim bucket As CurveBucket = CollectViewCurves(v)
            Dim placedPairs As New HashSet(Of String)(StringComparer.Ordinal)

            If addHorizontal Then
                If includeOverallExtremes Then
                    Dim hExt As CurvePair = FindExtremeCurves(bucket, True)
                    count += AddLinearPairDimension(doc, sheet, v, slot, hExt, DimensionTypeEnum.kHorizontalDimensionType, False, placedPairs, globalDimensionKeys, dedupeScope)
                    If count >= maxToAdd Then Return count
                End If

                If preferLocal Then
                    Dim hLocal As CurvePair = FindInnerParallelPairs(bucket, True)
                    count += AddLinearPairDimension(doc, sheet, v, slot, hLocal, DimensionTypeEnum.kHorizontalDimensionType, True, placedPairs, globalDimensionKeys, dedupeScope)
                    If count >= maxToAdd Then Return count
                End If

                If includeFeatureOffsets Then
                    Dim pairs As List(Of CurvePair) = FindFeatureOffsets(bucket, True, allowOuterFeatureOffsets)
                    For Each pair As CurvePair In pairs
                        count += AddLinearPairDimension(doc, sheet, v, slot, pair, DimensionTypeEnum.kHorizontalDimensionType, True, placedPairs, globalDimensionKeys, dedupeScope)
                        If count >= maxToAdd Then Return count
                    Next
                End If
            End If

            If addVertical Then
                If includeOverallExtremes Then
                    Dim vExt As CurvePair = FindExtremeCurves(bucket, False)
                    count += AddLinearPairDimension(doc, sheet, v, slot, vExt, DimensionTypeEnum.kVerticalDimensionType, False, placedPairs, globalDimensionKeys, dedupeScope)
                    If count >= maxToAdd Then Return count
                End If

                If preferLocal Then
                    Dim vLocal As CurvePair = FindInnerParallelPairs(bucket, False)
                    count += AddLinearPairDimension(doc, sheet, v, slot, vLocal, DimensionTypeEnum.kVerticalDimensionType, True, placedPairs, globalDimensionKeys, dedupeScope)
                    If count >= maxToAdd Then Return count
                End If

                If includeFeatureOffsets Then
                    Dim pairs As List(Of CurvePair) = FindFeatureOffsets(bucket, False, allowOuterFeatureOffsets)
                    For Each pair As CurvePair In pairs
                        count += AddLinearPairDimension(doc, sheet, v, slot, pair, DimensionTypeEnum.kVerticalDimensionType, True, placedPairs, globalDimensionKeys, dedupeScope)
                        If count >= maxToAdd Then Return count
                    Next
                End If
            End If
        Catch ex As Exception
            Debug.Print("WARN TryAddTrueDimensions: " & ex.Message)
        End Try
        Return count
    End Function

    Private Function CollectViewCurves(v As DrawingView) As CurveBucket
        Dim bucket As New CurveBucket()
        If v Is Nothing Then Return bucket
        Try
            Dim curves As DrawingCurvesEnumerator = v.DrawingCurves
            If curves Is Nothing Then Return bucket

            Dim tol As Double = Math.Max(0.02, Math.Min(v.Width, v.Height) * 0.03)
            For i As Integer = 1 To curves.Count
                Dim c As DrawingCurve = curves.Item(i)
                If c Is Nothing Then Continue For
                Dim rb As Box2d = c.RangeBox
                Dim dx As Double = Math.Abs(rb.MaxPoint.X - rb.MinPoint.X)
                Dim dy As Double = Math.Abs(rb.MaxPoint.Y - rb.MinPoint.Y)
                Dim outer As Boolean = (Math.Abs(rb.MinPoint.X - v.Left) <= tol OrElse Math.Abs(rb.MaxPoint.X - (v.Left + v.Width)) <= tol OrElse Math.Abs(rb.MinPoint.Y - (v.Top - v.Height)) <= tol OrElse Math.Abs(rb.MaxPoint.Y - v.Top) <= tol)

                If c.CurveType = Curve2dTypeEnum.kCircularArcCurve2d OrElse c.CurveType = Curve2dTypeEnum.kCircleCurve2d Then
                    bucket.Arcs.Add(c)
                ElseIf dx >= dy * 2.2 Then
                    If outer Then
                        bucket.OuterHorizontal.Add(c)
                    Else
                        bucket.InnerHorizontal.Add(c)
                    End If
                ElseIf dy >= dx * 2.2 Then
                    If outer Then
                        bucket.OuterVertical.Add(c)
                    Else
                        bucket.InnerVertical.Add(c)
                    End If
                Else
                    bucket.Sloped.Add(c)
                End If
            Next
        Catch ex As Exception
            Debug.Print("WARN CollectViewCurves: " & ex.Message)
        End Try
        Return bucket
    End Function

    Private Function FindExtremeCurves(bucket As CurveBucket, horizontalDim As Boolean) As CurvePair
        Dim pair As New CurvePair()
        Dim pool As List(Of DrawingCurve) = If(horizontalDim, bucket.OuterVertical, bucket.OuterHorizontal)
        If pool Is Nothing OrElse pool.Count < 2 Then pool = If(horizontalDim, bucket.InnerVertical, bucket.InnerHorizontal)
        If pool Is Nothing OrElse pool.Count < 2 Then
            Dim all As New List(Of DrawingCurve)()
            If horizontalDim Then
                all.AddRange(bucket.OuterVertical) : all.AddRange(bucket.InnerVertical)
            Else
                all.AddRange(bucket.OuterHorizontal) : all.AddRange(bucket.InnerHorizontal)
            End If
            pool = all
        End If
        If pool Is Nothing OrElse pool.Count < 2 Then Return pair

        Dim minCurve As DrawingCurve = Nothing
        Dim maxCurve As DrawingCurve = Nothing
        Dim minVal As Double = Double.MaxValue
        Dim maxVal As Double = Double.MinValue
        For Each c As DrawingCurve In pool
            Dim center As Double = GetCurveCenterCoordinate(c, horizontalDim)
            If center < minVal Then minVal = center : minCurve = c
            If center > maxVal Then maxVal = center : maxCurve = c
        Next
        pair.First = minCurve
        pair.Second = maxCurve
        Return pair
    End Function

    Private Function FindInnerParallelPairs(bucket As CurveBucket, horizontalDim As Boolean) As CurvePair
        Dim pair As New CurvePair()
        Dim pool As List(Of DrawingCurve) = If(horizontalDim, bucket.InnerVertical, bucket.InnerHorizontal)
        If pool Is Nothing OrElse pool.Count < 2 Then Return pair

        Dim bestGap As Double = -1.0
        For i As Integer = 0 To pool.Count - 2
            For j As Integer = i + 1 To pool.Count - 1
                Dim gap As Double = Math.Abs(GetCurveCenterCoordinate(pool(j), horizontalDim) - GetCurveCenterCoordinate(pool(i), horizontalDim))
                If gap > bestGap Then
                    bestGap = gap
                    pair.First = pool(i)
                    pair.Second = pool(j)
                End If
            Next
        Next
        Return pair
    End Function

    Private Function TryAddOuterInnerDimension(doc As DrawingDocument,
                                               sheet As Sheet,
                                               v As DrawingView,
                                               slot As SlotRect,
                                               horizontalDim As Boolean,
                                               globalDimensionKeys As HashSet(Of String),
                                               dedupeScope As String) As Integer
        If sheet Is Nothing OrElse v Is Nothing Then Return 0
        Try
            Dim bucket As CurveBucket = CollectViewCurves(v)
            Dim pair As CurvePair = FindOuterInnerPairByRange(bucket, horizontalDim)
            If pair Is Nothing OrElse pair.First Is Nothing OrElse pair.Second Is Nothing Then Return 0
            Dim placed As New HashSet(Of String)(StringComparer.Ordinal)
            Dim dt As DimensionTypeEnum = If(horizontalDim, DimensionTypeEnum.kHorizontalDimensionType, DimensionTypeEnum.kVerticalDimensionType)
            Return AddLinearPairDimension(doc, sheet, v, slot, pair, dt, True, placed, globalDimensionKeys, dedupeScope)
        Catch ex As Exception
            Debug.Print("WARN TryAddOuterInnerDimension: " & ex.Message)
        End Try
        Return 0
    End Function

    Private Function FindOuterInnerPairByRange(bucket As CurveBucket, horizontalDim As Boolean) As CurvePair
        Dim pair As New CurvePair()
        If bucket Is Nothing Then Return pair

        Dim outers As List(Of DrawingCurve) = If(horizontalDim, bucket.OuterVertical, bucket.OuterHorizontal)
        Dim inners As List(Of DrawingCurve) = If(horizontalDim, bucket.InnerVertical, bucket.InnerHorizontal)
        If outers Is Nothing OrElse inners Is Nothing OrElse outers.Count = 0 OrElse inners.Count = 0 Then Return pair

        Dim bestGap As Double = Double.MaxValue
        For Each o As DrawingCurve In outers
            Dim oc As Double = GetCurveCenterCoordinate(o, horizontalDim)
            For Each i As DrawingCurve In inners
                Dim ic As Double = GetCurveCenterCoordinate(i, horizontalDim)
                Dim gap As Double = Math.Abs(oc - ic)
                If gap > 0.0001 AndAlso gap < bestGap Then
                    bestGap = gap
                    pair.First = o
                    pair.Second = i
                End If
            Next
        Next
        Return pair
    End Function

    Private Function TryAddSlopeHeightByExtremums(doc As DrawingDocument, sheet As Sheet, v As DrawingView, slot As SlotRect, globalDimensionKeys As HashSet(Of String), dedupeScope As String) As Integer
        Try
            Dim bucket As CurveBucket = CollectViewCurves(v)
            If bucket Is Nothing Then Return 0
            Dim allCurves As New List(Of DrawingCurve)()
            allCurves.AddRange(bucket.OuterHorizontal)
            allCurves.AddRange(bucket.InnerHorizontal)
            allCurves.AddRange(bucket.Sloped)
            If allCurves.Count < 2 Then Return 0

            Dim low As DrawingCurve = Nothing
            Dim high As DrawingCurve = Nothing
            Dim lowY As Double = Double.MaxValue
            Dim highY As Double = Double.MinValue
            For Each c As DrawingCurve In allCurves
                Dim y As Double = GetCurveCenterCoordinate(c, False)
                If y < lowY Then lowY = y : low = c
                If y > highY Then highY = y : high = c
            Next

            If low Is Nothing OrElse high Is Nothing Then Return 0
            Dim pair As New CurvePair()
            pair.First = low
            pair.Second = high
            Dim placed As New HashSet(Of String)(StringComparer.Ordinal)
            Return AddLinearPairDimension(doc, sheet, v, slot, pair, DimensionTypeEnum.kVerticalDimensionType, True, placed, globalDimensionKeys, dedupeScope)
        Catch ex As Exception
            Debug.Print("WARN TryAddSlopeHeightByExtremums: " & ex.Message)
        End Try
        Return 0
    End Function

    Private Function TryAddSlopedProjectionDimension(doc As DrawingDocument, sheet As Sheet, v As DrawingView, slot As SlotRect, globalDimensionKeys As HashSet(Of String), dedupeScope As String) As Integer
        Try
            Dim bucket As CurveBucket = CollectViewCurves(v)
            If bucket Is Nothing OrElse bucket.Sloped Is Nothing OrElse bucket.Sloped.Count = 0 Then Return 0
            Dim minC As DrawingCurve = Nothing
            Dim maxC As DrawingCurve = Nothing
            Dim minX As Double = Double.MaxValue
            Dim maxX As Double = Double.MinValue
            For Each c As DrawingCurve In bucket.Sloped
                Dim x As Double = GetCurveCenterCoordinate(c, True)
                If x < minX Then minX = x : minC = c
                If x > maxX Then maxX = x : maxC = c
            Next
            If minC Is Nothing OrElse maxC Is Nothing Then Return 0

            Dim pair As New CurvePair()
            pair.First = minC
            pair.Second = maxC
            Dim placed As New HashSet(Of String)(StringComparer.Ordinal)
            Return AddLinearPairDimension(doc, sheet, v, slot, pair, DimensionTypeEnum.kHorizontalDimensionType, True, placed, globalDimensionKeys, dedupeScope)
        Catch ex As Exception
            Debug.Print("WARN TryAddSlopedProjectionDimension: " & ex.Message)
        End Try
        Return 0
    End Function

    Private Function FindFeatureOffsets(bucket As CurveBucket, horizontalDim As Boolean, Optional allowOuterCurves As Boolean = True) As List(Of CurvePair)
        Dim result As New List(Of CurvePair)()
        Dim pool As New List(Of DrawingCurve)()
        If horizontalDim Then
            pool.AddRange(bucket.InnerVertical)
            If allowOuterCurves Then pool.AddRange(bucket.OuterVertical)
        Else
            pool.AddRange(bucket.InnerHorizontal)
            If allowOuterCurves Then pool.AddRange(bucket.OuterHorizontal)
        End If
        If pool.Count < 3 Then Return result

        pool.Sort(Function(a As DrawingCurve, b As DrawingCurve) GetCurveCenterCoordinate(a, horizontalDim).CompareTo(GetCurveCenterCoordinate(b, horizontalDim)))
        For i As Integer = 0 To pool.Count - 2
            Dim pair As New CurvePair()
            pair.First = pool(i)
            pair.Second = pool(i + 1)
            result.Add(pair)
            If result.Count >= 2 Then Exit For
        Next
        Return result
    End Function

    Private Function AddLinearPairDimension(doc As DrawingDocument,
                                            sheet As Sheet,
                                            v As DrawingView,
                                            slot As SlotRect,
                                            pair As CurvePair,
                                            dimType As DimensionTypeEnum,
                                            localOffset As Boolean,
                                            placedPairs As HashSet(Of String),
                                            Optional globalDimensionKeys As HashSet(Of String) = Nothing,
                                            Optional dedupeScope As String = "") As Integer
        If pair Is Nothing OrElse pair.First Is Nothing OrElse pair.Second Is Nothing Then Return 0
        Dim key As String = BuildCurvePairKey(pair.First, pair.Second)
        If placedPairs.Contains(key) Then Return 0
        Dim globalKey As String = dedupeScope & "|" & CInt(dimType).ToString() & "|" & key
        If globalDimensionKeys IsNot Nothing AndAlso globalDimensionKeys.Contains(globalKey) Then Return 0
        Try
            Dim i1 As GeometryIntent = sheet.CreateGeometryIntent(pair.First, PointIntentEnum.kMidPointIntent)
            Dim i2 As GeometryIntent = sheet.CreateGeometryIntent(pair.Second, PointIntentEnum.kMidPointIntent)
            Dim p As Point2d = Nothing
            If dimType = DimensionTypeEnum.kHorizontalDimensionType Then
                Dim py As Double = If(localOffset, Math.Max(slot.B + Cm(doc, 1.5), v.Top - v.Height - Cm(doc, 1.8)), Math.Min(slot.T - Cm(doc, 2.0), v.Top + Cm(doc, 3.5)))
                p = _app.TransientGeometry.CreatePoint2d((slot.L + slot.R) / 2.0, py)
            Else
                Dim px As Double = If(localOffset, Math.Max(slot.L + Cm(doc, 1.5), v.Left - Cm(doc, 1.8)), Math.Min(slot.R - Cm(doc, 1.0), v.Left + v.Width + Cm(doc, 2.6)))
                p = _app.TransientGeometry.CreatePoint2d(px, (slot.B + slot.T) / 2.0)
            End If
            sheet.DrawingDimensions.GeneralDimensions.AddLinear(p, i1, i2, dimType)
            placedPairs.Add(key)
            If globalDimensionKeys IsNot Nothing Then globalDimensionKeys.Add(globalKey)
            Return 1
        Catch ex As Exception
            Debug.Print("WARN AddLinearPairDimension: " & ex.Message)
        End Try
        Return 0
    End Function

    Private Function GetCurveCenterCoordinate(c As DrawingCurve, horizontalDim As Boolean) As Double
        If c Is Nothing Then Return 0.0
        Dim rb As Box2d = c.RangeBox
        If horizontalDim Then Return (rb.MinPoint.X + rb.MaxPoint.X) / 2.0
        Return (rb.MinPoint.Y + rb.MaxPoint.Y) / 2.0
    End Function

    Private Function BuildCurvePairKey(a As DrawingCurve, b As DrawingCurve) As String
        Dim h1 As Integer = RuntimeHelpers.GetHashCode(a)
        Dim h2 As Integer = RuntimeHelpers.GetHashCode(b)
        If h1 < h2 Then Return h1.ToString() & "|" & h2.ToString()
        Return h2.ToString() & "|" & h1.ToString()
    End Function

    Private Function AddFallbackDimensionNotes(doc As DrawingDocument, sheet As Sheet,
                                               v As DrawingView, slot As SlotRect,
                                               intent As DimensionIntentId,
                                               role As ViewRole,
                                               measure As ViewMeasure,
                                               modelSize As ModelOverallExtents,
                                               noteKeys As HashSet(Of String)) As Boolean
        Try
            If sheet Is Nothing OrElse v Is Nothing Then Return False
            Dim text As String = BuildFallbackIntentText(intent, role, modelSize)
            If String.IsNullOrWhiteSpace(text) Then Return False

            Dim minX As Double = Math.Min(slot.L, slot.R)
            Dim maxX As Double = Math.Max(slot.L, slot.R)
            Dim minY As Double = Math.Min(slot.B, slot.T)
            Dim maxY As Double = Math.Max(slot.B, slot.T)

            Dim px As Double = v.Left + v.Width * 0.5
            Dim py As Double = v.Top - v.Height * 0.5

            If IsRadiusIntent(intent) Then
                px = v.Left + v.Width + Cm(doc, 1.6)
                py = v.Top + Cm(doc, 1.0)
            ElseIf IsHorizontalIntent(intent) Then
                px = v.Left + v.Width * 0.5
                If (RuntimeHelpers.GetHashCode(v) Mod 2) = 0 Then
                    py = v.Top + Cm(doc, 1.2)
                Else
                    py = v.Top - v.Height - Cm(doc, 1.2)
                End If
            ElseIf IsVerticalIntent(intent) Then
                py = v.Top - v.Height * 0.5
                px = v.Left + v.Width + Cm(doc, 1.5)
            End If

            px = Math.Max(minX + Cm(doc, 1.0), Math.Min(maxX - Cm(doc, 1.0), px))
            py = Math.Max(minY + Cm(doc, 0.8), Math.Min(maxY - Cm(doc, 0.8), py))

            Dim viewId As String = RuntimeHelpers.GetHashCode(v).ToString()
            Dim dedupe As String = viewId & "|" & intent.ToString() & "|" & text
            If IsOverallIntent(intent) Then
                Dim valueMm As Integer = 0
                If Not TryGetFallbackOverallValueMm(intent, modelSize, valueMm) Then Return False
                Dim kind As String = ResolveOverallPhysicalDimensionKind(intent, role)
                dedupe = "overall|" & intent.ToString() & "|" & kind & "|" & valueMm.ToString() & "|" & viewId
            End If
            If noteKeys IsNot Nothing AndAlso noteKeys.Contains(dedupe) Then Return False

            Dim notePt As Point2d = _app.TransientGeometry.CreatePoint2d(px, py)
            sheet.DrawingNotes.GeneralNotes.AddFitted(notePt, text)
            If noteKeys IsNot Nothing Then noteKeys.Add(dedupe)
            Return True
        Catch ex As Exception
            Debug.Print("WARN AddFallbackDimensionNotes: " & ex.Message)
        End Try
        Return False
    End Function

    Private Function BuildFallbackIntentText(intent As DimensionIntentId,
                                             role As ViewRole,
                                             modelSize As ModelOverallExtents) As String
        Dim overallKind As String = ResolveOverallPhysicalDimensionKind(intent, role)
        If overallKind <> "unknown" Then
            Dim valueMm As Integer = 0
            If TryGetFallbackOverallValueMm(intent, modelSize, valueMm) Then
                Return valueMm.ToString() & " мм"
            End If
            Return String.Empty
        End If

        Select Case intent
            Case DimensionIntentId.RadiusMain, DimensionIntentId.RadiusSecondary, DimensionIntentId.VisibleRadius,
                 DimensionIntentId.StepHeight, DimensionIntentId.SlopeHeightHigh, DimensionIntentId.SlopeHeightLow,
                 DimensionIntentId.LipDepth, DimensionIntentId.LipHeight, DimensionIntentId.RecessOffset,
                 DimensionIntentId.RecessDepth, DimensionIntentId.Chamfer, DimensionIntentId.EndCutLength,
                 DimensionIntentId.EdgeBandWidth, DimensionIntentId.ProfileHeight, DimensionIntentId.ProfileDepth
                Return String.Empty
            Case Else
                Return String.Empty
        End Select
    End Function

    Private Function TryGetFallbackOverallValueMm(intent As DimensionIntentId,
                                                   modelSize As ModelOverallExtents,
                                                   ByRef valueMm As Integer) As Boolean
        valueMm = 0
        If modelSize Is Nothing OrElse Not modelSize.IsValid Then Return False
        Select Case intent
            Case DimensionIntentId.OverallLength, DimensionIntentId.ChordOrSpan
                valueMm = modelSize.LengthMm
            Case DimensionIntentId.OverallWidth, DimensionIntentId.OverallHeight
                valueMm = modelSize.WidthMm
            Case DimensionIntentId.OverallThickness
                valueMm = modelSize.ThicknessMm
            Case Else
                Return False
        End Select
        Return (valueMm > 0)
    End Function

    Private Function ResolveOverallPhysicalDimensionKind(intent As DimensionIntentId,
                                                         role As ViewRole) As String
        Select Case intent
            Case DimensionIntentId.OverallLength, DimensionIntentId.ChordOrSpan
                Return "length"
            Case DimensionIntentId.OverallWidth, DimensionIntentId.OverallHeight
                Return "width"
            Case DimensionIntentId.OverallThickness
                Return "thickness"
            Case Else
                Return "unknown"
        End Select
    End Function

    Private Function ForceVisibleFallbackDimensions(doc As DrawingDocument,
                                                    sheet As Sheet,
                                                    placedViews As Dictionary(Of ViewRole, DrawingView),
                                                    roleMap As RoleMap,
                                                    plan As DimensionPlan,
                                                    modelSize As ModelOverallExtents) As Integer
        Dim added As Integer = 0
        Dim noteKeys As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        For Each kvp As KeyValuePair(Of ViewRole, DrawingView) In placedViews
            If kvp.Key = ViewRole.IsoReference Then Continue For
            If kvp.Value Is Nothing Then Continue For
            Dim v As DrawingView = kvp.Value
            Dim slot As SlotRect = New SlotRect(v.Left, v.Left + v.Width, v.Top - v.Height, v.Top)
            Dim m As ViewMeasure = If(roleMap IsNot Nothing, roleMap.GetMeasure(kvp.Key), Nothing)

            ' When real 3D extents are known, do not re-introduce projected overall dimensions here.
            If modelSize Is Nothing OrElse Not modelSize.IsValid Then
                Dim realAdded As Integer = TryAddTrueDimensions(doc, sheet, v, slot, True, True, False, False)
                added += realAdded
                If realAdded > 0 Then Continue For
            End If

            ' Fallback notes for the two most important intents
            If AddFallbackDimensionNotes(doc, sheet, v, slot, DimensionIntentId.OverallLength, kvp.Key, m, modelSize, noteKeys) Then added += 1
            If AddFallbackDimensionNotes(doc, sheet, v, slot, DimensionIntentId.OverallHeight, kvp.Key, m, modelSize, noteKeys) Then added += 1
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

            Dim fallbackStyle As DrawingViewStyleEnum = measure.Style
            If measure.Orientation = ViewOrientationTypeEnum.kIsoTopRightViewOrientation AndAlso measure.Style = DrawingViewStyleEnum.kShadedDrawingViewStyle Then
                fallbackStyle = DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle
            End If

            Dim v As DrawingView = TryCreateBaseView(sheet,
                                                     modelDoc,
                                                     pt,
                                                     fit.Scale,
                                                     measure.Orientation,
                                                     measure.Style,
                                                     fallbackStyle)

            If v Is Nothing Then Return Nothing

            Dim placedName As String = "__PLACED__" & measure.Key & "__" & Guid.NewGuid().ToString("N").Substring(0, 8)
            Try
                v.Name = placedName
            Catch
            End Try
            v = TryGetViewByName(sheet, placedName)
            If v Is Nothing Then Return Nothing

            If Not WaitForViewReady(v, 5000, False) Then
                Debug.Print("WARN final view model='" & GetModelPathForLog(modelDoc) & "' orientation=" & measure.Orientation.ToString() & " style=" & measure.Style.ToString() & " stage=wait timeout=5000")
                Return Nothing
            End If

            ' Rotate 90° if FitResult requested it
            If fit.Rotate90 Then
                Try
                    v.Rotation = Math.PI / 2.0
                Catch
                End Try

                v = TryGetViewByName(sheet, placedName)
                If v Is Nothing Then Return Nothing
                If Not WaitForViewReady(v, 2000, False) Then
                    Debug.Print("WARN final view model='" & GetModelPathForLog(modelDoc) & "' orientation=" & measure.Orientation.ToString() & " style=" & measure.Style.ToString() & " stage=wait-after-rotate timeout=2000")
                End If
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
        Dim probeName As String = String.Empty
        Dim createdStyle As DrawingViewStyleEnum = style
        Dim modelPath As String = GetModelPathForLog(modelDoc)
        Dim orientationLabel As String = orientation.ToString()
        Dim doc As DrawingDocument = TryCast(sheet.Parent, DrawingDocument)
        Try
            Dim pt As Point2d = GetProbePoint(doc, sheet)
            probeView = TryCreateProbeBaseView(sheet,
                                               modelDoc,
                                               pt,
                                               PROBE_SCALE,
                                               orientation,
                                               style,
                                               key,
                                               modelPath,
                                               createdStyle)
            If probeView Is Nothing Then
                Debug.Print("WARN probe model='" & modelPath & "' orientation=" & orientationLabel & " style=" & style.ToString() & " stage=create error=base view was not created")
                Return Nothing
            End If

            probeName = "__PROBE__" & key & "__" & Guid.NewGuid().ToString("N").Substring(0, 8)
            Try
                probeView.Name = probeName
            Catch exName As Exception
                Debug.Print("WARN probe model='" & modelPath & "' orientation=" & orientationLabel & " style=" & createdStyle.ToString() & " stage=create error=name assign failed: " & exName.Message)
            End Try

            Try
                If doc IsNot Nothing Then doc.Update2(True)
            Catch exUpdate As Exception
                Debug.Print("WARN probe model='" & modelPath & "' orientation=" & orientationLabel & " style=" & createdStyle.ToString() & " stage=update error=" & exUpdate.Message)
            End Try

            If Not String.IsNullOrEmpty(probeName) Then
                probeView = TryGetViewByName(sheet, probeName)
            End If
            If probeView Is Nothing Then
                Debug.Print("WARN probe model='" & modelPath & "' orientation=" & orientationLabel & " style=" & createdStyle.ToString() & " stage=wait error=probe view reference lost")
                Return Nothing
            End If

            If Not WaitForViewReady(probeView, 5000, False) Then
                Debug.Print("WARN probe model='" & modelPath & "' orientation=" & orientationLabel & " style=" & createdStyle.ToString() & " stage=wait timeout=5000")
                Return Nothing
            End If

            Dim m As New ViewMeasure()
            m.Key = key
            m.Caption = caption
            m.Orientation = orientation
            m.Style = createdStyle
            Try
                m.UnitW = probeView.Width / PROBE_SCALE
                m.UnitH = probeView.Height / PROBE_SCALE
                m.BoundingArea = m.UnitW * m.UnitH
                If m.UnitH > 0 Then
                    m.AspectRatio = m.UnitW / m.UnitH
                Else
                    m.AspectRatio = 1.0
                End If
            Catch exSize As Exception
                Debug.Print("WARN probe model='" & modelPath & "' orientation=" & orientationLabel & " style=" & createdStyle.ToString() & " stage=measure-size error=" & exSize.Message)
                Return Nothing
            End Try
            If m.UnitW <= 0.0001 OrElse m.UnitH <= 0.0001 Then
                Debug.Print("WARN probe model='" & modelPath & "' orientation=" & orientationLabel & " style=" & createdStyle.ToString() & " stage=measure-size error=non-positive size w=" & m.UnitW.ToString("0.####") & " h=" & m.UnitH.ToString("0.####"))
                Return Nothing
            End If

            ' Curve analysis should not abort successful size probe.
            Try
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
                    Debug.Print("INFO probe success key=" & key & " uw=" & m.UnitW.ToString("0.####") & " uh=" & m.UnitH.ToString("0.####") & " curves=" & m.CurveCount.ToString())
                Else
                    Debug.Print("INFO probe success key=" & key & " uw=" & m.UnitW.ToString("0.####") & " uh=" & m.UnitH.ToString("0.####") & " curves=<none>")
                End If
            Catch exCurve As Exception
                Debug.Print("WARN probe model='" & modelPath & "' orientation=" & orientationLabel & " style=" & createdStyle.ToString() & " stage=measure-curves error=" & exCurve.Message)
                Debug.Print("INFO probe success key=" & key & " uw=" & m.UnitW.ToString("0.####") & " uh=" & m.UnitH.ToString("0.####") & " curves=<failed>")
            End Try
            Return m
        Catch ex As Exception
            Debug.Print("WARN probe model='" & modelPath & "' orientation=" & orientationLabel & " style=" & createdStyle.ToString() & " stage=create error=" & ex.Message)
        Finally
            If Not String.IsNullOrEmpty(probeName) Then
                probeView = TryGetViewByName(sheet, probeName)
            End If
            If probeView IsNot Nothing Then
                Try
                    probeView.Delete()
                Catch
                End Try
            End If
        End Try
        Return Nothing
    End Function

    Private Function GetModelPathForLog(modelDoc As Document) As String
        If modelDoc Is Nothing Then Return "<null>"
        Try
            If Not String.IsNullOrEmpty(modelDoc.FullFileName) Then Return modelDoc.FullFileName
        Catch
        End Try
        Try
            If Not String.IsNullOrEmpty(modelDoc.DisplayName) Then Return modelDoc.DisplayName
        Catch
        End Try
        Return "<unknown>"
    End Function

    Private Function TryGetViewByName(sheet As Sheet, name As String) As DrawingView
        If sheet Is Nothing OrElse String.IsNullOrEmpty(name) Then Return Nothing
        Try
            For i As Integer = 1 To sheet.DrawingViews.Count
                Dim v As DrawingView = sheet.DrawingViews.Item(i)
                If v Is Nothing Then Continue For
                If String.Equals(v.Name, name, StringComparison.OrdinalIgnoreCase) Then Return v
            Next
        Catch
        End Try
        Return Nothing
    End Function

    Private Function WaitForViewReady(v As DrawingView,
                                      timeoutMs As Integer,
                                      Optional requirePrecise As Boolean = False) As Boolean
        If v Is Nothing Then Return False
        Dim startTick As Integer = System.Environment.TickCount
        Do
            Try
                If v IsNot Nothing AndAlso v.IsUpdateComplete Then
                    Dim w As Double = v.Width
                    Dim h As Double = v.Height
                    If w > 0.0001 AndAlso h > 0.0001 Then
                        If requirePrecise Then
                            Try
                                If Not v.IsRasterView Then Return True
                            Catch
                                Return True
                            End Try
                        Else
                            Return True
                        End If
                    End If
                End If
            Catch
                ' COM can be transient while view rebuilds.
            End Try

            System.Windows.Forms.Application.DoEvents()
            System.Threading.Thread.Sleep(50)
        Loop While (System.Environment.TickCount - startTick) < timeoutMs

        If requirePrecise Then
            Try
                If v IsNot Nothing AndAlso v.IsUpdateComplete AndAlso v.Width > 0.0001 AndAlso v.Height > 0.0001 Then Return True
            Catch
            End Try
        End If
        Return False
    End Function

    Private Function GetProbeStyleCandidates(orientation As ViewOrientationTypeEnum,
                                             preferredStyle As DrawingViewStyleEnum) As List(Of DrawingViewStyleEnum)
        Dim styles As New List(Of DrawingViewStyleEnum)()
        If orientation = ViewOrientationTypeEnum.kIsoTopRightViewOrientation Then
            styles.Add(DrawingViewStyleEnum.kShadedDrawingViewStyle)
            styles.Add(DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle)
            styles.Add(DrawingViewStyleEnum.kHiddenLineDrawingViewStyle)
        Else
            styles.Add(DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle)
            styles.Add(DrawingViewStyleEnum.kHiddenLineDrawingViewStyle)
        End If
        If Not styles.Contains(preferredStyle) Then styles.Insert(0, preferredStyle)
        Return styles
    End Function

    Private Function TryCreateProbeBaseView(sheet As Sheet,
                                            modelDoc As Document,
                                            pt As Point2d,
                                            scale As Double,
                                            orientation As ViewOrientationTypeEnum,
                                            preferredStyle As DrawingViewStyleEnum,
                                            key As String,
                                            modelPath As String,
                                            ByRef usedStyle As DrawingViewStyleEnum) As DrawingView
        Dim styles As List(Of DrawingViewStyleEnum) = GetProbeStyleCandidates(orientation, preferredStyle)
        For Each candidateStyle As DrawingViewStyleEnum In styles
            Try
                Dim v As DrawingView = sheet.DrawingViews.AddBaseView(TryCast(modelDoc, Inventor.Document),
                                                                      pt,
                                                                      scale,
                                                                      orientation,
                                                                      candidateStyle)
                usedStyle = candidateStyle
                Return v
            Catch ex As Exception
                Debug.Print("WARN probe model='" & modelPath & "' orientation=" & orientation.ToString() & " style=" & candidateStyle.ToString() & " stage=create error=" & ex.Message)
            End Try
        Next
        Return Nothing
    End Function

    Private Function GetProbePoint(doc As DrawingDocument, sheet As Sheet) As Point2d
        Dim safe As SlotRect = GetSheetSafeRect(doc, sheet)
        Dim px As Double = safe.L + RectW(safe) * 0.06
        Dim py As Double = safe.B + RectH(safe) * 0.06
        Return _app.TransientGeometry.CreatePoint2d(px, py)
    End Function

    Private Function TryCreateBaseView(sheet As Sheet,
                                       modelDoc As Document,
                                       pt As Point2d,
                                       scale As Double,
                                       orientation As ViewOrientationTypeEnum,
                                       primaryStyle As DrawingViewStyleEnum,
                                       fallbackStyle As DrawingViewStyleEnum) As DrawingView
        Try
            Return sheet.DrawingViews.AddBaseView(TryCast(modelDoc, Inventor.Document),
                                                  pt,
                                                  scale,
                                                  orientation,
                                                  primaryStyle)
        Catch ex As Exception
            If primaryStyle <> fallbackStyle Then
                Debug.Print("WARN base view style " & primaryStyle.ToString() & " failed, fallback " & fallbackStyle.ToString() & ": " & ex.Message)
                Try
                    Return sheet.DrawingViews.AddBaseView(TryCast(modelDoc, Inventor.Document),
                                                          pt,
                                                          scale,
                                                          orientation,
                                                          fallbackStyle)
                Catch ex2 As Exception
                    Debug.Print("WARN base view fallback failed " & orientation.ToString() & ": " & ex2.Message)
                End Try
            Else
                Debug.Print("WARN base view create failed " & orientation.ToString() & "/" & primaryStyle.ToString() & ": " & ex.Message)
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
        Dim safeR As Double = ww - fo
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


    Private Function GetModelOverallExtentsMm(modelDoc As Document) As ModelOverallExtents
        Dim result As New ModelOverallExtents()
        If modelDoc Is Nothing Then Return result
        Try
            Dim rb As Box = Nothing
            If TypeOf modelDoc Is PartDocument Then
                rb = DirectCast(modelDoc, PartDocument).ComponentDefinition.RangeBox
            ElseIf TypeOf modelDoc Is AssemblyDocument Then
                rb = DirectCast(modelDoc, AssemblyDocument).ComponentDefinition.RangeBox
            End If
            If rb Is Nothing Then Return result

            Dim sx As Double = Math.Abs(rb.MaxPoint.X - rb.MinPoint.X) * 10.0
            Dim sy As Double = Math.Abs(rb.MaxPoint.Y - rb.MinPoint.Y) * 10.0
            Dim sz As Double = Math.Abs(rb.MaxPoint.Z - rb.MinPoint.Z) * 10.0

            result.Xmm = CInt(Math.Round(sx))
            result.Ymm = CInt(Math.Round(sy))
            result.Zmm = CInt(Math.Round(sz))

            Dim a As Integer = result.Xmm
            Dim b As Integer = result.Ymm
            Dim c As Integer = result.Zmm
            Dim tmp As Integer
            If a < b Then tmp = a : a = b : b = tmp
            If b < c Then tmp = b : b = c : c = tmp
            If a < b Then tmp = a : a = b : b = tmp

            result.LengthMm = a
            result.WidthMm = b
            result.ThicknessMm = c
            result.IsValid = (result.LengthMm > 0 AndAlso result.WidthMm > 0 AndAlso result.ThicknessMm > 0)
        Catch ex As Exception
            Debug.Print("WARN GetModelOverallExtentsMm: " & ex.Message)
        End Try
        Return result
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
            Dim target As BorderDefinition = existing
            If target Is Nothing Then
                Try
                    _app.SilentOperation = True
                    target = doc.BorderDefinitions.Add(BORDER_NAME)
                Finally
                    _app.SilentOperation = False
                End Try
            End If
            If target Is Nothing Then
                Debug.Print("WARN EnsureBorder: definition '" & BORDER_NAME & "' could not be created.")
                Return
            End If

            Dim sk As DrawingSketch = Nothing
            target.Edit(sk)
            Try
                ClearSketch(sk)
                DrawBorderDefinition(doc, sk)
            Finally
                target.ExitEdit(True)
            End Try
            If existing Is Nothing Then
                Debug.Print("EnsureBorder: created: " & BORDER_NAME)
            Else
                Debug.Print("EnsureBorder: refreshed: " & BORDER_NAME)
            End If
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

            Dim created As TitleBlockDefinition = Nothing
            Try
                _app.SilentOperation = True
                created = doc.TitleBlockDefinitions.Add(TB_NAME)
            Finally
                _app.SilentOperation = False
            End Try
            If created Is Nothing Then
                Debug.Print("WARN EnsureTitleBlock: definition '" & TB_NAME & "' could not be created.")
                Return
            End If

            Dim sk As DrawingSketch = Nothing
            created.Edit(sk)
            Try
                ClearSketch(sk)
                DrawTitleBlockDefinition(doc, sk)
            Finally
                created.ExitEdit(True)
            End Try
            Debug.Print("EnsureTitleBlock: created: " & TB_NAME)
        Catch ex As Exception
            Debug.Print("WARN EnsureTitleBlock: " & ex.Message)
        End Try
    End Sub

    Private Sub ClearSketch(sk As DrawingSketch)
        If sk Is Nothing Then Return
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
    End Sub

    Private Sub DrawBorderDefinition(doc As DrawingDocument, sk As DrawingSketch)
        sk.SketchLines.AddByTwoPoints(P(0, 0), P(0.0001, 0.0001))
        sk.SketchLines.AddByTwoPoints(P(Cm(doc, A3_W_MM), Cm(doc, A3_H_MM)), P(Cm(doc, A3_W_MM) - 0.0001, Cm(doc, A3_H_MM) - 0.0001))
        sk.SketchLines.AddAsTwoPointRectangle(P(Cm(doc, FRAME_L_MM), Cm(doc, FRAME_O_MM)), P(Cm(doc, A3_W_MM - FRAME_O_MM), Cm(doc, A3_H_MM - FRAME_O_MM)))
    End Sub

    Private Sub DrawTitleBlockDefinition(doc As DrawingDocument, sk As DrawingSketch)
        ' Фикс смещения влево: штамп привязываем к правой кромке листа.
        Dim x2 As Double = 0.0
        Dim y1 As Double = Cm(doc, FRAME_O_MM)
        Dim x1 As Double = x2 - Cm(doc, TB_W_MM)
        Dim y2 As Double = y1 + Cm(doc, TB_H_MM)

        sk.SketchLines.AddByTwoPoints(P(0, 0), P(-0.0001, 0.0001))
        sk.SketchLines.AddAsTwoPointRectangle(P(x1, y1), P(x2, y2))

        VL(doc, sk, x1, y1, 7, 0, 55) : VL(doc, sk, x1, y1, 17, 0, 55)
        VL(doc, sk, x1, y1, 27, 0, 55) : VL(doc, sk, x1, y1, 42, 0, 55)
        VL(doc, sk, x1, y1, 57, 0, 55) : VL(doc, sk, x1, y1, 67, 0, 55)
        VL(doc, sk, x1, y1, 137, 0, 40) : VL(doc, sk, x1, y1, 152, 15, 40)
        VL(doc, sk, x1, y1, 167, 15, 40)

        Dim y As Double
        For y = 5.0 To 30.0 Step 5.0
            HL(doc, sk, x1, y1, 0, 67, y)
        Next
        HL(doc, sk, x1, y1, 0, 185, 15) : HL(doc, sk, x1, y1, 0, 67, 35)
        HL(doc, sk, x1, y1, 137, 185, 35) : HL(doc, sk, x1, y1, 0, 185, 40)
        HL(doc, sk, x1, y1, 0, 67, 45) : HL(doc, sk, x1, y1, 0, 67, 50)

        Lbl(doc, sk, x1, y1, 0, 35, 7, 40, "Изм.")
        Lbl(doc, sk, x1, y1, 7, 35, 17, 40, "Кол.уч")
        Lbl(doc, sk, x1, y1, 17, 35, 27, 40, "Лист")
        Lbl(doc, sk, x1, y1, 27, 35, 42, 40, "№ doc.")
        Lbl(doc, sk, x1, y1, 42, 35, 57, 40, "Подп.")
        Lbl(doc, sk, x1, y1, 57, 35, 67, 40, "Дата")
        Lbl(doc, sk, x1, y1, 137, 35, 152, 40, "Стадия")
        Lbl(doc, sk, x1, y1, 152, 35, 167, 40, "Лист")
        Lbl(doc, sk, x1, y1, 167, 35, 185, 40, "Листов")

        Prm(doc, sk, x1, y1, 67, 40, 185, 55, "CODE")
        Prm(doc, sk, x1, y1, 67, 15, 137, 40, "PROJECT_NAME")
        Prm(doc, sk, x1, y1, 67, 0, 137, 15, "DRAWING_NAME")
        Prm(doc, sk, x1, y1, 137, 0, 185, 15, "ORG_NAME")
        Prm(doc, sk, x1, y1, 137, 15, 152, 35, "STAGE")
        Prm(doc, sk, x1, y1, 152, 15, 167, 35, "SHEET")
        Prm(doc, sk, x1, y1, 167, 15, 185, 35, "SHEETS")
    End Sub

    Private Function P(x As Double, y As Double) As Point2d
        Return _app.TransientGeometry.CreatePoint2d(x, y)
    End Function

    Private Sub VL(doc As DrawingDocument, sk As DrawingSketch, x1 As Double, y1 As Double, dxMm As Double, y0Mm As Double, y1Mm As Double)
        sk.SketchLines.AddByTwoPoints(P(x1 + Cm(doc, dxMm), y1 + Cm(doc, y0Mm)), P(x1 + Cm(doc, dxMm), y1 + Cm(doc, y1Mm)))
    End Sub

    Private Sub HL(doc As DrawingDocument, sk As DrawingSketch, x1 As Double, y1 As Double, x0Mm As Double, x1Mm As Double, dyMm As Double)
        sk.SketchLines.AddByTwoPoints(P(x1 + Cm(doc, x0Mm), y1 + Cm(doc, dyMm)), P(x1 + Cm(doc, x1Mm), y1 + Cm(doc, dyMm)))
    End Sub

    Private Sub Lbl(doc As DrawingDocument, sk As DrawingSketch, x1 As Double, y1 As Double, x0Mm As Double, y0Mm As Double, x1Mm As Double, y1Mm As Double, caption As String)
        Dim tb As Inventor.TextBox = sk.TextBoxes.AddByRectangle(P(x1 + Cm(doc, x0Mm), y1 + Cm(doc, y0Mm)), P(x1 + Cm(doc, x1Mm), y1 + Cm(doc, y1Mm)), caption)
        tb.HorizontalJustification = HorizontalTextAlignmentEnum.kAlignTextCenter
        tb.VerticalJustification = VerticalTextAlignmentEnum.kAlignTextMiddle
        ApplyTitleTextStyle(doc, tb)
    End Sub

    Private Sub Prm(doc As DrawingDocument, sk As DrawingSketch, x1 As Double, y1 As Double, x0Mm As Double, y0Mm As Double, x1Mm As Double, y1Mm As Double, promptName As String)
        Dim tb As Inventor.TextBox = sk.TextBoxes.AddByRectangle(P(x1 + Cm(doc, x0Mm), y1 + Cm(doc, y0Mm)), P(x1 + Cm(doc, x1Mm), y1 + Cm(doc, y1Mm)), "<Prompt>" & promptName & "</Prompt>")
        tb.HorizontalJustification = HorizontalTextAlignmentEnum.kAlignTextCenter
        tb.VerticalJustification = VerticalTextAlignmentEnum.kAlignTextMiddle
        ApplyTitleTextStyle(doc, tb)
    End Sub

    Private Sub ApplyTitleTextStyle(doc As DrawingDocument, tb As Inventor.TextBox)
        Try
            tb.Style.FontSize = Cm(doc, TITLE_TEXT_HEIGHT_MM)
        Catch
        End Try
    End Sub

    Private Function MakeUniqueSheetName(doc As DrawingDocument, baseName As String) As String
        Dim candidate As String = baseName
        Dim suffix As Integer = 2
        While SheetNameExists(doc, candidate)
            candidate = baseName & "_" & suffix.ToString()
            suffix += 1
        End While
        Return candidate
    End Function

    Private Function SheetNameExists(doc As DrawingDocument, sheetName As String) As Boolean
        For Each existing As Sheet In doc.Sheets
            If String.Equals(existing.Name, sheetName, StringComparison.OrdinalIgnoreCase) Then
                Return True
            End If
        Next
        Return False
    End Function

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

Public Class ModelOverallExtents
    Public Xmm As Integer
    Public Ymm As Integer
    Public Zmm As Integer
    Public LengthMm As Integer
    Public WidthMm As Integer
    Public ThicknessMm As Integer
    Public IsValid As Boolean
End Class

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
    Public SourceModelRaw As String = String.Empty
    Public ExcelRow As Integer = 0
    Public Prompts As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
End Class

Public Class XlsxReader
    Public Shared Function Load(excelPath As String,
                                workspacePath As String,
                                sheetTab As String) As List(Of AlbumItem)
        Dim result As New List(Of AlbumItem)()
        Dim xl As Object = Nothing
        Dim wb As Object = Nothing
        Dim ws As Object = Nothing
        Dim used As Object = Nothing
        Try
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
            Try
                wb = xl.Workbooks.Open(excelPath)
            Catch ex As Exception
                System.Windows.Forms.MessageBox.Show(
                    "Не удалось открыть Excel:" & vbCrLf & ex.Message,
                    "Ошибка XlsxReader")
                Return result
            End Try

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
                Return result
            End If

            used = ws.UsedRange
            Dim firstRow As Integer = CInt(used.Row)
            Dim firstCol As Integer = CInt(used.Column)
            Dim lastRow As Integer = firstRow + CInt(used.Rows.Count) - 1
            Dim lastCol As Integer = firstCol + CInt(used.Columns.Count) - 1
            Dim headerRow As Integer = 0
            Dim canonicalByColumn As New Dictionary(Of Integer, String)()
            For r As Integer = firstRow To Math.Min(firstRow + 9, lastRow)
                Dim rowMap As New Dictionary(Of Integer, String)()
                Dim hasModelPath As Boolean = False
                For c As Integer = firstCol To lastCol
                    Dim raw As String = SafeCellText(ws.Cells(r, c).Value)
                    Dim canonical As String = ResolveHeaderAlias(raw)
                    If Not String.IsNullOrWhiteSpace(canonical) Then
                        rowMap(c) = canonical
                        If String.Equals(canonical, "MODEL_PATH", StringComparison.OrdinalIgnoreCase) Then
                            hasModelPath = True
                        End If
                    End If
                Next
                If hasModelPath Then
                    headerRow = r
                    canonicalByColumn = rowMap
                    Exit For
                End If
            Next

            If headerRow = 0 Then
                Dim hdrErr As String = "Не найден заголовок MODEL_PATH (или алиасы MODELPATH/MODEL/P/ПУТЬ/ФАЙЛ) в первых 10 строках." & vbCrLf &
                                       "Файл: " & excelPath & vbCrLf &
                                       "Лист: " & sheetTab
                Debug.Print("XlsxReader.Load: " & hdrErr)
                System.Windows.Forms.MessageBox.Show(hdrErr, "Ошибка XlsxReader")
                Return result
            End If

            Dim modelIdx As Integer = 0
            For Each kv As KeyValuePair(Of Integer, String) In canonicalByColumn
                If String.Equals(kv.Value, "MODEL_PATH", StringComparison.OrdinalIgnoreCase) Then
                    modelIdx = kv.Key
                    Exit For
                End If
            Next
            If modelIdx = 0 Then
                Dim idxErr As String = "В строке заголовков не определена колонка MODEL_PATH."
                Debug.Print("XlsxReader.Load: " & idxErr)
                System.Windows.Forms.MessageBox.Show(idxErr, "Ошибка XlsxReader")
                Return result
            End If

            For r As Integer = headerRow + 1 To lastRow
                Dim modelRaw As String = SafeCellText(ws.Cells(r, modelIdx).Value)
                If String.IsNullOrWhiteSpace(modelRaw) Then Continue For

                Dim tried As New List(Of String)()
                Dim modelPath As String = ResolveModelPath(modelRaw, workspacePath, excelPath, tried)

                Dim item As New AlbumItem()
                item.ExcelRow = r
                item.SourceModelRaw = modelRaw
                item.ModelPath = modelPath

                For c As Integer = firstCol To lastCol
                    Dim canonical As String = String.Empty
                    If canonicalByColumn.ContainsKey(c) Then
                        canonical = canonicalByColumn(c)
                    End If
                    If String.IsNullOrWhiteSpace(canonical) Then Continue For

                    Dim v As String = SafeCellText(ws.Cells(r, c).Value)
                    If Not String.IsNullOrWhiteSpace(v) Then
                        item.Prompts(canonical) = v
                    End If
                Next

                If String.IsNullOrWhiteSpace(modelPath) Then
                    Debug.Print("XlsxReader.Load unresolved: row=" & r.ToString() & "; source='" & modelRaw & "'; tried=" & String.Join(" | ", tried.ToArray()))
                End If

                result.Add(item)
            Next

        Catch ex As Exception
            Debug.Print("XlsxReader.Load error: " & ex.Message)
        Finally
            If wb IsNot Nothing Then
                Try
                    wb.Close(False)
                Catch
                End Try
            End If
            If xl IsNot Nothing Then
                Try
                    xl.Quit()
                Catch
                End Try
            End If
            ReleaseComObjectSafe(used)
            ReleaseComObjectSafe(ws)
            ReleaseComObjectSafe(wb)
            ReleaseComObjectSafe(xl)
            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
        Return result
    End Function

    Private Shared Sub ReleaseComObjectSafe(comObj As Object)
        If comObj Is Nothing Then Return
        Try
            Marshal.FinalReleaseComObject(comObj)
        Catch
            Try
                Marshal.ReleaseComObject(comObj)
            Catch
            End Try
        End Try
    End Sub

    Private Shared Function SafeCellText(v As Object) As String
        If v Is Nothing Then Return String.Empty
        Dim s As String = CStr(v)
        If s Is Nothing Then Return String.Empty
        Return s.Trim()
    End Function

    Private Shared Function NormalizeHeader(hdr As String) As String
        If hdr Is Nothing Then Return String.Empty
        Dim u As String = hdr.Trim().ToUpperInvariant()
        u = u.Replace(" ", "")
        u = u.Replace("_", "")
        u = u.Replace("-", "")
        u = u.Replace(".", "")
        Return u
    End Function

    Private Shared Function ResolveHeaderAlias(hdr As String) As String
        Dim n As String = NormalizeHeader(hdr)
        If String.IsNullOrWhiteSpace(n) Then Return String.Empty

        Select Case n
            Case "MODELPATH", "MODEL", "P", "ПУТЬ", "ФАЙЛ"
                Return "MODEL_PATH"
            Case "CODE", "ШИФР", "АРТИКУЛ", "ОБОЗНАЧЕНИЕ"
                Return "CODE"
            Case "PROJECTNAME", "PROJECT", "ОБЪЕКТ", "ПРОЕКТ"
                Return "PROJECT_NAME"
            Case "DRAWINGNAME", "TITLE", "НАИМЕНОВАНИЕ"
                Return "DRAWING_NAME"
            Case "ORGNAME", "ОРГАНИЗАЦИЯ", "КОМПАНИЯ"
                Return "ORG_NAME"
            Case "SHEET", "ЛИСТ"
                Return "SHEET"
            Case "SHEETS", "ЛИСТОВ"
                Return "SHEETS"
            Case "STAGE", "СТАДИЯ"
                Return "STAGE"
        End Select

        Return String.Empty
    End Function

    Private Shared Function ResolveModelPath(inputPath As String,
                                             workspacePath As String,
                                             excelPath As String,
                                             ByRef triedCandidates As List(Of String)) As String
        triedCandidates = New List(Of String)()
        Dim raw As String = If(inputPath, String.Empty).Trim()
        If String.IsNullOrWhiteSpace(raw) Then Return String.Empty

        Dim hasExt As Boolean = Not String.IsNullOrWhiteSpace(System.IO.Path.GetExtension(raw))

        Dim candidates As New List(Of String)()
        AddUniqueCandidate(candidates, raw)

        If Not System.IO.Path.IsPathRooted(raw) Then
            If Not String.IsNullOrWhiteSpace(workspacePath) Then
                AddUniqueCandidate(candidates, System.IO.Path.Combine(workspacePath, raw))
            End If

            Dim excelDir As String = String.Empty
            Try
                excelDir = System.IO.Path.GetDirectoryName(excelPath)
            Catch
            End Try
            If Not String.IsNullOrWhiteSpace(excelDir) Then
                AddUniqueCandidate(candidates, System.IO.Path.Combine(excelDir, raw))
            End If
        End If

        If Not hasExt Then
            Dim snapshot As List(Of String) = New List(Of String)(candidates)
            For Each c As String In snapshot
                AddUniqueCandidate(candidates, c & ".ipt")
            Next
        End If

        For Each c As String In candidates
            triedCandidates.Add(c)
            Try
                If System.IO.File.Exists(c) Then
                    Return c
                End If
            Catch
            End Try
        Next

        Return String.Empty
    End Function

    Private Shared Sub AddUniqueCandidate(list As List(Of String), candidate As String)
        If list Is Nothing Then Return
        If String.IsNullOrWhiteSpace(candidate) Then Return
        Dim c As String = candidate.Trim()
        For Each existing As String In list
            If String.Equals(existing, c, StringComparison.OrdinalIgnoreCase) Then
                Return
            End If
        Next
        list.Add(c)
    End Sub
End Class
