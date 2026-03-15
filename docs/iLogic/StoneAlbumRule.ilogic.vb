' ================================================================
' StoneAlbumRule.ilogic.vb  –  v3.19
' Архитектура точно повторяет рабочий VBA RKM_IdwAlbum.bas
' Источник: vba-inventor / RKM_IdwAlbum.bas, RKM_FrameBorder.bas,
'           RKM_TitleBlockPrompted.bas, RKM_Excel.bas
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
    Private Const ADD_VIEW_NOTES As Boolean = True
    Private Const ADD_VIEW_DIMENSIONS As Boolean = True
    Private Const DIMENSIONS_ON_MAIN_ISO As Boolean = False
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

        Dim mFront As ViewMeasure = MeasureView(sheet, modelDoc, ViewOrientationTypeEnum.kFrontViewOrientation, DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle, "FRONT", "Вид спереди")
        Dim mBack As ViewMeasure = MeasureView(sheet, modelDoc, ViewOrientationTypeEnum.kBackViewOrientation, DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle, "BACK", "Вид сзади")
        Dim mTop As ViewMeasure = MeasureView(sheet, modelDoc, ViewOrientationTypeEnum.kTopViewOrientation, DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle, "TOP", "Вид сверху")
        Dim mLeft As ViewMeasure = MeasureView(sheet, modelDoc, ViewOrientationTypeEnum.kLeftViewOrientation, DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle, "LEFT", "Вид слева")
        Dim mRight As ViewMeasure = MeasureView(sheet, modelDoc, ViewOrientationTypeEnum.kRightViewOrientation, DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle, "RIGHT", "Вид справа")
        Dim mIso As ViewMeasure = MeasureView(sheet, modelDoc, ViewOrientationTypeEnum.kIsoTopRightViewOrientation, DrawingViewStyleEnum.kShadedDrawingViewStyle, "ISO", "Изометрия")

        If ALBUM_MODE_VISUAL AndAlso mIso Is Nothing Then
            Debug.Print("WARN: не удалось измерить изометрический shaded вид")
            Return False
        End If
        Dim all2D As List(Of ViewMeasure) = BuildAll2DCandidates(mFront, mBack, mTop, mLeft, mRight)
        If all2D Is Nothing OrElse all2D.Count < 2 Then
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
            Debug.Print("WARN: не найден корректный layout визуализации")
            Return False
        End If

        Debug.Print("Layout template=" & best.TemplateName & ", score=" & String.Format("{0:F3}", best.Score))

        Dim placedViews As Dictionary(Of ViewRole, DrawingView) = PlaceViewsByTemplate(sheet, modelDoc, best)
        If placedViews Is Nothing OrElse placedViews.Count = 0 Then
            Debug.Print("WARN: не удалось разместить виды по template")
            Return False
        End If

        If ADD_VIEW_NOTES Then
            AddViewRoleCaptions(doc, sheet, best, placedViews)
        End If

        If ADD_VIEW_DIMENSIONS Then
            Dim dimPlan As DimensionPlan = BuildDimensionPlan(descriptor, roleMap)
            DebugPrintDimensionPlan(dimPlan)
            ApplyDimensionPlan(doc, sheet, placedViews, roleMap, dimPlan)
        End If

        Return True
    End Function

    Private Sub DebugPrintDescriptor(d As PartDescriptor)
        If d Is Nothing Then Return
        Debug.Print("Part family=" & d.Family.ToString() &
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
            Return d
        End If

        Dim maxAr As Double = 0
        Dim maxSlope As Double = 0
        Dim maxArcRatio As Double = 0
        Dim maxProfile As Double = 0
        Dim maxPlan As Double = 0
        Dim totalArea As Double = 0

        For Each m As ViewMeasure In measures
            If m Is Nothing Then Continue For
            maxAr = Math.Max(maxAr, m.AspectRatio)
            maxSlope = Math.Max(maxSlope, m.SlopeScore)
            maxProfile = Math.Max(maxProfile, m.ProfileComplexityScore)
            maxPlan = Math.Max(maxPlan, m.PlanComplexityScore)
            totalArea += m.BoundingArea
            If m.CurveCount > 0 Then
                maxArcRatio = Math.Max(maxArcRatio, CDbl(m.ArcCount + m.CircleCount) / CDbl(Math.Max(1, m.CurveCount)))
            End If
        Next

        d.IsLong = (maxAr >= 3.6)
        d.IsThin = (maxAr >= 6.0)
        d.HasSlope = (maxSlope >= 0.26)
        d.HasComplexProfile = (maxProfile >= 1.2)
        d.HasRadialPlan = (maxArcRatio >= 0.22)
        d.HasPlanTaper = (maxSlope >= 0.18 AndAlso maxPlan >= 0.95)
        d.HasDominantPlan = HasDominantByKey(measures, "TOP")
        d.HasDominantFacade = HasDominantByKey(measures, "FRONT") OrElse HasDominantByKey(measures, "BACK")
        d.HasDovetailEnds = (maxProfile >= 1.55 AndAlso maxSlope >= 0.18)
        d.HasDecorativeRecess = (maxPlan >= 1.25 AndAlso maxProfile >= 1.25)
        d.HasEdgeRadiusOrDrip = d.HasRadialPlan OrElse maxArcRatio >= 0.15

        If d.HasSlope Then
            d.Family = PartFamily.Sloped
        ElseIf d.HasRadialPlan Then
            d.Family = PartFamily.Radial
        ElseIf d.IsLong Then
            d.Family = PartFamily.Linear
        Else
            d.Family = PartFamily.Plate
        End If

        Return d
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

        Dim used As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim orderedPairs As List(Of AuxPair) = BuildOrderedViewPairs(measures)

        AssignRole(map, ViewRole.PlanContour, PickBestForRole(ViewRole.PlanContour, descriptor, measures, used))
        AssignRole(map, ViewRole.LongitudinalFacade, PickBestForRole(ViewRole.LongitudinalFacade, descriptor, measures, used))
        AssignRole(map, ViewRole.CrossProfile, PickBestForRole(ViewRole.CrossProfile, descriptor, measures, used))

        Dim slopeCandidate As ViewMeasure = PickBestForRole(ViewRole.SlopeView, descriptor, measures, used)
        If slopeCandidate Is Nothing AndAlso descriptor IsNot Nothing AndAlso descriptor.HasSlope AndAlso orderedPairs.Count > 0 Then
            slopeCandidate = orderedPairs(0).A
        End If
        AssignRole(map, ViewRole.SlopeView, slopeCandidate)

        AssignRole(map, ViewRole.ThicknessView, PickBestForRole(ViewRole.ThicknessView, descriptor, measures, used))
        AssignRole(map, ViewRole.EndFace, PickBestForRole(ViewRole.EndFace, descriptor, measures, used))
        AssignRole(map, ViewRole.MainContour, PickMainContour(descriptor, map, measures, used))
        AssignRole(map, ViewRole.IsoReference, isoMeasure)
        Return map
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

        result.Add(NewLayoutTemplate("SLOPED_PLAN_COMPLEX", PartFamily.Sloped, ViewRole.PlanContour, ViewRole.SlopeView, ViewRole.IsoReference,
                                     InsetRect(New SlotRect(safe.L, safe.L + w * 0.60, safe.B + h * 0.28, safe.T), gap * 0.18),
                                     InsetRect(New SlotRect(safe.L + w * 0.63, safe.R, safe.B + h * 0.52, safe.T), gap * 0.18),
                                     InsetRect(New SlotRect(safe.L + w * 0.60, safe.R, safe.B, safe.B + h * 0.48), gap * 0.18)))
        result.Add(NewLayoutTemplate("SLOPED_FACADE_MAIN", PartFamily.Sloped, ViewRole.SlopeView, ViewRole.PlanContour, ViewRole.IsoReference,
                                     InsetRect(New SlotRect(safe.L, safe.L + w * 0.38, safe.B, safe.T), gap * 0.18),
                                     InsetRect(New SlotRect(safe.L + w * 0.42, safe.R, safe.B + h * 0.56, safe.T), gap * 0.18),
                                     InsetRect(New SlotRect(safe.L + w * 0.52, safe.R, safe.B, safe.B + h * 0.50), gap * 0.18)))

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
        If mainM Is Nothing OrElse auxM Is Nothing OrElse isoM Is Nothing Then Return Nothing

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
        p.Aux2Fit = ScaleToFit(p.Aux2Slot, isoM, ISO_SCALE_MARGIN)
        If p.MainFit Is Nothing OrElse p.Aux1Fit Is Nothing OrElse p.Aux2Fit Is Nothing Then Return Nothing

        Dim mainFill As Double = (p.MainFit.ProjectedW * p.MainFit.ProjectedH) / Math.Max(0.0001, RectW(p.MainSlot) * RectH(p.MainSlot))
        Dim auxFill As Double = (p.Aux1Fit.ProjectedW * p.Aux1Fit.ProjectedH) / Math.Max(0.0001, RectW(p.Aux1Slot) * RectH(p.Aux1Slot))
        Dim complement As Double = Math.Abs(mainM.AspectRatio - auxM.AspectRatio)
        p.Score = mainFill * 0.55 + auxFill * 0.25 + Math.Min(1.0, complement / 3.0) * 0.2

        If descriptor IsNot Nothing Then
            If descriptor.Family = PartFamily.Sloped AndAlso t.MainRole = ViewRole.SlopeView Then p.Score += 0.2
            If descriptor.Family = PartFamily.Linear AndAlso t.MainRole = ViewRole.LongitudinalFacade Then p.Score += 0.12
            If descriptor.Family = PartFamily.Radial AndAlso String.Equals(t.TemplateName, "RADIAL_PROFILE_HEAVY", StringComparison.OrdinalIgnoreCase) AndAlso descriptor.HasComplexProfile Then
                p.Score += 0.2
            End If
        End If
        Return p
    End Function

    Private Function PlaceViewsByTemplate(sheet As Sheet, modelDoc As Document, plan As LayoutPlan) As Dictionary(Of ViewRole, DrawingView)
        Dim placed As New Dictionary(Of ViewRole, DrawingView)()
        If plan Is Nothing Then Return placed

        Dim mainV As DrawingView = PlaceViewInSlot(sheet, modelDoc, plan.MainMeasure, plan.MainFit, plan.MainSlot)
        Dim auxV As DrawingView = PlaceViewInSlot(sheet, modelDoc, plan.Aux1Measure, plan.Aux1Fit, plan.Aux1Slot)
        Dim isoV As DrawingView = PlaceViewInSlot(sheet, modelDoc, plan.Aux2Measure, plan.Aux2Fit, plan.Aux2Slot)

        If mainV Is Nothing OrElse auxV Is Nothing OrElse isoV Is Nothing Then Return Nothing
        placed(plan.MainRole) = mainV
        placed(plan.AuxRole) = auxV
        placed(plan.IsoRole) = isoV
        Return placed
    End Function

    Private Sub AddViewRoleCaptions(doc As DrawingDocument,
                                    sheet As Sheet,
                                    plan As LayoutPlan,
                                    placed As Dictionary(Of ViewRole, DrawingView))
        If plan Is Nothing OrElse placed Is Nothing Then Return
        Try
            AddViewAnnotations(doc, sheet, placed(plan.MainRole), plan.MainSlot, plan.MainMeasure, False)
            AddViewAnnotations(doc, sheet, placed(plan.AuxRole), plan.Aux1Slot, plan.Aux1Measure, False)
            AddViewAnnotations(doc, sheet, placed(plan.IsoRole), plan.Aux2Slot, plan.Aux2Measure, False)
        Catch ex As Exception
            Debug.Print("WARN AddViewRoleCaptions: " & ex.Message)
        End Try
    End Sub

    Private Function BuildDimensionPlan(descriptor As PartDescriptor,
                                        roleMap As RoleMap) As DimensionPlan
        Dim plan As New DimensionPlan()
        If descriptor Is Nothing Then Return plan

        Select Case descriptor.Family
            Case PartFamily.Plate
                plan.Intents.Add(NewIntent(DimensionIntentId.OverallLength, ViewRole.PlanContour, 0, True))
                plan.Intents.Add(NewIntent(DimensionIntentId.OverallWidth, ViewRole.PlanContour, 0, True))
                plan.Intents.Add(NewIntent(DimensionIntentId.OverallThickness, ViewRole.ThicknessView, 0, True))
                If descriptor.HasComplexProfile Then
                    plan.Intents.Add(NewIntent(DimensionIntentId.ProfileDepth, ViewRole.CrossProfile, 1, True))
                End If
            Case PartFamily.Linear
                plan.Intents.Add(NewIntent(DimensionIntentId.OverallLength, ViewRole.LongitudinalFacade, 0, True))
                plan.Intents.Add(NewIntent(DimensionIntentId.OverallHeight, ViewRole.LongitudinalFacade, 0, True))
                plan.Intents.Add(NewIntent(DimensionIntentId.OverallThickness, ViewRole.CrossProfile, 0, True))
                plan.Intents.Add(NewIntent(DimensionIntentId.ProfileHeight, ViewRole.CrossProfile, 1, True))
                plan.Intents.Add(NewIntent(DimensionIntentId.ProfileDepth, ViewRole.CrossProfile, 1, True))
                If descriptor.HasComplexProfile Then
                    plan.Intents.Add(NewIntent(DimensionIntentId.StepHeight, ViewRole.CrossProfile, 2, True))
                End If
            Case PartFamily.Radial
                plan.Intents.Add(NewIntent(DimensionIntentId.ChordOrSpan, ViewRole.PlanContour, 0, True))
                plan.Intents.Add(NewIntent(DimensionIntentId.OverallWidth, ViewRole.PlanContour, 0, True))
                plan.Intents.Add(NewIntent(DimensionIntentId.RadiusMain, ViewRole.PlanContour, 1, True))
                plan.Intents.Add(NewIntent(DimensionIntentId.OverallThickness, ViewRole.CrossProfile, 0, True))
                plan.Intents.Add(NewIntent(DimensionIntentId.ProfileHeight, ViewRole.CrossProfile, 1, True))
            Case PartFamily.Sloped
                plan.Intents.Add(NewIntent(DimensionIntentId.OverallLength, ViewRole.SlopeView, 0, True))
                plan.Intents.Add(NewIntent(DimensionIntentId.SlopeHeightHigh, ViewRole.SlopeView, 0, True))
                plan.Intents.Add(NewIntent(DimensionIntentId.SlopeHeightLow, ViewRole.SlopeView, 0, True))
                plan.Intents.Add(NewIntent(DimensionIntentId.OverallThickness, ViewRole.ThicknessView, 1, True))
                plan.Intents.Add(NewIntent(DimensionIntentId.OverallWidth, ViewRole.PlanContour, 1, True))
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
                                   plan As DimensionPlan)
        If plan Is Nothing OrElse placedViews Is Nothing Then Return
        Dim placedIntents As New HashSet(Of DimensionIntentId)()

        plan.Intents.Sort(Function(a As DimensionIntent, b As DimensionIntent) a.Priority.CompareTo(b.Priority))
        For Each intent As DimensionIntent In plan.Intents
            If placedIntents.Contains(intent.IntentId) Then Continue For
            If intent.PreferredRole = ViewRole.IsoReference Then Continue For
            Dim v As DrawingView = Nothing
            If placedViews.ContainsKey(intent.PreferredRole) Then v = placedViews(intent.PreferredRole)
            If v Is Nothing Then Continue For

            Dim slot As SlotRect = ResolveSlotByRole(placedViews, intent.PreferredRole)
            Dim addH As Boolean = IsHorizontalIntent(intent.IntentId)
            Dim addV As Boolean = IsVerticalIntent(intent.IntentId)

            Dim added As Integer = TryAddTrueDimensions(doc, sheet, v, slot, addH, addV)
            If added > 0 Then
                placedIntents.Add(intent.IntentId)
            ElseIf intent.AllowFallbackNote Then
                Dim m As ViewMeasure = roleMap.GetMeasure(intent.PreferredRole)
                AddFallbackDimensionNotes(doc, sheet, v, slot, addH, addV, m)
                placedIntents.Add(intent.IntentId)
                Debug.Print("Fallback dimension note for " & intent.IntentId.ToString())
            End If
        Next
    End Sub

    Private Function ResolveSlotByRole(placedViews As Dictionary(Of ViewRole, DrawingView), role As ViewRole) As SlotRect
        Dim v As DrawingView = Nothing
        If placedViews.ContainsKey(role) Then v = placedViews(role)
        If v Is Nothing Then Return New SlotRect(0, 0, 0, 0)
        Return New SlotRect(v.Left, v.Left + v.Width, v.Top - v.Height, v.Top)
    End Function

    Private Function IsHorizontalIntent(id As DimensionIntentId) As Boolean
        Return (id = DimensionIntentId.OverallLength OrElse
                id = DimensionIntentId.OverallWidth OrElse
                id = DimensionIntentId.ChordOrSpan OrElse
                id = DimensionIntentId.ProfileDepth)
    End Function

    Private Function IsVerticalIntent(id As DimensionIntentId) As Boolean
        Return (id = DimensionIntentId.OverallHeight OrElse
                id = DimensionIntentId.OverallThickness OrElse
                id = DimensionIntentId.ProfileHeight OrElse
                id = DimensionIntentId.StepHeight OrElse
                id = DimensionIntentId.SlopeHeightHigh OrElse
                id = DimensionIntentId.SlopeHeightLow)
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
            sheet.DrawingNotes.Add(_app.TransientGeometry.CreatePoint2d(px, capY), caption)
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
        added += TryAddTrueDimensions(doc, sheet, v, slot, addHorizontal, addVertical)

        If detailedProfile Then
            If added < 2 Then AddFallbackDimensionNotes(doc, sheet, v, slot, True, True, measure)
        ElseIf facadeLike Then
            If lightDim Then
                If added < 1 Then AddFallbackDimensionNotes(doc, sheet, v, slot, False, True, measure)
            Else
                If added < 1 Then AddFallbackDimensionNotes(doc, sheet, v, slot, False, True, measure)
                If added < 2 Then AddFallbackDimensionNotes(doc, sheet, v, slot, True, False, measure)
            End If
        Else
            If added < 2 Then AddFallbackDimensionNotes(doc, sheet, v, slot, True, True, measure)
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
                    Dim p As Point2d = _app.TransientGeometry.CreatePoint2d((slot.L + slot.R) / 2.0, Math.Min(slot.T - Cm(doc, 2.0), v.Top + Cm(doc, 3.5)))
                    sheet.DrawingDimensions.GeneralDimensions.AddLinear(p, i1, i2, DimensionTypeEnum.kHorizontalDimensionType)
                    count += 1
                Catch
                End Try
            End If

            If addVertical AndAlso minYCurve IsNot Nothing AndAlso maxYCurve IsNot Nothing Then
                Try
                    Dim j1 As GeometryIntent = sheet.CreateGeometryIntent(minYCurve, PointIntentEnum.kMidPointIntent)
                    Dim j2 As GeometryIntent = sheet.CreateGeometryIntent(maxYCurve, PointIntentEnum.kMidPointIntent)
                    Dim p2 As Point2d = _app.TransientGeometry.CreatePoint2d(Math.Min(slot.R - Cm(doc, 1.0), v.Left + v.Width + Cm(doc, 2.6)), (slot.B + slot.T) / 2.0)
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
                                          addVertical As Boolean,
                                          measure As ViewMeasure)
        Try
            Dim sc As Double = v.Scale
            If sc <= 0.0001 Then Return

            Dim realWmm As Double = Math.Round(v.Width / sc * 10.0)
            Dim realHmm As Double = Math.Round(v.Height / sc * 10.0)
            Dim isProfile As Boolean = IsProfileLikeMeasure(measure)
            Dim isLong As Boolean = IsLongitudinalMeasure(measure)

            If addHorizontal AndAlso realWmm >= VIEW_DIM_MIN_MM Then
                Dim px As Double = Math.Max(slot.L + Cm(doc, 3.0), Math.Min(slot.R - Cm(doc, 3.0), v.Left + v.Width / 2.0))
                Dim py As Double = Math.Min(slot.T - Cm(doc, 2.2), v.Top + Cm(doc, 3.7))
                If isLong Then py = Math.Min(slot.T - Cm(doc, 2.0), v.Top + Cm(doc, 4.2))
                If isProfile Then py = Math.Min(slot.T - Cm(doc, 2.1), v.Top + Cm(doc, 3.4))
                py = Math.Max(slot.B + Cm(doc, 2.0), py)

                If RectW(slot) >= Cm(doc, 22.0) Then
                    sheet.DrawingNotes.Add(_app.TransientGeometry.CreatePoint2d(px, py),
                                           "↔ " & String.Format("{0:F0} мм", realWmm))
                End If
            End If

            If addVertical AndAlso realHmm >= VIEW_DIM_MIN_MM Then
                Dim rightSpace As Double = slot.R - (v.Left + v.Width)
                Dim leftSpace As Double = v.Left - slot.L
                Dim px2 As Double
                If rightSpace >= leftSpace Then
                    px2 = Math.Min(slot.R - Cm(doc, 1.2), v.Left + v.Width + Cm(doc, 2.2))
                Else
                    px2 = Math.Max(slot.L + Cm(doc, 1.2), v.Left - Cm(doc, 2.2))
                End If

                Dim py2 As Double = v.Top - v.Height / 2.0
                If isLong Then
                    py2 = Math.Max(slot.B + Cm(doc, 2.0), Math.Min(slot.T - Cm(doc, 2.0), v.Top - v.Height * 0.45))
                ElseIf isProfile Then
                    py2 = Math.Max(slot.B + Cm(doc, 2.0), Math.Min(slot.T - Cm(doc, 2.0), v.Top - v.Height * 0.5))
                End If

                If RectH(slot) >= Cm(doc, 18.0) Then
                    sheet.DrawingNotes.Add(_app.TransientGeometry.CreatePoint2d(px2, py2),
                                           "↕ " & String.Format("{0:F0} мм", realHmm))
                End If
            End If
        Catch ex As Exception
            Debug.Print("WARN AddFallbackDimensionNotes: " & ex.Message)
        End Try
    End Sub

    Private Function MeasureView(sheet As Sheet, modelDoc As Document,
                                 orient As ViewOrientationTypeEnum,
                                 style As DrawingViewStyleEnum,
                                 key As String,
                                 caption As String) As ViewMeasure
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
            m.Key = key
            m.Caption = caption

            Dim totalLen As Double = 0
            Dim totalHLen As Double = 0
            Dim totalVLen As Double = 0

            Try
                Dim curves As DrawingCurvesEnumerator = probe.DrawingCurves
                m.CurveCount = curves.Count
                For i As Integer = 1 To curves.Count
                    Dim c As DrawingCurve = curves.Item(i)
                    If c Is Nothing Then Continue For
                    Dim rb As Box2d = c.RangeBox
                    Dim dx As Double = Math.Abs(rb.MaxPoint.X - rb.MinPoint.X)
                    Dim dy As Double = Math.Abs(rb.MaxPoint.Y - rb.MinPoint.Y)
                    Dim segLen As Double = Math.Sqrt(dx * dx + dy * dy)
                    totalLen += segLen
                    If dx >= dy Then totalHLen += segLen
                    If dy >= dx Then totalVLen += segLen

                    Select Case c.CurveType
                        Case Curve2dTypeEnum.kLineSegmentCurve2d
                            m.LineCount += 1
                            If dx > 0.0001 AndAlso dy > 0.0001 Then m.NonAxisEdgeCount += 1
                        Case Curve2dTypeEnum.kCircularArcCurve2d
                            m.ArcCount += 1
                        Case Curve2dTypeEnum.kCircleCurve2d
                            m.CircleCount += 1
                    End Select

                    If dx > 0.0001 AndAlso dy > 0.0001 Then m.HasSlopeEdges = True
                    If rb.MinPoint.X > (probe.Left + probe.Width * 0.1) AndAlso rb.MaxPoint.X < (probe.Left + probe.Width * 0.9) AndAlso _
                       rb.MinPoint.Y > (probe.Top - probe.Height * 0.9) AndAlso rb.MaxPoint.Y < (probe.Top - probe.Height * 0.1) Then
                        m.InnerContourCount += 1
                    End If
                Next
            Catch
                m.CurveCount = 0
            End Try

            m.HasArcs = (m.ArcCount + m.CircleCount) > 0
            m.BoundingArea = Math.Max(0.0001, m.UnitW * m.UnitH)

            Try
                Dim mx As Double = Math.Max(m.UnitW, m.UnitH)
                Dim mn As Double = Math.Max(0.0001, Math.Min(m.UnitW, m.UnitH))
                m.AspectRatio = mx / mn
            Catch
                m.AspectRatio = 1.0
            End Try

            m.HorizontalBias = totalHLen / Math.Max(0.0001, totalLen)
            m.VerticalBias = totalVLen / Math.Max(0.0001, totalLen)
            m.LongEdgeBias = Math.Max(m.HorizontalBias, m.VerticalBias)
            m.SlopeScore = CDbl(m.NonAxisEdgeCount) / Math.Max(1.0, CDbl(m.LineCount))
            m.ProfileComplexityScore = CDbl(m.InnerContourCount) * 0.35 + CDbl(m.CurveCount) / 18.0 + CDbl(m.NonAxisEdgeCount) * 0.15
            m.PlanComplexityScore = CDbl(m.ArcCount + m.CircleCount) * 0.25 + CDbl(m.CurveCount) / 20.0

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
    Public Enum LayoutArchetype
        RadialSegment
        LongLinear
        PlateBlock
    End Enum

    Public Enum PartFamily
        Plate
        Linear
        Radial
        Sloped
    End Enum

    Public Enum ViewRole
        MainContour
        PlanContour
        LongitudinalFacade
        CrossProfile
        SlopeView
        EndFace
        ThicknessView
        IsoReference
    End Enum

    Public Enum DimensionIntentId
        OverallLength
        OverallWidth
        OverallHeight
        OverallThickness
        ProfileDepth
        ProfileHeight
        StepHeight
        SlopeHeightHigh
        SlopeHeightLow
        RadiusMain
        RadiusSecondary
        ChordOrSpan
    End Enum

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
        Public Key As String
        Public Caption As String
        Public CurveCount As Integer
        Public AspectRatio As Double

        Public LineCount As Integer
        Public ArcCount As Integer
        Public CircleCount As Integer
        Public HasArcs As Boolean
        Public HasSlopeEdges As Boolean
        Public NonAxisEdgeCount As Integer
        Public InnerContourCount As Integer
        Public BoundingArea As Double
        Public LongEdgeBias As Double
        Public VerticalBias As Double
        Public HorizontalBias As Double
        Public ProfileComplexityScore As Double
        Public PlanComplexityScore As Double
        Public SlopeScore As Double
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
        Public PatternName As String
        Public Archetype As LayoutArchetype
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

        Public MainRole As ViewRole
        Public AuxRole As ViewRole
        Public IsoRole As ViewRole
        Public TemplateName As String
        Public Score As Double
    End Class

    Public Class PartDescriptor
        Public Family As PartFamily
        Public IsLong As Boolean
        Public IsThin As Boolean
        Public HasDominantPlan As Boolean
        Public HasDominantFacade As Boolean
        Public HasComplexProfile As Boolean
        Public HasRadialPlan As Boolean
        Public HasSlope As Boolean
        Public HasPlanTaper As Boolean
        Public HasDovetailEnds As Boolean
        Public HasDecorativeRecess As Boolean
        Public HasEdgeRadiusOrDrip As Boolean
    End Class

    Public Class RoleMap
        Public ByRole As New Dictionary(Of ViewRole, ViewMeasure)()
        Public Function GetMeasure(role As ViewRole) As ViewMeasure
            If ByRole.ContainsKey(role) Then Return ByRole(role)
            Return Nothing
        End Function
    End Class

    Public Class LayoutTemplate
        Public TemplateName As String
        Public Family As PartFamily
        Public RequiredRoles As New List(Of ViewRole)()
        Public MainRole As ViewRole
        Public AuxRole As ViewRole
        Public IsoRole As ViewRole
        Public MainSlot As SlotRect
        Public AuxSlot As SlotRect
        Public IsoSlot As SlotRect
    End Class

    Public Class DimensionIntent
        Public IntentId As DimensionIntentId
        Public PreferredRole As ViewRole
        Public Priority As Integer
        Public AllowFallbackNote As Boolean
    End Class

    Public Class DimensionPlan
        Public Intents As New List(Of DimensionIntent)()
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
