' ================================================================
' RKM_SPDS_A3_Frame.ilogic.vb  –  v2.0
' Применить рамку СПДС А3 + штамп Форма 3 на активный лист
' Запускается одной кнопкой — без Excel, без моделей.
'
' Источник: vba-inventor / RKM_FrameBorder.bas + RKM_TitleBlockPrompted.bas
' ================================================================

Option Explicit On

Imports Inventor
Imports System

Sub Main()
    Dim doc As DrawingDocument = TryCast(ThisApplication.ActiveDocument, DrawingDocument)
    If doc Is Nothing Then
        System.Windows.Forms.MessageBox.Show("Откройте .idw документ.", "Ошибка")
        Return
    End If

    Dim sheet As Sheet = doc.ActiveSheet
    If sheet Is Nothing Then
        System.Windows.Forms.MessageBox.Show("Нет активного листа.", "Ошибка")
        Return
    End If

    ' Установить А3 альбомная
    Try
        sheet.Size        = DrawingSheetSizeEnum.kA3DrawingSheetSize
        sheet.Orientation = PageOrientationTypeEnum.kLandscapePageOrientation
    Catch : End Try

    ' Проверить размер
    Dim wMm As Double = doc.UnitsOfMeasure.ConvertUnits(sheet.Width,  UnitsTypeEnum.kCentimeterLengthUnits, UnitsTypeEnum.kMillimeterLengthUnits)
    Dim hMm As Double = doc.UnitsOfMeasure.ConvertUnits(sheet.Height, UnitsTypeEnum.kCentimeterLengthUnits, UnitsTypeEnum.kMillimeterLengthUnits)

    If Math.Abs(wMm - 420.0) > 0.5 OrElse Math.Abs(hMm - 297.0) > 0.5 Then
        System.Windows.Forms.MessageBox.Show("Лист не А3: " & FormatNumber(wMm,1) & " × " & FormatNumber(hMm,1) & " мм", "Ошибка")
        Return
    End If

    Dim framer As New SpdsFramer(ThisApplication)

    Dim borderDef    As BorderDefinition    = framer.EnsureBorderDef(doc)
    Dim titleDef     As TitleBlockDefinition = framer.EnsureTitleBlockDef(doc)

    framer.ApplyBorder(sheet, borderDef)
    framer.ApplyTitleBlock(sheet, titleDef, GetDefaultPrompts())

    System.Windows.Forms.MessageBox.Show("Рамка СПДС А3 и штамп Форма 3 применены.", "Готово")
End Sub

Private Function GetDefaultPrompts() As Dictionary(Of String, String)
    Return New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
        {"CODE",          "000-2026-АР"},
        {"PROJECT_NAME",  "Наименование объекта"},
        {"DRAWING_NAME",  "Эскиз изделия"},
        {"ORG_NAME",      "ООО «Организация»"},
        {"STAGE",         "ИД"},
        {"SHEET",         "1"},
        {"SHEETS",        "1"}
    }
End Function

' ================================================================
'  КЛАСС ОБЁРТКА (можно переиспользовать из StoneAlbumRule)
' ================================================================
Public Class SpdsFramer

    Private ReadOnly _app As Inventor.Application

    Private Const A3_W_MM     As Double = 420.0
    Private Const A3_H_MM     As Double = 297.0
    Private Const FRAME_L_MM  As Double = 20.0
    Private Const FRAME_O_MM  As Double = 5.0
    Private Const TB_W_MM     As Double = 185.0
    Private Const TB_H_MM     As Double = 55.0
    Private Const BORDER_NAME As String = "RKM_SPDS_A3_BORDER_V12"
    Private Const TB_NAME     As String = "RKM_SPDS_A3_FORM3_V17"

    Public Sub New(app As Inventor.Application)
        _app = app
    End Sub

    ' --- Рамка ---
    Public Function EnsureBorderDef(doc As DrawingDocument) As BorderDefinition
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
            ' Очистить старую геометрию
            For i As Integer = sk.SketchLines.Count To 1 Step -1
                Try
                    sk.SketchLines.Item(i).Delete()
                Catch
                End Try
            Next

            ' Микро-якоря для фиксации листа
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

    Public Sub ApplyBorder(sheet As Sheet, def As BorderDefinition)
        Try
            If sheet.Border IsNot Nothing Then sheet.Border.Delete()
        Catch
        End Try
        _app.SilentOperation = True
        sheet.AddCustomBorder(def)
        _app.SilentOperation = False
    End Sub

    ' --- Штамп ---
    Public Function EnsureTitleBlockDef(doc As DrawingDocument) As TitleBlockDefinition
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
            ' Полная очистка скетча (линии + текст)
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

            DrawGeometry(doc, sk)
            DrawLabels(doc, sk)
        Finally
            def.ExitEdit(True)
        End Try

        Return def
    End Function

    Public Sub ApplyTitleBlock(sheet As Sheet, def As TitleBlockDefinition, prompts As Dictionary(Of String, String))
        Try
            If sheet.TitleBlock IsNot Nothing Then sheet.TitleBlock.Delete()
        Catch
        End Try

        Dim order As String() = {"CODE","PROJECT_NAME","DRAWING_NAME","ORG_NAME","STAGE","SHEET","SHEETS"}
        Dim ps(8) As String
        For i As Integer = 0 To order.Length - 1
            Dim v As String = String.Empty
            If prompts IsNot Nothing Then prompts.TryGetValue(order(i), v)
            ps(i + 1) = If(String.IsNullOrEmpty(v), String.Empty, v)
        Next

        Try
            _app.SilentOperation = True
            sheet.AddTitleBlock(def, Nothing, ps)
        Catch ex As Exception
            Debug.Print("WARN:  AddTitleBlock: " & ex.Message)
        Finally
            _app.SilentOperation = False
        End Try
    End Sub

    ' --- Геометрия штампа ---
    Private Sub DrawGeometry(doc As DrawingDocument, sk As DrawingSketch)
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
            HL(doc, sk, x1, y1,   0,  67, y)
        Next
        HL(doc, sk, x1, y1,   0, 185, 15) : HL(doc, sk, x1, y1,   0,  67, 35)
        HL(doc, sk, x1, y1, 137, 185, 35) : HL(doc, sk, x1, y1,   0, 185, 40)
        HL(doc, sk, x1, y1,   0,  67, 45) : HL(doc, sk, x1, y1,   0,  67, 50)
    End Sub

    Private Sub DrawLabels(doc As DrawingDocument, sk As DrawingSketch)
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

    Private Sub VL(doc As DrawingDocument, sk As DrawingSketch, x0 As Double, y0 As Double, atMm As Double, yFr As Double, yTo As Double)
        sk.SketchLines.AddByTwoPoints(P(x0+Cm(doc,atMm), y0+Cm(doc,yFr)), P(x0+Cm(doc,atMm), y0+Cm(doc,yTo)))
    End Sub
    Private Sub HL(doc As DrawingDocument, sk As DrawingSketch, x0 As Double, y0 As Double, xFr As Double, xTo As Double, atMm As Double)
        sk.SketchLines.AddByTwoPoints(P(x0+Cm(doc,xFr), y0+Cm(doc,atMm)), P(x0+Cm(doc,xTo), y0+Cm(doc,atMm)))
    End Sub
    Private Sub Lbl(doc As DrawingDocument, sk As DrawingSketch, x0 As Double, y0 As Double, l As Double, b As Double, r As Double, t As Double, text As String)
        Dim tb As Inventor.TextBox = sk.TextBoxes.AddByRectangle(P(x0+Cm(doc,l), y0+Cm(doc,b)), P(x0+Cm(doc,r), y0+Cm(doc,t)), text)
        tb.HorizontalJustification = HorizontalTextAlignmentEnum.kAlignTextCenter
        tb.VerticalJustification   = VerticalTextAlignmentEnum.kAlignTextMiddle
    End Sub
    Private Sub Prm(doc As DrawingDocument, sk As DrawingSketch, x0 As Double, y0 As Double, l As Double, b As Double, r As Double, t As Double, name As String)
        Dim tb As Inventor.TextBox = sk.TextBoxes.AddByRectangle(P(x0+Cm(doc,l), y0+Cm(doc,b)), P(x0+Cm(doc,r), y0+Cm(doc,t)), "<Prompt>" & name & "</Prompt>")
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

End Class
