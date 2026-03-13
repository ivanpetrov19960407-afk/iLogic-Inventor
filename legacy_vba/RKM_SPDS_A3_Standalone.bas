Attribute VB_Name = "RKM_SPDS_A3_Standalone"
Option Explicit

' Геометрические константы рамки (в мм)
Private Const A3_W_MM As Double = 420#
Private Const A3_H_MM As Double = 297#
Private Const FRAME_LEFT_MM As Double = 20#    ' Отступ слева под подшивку
Private Const FRAME_OTHER_MM As Double = 5#     ' Остальные отступы

' Константы штампа Форма 3 (основная надпись)
Private Const TB_W_MM As Double = 185#          ' Ширина штампа
Private Const TB_H_MM As Double = 55#           ' Высота штампа
Private Const A3_TOL_MM As Double = 0.05        ' Допуск точности

' Основная процедура создания рамки
Public Sub Rkm_CreateOrApplyA3Frame_SPDS()
    Dim oDoc As DrawingDocument
    Dim oSheet As Sheet
    Dim oBorderDef As BorderDefinition

    On Error GoTo EH
    Set oDoc = GetActiveDrawingDocument(ThisApplication)
    If oDoc Is Nothing Then Exit Sub

    ' Проверка возможности редактирования (модуль RKM_Utils)
    If Not CanEditDrawingResources(ThisApplication) Then Exit Sub

    ' Установка формата А3 Landscape
    Set oSheet = EnsureA3LandscapeSheet(oDoc)
    If oSheet Is Nothing Then Exit Sub

    ' ГЕОМЕТРИЯ: Отрисовка внутренней рамки
    Set oBorderDef = EnsureSpdsA3BorderDefinition(oDoc)
    ApplySpdsBorderToSheet oSheet, oBorderDef

    Debug.Print "LOG: SPDS A3 frame applied to sheet: " & oSheet.Name
    Exit Sub
EH:
    MsgBox "Ошибка при создании рамки: " & Err.Description, vbCritical
End Sub

' Создание определения рамки в ресурсах чертежа
Private Function EnsureSpdsA3BorderDefinition(ByVal oDoc As DrawingDocument) As BorderDefinition
    Dim oDef As BorderDefinition
    Dim oSketch As DrawingSketch

    ' Поиск существующей или создание новой
    Set oDef = FindBorderDefinition(oDoc, "RKM_SPDS_A3_BORDER")
    If oDef Is Nothing Then
        Set oDef = oDoc.BorderDefinitions.Add("RKM_SPDS_A3_BORDER")
    End If

    oDef.Edit oSketch
    ClearSketch oSketch ' Очистка старой геометрии
    
    ' Отрисовка двух прямоугольников: внешний (обрезной) и внутренний (рамка)
    DrawSpdsBorderGeometry oSketch
    
    oDef.ExitEdit True
    Set EnsureSpdsA3BorderDefinition = oDef
End Function

' Отрисовка линий рамки
Private Sub DrawSpdsBorderGeometry(ByVal oSketch As DrawingSketch)
    Dim xMax As Double, yMax As Double
    Dim ix1 As Double, iy1 As Double, ix2 As Double, iy2 As Double

    xMax = A3_W_MM * 0.1 ' Перевод в см для Inventor
    yMax = A3_H_MM * 0.1

    ix1 = FRAME_LEFT_MM * 0.1
    iy1 = FRAME_OTHER_MM * 0.1
    ix2 = (A3_W_MM - FRAME_OTHER_MM) * 0.1
    iy2 = (A3_H_MM - FRAME_OTHER_MM) * 0.1

    ' Внешний контур листа
    oSketch.SketchLines.AddAsTwoPointRectangle P2d(0, 0), P2d(xMax, yMax)
    ' Внутренняя рамка чертежа
    oSketch.SketchLines.AddAsTwoPointRectangle P2d(ix1, iy1), P2d(ix2, iy2)
End Sub
