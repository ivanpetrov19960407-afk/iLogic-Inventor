Attribute VB_Name = "RKM_Utils"
Option Explicit

' NOTE: Historical legacy VBA reference excerpt; this file is not a guaranteed full importable module.

Public Const RKM_BORDER_NAME As String = "RKM_SPDS_A3_BORDER"
Public Const MM_TO_CM As Double = 0.1
Public Const A3_WIDTH_MM As Double = 420#
Public Const A3_HEIGHT_MM As Double = 297#
Public Const FRAME_LEFT_MM As Double = 20#
Public Const FRAME_OTHER_MM As Double = 5#
Public Const TITLE_W_MM As Double = 185#
Public Const TITLE_H_MM As Double = 55#
Public Const DIM_TOLERANCE_MM As Double = 0.05

Public Function MmToCm(ByVal oDoc As DrawingDocument, ByVal valueMm As Double) As Double
    Dim oUom As UnitsOfMeasure
    If oDoc Is Nothing Then
        MmToCm = valueMm * MM_TO_CM
        Exit Function
    End If
    Set oUom = oDoc.UnitsOfMeasure
    MmToCm = oUom.ConvertUnits(valueMm, kMillimeterLengthUnits, kCentimeterLengthUnits)
End Function

Public Function CmToMm(ByVal oDoc As DrawingDocument, ByVal valueCm As Double) As Double
    Dim oUom As UnitsOfMeasure
    If oDoc Is Nothing Then
        CmToMm = valueCm / MM_TO_CM
        Exit Function
    End If
    Set oUom = oDoc.UnitsOfMeasure
    CmToMm = oUom.ConvertUnits(valueCm, kCentimeterLengthUnits, kMillimeterLengthUnits)
End Function

Public Function Pt(ByVal x As Double, ByVal y As Double) As Point2d
    Set Pt = ThisApplication.TransientGeometry.CreatePoint2d(x, y)
End Function

Public Function CanEditDrawingResources(ByVal oApp As Inventor.Application) As Boolean
    Dim eo As Object
    CanEditDrawingResources = False
    If oApp Is Nothing Then Exit Function
    On Error Resume Next
    Set eo = oApp.ActiveEditObject
    On Error GoTo 0
    If Not eo Is Nothing Then
        If TypeOf eo Is DrawingSketch Or TypeOf eo Is Sketch Then
            Exit Function
        End If
    End If
    CanEditDrawingResources = True
End Function

Public Function GetActiveDrawingDocument(ByVal oApp As Inventor.Application) As DrawingDocument
    If oApp Is Nothing Then Exit Function
    If oApp.ActiveDocument Is Nothing Then Exit Function
    If oApp.ActiveDocument.DocumentType <> kDrawingDocumentObject Then Exit Function
    Set GetActiveDrawingDocument = oApp.ActiveDocument
End Function
