Attribute VB_Name = "RKM_TitleBlockPrompted"
Option Explicit

' Имя определения штампа в ресурсах
Public Const RKM_TITLEBLOCK_NAME As String = "RKM_SPDS_A3_FORM3"

' Геометрия штампа (Форма 3 по ГОСТ)
Private Const TITLE_W_MM As Double = 185#
Private Const TITLE_H_MM As Double = 55#

' Основная функция: Применяет штамп и заполняет его данными
Public Sub ApplyRkmTitleBlockToSheetWithPrompts(ByVal oSheet As Sheet, ByVal oDef As TitleBlockDefinition, Optional ByVal promptData As Variant)
    Dim promptStrings() As String
    Dim newTitleBlock As TitleBlock

    On Error GoTo AddTitleBlockFailed
    If oSheet Is Nothing Or oDef Is Nothing Then Exit Sub

    ' Подготовка массива строк для полей штампа (Prompted Entries)
    ' Порядок: Код, Проект, Название чертежа, Организация, Стадия, Лист, Листов
    promptStrings = BuildPromptStringsFromAny(promptData)

    ' Тихая замена старого штампа на новый с заполненными полями
    ThisApplication.SilentOperation = True
    On Error Resume Next
    If Not oSheet.TitleBlock Is Nothing Then oSheet.TitleBlock.Delete
    On Error GoTo AddTitleBlockFailed
    
    Set newTitleBlock = oSheet.AddTitleBlock(oDef, , promptStrings)
    ThisApplication.SilentOperation = False
    Exit Sub

AddTitleBlockFailed:
    ThisApplication.SilentOperation = False
    Debug.Print "LOG: Ошибка заполнения штампа на листе " & oSheet.Name
End Sub

' Названия ключей, которые мы ищем в данных из Excel
Public Function GetPromptOrder() As Variant
    GetPromptOrder = Array("CODE", "PROJECT_NAME", "DRAWING_NAME", "ORG_NAME", "STAGE", "SHEET", "SHEETS")
End Function

' Заполнение массива строк на основе Dictionary (из модуля RKM_Excel)
Private Sub FillPromptStringsFromMap(ByRef prompts() As String, ByVal promptMap As Object, ByVal promptOrder As Variant)
    Dim i As Long
    Dim keyName As String

    For i = LBound(promptOrder) To UBound(promptOrder)
        keyName = CStr(promptOrder(i))
        ' Если данных нет, оставляем поле пустым или берем дефолт
        If Not promptMap Is Nothing Then
            If promptMap.Exists(keyName) Then
                prompts(i + 1) = CStr(promptMap(keyName))
            End If
        End If
    Next i
End Sub
