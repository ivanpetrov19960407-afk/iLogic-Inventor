Attribute VB_Name = "RKM_Excel"
Option Explicit

' NOTE: Historical legacy VBA reference excerpt; this file is not a guaranteed full importable module.

Private Const DEFAULT_SHEET_NAME As String = "ALBUM"
Private Const HEADER_ROW_INDEX As Long = 1
Private Const HEADER_SCAN_ROWS As Long = 10

' Загрузка списка изделий из Excel 
Public Function LoadAlbumItemsFromExcel(ByVal excelPath As String, Optional ByVal workspacePath As String = "", Optional ByVal sheetName As String = DEFAULT_SHEET_NAME) As Collection
    Dim xlApp As Object, xlBook As Object, xlSheet As Object
    Dim headerMap As Object, i As Long, rowsCount As Long
    Dim modelPathRaw As String, resolvedModelPath As String
    Dim item As Object, promptMap As Object

    On Error GoTo EH
    If Len(Trim$(excelPath)) = 0 Then Exit Function

    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    xlApp.DisplayAlerts = False

    Set xlBook = xlApp.Workbooks.Open(excelPath)
    Set xlSheet = ResolveAlbumWorksheet(xlBook, sheetName)
    
    If xlSheet Is Nothing Then
        Err.Raise vbObjectError + 3100, "LoadAlbumItemsFromExcel", "Worksheet '" & sheetName & "' was not found."
    End If

    Set headerMap = ReadHeaderMap(xlSheet, DetectHeaderRowIndex(xlSheet))
    If Not headerMap.Exists("MODEL_PATH") Then
        Err.Raise vbObjectError + 3101, "LoadAlbumItemsFromExcel", "Required header MODEL_PATH is missing."
    End If

    rowsCount = xlSheet.Cells(xlSheet.Rows.Count, CLng(headerMap("MODEL_PATH"))).End(-4162).Row
    Set LoadAlbumItemsFromExcel = New Collection

    For i = CLng(headerMap("__HEADER_ROW")) + 1 To rowsCount
        modelPathRaw = Trim$(CStr(xlSheet.Cells(i, CLng(headerMap("MODEL_PATH"))).Value))
        If Len(modelPathRaw) > 0 Then
            resolvedModelPath = ResolveModelPath(modelPathRaw, workspacePath, excelPath)
            
            If Len(resolvedModelPath) > 0 Then
                Set item = CreateObject("Scripting.Dictionary")
                item.CompareMode = 1
                item("MODEL_PATH") = resolvedModelPath
                Set item("PROMPTS") = CreatePromptMapFromRow(xlSheet, headerMap, i)
                LoadAlbumItemsFromExcel.Add item
            End If
        End If
    Next i

CleanExit:
    On Error Resume Next
    xlBook.Close False
    xlApp.Quit
    Exit Function
EH:
    Debug.Print "LOG: Excel load failed. Err=" & Err.Number
    Resume CleanExit
End Function

' Умный поиск заголовков (алиасы) 
Private Function ResolveHeaderAlias(ByVal rawHeader As String) As String
    Dim n As String
    n = UCase$(Trim$(rawHeader))
    
    Select Case n
        Case "MODEL_PATH", "MODEL", "P", "ПУТЬ", "ФАЙЛ"
            ResolveHeaderAlias = "MODEL_PATH"
        Case "CODE", "ШИФР", "АРТИКУЛ", "ОБОЗНАЧЕНИЕ"
            ResolveHeaderAlias = "CODE"
        Case "PROJECT_NAME", "PROJECT", "ОБЪЕКТ", "ПРОЕКТ"
            ResolveHeaderAlias = "PROJECT_NAME"
        Case "DRAWING_NAME", "TITLE", "НАИМЕНОВАНИЕ"
            ResolveHeaderAlias = "DRAWING_NAME"
        Case "ORG_NAME", "ОРГАНИЗАЦИЯ", "КОМПАНИЯ"
            ResolveHeaderAlias = "ORG_NAME"
        Case "SHEET", "ЛИСТ"
            ResolveHeaderAlias = "SHEET"
        Case "SHEETS", "ЛИСТОВ"
            ResolveHeaderAlias = "SHEETS"
    End Select
End Function

' Поиск пути к модели (с учетом папки проекта или Excel) 
Private Function ResolveModelPath(ByVal inputPath As String, ByVal workspacePath As String, ByVal excelPath As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(inputPath) Then
        ResolveModelPath = fso.GetAbsolutePathName(inputPath)
    ElseIf Len(workspacePath) > 0 And fso.FileExists(fso.BuildPath(workspacePath, inputPath)) Then
        ResolveModelPath = fso.GetAbsolutePathName(fso.BuildPath(workspacePath, inputPath))
    Else
        Dim excelFolder As String
        excelFolder = fso.GetParentFolderName(excelPath)
        If fso.FileExists(fso.BuildPath(excelFolder, inputPath)) Then
            ResolveModelPath = fso.GetAbsolutePathName(fso.BuildPath(excelFolder, inputPath))
        End If
    End If
End Function

' Остальные вспомогательные функции (ReadHeaderMap, CreatePromptMapFromRow и др.) 
