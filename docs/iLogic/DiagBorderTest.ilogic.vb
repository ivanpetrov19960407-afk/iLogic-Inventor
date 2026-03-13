' DiagBorderTest.ilogic.vb
' Диагностика: создаём 1 лист, пробуем добавить рамку и штамп.
' Все ошибки выводятся в MessageBox.

Option Explicit On
Imports Inventor
Imports System

Sub Main()
    Dim doc As DrawingDocument = TryCast(ThisApplication.ActiveDocument, DrawingDocument)
    If doc Is Nothing Then
        System.Windows.Forms.MessageBox.Show("Откройте .idw документ.", "Ошибка")
        Return
    End If

    Dim log As New System.Text.StringBuilder()
    log.AppendLine("=== ДИАГНОСТИКА РАМКИ И ШТАМПА ===")
    log.AppendLine("Документ: " & doc.FullFileName)
    log.AppendLine("")

    ' --- Шаг 1: BorderDefinitions ---
    Dim borderCount As Integer = 0
    Try
        borderCount = doc.BorderDefinitions.Count
        log.AppendLine("BorderDefinitions.Count = " & borderCount)
        For i As Integer = 1 To borderCount
            log.AppendLine("  Border[" & i & "] = '" & doc.BorderDefinitions.Item(i).Name & "'")
        Next
    Catch ex As Exception
        log.AppendLine("ОШИБКА BorderDefinitions.Count: " & ex.Message)
    End Try

    ' --- Шаг 2: TitleBlockDefinitions ---
    Dim tbCount As Integer = 0
    Try
        tbCount = doc.TitleBlockDefinitions.Count
        log.AppendLine("TitleBlockDefinitions.Count = " & tbCount)
        For i As Integer = 1 To tbCount
            log.AppendLine("  TB[" & i & "] = '" & doc.TitleBlockDefinitions.Item(i).Name & "'")
        Next
    Catch ex As Exception
        log.AppendLine("ОШИБКА TitleBlockDefinitions.Count: " & ex.Message)
    End Try

    ' --- Шаг 3: Создаём определение рамки ---
    Dim borderDef As BorderDefinition = Nothing
    Dim BORDER_NAME As String = "RKM_SPDS_A3_BORDER_V12"
    log.AppendLine("")
    log.AppendLine("--- Создаём рамку '" & BORDER_NAME & "' ---")
    Try
        borderDef = doc.BorderDefinitions.Item(BORDER_NAME)
        log.AppendLine("Рамка уже существует, используем.")
    Catch
        Try
            ThisApplication.SilentOperation = True
            borderDef = doc.BorderDefinitions.Add(BORDER_NAME)
            ThisApplication.SilentOperation = False
            log.AppendLine("Рамка создана через .Add()")
        Catch ex As Exception
            ThisApplication.SilentOperation = False
            log.AppendLine("ОШИБКА Add рамки: " & ex.Message)
        End Try
    End Try

    If borderDef IsNot Nothing Then
        Try
            Dim sk As DrawingSketch = Nothing
            borderDef.Edit(sk)
            Try
                sk.SketchLines.AddByTwoPoints(
                    ThisApplication.TransientGeometry.CreatePoint2d(0,0),
                    ThisApplication.TransientGeometry.CreatePoint2d(0.0001,0.0001))
                sk.SketchLines.AddByTwoPoints(
                    ThisApplication.TransientGeometry.CreatePoint2d(42,29.7),
                    ThisApplication.TransientGeometry.CreatePoint2d(41.9999,29.6999))
                sk.SketchLines.AddAsTwoPointRectangle(
                    ThisApplication.TransientGeometry.CreatePoint2d(2.0, 0.5),
                    ThisApplication.TransientGeometry.CreatePoint2d(41.5, 29.2))
                log.AppendLine("Скетч рамки нарисован OK")
            Finally
                borderDef.ExitEdit(True)
            End Try
        Catch ex As Exception
            log.AppendLine("ОШИБКА Edit/ExitEdit рамки: " & ex.Message)
        End Try
    End If

    ' --- Шаг 4: Создаём определение штампа ---
    Dim tbDef As TitleBlockDefinition = Nothing
    Dim TB_NAME As String = "RKM_SPDS_A3_FORM3_V17"
    log.AppendLine("")
    log.AppendLine("--- Создаём штамп '" & TB_NAME & "' ---")
    Try
        tbDef = doc.TitleBlockDefinitions.Item(TB_NAME)
        log.AppendLine("Штамп уже существует, используем.")
    Catch
        Try
            ThisApplication.SilentOperation = True
            tbDef = doc.TitleBlockDefinitions.Add(TB_NAME)
            ThisApplication.SilentOperation = False
            log.AppendLine("Штамп создан через .Add()")
        Catch ex As Exception
            ThisApplication.SilentOperation = False
            log.AppendLine("ОШИБКА Add штампа: " & ex.Message)
        End Try
    End Try

    If tbDef IsNot Nothing Then
        Try
            Dim sk2 As DrawingSketch = Nothing
            tbDef.Edit(sk2)
            Try
                sk2.SketchLines.AddByTwoPoints(
                    ThisApplication.TransientGeometry.CreatePoint2d(0,0),
                    ThisApplication.TransientGeometry.CreatePoint2d(-0.0001,0.0001))
                sk2.SketchLines.AddAsTwoPointRectangle(
                    ThisApplication.TransientGeometry.CreatePoint2d(-19.0, 0.5),
                    ThisApplication.TransientGeometry.CreatePoint2d(-0.5, 6.0))
                Dim tb As Inventor.TextBox = sk2.TextBoxes.AddByRectangle(
                    ThisApplication.TransientGeometry.CreatePoint2d(-19.0, 4.5),
                    ThisApplication.TransientGeometry.CreatePoint2d(-0.5, 6.0),
                    "<Prompt>CODE</Prompt>")
                log.AppendLine("Скетч штампа нарисован OK")
            Finally
                tbDef.ExitEdit(True)
            End Try
        Catch ex As Exception
            log.AppendLine("ОШИБКА Edit/ExitEdit штампа: " & ex.Message)
        End Try
    End If

    ' --- Шаг 5: Создаём тестовый лист ---
    Dim sheet As Sheet = Nothing
    log.AppendLine("")
    log.AppendLine("--- Создаём тестовый лист ---")
    Try
        sheet = doc.Sheets.Add(
            DrawingSheetSizeEnum.kA3DrawingSheetSize,
            PageOrientationTypeEnum.kLandscapePageOrientation)
        sheet.Name = "DIAG_TEST"
        sheet.Activate()
        log.AppendLine("Лист создан: " & sheet.Name)
        log.AppendLine("Размер: " & sheet.Width & " x " & sheet.Height)
    Catch ex As Exception
        log.AppendLine("ОШИБКА создания листа: " & ex.Message)
    End Try

    ' --- Шаг 6: AddCustomBorder ---
    If sheet IsNot Nothing AndAlso borderDef IsNot Nothing Then
        log.AppendLine("")
        log.AppendLine("--- AddCustomBorder ---")
        ' Попытка 1: без SilentOperation
        Try
            sheet.AddCustomBorder(borderDef)
            log.AppendLine("AddCustomBorder БЕЗ SilentOperation — OK")
        Catch ex As Exception
            log.AppendLine("AddCustomBorder БЕЗ SilentOperation — ОШИБКА: " & ex.Message)
            ' Попытка 2: с SilentOperation
            Try
                ThisApplication.SilentOperation = True
                sheet.AddCustomBorder(borderDef)
                ThisApplication.SilentOperation = False
                log.AppendLine("AddCustomBorder С SilentOperation=True — OK")
            Catch ex2 As Exception
                ThisApplication.SilentOperation = False
                log.AppendLine("AddCustomBorder С SilentOperation=True — ОШИБКА: " & ex2.Message)
            End Try
        End Try
        ' Проверяем результат
        Try
            If sheet.Border IsNot Nothing Then
                log.AppendLine("sheet.Border после операции: ЕСТЬ (" & sheet.Border.Name & ")")
            Else
                log.AppendLine("sheet.Border после операции: Nothing (рамка НЕ применилась)")
            End If
        Catch ex As Exception
            log.AppendLine("Проверка sheet.Border: ошибка: " & ex.Message)
        End Try
    End If

    ' --- Шаг 7: AddTitleBlock ---
    If sheet IsNot Nothing AndAlso tbDef IsNot Nothing Then
        log.AppendLine("")
        log.AppendLine("--- AddTitleBlock ---")
        Dim ps(1) As String
        ps(1) = "ТЕСТ"
        ' Попытка 1: без SilentOperation
        Try
            sheet.AddTitleBlock(tbDef, Nothing, ps)
            log.AppendLine("AddTitleBlock БЕЗ SilentOperation — OK")
        Catch ex As Exception
            log.AppendLine("AddTitleBlock БЕЗ SilentOperation — ОШИБКА: " & ex.Message)
            ' Попытка 2: с SilentOperation
            Try
                ThisApplication.SilentOperation = True
                sheet.AddTitleBlock(tbDef, Nothing, ps)
                ThisApplication.SilentOperation = False
                log.AppendLine("AddTitleBlock С SilentOperation=True — OK")
            Catch ex2 As Exception
                ThisApplication.SilentOperation = False
                log.AppendLine("AddTitleBlock С SilentOperation=True — ОШИБКА: " & ex2.Message)
            End Try
        End Try
        Try
            If sheet.TitleBlock IsNot Nothing Then
                log.AppendLine("sheet.TitleBlock после операции: ЕСТЬ (" & sheet.TitleBlock.Name & ")")
            Else
                log.AppendLine("sheet.TitleBlock после операции: Nothing")
            End If
        Catch ex As Exception
            log.AppendLine("Проверка sheet.TitleBlock: " & ex.Message)
        End Try
    End If

    ' --- Показываем результат ---
    System.Windows.Forms.MessageBox.Show(
        log.ToString(),
        "Диагностика рамки и штампа",
        System.Windows.Forms.MessageBoxButtons.OK,
        System.Windows.Forms.MessageBoxIcon.Information)
End Sub
