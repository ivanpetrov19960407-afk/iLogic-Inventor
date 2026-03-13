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
            Try
                ThisApplication.SilentOperation = True
                borderDef = doc.BorderDefinitions.Add(BORDER_NAME)
                log.AppendLine("Рамка создана через .Add()")
            Finally
                ThisApplication.SilentOperation = False
            End Try
        Catch ex As Exception
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
            Try
                ThisApplication.SilentOperation = True
                tbDef = doc.TitleBlockDefinitions.Add(TB_NAME)
                log.AppendLine("Штамп создан через .Add()")
            Finally
                ThisApplication.SilentOperation = False
            End Try
        Catch ex As Exception
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

    ' --- Шаг 6: AddBorder ---
    If sheet IsNot Nothing AndAlso borderDef IsNot Nothing Then
        log.AppendLine("")
        log.AppendLine("--- AddBorder ---")
        Try
            If sheet.Border IsNot Nothing Then
                sheet.Border.Delete()
            End If
            log.AppendLine("Существующая рамка удалена перед AddBorder")
        Catch ex As Exception
            log.AppendLine("Удаление существующей рамки: " & ex.Message)
        End Try

        Dim addBorderOk As Boolean = False
        Try
            Try
                ThisApplication.SilentOperation = True
                sheet.AddBorder(borderDef)
                addBorderOk = True
            Finally
                ThisApplication.SilentOperation = False
            End Try
            log.AppendLine("AddBorder OK")
        Catch ex As Exception
            log.AppendLine("AddBorder ОШИБКА: " & ex.Message)
        End Try

        log.AppendLine("AddBorder success flag = " & addBorderOk)
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
        Dim ps() As String = New String(6) {}
        ps(0) = "ТЕСТ-КОД"
        ps(1) = "Тестовый проект"
        ps(2) = "Тестовый чертёж"
        ps(3) = "Тестовая организация"
        ps(4) = "ИД"
        ps(5) = "1"
        ps(6) = "1"

        Try
            If sheet.TitleBlock IsNot Nothing Then
                sheet.TitleBlock.Delete()
            End If
            log.AppendLine("Существующий штамп удалён перед AddTitleBlock")
        Catch ex As Exception
            log.AppendLine("Удаление существующего штампа: " & ex.Message)
        End Try

        Dim addTitleBlockOk As Boolean = False
        Try
            Try
                ThisApplication.SilentOperation = True
                sheet.AddTitleBlock(tbDef, , ps)
                addTitleBlockOk = True
            Finally
                ThisApplication.SilentOperation = False
            End Try
            log.AppendLine("AddTitleBlock OK")
        Catch ex As Exception
            log.AppendLine("AddTitleBlock ОШИБКА: " & ex.Message)
        End Try

        log.AppendLine("AddTitleBlock success flag = " & addTitleBlockOk)
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
