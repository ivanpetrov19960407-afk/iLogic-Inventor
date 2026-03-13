# iLogic-Inventor — Альбом каменных изделий

## Что это

iLogic правило для Autodesk Inventor, которое автоматически собирает IDW-альбом из 3D-моделей каменных изделий (`.ipt`).

- Читает Excel-файл со списком моделей (лист `ALBUM`, 25 строк, пути к `.ipt`)
- Для каждой модели создаёт лист A3 альбомный в IDW
- Рисует рамку СПДС (Форма 3) через `AddBorder`
- Применяет штамп (TitleBlock) через `AddTitleBlock` с данными из Excel
- Размещает 4 вида: LARGE (фронт/профиль), SMALL (торец), WIDE (сверху/сбоку), ISO (изометрия тонированная)

## Основной файл

`docs/iLogic/StoneAlbumRule.ilogic.vb` — **текущая рабочая версия**

Вставить в Inventor: Управление → iLogic → Правила → Создать → вставить код → запустить из открытого `.idw`.

## Диагностическое правило

`docs/iLogic/DiagBorderTest.ilogic.vb` — диагностика наличия рамки и штампа в документе

## Константы в коде (критически важные)

```
BORDER_NAME = "RKM_SPDS_A3_BORDER_V12"   ← имя рамки в Образец.idw
TB_NAME     = "RKM_SPDS_A3_FORM3_V17"    ← имя штампа в Образец.idw
SHEET_PFX   = "ALB_"                     ← префикс создаваемых листов
```

Если в документе нет этих определений — правило создаёт их программно.

## Excel (ALBUM-RKM.xlsx)

Лист `ALBUM`. Колонки:
- `MODEL_PATH` — путь к `.ipt` (относительный от папки workspace или абсолютный)
- `CODE`, `PROJECT_NAME`, `DRAWING_NAME`, `ORG_NAME`, `STAGE` — поля штампа
- Номера листов `SHEET`/`SHEETS` заполняются автоматически

## Критические ограничения iLogic (НЕ НАРУШАТЬ)

1. **NO** inline `Try : код : Catch : End Try` на одной строке — каждый оператор на отдельной строке
2. `Alias` — зарезервировано VB.NET, не использовать как имя переменной
3. `shared` — зарезервировано, не использовать
4. **Нет** `Imports System.IO` — только полные имена: `System.IO.Path.`, `System.IO.File.`
5. **Нет** `Imports System.Text.RegularExpressions` — только `System.Text.RegularExpressions.Regex.`
6. `sheet.Border` и `sheet.TitleBlock` — **ReadOnly**, присваивать нельзя
7. `AddCustomBorder` **НЕ СУЩЕСТВУЕТ** в iLogic → правильный метод: **`AddBorder`**
8. `SilentOperation = True` нужен вокруг `AddBorder` и `AddTitleBlock`

## Архитектура кода

```
Main()
  └── AlbumBuilder.Build()
        ├── EnsureBorder(doc)      ← создаёт/пересоздаёт BorderDefinition через Edit/ExitEdit
        ├── EnsureTitleBlock(doc)  ← если уже есть — не трогает геометрию
        ├── PurgeAlbumSheets()     ← удаляет старые ALB_* листы
        └── BuildOneSheet()        ← для каждой модели из Excel
              ├── doc.Sheets.Add() → лист ALB_NNN
              ├── sheet.AddBorder(borderDef)           [SilentOperation=True]
              ├── sheet.AddTitleBlock(tbDef, , ps)     [SilentOperation=True]
              ├── Documents.Open(modelPath)
              └── PlaceViewsSlotBased()
                    ├── MeasureView × 4 (probe-виды на масштабе 0.1)
                    ├── 6 перестановок → выбрать лучший матчинг вид→слот
                    ├── PlaceViewInSlot × 4
                    └── AddDimNotes × 4 (текстовые ноты с размерами в мм)
```

## Геометрия слотов (А3, абсолютные значения мм)

```
Рабочая зона: X=[20, 415], Y=[65, 292]  (над штампом 55мм)

┌────────────────────────────────────────────────────┐ Y=292
│  SMALL (20..119, 229..292)  │  WIDE (124..415, 229..292)  │
├───────────────────────┬────────────────────────────┤ Y=229
│                       │                            │
│   LARGE               │    ISO                     │
│  (20..276, 65..224)   │  (281..415, 65..224)       │
│                       │                            │
└───────────────────────┴────────────────────────────┘ Y=65
X=20                  X=281                        X=415
```

## Статус (последнее обновление: март 2026)

| Что | Статус |
|-----|--------|
| Рамка AddBorder | ✅ работает |
| Штамп AddTitleBlock | ✅ работает (ps массив 7 элементов, 0-based) |
| Чтение Excel (ZIP+XML, без COM) | ✅ работает |
| Layout видов (4 слота) | 🔧 в работе — слоты фиксированы, масштаб уточняется |
| Размеры на видах | ✅ DrawingNotes с мм |

## VBA-источники (legacy, только для справки)

`legacy_vba/` или `/home/user/workspace/*.bas` на рабочей машине:
- `RKM_IdwAlbum-4.bas` — главный VBA (2128 строк), эталон архитектуры
- `RKM_FrameBorder-3.bas` — VBA рамка
- `RKM_TitleBlockPrompted-7.bas` — VBA штамп (7 промптов: CODE/PROJECT_NAME/DRAWING_NAME/ORG_NAME/STAGE/SHEET/SHEETS)
- `RKM_Excel-2.bas` — VBA Excel reader
