# Промт для миграции VBA `Work.ivb` в iLogic

Ниже — готовый промт для Codex/LLM, который помогает перенести логику старого VBA-макроса Inventor в структуру iLogic (VB.NET).

## Copy-Paste промт

**Role:** Ты — эксперт по автоматизации Autodesk Inventor (API) и опытный разработчик на VB.NET/C#.

**Task:** Мне нужно спроектировать архитектуру и написать код для системы автоматизации создания альбомов чертежей (.idw). Проект должен заменить мой старый VBA-макрос, сохранив следующую логику:

1. **Модуль Excel (`RKM_Excel`):**
   - Чтение таблицы Excel, где колонки — это пути к моделям (.ipt), а заголовки соответствуют именам полей в штампе чертежа (Title Block).
   - Реализуй умный поиск заголовков (алиасы): например, если колонка называется `CODE` или `Шифр`, она должна мапиться на одно и то же поле.

2. **Модуль оформления (`RKM_FrameBorder`, `RKM_SPDS`):**
   - Программное создание листов формата А3 в альбомной ориентации.
   - Отрисовка рамки по ГОСТ/СПДС и основной надписи (Форма 3).

3. **Модуль размещения видов (`RKM_IdwAlbum`):**
   - Автоматическое создание 3 ортогональных видов (спереди, сверху, слева) и одного изометрического вида.
   - **Важно:** реализуй алгоритм `FindBestScaleFor3Views`, который подбирает масштаб так, чтобы виды вписались в рабочую область листа, не перекрывая штамп.

4. **Модуль размеров (`AutoDimension`):**
   - Поиск крайних точек геометрии на видах и автоматическая простановка габаритных размеров (длина, высота, толщина изделия).
   - Используй логику из `AddHorizontalOverallDimension` и `AddVerticalOverallDimension`.

5. **Контекст проекта:**
   - Это изделия из натурального камня (гранит, мрамор).
   - Учитывай точность до миллиметра и перевод единиц (мм в см для расчетов).

**Output:** Напиши базовую структуру проекта на **iLogic (VB.NET)**. Раздели код на логические функции, чтобы я мог легко переносить их в Telegram-бот в будущем.

## Рекомендация по использованию

1. Вставьте промт в Codex/LLM.
2. Полученный код разложите по блокам:
   - `ExcelReader`
   - `SheetBuilder`
   - `ViewLayout`
   - `Dimensioning`
3. Сверяйте итоговую логику с исходными VBA-модулями из `legacy_vba/`.

---

## Prompt: исправление ошибок компиляции iLogic (RU)

Ниже — готовый промт для случаев, когда iLogic/VB.NET падает с ошибками вида:
- `Required one of these operators: 'Dim', 'Const', 'Public', ...`
- `Declaration expected`
- ошибка на ранних строках (часто `line 3`) из-за незакомментированного русского текста.

### Copy-Paste промт

**Role:** Senior Autodesk Inventor Developer (iLogic/VB.NET expert).

**Task:** Fix compilation errors in the provided iLogic code.

**Context:**  
I get errors like:
- `Required one of these operators: 'Dim', 'Const', 'Public', etc.`
- `Declaration expected` (for example at line 3)

This is usually caused by unquoted Russian text that is parsed as code instead of comments.

**Code to Fix (paste full script):**
```vbnet
[ВСТАВЬТЕ СЮДА ВЕСЬ ТЕКСТ ВАШЕГО СКРИПТА]
```

**Instructions for AI:**
1. **Fix Syntax**
   - Ensure all Russian descriptions/notes are commented out with a leading `'`.
   - Example:  
     `Установить А3 альбомная` → `' Установить А3 альбомная`.
2. **Strict iLogic/VB.NET**
   - Verify every non-comment line is valid VB.NET declaration or statement.
   - Ensure all executable lines are inside `Sub Main()` (or inside class members where appropriate).
   - Check `SpdsFramer` class members for valid declarations.
3. **Encoding**
   - Keep UTF-8 encoding to preserve Cyrillic text in comments and strings.
4. **Formatting**
   - Do not shorten code.
   - Return the full corrected document, ready for paste into iLogic Rule Editor.

### Что проверить вручную (быстрый чек-лист)

1. На проблемной ранней строке (например, строка 3) русский текст должен начинаться с `'`.
2. Перед `Sub ApplyBorder` и другими `Sub/Function` не должно быть «голого» текста.
3. Любая кириллица в коде должна быть:
   - либо в строковом литерале (`"Наименование объекта"`),
   - либо в комментарии (`' Наименование объекта`).
4. После вставки в iLogic убедитесь, что комментарии подсвечены как комментарии (обычно серым).

### Важно

Если не вставить **полный исходный скрипт**, модель не сможет безопасно исправить все места и вернёт только общие рекомендации.
