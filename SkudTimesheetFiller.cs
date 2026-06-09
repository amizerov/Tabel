namespace Tabel;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

internal enum PersonSearchResult
{
    Found,
    NotFound,
    AmbiguousFound
}

internal class PersonSearchInfo
{
    public PersonSearchResult Status { get; set; }
    public string? PersonKey { get; set; }
    public string SearchKey { get; set; } = string.Empty;
    public List<string> FoundVariants { get; set; } = new();
}

public static class SkudExcelInteropFiller
{
    public static event EventHandler<SkudProgressEventArgs>? ProgressChanged;

    public static Task<string> FillTimesheetAsync(string timesheetPath, string skudPath)
    {
        return RunStaTask(() => FillTimesheetFile(timesheetPath, skudPath));
    }

    public static Task<string> FillTimesheetFromFolderAsync(string folderPath)
    {
        return RunStaTask(() => FillTimesheetFromFolder(folderPath));
    }

    private static Task<T> RunStaTask<T>(Func<T> action)
    {
        var tcs = new TaskCompletionSource<T>();

        var thread = new Thread(() =>
        {
            try
            {
                T result = action();
                tcs.SetResult(result);
            }
            catch (Exception ex)
            {
                tcs.SetException(ex);
            }
        });

        thread.IsBackground = true;
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();

        return tcs.Task;
    }

    public static string FillTimesheetFile(string timesheetPath, string skudPath)
    {
        if (!File.Exists(timesheetPath))
            throw new FileNotFoundException("Табель подрядчика не найден.", timesheetPath);

        if (!File.Exists(skudPath))
            throw new FileNotFoundException("Выгрузка СКУД не найдена.", skudPath);

        string outputPath = Path.Combine(
            Path.GetDirectoryName(timesheetPath) ?? "",
            Path.GetFileNameWithoutExtension(timesheetPath) + "_заполнен.xlsx");

        Type? excelType = Type.GetTypeFromProgID("Excel.Application");

        if (excelType == null)
            throw new Exception("Microsoft Excel не установлен или COM-объект Excel.Application недоступен.");

        dynamic? excel = null;
        dynamic? skudBook = null;
        dynamic? timesheetBook = null;

        try
        {
            Report("Запуск Excel", "Запускаю Microsoft Excel...", 0, 100, 5);

            excel = Activator.CreateInstance(excelType);
            excel.Visible = false;
            excel.DisplayAlerts = false;

            try
            {
                excel.ScreenUpdating = false;
                excel.EnableEvents = false;
                excel.Calculation = -4135;
            }
            catch { }

            Report("Открытие выгрузки СКУД", "Открываю выгрузку СКУД...", 0, 100, 10);

            skudBook = excel.Workbooks.Open(skudPath);
            dynamic skudSheet = skudBook.Worksheets[1];

            var eventsByPerson = ReadSkudExportFast(skudSheet);

            skudBook.Close(false);
            ReleaseComObject(skudBook);
            skudBook = null;

            Report("Открытие табеля подрядчика", "Открываю табель подрядчика...", 0, 100, 45);

            timesheetBook = excel.Workbooks.Open(timesheetPath);
            dynamic timesheetSheet = timesheetBook.Worksheets[1];

            FillTimesheet(timesheetSheet, eventsByPerson);

            Report("Сохранение", "Сохраняю заполненный табель...", 0, 100, 95);

            if (File.Exists(outputPath))
                File.Delete(outputPath);

            timesheetBook.SaveAs(outputPath);

            timesheetBook.Close(false);
            ReleaseComObject(timesheetBook);
            timesheetBook = null;

            Report("Готово", "Табель успешно заполнен.", 100, 100, 100, true);

            return outputPath;
        }
        finally
        {
            try
            {
                if (skudBook != null)
                    skudBook.Close(false);
            }
            catch { }

            try
            {
                if (timesheetBook != null)
                    timesheetBook.Close(false);
            }
            catch { }

            if (skudBook != null)
                ReleaseComObject(skudBook);

            if (timesheetBook != null)
                ReleaseComObject(timesheetBook);

            if (excel != null)
            {
                try
                {
                    excel.Quit();
                }
                catch { }

                ReleaseComObject(excel);
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }

    public static string FillTimesheetFromFolder(string folderPath)
    {
        string exportPath = Path.Combine(folderPath, "выгрузка_СКУД.xlsx");
        string timesheetPath = Path.Combine(folderPath, "Табель_СКУД.xlsx");
        string outputPath = Path.Combine(folderPath, "Табель_СКУД_заполнен.xlsx");

        Report("Проверка файлов", "Проверяю наличие файлов...", 0, 100, 0);

        if (!File.Exists(exportPath))
            throw new FileNotFoundException("Не найден файл выгрузка_СКУД.xlsx", exportPath);

        if (!File.Exists(timesheetPath))
            throw new FileNotFoundException("Не найден файл Табель_СКУД.xlsx", timesheetPath);

        Type? excelType = Type.GetTypeFromProgID("Excel.Application");

        if (excelType == null)
            throw new Exception("Microsoft Excel не установлен или COM-объект Excel.Application недоступен.");

        dynamic? excel = null;
        dynamic? exportBook = null;
        dynamic? timesheetBook = null;

        try
        {
            Report("Запуск Excel", "Запускаю Microsoft Excel...", 0, 100, 5);

            excel = Activator.CreateInstance(excelType);
            excel.Visible = false;
            excel.DisplayAlerts = false;

            try
            {
                excel.ScreenUpdating = false;
                excel.EnableEvents = false;
                excel.Calculation = -4135;
            }
            catch
            {
            }

            Report("Открытие выгрузки", "Открываю файл выгрузка_СКУД.xlsx...", 0, 100, 10);

            exportBook = excel.Workbooks.Open(exportPath);
            dynamic exportSheet = exportBook.Worksheets[1];

            var eventsByPerson = ReadSkudExportFast(exportSheet);

            exportBook.Close(false);
            ReleaseComObject(exportBook);
            exportBook = null;

            Report("Открытие табеля", "Открываю файл Табель_СКУД.xlsx...", 0, 100, 45);

            timesheetBook = excel.Workbooks.Open(timesheetPath);
            dynamic timesheetSheet = timesheetBook.Worksheets[1];

            FillTimesheet(timesheetSheet, eventsByPerson);

            Report("Сохранение", "Сохраняю заполненный табель...", 0, 100, 95);

            if (File.Exists(outputPath))
                File.Delete(outputPath);

            timesheetBook.SaveAs(outputPath);

            timesheetBook.Close(false);
            ReleaseComObject(timesheetBook);
            timesheetBook = null;

            Report("Готово", "Табель успешно заполнен.", 100, 100, 100, true);

            return outputPath;
        }
        finally
        {
            try
            {
                if (exportBook != null)
                    exportBook.Close(false);
            }
            catch { }

            try
            {
                if (timesheetBook != null)
                    timesheetBook.Close(false);
            }
            catch { }

            if (exportBook != null)
                ReleaseComObject(exportBook);

            if (timesheetBook != null)
                ReleaseComObject(timesheetBook);

            if (excel != null)
            {
                try
                {
                    excel.Quit();
                }
                catch { }

                ReleaseComObject(excel);
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }

    private static void FillTimesheet(
        dynamic sheet,
        Dictionary<string, Dictionary<DateTime, DayEvents>> eventsByPerson)
    {
        Report("Анализ табеля", "Читаю табель в память...", 0, 100, 50);

        dynamic usedRange = sheet.UsedRange;
        int usedRows = usedRange.Rows.Count;
        int usedCols = usedRange.Columns.Count;

        object[,] values = (object[,])usedRange.Value2;

        int fioHeaderRow = -1;
        int fioCol = -1;

        Report("Анализ табеля", "Ищу колонку ФИО и колонки дней месяца...", 0, 100, 50);

        // Поиск заголовка ФИО
        for (int row = 1; row <= Math.Min(50, usedRows); row++)
        {
            for (int col = 1; col <= usedCols; col++)
            {
                string text = GetValueText(values[row, col]);

                if (NormalizeHeader(text) == "фио")
                {
                    fioHeaderRow = row;
                    fioCol = col;
                    break;
                }
            }

            if (fioHeaderRow > 0)
                break;
        }

        if (fioHeaderRow < 0 || fioCol < 0)
            throw new Exception("В табеле не найдена колонка 'ФИО'.");

        var dayColumns = FindDayColumns(sheet, fioHeaderRow, usedCols);

        if (dayColumns.Count == 0)
            throw new Exception("В табеле не найдены колонки дней месяца: 1, 2, 3 ...");

        DateTime monthDate = GetMainMonth(eventsByPerson);
        int year = monthDate.Year;
        int month = monthDate.Month;
        int daysInMonth = DateTime.DaysInMonth(year, month);

        Report(
            "Анализ табеля",
            $"Месяц выгрузки определен как {month:00}.{year}. Дней в месяце: {daysInMonth}.",
            0,
            100,
            52);

        // Сборка списка ФИО из табеля за один проход через массив
        var personKeyIndex = BuildPersonKeyIndex(eventsByPerson);
        var peopleFromTimesheet = new List<(int row, string fio)>();

        for (int row = fioHeaderRow + 1; row <= usedRows; row++)
        {
            string fio = GetValueText(values[row, fioCol]);

            if (!string.IsNullOrWhiteSpace(fio))
                peopleFromTimesheet.Add((row, fio));
        }

        int totalPeople = peopleFromTimesheet.Count;

        Report(
            "Заполнение табеля",
            $"В табеле найдено сотрудников: {totalPeople}. Начинаю заполнять табель.",
            0,
            totalPeople,
            55);

        int filledCellsCount = 0;
        var notFoundFios = new List<string>();
        var ambiguousFios = new List<(string fio, List<string> variants)>();

        for (int i = 0; i < peopleFromTimesheet.Count; i++)
        {
            var (row, fioFromTimesheet) = peopleFromTimesheet[i];

            // Ищем ФИО в данных СКУД
            var searchResult = FindPersonInSkudData(fioFromTimesheet, eventsByPerson, personKeyIndex);

            if (searchResult.Status == PersonSearchResult.NotFound)
            {
                notFoundFios.Add(fioFromTimesheet);
                HighlightCellNotFound(sheet.Cells[row, fioCol], searchResult.SearchKey);
                continue;
            }
            else if (searchResult.Status == PersonSearchResult.AmbiguousFound)
            {
                ambiguousFios.Add((fioFromTimesheet, searchResult.FoundVariants));
                HighlightCellAmbiguous(sheet.Cells[row, fioCol], searchResult.SearchKey, searchResult.FoundVariants);
                continue;
            }

            // Процесс заполнения часов
            var personDays = eventsByPerson[searchResult.PersonKey!];

            foreach (var pair in dayColumns)
            {
                int day = pair.Key;
                int col = pair.Value;

                if (day < 1 || day > daysInMonth)
                    continue;

                DateTime date = new DateTime(year, month, day);

                dynamic cell = sheet.Cells[row, col];

                // НЕ затирать значение заранее
                if (!personDays.TryGetValue(date, out var dayEvents))
                    continue;

                var calc = dayEvents.Calculate();

                if (calc == null)
                    continue;

                // Если есть проблемные события - выделяем красным и добавляем комментарий, не меняем значение
                if (calc.HasProblemEvents)
                {
                    try
                    {
                        cell.Interior.Color = RGB(255, 150, 150);
                    }
                    catch { }

                    try
                    {
                        AddCommentSafe(cell, string.Join("; ", calc.ProblemMessages));
                    }
                    catch { }

                    // логируем
                    // не менять значение ячейки
                    continue;
                }

                // Если одна разрешенная отметка
                if (calc.IsSingleAllowedEvent)
                {
                    try { cell.Interior.Color = RGB(189, 215, 238); } catch { }
                    try { AddCommentSafe(cell, calc.Comment ?? ""); } catch { }
                    // Записать 2
                    cell.Value2 = 2;
                    SetNumberFormatSafe(cell);
                    filledCellsCount++;
                    continue;
                }

                if (calc.ShouldWriteValue && calc.Hours.HasValue)
                {
                    cell.Value2 = calc.Hours.Value;
                    SetNumberFormatSafe(cell);
                    filledCellsCount++;
                }
                else
                {
                    // не менять/очищать
                }
            }

            int current = i + 1;
            int percent = ScalePercent(current, totalPeople, 55, 90);

            if (current % 5 == 0 || current == totalPeople)
            {
                Report(
                    "Заполнение табеля",
                    $"Заполняю сотрудников: {current} из {totalPeople}",
                    current,
                    totalPeople,
                    percent);
            }
        }

        Report(
            "Итог заполнения",
            $"Сотрудников найдено в табеле: {totalPeople}. Заполнено ячеек с часами: {filledCellsCount}. " +
            $"Не найдено в СКУД: {notFoundFios.Count}. Неоднозначных: {ambiguousFios.Count}.",
            totalPeople,
            totalPeople,
            92);
    }

    private static PersonSearchInfo FindPersonInSkudData(
        string fioFromTimesheet,
        Dictionary<string, Dictionary<DateTime, DayEvents>> eventsByPerson,
        Dictionary<string, List<string>> personKeyIndex)
    {
        var result = new PersonSearchInfo();
        string normalizedFio = NormalizeFio(fioFromTimesheet);
        string searchKey = BuildPersonSearchKey(fioFromTimesheet);
        result.SearchKey = string.IsNullOrEmpty(searchKey) ? FirstWord(normalizedFio) : searchKey;

        // Шаг 1: точное совпадение полного ФИО
        if (!string.IsNullOrWhiteSpace(normalizedFio) && eventsByPerson.ContainsKey(normalizedFio))
        {
            result.Status = PersonSearchResult.Found;
            result.PersonKey = normalizedFio;
            return result;
        }

        // Шаг 2: поиск по ключу Фамилия + первая буква имени
        if (!string.IsNullOrWhiteSpace(searchKey) && personKeyIndex.TryGetValue(searchKey, out var keys))
        {
            if (keys.Count == 1)
            {
                result.Status = PersonSearchResult.Found;
                result.PersonKey = keys[0];
                return result;
            }
            else if (keys.Count > 1)
            {
                // если среди вариантов есть точное совпадение по полному ФИО - используем его
                if (!string.IsNullOrWhiteSpace(normalizedFio))
                {
                    var exact = keys.FirstOrDefault(k => string.Equals(k, normalizedFio, StringComparison.OrdinalIgnoreCase));
                    if (exact != null)
                    {
                        result.Status = PersonSearchResult.Found;
                        result.PersonKey = exact;
                        return result;
                    }
                }

                result.Status = PersonSearchResult.AmbiguousFound;
                result.FoundVariants = keys.ToList();
                return result;
            }
        }

        // Шаг 3: не найдено
        result.Status = PersonSearchResult.NotFound;
        return result;
    }

    private sealed class SkudEvent
    {
        public DateTime DateTime { get; set; }
        public string Message { get; set; } = string.Empty;
        public bool IsAllowedEnter { get; set; }
        public bool IsAllowedExit { get; set; }
        public bool IsAllowed => IsAllowedEnter || IsAllowedExit;
        public bool IsProblem { get; set; }
    }

    private static string NormalizeMessage(string text)
    {
        if (text == null)
            return string.Empty;

        string t = text.Replace('\u00A0', ' ').Trim().ToLowerInvariant().Replace('ё','е');
        t = Regex.Replace(t, @"\s+", " ");
        return t;
    }

    private static Dictionary<string, Dictionary<DateTime, DayEvents>> ReadSkudExportFast(dynamic workbook)
    {
        Report("Чтение выгрузки", "Ищу лист с необходимыми колонками и читаю данные...", 0, 100, 12);

        Type t = workbook.GetType();
        int sheetsCount = workbook.Worksheets.Count;

        Dictionary<string, Dictionary<DateTime, DayEvents>> result = new(StringComparer.OrdinalIgnoreCase);

        for (int si = 1; si <= sheetsCount; si++)
        {
            try
            {
                dynamic sheet = workbook.Worksheets[si];
                dynamic usedRange = sheet.UsedRange;
                int usedRows = usedRange.Rows.Count;
                int usedCols = usedRange.Columns.Count;
                object[,] values = (object[,])usedRange.Value2;

                // Найти колонки Date, Person, Message в первых 50 строк
                int headerRow = -1;
                int dateCol = -1;
                int personCol = -1;
                int messageCol = -1;

                for (int row = 1; row <= Math.Min(50, usedRows); row++)
                {
                    for (int col = 1; col <= usedCols; col++)
                    {
                        string txt = GetValueText(values[row, col]);
                        string n = NormalizeHeader(txt);

                        if (n == "дата записи") dateCol = col;
                        else if (n == "персона") personCol = col;
                        else if (n == "сообщение" || n == "message") messageCol = col;

                        if (dateCol > 0 && personCol > 0 && messageCol > 0)
                        {
                            headerRow = row;
                            break;
                        }
                    }

                    if (headerRow > 0) break;
                }

                if (headerRow <= 0 || dateCol <= 0 || personCol <= 0 || messageCol <= 0)
                {
                    // Этот лист не подходит
                    continue;
                }

                Report("Чтение листа СКУД", $"Читаю лист {si} - строки: {usedRows}", 0, usedRows, 20);

                for (int row = headerRow + 1; row <= usedRows; row++)
                {
                    try
                    {
                        object personObj = values[row, personCol];
                        object dateObj = values[row, dateCol];
                        object msgObj = values[row, messageCol];

                        if (personObj == null || dateObj == null)
                            continue;

                        string person = GetValueText(personObj);
                        if (string.IsNullOrWhiteSpace(person))
                            continue;

                        if (!TryReadDateTimeFromValue(dateObj, out DateTime eventDt))
                            continue;

                        string message = GetValueText(msgObj);
                        string normMsg = NormalizeMessage(message);

                        var ev = new SkudEvent { DateTime = eventDt, Message = message };

                        // Классифицируем сообщение
                        if (normMsg.Contains("вход разреш") || normMsg.Contains("вход разрешен") || normMsg.Contains("вход разрешён"))
                            ev.IsAllowedEnter = true;
                        else if (normMsg.Contains("выход разреш") || normMsg.Contains("выход разрешен") || normMsg.Contains("выход разрешён"))
                            ev.IsAllowedExit = true;
                        else if (normMsg.Contains("повторный вход запрещ") || normMsg.Contains("повторный выход запрещ") || normMsg.Contains("выход запрещ") || normMsg.StartsWith("запрещено") || normMsg.Contains("заблокирован") || normMsg.Contains("несогласован") || normMsg.Contains("несовпада"))
                            ev.IsProblem = true;

                        string personKey = NormalizeFio(person);
                        DateTime date = eventDt.Date;

                        if (!result.TryGetValue(personKey, out var personDays))
                        {
                            personDays = new Dictionary<DateTime, DayEvents>();
                            result[personKey] = personDays;
                        }

                        if (!personDays.TryGetValue(date, out var dayEvents))
                        {
                            dayEvents = new DayEvents();
                            personDays[date] = dayEvents;
                        }

                        dayEvents.AddEvent(ev);
                    }
                    catch { }
                }

                // Если дошли сюда и считали хотя бы что-то — считаем чтение успешным
                if (result.Count > 0)
                    return result;
            }
            catch { }
        }

        throw new Exception("В файле СКУД не найден лист с обязательными колонками: Дата записи, Персона, Сообщение");
    }

    private record TimesheetStructure(
        int HeaderRow,
        int DateRow,
        int FirstEmployeeRow,
        int FioCol,
        Dictionary<DateTime,int> DateColumns,
        Dictionary<int,int> DayNumberColumns,
        int? TotalHoursCol
    );

    private static TimesheetStructure DetectTimesheetStructure(dynamic sheet, object[,] values, int usedRows, int usedCols, DateTime? fallbackMonth)
    {
        // Ищем строку заголовков по ячейке с "ФИО"
        int headerRow = -1;
        int fioCol = -1;

        for (int row = 1; row <= Math.Min(15, usedRows); row++)
        {
            for (int col = 1; col <= usedCols; col++)
            {
                string text = GetValueText(values[row, col]);
                if (NormalizeHeader(text) == "фио")
                {
                    headerRow = row;
                    fioCol = col;
                    break;
                }
            }

            if (headerRow > 0)
                break;
        }

        if (headerRow < 0 || fioCol < 0)
            throw new Exception("Не найдена строка заголовков с колонкой 'ФИО' в табеле.");

        // Ищем строку дат в пределах headerRow..headerRow+3
        int dateRow = -1;
        var dateColumns = new Dictionary<DateTime,int>();
        var dayNumberColumns = new Dictionary<int,int>();

        for (int row = headerRow; row <= Math.Min(headerRow + 3, usedRows); row++)
        {
            bool foundAny = false;

            for (int col = 1; col <= usedCols; col++)
            {
                if (col == fioCol) continue;

                object raw = values[row, col];

                if (raw == null)
                    continue;

                // Если double - Excel serial date
                if (raw is double d)
                {
                    try
                    {
                        DateTime dt = DateTime.FromOADate(d);
                        foundAny = true;
                        // сохраним предположительно - дата без года/месяца пока не известны
                        dateColumns[dt.Date] = col;
                    }
                    catch { }
                }
                else
                {
                    string txt = GetValueText(raw);
                    if (int.TryParse(txt, out int day) && day >=1 && day <=31)
                    {
                        foundAny = true;
                        dayNumberColumns[day] = col;
                    }
                    else
                    {
                        // попробуем распарсить как дата
                        if (DateTime.TryParse(txt, out DateTime dt2))
                        {
                            foundAny = true;
                            dateColumns[dt2.Date] = col;
                        }
                    }
                }
            }

            if (foundAny)
            {
                dateRow = row;
                break;
            }
        }

        if (dateRow < 0)
            throw new Exception("Не найдена строка дат рядом с заголовком табеля.");

        // Теперь попытка сопоставить dayNumberColumns с fallbackMonth
        if (dateColumns.Count == 0 && dayNumberColumns.Count > 0)
        {
            if (fallbackMonth == null)
                throw new Exception("В табеле найдены только номера дней без месяца. Нельзя однозначно определить месяц для дат.");

            int year = fallbackMonth.Value.Year;
            int month = fallbackMonth.Value.Month;

            foreach (var kv in dayNumberColumns)
            {
                int day = kv.Key;
                int col = kv.Value;
                try
                {
                    var dt = new DateTime(year, month, day);
                    dateColumns[dt] = col;
                }
                catch { }
            }
        }

        // Определим первую строку сотрудников - строка после dateRow
        int firstEmployeeRow = dateRow + 1;

        // Попробуем найти колонку итого часов - ищем заголовок содержащий "итог" или "итого"
        int? totalCol = null;
        for (int col = 1; col <= usedCols; col++)
        {
            string h = GetValueText(values[headerRow, col]);
            string nh = NormalizeHeader(h);
            if (nh.Contains("итог") || nh.Contains("итого") || nh.Contains("итог часов"))
            {
                totalCol = col;
                break;
            }
        }

        return new TimesheetStructure(headerRow, dateRow, firstEmployeeRow, fioCol, dateColumns, dayNumberColumns, totalCol);
    }

    private static Dictionary<int, int> FindDayColumns(dynamic sheet, int headerRow, int usedCols)
    {
        var result = new Dictionary<int, int>();

        dynamic usedRange = sheet.UsedRange;
        object[,] values = (object[,])usedRange.Value2;

        for (int col = 1; col <= usedCols; col++)
        {
            string text = GetValueText(values[headerRow, col]);

            if (int.TryParse(text, out int day))
            {
                if (day >= 1 && day <= 31)
                    result[day] = col;
            }
        }

        return result;
    }

    private static DateTime GetMainMonth(Dictionary<string, Dictionary<DateTime, DayEvents>> eventsByPerson)
    {
        var counter = new Dictionary<(int Year, int Month), int>();

        foreach (var person in eventsByPerson.Values)
        {
            foreach (var date in person.Keys)
            {
                var key = (date.Year, date.Month);

                if (!counter.ContainsKey(key))
                    counter[key] = 0;

                counter[key]++;
            }
        }

        var main = counter.OrderByDescending(x => x.Value).First().Key;

        return new DateTime(main.Year, main.Month, 1);
    }

    private sealed class DayEvents
    {
        public List<SkudEvent> Events { get; } = new List<SkudEvent>();

        public void AddEvent(SkudEvent ev)
        {
            Events.Add(ev);
            Events.Sort((a,b) => a.DateTime.CompareTo(b.DateTime));
        }

        public DayCalculationResult Calculate()
        {
            var res = new DayCalculationResult();

            if (Events == null || Events.Count == 0)
            {
                res.HasAllowedEvents = false;
                res.ShouldWriteValue = false;
                return res;
            }

            var allowed = Events.Where(e => e.IsAllowed).ToList();
            var problems = Events.Where(e => e.IsProblem).ToList();

            res.ProblemMessages = problems.Select(p => p.Message).ToList();
            res.HasProblemEvents = problems.Count > 0;
            res.HasAllowedEvents = allowed.Count > 0;

            if (res.HasProblemEvents)
            {
                // при проблемных сообщениях не менять значение ячейки
                res.ShouldWriteValue = false;
                res.Comment = string.Join("; ", res.ProblemMessages);
                return res;
            }

            if (!res.HasAllowedEvents)
            {
                res.ShouldWriteValue = false;
                return res;
            }

            if (allowed.Count == 1)
            {
                res.IsSingleAllowedEvent = true;
                res.Hours = 2;
                res.ShouldWriteValue = true;
                res.Comment = "За дату найдена только одна разрешенная отметка СКУД";
                return res;
            }

            // 2 или более разрешенных - взять первый и последний
            var first = allowed.First().DateTime;
            var last = allowed.Last().DateTime;

            double totalHours = (last - first).TotalHours - 1.0;
            if (totalHours < 0)
                totalHours = 0;

            int hoursWhole = (int)Math.Floor(totalHours);
            double minutes = (totalHours - hoursWhole) * 60.0;

            if (minutes >= 40)
                hoursWhole += 1;

            res.Hours = hoursWhole;
            res.ShouldWriteValue = true;

            return res;
        }
    }

    private class DayCalculationResult
    {
        public bool HasAllowedEvents { get; set; }
        public bool HasProblemEvents { get; set; }
        public bool ShouldWriteValue { get; set; }
        public double? Hours { get; set; }
        public bool IsSingleAllowedEvent { get; set; }
        public List<string> ProblemMessages { get; set; } = new List<string>();
        public string? Comment { get; set; }
    }

    private static bool TryReadDateTimeFromValue(object value, out DateTime result)
    {
        result = default;

        try
        {
            if (value == null)
                return false;

            if (value is double d)
            {
                result = DateTime.FromOADate(d);
                return true;
            }

            if (value is DateTime dt)
            {
                result = dt;
                return true;
            }

            string text = Convert.ToString(value)?.Trim() ?? "";

            if (string.IsNullOrWhiteSpace(text))
                return false;

            var ru = CultureInfo.GetCultureInfo("ru-RU");

            string[] formats =
            {
                "dd.MM.yyyy H:mm:ss",
                "dd.MM.yyyy HH:mm:ss",
                "dd.MM.yyyy H:mm",
                "dd.MM.yyyy HH:mm",
                "yyyy-MM-dd H:mm:ss",
                "yyyy-MM-dd HH:mm:ss",
                "yyyy-MM-dd H:mm",
                "yyyy-MM-dd HH:mm"
            };

            if (DateTime.TryParseExact(text, formats, ru, DateTimeStyles.None, out result))
                return true;

            return DateTime.TryParse(text, ru, DateTimeStyles.None, out result);
        }
        catch
        {
            return false;
        }
    }

    private static bool TryReadDateTime(dynamic cell, out DateTime value)
    {
        value = default;

        try
        {
            object raw = cell.Value2;

            if (raw == null)
                return false;

            if (raw is double d)
            {
                value = DateTime.FromOADate(d);
                return true;
            }

            if (raw is DateTime dt)
            {
                value = dt;
                return true;
            }

            string text = Convert.ToString(raw)?.Trim() ?? "";

            if (string.IsNullOrWhiteSpace(text))
                return false;

            var ru = CultureInfo.GetCultureInfo("ru-RU");

            string[] formats =
            {
                "dd.MM.yyyy H:mm:ss",
                "dd.MM.yyyy HH:mm:ss",
                "dd.MM.yyyy H:mm",
                "dd.MM.yyyy HH:mm",
                "yyyy-MM-dd H:mm:ss",
                "yyyy-MM-dd HH:mm:ss",
                "yyyy-MM-dd H:mm",
                "yyyy-MM-dd HH:mm"
            };

            if (DateTime.TryParseExact(text, formats, ru, DateTimeStyles.None, out value))
                return true;

            return DateTime.TryParse(text, ru, DateTimeStyles.None, out value);
        }
        catch
        {
            return false;
        }
    }

    private static string GetCellText(dynamic cell)
    {
        try
        {
            object raw = cell.Value2;

            if (raw == null)
                return "";

            return Convert.ToString(raw)?.Trim() ?? "";
        }
        catch
        {
            return "";
        }
    }

    private static string GetValueText(object value)
    {
        try
        {
            if (value == null)
                return "";

            return Convert.ToString(value)?.Trim() ?? "";
        }
        catch
        {
            return "";
        }
    }

    private static string BuildPersonSearchKey(string fio)
        {
            if (string.IsNullOrWhiteSpace(fio))
                return string.Empty;

            string normalized = NormalizeFio(fio);
            var parts = normalized.Split(' ', StringSplitOptions.RemoveEmptyEntries);

            if (parts.Length < 2)
                return string.Empty; // нельзя сформировать ключ

            string surname = parts[0];
            string firstInitial = parts[1].Substring(0, 1);

            return surname + " " + firstInitial;
        }

        private static Dictionary<string, List<string>> BuildPersonKeyIndex(
            Dictionary<string, Dictionary<DateTime, DayEvents>> eventsByPerson)
        {
            var result = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

            foreach (string fullFio in eventsByPerson.Keys)
            {
                string key = BuildPersonSearchKey(fullFio);

                if (string.IsNullOrWhiteSpace(key))
                    continue;

                if (!result.TryGetValue(key, out var list))
                {
                    list = new List<string>();
                    result[key] = list;
                }

                list.Add(fullFio);
            }

            return result;
        }

    private static string FirstWord(string text)
    {
        text = NormalizeFio(text);

        if (string.IsNullOrWhiteSpace(text))
            return "";

        int spaceIndex = text.IndexOf(' ');

        if (spaceIndex < 0)
            return text;

        return text.Substring(0, spaceIndex);
    }

    private static string NormalizeFio(string text)
    {
        return NormalizeSpaces(text)
            .ToLowerInvariant()
            .Replace('ё', 'е');
    }

    private static string NormalizeHeader(string text)
    {
        return NormalizeSpaces(text)
            .ToLowerInvariant()
            .Replace('ё', 'е');
    }

    private static string NormalizeSpaces(string text)
    {
        if (text == null)
            return "";

        text = text.Replace('\u00A0', ' ');
        text = Regex.Replace(text, @"\s+", " ");

        return text.Trim();
    }

    private static int ScalePercent(int current, int total, int fromPercent, int toPercent)
    {
        if (total <= 0)
            return fromPercent;

        double k = (double)current / total;

        if (k < 0)
            k = 0;

        if (k > 1)
            k = 1;

        return fromPercent + (int)Math.Round((toPercent - fromPercent) * k);
    }

    private static void SetNumberFormatSafe(dynamic cell)
    {
        try
        {
            cell.NumberFormatLocal = "0";
            return;
        }
        catch
        {
        }

        try
        {
            cell.NumberFormat = "0";
            return;
        }
        catch
        {
        }
    }

    private static void HighlightCellNotFound(dynamic cell, string searchKey)
    {
        try
        {
            // Светло-красный цвет (RGB: 255, 200, 200)
            cell.Interior.Color = RGB(255, 200, 200);
        }
        catch { }

        try
        {
            string comment = $"ФИО не найдено в выгрузке СКУД.\nКлюч поиска: {searchKey}";
            cell.AddComment(comment);
        }
        catch { }
    }

    private static void HighlightCellAmbiguous(dynamic cell, string searchKey, List<string> variants)
    {
        try
        {
            // Оранжевый цвет (RGB: 255, 165, 0)
            cell.Interior.Color = RGB(255, 165, 0);
        }
        catch { }

        try
        {
            string variantsStr = string.Join(", ", variants);
            string comment = $"Неоднозначное совпадение в СКУД.\nКлюч поиска: {searchKey}\nНайденные варианты: {variantsStr}";
            cell.AddComment(comment);
        }
        catch { }
    }

    private static int RGB(int red, int green, int blue)
    {
        return red + (green << 8) + (blue << 16);
    }

    private static void AddCommentSafe(dynamic cell, string commentText)
    {
        try
        {
            // удалить старый комментарий, если есть
            try { var old = cell.Comment; if (old != null) old.Delete(); } catch { }
            cell.AddComment(commentText);
            try { cell.Comment.Shape.TextFrame.AutoSize = true; } catch { }
        }
        catch { }
    }

    private static void Report(
        string stage,
        string message,
        int current,
        int total,
        int percent,
        bool isFinished = false)
    {
        ProgressChanged?.Invoke(
            null,
            new SkudProgressEventArgs(
                stage,
                message,
                current,
                total,
                percent,
                isFinished));
    }

    private static void ReleaseComObject(object obj)
    {
        try
        {
            Marshal.FinalReleaseComObject(obj);
        }
        catch
        {
            // COM-объект уже мог быть освобожден.
        }
    }
}
