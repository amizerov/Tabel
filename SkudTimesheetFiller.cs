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
    AmbiguouslFound
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
        var surnameIndex = BuildSurnameIndex(eventsByPerson);
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
            var searchResult = FindPersonInSkudData(fioFromTimesheet, eventsByPerson, surnameIndex);

            if (searchResult.Status == PersonSearchResult.NotFound)
            {
                notFoundFios.Add(fioFromTimesheet);
                HighlightCellNotFound(sheet.Cells[row, fioCol], searchResult.SearchKey);
                continue;
            }
            else if (searchResult.Status == PersonSearchResult.AmbiguouslFound)
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

                cell.Value2 = null;

                if (!personDays.TryGetValue(date, out var dayEvents))
                    continue;

                double hours = dayEvents.GetHours();

                cell.Value2 = hours;
                SetNumberFormatSafe(cell);

                filledCellsCount++;
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
        Dictionary<string, List<string>> surnameIndex)
    {
        var result = new PersonSearchInfo();
        string normalizedFio = NormalizeFio(fioFromTimesheet);
        result.SearchKey = FirstWord(normalizedFio);

        // Шаг 1: точное совпадение
        if (eventsByPerson.ContainsKey(normalizedFio))
        {
            result.Status = PersonSearchResult.Found;
            result.PersonKey = normalizedFio;
            return result;
        }

        // Шаг 2 & 3: поиск по ключу Фамилия + первая буква имени
        string surname = FirstWord(normalizedFio);

        if (surnameIndex.TryGetValue(surname, out var keys))
        {
            if (keys.Count == 1)
            {
                result.Status = PersonSearchResult.Found;
                result.PersonKey = keys[0];
                return result;
            }
            else if (keys.Count > 1)
            {
                result.Status = PersonSearchResult.AmbiguouslFound;
                result.FoundVariants = keys.ToList();
                return result;
            }
        }

        // Шаг 4: не найдено
        result.Status = PersonSearchResult.NotFound;
        return result;
    }

    private static Dictionary<string, Dictionary<DateTime, DayEvents>> ReadSkudExportFast(dynamic sheet)
    {
        Report("Чтение выгрузки", "Читаю диапазон Excel в память одним блоком...", 0, 100, 12);

        dynamic usedRange = sheet.UsedRange;

        int usedRows = usedRange.Rows.Count;
        int usedCols = usedRange.Columns.Count;

        object[,] values = (object[,])usedRange.Value2;

        int headerRow = -1;
        int dateCol = -1;
        int personCol = -1;

        Report("Анализ выгрузки", "Ищу колонки Дата записи и Персона...", 0, 100, 14);

        // Поиск колонок в первых 50 строках
        for (int row = 1; row <= Math.Min(50, usedRows); row++)
        {
            if (headerRow > 0 && dateCol > 0 && personCol > 0)
                break;

            for (int col = 1; col <= usedCols; col++)
            {
                string text = GetValueText(values[row, col]);
                string normalized = NormalizeHeader(text);

                if (normalized == "дата записи")
                    dateCol = col;
                else if (normalized == "персона")
                    personCol = col;

                if (dateCol > 0 && personCol > 0)
                {
                    headerRow = row;
                    break;
                }
            }
        }

        if (headerRow < 0 || dateCol < 0)
            throw new Exception("В выгрузке СКУД не найдена колонка 'Дата записи'.");

        if (personCol < 0)
            throw new Exception("В выгрузке СКУД не найдена колонка 'Персона'.");

        Report("Обработка данных", "Обрабатываю данные выгрузки...", 0, 100, 20);

        int totalRows = usedRows - headerRow;
        var result = new Dictionary<string, Dictionary<DateTime, DayEvents>>(StringComparer.OrdinalIgnoreCase);

        // Обработка данных одним проходом без промежуточных проверок
        int processedRows = 0;
        for (int row = headerRow + 1; row <= usedRows; row++)
        {
            try
            {
                object personObj = values[row, personCol];
                object dateObj = values[row, dateCol];

                if (personObj == null || dateObj == null)
                    continue;

                string person = GetValueText(personObj);
                if (string.IsNullOrWhiteSpace(person))
                    continue;

                if (!TryReadDateTimeFromValue(dateObj, out DateTime eventDateTime))
                    continue;

                string personKey = NormalizeFio(person);
                DateTime date = eventDateTime.Date;

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

                dayEvents.Add(eventDateTime);
                processedRows++;
            }
            catch
            {
                // Игнорируем ошибки в отдельных строках
            }

            // Реже обновляем прогресс (каждые 200 строк)
            if (processedRows > 0 && (row - headerRow) % 200 == 0)
            {
                int percent = ScalePercent(row - headerRow, totalRows, 20, 40);
                Report(
                    "Чтение выгрузки",
                    $"Обработано строк: {processedRows} из {totalRows}",
                    row - headerRow,
                    totalRows,
                    percent);
            }
        }

        Report("Чтение выгрузки", $"Завершено. Найдено персон: {result.Count}", totalRows, totalRows, 40);

        if (result.Count == 0)
            throw new Exception("В выгрузке СКУД не найдено данных.");

        return result;
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
        private DateTime _first = DateTime.MaxValue;
        private DateTime _last = DateTime.MinValue;

        public int Count { get; private set; }

        public void Add(DateTime dateTime)
        {
            if (dateTime < _first)
                _first = dateTime;

            if (dateTime > _last)
                _last = dateTime;

            Count++;
        }

        public double GetHours()
        {
            if (Count == 1)
                return 2;

            if (Count > 1)
            {
                double hours = (_last - _first).TotalHours - 1;
                if (hours < 0)
                    return 0;
                return Math.Ceiling(hours);
            }

            return 0;
        }
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

    private static Dictionary<string, List<string>> BuildSurnameIndex(
        Dictionary<string, Dictionary<DateTime, DayEvents>> eventsByPerson)
    {
        var result = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

        foreach (string fullFio in eventsByPerson.Keys)
        {
            string surname = FirstWord(fullFio);

            if (string.IsNullOrWhiteSpace(surname))
                continue;

            if (!result.TryGetValue(surname, out var list))
            {
                list = new List<string>();
                result[surname] = list;
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
