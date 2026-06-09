namespace Tabel;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

public static class SkudTimesheetComparator
{
    public static event EventHandler<SkudProgressEventArgs>? ProgressChanged;

    public static Task<string> CompareTimesheetsAsync(string originalPath, string filledPath)
    {
        return RunStaTask(() => CompareTimesheets(originalPath, filledPath));
    }

    public static string CompareTimesheets(string originalPath, string filledPath)
    {
        if (!File.Exists(originalPath))
            throw new FileNotFoundException("Исходный табель не найден.", originalPath);

        if (!File.Exists(filledPath))
            throw new FileNotFoundException("Заполненный табель не найден.", filledPath);

        string outputPath = Path.Combine(
            Path.GetDirectoryName(originalPath) ?? "",
            Path.GetFileNameWithoutExtension(originalPath) + "_сравнение.xlsx");

        string protocolPath = Path.Combine(
            Path.GetDirectoryName(originalPath) ?? "",
            Path.GetFileNameWithoutExtension(originalPath) + "_протокол_сравнения.txt");

        Type? excelType = Type.GetTypeFromProgID("Excel.Application");
        if (excelType == null)
            throw new Exception("Microsoft Excel не установлен или COM-объект Excel.Application недоступен.");

        var protocol = new List<string>();
        dynamic? excel = null;
        dynamic? originalBook = null;
        dynamic? filledBook = null;
        dynamic? outputBook = null;

        try
        {
            Report("Запуск Excel", "Запускаю Microsoft Excel...", 0, 100, 5);

            if (File.Exists(outputPath))
                File.Delete(outputPath);

            File.Copy(filledPath, outputPath, true);

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

            Report("Чтение табелей", "Читаю исходный и заполненный табели...", 0, 100, 15);

            originalBook = excel.Workbooks.Open(originalPath);
            dynamic originalSheet = FindWorksheetWithTimesheet(originalBook, out TimesheetStructure originalStructure);
            object[,] originalValues = ToArray2D(((dynamic)originalSheet.UsedRange).Value2);
            int originalRows = originalSheet.UsedRange.Rows.Count;
            int originalCols = originalSheet.UsedRange.Columns.Count;
            originalStructure = originalStructure with { Values = originalValues, UsedRows = originalRows, UsedCols = originalCols };

            filledBook = excel.Workbooks.Open(filledPath);
            dynamic filledSheet = FindWorksheetWithTimesheet(filledBook, out TimesheetStructure filledStructure);
            object[,] filledValues = ToArray2D(((dynamic)filledSheet.UsedRange).Value2);
            int filledRows = filledSheet.UsedRange.Rows.Count;
            int filledCols = filledSheet.UsedRange.Columns.Count;
            filledStructure = filledStructure with { Values = filledValues, UsedRows = filledRows, UsedCols = filledCols };

            originalBook.Close(false);
            filledBook.Close(false);
            ReleaseComObject(originalBook);
            ReleaseComObject(filledBook);
            originalBook = null;
            filledBook = null;

            outputBook = excel.Workbooks.Open(outputPath);
            dynamic outputSheet = FindWorksheetWithTimesheet(outputBook, out TimesheetStructure _);

            Report("Сравнение табелей", "Сравниваю дневные часы...", 0, 100, 40);

            var stats = PerformComparisonWithData(
                originalValues,
                originalRows,
                originalCols,
                filledValues,
                filledRows,
                filledCols,
                originalStructure,
                filledStructure,
                outputSheet,
                protocol);

            Report("Сохранение", "Сохраняю файл сравнения...", 0, 100, 95);
            outputBook.Save();
            outputBook.Close(false);
            ReleaseComObject(outputBook);
            outputBook = null;

            WriteProtocol(protocolPath, protocol);

            string summary =
                $"Сравнение завершено. Проверено сотрудников: {stats.CheckedEmployees}. " +
                $"Проверено ячеек: {stats.CheckedCells}. Существенных расхождений: {stats.SignificantDifferences}. " +
                $"Допустимых расхождений: {stats.AllowedDifferences}. Ошибок чтения: {stats.ReadErrors}. Файл: {outputPath}";

            Report("Завершено", summary, 100, 100, 100, true);
            return outputPath;
        }
        catch (Exception ex)
        {
            protocol.Add($"Критическая ошибка; {ex.Message}");
            try { WriteProtocol(protocolPath, protocol); } catch { }
            Report("ОШИБКА", ex.Message, 0, 100, 100, true);
            throw;
        }
        finally
        {
            try { if (outputBook != null) outputBook.Close(false); } catch { }
            try { if (originalBook != null) originalBook.Close(false); } catch { }
            try { if (filledBook != null) filledBook.Close(false); } catch { }

            if (outputBook != null)
                ReleaseComObject(outputBook);

            if (originalBook != null)
                ReleaseComObject(originalBook);

            if (filledBook != null)
                ReleaseComObject(filledBook);

            if (excel != null)
            {
                try { excel.Quit(); } catch { }
                ReleaseComObject(excel);
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }

    private static ComparisonStats PerformComparisonWithData(
        object[,] originalValues,
        int originalRows,
        int originalCols,
        object[,] filledValues,
        int filledRows,
        int filledCols,
        TimesheetStructure originalStructure,
        TimesheetStructure filledStructure,
        dynamic outputSheet,
        List<string> protocol)
    {
        if (originalStructure.DayColumns.Count == 0 || filledStructure.DayColumns.Count == 0)
            throw new Exception("В табеле не найдены колонки дней месяца.");

        var stats = new ComparisonStats();
        var originalEmployees = ReadEmployees(originalStructure);
        var filledEmployees = ReadEmployees(filledStructure);
        var filledIndex = BuildEmployeeIndex(filledEmployees);
        int totalComparisons = Math.Max(1, originalStructure.DayColumns.Count * Math.Max(1, originalEmployees.Count));
        int currentComparison = 0;

        protocol.Add("ФИО;Дата/день;Табель;СКУД;Разница;Статус");

        foreach (var originalEmployee in originalEmployees)
        {
            var match = MatchFilledEmployee(originalEmployee, filledStructure, filledIndex);

            if (match.Status == MatchStatus.NotFound)
            {
                protocol.Add($"{originalEmployee.OriginalName};;;;;Строка не найдена в заполненном табеле");
                stats.ReadErrors++;
                currentComparison += originalStructure.DayColumns.Count;
                continue;
            }

            if (match.Status == MatchStatus.Ambiguous)
            {
                protocol.Add($"{originalEmployee.OriginalName};;;;;Неоднозначное сопоставление строк: {string.Join(", ", match.Candidates.Select(x => x.OriginalName))}");
                stats.ReadErrors++;
                currentComparison += originalStructure.DayColumns.Count;
                continue;
            }

            var filledEmployee = match.Employee!;
            stats.CheckedEmployees++;

            foreach (var pair in originalStructure.DayColumns)
            {
                int day = pair.Key;
                int originalCol = pair.Value;

                if (!filledStructure.DayColumns.TryGetValue(day, out int filledCol))
                {
                    protocol.Add($"{originalEmployee.OriginalName};{day};;;;Дневная колонка не найдена в заполненном табеле");
                    stats.ReadErrors++;
                    continue;
                }

                object originalValue = GetArrayValue(originalValues, originalEmployee.RowIndex, originalCol, originalRows, originalCols);
                object filledValue = GetArrayValue(filledValues, filledEmployee.RowIndex, filledCol, filledRows, filledCols);

                var originalCell = ReadHourCell(originalValue);
                var filledCell = ReadHourCell(filledValue);

                if (originalCell.IsEmpty && filledCell.IsEmpty)
                {
                }
                else if (originalCell.IsNumber && filledCell.IsNumber)
                {
                    stats.CheckedCells++;
                    CompareNumericCells(
                        outputSheet,
                        filledEmployee.RowIndex,
                        filledCol,
                        originalEmployee.OriginalName,
                        day,
                        originalCell.Value,
                        filledCell.Value,
                        protocol,
                        stats);
                }
                else if ((originalCell.IsNumber && filledCell.IsEmpty) || (originalCell.IsEmpty && filledCell.IsNumber))
                {
                    if ((originalCell.IsEmpty && filledCell.IsZero) || (filledCell.IsEmpty && originalCell.IsZero))
                    {
                        protocol.Add($"{originalEmployee.OriginalName};{day};Табель={FormatDisplay(originalCell)};СКУД={FormatDisplay(filledCell)};;Нет расхождения: пусто и 0 трактуются как отсутствие часов.");
                        continue;
                    }

                    stats.CheckedCells++;
                    stats.SignificantDifferences++;
                    double originalHours = originalCell.IsNumber ? originalCell.Value : 0;
                    double filledHours = filledCell.IsNumber ? filledCell.Value : 0;
                    double difference = Math.Abs(filledHours - originalHours);
                    HighlightCellYellow(outputSheet.Cells[filledEmployee.RowIndex, filledCol]);
                    string comment = $"Разница: {FormatNumber(difference)} час. Табель: {FormatDisplay(originalCell)}. СКУД: {FormatDisplay(filledCell)}.";
                    AddCommentSafe(outputSheet.Cells[filledEmployee.RowIndex, filledCol], comment);
                    protocol.Add($"{originalEmployee.OriginalName};{day};Табель={FormatDisplay(originalCell)};СКУД={FormatDisplay(filledCell)};Разница={FormatNumber(difference)};Существенное расхождение");
                }
                else
                {
                    stats.ReadErrors++;
                    protocol.Add($"{originalEmployee.OriginalName};{day};Табель={originalCell.Text};СКУД={filledCell.Text};;Ошибка чтения / ручная проверка");
                }

                currentComparison++;
                if (currentComparison % 50 == 0)
                {
                    int percent = 40 + (int)((double)currentComparison / totalComparisons * 50);
                    Report("Сравнение табелей", $"Сравнено ячеек: {currentComparison} из {totalComparisons}", currentComparison, totalComparisons, percent);
                }
            }
        }

        return stats;
    }

    private static void CompareNumericCells(
        dynamic outputSheet,
        int outputRow,
        int outputCol,
        string fio,
        int day,
        double originalHours,
        double filledHours,
        List<string> protocol,
        ComparisonStats stats)
    {
        double difference = Math.Abs(filledHours - originalHours);

        if (difference == 0)
            return;

        string comment = $"Разница: {FormatNumber(difference)} час. Табель: {FormatNumber(originalHours)}. СКУД: {FormatNumber(filledHours)}.";

        if (difference <= 1)
        {
            stats.AllowedDifferences++;
            HighlightCellGreen(outputSheet.Cells[outputRow, outputCol]);
            AddCommentSafe(outputSheet.Cells[outputRow, outputCol], comment);
            protocol.Add($"{fio};{day};Табель={originalHours};СКУД={filledHours};Разница={difference};Допустимое расхождение");
        }
        else
        {
            stats.SignificantDifferences++;
            HighlightCellYellow(outputSheet.Cells[outputRow, outputCol]);
            AddCommentSafe(outputSheet.Cells[outputRow, outputCol], comment);
            protocol.Add($"{fio};{day};Табель={originalHours};СКУД={filledHours};Разница={difference};Существенное расхождение");
        }
    }

    private static dynamic FindWorksheetWithTimesheet(dynamic workbook, out TimesheetStructure structure)
    {
        int sheetsCount = workbook.Worksheets.Count;

        for (int i = 1; i <= sheetsCount; i++)
        {
            dynamic sheet = workbook.Worksheets[i];

            try
            {
                dynamic usedRange = sheet.UsedRange;
                object[,] values = ToArray2D(usedRange.Value2);
                int usedRows = usedRange.Rows.Count;
                int usedCols = usedRange.Columns.Count;

                if (TryDetectTimesheetStructure(values, usedRows, usedCols, out structure))
                    return sheet;
            }
            catch { }
        }

        dynamic fallbackSheet = workbook.Worksheets[1];
        dynamic fallbackRange = fallbackSheet.UsedRange;
        object[,] fallbackValues = ToArray2D(fallbackRange.Value2);
        structure = BuildFallbackStructure(fallbackValues, fallbackRange.Rows.Count, fallbackRange.Columns.Count);
        return fallbackSheet;
    }

    private static bool TryDetectTimesheetStructure(object[,] values, int usedRows, int usedCols, out TimesheetStructure structure)
    {
        int fioCol = FindFioColumn(values, usedRows, usedCols, out int headerRow);

        if (fioCol < 0)
        {
            structure = default!;
            return false;
        }

        int dateRow = -1;
        Dictionary<int, int> dayColumns = new();

        for (int row = 1; row <= Math.Min(15, usedRows); row++)
        {
            var current = ReadDayColumns(values, row, usedCols);
            if (current.Count > dayColumns.Count)
            {
                dayColumns = current;
                dateRow = row;
            }
        }

        if (dayColumns.Count < 5)
        {
            structure = BuildFallbackStructure(values, usedRows, usedCols);

            if (structure.FioCol <= 0 || structure.DayColumns.Count == 0)
                return false;

            return true;
        }

        structure = new TimesheetStructure(
            headerRow,
            dateRow,
            Math.Max(headerRow, dateRow) + 1,
            fioCol,
            dayColumns,
            FindTotalColumn(values, Math.Max(headerRow, dateRow), usedCols),
            values,
            usedRows,
            usedCols);

        return true;
    }

    private static TimesheetStructure BuildFallbackStructure(object[,] values, int usedRows, int usedCols)
    {
        var dayColumns = new Dictionary<int, int>();

        if (usedCols >= 4)
        {
            for (int col = 4; col <= Math.Min(31, usedCols); col++)
                dayColumns[col - 3] = col;
        }

        return new TimesheetStructure(
            HeaderRow: 4,
            DateRow: 4,
            FirstEmployeeRow: 5,
            FioCol: usedCols >= 2 ? 2 : 1,
            DayColumns: dayColumns,
            TotalHoursCol: usedCols >= 32 ? 32 : null,
            Values: values,
            UsedRows: usedRows,
            UsedCols: usedCols);
    }

    private static Dictionary<int, int> ReadDayColumns(object[,] values, int row, int usedCols)
    {
        var result = new Dictionary<int, int>();

        for (int col = 1; col <= usedCols; col++)
        {
            if (TryReadDay(values[row, col], out int day))
                result[day] = col;
        }

        return result;
    }

    private static int? FindTotalColumn(object[,] values, int row, int usedCols)
    {
        if (row > 0)
        {
            for (int col = 1; col <= usedCols; col++)
            {
                string header = NormalizeHeader(GetValueText(values[row, col]));
                if (header.Contains("итог"))
                    return col;
            }
        }

        return usedCols >= 32 ? 32 : null;
    }

    private static List<TimesheetEmployee> ReadEmployees(TimesheetStructure structure)
    {
        var result = new List<TimesheetEmployee>();

        for (int row = structure.FirstEmployeeRow; row <= structure.UsedRows; row++)
        {
            string fio = GetValueText(GetArrayValue(structure.Values, row, structure.FioCol, structure.UsedRows, structure.UsedCols));
            if (!LooksLikePersonName(fio))
                continue;

            string normalized = NormalizeName(fio);
            result.Add(new TimesheetEmployee(row, fio, normalized, BuildShortNameKey(normalized)));
        }

        return result;
    }

    private static EmployeeIndex BuildEmployeeIndex(List<TimesheetEmployee> employees)
    {
        var index = new EmployeeIndex();

        foreach (var employee in employees)
        {
            AddToIndex(index.ByFullName, employee.NormalizedName, employee);

            if (!string.IsNullOrWhiteSpace(employee.ShortKey))
                AddToIndex(index.ByShortKey, employee.ShortKey, employee);
        }

        return index;
    }

    private static void AddToIndex(Dictionary<string, List<TimesheetEmployee>> index, string key, TimesheetEmployee employee)
    {
        if (!index.TryGetValue(key, out var list))
        {
            list = new List<TimesheetEmployee>();
            index[key] = list;
        }

        list.Add(employee);
    }

    private static RowMatch MatchFilledEmployee(TimesheetEmployee originalEmployee, TimesheetStructure filledStructure, EmployeeIndex filledIndex)
    {
        string sameRowFio = GetValueText(GetArrayValue(filledStructure.Values, originalEmployee.RowIndex, filledStructure.FioCol, filledStructure.UsedRows, filledStructure.UsedCols));
        string normalizedSameRowFio = NormalizeName(sameRowFio);

        if (!string.IsNullOrWhiteSpace(normalizedSameRowFio) && normalizedSameRowFio == originalEmployee.NormalizedName)
        {
            return RowMatch.Found(new TimesheetEmployee(
                originalEmployee.RowIndex,
                sameRowFio,
                normalizedSameRowFio,
                BuildShortNameKey(normalizedSameRowFio)));
        }

        if (filledIndex.ByFullName.TryGetValue(originalEmployee.NormalizedName, out var fullMatches))
            return BuildRowMatch(fullMatches);

        if (!string.IsNullOrWhiteSpace(originalEmployee.ShortKey)
            && filledIndex.ByShortKey.TryGetValue(originalEmployee.ShortKey, out var shortMatches))
            return BuildRowMatch(shortMatches);

        return RowMatch.NotFound();
    }

    private static RowMatch BuildRowMatch(List<TimesheetEmployee> matches)
    {
        var unique = matches
            .GroupBy(x => x.RowIndex)
            .Select(x => x.First())
            .ToList();

        if (unique.Count == 1)
            return RowMatch.Found(unique[0]);

        return RowMatch.Ambiguous(unique);
    }

    private static HourCell ReadHourCell(object value)
    {
        if (IsEmpty(value))
            return HourCell.Empty();

        if (TryConvertToDouble(value, out double number))
            return HourCell.Number(number, GetValueText(value));

        return HourCell.TextValue(GetValueText(value));
    }

    private static string FormatDisplay(HourCell cell)
    {
        if (cell.IsEmpty)
            return "пусто";

        if (cell.IsNumber)
            return FormatNumber(cell.Value);

        return cell.Text;
    }

    private static bool IsZeroValue(double value)
    {
        return Math.Abs(value) < 0.000001;
    }

    private static string BuildShortNameKey(string fio)
    {
        string normalized = NormalizeName(fio);
        var parts = normalized.Split(' ', StringSplitOptions.RemoveEmptyEntries);

        if (parts.Length < 2)
            return string.Empty;

        return parts[0] + " " + parts[1][0];
    }

    private static int FindFioColumn(object[,] values, int usedRows, int usedCols, out int headerRow)
    {
        headerRow = -1;

        for (int row = 1; row <= Math.Min(50, usedRows); row++)
        {
            for (int col = 1; col <= usedCols; col++)
            {
                string header = NormalizeHeader(GetValueText(values[row, col]));

                if (header == "фио" || header.Contains("фамилия"))
                {
                    headerRow = row;
                    return col;
                }
            }
        }

        return -1;
    }

    private static bool TryReadDay(object value, out int day)
    {
        day = 0;

        if (value == null)
            return false;

        if (value is double d)
        {
            if (d >= 1 && d <= 31 && Math.Abs(d - Math.Round(d)) < 0.001)
            {
                day = (int)Math.Round(d);
                return true;
            }

            try
            {
                var date = DateTime.FromOADate(d);
                if (date.Year > 1900)
                {
                    day = date.Day;
                    return true;
                }
            }
            catch { }
        }

        string text = GetValueText(value);
        if (int.TryParse(text, out int parsed) && parsed >= 1 && parsed <= 31)
        {
            day = parsed;
            return true;
        }

        if (DateTime.TryParse(text, CultureInfo.GetCultureInfo("ru-RU"), DateTimeStyles.None, out DateTime dt))
        {
            day = dt.Day;
            return true;
        }

        return false;
    }
    private static bool TryConvertToDouble(object value, out double result)
    {
        result = 0;

        if (value == null)
            return false;

        if (value is double d)
        {
            result = d;
            return true;
        }

        string text = Convert.ToString(value)?.Trim() ?? "";
        if (string.IsNullOrWhiteSpace(text))
            return false;

        return double.TryParse(text, NumberStyles.Any, CultureInfo.CurrentCulture, out result)
            || double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out result);
    }

    private static object GetArrayValue(object[,] values, int row, int col, int rows, int cols)
    {
        if (row <= 0 || col <= 0 || row > rows || col > cols)
            return null!;

        return values[row, col];
    }

    private static bool IsEmpty(object value)
    {
        return value == null || string.IsNullOrWhiteSpace(Convert.ToString(value));
    }

    private static bool LooksLikePersonName(string value)
    {
        string normalized = NormalizeName(value);

        if (string.IsNullOrWhiteSpace(normalized))
            return false;

        if (normalized.Contains("итого") || normalized.Contains("всего") || normalized.Contains("табель") || normalized.Contains("фио"))
            return false;

        return normalized.Split(' ', StringSplitOptions.RemoveEmptyEntries).Length >= 2;
    }

    private static string NormalizeName(string text)
    {
        if (text == null)
            return string.Empty;

        text = text.Replace('\u00A0', ' ').ToLowerInvariant().Replace('ё', 'е');
        text = Regex.Replace(text, @"[^\p{L}\s]", " ");
        text = Regex.Replace(text, @"\s+", " ");
        return text.Trim();
    }

    private static string NormalizeHeader(string text)
    {
        if (text == null)
            return string.Empty;

        text = text.Replace('\u00A0', ' ').ToLowerInvariant().Replace('ё', 'е');
        text = Regex.Replace(text, @"\s+", " ");
        return text.Trim();
    }

    private static object[,] ToArray2D(object value)
    {
        if (value is object[,] array)
            return array;

        var result = new object[2, 2];
        result[1, 1] = value;
        return result;
    }

    private static string GetValueText(object value)
    {
        try
        {
            return value == null ? string.Empty : Convert.ToString(value)?.Trim() ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    private static string FormatNumber(double value)
    {
        return value.ToString("0.##", CultureInfo.GetCultureInfo("ru-RU"));
    }

    private static void HighlightCellYellow(dynamic cell)
    {
        try { cell.Interior.Color = RGB(255, 255, 0); } catch { }
    }

    private static void HighlightCellGreen(dynamic cell)
    {
        try { cell.Interior.Color = RGB(198, 239, 206); } catch { }
    }

    private static void AddCommentSafe(dynamic cell, string commentText)
    {
        try
        {
            try { var old = cell.Comment; if (old != null) old.Delete(); } catch { }
            cell.AddComment(commentText);
            try { cell.Comment.Shape.TextFrame.AutoSize = true; } catch { }
        }
        catch { }
    }

    private static void WriteProtocol(string path, List<string> protocol)
    {
        var sb = new StringBuilder();
        sb.AppendLine("# Протокол сравнения");

        if (protocol.Count == 0)
            sb.AppendLine("Нет записей.");
        else
            foreach (string line in protocol)
                sb.AppendLine(line);

        File.WriteAllText(path, sb.ToString(), Encoding.UTF8);
    }

    private static Task<T> RunStaTask<T>(Func<T> action)
    {
        var tcs = new TaskCompletionSource<T>();
        var thread = new Thread(() =>
        {
            try { tcs.SetResult(action()); }
            catch (Exception ex) { tcs.SetException(ex); }
        });

        thread.IsBackground = true;
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        return tcs.Task;
    }

    private static int RGB(int red, int green, int blue)
    {
        return red + (green << 8) + (blue << 16);
    }

    private static void Report(string stage, string message, int current, int total, int percent, bool isFinished = false)
    {
        ProgressChanged?.Invoke(null, new SkudProgressEventArgs(stage, message, current, total, percent, isFinished));
    }

    private static void ReleaseComObject(object obj)
    {
        try { Marshal.FinalReleaseComObject(obj); } catch { }
    }

    private sealed record TimesheetStructure(
        int HeaderRow,
        int DateRow,
        int FirstEmployeeRow,
        int FioCol,
        Dictionary<int, int> DayColumns,
        int? TotalHoursCol,
        object[,] Values,
        int UsedRows,
        int UsedCols);

    private sealed record TimesheetEmployee(
        int RowIndex,
        string OriginalName,
        string NormalizedName,
        string ShortKey);

    private sealed class EmployeeIndex
    {
        public Dictionary<string, List<TimesheetEmployee>> ByFullName { get; } = new(StringComparer.OrdinalIgnoreCase);
        public Dictionary<string, List<TimesheetEmployee>> ByShortKey { get; } = new(StringComparer.OrdinalIgnoreCase);
    }

    private enum MatchStatus
    {
        Found,
        NotFound,
        Ambiguous
    }

    private sealed class RowMatch
    {
        public MatchStatus Status { get; private init; }
        public TimesheetEmployee? Employee { get; private init; }
        public List<TimesheetEmployee> Candidates { get; private init; } = new();

        public static RowMatch Found(TimesheetEmployee employee)
        {
            return new RowMatch { Status = MatchStatus.Found, Employee = employee };
        }

        public static RowMatch NotFound()
        {
            return new RowMatch { Status = MatchStatus.NotFound };
        }

        public static RowMatch Ambiguous(List<TimesheetEmployee> candidates)
        {
            return new RowMatch { Status = MatchStatus.Ambiguous, Candidates = candidates };
        }
    }

    private sealed class HourCell
    {
        public bool IsEmpty { get; private init; }
        public bool IsNumber { get; private init; }
        public bool IsZero => IsNumber && IsZeroValue(Value);
        public double Value { get; private init; }
        public string Text { get; private init; } = string.Empty;

        public static HourCell Empty()
        {
            return new HourCell { IsEmpty = true, Text = string.Empty };
        }

        public static HourCell Number(double value, string text)
        {
            return new HourCell { IsNumber = true, Value = value, Text = text };
        }

        public static HourCell TextValue(string text)
        {
            return new HourCell { Text = text };
        }
    }

    private sealed class ComparisonStats
    {
        public int CheckedEmployees { get; set; }
        public int CheckedCells { get; set; }
        public int SignificantDifferences { get; set; }
        public int AllowedDifferences { get; set; }
        public int ReadErrors { get; set; }
    }
}
