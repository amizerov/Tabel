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
        dynamic? outputBook = null;
        dynamic? originalSheet = null;
        dynamic? filledSheet = null;

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
            originalSheet = FindWorksheetWithTimesheet(originalBook, out TimesheetStructure originalStructure);
            object[,] originalValues = originalStructure.Values;
            int originalRows = originalStructure.UsedRows;
            int originalCols = originalStructure.UsedCols;

            ReleaseComObject(originalSheet);
            originalSheet = null;
            originalBook.Close(false);
            ReleaseComObject(originalBook);
            originalBook = null;

            outputBook = excel.Workbooks.Open(outputPath);
            filledSheet = FindWorksheetWithTimesheet(outputBook, out TimesheetStructure filledStructure);
            object[,] filledValues = filledStructure.Values;
            int filledRows = filledStructure.UsedRows;
            int filledCols = filledStructure.UsedCols;

            AlignDateColumns(ref originalStructure, ref filledStructure);
            EnsureTotalColumnFormatting(filledSheet, filledStructure, protocol);

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
                filledSheet,
                protocol);

            Report("Сохранение", "Сохраняю файл сравнения...", 0, 100, 95);
            TabelWorkbookMetadata.SetRole(outputBook, TabelWorkbookMetadata.ComparisonRole);
            outputBook.Save();
            ReleaseComObject(filledSheet);
            filledSheet = null;
            outputBook.Close(false);
            ReleaseComObject(outputBook);
            outputBook = null;

            WriteProtocol(protocolPath, protocol);

            string summary =
                $"Сравнение завершено. Проверено сотрудников: {stats.CheckedEmployees}. " +
                $"Проверено ячеек: {stats.CheckedCells}. СКУД меньше заявленного: {stats.SkudLowerDifferences}. " +
                $"СКУД больше заявленного: {stats.SkudHigherDifferences}. Специальных статусов: {stats.SpecialStatuses}. " +
                $"Ошибок чтения: {stats.ReadErrors}. Файл: {outputPath}";

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
            if (filledSheet != null)
                ReleaseComObject(filledSheet);
            if (originalSheet != null)
                ReleaseComObject(originalSheet);

            try { if (outputBook != null) outputBook.Close(false); } catch { }
            try { if (originalBook != null) originalBook.Close(false); } catch { }

            if (outputBook != null)
                ReleaseComObject(outputBook);

            if (originalBook != null)
                ReleaseComObject(originalBook);

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

        protocol.Add("ФИО;Дата/день;Заявлено;СКУД;Итоговое значение;Разница;Статус");

        foreach (var originalEmployee in originalEmployees)
        {
            var match = MatchFilledEmployee(originalEmployee, filledStructure, filledIndex);

            if (match.Status == MatchStatus.NotFound)
            {
                protocol.Add($"{originalEmployee.OriginalName};;;;;;Сотрудник не найден в заполненном табеле");
                stats.ReadErrors++;
                currentComparison += originalStructure.DayColumns.Count;
                continue;
            }

            if (match.Status == MatchStatus.Ambiguous)
            {
                string variants = string.Join(", ", match.Candidates.Select(x => x.OriginalName));
                try
                {
                    dynamic fioCell = outputSheet.Cells[originalEmployee.RowIndex, filledStructure.FioCol];
                    HighlightCellOrange(fioCell);
                    AddOrAppendComparisonComment(
                        fioCell,
                        $"Неоднозначное сопоставление ФИО. Ключ поиска: {originalEmployee.ShortKey}. Найденные варианты: {variants}.");
                }
                catch { }
                protocol.Add($"{originalEmployee.OriginalName};;;;;;Неоднозначное сопоставление ФИО. Ключ поиска: {originalEmployee.ShortKey}. Найденные варианты: {variants}");
                stats.ReadErrors++;
                currentComparison += originalStructure.DayColumns.Count;
                continue;
            }

            var filledEmployee = match.Employee!;
            stats.CheckedEmployees++;

            string fioComment = GetCellCommentText(outputSheet.Cells[filledEmployee.RowIndex, filledStructure.FioCol]);
            if (fioComment.Contains("Неоднозначное сопоставление ФИО", StringComparison.OrdinalIgnoreCase))
            {
                try { HighlightCellOrange(outputSheet.Cells[filledEmployee.RowIndex, filledStructure.FioCol]); } catch { }
                protocol.Add($"{originalEmployee.OriginalName};;;;;;{FlattenForProtocol(fioComment)}");
                stats.SpecialStatuses++;
                currentComparison += originalStructure.DayColumns.Count;

                RecalculateOutputTotal(
                    outputSheet,
                    filledEmployee.RowIndex,
                    filledStructure,
                    originalEmployee.OriginalName,
                    protocol,
                    stats);

                continue;
            }

            foreach (var pair in originalStructure.DayColumns)
            {
                DateTime date = pair.Key;
                int originalCol = pair.Value;

                if (!filledStructure.DayColumns.TryGetValue(date, out int filledCol))
                {
                    protocol.Add($"{originalEmployee.OriginalName};{FormatDateKey(date)};;;;;Дневная колонка не найдена в заполненном табеле");
                    stats.ReadErrors++;
                    continue;
                }

                object originalValue = GetArrayValue(originalValues, originalEmployee.RowIndex, originalCol, originalRows, originalCols);
                object filledValue = GetArrayValue(filledValues, filledEmployee.RowIndex, filledCol, filledRows, filledCols);

                var originalCell = ReadHourCell(originalValue);
                var filledCell = ReadHourCell(filledValue);
                dynamic outputCell = outputSheet.Cells[filledEmployee.RowIndex, filledCol];
                string outputComment = GetCellCommentText(outputCell);
                CellSpecialStatus specialStatus = DetectSpecialStatus(outputCell, outputComment);

                if (originalCell.IsNumber && specialStatus is CellSpecialStatus.NoAllowedEvents
                    or CellSpecialStatus.EmployeeNotFound
                    or CellSpecialStatus.ManualReview)
                {
                    stats.CheckedCells++;
                    stats.SpecialStatuses++;
                    HandleMissingSkudData(
                        outputCell,
                        originalEmployee.OriginalName,
                        date,
                        originalCell.Value,
                        specialStatus,
                        outputComment,
                        protocol);
                }

                else if (originalCell.IsEmpty && filledCell.IsEmpty)
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
                        date,
                        originalCell.Value,
                        filledCell.Value,
                        specialStatus,
                        protocol,
                        stats);
                }
                else if ((originalCell.IsNumber && filledCell.IsEmpty) || (originalCell.IsEmpty && filledCell.IsNumber))
                {
                    stats.CheckedCells++;
                    double originalHours = originalCell.IsNumber ? originalCell.Value : 0;
                    double filledHours = filledCell.IsNumber ? filledCell.Value : 0;

                    CompareNumericCells(
                        outputSheet,
                        filledEmployee.RowIndex,
                        filledCol,
                        originalEmployee.OriginalName,
                        date,
                        originalHours,
                        filledHours,
                        specialStatus,
                        protocol,
                        stats);
                }
                else
                {
                    stats.ReadErrors++;
                    protocol.Add($"{originalEmployee.OriginalName};{FormatDateKey(date)};{originalCell.Text};{filledCell.Text};;;Ошибка чтения / ручная проверка");
                }

                currentComparison++;
                if (currentComparison % 50 == 0)
                {
                    int percent = 40 + (int)((double)currentComparison / totalComparisons * 50);
                    Report("Сравнение табелей", $"Сравнено ячеек: {currentComparison} из {totalComparisons}", currentComparison, totalComparisons, percent);
                }
            }

            RecalculateOutputTotal(
                outputSheet,
                filledEmployee.RowIndex,
                filledStructure,
                originalEmployee.OriginalName,
                protocol,
                stats);
        }

        return stats;
    }

    private static void CompareNumericCells(
        dynamic outputSheet,
        int outputRow,
        int outputCol,
        string fio,
        DateTime date,
        double originalHours,
        double filledHours,
        CellSpecialStatus specialStatus,
        List<string> protocol,
        ComparisonStats stats)
    {
        dynamic cell = outputSheet.Cells[outputRow, outputCol];
        double difference = filledHours - originalHours;
        string declared = FormatNumber(originalHours);
        string skud = FormatNumber(filledHours);
        bool isSingleAllowedEvent = specialStatus == CellSpecialStatus.SingleAllowedEvent;

        if (isSingleAllowedEvent)
        {
            stats.SpecialStatuses++;
            try { HighlightCellBlue(cell); } catch { }
            string singleComment = Math.Abs(filledHours - 2) < 1e-9
                ? "За день найдена только одна разрешённая отметка СКУД. Установлено 2 часа."
                : $"За день найдена только одна разрешённая отметка СКУД. Расчёт выполнен по правилу для одной отметки. По СКУД: {skud} ч.";
            try { AddOrAppendComparisonComment(cell, singleComment); } catch { }
        }

        if (Math.Abs(difference) < 1e-9)
        {
            if (!isSingleAllowedEvent)
                try { ClearCellFill(cell); } catch { }

            protocol.Add($"{fio};{FormatDateKey(date)};{declared};{skud};{skud};0;" +
                (isSingleAllowedEvent ? "Одна разрешённая отметка СКУД" : "Совпадает"));
            return;
        }

        if (difference > 0)
        {
            stats.SkudHigherDifferences++;
            cell.Value2 = originalHours;
            if (!isSingleAllowedEvent)
                try { ClearCellFill(cell); } catch { }

            string comment = $"Заявлено: {declared} ч. По СКУД: {skud} ч. Фактическое время по СКУД превышает заявленное. " +
                $"В табеле оставлено заявленное значение: {declared} ч.";
            try { AddOrAppendComparisonComment(cell, comment); } catch { }
            protocol.Add($"{fio};{FormatDateKey(date)};{declared};{skud};{declared};{FormatNumber(difference)};" +
                (isSingleAllowedEvent ? "Одна разрешённая отметка СКУД" : "СКУД больше заявленного — оставлены заявленные часы"));
            return;
        }

        stats.SkudLowerDifferences++;
        if (!isSingleAllowedEvent)
            try { HighlightCellYellow(cell); } catch { }

        double shortage = originalHours - filledHours;
        string lowerComment = $"Заявлено: {declared} ч. По СКУД: {skud} ч. Фактическое время по СКУД меньше заявленного на {FormatNumber(shortage)} ч.";
        try { AddOrAppendComparisonComment(cell, lowerComment); } catch { }
        protocol.Add($"{fio};{FormatDateKey(date)};{declared};{skud};{skud};{FormatNumber(shortage)};" +
            (isSingleAllowedEvent ? "Одна разрешённая отметка СКУД" : "СКУД меньше заявленного"));
    }

    private static void RecalculateOutputTotal(
        dynamic sheet,
        int row,
        TimesheetStructure structure,
        string fio,
        List<string> protocol,
        ComparisonStats stats)
    {
        if (!structure.TotalHoursCol.HasValue)
            return;

        try
        {
            int totalColumn = structure.TotalHoursCol.Value;
            if (structure.DayColumns.Values.Contains(totalColumn))
            {
                stats.ReadErrors++;
                protocol.Add($"{fio};;;;;;Итоговая колонка совпала с дневной колонкой. Пересчёт итога пропущен");
                return;
            }

            dynamic totalCell = sheet.Cells[row, totalColumn];
            string totalComment = GetCellCommentText(totalCell);
            if (totalComment.Contains("Итог не рассчитан", StringComparison.OrdinalIgnoreCase))
            {
                try { totalCell.ClearContents(); } catch { try { totalCell.Value2 = null; } catch { } }
                try { HighlightCellRed(totalCell); } catch { }
                protocol.Add($"{fio};;;;;;{FlattenForProtocol(totalComment)}");
                stats.SpecialStatuses++;
                return;
            }

            double total = 0;
            foreach (int col in structure.DayColumns.Values.Distinct())
            {
                object value = sheet.Cells[row, col].Value2;
                if (TryConvertToDouble(value, out double hours))
                    total += hours;
            }

            try
            {
                if (totalCell.HasFormula == true)
                {
                    totalCell.Calculate();
                    return;
                }
            }
            catch { }

            totalCell.Value2 = total;
        }
        catch (Exception ex)
        {
            stats.ReadErrors++;
            protocol.Add($"{fio};;;;;;Не удалось пересчитать итоговую колонку: {FlattenForProtocol(ex.Message)}");
        }
    }

    private static void EnsureTotalColumnFormatting(
        dynamic sheet,
        TimesheetStructure structure,
        List<string> protocol)
    {
        if (!structure.TotalHoursCol.HasValue || structure.DayColumns.Count == 0)
            return;

        int totalColumn = structure.TotalHoursCol.Value;
        int lastDayColumn = structure.DayColumns.Values.Max();

        if (structure.DayColumns.Values.Contains(totalColumn))
            return;

        try
        {
            int headerRow = structure.DateRow > 0 ? structure.DateRow : structure.HeaderRow;
            dynamic sourceHeader = sheet.Cells[headerRow, lastDayColumn];
            dynamic totalHeader = sheet.Cells[headerRow, totalColumn];

            try
            {
                sourceHeader.Copy();
                totalHeader.PasteSpecial(-4122); // xlPasteFormats
            }
            catch { }

            if (string.IsNullOrWhiteSpace(GetValueText(totalHeader.Value2)))
                totalHeader.Value2 = "Итого часов";

            try { totalHeader.Font.Bold = true; } catch { }
            try { totalHeader.HorizontalAlignment = -4108; } catch { } // xlCenter
            try { totalHeader.VerticalAlignment = -4108; } catch { }
            try { totalHeader.WrapText = true; } catch { }

            for (int row = structure.FirstEmployeeRow; row <= structure.UsedRows; row++)
            {
                dynamic sourceCell = sheet.Cells[row, lastDayColumn];
                dynamic totalCell = sheet.Cells[row, totalColumn];

                try
                {
                    sourceCell.Copy();
                    totalCell.PasteSpecial(-4122); // xlPasteFormats
                }
                catch { }

                string totalComment = GetCellCommentText(totalCell);
                if (totalComment.Contains("Итог не рассчитан", StringComparison.OrdinalIgnoreCase))
                {
                    try { HighlightCellRed(totalCell); } catch { }
                }
                else
                {
                    try { totalCell.Interior.Pattern = -4142; } catch { } // xlNone
                }
                try { totalCell.NumberFormat = "0"; } catch { }
                try { totalCell.HorizontalAlignment = -4108; } catch { }
                try { totalCell.VerticalAlignment = -4108; } catch { }
            }

            for (int row = headerRow; row <= structure.UsedRows; row++)
            {
                dynamic lastDayCell = sheet.Cells[row, lastDayColumn];
                dynamic totalCell = sheet.Cells[row, totalColumn];

                try { lastDayCell.Borders[10].LineStyle = 1; } catch { } // xlEdgeRight, xlContinuous
                try { lastDayCell.Borders[10].Weight = 2; } catch { } // xlThin
                try { totalCell.Borders[7].LineStyle = 1; } catch { } // xlEdgeLeft
                try { totalCell.Borders[7].Weight = 2; } catch { }
                try { totalCell.Borders[10].LineStyle = 1; } catch { }
                try { totalCell.Borders[10].Weight = -4138; } catch { } // xlMedium
            }

            try
            {
                double lastDayWidth = Convert.ToDouble(
                    sheet.Columns[lastDayColumn].ColumnWidth,
                    CultureInfo.InvariantCulture);
                sheet.Columns[totalColumn].ColumnWidth = Math.Max(7d, lastDayWidth);
            }
            catch { }

            try { sheet.Application.CutCopyMode = false; } catch { }
        }
        catch (Exception ex)
        {
            protocol.Add($";;;;;;Не удалось оформить итоговую колонку: {FlattenForProtocol(ex.Message)}");
        }
    }

    private static dynamic FindWorksheetWithTimesheet(dynamic workbook, out TimesheetStructure structure)
    {
        int sheetsCount = workbook.Worksheets.Count;
        Exception? lastStructureError = null;

        for (int i = 1; i <= sheetsCount; i++)
        {
            dynamic sheet = workbook.Worksheets[i];
            dynamic? usedRange = null;

            try
            {
                ExcelWorksheetData data = ExcelWorksheetReader.Read(sheet);
                object[,] values = data.Values;
                int usedRows = data.Rows;
                int usedCols = data.Columns;

                if (TryDetectTimesheetStructure(values, usedRows, usedCols, out structure))
                    return sheet;
            }
            catch (Exception ex)
            {
                lastStructureError = ex;
            }
            finally
            {
                if (usedRange != null)
                    ReleaseComObject(usedRange);
            }
        }

        throw lastStructureError ?? new Exception("Не удалось определить структуру табеля для сравнения. Не найдены колонка ФИО и строка с датами.");
    }

    private static bool TryDetectTimesheetStructure(object[,] values, int usedRows, int usedCols, out TimesheetStructure structure)
    {
        int fioCol = FindFioColumn(values, usedRows, usedCols, out int headerRow);

        if (fioCol < 0)
            throw new Exception("В табеле не найдена колонка ФИО.");

        int dateRow = -1;
        Dictionary<DateTime, int> dayColumns = new();

        foreach (int row in GetDateRowSearchOrder(headerRow, usedRows))
        {
            var current = ReadDayColumns(values, row, usedCols);
            if (current.Count > dayColumns.Count)
            {
                dayColumns = current;
                dateRow = row;
            }
        }

        if (dayColumns.Count < 5)
            throw new Exception("В табеле не найдена строка с датами или номерами дней месяца.");

        structure = new TimesheetStructure(
            headerRow,
            dateRow,
            Math.Max(headerRow, dateRow) + 1,
            fioCol,
            dayColumns,
            FindTotalColumn(values, Math.Max(headerRow, dateRow), usedCols, dayColumns),
            values,
            usedRows,
            usedCols,
            dayColumns.Keys.All(IsSyntheticDateKey));

        return true;
    }

    private static void AlignDateColumns(ref TimesheetStructure originalStructure, ref TimesheetStructure filledStructure)
    {
        if (originalStructure.UsesSyntheticDayNumbers && !filledStructure.UsesSyntheticDayNumbers)
            originalStructure = originalStructure with { DayColumns = RebaseSyntheticDayColumns(originalStructure.DayColumns, filledStructure.DayColumns) };

        if (filledStructure.UsesSyntheticDayNumbers && !originalStructure.UsesSyntheticDayNumbers)
            filledStructure = filledStructure with { DayColumns = RebaseSyntheticDayColumns(filledStructure.DayColumns, originalStructure.DayColumns) };
    }

    private static Dictionary<DateTime, int> RebaseSyntheticDayColumns(Dictionary<DateTime, int> source, Dictionary<DateTime, int> target)
    {
        var targetMonths = target.Keys
            .Where(x => !IsSyntheticDateKey(x))
            .Select(x => new DateTime(x.Year, x.Month, 1))
            .Distinct()
            .ToList();

        if (targetMonths.Count != 1)
            return source;

        DateTime month = targetMonths[0];
        var result = new Dictionary<DateTime, int>();

        foreach (var pair in source)
        {
            int day = pair.Key.Day;
            if (day <= DateTime.DaysInMonth(month.Year, month.Month))
                result[new DateTime(month.Year, month.Month, day)] = pair.Value;
        }

        return result;
    }

    private static bool IsSyntheticDateKey(DateTime date)
    {
        return date.Year == 2000 && date.Month == 1;
    }

    private static Dictionary<DateTime, int> ReadDayColumns(object[,] values, int row, int usedCols)
    {
        var result = new Dictionary<DateTime, int>();
        DateTime? rowMonth = TryFindRowMonth(values, row, usedCols);

        for (int col = 1; col <= usedCols; col++)
        {
            if (TryReadDateColumn(values[row, col], rowMonth, out DateTime date))
                result[date.Date] = col;
        }

        return result;
    }

    private static IEnumerable<int> GetDateRowSearchOrder(int headerRow, int usedRows)
    {
        int[] offsets = { 1, 0, 2, -1, 3, -2 };
        var seen = new HashSet<int>();

        foreach (int offset in offsets)
        {
            int row = headerRow + offset;
            if (row >= 1 && row <= usedRows && seen.Add(row))
                yield return row;
        }
    }

    private static DateTime? TryFindRowMonth(object[,] values, int row, int usedCols)
    {
        for (int col = 1; col <= usedCols; col++)
        {
            if (TryReadFullDate(values[row, col], out DateTime date))
                return new DateTime(date.Year, date.Month, 1);
        }

        return null;
    }

    private static int? FindTotalColumn(
        object[,] values,
        int row,
        int usedCols,
        Dictionary<DateTime, int> dayColumns)
    {
        if (row > 0)
        {
            for (int col = 1; col <= usedCols; col++)
            {
                string header = NormalizeHeader(GetValueText(values[row, col]));
                if ((header.Contains("итог") || header.Contains("всего"))
                    && !dayColumns.Values.Contains(col))
                    return col;
            }
        }

        if (dayColumns.Count == 0)
            return null;

        return dayColumns.Values.Max() + 1;
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
        {
            RowMatch shortMatch = BuildRowMatch(shortMatches);
            if (shortMatch.Status == MatchStatus.Found)
                return shortMatch;
        }

        var fuzzy = FioFuzzyMatcher.Find(
            originalEmployee.NormalizedName,
            filledIndex.ByFullName.Values.SelectMany(x => x).GroupBy(x => x.RowIndex).Select(x => x.First()),
            candidate => candidate.NormalizedName);

        if (fuzzy.Status == FuzzyMatchStatus.Found)
            return RowMatch.Found(fuzzy.Match!);

        if (fuzzy.Status == FuzzyMatchStatus.Ambiguous)
            return RowMatch.Ambiguous(fuzzy.Candidates.ToList());

        if (!string.IsNullOrWhiteSpace(originalEmployee.ShortKey)
            && filledIndex.ByShortKey.TryGetValue(originalEmployee.ShortKey, out shortMatches))
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

                if (IsFioHeader(header))
                {
                    headerRow = row;
                    return col;
                }
            }
        }

        return -1;
    }

    private static bool TryReadDateColumn(object value, DateTime? rowMonth, out DateTime date)
    {
        date = default;

        if (value == null)
            return false;

        if (TryReadFullDate(value, out DateTime fullDate))
        {
            date = fullDate.Date;
            return true;
        }

        if (value is double d)
        {
            if (d >= 1 && d <= 31 && Math.Abs(d - Math.Round(d)) < 0.001)
            {
                int day = (int)Math.Round(d);
                DateTime month = rowMonth ?? new DateTime(2000, 1, 1);
                if (day <= DateTime.DaysInMonth(month.Year, month.Month))
                {
                    date = new DateTime(month.Year, month.Month, day);
                    return true;
                }
            }
        }

        string text = GetValueText(value);
        if (int.TryParse(text, out int parsed) && parsed >= 1 && parsed <= 31)
        {
            DateTime month = rowMonth ?? new DateTime(2000, 1, 1);
            if (parsed <= DateTime.DaysInMonth(month.Year, month.Month))
            {
                date = new DateTime(month.Year, month.Month, parsed);
                return true;
            }
        }

        return false;
    }

    private static bool TryReadFullDate(object value, out DateTime date)
    {
        date = default;

        if (value == null)
            return false;

        if (value is double d)
        {
            if (d >= 1 && d <= 31 && Math.Abs(d - Math.Round(d)) < 0.001)
                return false;

            try
            {
                var excelDate = DateTime.FromOADate(d);
                if (excelDate.Year > 1900)
                {
                    date = excelDate.Date;
                    return true;
                }
            }
            catch { }
        }

        string text = GetValueText(value);
        if (int.TryParse(text, out _))
            return false;

        if (DateTime.TryParse(text, CultureInfo.GetCultureInfo("ru-RU"), DateTimeStyles.None, out DateTime parsed))
        {
            date = parsed.Date;
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

        var parts = normalized.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length < 2)
            return false;

        return !IsServiceRowText(normalized);
    }

    private static bool IsServiceRowText(string normalized)
    {
        return IsFioHeader(normalized)
            || normalized == "итого"
            || normalized == "всего"
            || normalized.StartsWith("итого ")
            || normalized.StartsWith("всего ")
            || normalized.Contains("итого часов")
            || normalized.Contains("сводный табель");
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

    private static bool IsFioHeader(string header)
    {
        string normalized = NormalizeHeader(header);
        string compact = Regex.Replace(normalized, @"[.\s]+", string.Empty);

        return compact == "фио" || normalized.Contains("фамилия", StringComparison.Ordinal);
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

    private static string FormatDateKey(DateTime date)
    {
        return date.Year == 2000 && date.Month == 1
            ? date.Day.ToString(CultureInfo.InvariantCulture)
            : date.ToString("dd.MM.yyyy", CultureInfo.InvariantCulture);
    }

    private static void HighlightCellYellow(dynamic cell)
    {
        try { cell.Interior.Color = RGB(255, 255, 0); } catch { }
    }

    private static void HighlightCellBlue(dynamic cell)
    {
        try { cell.Interior.Color = RGB(189, 215, 238); } catch { }
    }

    private static void HighlightCellRed(dynamic cell)
    {
        try { cell.Interior.Color = RGB(255, 199, 206); } catch { }
    }

    private static void HighlightCellOrange(dynamic cell)
    {
        try { cell.Interior.Color = RGB(255, 165, 0); } catch { }
    }

    private static void ClearCellFill(dynamic cell)
    {
        try { cell.Interior.Pattern = -4142; } catch { }
    }

    private static string GetCellCommentText(dynamic cell)
    {
        try
        {
            var comment = cell.Comment;
            return comment == null
                ? string.Empty
                : Convert.ToString(comment.Text())?.Trim() ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    private static string FlattenForProtocol(string value)
    {
        return Regex.Replace(value ?? string.Empty, @"\s+", " ").Trim();
    }

    private static CellSpecialStatus DetectSpecialStatus(dynamic cell, string comment)
    {
        if (comment.Contains("сотрудник не найден в данных СКУД", StringComparison.OrdinalIgnoreCase))
            return CellSpecialStatus.EmployeeNotFound;

        if (comment.Contains("разрешённые отметки в СКУД отсутствуют", StringComparison.OrdinalIgnoreCase))
            return CellSpecialStatus.NoAllowedEvents;

        if (comment.Contains("Одна запись СКУД за день", StringComparison.OrdinalIgnoreCase)
            || comment.Contains("только одна разрешённая отметка СКУД", StringComparison.OrdinalIgnoreCase))
            return CellSpecialStatus.SingleAllowedEvent;

        if (HasCellColor(cell, RGB(255, 199, 206)))
            return CellSpecialStatus.ManualReview;

        if (HasCellColor(cell, RGB(189, 215, 238)))
            return CellSpecialStatus.SingleAllowedEvent;

        return CellSpecialStatus.None;
    }

    private static bool HasCellColor(dynamic cell, int expectedColor)
    {
        try { return Convert.ToInt32(cell.Interior.Color, CultureInfo.InvariantCulture) == expectedColor; }
        catch { return false; }
    }

    private static void HandleMissingSkudData(
        dynamic cell,
        string fio,
        DateTime date,
        double declaredHours,
        CellSpecialStatus status,
        string existingComment,
        List<string> protocol)
    {
        string declared = FormatNumber(declaredHours);
        string comparisonComment;
        string protocolStatus;

        if (status == CellSpecialStatus.EmployeeNotFound)
        {
            comparisonComment = "Смена заявлена, сотрудник не найден в данных СКУД.";
            protocolStatus = "Сотрудник не найден в СКУД";
        }
        else if (status == CellSpecialStatus.NoAllowedEvents)
        {
            comparisonComment = $"Смена заявлена на {declared} часов, но разрешённые отметки в СКУД отсутствуют.";
            protocolStatus = "Разрешённые отметки СКУД отсутствуют";
        }
        else
        {
            comparisonComment = string.IsNullOrWhiteSpace(existingComment)
                ? "Требуется ручная проверка данных СКУД."
                : $"Специальный статус сохранён: {FlattenForProtocol(existingComment)}";
            protocolStatus = string.IsNullOrWhiteSpace(existingComment)
                ? "Ручная проверка данных СКУД"
                : FlattenForProtocol(existingComment);
        }

        try { HighlightCellRed(cell); } catch { }
        try { AddOrAppendComparisonComment(cell, comparisonComment); } catch { }
        string finalValue = status == CellSpecialStatus.ManualReview ? declared : string.Empty;
        protocol.Add($"{fio};{FormatDateKey(date)};{declared};;{finalValue};;{protocolStatus}");
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

    private static void AddOrAppendComparisonComment(dynamic cell, string comparisonText)
    {
        string block = "Сравнение: " + comparisonText;

        try
        {
            string existingText = string.Empty;

            try
            {
                var old = cell.Comment;
                if (old != null)
                    existingText = Convert.ToString(old.Text())?.Trim() ?? string.Empty;
            }
            catch { }

            string newText = string.IsNullOrWhiteSpace(existingText)
                ? block
                : existingText + Environment.NewLine + Environment.NewLine + block;

            try { var old = cell.Comment; if (old != null) old.Delete(); } catch { }
            cell.AddComment(newText);
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
        Dictionary<DateTime, int> DayColumns,
        int? TotalHoursCol,
        object[,] Values,
        int UsedRows,
        int UsedCols,
        bool UsesSyntheticDayNumbers);

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

    private enum CellSpecialStatus
    {
        None,
        SingleAllowedEvent,
        NoAllowedEvents,
        EmployeeNotFound,
        ManualReview
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
        public int SkudLowerDifferences { get; set; }
        public int SkudHigherDifferences { get; set; }
        public int SpecialStatuses { get; set; }
        public int ReadErrors { get; set; }
    }
}
