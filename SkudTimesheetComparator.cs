namespace Tabel;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

public static class SkudTimesheetComparator
{
    public static event EventHandler<SkudProgressEventArgs>? ProgressChanged;

    public static Task<string> CompareTimesheetsAsync(string originalPath, string filledPath)
    {
        return RunStaTask(() => CompareTimesheets(originalPath, filledPath));
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

    public static string CompareTimesheets(string originalPath, string filledPath)
    {
        if (!File.Exists(originalPath))
            throw new FileNotFoundException("Исходный табель не найден.", originalPath);

        if (!File.Exists(filledPath))
            throw new FileNotFoundException("Заполненный табель не найден.", filledPath);

        string outputPath = Path.Combine(
            Path.GetDirectoryName(originalPath) ?? "",
            Path.GetFileNameWithoutExtension(originalPath) + "_сравнение.xlsx");

        Type? excelType = Type.GetTypeFromProgID("Excel.Application");

        if (excelType == null)
            throw new Exception("Microsoft Excel не установлен или COM-объект Excel.Application недоступен.");

        dynamic? excel = null;
        dynamic? originalBook = null;
        dynamic? filledBook = null;
        dynamic? outputBook = null;

        try
        {
            Report("Запуск Excel", "Запускаю Microsoft Excel...", 0, 100, 5);

            excel = Activator.CreateInstance(excelType);
            excel.Visible = false;
            excel.DisplayAlerts = false;

            // Сначала читаем исходный табель в память и закрываем его
            Report("Чтение исходного табеля", "Читаю исходный табель в память...", 0, 100, 10);

            originalBook = excel.Workbooks.Open(originalPath);
            dynamic originalSheet = originalBook.Worksheets[1];

            object[,] originalValues = ((dynamic)originalSheet.UsedRange).Value2;
            int originalRows = originalSheet.UsedRange.Rows.Count;
            int originalCols = originalSheet.UsedRange.Columns.Count;

            originalBook.Close(false);
            ReleaseComObject(originalSheet);
            ReleaseComObject(originalBook);
            originalBook = null;
            originalSheet = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();

            // Читаем заполненный табель в память и закрываем его
            Report("Чтение заполненного табеля", "Читаю заполненный табель в память...", 0, 100, 20);

            filledBook = excel.Workbooks.Open(filledPath);
            dynamic filledSheet = filledBook.Worksheets[1];

            object[,] filledValues = ((dynamic)filledSheet.UsedRange).Value2;
            int filledRows = filledSheet.UsedRange.Rows.Count;
            int filledCols = filledSheet.UsedRange.Columns.Count;

            filledBook.Close(false);
            ReleaseComObject(filledSheet);
            ReleaseComObject(filledBook);
            filledBook = null;
            filledSheet = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();

            // Копируем исходный табель в выходной файл
            Report("Подготовка выходного файла", "Создаю копию исходного табеля...", 0, 100, 30);

            if (File.Exists(outputPath))
                File.Delete(outputPath);

            // Копируем файл
            File.Copy(originalPath, outputPath, true);

            // Открываем выходной файл для применения форматирования
            outputBook = excel.Workbooks.Open(outputPath);
            dynamic outputSheet = outputBook.Worksheets[1];

            // Выполняем сравнение с использованием загруженных данных
            Report("Сравнение табелей", "Сравниваю табели...", 0, 100, 40);

            PerformComparisonWithData(
                originalValues, originalRows, originalCols,
                filledValues, filledRows, filledCols,
                outputSheet);

            Report("Сохранение", "Сохраняю файл сравнения...", 0, 100, 95);

            outputBook.Save();
            outputBook.Close(false);
            ReleaseComObject(outputSheet);
            ReleaseComObject(outputBook);
            outputBook = null;

            Report("Завершено", $"Файл сравнения создан: {outputPath}", 100, 100, 100, true);

            return outputPath;
        }
        catch (Exception ex)
        {
            Report("ОШИБКА", ex.Message, 0, 100, 100, true);
            throw;
        }
        finally
        {
            try
            {
                if (outputBook != null)
                    outputBook.Close(false);
            }
            catch { }

            if (originalBook != null)
            {
                try
                {
                    originalBook.Close(false);
                }
                catch { }
                ReleaseComObject(originalBook);
            }

            if (filledBook != null)
            {
                try
                {
                    filledBook.Close(false);
                }
                catch { }
                ReleaseComObject(filledBook);
            }

            if (outputBook != null)
                ReleaseComObject(outputBook);

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

    private static void PerformComparisonWithData(
        object[,] originalValues, int originalRows, int originalCols,
        object[,] filledValues, int filledRows, int filledCols,
        dynamic outputSheet)
    {
        int usedRows = Math.Max(originalRows, filledRows);
        int usedCols = Math.Max(originalCols, filledCols);

        // Ищем колонку ФИО
        int fioHeaderRow = -1;
        int fioCol = -1;

        Report("Анализ табелей", "Ищу колонку ФИО...", 0, 100, 40);

        for (int row = 1; row <= Math.Min(50, usedRows); row++)
        {
            for (int col = 1; col <= usedCols; col++)
            {
                string text = GetValueText(originalValues[row, col]);

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

        // Ищем дневные колонки
        var dayColumns = FindDayColumnsWithValues(originalValues, fioHeaderRow, usedCols);

        if (dayColumns.Count == 0)
            throw new Exception("В табеле не найдены колонки дней месяца.");

        int totalComparisons = dayColumns.Count * (usedRows - fioHeaderRow);
        int currentComparison = 0;
        int highlightedCells = 0;

        Report("Сравнение табелей", "Начинаю сравнение ячеек часов...", 0, 100, 45);

        for (int row = fioHeaderRow + 1; row <= usedRows; row++)
        {
            string fio = GetValueText(originalValues[row, fioCol]);

            if (string.IsNullOrWhiteSpace(fio))
            {
                currentComparison += dayColumns.Count;
                continue;
            }

            foreach (var pair in dayColumns)
            {
                int col = pair.Value;

                try
                {
                    object originalValue = row <= originalRows && col <= originalCols ? originalValues[row, col] : null;
                    object filledValue = row <= filledRows && col <= filledCols ? filledValues[row, col] : null;

                    double originalHours = ConvertToDouble(originalValue);
                    double filledHours = ConvertToDouble(filledValue);

                    double difference = originalHours - filledHours;

                    if (difference > 1)
                    {
                        // Жёлтое выделение
                        HighlightCellYellow(outputSheet.Cells[row, col]);

                        string comment = $"Исходный табель: {originalHours}\nЗаполненный табель: {filledHours}\nРазница: {difference:F1} часа";
                        AddCommentSafe(outputSheet.Cells[row, col], comment);

                        highlightedCells++;
                    }
                    else if (difference < -1)
                    {
                        // Зелёное выделение если заполненный > исходного на 1
                        HighlightCellGreen(outputSheet.Cells[row, col]);
                        highlightedCells++;
                    }
                }
                catch { }

                currentComparison++;
                if (currentComparison % 50 == 0)
                {
                    int percent = 40 + (int)((double)currentComparison / totalComparisons * 50);
                    Report("Сравнение табелей", $"Сравнено ячеек: {currentComparison} из {totalComparisons}", currentComparison, totalComparisons, percent);
                }
            }
        }

        Report("Итог сравнения", $"Выделено ячеек: {highlightedCells}", totalComparisons, totalComparisons, 90);
    }

    private static double ConvertToDouble(object value)
    {
        if (value == null)
            return 0;

        if (value is double d)
            return d;

        if (value is int i)
            return i;

        if (double.TryParse(value.ToString(), out double result))
            return result;

        return 0;
    }

    private static void HighlightCellYellow(dynamic cell)
    {
        try
        {
            cell.Interior.Color = RGB(255, 255, 0);
        }
        catch { }
    }

    private static void HighlightCellGreen(dynamic cell)
    {
        try
        {
            cell.Interior.Color = RGB(0, 255, 0);
        }
        catch { }
    }

    private static void AddCommentSafe(dynamic cell, string commentText)
    {
        try
        {
            dynamic comment = cell.AddComment(commentText);
            comment.Shape.TextFrame.AutoSize = true;
        }
        catch { }
    }

    private static int RGB(int red, int green, int blue)
    {
        return red + (green << 8) + (blue << 16);
    }

    private static Dictionary<int, int> FindDayColumnsWithValues(object[,] values, int headerRow, int usedCols)
    {
        var result = new Dictionary<int, int>();

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

    private static string NormalizeHeader(string text)
    {
        if (text == null)
            return "";

        text = text.Replace('\u00A0', ' ');
        System.Text.RegularExpressions.Regex.Replace(text, @"\s+", " ");
        text = text.Trim().ToLowerInvariant().Replace('ё', 'е');

        return text;
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
