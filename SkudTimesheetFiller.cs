namespace Tabel;

using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

internal enum PersonSearchResult
{
    Found,
    NotFound,
    AmbiguousFound
}

internal sealed class PersonSearchInfo
{
    public PersonSearchResult Status { get; set; }
    public string? PersonKey { get; set; }
    public string SearchKey { get; set; } = string.Empty;
    public string StatusText { get; set; } = string.Empty;
    public List<string> FoundVariants { get; set; } = new();
}

public static class SkudExcelInteropFiller
{
    private const double MaxDeclaredShiftHours = 23d;

    private enum InputFileExpectation
    {
        OriginalTimesheet,
        FilledTimesheet,
        SkudExport
    }

    public static event EventHandler<SkudProgressEventArgs>? ProgressChanged;

    public static Task<string> FillTimesheetAsync(string timesheetPath, string skudPath)
    {
        return RunStaTask(() => FillTimesheetFile(timesheetPath, skudPath));
    }

    public static Task ValidateOriginalTimesheetInputAsync(string path)
    {
        return ValidationWorker.Instance.ValidateAsync(path, InputFileExpectation.OriginalTimesheet);
    }

    public static Task ValidateFilledTimesheetInputAsync(string path)
    {
        return ValidationWorker.Instance.ValidateAsync(path, InputFileExpectation.FilledTimesheet);
    }

    public static Task ValidateSkudInputAsync(string path)
    {
        return ValidationWorker.Instance.ValidateAsync(path, InputFileExpectation.SkudExport);
    }

    public static void ShutdownValidationSession()
    {
        ValidationWorker.Shutdown();
    }

    private static void ValidateInputFile(
        string path,
        InputFileExpectation expectation,
        dynamic excel)
    {
        if (!File.Exists(path))
            throw new FileNotFoundException("Выбранный файл не найден.", path);

        if (!string.Equals(Path.GetExtension(path), ".xlsx", StringComparison.OrdinalIgnoreCase))
            throw new Exception("Поддерживаются только Excel-файлы формата .xlsx.");

        dynamic? workbooks = null;
        dynamic? workbook = null;
        dynamic? timesheetSheet = null;

        try
        {
            workbooks = excel.Workbooks;
            workbook = workbooks.Open(path, 0, true);

            if (expectation == InputFileExpectation.SkudExport)
            {
                ValidateSkudColumns(workbook);
            }
            else
            {
                string workbookRole = TabelWorkbookMetadata.GetExplicitRole(workbook);
                ValidateWorkbookRole(workbookRole, expectation);
                var protocol = new ProcessingProtocol();

                timesheetSheet = FindTimesheetWorksheet(
                    workbook,
                    new DateTime(2000, 1, 1),
                    protocol,
                    out TimesheetStructure _,
                    maxRows: 1000,
                    maxColumns: 512);

                ReleaseComObject(timesheetSheet);
                timesheetSheet = null;

                if (string.IsNullOrWhiteSpace(workbookRole)
                    && TabelWorkbookMetadata.IsLegacyComparison(workbook))
                    throw new Exception("Эта книга определена как старый файл результата сравнения по комментариям или цветовой разметке и не может использоваться как входной файл.");
            }
        }
        finally
        {
            if (timesheetSheet != null)
                ReleaseComObject(timesheetSheet);

            try { if (workbook != null) workbook.Close(false); } catch { }
            if (workbook != null)
                ReleaseComObject(workbook);
            if (workbooks != null)
                ReleaseComObject(workbooks);
        }
    }

    private sealed class ValidationWorker : IDisposable
    {
        private sealed record Request(
            string Path,
            InputFileExpectation Expectation,
            TaskCompletionSource<bool> Completion);

        private sealed record CacheEntry(long Length, DateTime LastWriteTimeUtc);

        private static readonly object SyncRoot = new();
        private static ValidationWorker? _instance;

        private readonly BlockingCollection<Request> _requests = new();
        private readonly Dictionary<(string Path, InputFileExpectation Expectation), CacheEntry> _cache = new();
        private readonly Thread _thread;
        private bool _disposed;

        public static ValidationWorker Instance
        {
            get
            {
                lock (SyncRoot)
                    return _instance ??= new ValidationWorker();
            }
        }

        private ValidationWorker()
        {
            _thread = new Thread(ProcessRequests)
            {
                IsBackground = true,
                Name = "Tabel Excel validation"
            };
            _thread.SetApartmentState(ApartmentState.STA);
            _thread.Start();
        }

        public Task ValidateAsync(string path, InputFileExpectation expectation)
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(ValidationWorker));

            var completion = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            _requests.Add(new Request(path, expectation, completion));
            return completion.Task;
        }

        public static void Shutdown()
        {
            ValidationWorker? worker;

            lock (SyncRoot)
            {
                worker = _instance;
                _instance = null;
            }

            worker?.Dispose();
        }

        private void ProcessRequests()
        {
            foreach (Request request in _requests.GetConsumingEnumerable())
            {
                dynamic? excel = null;

                try
                {
                    var file = new FileInfo(request.Path);
                    var key = (Path.GetFullPath(request.Path), request.Expectation);
                    var stamp = new CacheEntry(file.Length, file.LastWriteTimeUtc);

                    if (!_cache.TryGetValue(key, out CacheEntry? cached) || cached != stamp)
                    {
                        excel = CreateValidationExcel();
                        ValidateInputFile(request.Path, request.Expectation, excel);
                        _cache[key] = stamp;
                    }

                    request.Completion.TrySetResult(true);
                }
                catch (Exception ex)
                {
                    request.Completion.TrySetException(ex);
                }
                finally
                {
                    if (excel != null)
                    {
                        try { excel.Quit(); } catch { }
                        ReleaseComObject(excel);
                    }

                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }
        }

        private static dynamic CreateValidationExcel()
        {
            Type? excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null)
                throw new Exception("Microsoft Excel не установлен или COM-объект Excel.Application недоступен.");

            dynamic excel = Activator.CreateInstance(excelType)
                ?? throw new Exception("Не удалось запустить Microsoft Excel.");
            excel.Visible = false;
            excel.DisplayAlerts = false;
            excel.ScreenUpdating = false;
            excel.EnableEvents = false;
            try { excel.Calculation = -4135; } catch { }
            try { excel.AskToUpdateLinks = false; } catch { }
            return excel;
        }

        public void Dispose()
        {
            if (_disposed)
                return;

            _disposed = true;
            _requests.CompleteAdding();

            if (Thread.CurrentThread != _thread)
                _thread.Join(TimeSpan.FromSeconds(10));

            _requests.Dispose();
        }
    }

    private static void ValidateSkudColumns(dynamic workbook)
    {
        int sheetCount = workbook.Worksheets.Count;

        for (int sheetIndex = 1; sheetIndex <= sheetCount; sheetIndex++)
        {
            dynamic? sheet = null;
            dynamic? usedRange = null;

            try
            {
                sheet = workbook.Worksheets[sheetIndex];
                ExcelWorksheetData data = ExcelWorksheetReader.Read(sheet, maxRows: 50, maxColumns: 512);
                object[,] values = data.Values;
                int usedRows = data.Rows;
                int usedCols = data.Columns;

                if (TryFindSkudColumns(values, usedRows, usedCols, out _, out _, out _, out _))
                    return;
            }
            finally
            {
                if (usedRange != null)
                    ReleaseComObject(usedRange);
                if (sheet != null)
                    ReleaseComObject(sheet);
            }
        }

        throw new Exception("Не найдены обязательные колонки СКУД: Дата записи, Персона, Сообщение.");
    }

    private static void ValidateWorkbookRole(string workbookRole, InputFileExpectation expectation)
    {
        if (string.IsNullOrWhiteSpace(workbookRole))
            return;

        if (string.Equals(workbookRole, TabelWorkbookMetadata.ComparisonRole, StringComparison.OrdinalIgnoreCase))
            throw new Exception("Эта книга помечена как файл результата сравнения и не может использоваться как входной файл.");

        if (expectation == InputFileExpectation.OriginalTimesheet
            && string.Equals(workbookRole, TabelWorkbookMetadata.FilledTimesheetRole, StringComparison.OrdinalIgnoreCase))
            throw new Exception("Эта книга помечена как заполненный табель и не может использоваться как исходный табель подрядчика.");

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

        string protocolPath = Path.Combine(
            Path.GetDirectoryName(timesheetPath) ?? "",
            Path.GetFileNameWithoutExtension(timesheetPath) + "_протокол.txt");

        Type? excelType = Type.GetTypeFromProgID("Excel.Application");
        if (excelType == null)
            throw new Exception("Microsoft Excel не установлен или COM-объект Excel.Application недоступен.");

        var protocol = new ProcessingProtocol();
        dynamic? excel = null;
        dynamic? skudBook = null;
        dynamic? outputBook = null;

        try
        {
            Report("Запуск Excel", "Запускаю Microsoft Excel...", 0, 100, 5);

            if (File.Exists(outputPath))
                File.Delete(outputPath);

            File.Copy(timesheetPath, outputPath, true);
            protocol.General.Add($"Создана копия исходного табеля: {outputPath}");

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

            Report("Чтение СКУД", "Ищу лист СКУД и читаю события...", 0, 100, 10);
            skudBook = excel.Workbooks.Open(skudPath);
            var skudRead = ReadSkudEvents(skudBook, skudPath, protocol);

            skudBook.Close(false);
            ReleaseComObject(skudBook);
            skudBook = null;

            if (skudRead.EventsByPerson.Count == 0)
                throw new Exception("В файле СКУД не найдено допустимых событий прохода.");

            Report("Открытие табеля", "Открываю копию табеля для заполнения...", 0, 100, 45);
            outputBook = excel.Workbooks.Open(outputPath);

            var fillResult = FillTimesheet(outputBook, skudRead.EventsByPerson, skudRead.MainMonth, skudRead.Months, protocol);

            Report("Сохранение", "Сохраняю заполненный табель...", 0, 100, 95);
            TabelWorkbookMetadata.SetRole(outputBook, TabelWorkbookMetadata.FilledTimesheetRole);
            outputBook.Save();
            outputBook.Close(false);
            ReleaseComObject(outputBook);
            outputBook = null;

            WriteProtocol(protocolPath, protocol);

            string summary =
                $"Обработка завершена. Сотрудников в табеле: {fillResult.TotalEmployees}. " +
                $"Сопоставлено со СКУД: {fillResult.MatchedEmployees}. Не найдено: {fillResult.NotFoundEmployees}. " +
                $"Неоднозначно: {fillResult.AmbiguousEmployees}. Заполнено ячеек: {fillResult.FilledCells}. " +
                $"Ошибок: {protocol.Errors.Count}. Файл: {outputPath}";

            Report("Готово", summary, 100, 100, 100, true);
            return outputPath;
        }
        catch (Exception ex)
        {
            protocol.Errors.Add($"Критическая ошибка; {ex.Message}");
            try { WriteProtocol(protocolPath, protocol); } catch { }
            Report("ОШИБКА", ex.Message, 0, 100, 100, true);
            throw;
        }
        finally
        {
            try { if (skudBook != null) skudBook.Close(false); } catch { }
            try { if (outputBook != null) outputBook.Close(false); } catch { }

            if (skudBook != null)
                ReleaseComObject(skudBook);

            if (outputBook != null)
                ReleaseComObject(outputBook);

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

    private static FillResult FillTimesheet(
        dynamic workbook,
        Dictionary<string, Dictionary<DateTime, DayEvents>> eventsByPerson,
        DateTime month,
        IReadOnlyCollection<DateTime> skudMonths,
        ProcessingProtocol protocol)
    {
        dynamic sheet = FindTimesheetWorksheet(workbook, month, protocol, out TimesheetStructure structure);
        EnsureTotalColumnFormatting(sheet, structure, protocol);

        if (skudMonths.Count > 1 && HasOnlyDayNumberDateHeaders(structure))
            throw new Exception("В СКУД найдены записи за несколько месяцев, а в табеле даты указаны только номерами дней. Укажите в табеле полные даты или используйте выгрузку СКУД за один месяц.");

        Report("Анализ табеля", $"Лист табеля найден. Строка дат: {structure.DateRow}, колонка ФИО: {ColumnName(structure.FioCol)}.", 0, 100, 50);

        var personKeyIndex = BuildPersonKeyIndex(eventsByPerson);
        var employees = ReadTimesheetEmployees(structure, protocol);
        var result = new FillResult { TotalEmployees = employees.Count };

        int daysInMonth = DateTime.DaysInMonth(month.Year, month.Month);

        Report("Заполнение табеля", $"В табеле найдено сотрудников: {employees.Count}.", 0, employees.Count, 55);

        for (int i = 0; i < employees.Count; i++)
        {
            var employee = employees[i];
            HashSet<int> invalidDeclaredColumns = MarkInvalidDeclaredShiftHours(
                sheet,
                structure,
                employee,
                month,
                daysInMonth,
                protocol);
            var search = FindPersonInSkudData(employee.OriginalName, eventsByPerson, personKeyIndex);

            protocol.Matching.Add($"{employee.OriginalName}; ключ={search.SearchKey}; статус={search.StatusText}; варианты={string.Join(", ", search.FoundVariants)}");

            if (search.Status == PersonSearchResult.NotFound)
            {
                result.NotFoundEmployees++;

                foreach (var dateColumn in structure.DateColumns.OrderBy(x => x.Key))
                {
                    DateTime date = dateColumn.Key.Date;
                    int col = dateColumn.Value;

                    if (date.Month != month.Month || date.Year != month.Year || date.Day > daysInMonth)
                        continue;

                    if (invalidDeclaredColumns.Contains(col))
                        continue;

                    if (!TryGetDeclaredShiftHours(structure, employee.RowIndex, col, out _))
                        continue;

                    dynamic cell = sheet.Cells[employee.RowIndex, col];
                    MarkDeclaredShiftWithoutEmployeeInSkud(cell);
                    protocol.Filling.Add(
                        $"{employee.OriginalName}; {date:dd.MM.yyyy}; часы=; " +
                        "статус=Смена заявлена, сотрудник не найден в данных СКУД");
                }

                FinalizeEmployeeTotal(sheet, employee.RowIndex, structure, invalidDeclaredColumns, protocol);

                continue;
            }

            if (search.Status == PersonSearchResult.AmbiguousFound)
            {
                result.AmbiguousEmployees++;
                HighlightCellAmbiguous(sheet.Cells[employee.RowIndex, structure.FioCol], search.SearchKey, search.FoundVariants);

                FinalizeEmployeeTotal(sheet, employee.RowIndex, structure, invalidDeclaredColumns, protocol);

                continue;
            }

            result.MatchedEmployees++;
            var personDays = eventsByPerson[search.PersonKey!];

            foreach (var dateColumn in structure.DateColumns.OrderBy(x => x.Key))
            {
                DateTime date = dateColumn.Key.Date;
                int col = dateColumn.Value;

                if (date.Month != month.Month || date.Year != month.Year || date.Day > daysInMonth)
                    continue;

                // Обрабатываем только даты, на которые в исходном табеле была заявлена смена.
                // Пустые и незапланированные даты не заполняем, не комментируем и не выделяем.
                if (invalidDeclaredColumns.Contains(col))
                    continue;

                if (!TryGetDeclaredShiftHours(structure, employee.RowIndex, col, out double declaredHours))
                    continue;

                if (!personDays.TryGetValue(date, out var dayEvents) || !dayEvents.HasAnyEvents)
                {
                    dynamic missingCell = sheet.Cells[employee.RowIndex, col];
                    MarkDeclaredShiftWithoutAllowedSkudEvents(missingCell, declaredHours);
                    protocol.Filling.Add(
                        $"{employee.OriginalName}; {date:dd.MM.yyyy}; заявлено={FormatHours(declaredHours)}; часы=; " +
                        "статус=Разрешённые отметки в СКУД отсутствуют");
                    continue;
                }

                var calc = dayEvents.Calculate();

                if (calc.ProblemEventCount > 0)
                {
                    protocol.Filling.Add(
                        $"{employee.OriginalName}; {date:dd.MM.yyyy}; проблемных сообщений СКУД={calc.ProblemEventCount}; " +
                        $"сообщения={string.Join("; ", calc.ProblemMessages)}");
                }

                if (calc.EventCount == 0)
                {
                    dynamic missingAllowedCell = sheet.Cells[employee.RowIndex, col];
                    MarkDeclaredShiftWithoutAllowedSkudEvents(missingAllowedCell, declaredHours);
                    protocol.Filling.Add(
                        $"{employee.OriginalName}; {date:dd.MM.yyyy}; заявлено={FormatHours(declaredHours)}; часы=; " +
                        "статус=Разрешённые отметки в СКУД отсутствуют");
                    continue;
                }

                if (calc.RequiresManualReview)
                {
                    dynamic reviewCell = sheet.Cells[employee.RowIndex, col];
                    invalidDeclaredColumns.Add(col);
                    ClearCellContentsSafe(reviewCell);
                    try { reviewCell.Interior.Color = RGB(255, 199, 206); } catch { }
                    AddCommentSafe(reviewCell, calc.Status);
                    protocol.Filling.Add(
                        $"{employee.OriginalName}; {date:dd.MM.yyyy}; событий={calc.EventCount}; " +
                        $"первое={FormatTime(calc.FirstEvent)}; последнее={FormatTime(calc.LastEvent)}; часы=; статус={calc.Status}");
                    continue;
                }

                if (!calc.ShouldWriteValue || !calc.Hours.HasValue)
                    continue;

                dynamic cell = sheet.Cells[employee.RowIndex, col];
                cell.Value2 = calc.Hours.Value;
                SetNumberFormatSafe(cell);

                if (calc.IsSingleAllowedEvent)
                {
                    try { cell.Interior.Color = RGB(189, 215, 238); } catch { }
                    AddCommentSafe(cell, "Одна запись СКУД за день. Установлено 2 часа.");
                }

                result.FilledCells++;
                protocol.Filling.Add(
                    $"{employee.OriginalName}; {date:dd.MM.yyyy}; событий={calc.EventCount}; " +
                    $"первое={FormatTime(calc.FirstEvent)}; последнее={FormatTime(calc.LastEvent)}; часы={calc.Hours}; статус={calc.Status}");
            }

            FinalizeEmployeeTotal(sheet, employee.RowIndex, structure, invalidDeclaredColumns, protocol);

            int current = i + 1;
            if (current % 5 == 0 || current == employees.Count)
                Report("Заполнение табеля", $"Заполняю сотрудников: {current} из {employees.Count}", current, employees.Count, ScalePercent(current, employees.Count, 55, 90));
        }

        return result;
    }

    private static dynamic FindTimesheetWorksheet(
        dynamic workbook,
        DateTime month,
        ProcessingProtocol protocol,
        out TimesheetStructure structure,
        int? maxRows = null,
        int? maxColumns = null)
    {
        int sheetsCount = workbook.Worksheets.Count;
        Exception? lastStructureError = null;

        for (int i = 1; i <= sheetsCount; i++)
        {
            dynamic sheet = workbook.Worksheets[i];
            dynamic? usedRange = null;

            try
            {
                ExcelWorksheetData data = ExcelWorksheetReader.Read(sheet, maxRows, maxColumns);
                object[,] values = data.Values;
                int usedRows = data.Rows;
                int usedCols = data.Columns;

                if (TryDetectTimesheetStructure(values, usedRows, usedCols, month, protocol, out structure))
                    return sheet;
            }
            catch (Exception ex)
            {
                lastStructureError = ex;
                protocol.Errors.Add($"Табель; лист {i}; ошибка анализа структуры; {ex.Message}");
            }
            finally
            {
                if (usedRange != null)
                    ReleaseComObject(usedRange);
            }
        }

        throw lastStructureError ?? new Exception("Не удалось определить лист табеля. Не найдены ФИО сотрудников или строка с датами.");
    }

    private static bool TryDetectTimesheetStructure(
        object[,] values,
        int usedRows,
        int usedCols,
        DateTime month,
        ProcessingProtocol protocol,
        out TimesheetStructure structure)
    {
        structure = default!;
        int fioCol = FindFioColumn(values, usedRows, usedCols, out int headerRow);

        if (fioCol < 0)
            throw new Exception("В табеле не найдена колонка ФИО.");

        var dateColumns = new Dictionary<DateTime, int>();
        int dateRow = -1;
        var dateRowCandidates = new List<(int Row, Dictionary<DateTime, int> Columns)>();

        foreach (int row in GetDateRowSearchOrder(headerRow, usedRows))
        {
            var detected = ReadDateColumns(values, row, usedCols, month);
            if (detected.Count >= 5)
                dateRowCandidates.Add((row, detected));
        }

        if (dateRowCandidates.Count > 0)
        {
            var selected = dateRowCandidates
                .OrderBy(x => Math.Abs(x.Row - headerRow))
                .ThenBy(x => x.Row)
                .First();

            dateRow = selected.Row;
            dateColumns = selected.Columns;

            if (dateRowCandidates.Count > 1)
            {
                string rows = string.Join(", ", dateRowCandidates.Select(x => x.Row));
                protocol.General.Add($"Табель; найдено несколько возможных строк дат: {rows}. Выбрана ближайшая к строке ФИО: {dateRow}.");
            }
        }

        if (dateColumns.Count == 0)
            throw new Exception("В табеле не найдена строка с датами или номерами дней месяца.");

        if (!HasEmployeeLikeRows(values, usedRows, fioCol, Math.Max(headerRow, dateRow) + 1))
            return false;

        int? totalCol = FindTotalColumn(values, headerRow, usedCols)
            ?? FindTotalColumn(values, dateRow, usedCols)
            ?? FindTotalColumnAfterLastMonthDate(dateColumns, month);
        structure = new TimesheetStructure(
            Math.Max(headerRow, dateRow),
            dateRow,
            Math.Max(headerRow, dateRow) + 1,
            fioCol,
            dateColumns,
            totalCol,
            values,
            usedRows);

        return true;
    }

    private static SkudReadResult ReadSkudEvents(dynamic workbook, string filePath, ProcessingProtocol protocol)
    {
        int sheetsCount = workbook.Worksheets.Count;
        var result = new Dictionary<string, Dictionary<DateTime, DayEvents>>(StringComparer.OrdinalIgnoreCase);
        var monthCounter = new Dictionary<(int Year, int Month), int>();

        for (int sheetIndex = 1; sheetIndex <= sheetsCount; sheetIndex++)
        {
            dynamic sheet = workbook.Worksheets[sheetIndex];
            dynamic? usedRange = null;

            try
            {
                ExcelWorksheetData data = ExcelWorksheetReader.Read(sheet);
                object[,] values = data.Values;
                int usedRows = data.Rows;
                int usedCols = data.Columns;

                if (!TryFindSkudColumns(values, usedRows, usedCols, out int headerRow, out int dateCol, out int personCol, out int messageCol))
                    continue;

                protocol.General.Add($"Лист СКУД найден: {sheetIndex}. Строка заголовков: {headerRow}.");

                for (int row = headerRow + 1; row <= usedRows; row++)
                {
                    string person = GetValueText(values[row, personCol]);
                    string message = GetValueText(values[row, messageCol]);

                    if (string.IsNullOrWhiteSpace(person))
                        continue;

                    if (!TryReadDateTimeFromValue(values[row, dateCol], out DateTime eventDateTime))
                    {
                        protocol.Errors.Add($"СКУД; {Path.GetFileName(filePath)}; лист {sheetIndex}; строка {row}; колонка {ColumnName(dateCol)}; не прочитана дата; {GetValueText(values[row, dateCol])}");
                        continue;
                    }

                    bool isAllowed = IsAllowedSkudMessage(message);
                    bool isProblem = IsProblemSkudMessage(message);

                    if (!isAllowed && !isProblem)
                    {
                        protocol.Errors.Add($"СКУД; {Path.GetFileName(filePath)}; лист {sheetIndex}; строка {row}; сообщение пропущено; {message}");
                        continue;
                    }

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

                    var skudEvent = new SkudEvent
                    {
                        OriginalPerson = person,
                        DateTime = eventDateTime,
                        Message = message,
                        SourceRowIndex = row
                    };

                    if (isAllowed)
                        dayEvents.AddAllowedEvent(skudEvent);
                    else
                        dayEvents.AddProblemEvent(skudEvent);

                    var monthKey = (eventDateTime.Year, eventDateTime.Month);
                    monthCounter[monthKey] = monthCounter.TryGetValue(monthKey, out int count) ? count + 1 : 1;
                }

                if (result.Count > 0)
                {
                    var main = monthCounter.OrderByDescending(x => x.Value).First().Key;
                    var months = monthCounter.Keys
                        .Select(x => new DateTime(x.Year, x.Month, 1))
                        .OrderBy(x => x)
                        .ToList();

                    return new SkudReadResult(result, new DateTime(main.Year, main.Month, 1), months);
                }
            }
            catch (Exception ex)
            {
                protocol.Errors.Add($"СКУД; {Path.GetFileName(filePath)}; лист {sheetIndex}; ошибка чтения; {ex.Message}");
            }
            finally
            {
                if (usedRange != null)
                    ReleaseComObject(usedRange);
            }
        }

        throw new Exception("Не удалось найти лист СКУД. Не найдены обязательные колонки: Дата записи, Персона, Сообщение.");
    }

    private static bool TryFindSkudColumns(object[,] values, int usedRows, int usedCols, out int headerRow, out int dateCol, out int personCol, out int messageCol)
    {
        headerRow = -1;
        dateCol = -1;
        personCol = -1;
        messageCol = -1;

        for (int row = 1; row <= Math.Min(50, usedRows); row++)
        {
            int currentDateCol = -1;
            int currentPersonCol = -1;
            int currentMessageCol = -1;

            for (int col = 1; col <= usedCols; col++)
            {
                string header = NormalizeHeader(GetValueText(values[row, col]));

                if (header == "дата записи")
                    currentDateCol = col;
                else if (header == "персона")
                    currentPersonCol = col;
                else if (header == "сообщение" || header == "message")
                    currentMessageCol = col;
            }

            if (currentDateCol > 0 && currentPersonCol > 0 && currentMessageCol > 0)
            {
                headerRow = row;
                dateCol = currentDateCol;
                personCol = currentPersonCol;
                messageCol = currentMessageCol;
                return true;
            }
        }

        return false;
    }

    private static List<TimesheetEmployee> ReadTimesheetEmployees(TimesheetStructure structure, ProcessingProtocol protocol)
    {
        var result = new List<TimesheetEmployee>();

        for (int row = structure.FirstEmployeeRow; row <= structure.UsedRows; row++)
        {
            string fio = GetValueText(structure.Values[row, structure.FioCol]);

            if (string.IsNullOrWhiteSpace(fio))
                continue;

            if (!LooksLikePersonName(fio))
                continue;

            result.Add(new TimesheetEmployee(row, fio));
        }

        if (result.Count == 0)
            protocol.Errors.Add("Табель; не найдены строки сотрудников в колонке ФИО.");

        return result;
    }

    private static PersonSearchInfo FindPersonInSkudData(
        string fioFromTimesheet,
        Dictionary<string, Dictionary<DateTime, DayEvents>> eventsByPerson,
        Dictionary<string, List<string>> personKeyIndex)
    {
        string normalizedFio = NormalizeFio(fioFromTimesheet);
        string searchKey = BuildPersonSearchKey(fioFromTimesheet);
        var result = new PersonSearchInfo { SearchKey = string.IsNullOrWhiteSpace(searchKey) ? FirstWord(normalizedFio) : searchKey };

        if (!string.IsNullOrWhiteSpace(normalizedFio) && eventsByPerson.ContainsKey(normalizedFio))
        {
            result.Status = PersonSearchResult.Found;
            result.PersonKey = normalizedFio;
            result.StatusText = "найден по полному ФИО";
            result.FoundVariants.Add(normalizedFio);
            return result;
        }

        List<string>? shortKeyCandidates = null;

        if (!string.IsNullOrWhiteSpace(searchKey) && personKeyIndex.TryGetValue(searchKey, out var keys))
        {
            shortKeyCandidates = keys
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Where(candidate => IsSafeShortKeyMatch(fioFromTimesheet, candidate))
                .ToList();

            if (shortKeyCandidates.Count == 1)
            {
                string candidate = shortKeyCandidates[0];
                result.Status = PersonSearchResult.Found;
                result.PersonKey = candidate;
                result.StatusText = "найден по краткому ключу";
                result.FoundVariants.Add(candidate);
                return result;
            }
        }

        FuzzyMatchResult<string> fuzzy = FioFuzzyMatcher.Find(
            normalizedFio,
            eventsByPerson.Keys,
            candidate => candidate);

        if (fuzzy.Status == FuzzyMatchStatus.Found)
        {
            string candidate = fuzzy.Match!;
            result.Status = PersonSearchResult.Found;
            result.PersonKey = candidate;
            result.StatusText = $"найден по нечёткому совпадению: расстояние {fuzzy.Distance}, отличие {fuzzy.DifferencePercent:0.#}%";
            result.FoundVariants.Add(candidate);
            return result;
        }

        if (fuzzy.Status == FuzzyMatchStatus.Ambiguous)
        {
            result.Status = PersonSearchResult.AmbiguousFound;
            result.StatusText = $"неоднозначное нечёткое совпадение: расстояние {fuzzy.Distance}";
            result.FoundVariants = fuzzy.Candidates.ToList();
            return result;
        }

        if (shortKeyCandidates is { Count: > 0 })
        {
            result.Status = PersonSearchResult.AmbiguousFound;
            result.StatusText = shortKeyCandidates.Count == 1
                ? "сомнительное совпадение по краткому ключу"
                : "неоднозначно";
            result.FoundVariants = shortKeyCandidates;
            return result;
        }

        result.Status = PersonSearchResult.NotFound;
        result.StatusText = "не найден";
        return result;
    }

    private static bool IsSafeShortKeyMatch(string timesheetName, string skudName)
    {
        var timesheetParts = ParseFioParts(timesheetName);
        var skudParts = ParseFioParts(skudName);

        if (string.IsNullOrWhiteSpace(timesheetParts.Surname)
            || string.IsNullOrWhiteSpace(timesheetParts.FirstName)
            || string.IsNullOrWhiteSpace(skudParts.Surname)
            || string.IsNullOrWhiteSpace(skudParts.FirstName))
            return true;

        if (!string.Equals(timesheetParts.Surname, skudParts.Surname, StringComparison.OrdinalIgnoreCase))
            return false;

        bool timesheetHasFullFirstName = timesheetParts.FirstName.Length > 1;
        bool skudHasFullFirstName = skudParts.FirstName.Length > 1;

        if (timesheetHasFullFirstName && skudHasFullFirstName)
            return string.Equals(timesheetParts.FirstName, skudParts.FirstName, StringComparison.OrdinalIgnoreCase);

        return timesheetParts.FirstName[0] == skudParts.FirstName[0];
    }

    private static FioParts ParseFioParts(string fio)
    {
        var parts = NormalizeFio(fio).Split(' ', StringSplitOptions.RemoveEmptyEntries);

        return new FioParts(
            parts.Length > 0 ? parts[0] : string.Empty,
            parts.Length > 1 ? parts[1] : string.Empty,
            parts.Length > 2 ? parts[2] : string.Empty);
    }

    private static Dictionary<string, List<string>> BuildPersonKeyIndex(Dictionary<string, Dictionary<DateTime, DayEvents>> eventsByPerson)
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

    private static Dictionary<DateTime, int> ReadDateColumns(object[,] values, int row, int usedCols, DateTime month)
    {
        var result = new Dictionary<DateTime, int>();

        for (int col = 1; col <= usedCols; col++)
        {
            if (TryReadDayHeader(values[row, col], month, out DateTime date))
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

    private static bool HasOnlyDayNumberDateHeaders(TimesheetStructure structure)
    {
        if (structure.DateColumns.Count == 0 || structure.DateRow <= 0)
            return false;

        foreach (int col in structure.DateColumns.Values.Distinct())
        {
            object value = structure.Values[structure.DateRow, col];

            if (!IsDayNumberHeader(value))
                return false;
        }

        return true;
    }

    private static bool IsDayNumberHeader(object value)
    {
        if (value == null)
            return false;

        if (value is double d)
            return d >= 1 && d <= 31 && Math.Abs(d - Math.Round(d)) < 0.001;

        string text = GetValueText(value);
        return int.TryParse(text, out int day) && day >= 1 && day <= 31;
    }

    private static bool TryReadDayHeader(object value, DateTime month, out DateTime date)
    {
        date = default;

        if (value == null)
            return false;

        if (value is double d)
        {
            if (d >= 1 && d <= 31 && Math.Abs(d - Math.Round(d)) < 0.001)
            {
                int day = (int)Math.Round(d);
                if (day <= DateTime.DaysInMonth(month.Year, month.Month))
                {
                    date = new DateTime(month.Year, month.Month, day);
                    return true;
                }
            }

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
        if (int.TryParse(text, out int intDay) && intDay >= 1 && intDay <= DateTime.DaysInMonth(month.Year, month.Month))
        {
            date = new DateTime(month.Year, month.Month, intDay);
            return true;
        }

        if (DateTime.TryParse(text, CultureInfo.GetCultureInfo("ru-RU"), DateTimeStyles.None, out DateTime parsed))
        {
            date = parsed.Date;
            return true;
        }

        return false;
    }

    private static bool HasEmployeeLikeRows(object[,] values, int usedRows, int fioCol, int firstRow)
    {
        int found = 0;

        for (int row = Math.Max(1, firstRow); row <= Math.Min(usedRows, firstRow + 80); row++)
        {
            if (LooksLikePersonName(GetValueText(values[row, fioCol])))
                found++;

            if (found >= 1)
                return true;
        }

        return false;
    }

    private static bool LooksLikePersonName(string value)
    {
        string normalized = NormalizeFio(value);
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

    private static int? FindTotalColumn(object[,] values, int row, int usedCols)
    {
        if (row <= 0)
            return null;

        for (int col = 1; col <= usedCols; col++)
        {
            string header = NormalizeHeader(GetValueText(values[row, col]));
            if (header.Contains("итог") || header.Contains("всего"))
                return col;
        }

        return null;
    }

    private static int? FindTotalColumnAfterLastMonthDate(
        Dictionary<DateTime, int> dateColumns,
        DateTime month)
    {
        var monthDateColumns = dateColumns
            .Where(x => x.Key.Year == month.Year && x.Key.Month == month.Month)
            .Select(x => x.Value)
            .ToList();

        if (monthDateColumns.Count == 0)
            return null;

        return monthDateColumns.Max() + 1;
    }

    private static HashSet<int> MarkInvalidDeclaredShiftHours(
        dynamic sheet,
        TimesheetStructure structure,
        TimesheetEmployee employee,
        DateTime month,
        int daysInMonth,
        ProcessingProtocol protocol)
    {
        var invalidColumns = new HashSet<int>();

        foreach (var dateColumn in structure.DateColumns.OrderBy(x => x.Key))
        {
            DateTime date = dateColumn.Key.Date;
            int col = dateColumn.Value;

            if (date.Month != month.Month || date.Year != month.Year || date.Day > daysInMonth)
                continue;

            if (!TryGetDeclaredShiftHours(structure, employee.RowIndex, col, out double declaredHours)
                || declaredHours <= MaxDeclaredShiftHours)
                continue;

            invalidColumns.Add(col);
            dynamic cell = sheet.Cells[employee.RowIndex, col];
            ClearCellContentsSafe(cell);
            try { cell.Interior.Color = RGB(255, 199, 206); } catch { }

            string message =
                $"Заявленная смена {FormatHours(declaredHours)} ч превышает допустимый максимум " +
                $"{FormatHours(MaxDeclaredShiftHours)} ч. Требуется ручная проверка.";
            AddCommentSafe(cell, message);
            protocol.Filling.Add(
                $"{employee.OriginalName}; {date:dd.MM.yyyy}; заявлено={FormatHours(declaredHours)}; часы=; " +
                $"статус=Превышен максимум заявленной смены {FormatHours(MaxDeclaredShiftHours)} ч — ручная проверка");
        }

        return invalidColumns;
    }

    private static void FinalizeEmployeeTotal(
        dynamic sheet,
        int row,
        TimesheetStructure structure,
        HashSet<int> invalidDeclaredColumns,
        ProcessingProtocol protocol)
    {
        if (!structure.TotalHoursCol.HasValue)
            return;

        if (invalidDeclaredColumns.Count == 0)
        {
            RecalculateTotal(sheet, row, structure, protocol);
            return;
        }

        dynamic totalCell = sheet.Cells[row, structure.TotalHoursCol.Value];
        ClearCellContentsSafe(totalCell);
        try { totalCell.Interior.Color = RGB(255, 199, 206); } catch { }
        AddCommentSafe(
            totalCell,
            $"Итог не рассчитан: найдены смены больше {FormatHours(MaxDeclaredShiftHours)} ч. Требуется ручная проверка.");
        protocol.Errors.Add(
            $"Табель; строка {row}; итог не рассчитан; найдены смены больше {FormatHours(MaxDeclaredShiftHours)} ч");
    }

    private static void RecalculateTotal(dynamic sheet, int row, TimesheetStructure structure, ProcessingProtocol protocol)
    {
        try
        {
            int totalColumn = structure.TotalHoursCol!.Value;
            if (structure.DateColumns.Values.Contains(totalColumn))
            {
                protocol.Errors.Add($"Табель; строка {row}; итоговая колонка совпала с дневной колонкой; пересчёт итога пропущен");
                return;
            }

            double total = 0;

            foreach (int col in structure.DateColumns.Values.Distinct())
            {
                object value = sheet.Cells[row, col].Value2;
                if (TryConvertToDouble(value, out double hours))
                    total += hours;
            }

            dynamic totalCell = sheet.Cells[row, totalColumn];

            try
            {
                if (totalCell.HasFormula == true)
                {
                    try { totalCell.Calculate(); } catch { }
                    return;
                }
            }
            catch { }

            totalCell.Value2 = total;
            SetNumberFormatSafe(totalCell);
        }
        catch (Exception ex)
        {
            protocol.Errors.Add($"Табель; строка {row}; не удалось пересчитать итоговую колонку; {ex.Message}");
        }
    }

    private static void EnsureTotalColumnFormatting(
        dynamic sheet,
        TimesheetStructure structure,
        ProcessingProtocol protocol)
    {
        if (!structure.TotalHoursCol.HasValue || structure.DateColumns.Count == 0)
            return;

        int totalColumn = structure.TotalHoursCol.Value;
        int lastDayColumn = structure.DateColumns.Values.Max();

        if (structure.DateColumns.Values.Contains(totalColumn))
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

                try { totalCell.Interior.Pattern = -4142; } catch { } // xlNone
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
            protocol.Errors.Add($"Табель; не удалось оформить итоговую колонку; {ex.Message}");
        }
    }

    private static bool IsAllowedSkudMessage(string message)
    {
        string text = NormalizeMessage(message);

        if (string.IsNullOrWhiteSpace(text))
            return false;

        if (text.Contains("запрещ") || text.Contains("заблок") || text.Contains("несоглас") || text.Contains("несовпад"))
            return false;

        return text.Contains("вход разрешен")
            || text.Contains("выход разрешен")
            || text.Contains("проход разрешен")
            || text.Contains("доступ разрешен");
    }

    private static bool IsProblemSkudMessage(string message)
    {
        string text = NormalizeMessage(message);

        if (string.IsNullOrWhiteSpace(text))
            return false;

        return text.Contains("запрещ")
            || text.Contains("заблок")
            || text.Contains("несоглас")
            || text.Contains("несовпад")
            || text.Contains("повтор")
            || text.Contains("ошиб")
            || text.Contains("отказ")
            || text.Contains("наруш")
            || text.Contains("пропуск")
            || text.Contains("заявк")
            || text.Contains("служеб")
            || text.Contains("неизвест")
            || text.Contains("недейств")
            || text.Contains("истек")
            || text.Contains("аннулир")
            || text.Contains("закрыт");
    }

    private static string NormalizeMessage(string text)
    {
        if (text == null)
            return string.Empty;

        text = text.Replace('\u00A0', ' ').Trim().ToLowerInvariant().Replace('ё', 'е');
        return Regex.Replace(text, @"\s+", " ");
    }

    private static string NormalizeFio(string text)
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

    private static string BuildPersonSearchKey(string fio)
    {
        string normalized = NormalizeFio(fio);
        var parts = normalized.Split(' ', StringSplitOptions.RemoveEmptyEntries);

        if (parts.Length < 2)
            return string.Empty;

        return parts[0] + " " + parts[1][0];
    }

    private static string FirstWord(string text)
    {
        text = NormalizeFio(text);
        int spaceIndex = text.IndexOf(' ');
        return spaceIndex < 0 ? text : text[..spaceIndex];
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

            return DateTime.TryParseExact(text, formats, ru, DateTimeStyles.None, out result)
                || DateTime.TryParse(text, ru, DateTimeStyles.None, out result);
        }
        catch
        {
            return false;
        }
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
        return double.TryParse(text, NumberStyles.Any, CultureInfo.CurrentCulture, out result)
            || double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out result);
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

    private static void SetNumberFormatSafe(dynamic cell)
    {
        try { cell.NumberFormatLocal = "0"; return; } catch { }
        try { cell.NumberFormat = "0"; } catch { }
    }

    private static bool TryGetDeclaredShiftHours(TimesheetStructure structure, int row, int col, out double hours)
    {
        hours = 0;

        try
        {
            object value = structure.Values[row, col];

            if (!TryConvertToDouble(value, out double parsedHours))
                return false;

            if (double.IsNaN(parsedHours) || double.IsInfinity(parsedHours) || parsedHours <= 0)
                return false;

            hours = parsedHours;
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static void MarkDeclaredShiftWithoutAllowedSkudEvents(dynamic cell, double declaredHours)
    {
        ClearCellContentsSafe(cell);
        try { cell.Interior.Color = RGB(255, 199, 206); } catch { }
        AddCommentSafe(
            cell,
            $"Смена заявлена на {FormatHours(declaredHours)} часов, но разрешённые отметки в СКУД отсутствуют");
    }

    private static void MarkDeclaredShiftWithoutEmployeeInSkud(dynamic cell)
    {
        ClearCellContentsSafe(cell);
        try { cell.Interior.Color = RGB(255, 199, 206); } catch { }
        AddCommentSafe(cell, "Смена заявлена, сотрудник не найден в данных СКУД");
    }

    private static void ClearCellContentsSafe(dynamic cell)
    {
        try { cell.ClearContents(); return; } catch { }
        try { cell.Value2 = null; } catch { }
    }

    private static string FormatHours(double hours)
    {
        return hours.ToString("0.##", CultureInfo.CurrentCulture);
    }

    private static void HighlightCellAmbiguous(dynamic cell, string searchKey, List<string> variants)
    {
        try { cell.Interior.Color = RGB(255, 165, 0); } catch { }
        AddCommentSafe(cell, $"Неоднозначное сопоставление ФИО.\nКлюч поиска: {searchKey}\nНайденные варианты: {string.Join(", ", variants)}");
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

    private static void WriteProtocol(string path, ProcessingProtocol protocol)
    {
        var sb = new StringBuilder();
        AppendSection(sb, "Общее", protocol.General);
        AppendSection(sb, "Протокол сопоставления", protocol.Matching);
        AppendSection(sb, "Протокол заполнения", protocol.Filling);
        AppendSection(sb, "Протокол ошибок", protocol.Errors);
        File.WriteAllText(path, sb.ToString(), Encoding.UTF8);
    }

    private static void AppendSection(StringBuilder sb, string title, List<string> lines)
    {
        sb.AppendLine($"# {title}");
        if (lines.Count == 0)
            sb.AppendLine("Нет записей.");
        else
            foreach (string line in lines)
                sb.AppendLine(line);

        sb.AppendLine();
    }

    private static string FormatTime(DateTime? value)
    {
        return value.HasValue ? value.Value.ToString("HH:mm", CultureInfo.InvariantCulture) : "";
    }

    private static string ColumnName(int column)
    {
        string name = string.Empty;
        while (column > 0)
        {
            int modulo = (column - 1) % 26;
            name = Convert.ToChar('A' + modulo) + name;
            column = (column - modulo) / 26;
        }

        return name;
    }

    private static int ScalePercent(int current, int total, int fromPercent, int toPercent)
    {
        if (total <= 0)
            return fromPercent;

        double k = Math.Clamp((double)current / total, 0, 1);
        return fromPercent + (int)Math.Round((toPercent - fromPercent) * k);
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

    private sealed class SkudEvent
    {
        public string OriginalPerson { get; set; } = string.Empty;
        public DateTime DateTime { get; set; }
        public string Message { get; set; } = string.Empty;
        public int SourceRowIndex { get; set; }
    }

    private sealed class DayEvents
    {
        private readonly List<SkudEvent> _allowedEvents = new();
        private readonly List<SkudEvent> _problemEvents = new();

        public bool HasAnyEvents => _allowedEvents.Count > 0 || _problemEvents.Count > 0;

        public void AddAllowedEvent(SkudEvent ev)
        {
            _allowedEvents.Add(ev);
            _allowedEvents.Sort((a, b) => a.DateTime.CompareTo(b.DateTime));
        }

        public void AddProblemEvent(SkudEvent ev)
        {
            _problemEvents.Add(ev);
            _problemEvents.Sort((a, b) => a.DateTime.CompareTo(b.DateTime));
        }

        public DayCalculationResult Calculate()
        {
            var result = new DayCalculationResult
            {
                EventCount = _allowedEvents.Count,
                ProblemEventCount = _problemEvents.Count,
                ProblemMessages = _problemEvents
                    .Select(x => $"{x.DateTime:HH:mm} {x.Message}")
                    .Distinct()
                    .ToList()
            };

            if (_allowedEvents.Count == 0)
            {
                result.Status = _problemEvents.Count > 0
                    ? "Нет разрешенных записей СКУД за день. Проблемные сообщения добавлены только в протокол."
                    : string.Empty;
                return result;
            }

            result.FirstEvent = _allowedEvents.First().DateTime;
            result.LastEvent = _allowedEvents.Last().DateTime;

            if (_allowedEvents.Count == 1)
            {
                result.Hours = 2;
                result.ShouldWriteValue = true;
                result.IsSingleAllowedEvent = true;
                result.Status = "Одна запись СКУД за день. Установлено 2 часа.";
                return result;
            }

            TimeSpan duration = result.LastEvent.Value - result.FirstEvent.Value - TimeSpan.FromHours(1);

            if (duration < TimeSpan.Zero)
                duration = TimeSpan.Zero;

            int wholeHours = (int)Math.Floor(duration.TotalHours);
            int roundedHours = duration.Minutes >= 50 ? wholeHours + 1 : wholeHours;

            if (roundedHours > MaxDeclaredShiftHours)
            {
                result.RequiresManualReview = true;
                result.Status =
                    $"По отметкам СКУД рассчитано {roundedHours} ч, что превышает допустимый максимум " +
                    $"{FormatHours(MaxDeclaredShiftHours)} ч. Требуется ручная проверка.";
                return result;
            }

            result.Hours = roundedHours;
            result.ShouldWriteValue = true;
            result.Status = "Рассчитано по первому и последнему событию дня.";
            return result;
        }
    }

    private sealed class DayCalculationResult
    {
        public bool ShouldWriteValue { get; set; }
        public double? Hours { get; set; }
        public int EventCount { get; set; }
        public int ProblemEventCount { get; set; }
        public DateTime? FirstEvent { get; set; }
        public DateTime? LastEvent { get; set; }
        public bool IsSingleAllowedEvent { get; set; }
        public bool RequiresManualReview { get; set; }
        public List<string> ProblemMessages { get; set; } = new();
        public string Status { get; set; } = string.Empty;
    }

    private sealed record SkudReadResult(
        Dictionary<string, Dictionary<DateTime, DayEvents>> EventsByPerson,
        DateTime MainMonth,
        IReadOnlyList<DateTime> Months);

    private sealed record TimesheetStructure(
        int HeaderRow,
        int DateRow,
        int FirstEmployeeRow,
        int FioCol,
        Dictionary<DateTime, int> DateColumns,
        int? TotalHoursCol,
        object[,] Values,
        int UsedRows);

    private sealed record TimesheetEmployee(int RowIndex, string OriginalName);

    private sealed record FioParts(string Surname, string FirstName, string Patronymic);

    private sealed class ProcessingProtocol
    {
        public List<string> General { get; } = new();
        public List<string> Matching { get; } = new();
        public List<string> Filling { get; } = new();
        public List<string> Errors { get; } = new();
    }

    private sealed class FillResult
    {
        public int TotalEmployees { get; set; }
        public int MatchedEmployees { get; set; }
        public int NotFoundEmployees { get; set; }
        public int AmbiguousEmployees { get; set; }
        public int FilledCells { get; set; }
    }
}
