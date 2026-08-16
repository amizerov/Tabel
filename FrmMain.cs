namespace Tabel;

using System.Reflection;

public partial class FrmMain : Form
{
    private enum SelectedFileRole
    {
        OriginalTimesheet,
        SkudExport,
        FilledTimesheet
    }

    private string? _fillTimesheetPath;
    private string? _skudPath;
    private string? _compareOriginalPath;
    private string? _compareFilledPath;
    private string? _lastFillerSummary;
    private string? _lastComparatorSummary;
    private bool _isBusy;

    public FrmMain()
    {
        InitializeComponent();
        Text = $"Табель СКУД — версия {GetAppVersion()}";

        SkudExcelInteropFiller.ProgressChanged += SkudExcelInteropFiller_ProgressChanged;
        SkudTimesheetComparator.ProgressChanged += SkudTimesheetComparator_ProgressChanged;

        UpdateFillResultPreview();
        UpdateCompareResultPreview();
        UpdateActionState();
        ShowFillStep();
    }

    protected override void OnFormClosed(FormClosedEventArgs e)
    {
        SkudExcelInteropFiller.ProgressChanged -= SkudExcelInteropFiller_ProgressChanged;
        SkudTimesheetComparator.ProgressChanged -= SkudTimesheetComparator_ProgressChanged;
        base.OnFormClosed(e);
    }

    private static string GetAppVersion() =>
        Assembly.GetExecutingAssembly().GetName().Version?.ToString(3) ?? "неизвестна";

    private async void btnSelectFillTimesheet_Click(object sender, EventArgs e)
    {
        string? path = SelectExcelFile("Выберите исходный табель подрядчика");
        if (path == null) return;

        _fillTimesheetPath = null;
        txtFillTimesheet.Clear();
        UpdateFillResultPreview();
        UpdateActionState();

        if (!await ValidateSelectedFileAsync(path, SelectedFileRole.OriginalTimesheet, "Исходный табель подрядчика"))
            return;

        _fillTimesheetPath = path;
        txtFillTimesheet.Text = path;
        UpdateFillResultPreview();
        UpdateActionState();
    }

    private async void btnSelectSkud_Click(object sender, EventArgs e)
    {
        string? path = SelectExcelFile("Выберите выгрузку СКУД");
        if (path == null) return;

        _skudPath = null;
        txtSkud.Clear();
        UpdateActionState();

        if (PathsAreEqual(path, _fillTimesheetPath))
        {
            ShowWrongFileMessage("Выгрузка СКУД", "Нельзя выбрать исходный табель одновременно как выгрузку СКУД.");
            return;
        }

        if (!await ValidateSelectedFileAsync(path, SelectedFileRole.SkudExport, "Выгрузка СКУД"))
            return;

        _skudPath = path;
        txtSkud.Text = path;
        UpdateActionState();
    }

    private async void btnSelectCompareOriginal_Click(object sender, EventArgs e)
    {
        string? path = SelectExcelFile("Выберите исходный табель подрядчика");
        if (path == null) return;

        _compareOriginalPath = null;
        txtCompareOriginal.Clear();
        UpdateCompareResultPreview();
        UpdateActionState();

        if (PathsAreEqual(path, _compareFilledPath))
        {
            ShowWrongFileMessage("Исходный табель подрядчика", "Исходный и заполненный табели должны быть разными файлами.");
            return;
        }

        if (!await ValidateSelectedFileAsync(path, SelectedFileRole.OriginalTimesheet, "Исходный табель подрядчика"))
            return;

        _compareOriginalPath = path;
        txtCompareOriginal.Text = path;
        UpdateCompareResultPreview();
        UpdateActionState();
    }

    private async void btnSelectCompareFilled_Click(object sender, EventArgs e)
    {
        string? path = SelectExcelFile("Выберите заполненный табель");
        if (path == null) return;

        _compareFilledPath = null;
        txtCompareFilled.Clear();
        UpdateCompareResultPreview();
        UpdateActionState();

        if (PathsAreEqual(path, _compareOriginalPath))
        {
            ShowWrongFileMessage("Заполненный табель", "Исходный и заполненный табели должны быть разными файлами.");
            return;
        }

        if (!await ValidateSelectedFileAsync(path, SelectedFileRole.FilledTimesheet, "Заполненный табель"))
            return;

        _compareFilledPath = path;
        txtCompareFilled.Text = path;
        UpdateCompareResultPreview();
        UpdateActionState();
    }

    private static string? SelectExcelFile(string title)
    {
        using var dialog = new OpenFileDialog
        {
            Title = title,
            Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*",
            RestoreDirectory = true,
            CheckFileExists = true
        };

        return dialog.ShowDialog() == DialogResult.OK ? dialog.FileName : null;
    }

    private async Task<bool> ValidateSelectedFileAsync(string path, SelectedFileRole role, string fieldName)
    {
        SetBusy(true);
        progressMain.Style = ProgressBarStyle.Marquee;
        progressMain.MarqueeAnimationSpeed = 25;
        lblProgressStatus.Text = $"Проверяю файл для поля «{fieldName}»...";
        AddLog($"Проверка файла ({fieldName}): {path}");

        try
        {
            if (role == SelectedFileRole.SkudExport)
                await SkudExcelInteropFiller.ValidateSkudInputAsync(path);
            else if (role == SelectedFileRole.FilledTimesheet)
                await SkudExcelInteropFiller.ValidateFilledTimesheetInputAsync(path);
            else
                await SkudExcelInteropFiller.ValidateOriginalTimesheetInputAsync(path);

            lblProgressStatus.Text = $"Файл для поля «{fieldName}» проверен.";
            AddLog("Файл принят.");
            return true;
        }
        catch (Exception ex)
        {
            string message = $"Выбранный файл не подходит для поля «{fieldName}»." +
                Environment.NewLine + Environment.NewLine + ex.Message;
            AddLog("Файл отклонён: " + ex.Message);
            lblProgressStatus.Text = "Выбран неправильный файл.";
            MessageBox.Show(message, "Неправильный файл", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return false;
        }
        finally
        {
            progressMain.MarqueeAnimationSpeed = 0;
            progressMain.Style = ProgressBarStyle.Blocks;
            progressMain.Value = 0;
            SetBusy(false);
        }
    }

    private void ShowWrongFileMessage(string fieldName, string reason)
    {
        lblProgressStatus.Text = "Выбран неправильный файл.";
        AddLog($"Файл для поля «{fieldName}» отклонён: {reason}");
        MessageBox.Show(
            $"Выбранный файл не подходит для поля «{fieldName}»." + Environment.NewLine + Environment.NewLine + reason,
            "Неправильный файл",
            MessageBoxButtons.OK,
            MessageBoxIcon.Warning);
    }

    private async void btnProcessTimesheet_Click(object sender, EventArgs e)
    {
        if (string.IsNullOrWhiteSpace(_fillTimesheetPath) || string.IsNullOrWhiteSpace(_skudPath))
            return;

        string timesheetPath = _fillTimesheetPath;
        string skudPath = _skudPath;
        SetBusy(true);
        _lastFillerSummary = null;
        BeginOperation("Начинаю заполнение табеля по СКУД.");
        AddLog("Табель: " + timesheetPath);
        AddLog("СКУД: " + skudPath);

        try
        {
            string result = await SkudExcelInteropFiller.FillTimesheetAsync(timesheetPath, skudPath);
            PrepareComparison(timesheetPath, result);
            ShowSuccess(result, _lastFillerSummary ?? "Табель успешно заполнен.", "Файл результата:");
        }
        catch (Exception ex)
        {
            ShowOperationError(ex);
        }
        finally
        {
            SetBusy(false);
        }
    }

    private async void btnCompareTimesheets_Click(object sender, EventArgs e)
    {
        if (string.IsNullOrWhiteSpace(_compareOriginalPath)
            || string.IsNullOrWhiteSpace(_compareFilledPath)
            || PathsAreEqual(_compareOriginalPath, _compareFilledPath))
            return;

        string originalPath = _compareOriginalPath;
        string filledPath = _compareFilledPath;
        SetBusy(true);
        _lastComparatorSummary = null;
        BeginOperation("Начинаю сравнение табелей.");
        AddLog("Исходный табель: " + originalPath);
        AddLog("Заполненный табель: " + filledPath);

        try
        {
            string result = await SkudTimesheetComparator.CompareTimesheetsAsync(originalPath, filledPath);
            ShowSuccess(result, _lastComparatorSummary ?? "Сравнение табелей завершено.", "Файл сравнения:");
        }
        catch (Exception ex)
        {
            ShowOperationError(ex);
        }
        finally
        {
            SetBusy(false);
        }
    }

    private void ShowSuccess(string result, string summary, string resultCaption)
    {
        AddLog("");
        AddLog("Готово.");
        AddLog(resultCaption);
        AddLog(result);

        string dialogText = summary.Contains(result, StringComparison.OrdinalIgnoreCase)
            ? summary
            : summary + Environment.NewLine + result;

        MessageBox.Show(dialogText, "Готово", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private void BeginOperation(string message)
    {
        txtLog.Clear();
        progressMain.Value = 0;
        lblProgressStatus.Text = message;
        AddLog(message);
    }

    private void ShowOperationError(Exception ex)
    {
        AddLog("");
        AddLog("ОШИБКА:");
        AddLog(ex.Message);
        SetProgress(0, "Операция завершилась с ошибкой.");
        MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
    }

    private void SetBusy(bool isBusy)
    {
        _isBusy = isBusy;
        pnlWizard.Enabled = !isBusy;
        UpdateActionState();
    }

    private void UpdateActionState()
    {
        btnProcessTimesheet.Enabled = !_isBusy
            && !string.IsNullOrWhiteSpace(_fillTimesheetPath)
            && !string.IsNullOrWhiteSpace(_skudPath);

        btnCompareTimesheets.Enabled = !_isBusy
            && !string.IsNullOrWhiteSpace(_compareOriginalPath)
            && !string.IsNullOrWhiteSpace(_compareFilledPath)
            && !PathsAreEqual(_compareOriginalPath, _compareFilledPath);
    }

    private void UpdateFillResultPreview()
    {
        lblFillResultPath.Text = BuildResultPath(_fillTimesheetPath, "_заполнен.xlsx")
            ?? "Будет показан после выбора исходного табеля.";
    }

    private void PrepareComparison(string originalPath, string filledPath)
    {
        _compareOriginalPath = originalPath;
        _compareFilledPath = filledPath;
        txtCompareOriginal.Text = originalPath;
        txtCompareFilled.Text = filledPath;

        UpdateCompareResultPreview();
        UpdateActionState();
        btnGoToCompare.Visible = true;
        ShowComparisonStep();
    }

    private void btnBackToFill_Click(object sender, EventArgs e)
    {
        ShowFillStep();
    }

    private void btnGoToCompare_Click(object sender, EventArgs e)
    {
        ShowComparisonStep();
    }

    private void ShowFillStep()
    {
        grpCompareStep.Visible = false;
        grpFillStep.Visible = true;
        grpFillStep.BringToFront();
    }

    private void ShowComparisonStep()
    {
        grpFillStep.Visible = false;
        grpCompareStep.Visible = true;
        grpCompareStep.BringToFront();
    }

    private void UpdateCompareResultPreview()
    {
        if (PathsAreEqual(_compareOriginalPath, _compareFilledPath))
        {
            lblCompareResultPath.ForeColor = Color.Firebrick;
            lblCompareResultPath.Text = "Исходный и заполненный табели должны быть разными файлами.";
            return;
        }

        lblCompareResultPath.ForeColor = SystemColors.GrayText;
        lblCompareResultPath.Text = BuildResultPath(_compareOriginalPath, "_сравнение.xlsx")
            ?? "Будет показан после выбора исходного табеля.";
    }

    private static string? BuildResultPath(string? sourcePath, string suffix)
    {
        if (string.IsNullOrWhiteSpace(sourcePath)) return null;

        return Path.Combine(
            Path.GetDirectoryName(sourcePath) ?? string.Empty,
            Path.GetFileNameWithoutExtension(sourcePath) + suffix);
    }

    private static bool PathsAreEqual(string? firstPath, string? secondPath)
    {
        if (string.IsNullOrWhiteSpace(firstPath) || string.IsNullOrWhiteSpace(secondPath))
            return false;

        try
        {
            return string.Equals(Path.GetFullPath(firstPath), Path.GetFullPath(secondPath), StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return string.Equals(firstPath, secondPath, StringComparison.OrdinalIgnoreCase);
        }
    }

    private void SkudExcelInteropFiller_ProgressChanged(object? sender, SkudProgressEventArgs e)
    {
        HandleProgress(e);
        if (e.IsFinished) _lastFillerSummary = e.Message;
    }

    private void SkudTimesheetComparator_ProgressChanged(object? sender, SkudProgressEventArgs e)
    {
        HandleProgress(e);
        if (e.IsFinished) _lastComparatorSummary = e.Message;
    }

    private void HandleProgress(SkudProgressEventArgs e)
    {
        string text = e.Total > 0
            ? $"{DateTime.Now:HH:mm:ss} | {e.Stage}: {e.Message} ({e.Percent}%)"
            : $"{DateTime.Now:HH:mm:ss} | {e.Stage}: {e.Message}";

        AddLog(text);
        SetProgress(e.Percent, $"{e.Stage}: {e.Message}");
    }

    private void SetProgress(int percent, string status)
    {
        if (InvokeRequired)
        {
            BeginInvoke(new Action(() => SetProgress(percent, status)));
            return;
        }

        progressMain.Value = Math.Clamp(percent, progressMain.Minimum, progressMain.Maximum);
        lblProgressStatus.Text = status;
    }

    private void AddLog(string text)
    {
        if (txtLog.InvokeRequired)
        {
            txtLog.BeginInvoke(new Action(() => AddLog(text)));
            return;
        }

        txtLog.AppendText(text + Environment.NewLine);
        txtLog.SelectionStart = txtLog.TextLength;
        txtLog.ScrollToCaret();
    }
}
