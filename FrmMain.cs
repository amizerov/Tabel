namespace Tabel;

public partial class FrmMain : Form
{
    private string? _timesheetPath;
    private string? _skudPath;
    private string? _lastFillerSummary;
    private string? _lastComparatorSummary;

    public FrmMain()
    {
        InitializeComponent();

        SkudExcelInteropFiller.ProgressChanged += SkudExcelInteropFiller_ProgressChanged;
        SkudTimesheetComparator.ProgressChanged += SkudTimesheetComparator_ProgressChanged;
    }

    protected override void OnFormClosed(FormClosedEventArgs e)
    {
        SkudExcelInteropFiller.ProgressChanged -= SkudExcelInteropFiller_ProgressChanged;
        SkudTimesheetComparator.ProgressChanged -= SkudTimesheetComparator_ProgressChanged;

        base.OnFormClosed(e);
    }

    private void btnSelectTabel_Click(object sender, EventArgs e)
    {
        using (var ofd = new OpenFileDialog())
        {
            ofd.Title = "Выберите табель подрядчика";
            ofd.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*";
            ofd.RestoreDirectory = true;

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                _timesheetPath = ofd.FileName;
                txtTabel.Text = _timesheetPath;
            }
        }
    }

    private void btnSelectSkud_Click(object sender, EventArgs e)
    {
        using (var ofd = new OpenFileDialog())
        {
            ofd.Title = "Выберите выгрузку СКУД";
            ofd.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*";
            ofd.RestoreDirectory = true;

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                _skudPath = ofd.FileName;
                txtSkud.Text = _skudPath;
            }
        }
    }

    private async void btnDoIt_Click(object sender, EventArgs e)
    {
        if (string.IsNullOrWhiteSpace(_timesheetPath))
        {
            MessageBox.Show("Сначала выберите табель подрядчика.");
            return;
        }

        if (string.IsNullOrWhiteSpace(_skudPath))
        {
            MessageBox.Show("Сначала выберите выгрузку СКУД.");
            return;
        }

        btnDoIt.Enabled = false;
        btnSelectTabel.Enabled = false;
        btnSelectSkud.Enabled = false;
        _lastFillerSummary = null;

        txtLog.Clear();

        AddLog("Начинаю заполнение табеля по СКУД.");
        AddLog("Табель: " + _timesheetPath);
        AddLog("СКУД: " + _skudPath);

        try
        {
            string result = await SkudExcelInteropFiller.FillTimesheetAsync(_timesheetPath, _skudPath);

            AddLog("");
            AddLog("Готово.");
            AddLog("Файл результата:");
            AddLog(result);

            MessageBox.Show(
                (_lastFillerSummary ?? "Табель успешно заполнен.") + Environment.NewLine + result,
                "Готово",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            AddLog("");
            AddLog("ОШИБКА:");
            AddLog(ex.Message);

            MessageBox.Show(
                ex.Message,
                "Ошибка",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
        finally
        {
            btnDoIt.Enabled = true;
            btnSelectTabel.Enabled = true;
            btnSelectSkud.Enabled = true;
        }
    }

    private async void btnCompare_Click(object sender, EventArgs e)
    {
        if (string.IsNullOrWhiteSpace(_timesheetPath))
        {
            MessageBox.Show("Сначала выберите исходный табель подрядчика.");
            return;
        }

        using (var ofd = new OpenFileDialog())
        {
            ofd.Title = "Выберите заполненный табель";
            ofd.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*";
            ofd.RestoreDirectory = true;

            if (ofd.ShowDialog() != DialogResult.OK)
                return;

            string filledPath = ofd.FileName;

            btnDoIt.Enabled = false;
            btnSelectTabel.Enabled = false;
            btnSelectSkud.Enabled = false;
            btnCompare.Enabled = false;
            _lastComparatorSummary = null;

            txtLog.Clear();

            AddLog("Начинаю сравнение табелей.");
            AddLog("Исходный табель: " + _timesheetPath);
            AddLog("Заполненный табель: " + filledPath);

            try
            {
                string result = await SkudTimesheetComparator.CompareTimesheetsAsync(_timesheetPath, filledPath);

                AddLog("");
                AddLog("Готово.");
                AddLog("Файл сравнения:");
                AddLog(result);

                MessageBox.Show(
                    (_lastComparatorSummary ?? "Сравнение табелей завершено.") + Environment.NewLine + result,
                    "Готово",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                AddLog("");
                AddLog("ОШИБКА:");
                AddLog(ex.Message);

                MessageBox.Show(
                    ex.Message,
                    "Ошибка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                btnDoIt.Enabled = true;
                btnSelectTabel.Enabled = true;
                btnSelectSkud.Enabled = true;
                btnCompare.Enabled = true;
            }
        }
    }

    private void SkudExcelInteropFiller_ProgressChanged(object? sender, SkudProgressEventArgs e)
    {
        string text;

        if (e.Total > 0)
        {
            text = $"{DateTime.Now:HH:mm:ss} | {e.Stage}: {e.Message} ({e.Percent}%)";
        }
        else
        {
            text = $"{DateTime.Now:HH:mm:ss} | {e.Stage}: {e.Message}";
        }

        AddLog(text);

        if (e.IsFinished)
            _lastFillerSummary = e.Message;
    }

    private void SkudTimesheetComparator_ProgressChanged(object? sender, SkudProgressEventArgs e)
    {
        string text;

        if (e.Total > 0)
        {
            text = $"{DateTime.Now:HH:mm:ss} | {e.Stage}: {e.Message} ({e.Percent}%)";
        }
        else
        {
            text = $"{DateTime.Now:HH:mm:ss} | {e.Stage}: {e.Message}";
        }

        AddLog(text);

        if (e.IsFinished)
            _lastComparatorSummary = e.Message;
    }

    private void AddLog(string text)
    {
        if (txtLog.InvokeRequired)
        {
            txtLog.BeginInvoke(new Action(() => AddLog(text)));
            return;
        }

        txtLog.AppendText(text + Environment.NewLine);

        // Автопрокрутка вниз
        txtLog.SelectionStart = txtLog.TextLength;
        txtLog.ScrollToCaret();
    }
}
