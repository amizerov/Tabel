namespace Tabel
{
    partial class FrmMain
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && components != null)
                components.Dispose();
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        private void InitializeComponent()
        {
            pnlWizard = new Panel();
            grpFillStep = new GroupBox();
            pnlFillHeader = new Panel();
            lblFillStepNumber = new Label();
            lblFillStepTitle = new Label();
            lblFillTimesheet = new Label();
            txtFillTimesheet = new TextBox();
            btnSelectFillTimesheet = new Button();
            lblSkud = new Label();
            txtSkud = new TextBox();
            btnSelectSkud = new Button();
            btnProcessTimesheet = new Button();
            lblFillResultCaption = new Label();
            lblFillResultPath = new Label();
            btnGoToCompare = new Button();
            grpCompareStep = new GroupBox();
            pnlCompareHeader = new Panel();
            lblCompareStepNumber = new Label();
            lblCompareStepTitle = new Label();
            lblCompareOriginal = new Label();
            txtCompareOriginal = new TextBox();
            btnSelectCompareOriginal = new Button();
            lblCompareFilled = new Label();
            txtCompareFilled = new TextBox();
            btnSelectCompareFilled = new Button();
            btnCompareTimesheets = new Button();
            lblCompareResultCaption = new Label();
            lblCompareResultPath = new Label();
            btnBackToFill = new Button();
            grpProgress = new GroupBox();
            progressMain = new ProgressBar();
            lblProgressStatus = new Label();
            txtLog = new TextBox();
            pnlWizard.SuspendLayout();
            grpFillStep.SuspendLayout();
            pnlFillHeader.SuspendLayout();
            grpCompareStep.SuspendLayout();
            pnlCompareHeader.SuspendLayout();
            grpProgress.SuspendLayout();
            SuspendLayout();
            // 
            // pnlWizard
            // 
            pnlWizard.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            pnlWizard.Controls.Add(grpCompareStep);
            pnlWizard.Controls.Add(grpFillStep);
            pnlWizard.Location = new Point(12, 4);
            pnlWizard.Name = "pnlWizard";
            pnlWizard.Size = new Size(816, 292);
            pnlWizard.TabIndex = 0;
            // 
            // grpFillStep
            // 
            grpFillStep.Controls.Add(lblFillTimesheet);
            grpFillStep.Controls.Add(txtFillTimesheet);
            grpFillStep.Controls.Add(btnSelectFillTimesheet);
            grpFillStep.Controls.Add(lblSkud);
            grpFillStep.Controls.Add(txtSkud);
            grpFillStep.Controls.Add(btnSelectSkud);
            grpFillStep.Controls.Add(btnProcessTimesheet);
            grpFillStep.Controls.Add(lblFillResultCaption);
            grpFillStep.Controls.Add(lblFillResultPath);
            grpFillStep.Controls.Add(pnlFillHeader);
            grpFillStep.Dock = DockStyle.Fill;
            grpFillStep.Name = "grpFillStep";
            grpFillStep.Padding = new Padding(16, 0, 16, 0);
            grpFillStep.Size = new Size(816, 292);
            grpFillStep.TabIndex = 0;
            grpFillStep.TabStop = false;
            grpFillStep.Text = "";
            // 
            // pnlFillHeader
            // 
            pnlFillHeader.BackColor = Color.FromArgb(0, 153, 219);
            pnlFillHeader.Controls.Add(lblFillStepNumber);
            pnlFillHeader.Controls.Add(lblFillStepTitle);
            pnlFillHeader.Controls.Add(btnGoToCompare);
            pnlFillHeader.Dock = DockStyle.Top;
            pnlFillHeader.Location = new Point(16, 16);
            pnlFillHeader.Name = "pnlFillHeader";
            pnlFillHeader.Size = new Size(784, 68);
            pnlFillHeader.TabIndex = 0;
            // 
            // lblFillStepNumber
            // 
            lblFillStepNumber.AutoSize = true;
            lblFillStepNumber.ForeColor = Color.White;
            lblFillStepNumber.Location = new Point(16, 9);
            lblFillStepNumber.Name = "lblFillStepNumber";
            lblFillStepNumber.Size = new Size(73, 15);
            lblFillStepNumber.TabIndex = 0;
            lblFillStepNumber.Text = "Шаг 1 из 2";
            // 
            // lblFillStepTitle
            // 
            lblFillStepTitle.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            lblFillStepTitle.AutoEllipsis = true;
            lblFillStepTitle.Font = new Font("Segoe UI", 17F, FontStyle.Bold);
            lblFillStepTitle.ForeColor = Color.White;
            lblFillStepTitle.Location = new Point(14, 25);
            lblFillStepTitle.Name = "lblFillStepTitle";
            lblFillStepTitle.Size = new Size(600, 34);
            lblFillStepTitle.TabIndex = 1;
            lblFillStepTitle.Text = "Заполнение табеля по СКУД";
            // 
            // lblFillTimesheet
            // 
            lblFillTimesheet.AutoSize = true;
            lblFillTimesheet.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            lblFillTimesheet.Location = new Point(20, 111);
            lblFillTimesheet.Name = "lblFillTimesheet";
            lblFillTimesheet.Size = new Size(207, 15);
            lblFillTimesheet.TabIndex = 0;
            lblFillTimesheet.Text = "1. Исходный табель подрядчика";
            // 
            // txtFillTimesheet
            // 
            txtFillTimesheet.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            txtFillTimesheet.Location = new Point(20, 132);
            txtFillTimesheet.Name = "txtFillTimesheet";
            txtFillTimesheet.PlaceholderText = "Выберите исходный табель подрядчика";
            txtFillTimesheet.ReadOnly = true;
            txtFillTimesheet.Size = new Size(642, 23);
            txtFillTimesheet.TabIndex = 1;
            // 
            // btnSelectFillTimesheet
            // 
            btnSelectFillTimesheet.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btnSelectFillTimesheet.Location = new Point(674, 131);
            btnSelectFillTimesheet.Name = "btnSelectFillTimesheet";
            btnSelectFillTimesheet.Size = new Size(112, 25);
            btnSelectFillTimesheet.TabIndex = 2;
            btnSelectFillTimesheet.Text = "Выбрать...";
            btnSelectFillTimesheet.UseVisualStyleBackColor = true;
            btnSelectFillTimesheet.Click += btnSelectFillTimesheet_Click;
            // 
            // lblSkud
            // 
            lblSkud.AutoSize = true;
            lblSkud.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            lblSkud.Location = new Point(20, 171);
            lblSkud.Name = "lblSkud";
            lblSkud.Size = new Size(117, 15);
            lblSkud.TabIndex = 3;
            lblSkud.Text = "2. Выгрузка СКУД";
            // 
            // txtSkud
            // 
            txtSkud.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            txtSkud.Location = new Point(20, 192);
            txtSkud.Name = "txtSkud";
            txtSkud.PlaceholderText = "Выберите выгрузку СКУД";
            txtSkud.ReadOnly = true;
            txtSkud.Size = new Size(642, 23);
            txtSkud.TabIndex = 4;
            // 
            // btnSelectSkud
            // 
            btnSelectSkud.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btnSelectSkud.Location = new Point(674, 191);
            btnSelectSkud.Name = "btnSelectSkud";
            btnSelectSkud.Size = new Size(112, 25);
            btnSelectSkud.TabIndex = 5;
            btnSelectSkud.Text = "Выбрать...";
            btnSelectSkud.UseVisualStyleBackColor = true;
            btnSelectSkud.Click += btnSelectSkud_Click;
            // 
            // btnProcessTimesheet
            // 
            btnProcessTimesheet.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btnProcessTimesheet.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            btnProcessTimesheet.Location = new Point(656, 240);
            btnProcessTimesheet.Name = "btnProcessTimesheet";
            btnProcessTimesheet.Size = new Size(130, 30);
            btnProcessTimesheet.TabIndex = 6;
            btnProcessTimesheet.Text = "Обработать";
            btnProcessTimesheet.UseVisualStyleBackColor = true;
            btnProcessTimesheet.Click += btnProcessTimesheet_Click;
            // 
            // lblFillResultCaption
            // 
            lblFillResultCaption.AutoSize = true;
            lblFillResultCaption.Location = new Point(20, 240);
            lblFillResultCaption.Name = "lblFillResultCaption";
            lblFillResultCaption.Size = new Size(159, 15);
            lblFillResultCaption.TabIndex = 7;
            lblFillResultCaption.Text = "Результат будет сохранён:";
            // 
            // lblFillResultPath
            // 
            lblFillResultPath.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            lblFillResultPath.AutoEllipsis = true;
            lblFillResultPath.ForeColor = SystemColors.GrayText;
            lblFillResultPath.Location = new Point(20, 264);
            lblFillResultPath.Name = "lblFillResultPath";
            lblFillResultPath.Size = new Size(620, 23);
            lblFillResultPath.TabIndex = 8;
            // 
            // btnGoToCompare
            // 
            btnGoToCompare.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btnGoToCompare.BackColor = Color.White;
            btnGoToCompare.Font = new Font("Segoe UI", 8.5F, FontStyle.Bold);
            btnGoToCompare.ForeColor = Color.FromArgb(0, 122, 204);
            btnGoToCompare.Location = new Point(669, 19);
            btnGoToCompare.Name = "btnGoToCompare";
            btnGoToCompare.Size = new Size(112, 30);
            btnGoToCompare.TabIndex = 9;
            btnGoToCompare.Text = "Сравнить →";
            btnGoToCompare.UseVisualStyleBackColor = false;
            btnGoToCompare.Visible = false;
            btnGoToCompare.Click += btnGoToCompare_Click;
            // 
            // grpCompareStep
            // 
            grpCompareStep.Controls.Add(lblCompareOriginal);
            grpCompareStep.Controls.Add(txtCompareOriginal);
            grpCompareStep.Controls.Add(btnSelectCompareOriginal);
            grpCompareStep.Controls.Add(lblCompareFilled);
            grpCompareStep.Controls.Add(txtCompareFilled);
            grpCompareStep.Controls.Add(btnSelectCompareFilled);
            grpCompareStep.Controls.Add(btnCompareTimesheets);
            grpCompareStep.Controls.Add(lblCompareResultCaption);
            grpCompareStep.Controls.Add(lblCompareResultPath);
            grpCompareStep.Controls.Add(pnlCompareHeader);
            grpCompareStep.Dock = DockStyle.Fill;
            grpCompareStep.Name = "grpCompareStep";
            grpCompareStep.Padding = new Padding(16, 0, 16, 0);
            grpCompareStep.Size = new Size(816, 292);
            grpCompareStep.TabIndex = 1;
            grpCompareStep.TabStop = false;
            grpCompareStep.Text = "";
            // 
            // pnlCompareHeader
            // 
            pnlCompareHeader.BackColor = Color.FromArgb(0, 153, 219);
            pnlCompareHeader.Controls.Add(lblCompareStepNumber);
            pnlCompareHeader.Controls.Add(lblCompareStepTitle);
            pnlCompareHeader.Controls.Add(btnBackToFill);
            pnlCompareHeader.Dock = DockStyle.Top;
            pnlCompareHeader.Location = new Point(16, 16);
            pnlCompareHeader.Name = "pnlCompareHeader";
            pnlCompareHeader.Size = new Size(784, 68);
            pnlCompareHeader.TabIndex = 0;
            // 
            // lblCompareStepNumber
            // 
            lblCompareStepNumber.AutoSize = true;
            lblCompareStepNumber.ForeColor = Color.White;
            lblCompareStepNumber.Location = new Point(16, 9);
            lblCompareStepNumber.Name = "lblCompareStepNumber";
            lblCompareStepNumber.Size = new Size(73, 15);
            lblCompareStepNumber.TabIndex = 0;
            lblCompareStepNumber.Text = "Шаг 2 из 2";
            // 
            // lblCompareStepTitle
            // 
            lblCompareStepTitle.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            lblCompareStepTitle.AutoEllipsis = true;
            lblCompareStepTitle.Font = new Font("Segoe UI", 17F, FontStyle.Bold);
            lblCompareStepTitle.ForeColor = Color.White;
            lblCompareStepTitle.Location = new Point(14, 25);
            lblCompareStepTitle.Name = "lblCompareStepTitle";
            lblCompareStepTitle.Size = new Size(600, 34);
            lblCompareStepTitle.TabIndex = 1;
            lblCompareStepTitle.Text = "Сравнение табелей";
            // 
            // lblCompareOriginal
            // 
            lblCompareOriginal.AutoSize = true;
            lblCompareOriginal.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            lblCompareOriginal.Location = new Point(20, 111);
            lblCompareOriginal.Name = "lblCompareOriginal";
            lblCompareOriginal.Size = new Size(207, 15);
            lblCompareOriginal.TabIndex = 0;
            lblCompareOriginal.Text = "1. Исходный табель подрядчика";
            // 
            // txtCompareOriginal
            // 
            txtCompareOriginal.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            txtCompareOriginal.Location = new Point(20, 132);
            txtCompareOriginal.Name = "txtCompareOriginal";
            txtCompareOriginal.PlaceholderText = "Выберите исходный табель подрядчика";
            txtCompareOriginal.ReadOnly = true;
            txtCompareOriginal.Size = new Size(642, 23);
            txtCompareOriginal.TabIndex = 1;
            // 
            // btnSelectCompareOriginal
            // 
            btnSelectCompareOriginal.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btnSelectCompareOriginal.Location = new Point(674, 131);
            btnSelectCompareOriginal.Name = "btnSelectCompareOriginal";
            btnSelectCompareOriginal.Size = new Size(112, 25);
            btnSelectCompareOriginal.TabIndex = 2;
            btnSelectCompareOriginal.Text = "Выбрать...";
            btnSelectCompareOriginal.UseVisualStyleBackColor = true;
            btnSelectCompareOriginal.Click += btnSelectCompareOriginal_Click;
            // 
            // lblCompareFilled
            // 
            lblCompareFilled.AutoSize = true;
            lblCompareFilled.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            lblCompareFilled.Location = new Point(20, 171);
            lblCompareFilled.Name = "lblCompareFilled";
            lblCompareFilled.Size = new Size(150, 15);
            lblCompareFilled.TabIndex = 3;
            lblCompareFilled.Text = "2. Заполненный табель";
            // 
            // txtCompareFilled
            // 
            txtCompareFilled.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            txtCompareFilled.Location = new Point(20, 192);
            txtCompareFilled.Name = "txtCompareFilled";
            txtCompareFilled.PlaceholderText = "Выберите заполненный табель";
            txtCompareFilled.ReadOnly = true;
            txtCompareFilled.Size = new Size(642, 23);
            txtCompareFilled.TabIndex = 4;
            // 
            // btnSelectCompareFilled
            // 
            btnSelectCompareFilled.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btnSelectCompareFilled.Location = new Point(674, 191);
            btnSelectCompareFilled.Name = "btnSelectCompareFilled";
            btnSelectCompareFilled.Size = new Size(112, 25);
            btnSelectCompareFilled.TabIndex = 5;
            btnSelectCompareFilled.Text = "Выбрать...";
            btnSelectCompareFilled.UseVisualStyleBackColor = true;
            btnSelectCompareFilled.Click += btnSelectCompareFilled_Click;
            // 
            // btnCompareTimesheets
            // 
            btnCompareTimesheets.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btnCompareTimesheets.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            btnCompareTimesheets.Location = new Point(656, 240);
            btnCompareTimesheets.Name = "btnCompareTimesheets";
            btnCompareTimesheets.Size = new Size(130, 30);
            btnCompareTimesheets.TabIndex = 6;
            btnCompareTimesheets.Text = "Сравнить";
            btnCompareTimesheets.UseVisualStyleBackColor = true;
            btnCompareTimesheets.Click += btnCompareTimesheets_Click;
            // 
            // lblCompareResultCaption
            // 
            lblCompareResultCaption.AutoSize = true;
            lblCompareResultCaption.Location = new Point(20, 240);
            lblCompareResultCaption.Name = "lblCompareResultCaption";
            lblCompareResultCaption.Size = new Size(159, 15);
            lblCompareResultCaption.TabIndex = 7;
            lblCompareResultCaption.Text = "Результат будет сохранён:";
            // 
            // lblCompareResultPath
            // 
            lblCompareResultPath.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            lblCompareResultPath.AutoEllipsis = true;
            lblCompareResultPath.ForeColor = SystemColors.GrayText;
            lblCompareResultPath.Location = new Point(20, 264);
            lblCompareResultPath.Name = "lblCompareResultPath";
            lblCompareResultPath.Size = new Size(620, 23);
            lblCompareResultPath.TabIndex = 8;
            // 
            // btnBackToFill
            // 
            btnBackToFill.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btnBackToFill.BackColor = Color.White;
            btnBackToFill.Font = new Font("Segoe UI", 8.5F, FontStyle.Bold);
            btnBackToFill.ForeColor = Color.FromArgb(0, 122, 204);
            btnBackToFill.Location = new Point(653, 19);
            btnBackToFill.Name = "btnBackToFill";
            btnBackToFill.Size = new Size(128, 30);
            btnBackToFill.TabIndex = 9;
            btnBackToFill.Text = "← Заполнение";
            btnBackToFill.UseVisualStyleBackColor = false;
            btnBackToFill.Click += btnBackToFill_Click;
            // 
            // grpProgress
            // 
            grpProgress.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            grpProgress.Controls.Add(progressMain);
            grpProgress.Controls.Add(lblProgressStatus);
            grpProgress.Controls.Add(txtLog);
            grpProgress.Location = new Point(12, 298);
            grpProgress.Name = "grpProgress";
            grpProgress.Padding = new Padding(12);
            grpProgress.Size = new Size(816, 328);
            grpProgress.TabIndex = 1;
            grpProgress.TabStop = false;
            grpProgress.Text = "Ход выполнения";
            // 
            // progressMain
            // 
            progressMain.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            progressMain.Location = new Point(15, 25);
            progressMain.Name = "progressMain";
            progressMain.Size = new Size(786, 22);
            progressMain.TabIndex = 0;
            // 
            // lblProgressStatus
            // 
            lblProgressStatus.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            lblProgressStatus.AutoEllipsis = true;
            lblProgressStatus.ForeColor = SystemColors.GrayText;
            lblProgressStatus.Location = new Point(15, 54);
            lblProgressStatus.Name = "lblProgressStatus";
            lblProgressStatus.Size = new Size(786, 20);
            lblProgressStatus.TabIndex = 1;
            lblProgressStatus.Text = "Выберите файлы и запустите нужную операцию.";
            // 
            // txtLog
            // 
            txtLog.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            txtLog.BackColor = SystemColors.Window;
            txtLog.Location = new Point(15, 77);
            txtLog.Multiline = true;
            txtLog.Name = "txtLog";
            txtLog.ReadOnly = true;
            txtLog.ScrollBars = ScrollBars.Vertical;
            txtLog.Size = new Size(786, 236);
            txtLog.TabIndex = 2;
            // 
            // FrmMain
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(840, 638);
            Controls.Add(grpProgress);
            Controls.Add(pnlWizard);
            MinimumSize = new Size(720, 580);
            Name = "FrmMain";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "Табель СКУД";
            pnlWizard.ResumeLayout(false);
            grpFillStep.ResumeLayout(false);
            grpFillStep.PerformLayout();
            pnlFillHeader.ResumeLayout(false);
            pnlFillHeader.PerformLayout();
            grpCompareStep.ResumeLayout(false);
            grpCompareStep.PerformLayout();
            pnlCompareHeader.ResumeLayout(false);
            pnlCompareHeader.PerformLayout();
            grpProgress.ResumeLayout(false);
            grpProgress.PerformLayout();
            ResumeLayout(false);
        }

        #endregion

        private Panel pnlWizard;
        private GroupBox grpFillStep;
        private GroupBox grpCompareStep;
        private Panel pnlFillHeader;
        private Label lblFillStepNumber;
        private Label lblFillStepTitle;
        private Panel pnlCompareHeader;
        private Label lblCompareStepNumber;
        private Label lblCompareStepTitle;
        private Label lblFillTimesheet;
        private TextBox txtFillTimesheet;
        private Button btnSelectFillTimesheet;
        private Label lblSkud;
        private TextBox txtSkud;
        private Button btnSelectSkud;
        private Button btnProcessTimesheet;
        private Label lblFillResultCaption;
        private Label lblFillResultPath;
        private Button btnGoToCompare;
        private Label lblCompareOriginal;
        private TextBox txtCompareOriginal;
        private Button btnSelectCompareOriginal;
        private Label lblCompareFilled;
        private TextBox txtCompareFilled;
        private Button btnSelectCompareFilled;
        private Button btnCompareTimesheets;
        private Label lblCompareResultCaption;
        private Label lblCompareResultPath;
        private Button btnBackToFill;
        private GroupBox grpProgress;
        private ProgressBar progressMain;
        private Label lblProgressStatus;
        private TextBox txtLog;
    }
}
