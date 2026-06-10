namespace Tabel
{
    partial class FrmMain
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            btnDoIt = new Button();
            txtTabel = new TextBox();
            btnSelectTabel = new Button();
            txtLog = new TextBox();
            btnCompare = new Button();
            btnSelectSkud = new Button();
            txtSkud = new TextBox();
            SuspendLayout();
            // 
            // btnDoIt
            // 
            btnDoIt.Location = new Point(515, 17);
            btnDoIt.Name = "btnDoIt";
            btnDoIt.Size = new Size(101, 44);
            btnDoIt.TabIndex = 0;
            btnDoIt.Text = "Обработать";
            btnDoIt.UseVisualStyleBackColor = true;
            btnDoIt.Click += btnDoIt_Click;
            // 
            // txtTabel
            // 
            txtTabel.Location = new Point(13, 14);
            txtTabel.Name = "txtTabel";
            txtTabel.ReadOnly = true;
            txtTabel.Size = new Size(420, 23);
            txtTabel.TabIndex = 1;
            txtTabel.Text = "Выберите табель подрядчика";
            // 
            // btnSelectTabel
            // 
            btnSelectTabel.Location = new Point(444, 14);
            btnSelectTabel.Name = "btnSelectTabel";
            btnSelectTabel.Size = new Size(41, 23);
            btnSelectTabel.TabIndex = 2;
            btnSelectTabel.Text = "...";
            btnSelectTabel.UseVisualStyleBackColor = true;
            btnSelectTabel.Click += btnSelectTabel_Click;
            // 
            // txtLog
            // 
            txtLog.Location = new Point(12, 81);
            txtLog.Multiline = true;
            txtLog.Name = "txtLog";
            txtLog.Size = new Size(604, 357);
            txtLog.TabIndex = 3;
            // 
            // btnCompare
            // 
            btnCompare.Location = new Point(529, 71);
            btnCompare.Name = "btnCompare";
            btnCompare.Size = new Size(79, 23);
            btnCompare.TabIndex = 4;
            btnCompare.Text = "Сравнить";
            btnCompare.UseVisualStyleBackColor = true;
            btnCompare.Click += btnCompare_Click;
            // 
            // btnSelectSkud
            // 
            btnSelectSkud.Location = new Point(444, 43);
            btnSelectSkud.Name = "btnSelectSkud";
            btnSelectSkud.Size = new Size(41, 23);
            btnSelectSkud.TabIndex = 5;
            btnSelectSkud.Text = "...";
            btnSelectSkud.UseVisualStyleBackColor = true;
            btnSelectSkud.Click += btnSelectSkud_Click;
            // 
            // txtSkud
            // 
            txtSkud.Location = new Point(13, 43);
            txtSkud.Name = "txtSkud";
            txtSkud.ReadOnly = true;
            txtSkud.Size = new Size(420, 23);
            txtSkud.TabIndex = 1;
            txtSkud.Text = "Выберите выгрузку СКУД";
            // 
            // FrmMain
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(627, 450);
            Controls.Add(btnSelectSkud);
            Controls.Add(btnCompare);
            Controls.Add(txtLog);
            Controls.Add(btnSelectTabel);
            Controls.Add(txtSkud);
            Controls.Add(txtTabel);
            Controls.Add(btnDoIt);
            FormBorderStyle = FormBorderStyle.FixedToolWindow;
            Name = "FrmMain";
            Text = "Табель СКУД версия 1006";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button btnDoIt;
        private TextBox txtTabel;
        private Button btnSelectTabel;
        private TextBox txtLog;
        private Button btnCheck;
        private Button btnCompare;
        private Button btnSelectSkud;
        private TextBox txtSkud;
    }
}
