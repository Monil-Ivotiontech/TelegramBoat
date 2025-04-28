namespace Json_Creator_Tool
{
    partial class Form1
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
            cmbType = new ComboBox();
            label1 = new Label();
            label2 = new Label();
            dt = new DateTimePicker();
            label3 = new Label();
            txtFrom = new NumericUpDown();
            txtTo = new NumericUpDown();
            label4 = new Label();
            txtAdd = new NumericUpDown();
            label5 = new Label();
            btnCancel = new Button();
            btnCreate = new Button();
            btnClose = new Button();
            ((System.ComponentModel.ISupportInitialize)txtFrom).BeginInit();
            ((System.ComponentModel.ISupportInitialize)txtTo).BeginInit();
            ((System.ComponentModel.ISupportInitialize)txtAdd).BeginInit();
            SuspendLayout();
            // 
            // cmbType
            // 
            cmbType.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbType.FormattingEnabled = true;
            cmbType.Items.AddRange(new object[] { "SENSEX", "NIFTY" });
            cmbType.Location = new Point(63, 42);
            cmbType.Name = "cmbType";
            cmbType.Size = new Size(151, 28);
            cmbType.TabIndex = 1;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(14, 46);
            label1.Name = "label1";
            label1.Size = new Size(43, 20);
            label1.TabIndex = 1;
            label1.Text = "Type:";
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(11, 79);
            label2.Name = "label2";
            label2.Size = new Size(44, 20);
            label2.TabIndex = 2;
            label2.Text = "Date:";
            // 
            // dt
            // 
            dt.Format = DateTimePickerFormat.Short;
            dt.Location = new Point(62, 76);
            dt.Name = "dt";
            dt.Size = new Size(151, 27);
            dt.TabIndex = 2;
            dt.Value = new DateTime(2025, 4, 21, 0, 0, 0, 0);
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new Point(10, 112);
            label3.Name = "label3";
            label3.Size = new Size(46, 20);
            label3.TabIndex = 3;
            label3.Text = "From:";
            // 
            // txtFrom
            // 
            txtFrom.Location = new Point(63, 109);
            txtFrom.Maximum = new decimal(new int[] { 1000000, 0, 0, 0 });
            txtFrom.Name = "txtFrom";
            txtFrom.Size = new Size(151, 27);
            txtFrom.TabIndex = 4;
            // 
            // txtTo
            // 
            txtTo.Location = new Point(62, 142);
            txtTo.Maximum = new decimal(new int[] { 1000000, 0, 0, 0 });
            txtTo.Name = "txtTo";
            txtTo.Size = new Size(151, 27);
            txtTo.TabIndex = 6;
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Location = new Point(28, 145);
            label4.Name = "label4";
            label4.Size = new Size(28, 20);
            label4.TabIndex = 5;
            label4.Text = "To:";
            // 
            // txtAdd
            // 
            txtAdd.Location = new Point(62, 175);
            txtAdd.Maximum = new decimal(new int[] { 1000, 0, 0, 0 });
            txtAdd.Name = "txtAdd";
            txtAdd.Size = new Size(151, 27);
            txtAdd.TabIndex = 8;
            // 
            // label5
            // 
            label5.AutoSize = true;
            label5.Location = new Point(16, 178);
            label5.Name = "label5";
            label5.Size = new Size(40, 20);
            label5.TabIndex = 7;
            label5.Text = "Add:";
            // 
            // btnCancel
            // 
            btnCancel.BackColor = Color.Coral;
            btnCancel.ForeColor = SystemColors.ActiveCaptionText;
            btnCancel.Location = new Point(16, 222);
            btnCancel.Name = "btnCancel";
            btnCancel.Size = new Size(94, 35);
            btnCancel.TabIndex = 9;
            btnCancel.Text = "Clear";
            btnCancel.UseVisualStyleBackColor = false;
            btnCancel.Click += btnCancel_Click;
            // 
            // btnCreate
            // 
            btnCreate.BackColor = Color.IndianRed;
            btnCreate.Location = new Point(119, 222);
            btnCreate.Name = "btnCreate";
            btnCreate.Size = new Size(94, 35);
            btnCreate.TabIndex = 10;
            btnCreate.Text = "Create";
            btnCreate.UseVisualStyleBackColor = false;
            btnCreate.Click += btnCreate_Click;
            // 
            // btnClose
            // 
            btnClose.BackColor = Color.IndianRed;
            btnClose.Location = new Point(181, 3);
            btnClose.Name = "btnClose";
            btnClose.Size = new Size(33, 32);
            btnClose.TabIndex = 11;
            btnClose.Text = "X";
            btnClose.UseVisualStyleBackColor = false;
            btnClose.Click += btnClose_Click;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = SystemColors.ActiveCaption;
            ClientSize = new Size(231, 272);
            Controls.Add(btnClose);
            Controls.Add(btnCreate);
            Controls.Add(btnCancel);
            Controls.Add(txtAdd);
            Controls.Add(label5);
            Controls.Add(txtTo);
            Controls.Add(label4);
            Controls.Add(txtFrom);
            Controls.Add(label3);
            Controls.Add(dt);
            Controls.Add(label2);
            Controls.Add(label1);
            Controls.Add(cmbType);
            FormBorderStyle = FormBorderStyle.None;
            Name = "Form1";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "JSON";
            Load += Form1_Load;
            ((System.ComponentModel.ISupportInitialize)txtFrom).EndInit();
            ((System.ComponentModel.ISupportInitialize)txtTo).EndInit();
            ((System.ComponentModel.ISupportInitialize)txtAdd).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private ComboBox cmbType;
        private Label label1;
        private Label label2;
        private DateTimePicker dt;
        private Label label3;
        private NumericUpDown txtFrom;
        private NumericUpDown txtTo;
        private Label label4;
        private NumericUpDown txtAdd;
        private Label label5;
        private Button btnCancel;
        private Button btnCreate;
        private Button btnClose;
    }
}
