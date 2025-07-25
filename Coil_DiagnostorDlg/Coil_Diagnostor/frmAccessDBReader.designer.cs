namespace Coil_Diagnostor
{
    partial class frmAccessDBReader
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
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
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmAccessDBReader));
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            this.panel1 = new System.Windows.Forms.Panel();
            this.chkReferenceValue = new System.Windows.Forms.CheckBox();
            this.label8 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.rboDRPI = new System.Windows.Forms.RadioButton();
            this.rboRCS = new System.Windows.Forms.RadioButton();
            this.teAccessDBFileName = new System.Windows.Forms.TextBox();
            this.btnClose = new System.Windows.Forms.Button();
            this.btnSave = new System.Windows.Forms.Button();
            this.nmHogi = new System.Windows.Forms.NumericUpDown();
            this.nmOh_Degree = new System.Windows.Forms.NumericUpDown();
            this.label7 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.btnAccessDBFileNameOpen = new System.Windows.Forms.Button();
            this.btnSearch = new System.Windows.Forms.Button();
            this.panel2 = new System.Windows.Forms.Panel();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.dgvAccessData = new System.Windows.Forms.DataGridView();
            this.dgvReferenceValue = new System.Windows.Forms.DataGridView();
            this.label9 = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nmHogi)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nmOh_Degree)).BeginInit();
            this.panel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvAccessData)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvReferenceValue)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(215)))), ((int)(((byte)(231)))), ((int)(((byte)(252)))));
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.chkReferenceValue);
            this.panel1.Controls.Add(this.label8);
            this.panel1.Controls.Add(this.label6);
            this.panel1.Controls.Add(this.label5);
            this.panel1.Controls.Add(this.rboDRPI);
            this.panel1.Controls.Add(this.rboRCS);
            this.panel1.Controls.Add(this.teAccessDBFileName);
            this.panel1.Controls.Add(this.btnClose);
            this.panel1.Controls.Add(this.btnSave);
            this.panel1.Controls.Add(this.nmHogi);
            this.panel1.Controls.Add(this.nmOh_Degree);
            this.panel1.Controls.Add(this.label7);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.btnAccessDBFileNameOpen);
            this.panel1.Controls.Add(this.btnSearch);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Margin = new System.Windows.Forms.Padding(5);
            this.panel1.Name = "panel1";
            this.panel1.Padding = new System.Windows.Forms.Padding(5);
            this.panel1.Size = new System.Drawing.Size(1284, 70);
            this.panel1.TabIndex = 0;
            // 
            // chkReferenceValue
            // 
            this.chkReferenceValue.AutoSize = true;
            this.chkReferenceValue.Location = new System.Drawing.Point(715, 40);
            this.chkReferenceValue.Name = "chkReferenceValue";
            this.chkReferenceValue.Size = new System.Drawing.Size(72, 16);
            this.chkReferenceValue.TabIndex = 43;
            this.chkReferenceValue.Tag = "";
            this.chkReferenceValue.Text = "가져오기";
            this.chkReferenceValue.UseVisualStyleBackColor = true;
            // 
            // label8
            // 
            this.label8.BackColor = System.Drawing.Color.Transparent;
            this.label8.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label8.Location = new System.Drawing.Point(589, 37);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(120, 23);
            this.label8.TabIndex = 42;
            this.label8.Text = "기준 값";
            this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label6
            // 
            this.label6.BackColor = System.Drawing.Color.Transparent;
            this.label6.Location = new System.Drawing.Point(509, 7);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(42, 23);
            this.label6.TabIndex = 41;
            this.label6.Text = "차수";
            this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label5
            // 
            this.label5.BackColor = System.Drawing.Color.Transparent;
            this.label5.Location = new System.Drawing.Point(227, 7);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(42, 23);
            this.label5.TabIndex = 40;
            this.label5.Text = "호기";
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // rboDRPI
            // 
            this.rboDRPI.AutoSize = true;
            this.rboDRPI.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.rboDRPI.Location = new System.Drawing.Point(775, 10);
            this.rboDRPI.Name = "rboDRPI";
            this.rboDRPI.Size = new System.Drawing.Size(56, 17);
            this.rboDRPI.TabIndex = 3;
            this.rboDRPI.Text = "DRPI";
            this.rboDRPI.UseVisualStyleBackColor = true;
            this.rboDRPI.CheckedChanged += new System.EventHandler(this.rboDRPI_CheckedChanged);
            // 
            // rboRCS
            // 
            this.rboRCS.AutoSize = true;
            this.rboRCS.Checked = true;
            this.rboRCS.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.rboRCS.Location = new System.Drawing.Point(715, 10);
            this.rboRCS.Name = "rboRCS";
            this.rboRCS.Size = new System.Drawing.Size(54, 17);
            this.rboRCS.TabIndex = 2;
            this.rboRCS.TabStop = true;
            this.rboRCS.Text = "RCS";
            this.rboRCS.UseVisualStyleBackColor = true;
            this.rboRCS.CheckedChanged += new System.EventHandler(this.rboRCS_CheckedChanged);
            // 
            // teAccessDBFileName
            // 
            this.teAccessDBFileName.Location = new System.Drawing.Point(136, 38);
            this.teAccessDBFileName.Name = "teAccessDBFileName";
            this.teAccessDBFileName.ReadOnly = true;
            this.teAccessDBFileName.Size = new System.Drawing.Size(397, 21);
            this.teAccessDBFileName.TabIndex = 4;
            // 
            // btnClose
            // 
            this.btnClose.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnClose.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnClose.BackgroundImage")));
            this.btnClose.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnClose.Font = new System.Drawing.Font("굴림", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.btnClose.ForeColor = System.Drawing.Color.White;
            this.btnClose.Location = new System.Drawing.Point(1146, 7);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(130, 54);
            this.btnClose.TabIndex = 12;
            this.btnClose.Text = "닫 기 (F12)";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnSave
            // 
            this.btnSave.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSave.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnSave.BackgroundImage")));
            this.btnSave.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnSave.Font = new System.Drawing.Font("굴림", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.btnSave.ForeColor = System.Drawing.Color.White;
            this.btnSave.Location = new System.Drawing.Point(1016, 7);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(130, 54);
            this.btnSave.TabIndex = 11;
            this.btnSave.Text = "저 장 (F3)";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // nmHogi
            // 
            this.nmHogi.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.nmHogi.Location = new System.Drawing.Point(137, 8);
            this.nmHogi.Maximum = new decimal(new int[] {
            100000000,
            0,
            0,
            0});
            this.nmHogi.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.nmHogi.Name = "nmHogi";
            this.nmHogi.Size = new System.Drawing.Size(88, 21);
            this.nmHogi.TabIndex = 1;
            this.nmHogi.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.nmHogi.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.nmHogi.ValueChanged += new System.EventHandler(this.nmOh_Degree_ValueChanged);
            this.nmHogi.Leave += new System.EventHandler(this.NumericUpDown_Leave);
            // 
            // nmOh_Degree
            // 
            this.nmOh_Degree.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.nmOh_Degree.Location = new System.Drawing.Point(427, 8);
            this.nmOh_Degree.Maximum = new decimal(new int[] {
            100000000,
            0,
            0,
            0});
            this.nmOh_Degree.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.nmOh_Degree.Name = "nmOh_Degree";
            this.nmOh_Degree.Size = new System.Drawing.Size(80, 21);
            this.nmOh_Degree.TabIndex = 1;
            this.nmOh_Degree.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.nmOh_Degree.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.nmOh_Degree.ValueChanged += new System.EventHandler(this.nmOh_Degree_ValueChanged);
            // 
            // label7
            // 
            this.label7.BackColor = System.Drawing.Color.Transparent;
            this.label7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label7.Location = new System.Drawing.Point(301, 7);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(120, 23);
            this.label7.TabIndex = 37;
            this.label7.Text = "OH 차수";
            this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.Color.Transparent;
            this.label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label1.Location = new System.Drawing.Point(11, 37);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(120, 23);
            this.label1.TabIndex = 38;
            this.label1.Text = "대상 DB 파일";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label2
            // 
            this.label2.BackColor = System.Drawing.Color.Transparent;
            this.label2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label2.Location = new System.Drawing.Point(589, 7);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(120, 23);
            this.label2.TabIndex = 38;
            this.label2.Text = "측정유형";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label4
            // 
            this.label4.BackColor = System.Drawing.Color.Transparent;
            this.label4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label4.Location = new System.Drawing.Point(11, 7);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(120, 23);
            this.label4.TabIndex = 38;
            this.label4.Text = "측정대상";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // btnAccessDBFileNameOpen
            // 
            this.btnAccessDBFileNameOpen.BackgroundImage = global::Coil_Diagnostor.Properties.Resources.Folder_Add_icon;
            this.btnAccessDBFileNameOpen.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnAccessDBFileNameOpen.Font = new System.Drawing.Font("굴림", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.btnAccessDBFileNameOpen.ForeColor = System.Drawing.Color.White;
            this.btnAccessDBFileNameOpen.Location = new System.Drawing.Point(535, 33);
            this.btnAccessDBFileNameOpen.Name = "btnAccessDBFileNameOpen";
            this.btnAccessDBFileNameOpen.Size = new System.Drawing.Size(30, 30);
            this.btnAccessDBFileNameOpen.TabIndex = 5;
            this.btnAccessDBFileNameOpen.Text = "가져오기 (F2)";
            this.btnAccessDBFileNameOpen.UseVisualStyleBackColor = true;
            this.btnAccessDBFileNameOpen.Click += new System.EventHandler(this.btnAccessDBFileNameOpen_Click);
            // 
            // btnSearch
            // 
            this.btnSearch.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSearch.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnSearch.BackgroundImage")));
            this.btnSearch.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnSearch.Font = new System.Drawing.Font("굴림", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.btnSearch.ForeColor = System.Drawing.Color.White;
            this.btnSearch.Location = new System.Drawing.Point(886, 7);
            this.btnSearch.Name = "btnSearch";
            this.btnSearch.Size = new System.Drawing.Size(130, 54);
            this.btnSearch.TabIndex = 10;
            this.btnSearch.Text = "가져오기 (F2)";
            this.btnSearch.UseVisualStyleBackColor = true;
            this.btnSearch.Click += new System.EventHandler(this.btnSearch_Click);
            // 
            // panel2
            // 
            this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel2.Controls.Add(this.progressBar1);
            this.panel2.Controls.Add(this.splitContainer1);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(0, 70);
            this.panel2.Name = "panel2";
            this.panel2.Padding = new System.Windows.Forms.Padding(3);
            this.panel2.Size = new System.Drawing.Size(1284, 731);
            this.panel2.TabIndex = 1;
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(251, 310);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(796, 40);
            this.progressBar1.TabIndex = 2;
            this.progressBar1.Visible = false;
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.Location = new System.Drawing.Point(3, 3);
            this.splitContainer1.Name = "splitContainer1";
            this.splitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.dgvAccessData);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.dgvReferenceValue);
            this.splitContainer1.Panel2.Controls.Add(this.label9);
            this.splitContainer1.Size = new System.Drawing.Size(1276, 723);
            this.splitContainer1.SplitterDistance = 450;
            this.splitContainer1.TabIndex = 4;
            // 
            // dgvAccessData
            // 
            this.dgvAccessData.AllowUserToAddRows = false;
            this.dgvAccessData.AllowUserToDeleteRows = false;
            this.dgvAccessData.AllowUserToResizeRows = false;
            this.dgvAccessData.BackgroundColor = System.Drawing.Color.White;
            this.dgvAccessData.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            this.dgvAccessData.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("굴림", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgvAccessData.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dgvAccessData.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvAccessData.Location = new System.Drawing.Point(0, 0);
            this.dgvAccessData.Name = "dgvAccessData";
            this.dgvAccessData.RowHeadersVisible = false;
            this.dgvAccessData.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;
            this.dgvAccessData.RowTemplate.Height = 23;
            this.dgvAccessData.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvAccessData.Size = new System.Drawing.Size(1276, 450);
            this.dgvAccessData.TabIndex = 1;
            this.dgvAccessData.RowHeadersWidthChanged += new System.EventHandler(this.dgvAccessData_RowHeadersWidthChanged);
            this.dgvAccessData.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvAccessData_CellClick);
            this.dgvAccessData.ColumnDividerWidthChanged += new System.Windows.Forms.DataGridViewColumnEventHandler(this.dgvAccessData_ColumnDividerWidthChanged);
            this.dgvAccessData.Scroll += new System.Windows.Forms.ScrollEventHandler(this.dgvAccessData_Scroll);
            // 
            // dgvReferenceValue
            // 
            this.dgvReferenceValue.AllowUserToAddRows = false;
            this.dgvReferenceValue.AllowUserToDeleteRows = false;
            this.dgvReferenceValue.AllowUserToResizeRows = false;
            this.dgvReferenceValue.BackgroundColor = System.Drawing.Color.White;
            this.dgvReferenceValue.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            this.dgvReferenceValue.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("굴림", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgvReferenceValue.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle2;
            this.dgvReferenceValue.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvReferenceValue.Location = new System.Drawing.Point(0, 27);
            this.dgvReferenceValue.Name = "dgvReferenceValue";
            this.dgvReferenceValue.RowHeadersVisible = false;
            this.dgvReferenceValue.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;
            this.dgvReferenceValue.RowTemplate.Height = 23;
            this.dgvReferenceValue.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvReferenceValue.Size = new System.Drawing.Size(1276, 242);
            this.dgvReferenceValue.TabIndex = 2;
            // 
            // label9
            // 
            this.label9.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.label9.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label9.Dock = System.Windows.Forms.DockStyle.Top;
            this.label9.Font = new System.Drawing.Font("굴림", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.label9.Location = new System.Drawing.Point(0, 0);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(1276, 27);
            this.label9.TabIndex = 3;
            this.label9.Text = "▣ 기준 값 정보";
            this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // frmAccessDBReader
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1284, 801);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "frmAccessDBReader";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "기존 코일 성능시험기 데이터 가져오기";
            this.Load += new System.EventHandler(this.frmAccessDBReader_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nmHogi)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nmOh_Degree)).EndInit();
            this.panel2.ResumeLayout(false);
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvAccessData)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvReferenceValue)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btnSearch;
        private System.Windows.Forms.NumericUpDown nmOh_Degree;
        private System.Windows.Forms.TextBox teAccessDBFileName;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnAccessDBFileNameOpen;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.RadioButton rboDRPI;
        private System.Windows.Forms.RadioButton rboRCS;
        private System.Windows.Forms.Panel panel2;
        public System.Windows.Forms.DataGridView dgvAccessData;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.NumericUpDown nmHogi;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.CheckBox chkReferenceValue;
        private System.Windows.Forms.Label label8;
        public System.Windows.Forms.DataGridView dgvReferenceValue;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.SplitContainer splitContainer1;
    }
}