namespace Coil_Diagnostor
{
    partial class frmTDR_RCSHistory
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmTDR_RCSHistory));
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            this.panel1 = new System.Windows.Forms.Panel();
            this.teOhDegreeT = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.teOhDegreeF = new System.Windows.Forms.TextBox();
            this.cboPowerCabinet = new System.Windows.Forms.ComboBox();
            this.label26 = new System.Windows.Forms.Label();
            this.cboControlRodName = new System.Windows.Forms.ComboBox();
            this.cboCoilName = new System.Windows.Forms.ComboBox();
            this.cboHogi = new System.Windows.Forms.ComboBox();
            this.label27 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.btnClose = new System.Windows.Forms.Button();
            this.btnFolderMove = new System.Windows.Forms.Button();
            this.btnReport = new System.Windows.Forms.Button();
            this.btnSearch = new System.Windows.Forms.Button();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.dgvMeasurement = new System.Windows.Forms.DataGridView();
            this.splitContainer2 = new System.Windows.Forms.SplitContainer();
            this.pbTDRImage = new System.Windows.Forms.PictureBox();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvMeasurement)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer2)).BeginInit();
            this.splitContainer2.Panel1.SuspendLayout();
            this.splitContainer2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pbTDRImage)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(215)))), ((int)(((byte)(231)))), ((int)(((byte)(252)))));
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.teOhDegreeT);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.teOhDegreeF);
            this.panel1.Controls.Add(this.cboPowerCabinet);
            this.panel1.Controls.Add(this.label26);
            this.panel1.Controls.Add(this.cboControlRodName);
            this.panel1.Controls.Add(this.cboCoilName);
            this.panel1.Controls.Add(this.cboHogi);
            this.panel1.Controls.Add(this.label27);
            this.panel1.Controls.Add(this.label6);
            this.panel1.Controls.Add(this.label7);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.btnClose);
            this.panel1.Controls.Add(this.btnFolderMove);
            this.panel1.Controls.Add(this.btnReport);
            this.panel1.Controls.Add(this.btnSearch);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Margin = new System.Windows.Forms.Padding(5);
            this.panel1.Name = "panel1";
            this.panel1.Padding = new System.Windows.Forms.Padding(5);
            this.panel1.Size = new System.Drawing.Size(1284, 71);
            this.panel1.TabIndex = 1;
            // 
            // teOhDegreeT
            // 
            this.teOhDegreeT.Location = new System.Drawing.Point(507, 8);
            this.teOhDegreeT.Name = "teOhDegreeT";
            this.teOhDegreeT.Size = new System.Drawing.Size(100, 21);
            this.teOhDegreeT.TabIndex = 41;
            this.teOhDegreeT.TextChanged += new System.EventHandler(this.teOhDegreeF_TextChanged);
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.Color.Transparent;
            this.label1.Location = new System.Drawing.Point(489, 8);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(16, 23);
            this.label1.TabIndex = 40;
            this.label1.Text = "~";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // teOhDegreeF
            // 
            this.teOhDegreeF.Location = new System.Drawing.Point(387, 9);
            this.teOhDegreeF.Name = "teOhDegreeF";
            this.teOhDegreeF.Size = new System.Drawing.Size(100, 21);
            this.teOhDegreeF.TabIndex = 39;
            this.teOhDegreeF.TextChanged += new System.EventHandler(this.teOhDegreeF_TextChanged);
            // 
            // cboPowerCabinet
            // 
            this.cboPowerCabinet.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboPowerCabinet.FormattingEnabled = true;
            this.cboPowerCabinet.Location = new System.Drawing.Point(137, 41);
            this.cboPowerCabinet.Name = "cboPowerCabinet";
            this.cboPowerCabinet.Size = new System.Drawing.Size(100, 20);
            this.cboPowerCabinet.TabIndex = 7;
            this.cboPowerCabinet.SelectedIndexChanged += new System.EventHandler(this.cboPowerCabinet_SelectedIndexChanged);
            // 
            // label26
            // 
            this.label26.BackColor = System.Drawing.Color.Transparent;
            this.label26.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label26.Location = new System.Drawing.Point(12, 40);
            this.label26.Name = "label26";
            this.label26.Size = new System.Drawing.Size(120, 23);
            this.label26.TabIndex = 38;
            this.label26.Text = "전력합";
            this.label26.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // cboControlRodName
            // 
            this.cboControlRodName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboControlRodName.FormattingEnabled = true;
            this.cboControlRodName.Location = new System.Drawing.Point(387, 41);
            this.cboControlRodName.Name = "cboControlRodName";
            this.cboControlRodName.Size = new System.Drawing.Size(100, 20);
            this.cboControlRodName.TabIndex = 8;
            this.cboControlRodName.SelectedIndexChanged += new System.EventHandler(this.cboControlRodName_SelectedIndexChanged);
            // 
            // cboCoilName
            // 
            this.cboCoilName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboCoilName.FormattingEnabled = true;
            this.cboCoilName.Location = new System.Drawing.Point(637, 41);
            this.cboCoilName.Name = "cboCoilName";
            this.cboCoilName.Size = new System.Drawing.Size(100, 20);
            this.cboCoilName.TabIndex = 9;
            this.cboCoilName.SelectedIndexChanged += new System.EventHandler(this.cboCoilName_SelectedIndexChanged);
            // 
            // cboHogi
            // 
            this.cboHogi.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboHogi.FormattingEnabled = true;
            this.cboHogi.Location = new System.Drawing.Point(137, 9);
            this.cboHogi.Name = "cboHogi";
            this.cboHogi.Size = new System.Drawing.Size(100, 20);
            this.cboHogi.TabIndex = 1;
            this.cboHogi.SelectedIndexChanged += new System.EventHandler(this.cboHogi_SelectedIndexChanged);
            // 
            // label27
            // 
            this.label27.BackColor = System.Drawing.Color.Transparent;
            this.label27.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label27.Location = new System.Drawing.Point(262, 40);
            this.label27.Name = "label27";
            this.label27.Size = new System.Drawing.Size(120, 23);
            this.label27.TabIndex = 37;
            this.label27.Text = "제어봉";
            this.label27.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label6
            // 
            this.label6.BackColor = System.Drawing.Color.Transparent;
            this.label6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label6.Location = new System.Drawing.Point(512, 40);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(120, 23);
            this.label6.TabIndex = 37;
            this.label6.Text = "코일명";
            this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label7
            // 
            this.label7.BackColor = System.Drawing.Color.Transparent;
            this.label7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label7.Location = new System.Drawing.Point(262, 8);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(120, 23);
            this.label7.TabIndex = 37;
            this.label7.Text = "OH 차수";
            this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label4
            // 
            this.label4.BackColor = System.Drawing.Color.Transparent;
            this.label4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label4.Location = new System.Drawing.Point(12, 8);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(120, 23);
            this.label4.TabIndex = 38;
            this.label4.Text = "측정대상";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // btnClose
            // 
            this.btnClose.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnClose.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnClose.BackgroundImage")));
            this.btnClose.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnClose.Font = new System.Drawing.Font("굴림", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.btnClose.ForeColor = System.Drawing.Color.White;
            this.btnClose.Location = new System.Drawing.Point(1165, 6);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(112, 57);
            this.btnClose.TabIndex = 11;
            this.btnClose.Text = "닫 기 (F12)";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnFolderMove
            // 
            this.btnFolderMove.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnFolderMove.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnFolderMove.BackgroundImage")));
            this.btnFolderMove.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnFolderMove.Font = new System.Drawing.Font("굴림", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.btnFolderMove.ForeColor = System.Drawing.Color.White;
            this.btnFolderMove.Location = new System.Drawing.Point(1023, 6);
            this.btnFolderMove.Name = "btnFolderMove";
            this.btnFolderMove.Size = new System.Drawing.Size(139, 57);
            this.btnFolderMove.TabIndex = 11;
            this.btnFolderMove.Text = "폴더바로가기 (F4)";
            this.btnFolderMove.UseVisualStyleBackColor = true;
            this.btnFolderMove.Click += new System.EventHandler(this.btnFolderMove_Click);
            // 
            // btnReport
            // 
            this.btnReport.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnReport.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnReport.BackgroundImage")));
            this.btnReport.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnReport.Font = new System.Drawing.Font("굴림", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.btnReport.ForeColor = System.Drawing.Color.White;
            this.btnReport.Location = new System.Drawing.Point(881, 6);
            this.btnReport.Name = "btnReport";
            this.btnReport.Size = new System.Drawing.Size(139, 57);
            this.btnReport.TabIndex = 11;
            this.btnReport.Text = "보고서 생성 (F3)";
            this.btnReport.UseVisualStyleBackColor = true;
            this.btnReport.Click += new System.EventHandler(this.btnReport_Click);
            // 
            // btnSearch
            // 
            this.btnSearch.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSearch.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnSearch.BackgroundImage")));
            this.btnSearch.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnSearch.Font = new System.Drawing.Font("굴림", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.btnSearch.ForeColor = System.Drawing.Color.White;
            this.btnSearch.Location = new System.Drawing.Point(766, 6);
            this.btnSearch.Name = "btnSearch";
            this.btnSearch.Size = new System.Drawing.Size(112, 57);
            this.btnSearch.TabIndex = 11;
            this.btnSearch.Text = "조 회 (F2)";
            this.btnSearch.UseVisualStyleBackColor = true;
            this.btnSearch.Click += new System.EventHandler(this.btnSearch_Click);
            // 
            // splitContainer1
            // 
            this.splitContainer1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.FixedPanel = System.Windows.Forms.FixedPanel.Panel1;
            this.splitContainer1.Location = new System.Drawing.Point(0, 71);
            this.splitContainer1.Name = "splitContainer1";
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.dgvMeasurement);
            this.splitContainer1.Panel1.Padding = new System.Windows.Forms.Padding(3);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.splitContainer2);
            this.splitContainer1.Size = new System.Drawing.Size(1284, 730);
            this.splitContainer1.SplitterDistance = 432;
            this.splitContainer1.TabIndex = 2;
            // 
            // dgvMeasurement
            // 
            this.dgvMeasurement.AllowUserToAddRows = false;
            this.dgvMeasurement.AllowUserToDeleteRows = false;
            this.dgvMeasurement.AllowUserToResizeRows = false;
            this.dgvMeasurement.BackgroundColor = System.Drawing.Color.White;
            this.dgvMeasurement.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            this.dgvMeasurement.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("굴림", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgvMeasurement.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle2;
            this.dgvMeasurement.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvMeasurement.Location = new System.Drawing.Point(3, 3);
            this.dgvMeasurement.Name = "dgvMeasurement";
            this.dgvMeasurement.RowHeadersVisible = false;
            this.dgvMeasurement.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;
            this.dgvMeasurement.RowTemplate.Height = 23;
            this.dgvMeasurement.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvMeasurement.Size = new System.Drawing.Size(424, 722);
            this.dgvMeasurement.TabIndex = 3;
            this.dgvMeasurement.RowHeadersWidthChanged += new System.EventHandler(this.dgvMeasurement_RowHeadersWidthChanged);
            this.dgvMeasurement.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvMeasurement_CellClick);
            this.dgvMeasurement.ColumnDividerWidthChanged += new System.Windows.Forms.DataGridViewColumnEventHandler(this.dgvMeasurement_ColumnDividerWidthChanged);
            this.dgvMeasurement.Scroll += new System.Windows.Forms.ScrollEventHandler(this.dgvMeasurement_Scroll);
            this.dgvMeasurement.Click += new System.EventHandler(this.dgvMeasurement_Click);
            this.dgvMeasurement.KeyUp += new System.Windows.Forms.KeyEventHandler(this.dgvMeasurement_KeyUp);
            // 
            // splitContainer2
            // 
            this.splitContainer2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.splitContainer2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer2.FixedPanel = System.Windows.Forms.FixedPanel.Panel1;
            this.splitContainer2.Location = new System.Drawing.Point(0, 0);
            this.splitContainer2.Name = "splitContainer2";
            this.splitContainer2.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer2.Panel1
            // 
            this.splitContainer2.Panel1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(239)))), ((int)(((byte)(237)))), ((int)(((byte)(241)))));
            this.splitContainer2.Panel1.Controls.Add(this.pbTDRImage);
            this.splitContainer2.Panel1.Padding = new System.Windows.Forms.Padding(3);
            // 
            // splitContainer2.Panel2
            // 
            this.splitContainer2.Panel2.Padding = new System.Windows.Forms.Padding(3);
            this.splitContainer2.Size = new System.Drawing.Size(848, 730);
            this.splitContainer2.SplitterDistance = 300;
            this.splitContainer2.TabIndex = 0;
            // 
            // pbTDRImage
            // 
            this.pbTDRImage.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pbTDRImage.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pbTDRImage.Location = new System.Drawing.Point(3, 3);
            this.pbTDRImage.Name = "pbTDRImage";
            this.pbTDRImage.Size = new System.Drawing.Size(840, 292);
            this.pbTDRImage.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pbTDRImage.TabIndex = 1;
            this.pbTDRImage.TabStop = false;
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // frmTDR_RCSHistory
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1284, 801);
            this.Controls.Add(this.splitContainer1);
            this.Controls.Add(this.panel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "frmTDR_RCSHistory";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "▣ TDR - RCS 이력";
            this.Load += new System.EventHandler(this.frmTDR_RCSHistory_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvMeasurement)).EndInit();
            this.splitContainer2.Panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer2)).EndInit();
            this.splitContainer2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pbTDRImage)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.ComboBox cboControlRodName;
        private System.Windows.Forms.ComboBox cboCoilName;
        private System.Windows.Forms.ComboBox cboPowerCabinet;
        private System.Windows.Forms.ComboBox cboHogi;
        private System.Windows.Forms.Label label27;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label26;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btnSearch;
        private System.Windows.Forms.SplitContainer splitContainer1;
        public System.Windows.Forms.DataGridView dgvMeasurement;
        private System.Windows.Forms.SplitContainer splitContainer2;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnFolderMove;
        private System.Windows.Forms.Button btnReport;
        private System.Windows.Forms.TextBox teOhDegreeT;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox teOhDegreeF;
        private System.Windows.Forms.PictureBox pbTDRImage;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
    }
}