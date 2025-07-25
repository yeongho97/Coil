namespace Coil_Diagnostor
{
    partial class frmMain
    {
        /// <summary>
        /// 필수 디자이너 변수입니다.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 사용 중인 모든 리소스를 정리합니다.
        /// </summary>
        /// <param name="disposing">관리되는 리소스를 삭제해야 하면 true이고, 그렇지 않으면 false입니다.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 디자이너에서 생성한 코드

        /// <summary>
        /// 디자이너 지원에 필요한 메서드입니다.
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마십시오.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmMain));
            this.btnConfiguration = new System.Windows.Forms.Button();
            this.btnExit = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.btnAccessDB = new System.Windows.Forms.Button();
            this.pnlMenu = new System.Windows.Forms.Panel();
            this.tblpnlMenu = new System.Windows.Forms.TableLayoutPanel();
            this.btnDRPIDiagnosis = new System.Windows.Forms.Button();
            this.btnDRPIHistory = new System.Windows.Forms.Button();
            this.btnDRPIReport = new System.Windows.Forms.Button();
            this.btnRCSDiagnosis = new System.Windows.Forms.Button();
            this.btnRCSHistory = new System.Windows.Forms.Button();
            this.btnRCSReport = new System.Windows.Forms.Button();
            this.btnTDR_DRPIDiagnosis = new System.Windows.Forms.Button();
            this.btnTDR_DRPIHistory = new System.Windows.Forms.Button();
            this.btnTDR_RCSDiagnosis = new System.Windows.Forms.Button();
            this.btnTDR_RCSHistory = new System.Windows.Forms.Button();
            this.btnSelfTest = new System.Windows.Forms.Button();
            this.btnSetOffset = new System.Windows.Forms.Button();
            this.pnlMenu.SuspendLayout();
            this.tblpnlMenu.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnConfiguration
            // 
            this.btnConfiguration.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnConfiguration.BackgroundImage")));
            this.btnConfiguration.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnConfiguration.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnConfiguration.Font = new System.Drawing.Font("굴림", 12.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.btnConfiguration.ForeColor = System.Drawing.Color.White;
            this.btnConfiguration.Location = new System.Drawing.Point(950, 25);
            this.btnConfiguration.Name = "btnConfiguration";
            this.btnConfiguration.Size = new System.Drawing.Size(105, 46);
            this.btnConfiguration.TabIndex = 0;
            this.btnConfiguration.Text = "환경 설정";
            this.btnConfiguration.UseVisualStyleBackColor = true;
            this.btnConfiguration.Click += new System.EventHandler(this.btnConfiguration_Click);
            // 
            // btnExit
            // 
            this.btnExit.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnExit.Image = global::Coil_Diagnostor.Properties.Resources.종료_버튼;
            this.btnExit.Location = new System.Drawing.Point(1065, 25);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(46, 46);
            this.btnExit.TabIndex = 15;
            this.btnExit.UseVisualStyleBackColor = true;
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.Transparent;
            this.label1.Font = new System.Drawing.Font("굴림", 32F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.label1.ForeColor = System.Drawing.SystemColors.WindowText;
            this.label1.Location = new System.Drawing.Point(22, 27);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(882, 43);
            this.label1.TabIndex = 14;
            this.label1.Text = "한빛 1,2호기 CRDM/DRPI 코일 성능시험기";
            // 
            // btnAccessDB
            // 
            this.btnAccessDB.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnAccessDB.BackgroundImage")));
            this.btnAccessDB.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnAccessDB.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnAccessDB.Font = new System.Drawing.Font("굴림", 12.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.btnAccessDB.ForeColor = System.Drawing.Color.White;
            this.btnAccessDB.Location = new System.Drawing.Point(789, 25);
            this.btnAccessDB.Name = "btnAccessDB";
            this.btnAccessDB.Size = new System.Drawing.Size(155, 46);
            this.btnAccessDB.TabIndex = 0;
            this.btnAccessDB.Text = "데이터 병합";
            this.btnAccessDB.UseVisualStyleBackColor = true;
            this.btnAccessDB.Visible = false;
            this.btnAccessDB.Click += new System.EventHandler(this.btnAccessDB_Click);
            // 
            // pnlMenu
            // 
            this.pnlMenu.BackColor = System.Drawing.Color.White;
            this.pnlMenu.Controls.Add(this.tblpnlMenu);
            this.pnlMenu.Location = new System.Drawing.Point(21, 639);
            this.pnlMenu.Name = "pnlMenu";
            this.pnlMenu.Size = new System.Drawing.Size(1075, 94);
            this.pnlMenu.TabIndex = 16;
            // 
            // tblpnlMenu
            // 
            this.tblpnlMenu.ColumnCount = 6;
            this.tblpnlMenu.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 16.66667F));
            this.tblpnlMenu.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 16.66667F));
            this.tblpnlMenu.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 16.66667F));
            this.tblpnlMenu.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 16.66667F));
            this.tblpnlMenu.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 16.66667F));
            this.tblpnlMenu.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 16.66667F));
            this.tblpnlMenu.Controls.Add(this.btnDRPIDiagnosis, 0, 1);
            this.tblpnlMenu.Controls.Add(this.btnDRPIHistory, 1, 1);
            this.tblpnlMenu.Controls.Add(this.btnDRPIReport, 2, 1);
            this.tblpnlMenu.Controls.Add(this.btnRCSDiagnosis, 0, 0);
            this.tblpnlMenu.Controls.Add(this.btnRCSHistory, 1, 0);
            this.tblpnlMenu.Controls.Add(this.btnRCSReport, 2, 0);
            this.tblpnlMenu.Controls.Add(this.btnTDR_DRPIDiagnosis, 3, 1);
            this.tblpnlMenu.Controls.Add(this.btnTDR_DRPIHistory, 4, 1);
            this.tblpnlMenu.Controls.Add(this.btnTDR_RCSDiagnosis, 3, 0);
            this.tblpnlMenu.Controls.Add(this.btnTDR_RCSHistory, 4, 0);
            this.tblpnlMenu.Controls.Add(this.btnSelfTest, 5, 1);
            this.tblpnlMenu.Controls.Add(this.btnSetOffset, 5, 0);
            this.tblpnlMenu.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tblpnlMenu.Location = new System.Drawing.Point(0, 0);
            this.tblpnlMenu.Name = "tblpnlMenu";
            this.tblpnlMenu.RowCount = 2;
            this.tblpnlMenu.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tblpnlMenu.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tblpnlMenu.Size = new System.Drawing.Size(1075, 94);
            this.tblpnlMenu.TabIndex = 0;
            // 
            // btnDRPIDiagnosis
            // 
            this.btnDRPIDiagnosis.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnDRPIDiagnosis.BackgroundImage")));
            this.btnDRPIDiagnosis.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnDRPIDiagnosis.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btnDRPIDiagnosis.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnDRPIDiagnosis.Font = new System.Drawing.Font("굴림", 12.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.btnDRPIDiagnosis.ForeColor = System.Drawing.Color.White;
            this.btnDRPIDiagnosis.Location = new System.Drawing.Point(3, 50);
            this.btnDRPIDiagnosis.Name = "btnDRPIDiagnosis";
            this.btnDRPIDiagnosis.Size = new System.Drawing.Size(173, 41);
            this.btnDRPIDiagnosis.TabIndex = 63;
            this.btnDRPIDiagnosis.Text = "DRPI 측정";
            this.btnDRPIDiagnosis.UseVisualStyleBackColor = true;
            this.btnDRPIDiagnosis.Click += new System.EventHandler(this.btnDRPIDiagnosis_Click);
            // 
            // btnDRPIHistory
            // 
            this.btnDRPIHistory.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnDRPIHistory.BackgroundImage")));
            this.btnDRPIHistory.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnDRPIHistory.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btnDRPIHistory.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnDRPIHistory.Font = new System.Drawing.Font("굴림", 12.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.btnDRPIHistory.ForeColor = System.Drawing.Color.White;
            this.btnDRPIHistory.Location = new System.Drawing.Point(182, 50);
            this.btnDRPIHistory.Name = "btnDRPIHistory";
            this.btnDRPIHistory.Size = new System.Drawing.Size(173, 41);
            this.btnDRPIHistory.TabIndex = 64;
            this.btnDRPIHistory.Text = "DRPI 이력";
            this.btnDRPIHistory.UseVisualStyleBackColor = true;
            this.btnDRPIHistory.Click += new System.EventHandler(this.btnDRPIHistory_Click);
            // 
            // btnDRPIReport
            // 
            this.btnDRPIReport.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnDRPIReport.BackgroundImage")));
            this.btnDRPIReport.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnDRPIReport.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btnDRPIReport.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnDRPIReport.Font = new System.Drawing.Font("굴림", 12.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.btnDRPIReport.ForeColor = System.Drawing.Color.White;
            this.btnDRPIReport.Location = new System.Drawing.Point(361, 50);
            this.btnDRPIReport.Name = "btnDRPIReport";
            this.btnDRPIReport.Size = new System.Drawing.Size(173, 41);
            this.btnDRPIReport.TabIndex = 65;
            this.btnDRPIReport.Text = "DRPI 보고서";
            this.btnDRPIReport.UseVisualStyleBackColor = true;
            this.btnDRPIReport.Click += new System.EventHandler(this.btnDRPIReport_Click);
            // 
            // btnRCSDiagnosis
            // 
            this.btnRCSDiagnosis.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnRCSDiagnosis.BackgroundImage")));
            this.btnRCSDiagnosis.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnRCSDiagnosis.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btnRCSDiagnosis.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnRCSDiagnosis.Font = new System.Drawing.Font("굴림", 12.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.btnRCSDiagnosis.ForeColor = System.Drawing.Color.White;
            this.btnRCSDiagnosis.Location = new System.Drawing.Point(3, 3);
            this.btnRCSDiagnosis.Name = "btnRCSDiagnosis";
            this.btnRCSDiagnosis.Size = new System.Drawing.Size(173, 41);
            this.btnRCSDiagnosis.TabIndex = 60;
            this.btnRCSDiagnosis.Text = "RCS 측정";
            this.btnRCSDiagnosis.UseVisualStyleBackColor = true;
            this.btnRCSDiagnosis.Click += new System.EventHandler(this.btnRCSDiagnosis_Click);
            // 
            // btnRCSHistory
            // 
            this.btnRCSHistory.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnRCSHistory.BackgroundImage")));
            this.btnRCSHistory.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnRCSHistory.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btnRCSHistory.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnRCSHistory.Font = new System.Drawing.Font("굴림", 12.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.btnRCSHistory.ForeColor = System.Drawing.Color.White;
            this.btnRCSHistory.Location = new System.Drawing.Point(182, 3);
            this.btnRCSHistory.Name = "btnRCSHistory";
            this.btnRCSHistory.Size = new System.Drawing.Size(173, 41);
            this.btnRCSHistory.TabIndex = 61;
            this.btnRCSHistory.Text = "RCS 이력";
            this.btnRCSHistory.UseVisualStyleBackColor = true;
            this.btnRCSHistory.Click += new System.EventHandler(this.btnRCSHistory_Click);
            // 
            // btnRCSReport
            // 
            this.btnRCSReport.BackgroundImage = global::Coil_Diagnostor.Properties.Resources.버튼;
            this.btnRCSReport.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnRCSReport.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btnRCSReport.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnRCSReport.Font = new System.Drawing.Font("굴림", 12.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.btnRCSReport.ForeColor = System.Drawing.Color.White;
            this.btnRCSReport.Location = new System.Drawing.Point(361, 3);
            this.btnRCSReport.Name = "btnRCSReport";
            this.btnRCSReport.Size = new System.Drawing.Size(173, 41);
            this.btnRCSReport.TabIndex = 62;
            this.btnRCSReport.Text = "RCS 보고서";
            this.btnRCSReport.UseVisualStyleBackColor = true;
            this.btnRCSReport.Click += new System.EventHandler(this.btnRCSReport_Click);
            // 
            // btnTDR_DRPIDiagnosis
            // 
            this.btnTDR_DRPIDiagnosis.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnTDR_DRPIDiagnosis.BackgroundImage")));
            this.btnTDR_DRPIDiagnosis.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnTDR_DRPIDiagnosis.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btnTDR_DRPIDiagnosis.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnTDR_DRPIDiagnosis.Font = new System.Drawing.Font("굴림", 12.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.btnTDR_DRPIDiagnosis.ForeColor = System.Drawing.Color.White;
            this.btnTDR_DRPIDiagnosis.Location = new System.Drawing.Point(540, 50);
            this.btnTDR_DRPIDiagnosis.Name = "btnTDR_DRPIDiagnosis";
            this.btnTDR_DRPIDiagnosis.Size = new System.Drawing.Size(173, 41);
            this.btnTDR_DRPIDiagnosis.TabIndex = 57;
            this.btnTDR_DRPIDiagnosis.Text = "TDR-DRPI 측정";
            this.btnTDR_DRPIDiagnosis.UseVisualStyleBackColor = true;
            this.btnTDR_DRPIDiagnosis.Click += new System.EventHandler(this.btnTDR_DRPIDiagnosis_Click);
            // 
            // btnTDR_DRPIHistory
            // 
            this.btnTDR_DRPIHistory.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnTDR_DRPIHistory.BackgroundImage")));
            this.btnTDR_DRPIHistory.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnTDR_DRPIHistory.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btnTDR_DRPIHistory.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnTDR_DRPIHistory.Font = new System.Drawing.Font("굴림", 12.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.btnTDR_DRPIHistory.ForeColor = System.Drawing.Color.White;
            this.btnTDR_DRPIHistory.Location = new System.Drawing.Point(719, 50);
            this.btnTDR_DRPIHistory.Name = "btnTDR_DRPIHistory";
            this.btnTDR_DRPIHistory.Size = new System.Drawing.Size(173, 41);
            this.btnTDR_DRPIHistory.TabIndex = 59;
            this.btnTDR_DRPIHistory.Text = "TDR-DRPI 이력";
            this.btnTDR_DRPIHistory.UseVisualStyleBackColor = true;
            this.btnTDR_DRPIHistory.Click += new System.EventHandler(this.btnTDR_DRPIHistory_Click);
            // 
            // btnTDR_RCSDiagnosis
            // 
            this.btnTDR_RCSDiagnosis.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnTDR_RCSDiagnosis.BackgroundImage")));
            this.btnTDR_RCSDiagnosis.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnTDR_RCSDiagnosis.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btnTDR_RCSDiagnosis.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnTDR_RCSDiagnosis.Font = new System.Drawing.Font("굴림", 12.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.btnTDR_RCSDiagnosis.ForeColor = System.Drawing.Color.White;
            this.btnTDR_RCSDiagnosis.Location = new System.Drawing.Point(540, 3);
            this.btnTDR_RCSDiagnosis.Name = "btnTDR_RCSDiagnosis";
            this.btnTDR_RCSDiagnosis.Size = new System.Drawing.Size(173, 41);
            this.btnTDR_RCSDiagnosis.TabIndex = 56;
            this.btnTDR_RCSDiagnosis.Text = "TDR-RCS 측정";
            this.btnTDR_RCSDiagnosis.UseVisualStyleBackColor = true;
            this.btnTDR_RCSDiagnosis.Click += new System.EventHandler(this.btnTDR_RCSDiagnosis_Click);
            // 
            // btnTDR_RCSHistory
            // 
            this.btnTDR_RCSHistory.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnTDR_RCSHistory.BackgroundImage")));
            this.btnTDR_RCSHistory.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnTDR_RCSHistory.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btnTDR_RCSHistory.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnTDR_RCSHistory.Font = new System.Drawing.Font("굴림", 12.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.btnTDR_RCSHistory.ForeColor = System.Drawing.Color.White;
            this.btnTDR_RCSHistory.Location = new System.Drawing.Point(719, 3);
            this.btnTDR_RCSHistory.Name = "btnTDR_RCSHistory";
            this.btnTDR_RCSHistory.Size = new System.Drawing.Size(173, 41);
            this.btnTDR_RCSHistory.TabIndex = 58;
            this.btnTDR_RCSHistory.Text = "TDR-RCS 이력";
            this.btnTDR_RCSHistory.UseVisualStyleBackColor = true;
            this.btnTDR_RCSHistory.Click += new System.EventHandler(this.btnTDR_RCSHistory_Click);
            // 
            // btnSelfTest
            // 
            this.btnSelfTest.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnSelfTest.BackgroundImage")));
            this.btnSelfTest.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnSelfTest.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btnSelfTest.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnSelfTest.Font = new System.Drawing.Font("굴림", 12.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.btnSelfTest.ForeColor = System.Drawing.Color.White;
            this.btnSelfTest.Location = new System.Drawing.Point(898, 50);
            this.btnSelfTest.Name = "btnSelfTest";
            this.btnSelfTest.Size = new System.Drawing.Size(174, 41);
            this.btnSelfTest.TabIndex = 44;
            this.btnSelfTest.Text = "Self Test";
            this.btnSelfTest.UseVisualStyleBackColor = true;
            this.btnSelfTest.Click += new System.EventHandler(this.btnSelfTest_Click);
            // 
            // btnSetOffset
            // 
            this.btnSetOffset.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnSetOffset.BackgroundImage")));
            this.btnSetOffset.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnSetOffset.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btnSetOffset.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnSetOffset.Font = new System.Drawing.Font("굴림", 12.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.btnSetOffset.ForeColor = System.Drawing.Color.White;
            this.btnSetOffset.Location = new System.Drawing.Point(898, 3);
            this.btnSetOffset.Name = "btnSetOffset";
            this.btnSetOffset.Size = new System.Drawing.Size(174, 41);
            this.btnSetOffset.TabIndex = 45;
            this.btnSetOffset.Text = "영점 보정";
            this.btnSetOffset.UseVisualStyleBackColor = true;
            this.btnSetOffset.Click += new System.EventHandler(this.btnSetOffset_Click);
            // 
            // frmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = global::Coil_Diagnostor.Properties.Resources.HB_Main;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.ClientSize = new System.Drawing.Size(1120, 760);
            this.Controls.Add(this.pnlMenu);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.btnConfiguration);
            this.Controls.Add(this.btnAccessDB);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "한울1, 2호기 계측설비 주요 릴레이 성능시험장비";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmMain_FormClosing);
            this.Load += new System.EventHandler(this.frmMain_Load);
            this.pnlMenu.ResumeLayout(false);
            this.tblpnlMenu.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button btnConfiguration;
        private System.Windows.Forms.Button btnExit;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnAccessDB;
        private System.Windows.Forms.Panel pnlMenu;
        private System.Windows.Forms.TableLayoutPanel tblpnlMenu;
        private System.Windows.Forms.Button btnDRPIDiagnosis;
        private System.Windows.Forms.Button btnDRPIHistory;
        private System.Windows.Forms.Button btnDRPIReport;
        private System.Windows.Forms.Button btnRCSDiagnosis;
        private System.Windows.Forms.Button btnRCSHistory;
        private System.Windows.Forms.Button btnRCSReport;
        private System.Windows.Forms.Button btnTDR_DRPIDiagnosis;
        private System.Windows.Forms.Button btnTDR_DRPIHistory;
        private System.Windows.Forms.Button btnTDR_RCSDiagnosis;
        private System.Windows.Forms.Button btnTDR_RCSHistory;
        private System.Windows.Forms.Button btnSelfTest;
        private System.Windows.Forms.Button btnSetOffset;
    }
}

