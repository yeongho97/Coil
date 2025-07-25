namespace Coil_Diagnostor
{
    partial class frmMainChanged
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmMainChanged));
            this.btnCoilDiagnostor = new System.Windows.Forms.Button();
            this.btnExit = new System.Windows.Forms.Button();
            this.btnMain = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.btnFilePathSetting = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnCoilDiagnostor
            // 
            this.btnCoilDiagnostor.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnCoilDiagnostor.BackgroundImage")));
            this.btnCoilDiagnostor.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnCoilDiagnostor.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnCoilDiagnostor.Font = new System.Drawing.Font("굴림", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.btnCoilDiagnostor.ForeColor = System.Drawing.Color.White;
            this.btnCoilDiagnostor.Location = new System.Drawing.Point(626, 12);
            this.btnCoilDiagnostor.Name = "btnCoilDiagnostor";
            this.btnCoilDiagnostor.Size = new System.Drawing.Size(140, 50);
            this.btnCoilDiagnostor.TabIndex = 2;
            this.btnCoilDiagnostor.Text = "기존 진단기 (F3)";
            this.btnCoilDiagnostor.UseVisualStyleBackColor = true;
            this.btnCoilDiagnostor.Click += new System.EventHandler(this.btnCoilDiagnostor_Click);
            // 
            // btnExit
            // 
            this.btnExit.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnExit.Image = global::Coil_Diagnostor.Properties.Resources.종료_버튼;
            this.btnExit.Location = new System.Drawing.Point(772, 14);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(46, 46);
            this.btnExit.TabIndex = 3;
            this.btnExit.UseVisualStyleBackColor = true;
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // btnMain
            // 
            this.btnMain.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnMain.BackgroundImage")));
            this.btnMain.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnMain.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnMain.Font = new System.Drawing.Font("굴림", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.btnMain.ForeColor = System.Drawing.Color.White;
            this.btnMain.Location = new System.Drawing.Point(482, 12);
            this.btnMain.Name = "btnMain";
            this.btnMain.Size = new System.Drawing.Size(140, 50);
            this.btnMain.TabIndex = 1;
            this.btnMain.Text = "신규 진단기 (F2)";
            this.btnMain.UseVisualStyleBackColor = true;
            this.btnMain.Click += new System.EventHandler(this.btnMain_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.Transparent;
            this.label1.Font = new System.Drawing.Font("굴림", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.label1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.label1.Location = new System.Drawing.Point(265, 96);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(486, 24);
            this.label1.TabIndex = 14;
            this.label1.Text = "한빛 1,2호기 CRDM/DRPI 코일 성능시험기";
            // 
            // btnFilePathSetting
            // 
            this.btnFilePathSetting.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnFilePathSetting.BackgroundImage")));
            this.btnFilePathSetting.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnFilePathSetting.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnFilePathSetting.Font = new System.Drawing.Font("굴림", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.btnFilePathSetting.ForeColor = System.Drawing.Color.White;
            this.btnFilePathSetting.Location = new System.Drawing.Point(338, 12);
            this.btnFilePathSetting.Name = "btnFilePathSetting";
            this.btnFilePathSetting.Size = new System.Drawing.Size(140, 50);
            this.btnFilePathSetting.TabIndex = 0;
            this.btnFilePathSetting.Text = "환경 설정 (F1)";
            this.btnFilePathSetting.UseVisualStyleBackColor = true;
            this.btnFilePathSetting.Click += new System.EventHandler(this.btnFilePathSetting_Click);
            // 
            // frmMainChanged
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("$this.BackgroundImage")));
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.ClientSize = new System.Drawing.Size(830, 347);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnMain);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.btnFilePathSetting);
            this.Controls.Add(this.btnCoilDiagnostor);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmMainChanged";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "한울1, 2호기 계측설비 주요 릴레이 성능시험장비";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmMain_FormClosing);
            this.Load += new System.EventHandler(this.frmMain_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnCoilDiagnostor;
        private System.Windows.Forms.Button btnExit;
        private System.Windows.Forms.Button btnMain;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnFilePathSetting;
    }
}

