namespace Coil_Diagnostor
{
    partial class frmFilePathSetting
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmFilePathSetting));
            this.teFilePathSetting = new System.Windows.Forms.TextBox();
            this.btnClose = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.btnFilePathSetting = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // teFilePathSetting
            // 
            this.teFilePathSetting.Location = new System.Drawing.Point(137, 22);
            this.teFilePathSetting.Name = "teFilePathSetting";
            this.teFilePathSetting.ReadOnly = true;
            this.teFilePathSetting.Size = new System.Drawing.Size(407, 21);
            this.teFilePathSetting.TabIndex = 39;
            // 
            // btnClose
            // 
            this.btnClose.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnClose.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnClose.BackgroundImage")));
            this.btnClose.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnClose.Font = new System.Drawing.Font("굴림", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.btnClose.ForeColor = System.Drawing.Color.White;
            this.btnClose.Location = new System.Drawing.Point(231, 58);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(130, 40);
            this.btnClose.TabIndex = 41;
            this.btnClose.Text = "닫 기 (F12)";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.Color.Transparent;
            this.label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label1.Location = new System.Drawing.Point(12, 21);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(120, 23);
            this.label1.TabIndex = 42;
            this.label1.Text = "기준 프로그램";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // btnFilePathSetting
            // 
            this.btnFilePathSetting.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnFilePathSetting.BackgroundImage = global::Coil_Diagnostor.Properties.Resources.Folder_Add_icon;
            this.btnFilePathSetting.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnFilePathSetting.Font = new System.Drawing.Font("굴림", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.btnFilePathSetting.ForeColor = System.Drawing.Color.White;
            this.btnFilePathSetting.Location = new System.Drawing.Point(545, 17);
            this.btnFilePathSetting.Name = "btnFilePathSetting";
            this.btnFilePathSetting.Size = new System.Drawing.Size(30, 30);
            this.btnFilePathSetting.TabIndex = 40;
            this.btnFilePathSetting.Text = "가져오기 (F2)";
            this.btnFilePathSetting.UseVisualStyleBackColor = true;
            this.btnFilePathSetting.Click += new System.EventHandler(this.btnFilePathSetting_Click);
            // 
            // frmFilePathSetting
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(215)))), ((int)(((byte)(231)))), ((int)(((byte)(252)))));
            this.ClientSize = new System.Drawing.Size(584, 110);
            this.Controls.Add(this.teFilePathSetting);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnFilePathSetting);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "frmFilePathSetting";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "기존 파일 경로 설정";
            this.Load += new System.EventHandler(this.frmFilePathSetting_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox teFilePathSetting;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnFilePathSetting;
    }
}