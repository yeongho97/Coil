namespace Coil_Diagnostor
{
    partial class frmTDR_RCSDiagnosis
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmTDR_RCSDiagnosis));
            this.panel2 = new System.Windows.Forms.Panel();
            this.panel9 = new System.Windows.Forms.Panel();
            this.pbTDRImage = new System.Windows.Forms.PictureBox();
            this.panel4 = new System.Windows.Forms.Panel();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.ledTDRMeter = new Coil_DiagnoseDlg.LedBulb();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.rboMeasurementItem_Move = new System.Windows.Forms.RadioButton();
            this.rboMeasurementItem_Up = new System.Windows.Forms.RadioButton();
            this.rboMeasurementItem_Stop = new System.Windows.Forms.RadioButton();
            this.groupBox7 = new System.Windows.Forms.GroupBox();
            this.btnClose = new System.Windows.Forms.Button();
            this.btnMeasurementStart = new System.Windows.Forms.Button();
            this.groupBox8 = new System.Windows.Forms.GroupBox();
            this.teOhDegree2 = new System.Windows.Forms.TextBox();
            this.lblOhDegree3 = new System.Windows.Forms.Label();
            this.lblOhDegree1 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.cboHogi = new System.Windows.Forms.ComboBox();
            this.panel5 = new System.Windows.Forms.Panel();
            this.panel3 = new System.Windows.Forms.Panel();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.panel2.SuspendLayout();
            this.panel9.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pbTDRImage)).BeginInit();
            this.panel4.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox7.SuspendLayout();
            this.groupBox8.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.panel5.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.panel9);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(0, 0);
            this.panel2.Name = "panel2";
            this.panel2.Padding = new System.Windows.Forms.Padding(5);
            this.panel2.Size = new System.Drawing.Size(1104, 671);
            this.panel2.TabIndex = 1;
            // 
            // panel9
            // 
            this.panel9.Controls.Add(this.pbTDRImage);
            this.panel9.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel9.Location = new System.Drawing.Point(5, 5);
            this.panel9.Name = "panel9";
            this.panel9.Size = new System.Drawing.Size(1094, 661);
            this.panel9.TabIndex = 3;
            // 
            // pbTDRImage
            // 
            this.pbTDRImage.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pbTDRImage.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pbTDRImage.Location = new System.Drawing.Point(0, 0);
            this.pbTDRImage.Margin = new System.Windows.Forms.Padding(0);
            this.pbTDRImage.Name = "pbTDRImage";
            this.pbTDRImage.Size = new System.Drawing.Size(1094, 661);
            this.pbTDRImage.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pbTDRImage.TabIndex = 0;
            this.pbTDRImage.TabStop = false;
            // 
            // panel4
            // 
            this.panel4.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(215)))), ((int)(((byte)(231)))), ((int)(((byte)(252)))));
            this.panel4.Controls.Add(this.groupBox4);
            this.panel4.Controls.Add(this.groupBox2);
            this.panel4.Controls.Add(this.groupBox7);
            this.panel4.Controls.Add(this.groupBox8);
            this.panel4.Controls.Add(this.groupBox1);
            this.panel4.Dock = System.Windows.Forms.DockStyle.Right;
            this.panel4.Location = new System.Drawing.Point(1104, 0);
            this.panel4.Name = "panel4";
            this.panel4.Padding = new System.Windows.Forms.Padding(2, 5, 5, 5);
            this.panel4.Size = new System.Drawing.Size(180, 801);
            this.panel4.TabIndex = 1;
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.ledTDRMeter);
            this.groupBox4.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.groupBox4.Location = new System.Drawing.Point(2, 609);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(173, 71);
            this.groupBox4.TabIndex = 21;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "TDR-Meter 통신상태";
            // 
            // ledTDRMeter
            // 
            this.ledTDRMeter.Location = new System.Drawing.Point(61, 20);
            this.ledTDRMeter.Margin = new System.Windows.Forms.Padding(2);
            this.ledTDRMeter.Name = "ledTDRMeter";
            this.ledTDRMeter.On = false;
            this.ledTDRMeter.Size = new System.Drawing.Size(45, 45);
            this.ledTDRMeter.TabIndex = 1;
            this.ledTDRMeter.Text = "ledBulb1";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.rboMeasurementItem_Move);
            this.groupBox2.Controls.Add(this.rboMeasurementItem_Up);
            this.groupBox2.Controls.Add(this.rboMeasurementItem_Stop);
            this.groupBox2.Dock = System.Windows.Forms.DockStyle.Top;
            this.groupBox2.Location = new System.Drawing.Point(2, 94);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(173, 114);
            this.groupBox2.TabIndex = 19;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "측정 항목";
            // 
            // rboMeasurementItem_Move
            // 
            this.rboMeasurementItem_Move.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.rboMeasurementItem_Move.Location = new System.Drawing.Point(28, 80);
            this.rboMeasurementItem_Move.Name = "rboMeasurementItem_Move";
            this.rboMeasurementItem_Move.Size = new System.Drawing.Size(123, 24);
            this.rboMeasurementItem_Move.TabIndex = 0;
            this.rboMeasurementItem_Move.Text = "이                  동";
            this.rboMeasurementItem_Move.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.rboMeasurementItem_Move.UseVisualStyleBackColor = true;
            this.rboMeasurementItem_Move.CheckedChanged += new System.EventHandler(this.rboMeasurementItem_CheckedChanged);
            // 
            // rboMeasurementItem_Up
            // 
            this.rboMeasurementItem_Up.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.rboMeasurementItem_Up.Location = new System.Drawing.Point(28, 50);
            this.rboMeasurementItem_Up.Name = "rboMeasurementItem_Up";
            this.rboMeasurementItem_Up.Size = new System.Drawing.Size(123, 24);
            this.rboMeasurementItem_Up.TabIndex = 0;
            this.rboMeasurementItem_Up.Text = "올                  림";
            this.rboMeasurementItem_Up.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.rboMeasurementItem_Up.UseVisualStyleBackColor = true;
            this.rboMeasurementItem_Up.CheckedChanged += new System.EventHandler(this.rboMeasurementItem_CheckedChanged);
            // 
            // rboMeasurementItem_Stop
            // 
            this.rboMeasurementItem_Stop.Checked = true;
            this.rboMeasurementItem_Stop.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.rboMeasurementItem_Stop.Location = new System.Drawing.Point(28, 20);
            this.rboMeasurementItem_Stop.Name = "rboMeasurementItem_Stop";
            this.rboMeasurementItem_Stop.Size = new System.Drawing.Size(123, 24);
            this.rboMeasurementItem_Stop.TabIndex = 0;
            this.rboMeasurementItem_Stop.TabStop = true;
            this.rboMeasurementItem_Stop.Text = "정                  지";
            this.rboMeasurementItem_Stop.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.rboMeasurementItem_Stop.UseVisualStyleBackColor = true;
            this.rboMeasurementItem_Stop.CheckedChanged += new System.EventHandler(this.rboMeasurementItem_CheckedChanged);
            // 
            // groupBox7
            // 
            this.groupBox7.Controls.Add(this.btnClose);
            this.groupBox7.Controls.Add(this.btnMeasurementStart);
            this.groupBox7.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.groupBox7.Location = new System.Drawing.Point(2, 680);
            this.groupBox7.Name = "groupBox7";
            this.groupBox7.Size = new System.Drawing.Size(173, 116);
            this.groupBox7.TabIndex = 18;
            this.groupBox7.TabStop = false;
            // 
            // btnClose
            // 
            this.btnClose.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.btnClose.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnClose.BackgroundImage")));
            this.btnClose.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnClose.Font = new System.Drawing.Font("굴림", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.btnClose.ForeColor = System.Drawing.Color.White;
            this.btnClose.Location = new System.Drawing.Point(5, 65);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(161, 50);
            this.btnClose.TabIndex = 4;
            this.btnClose.Text = "닫 기 (F12)";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnMeasurementStart
            // 
            this.btnMeasurementStart.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.btnMeasurementStart.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnMeasurementStart.BackgroundImage")));
            this.btnMeasurementStart.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnMeasurementStart.Font = new System.Drawing.Font("굴림", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.btnMeasurementStart.ForeColor = System.Drawing.Color.White;
            this.btnMeasurementStart.Location = new System.Drawing.Point(5, 14);
            this.btnMeasurementStart.Name = "btnMeasurementStart";
            this.btnMeasurementStart.Size = new System.Drawing.Size(161, 50);
            this.btnMeasurementStart.TabIndex = 0;
            this.btnMeasurementStart.Text = "측   정 (F2)";
            this.btnMeasurementStart.UseVisualStyleBackColor = true;
            this.btnMeasurementStart.Click += new System.EventHandler(this.btnMeasurementStart_Click);
            // 
            // groupBox8
            // 
            this.groupBox8.Controls.Add(this.teOhDegree2);
            this.groupBox8.Controls.Add(this.lblOhDegree3);
            this.groupBox8.Controls.Add(this.lblOhDegree1);
            this.groupBox8.Dock = System.Windows.Forms.DockStyle.Top;
            this.groupBox8.Location = new System.Drawing.Point(2, 49);
            this.groupBox8.Name = "groupBox8";
            this.groupBox8.Size = new System.Drawing.Size(173, 45);
            this.groupBox8.TabIndex = 10;
            this.groupBox8.TabStop = false;
            this.groupBox8.Text = "OH 차수";
            // 
            // teOhDegree2
            // 
            this.teOhDegree2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.teOhDegree2.Location = new System.Drawing.Point(41, 17);
            this.teOhDegree2.Name = "teOhDegree2";
            this.teOhDegree2.Size = new System.Drawing.Size(92, 21);
            this.teOhDegree2.TabIndex = 0;
            this.teOhDegree2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.teOhDegree2.TextChanged += new System.EventHandler(this.teOhDegree2_TextChanged);
            // 
            // lblOhDegree3
            // 
            this.lblOhDegree3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.lblOhDegree3.BackColor = System.Drawing.Color.Transparent;
            this.lblOhDegree3.Location = new System.Drawing.Point(135, 15);
            this.lblOhDegree3.Name = "lblOhDegree3";
            this.lblOhDegree3.Size = new System.Drawing.Size(33, 23);
            this.lblOhDegree3.TabIndex = 36;
            this.lblOhDegree3.Text = "차";
            this.lblOhDegree3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // lblOhDegree1
            // 
            this.lblOhDegree1.BackColor = System.Drawing.Color.Transparent;
            this.lblOhDegree1.Location = new System.Drawing.Point(6, 15);
            this.lblOhDegree1.Name = "lblOhDegree1";
            this.lblOhDegree1.Size = new System.Drawing.Size(38, 23);
            this.lblOhDegree1.TabIndex = 36;
            this.lblOhDegree1.Text = "제";
            this.lblOhDegree1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.cboHogi);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Top;
            this.groupBox1.Location = new System.Drawing.Point(2, 5);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(173, 44);
            this.groupBox1.TabIndex = 9;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "호기";
            // 
            // cboHogi
            // 
            this.cboHogi.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cboHogi.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboHogi.FormattingEnabled = true;
            this.cboHogi.Location = new System.Drawing.Point(6, 16);
            this.cboHogi.Name = "cboHogi";
            this.cboHogi.Size = new System.Drawing.Size(162, 20);
            this.cboHogi.TabIndex = 0;
            this.cboHogi.SelectedIndexChanged += new System.EventHandler(this.cboHogi_SelectedIndexChanged);
            // 
            // panel5
            // 
            this.panel5.Controls.Add(this.panel2);
            this.panel5.Controls.Add(this.panel3);
            this.panel5.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel5.Location = new System.Drawing.Point(0, 0);
            this.panel5.Name = "panel5";
            this.panel5.Size = new System.Drawing.Size(1104, 801);
            this.panel5.TabIndex = 2;
            // 
            // panel3
            // 
            this.panel3.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel3.Location = new System.Drawing.Point(0, 671);
            this.panel3.Margin = new System.Windows.Forms.Padding(3, 5, 3, 3);
            this.panel3.Name = "panel3";
            this.panel3.Padding = new System.Windows.Forms.Padding(5);
            this.panel3.Size = new System.Drawing.Size(1104, 130);
            this.panel3.TabIndex = 1;
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // frmTDR_RCSDiagnosis
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1284, 801);
            this.Controls.Add(this.panel5);
            this.Controls.Add(this.panel4);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "frmTDR_RCSDiagnosis";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "▣ TDR - RCS 측정";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmTDR_RCSDiagnosis_FormClosing);
            this.Load += new System.EventHandler(this.frmTDR_RCSDiagnosis_Load);
            this.panel2.ResumeLayout(false);
            this.panel9.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pbTDRImage)).EndInit();
            this.panel4.ResumeLayout(false);
            this.groupBox4.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox7.ResumeLayout(false);
            this.groupBox8.ResumeLayout(false);
            this.groupBox8.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.panel5.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Panel panel4;
        private System.Windows.Forms.GroupBox groupBox7;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnMeasurementStart;
        private System.Windows.Forms.GroupBox groupBox8;
        private System.Windows.Forms.Label lblOhDegree3;
        private System.Windows.Forms.Label lblOhDegree1;
        public System.Windows.Forms.TextBox teOhDegree2;
        private System.Windows.Forms.GroupBox groupBox1;
        public System.Windows.Forms.ComboBox cboHogi;
        private System.Windows.Forms.Panel panel5;
        private System.Windows.Forms.Panel panel9;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.RadioButton rboMeasurementItem_Move;
        private System.Windows.Forms.RadioButton rboMeasurementItem_Up;
        private System.Windows.Forms.RadioButton rboMeasurementItem_Stop;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.PictureBox pbTDRImage;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private Coil_DiagnoseDlg.LedBulb ledTDRMeter;
    }
}