namespace Coil_Diagnostor.UserControl
{
    partial class frmTDRChart
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
            DevExpress.XtraCharts.XYDiagram xyDiagram1 = new DevExpress.XtraCharts.XYDiagram();
            DevExpress.XtraCharts.Series series1 = new DevExpress.XtraCharts.Series();
            DevExpress.XtraCharts.SplineSeriesView splineSeriesView1 = new DevExpress.XtraCharts.SplineSeriesView();
            this.pnlTopInfo = new System.Windows.Forms.Panel();
            this.lblCursor2 = new System.Windows.Forms.Label();
            this.lblResult = new System.Windows.Forms.Label();
            this.lblCursordiff = new System.Windows.Forms.Label();
            this.lblCursor1 = new System.Windows.Forms.Label();
            this.pnlChart = new System.Windows.Forms.Panel();
            this.chartControl1 = new DevExpress.XtraCharts.ChartControl();
            this.pnlTopInfo.SuspendLayout();
            this.pnlChart.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.chartControl1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(xyDiagram1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(series1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(splineSeriesView1)).BeginInit();
            this.SuspendLayout();
            // 
            // pnlTopInfo
            // 
            this.pnlTopInfo.BackColor = System.Drawing.SystemColors.ButtonShadow;
            this.pnlTopInfo.Controls.Add(this.lblCursor2);
            this.pnlTopInfo.Controls.Add(this.lblResult);
            this.pnlTopInfo.Controls.Add(this.lblCursordiff);
            this.pnlTopInfo.Controls.Add(this.lblCursor1);
            this.pnlTopInfo.Dock = System.Windows.Forms.DockStyle.Top;
            this.pnlTopInfo.Location = new System.Drawing.Point(0, 0);
            this.pnlTopInfo.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.pnlTopInfo.Name = "pnlTopInfo";
            this.pnlTopInfo.Size = new System.Drawing.Size(942, 61);
            this.pnlTopInfo.TabIndex = 0;
            // 
            // lblCursor2
            // 
            this.lblCursor2.AutoSize = true;
            this.lblCursor2.Location = new System.Drawing.Point(562, 35);
            this.lblCursor2.Name = "lblCursor2";
            this.lblCursor2.Size = new System.Drawing.Size(28, 15);
            this.lblCursor2.TabIndex = 10;
            this.lblCursor2.Text = "0.0";
            // 
            // lblResult
            // 
            this.lblResult.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.lblResult.BackColor = System.Drawing.Color.White;
            this.lblResult.Font = new System.Drawing.Font("굴림", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.lblResult.ForeColor = System.Drawing.Color.Lime;
            this.lblResult.Location = new System.Drawing.Point(819, 15);
            this.lblResult.Name = "lblResult";
            this.lblResult.Size = new System.Drawing.Size(109, 31);
            this.lblResult.TabIndex = 8;
            this.lblResult.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // lblCursordiff
            // 
            this.lblCursordiff.BackColor = System.Drawing.Color.Transparent;
            this.lblCursordiff.Font = new System.Drawing.Font("굴림", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.lblCursordiff.ForeColor = System.Drawing.Color.Lime;
            this.lblCursordiff.Location = new System.Drawing.Point(675, 15);
            this.lblCursordiff.Name = "lblCursordiff";
            this.lblCursordiff.Size = new System.Drawing.Size(109, 31);
            this.lblCursordiff.TabIndex = 9;
            this.lblCursordiff.Text = "0.0";
            this.lblCursordiff.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // lblCursor1
            // 
            this.lblCursor1.AutoSize = true;
            this.lblCursor1.Location = new System.Drawing.Point(562, 12);
            this.lblCursor1.Name = "lblCursor1";
            this.lblCursor1.Size = new System.Drawing.Size(28, 15);
            this.lblCursor1.TabIndex = 0;
            this.lblCursor1.Text = "0.0";
            // 
            // pnlChart
            // 
            this.pnlChart.Controls.Add(this.chartControl1);
            this.pnlChart.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlChart.Location = new System.Drawing.Point(0, 61);
            this.pnlChart.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.pnlChart.Name = "pnlChart";
            this.pnlChart.Size = new System.Drawing.Size(942, 415);
            this.pnlChart.TabIndex = 1;
            // 
            // chartControl1
            // 
            xyDiagram1.AxisX.VisibleInPanesSerializable = "-1";
            xyDiagram1.AxisY.VisibleInPanesSerializable = "-1";
            this.chartControl1.Diagram = xyDiagram1;
            this.chartControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.chartControl1.Legend.LegendID = -1;
            this.chartControl1.Legend.Visibility = DevExpress.Utils.DefaultBoolean.False;
            this.chartControl1.Location = new System.Drawing.Point(0, 0);
            this.chartControl1.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.chartControl1.Name = "chartControl1";
            series1.Name = "Series 1";
            series1.SeriesID = 0;
            series1.View = splineSeriesView1;
            this.chartControl1.SeriesSerializable = new DevExpress.XtraCharts.Series[] {
        series1};
            this.chartControl1.Size = new System.Drawing.Size(942, 415);
            this.chartControl1.TabIndex = 0;
            this.chartControl1.ConstantLineMoved += new DevExpress.XtraCharts.ConstantLineMovedEventHandler(this.chartControl1_ConstantLineMoved);
            this.chartControl1.MouseMove += new System.Windows.Forms.MouseEventHandler(this.chartControl1_MouseMove);
            this.chartControl1.MouseUp += new System.Windows.Forms.MouseEventHandler(this.chartControl1_MouseUp);
            // 
            // frmTDRChart
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(942, 476);
            this.Controls.Add(this.pnlChart);
            this.Controls.Add(this.pnlTopInfo);
            this.DoubleBuffered = true;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Name = "frmTDRChart";
            this.pnlTopInfo.ResumeLayout(false);
            this.pnlTopInfo.PerformLayout();
            this.pnlChart.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(xyDiagram1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(splineSeriesView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(series1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.chartControl1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel pnlTopInfo;
        private System.Windows.Forms.Panel pnlChart;
        private System.Windows.Forms.Label lblCursor1;
        private System.Windows.Forms.Label lblCursor2;
        private System.Windows.Forms.Label lblResult;
        private System.Windows.Forms.Label lblCursordiff;
        private DevExpress.XtraCharts.ChartControl chartControl1;
    }
}