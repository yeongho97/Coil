using DevExpress.Charts.Native;
using DevExpress.XtraCharts;
using System;
using System.Data;
using System.Windows.Forms;

// 20240411 한인석
// TDR Chart : NI ui -> DevExpress 

namespace Coil_Diagnostor.UserControl
{
    public partial class frmTDRChart : System.Windows.Forms.Form
    {
        private XYDiagram diagram;                              
        private ConstantLine[] cls = new ConstantLine[2];       // 그래프가 다시 그려져도 constantLine Position 유지
        //private Annotation[] ats = new Annotation[2];

        private Series dataSeries;
        private Series markerSeries;
        private double[] cursordiff;
        private bool isCursorMoving = false;

        public frmTDRChart()
        {
            InitializeComponent();
            initControls();
            this.TopLevel = false;
            this.Dock = DockStyle.Fill;
        }

        private void initControls()
        {
            // chart diagram
            diagram = (XYDiagram)chartControl1.Diagram;
            diagram.DefaultPane.BackColor = System.Drawing.Color.DarkCyan;
            diagram.AxisX.GridLines.Visible = true;
            diagram.AxisX.NumericScaleOptions.GridSpacing = 50;

            // Zoom Option
            diagram.EnableAxisXZooming = true;
            diagram.ZoomingOptions.UseTouchDevice = true;   // todo : ?
            diagram.EnableAxisXScrolling = true;
            diagram.AxisX.WholeRange.AutoSideMargins = false;
            diagram.DependentAxesYRange = DevExpress.Utils.DefaultBoolean.False;

            // chart data series
            dataSeries = new Series("data", ViewType.Line);
            markerSeries = new Series("marker", ViewType.Point);

            LineSeriesView lsv = (LineSeriesView)dataSeries.View;
            lsv.Color = System.Drawing.Color.White;
            
            // constant Line : cursor
            cls[0] = new ConstantLine();
            cls[0].LineStyle.DashStyle = DashStyle.Dot;
            cls[0].LineStyle.Thickness = 1;
            cls[0].Color = System.Drawing.Color.Red;
            cls[0].Visible = true;
            cls[0].AxisValue = 0;
            cls[0].RuntimeMoving = true;
            diagram.AxisX.ConstantLines.Add(cls[0]);

            cls[1] = new ConstantLine();
            cls[1].LineStyle.Thickness = 1;
            cls[1].Color = System.Drawing.Color.Red;
            cls[1].Visible = true;
            cls[1].AxisValue = 0;
            cls[1].RuntimeMoving = true;
            diagram.AxisX.ConstantLines.Add(cls[1]);

            // cross : Constant Line - series 
            PointSeriesView psv = (PointSeriesView)markerSeries.View;
            psv.PointMarkerOptions.Kind = MarkerKind.ThinCross;
            psv.PointMarkerOptions.StarPointCount = 5;
            psv.PointMarkerOptions.Size = 20;

            chartControl1.Series.Add(markerSeries);
            chartControl1.Series.Add(dataSeries);
        }

        public int SetData(ref DataTable dt)
        {
            int count = 0;
            if (diagram is null) return count;

            if (dt.Rows.Count > 0)
            {
                dataSeries.Points.Clear();
                markerSeries.Points.Clear();

                cursordiff = new double[dt.Rows.Count];
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    dataSeries.Points.Add(new SeriesPoint(i, dt.Rows[i]["dataValue"]));
                    cursordiff[i] = Convert.ToDouble(dt.Rows[i][0]);
                    count++;
                }

                // Axis X Range
                diagram.AxisX.WholeRange.MinValue = 0;
                diagram.AxisX.WholeRange.MaxValue = dt.Rows.Count - 1;

                //int gridOffset = 10;
                //diagram.AxisX.NumericScaleOptions.GridSpacing = (dt.Rows.Count - gridOffset * 2) / 10;      // todo : check
                //diagram.AxisX.NumericScaleOptions.GridOffset = gridOffset;                                  // todo : check

                // Axis Y Range
                // Auto
                //diagram.AxisX.WholeRange.MaxValue = Math.Ceiling((dMaxValue + 1000d) / 1000d) * 1000d;
                //diagram.AxisX.WholeRange.MinValue = -dMaxYRenge;

                // Initailize Cursor Position
                if (Convert.ToInt32(cls[0].AxisValue) == 0)
                {
                    cls[0].AxisValue = dataSeries.Points.Count / 10;
                }
                if (Convert.ToInt32(cls[1].AxisValue) == 0)
                {
                    cls[1].AxisValue = dataSeries.Points.Count / 10 * 9;
                }

                markerSeries.Points.Add(dataSeries.Points[Convert.ToInt32(cls[0].AxisValue)]);
                markerSeries.Points.Add(dataSeries.Points[Convert.ToInt32(cls[1].AxisValue)]);
            }

            return count;
        }

        public void ClearChartData()
        {
            cls[0].AxisValue = 0;
            cls[1].AxisValue = 0;
            dataSeries.Points.Clear();
            markerSeries.Points.Clear();
        }

        public bool GetCursorDiff()
        {
            bool returnVal = GetCursorCheckedChanged();

            if (returnVal)
            {
                lblResult.Text = "적합";
                lblResult.ForeColor = System.Drawing.Color.Black;
            }
            else
            {
                lblResult.Text = "부적합";
                lblResult.ForeColor = System.Drawing.Color.Red;
            }

            return returnVal;
        }

        private bool GetCursorCheckedChanged()
        {
            bool returnVal = false;
            if (diagram is null) return returnVal;
                
            int idxCursor1 = Convert.ToInt32(diagram.AxisX.ConstantLines[0].AxisValue);
            int idxCursor2 = Convert.ToInt32(diagram.AxisX.ConstantLines[1].AxisValue);

            if (dataSeries.Points.Count < 1) return returnVal;

            //lblCursordiff.Text = (Convert.ToDouble(idxCursor2 - idxCursor1) / 10000).ToString("F3").Trim();
            lblCursordiff.Text = (cursordiff[idxCursor2] - cursordiff[idxCursor1]).ToString("F3").Trim();   // todo : tdr url download data check 
            //lblCursordiff.Text = ((cursordiff[idxCursor2] - cursordiff[idxCursor1]) / 10000).ToString("F3").Trim();

            int intValueCount = 0, intStartIndex = 0, intEndIndex = 0;

            if (idxCursor1 > idxCursor2)
            {
                intValueCount = idxCursor1 - idxCursor2;
                intStartIndex = idxCursor2;
                intEndIndex = idxCursor1;
            }
            else
            {
                intValueCount = idxCursor2 - idxCursor1;
                intStartIndex = idxCursor1;
                intEndIndex = idxCursor2;
            }

            double[] dValue = new double[intValueCount];
            double dMinValue = 0d, dMaxValue = 0d, dBaseValue = 0d;
            int intIndex = 0, intErrorCount = 0;

            dBaseValue = dataSeries.Points[idxCursor1].Values[0];

            for (int i = intStartIndex; i < intEndIndex; i++)
            {
                dValue[intIndex] = dataSeries.Points[i].Values[0];

                if (dMinValue > dValue[intIndex])
                    dMinValue = dValue[intIndex];

                if (dMaxValue < dValue[intIndex])
                    dMaxValue = dValue[intIndex];

                if ((dBaseValue + 100) < dValue[intIndex] || (dBaseValue - 100) > dValue[intIndex])
                    intErrorCount++;

                intIndex++;
            }

            if (intErrorCount > 0)
                returnVal = false;
            else
                returnVal = true;

            return returnVal;
        }

        private void chartControl1_ConstantLineMoved(object sender, ConstantLineMovedEventArgs e)
        {
            if(dataSeries.Points.Count == 0) return;

            isCursorMoving = true;
            GetCursorDiff();
            UpdateCursorLable();
        }

        private void chartControl1_MouseMove(object sender, MouseEventArgs e)
        {
            if(isCursorMoving)
            {
                markerSeries.Points.Clear();

                markerSeries.Points.Add(dataSeries.Points[Convert.ToInt32(cls[0].AxisValue)]);
                markerSeries.Points.Add(dataSeries.Points[Convert.ToInt32(cls[1].AxisValue)]);
            }
        }

        private void chartControl1_MouseUp(object sender, MouseEventArgs e)
        {
            // Ensure Axisvalue be integer
            if (isCursorMoving)
            {
                cls[0].AxisValue = Convert.ToInt32(cls[0].AxisValue);
                cls[1].AxisValue = Convert.ToInt32(cls[1].AxisValue);
            }

            isCursorMoving = false;
        }

        private void UpdateCursorLable()
        {
            if (diagram is null) return;

            lblCursor1.Text = dataSeries.Points[Convert.ToInt32(cls[0].AxisValue)].Values[0].ToString("F3");
            lblCursor2.Text = dataSeries.Points[Convert.ToInt32(cls[1].AxisValue)].Values[0].ToString("F3");
        }
    }
}
