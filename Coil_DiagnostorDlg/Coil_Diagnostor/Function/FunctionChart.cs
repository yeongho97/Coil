using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms.DataVisualization.Charting;

namespace Coil_Diagnostor.Function
{
    public class FunctionChart
    {
        //======================================
        // 차트 초기화.
        //======================================
        private string MakeAreaName(byte bteAreaNo)
        {
            return "ca" + bteAreaNo.ToString(); // 범례 접근명.
        }

        private void ClearChart(ref Chart cht)
        {
            try
            {
                cht.Series.Clear(); //시리즈 삭제
                cht.ChartAreas.Clear(); //차트 에어리어 삭제.
                cht.Titles.Clear();
                cht.Legends.Clear();
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
        }

        public void InitialChart(ref Chart cht, string strChartTitle, byte bteChartAreaCount)
        {
            try
            {
                // 이전 내용을 지운다.
                ClearChart(ref cht);

                // 차트 배경색 
                cht.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(211)))), ((int)(((byte)(223)))), ((int)(((byte)(240))))); //배경색
                cht.BackGradientStyle = System.Windows.Forms.DataVisualization.Charting.GradientStyle.TopBottom; // 그라디언트스타일 
                cht.BackSecondaryColor = System.Drawing.Color.White;
                cht.BorderlineColor = System.Drawing.Color.FromArgb(((int)(((byte)(26)))), ((int)(((byte)(59)))), ((int)(((byte)(105))))); // 경계색상
                cht.BorderlineDashStyle = System.Windows.Forms.DataVisualization.Charting.ChartDashStyle.Solid;
                cht.BorderlineWidth = 2; // 경계 선 폭
                cht.BorderSkin.SkinStyle = System.Windows.Forms.DataVisualization.Charting.BorderSkinStyle.Emboss; // 경계 스타일.

                //차트 제목 추가.
                System.Windows.Forms.DataVisualization.Charting.Title title1 = new System.Windows.Forms.DataVisualization.Charting.Title();
                title1.Name = "Default"; // 접근 명
                title1.Alignment = System.Drawing.ContentAlignment.MiddleCenter; // 타이틀 정열,
                title1.Font = new System.Drawing.Font("Trebuchet MS", 10.00F, System.Drawing.FontStyle.Bold);
                title1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(26)))), ((int)(((byte)(59)))), ((int)(((byte)(105)))));
                title1.Position.Auto = true;
                title1.ShadowColor = System.Drawing.Color.FromArgb(((int)(((byte)(32)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
                title1.ShadowOffset = 3;// 글자의 새도우 설정.
                title1.Text = strChartTitle; // 타이틀 제목
                cht.Titles.Add(title1); 
                title1 = null;

                //=================================================
                // 차트 영역 추가.
                for (byte b = 0; b < bteChartAreaCount; b++)
                {
                    AddChartArea(ref cht, b);
                }
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
        }

        //======================================
        // 차트 영역 추가.
        //======================================
        private void AddChartArea(ref Chart cht, byte bteAreaNo)
        {
            try
            {
                // 차트에어리어 명은 ca0~ca1 식으로 명명 한다.
                System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea1 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
                chartArea1.Name = MakeAreaName(bteAreaNo); //에어리어 명

                chartArea1.BackColor = cht.BackColor;
                chartArea1.BackGradientStyle = cht.BackGradientStyle;
                chartArea1.BorderColor = cht.BorderlineColor;
                chartArea1.BackSecondaryColor = cht.BackSecondaryColor;
                chartArea1.ShadowColor = System.Drawing.Color.Transparent;

                if (bteAreaNo >= 1)
                {
                    //chartArea1.AlignWithChartArea = "ca0";
                    chartArea1.AlignmentStyle = AreaAlignmentStyles.None;
                }

                cht.ChartAreas.Add(chartArea1); chartArea1 = null;

                cht.ChartAreas[MakeAreaName(bteAreaNo)].AxisX.IsMarginVisible = true; // x 축 선상에서 여유폭을 보이게 한다.
                //cht.ChartAreas["Default"].AxisX.is
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
        }

        //======================================
        // 차트에 범례 추가.
        //======================================
        private string MakeLegendName(byte bteAreaNo)
        {
            return "lg" + bteAreaNo.ToString(); // 범례 접근명.
        }

        public void MakeLegend(ref Chart cht, string strLegnedTitle, string strDockDirection, byte bteAreaNo)
        {
            try
            {
                System.Windows.Forms.DataVisualization.Charting.Legend legend1 = new System.Windows.Forms.DataVisualization.Charting.Legend();
                legend1.BackColor = System.Drawing.Color.Transparent;

                switch (strDockDirection)
                {
                    case "top":
                        legend1.Docking = Docking.Top; //범례를 위에 보인다.
                        legend1.LegendStyle = System.Windows.Forms.DataVisualization.Charting.LegendStyle.Row; //범례를 보일때 열 형식으로
                        break;
                    case "bottom":
                        legend1.Docking = Docking.Bottom; //범례를 위에 보인다.
                        legend1.LegendStyle = System.Windows.Forms.DataVisualization.Charting.LegendStyle.Row; //범례를 보일때 열 형식으로
                        break;
                    case "right":
                        legend1.Docking = Docking.Right; //범례를 위에 보인다.
                        legend1.LegendStyle = System.Windows.Forms.DataVisualization.Charting.LegendStyle.Column; //범례를 보일때 열 형식으로
                        break;
                    case "left":
                        legend1.Docking = Docking.Left; //범례를 위에 보인다.
                        legend1.LegendStyle = System.Windows.Forms.DataVisualization.Charting.LegendStyle.Column; //범례를 보일때 열 형식으로
                        break;
                    default:
                        break;
                }

                legend1.Name = MakeLegendName(bteAreaNo); // 범례 접근명.
                //legend1.Title = strLegnedTitle;// 범례의 타이틀명

                //legend1.TitleAlignment = System.Drawing.StringAlignment.Center; //정열

                //legend1.TitleSeparator = LegendSeparatorStyle.DoubleLine; // 타이틀 라인.
                //legend1.TitleSeparatorColor = System.Drawing.Color.Black;  // 분리선 색상.

                legend1.IsDockedInsideChartArea = false; // 범례를 차트에어리어 안에 넣어준다.
                legend1.DockedToChartArea = MakeAreaName(bteAreaNo); // 차트영역 설정,
                legend1.Enabled = false;

                cht.Legends.Add(legend1);
                legend1 = null;
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
        }

        //======================================
        // 차트에 컬럼 형식 추가.
        //======================================
        public void AddChartColumn(ref Chart cht, byte bteAreaNo, byte bteLegendNo, string strSeriesName
            , ref string[] aryX, ref double[] aryY, bool blnRate, System.Drawing.Color _color, int intBorderDashStyle
            , bool boolEnabled)
        {
            try
            {
                System.Windows.Forms.DataVisualization.Charting.Series series1 = new System.Windows.Forms.DataVisualization.Charting.Series();
                series1.ChartArea = MakeAreaName(bteAreaNo);
                series1.Legend = MakeLegendName(bteLegendNo); // 범례 이름명.
                series1.Name = strSeriesName; // 시리즈 접근명 및 범례에 나오는 이름.
                
                // 차트 유형
                series1.ChartType = SeriesChartType.Line;//.Column;

                // 모양.. 시리즈가 변동될때 마다 색상 변경.
                series1.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(180)))), ((int)(((byte)(26)))), ((int)(((byte)(59)))), ((int)(((byte)(105)))));
                series1.CustomProperties = "DrawingStyle=Cylinder"; // 실린터 형태의 막대 그래프
                series1.ShadowColor = System.Drawing.Color.Transparent;
                //series1.IsValueShownAsLabel = true; //막대 그래프의 값을 표시해준다.
                series1.MarkerSize = 5;    // 막대 그래프의 사이즈
                series1.MarkerStyle = MarkerStyle.Circle;   // 막대 그래프의 꼭대기 원형
                series1.Color = _color;

                if (intBorderDashStyle == 1)
                    series1.BorderDashStyle = ChartDashStyle.DashDotDot;    // 기본값
                else if (intBorderDashStyle == 2)
                        series1.BorderDashStyle = ChartDashStyle.Dot;       // 평균값
                else
                    series1.BorderDashStyle = ChartDashStyle.Solid;         // 표준값

                series1.Points.DataBindXY(aryX, aryY); // 자료 첨부
                series1.IsXValueIndexed = true;

                cht.ChartAreas[0].AxisX.Interval = 1; // X 축 증가 값 설정
                //cht.ChartAreas[0].AxisX.IntervalAutoMode = IntervalAutoMode.FixedCount;
                //cht.ChartAreas[0].AxisX.IntervalOffset = -1.5;
                cht.Series.Add(series1);
                series1 = null;

               // if (blnRate) { AddChartColumnRate(ref  cht, bteAreaNo, bteLegendNo, strSeriesName, ref  aryX, ref  aryY); }
           }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
        }
        
        //======================================
        // 비율 추가.
        //======================================
        public void AddChartColumnRate(ref Chart cht, byte bteAreaNo, byte bteLegendNo, string strSeriesName, ref string[] aryX, ref double[] aryY)
        {
            try
            {
                System.Windows.Forms.DataVisualization.Charting.Series series1 = new System.Windows.Forms.DataVisualization.Charting.Series();
                series1.ChartArea = MakeAreaName(bteAreaNo);
                series1.Legend = MakeLegendName(bteLegendNo); // 범례 이름명.
                series1.Name = strSeriesName; // 시리즈 접근명 및 범례에 나오는 이름.

                // 차트 유형
                series1.ChartType = SeriesChartType.Line;
                series1.CustomProperties = "DrawingStyle=Cylinder"; // 실린터 형태의 막대 그래프
                series1.ShadowColor = System.Drawing.Color.Transparent;
                series1.Label = "#PERCENT{P1}";

                //series1.Points.DataBindXY(aryX, aryY); // 자료 첨부

                //cht.Series.Add(series1);
                //series1 = null;
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
        }

        public void AddChartPie(ref Chart cht, byte bteAreaNo, byte bteLegendNo, string strSeriesName, ref string[] aryX, ref double[] aryY)
        {
            try
            {
                //CreateChartArea(ref cht, 0, 0);

                // 시리즈를 담는다.
                System.Windows.Forms.DataVisualization.Charting.Series series1 = new System.Windows.Forms.DataVisualization.Charting.Series();
                series1.ChartArea = MakeAreaName(bteAreaNo);
                series1.Legend = MakeLegendName(bteLegendNo); // 범례 이름명.
                series1.Name = strSeriesName; // 시리즈 접근명 및 범례에 나오는 이름.

                //막대형.
                series1.ChartType = SeriesChartType.Line;
                series1.Label = "#PERCENT{P1}";

                // 파이형./
                //series1.ChartType = SeriesChartType.Doughnut;
                //series1["PieLabelStyle"] = "Inside";
                //series1["DoughnutRadius"] = "50";

                //series1.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(180)))), ((int)(((byte)(26)))), ((int)(((byte)(59)))), ((int)(((byte)(105)))));
                //series1.ShadowColor = System.Drawing.Color.Transparent;
                //series1.IsValueShownAsLabel = true; //막대 그래프의 값을 표시해준다.

                //series1.Label = "#PERCENT{P0}";

                //series1.Points.DataBindXY(aryX, aryY);
                //cht.Series.Add(series1);
                //series1 = null;

                cht.ChartAreas["ca" + bteAreaNo.ToString()].Area3DStyle.Enable3D = true; //도넛이라 3D 형태로..
           }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
        }
    }// end class
}// end naespace
