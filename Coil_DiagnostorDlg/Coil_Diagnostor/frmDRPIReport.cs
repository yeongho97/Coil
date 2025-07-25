using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Windows.Forms.DataVisualization.Charting;
using Coil_Diagnostor.Function;

// Excel Export 용
using Excel = Microsoft.Office.Interop.Excel;

namespace Coil_Diagnostor
{
    public partial class frmDRPIReport : Form
    {
        protected Function.FunctionMeasureProcess m_MeasureProcess = new Function.FunctionMeasureProcess();
        protected Function.FunctionDataControl m_db = new Function.FunctionDataControl();
        protected Function.FunctionChart m_Chart = new Function.FunctionChart();
        protected frmMessageBox frmMB = new frmMessageBox();
        protected string strPlantName = "";
		protected bool boolFormLoad = false;
		protected string strMeasurementResult = "";
		protected string strTemperature_ReferenceValue = "";
		protected string strFrequency = "";
		protected string strVoltageLevel = "";
        protected string strMeasurementDate = "";
        protected string strReferenceHogi = "초기값";
        protected string strReferenceOHDegree = "제 0 차";

        // 평균치
        protected decimal[] dAveA_ControlRodRdc = new decimal[3];    // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dAveA_ControlRodRac = new decimal[3];    // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dAveA_ControlRodL = new decimal[3];      // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dAveA_ControlRodC = new decimal[3];      // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dAveA_ControlRodQ = new decimal[3];      // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dAveB_ControlRodRdc = new decimal[3];    // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dAveB_ControlRodRac = new decimal[3];    // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dAveB_ControlRodL = new decimal[3];      // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dAveB_ControlRodC = new decimal[3];      // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dAveB_ControlRodQ = new decimal[3];      // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dAveA_StopRdc = new decimal[3];          // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dAveA_StopRac = new decimal[3];          // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dAveA_StopL = new decimal[3];            // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dAveA_StopC = new decimal[3];            // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dAveA_StopQ = new decimal[3];            // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dAveB_StopRdc = new decimal[3];          // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dAveB_StopRac = new decimal[3];          // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dAveB_StopL = new decimal[3];            // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dAveB_StopC = new decimal[3];            // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dAveB_StopQ = new decimal[3];            // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)

        // 기준치
        protected decimal[] dRefA_ControlRodRdc = new decimal[3];    // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dRefA_ControlRodRac = new decimal[3];    // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dRefA_ControlRodL = new decimal[3];      // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dRefA_ControlRodC = new decimal[3];      // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dRefA_ControlRodQ = new decimal[3];      // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dRefB_ControlRodRdc = new decimal[3];    // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dRefB_ControlRodRac = new decimal[3];    // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dRefB_ControlRodL = new decimal[3];      // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dRefB_ControlRodC = new decimal[3];      // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dRefB_ControlRodQ = new decimal[3];      // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dRefA_StopRdc = new decimal[3];          // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dRefA_StopRac = new decimal[3];          // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dRefA_StopL = new decimal[3];            // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dRefA_StopC = new decimal[3];            // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dRefA_StopQ = new decimal[3];            // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dRefB_StopRdc = new decimal[3];          // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dRefB_StopRac = new decimal[3];          // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dRefB_StopL = new decimal[3];            // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dRefB_StopC = new decimal[3];            // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dRefB_StopQ = new decimal[3];            // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)

        System.Windows.Forms.DataVisualization.Charting.Series sSeries = new System.Windows.Forms.DataVisualization.Charting.Series();

        private Panel pnlRod = null;

        public frmDRPIReport()
        {
            CheckForIllegalCrossThreadCalls = false;
            InitializeComponent();
        }

        /// <summary>
        /// 폼 단축키 지정
        /// </summary>
        protected override bool ProcessCmdKey(ref System.Windows.Forms.Message msg, Keys keyData)
        {
            Keys key = keyData & ~(Keys.Shift | Keys.Control);

            switch (key)
            {
                case Keys.F2: // 초기화 버튼
                    btnConfirm.PerformClick();
                    break;
                case Keys.F12: // 닫기 버튼
                    btnClose.PerformClick();
                    break;
            }

            return base.ProcessCmdKey(ref msg, keyData);
        }

        /// <summary>
        /// Form Load
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmDRPIReport_Load(object sender, EventArgs e)
        {
            strPlantName = Gini.GetValue("Device", "PlantName").Trim();

            // ComboBox 데이터 설정
            SetComboBoxDataBinding();

            // 표형식 그리드 초기 설정
            SetDataGridViewInitialize();

            string plantName = Gini.GetValue("Device", "PlantName");
            SetRodPanel(plantName);
            SetCheckBoxDataBinding(pnlRod);

            chkMeasurementTarget_CheckedChanged(chkC, null);
        }

        // 20240402 한인석
        private void SetRodPanel(string plantName)
        {
            switch (plantName)
            {
                case "고리 1발전소":
                    pnlRod = pnlRod33;
                    break;
                case "한빛 1발전소":
                    pnlRod = pnlRod52;
                    break;
                default:
                    pnlRod = pnlRod52;
                    break;
            }
            pnlRod.Visible = true;

            btnCheckBoxAllClear52.PerformClick();
            btnCheckBoxAllClear33.PerformClick();
        }

        private void SetCheckBoxDataBinding(Panel p)
        {
            if (p == null) return;
            foreach (Control grv in p.Controls)
            {
                if (grv is GroupBox)
                {
                    string gbName = $"chk{grv.Name.Substring(2)}";
                    var cb = grv.Controls.Find(gbName, true).FirstOrDefault();
                    if (cb != null)
                    {
                        string key = $"RCSPowerCabinetItem_{cb.Text}";
                        string[] rodNames;
                        if(cb.Text.Replace(" ","") == "정지용") rodNames = Gini.GetValue("DRPI", "DRPIShutdownItem").Split(',');
                        else /*if ("제어용")*/                  rodNames = Gini.GetValue("DRPI", "DRPIControlItem" ).Split(',');
                        foreach (Control rod in grv.Controls)
                        {
                            CheckBox r = rod as CheckBox;
                            if (r.Name != cb.Name)
                            {
                                int i = Convert.ToInt32(r.Name.Substring(cb.Name.Length + 1));
                                if (i >= rodNames.Length || rodNames[i] == "N/A")
                                {
                                    r.Visible = false;
                                }
                                else
                                {
                                    r.Text = rodNames[i];
                                }
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// ComboBox 데이터 설정
        /// </summary>
        private void SetComboBoxDataBinding()
        {
            // 타입
            string[] strDRPITypeItem = Gini.GetValue("Combo", "DRPIReferenceValueLabelType").Split(',');

            cboDRPIType.Items.Clear();

            cboDRPIType.Items.Add("전체");

            for (int i = 0; i < strDRPITypeItem.Length; i++)
            {
                cboDRPIType.Items.Add(strDRPITypeItem[i].Trim());
            }

            cboDRPIType.SelectedIndex = 0;

            // 이력조회타입
            string[] strHistorySearchTypeItem = Gini.GetValue("Combo", "HistorySearchType_Item").Split(',');

            cboHistorySearchType.Items.Clear();

            for (int i = 0; i < strHistorySearchTypeItem.Length; i++)
            {
                cboHistorySearchType.Items.Add(strHistorySearchTypeItem[i].Trim());
            }

            cboHistorySearchType.SelectedIndex = 0;

            // 발전소 호기
            string[] strHogi = Gini.GetValue("Combo", "HogiData").Split(',');

            DataTable dt = new DataTable();
            dt = m_db.GetDRPIReferenceValueOhDegreeData(strPlantName.Trim());

            cboHogi.Items.Clear();

            // 기본 호기 설정
            for (int i = 0; i < strHogi.Length; i++)
            {
                cboHogi.Items.Add(strHogi[i].Trim());
            }

            if (dt != null && dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (dt.Rows[i]["Hogi"].ToString().Trim() == "1 호기"
                        || dt.Rows[i]["Hogi"].ToString().Trim() == "2 호기")
                        continue;

                    cboHogi.Items.Add(dt.Rows[i]["Hogi"].ToString().Trim());
                }
            }

            // 호기 선택
            if (Gini.GetValue("DRPI", "SelectDRPI_Hogi").Trim() == "")
            {
                // 최초 실행 시 기본 호기 선택
                cboHogi.SelectedIndex = 0;
            }
            else
            {
                // 직전 실행 시 기본 호기 선택
                cboHogi.SelectedItem = Gini.GetValue("DRPI", "SelectDRPI_Hogi").Trim();
            }

            // 주파수
            string[] strFrequencyItem = Gini.GetValue("Combo", "FrequencyItem").Split(',');

            cboFrequency.Items.Clear();

            cboFrequency.Items.Add("전체");

            for (int i = 0; i < strFrequencyItem.Length; i++)
            {
                cboFrequency.Items.Add(strFrequencyItem[i].Trim());
            }

            cboFrequency.SelectedIndex = 0;

            // 측정타입
            string[] strMeasurementTypeItem = Gini.GetValue("Combo", "RCSType_Item").Split(',');

            cboMeasurementType.Items.Clear();

            for (int i = 0; i < strMeasurementTypeItem.Length; i++)
            {
                cboMeasurementType.Items.Add(strMeasurementTypeItem[i].Trim());
            }

            cboMeasurementType.SelectedIndex = 1;

            // 비교대상 
            string[] strComparisonTargetItem = Gini.GetValue("Combo", "ComparisonTarget_Item").Split(',');

            cboComparisonTarget.Items.Clear();

            for (int i = 0; i < strComparisonTargetItem.Length; i++)
            {
                cboComparisonTarget.Items.Add(strComparisonTargetItem[i].Trim());
            }

            cboComparisonTarget.SelectedIndex = 0;

            // 그룹
            string[] strMeasurementGroup_Item = Gini.GetValue("Combo", "DRPIGroup_Item").Split(',');

            cboGroup.Items.Clear();

            cboGroup.Items.Add("전체");

            for (int i = 0; i < strMeasurementGroup_Item.Length; i++)
            {
                cboGroup.Items.Add(strMeasurementGroup_Item[i].Trim());
            }

            cboGroup.SelectedIndex = 0;
        }

        /// <summary>
        /// 표형식 그리드 초기 설정
        /// </summary>
        private void SetDataGridViewInitialize()
        {
            try
            {
                //----- DataGridView Column 추가
                DataGridViewTextBoxColumn Column1 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn Column2 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn Column3 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn Column4 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn Column5 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn Column6 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn Column7 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn Column8 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn Column9 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn Column10 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn Column11 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn Column12 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn Column13 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn Column14 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn Column15 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn Column16 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn Column17 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn Column18 = new DataGridViewTextBoxColumn();

                Column1.HeaderText = "제어봉";
                Column1.Name = "ControlName";
                Column1.Width = 60;
                Column1.ReadOnly = true;
                Column1.Visible = true;
                Column1.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column1.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column2.HeaderText = "그룹";
                Column2.Name = "DRPIGroup";
                Column2.Width = 60;
                Column2.ReadOnly = true;
                Column2.Visible = true;
                Column2.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column2.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column3.HeaderText = "코일명";
                Column3.Name = "CoilName";
                Column3.Width = 80;
                Column3.ReadOnly = true;
                Column3.Visible = true;
                Column3.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column3.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column4.HeaderText = "결과";
                Column4.Name = "Result";
                Column4.Width = 80;
                Column4.ReadOnly = true;
                Column4.Visible = true;
                Column4.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column4.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column5.HeaderText = "Index";
                Column5.Name = "CoilNumber";
                Column5.Width = 80;
                Column5.ReadOnly = true;
                Column5.Visible = false;
                Column5.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column5.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column6.HeaderText = "DC 저항(Ω)";
                Column6.Name = "DC_ResistanceValue";
                Column6.Width = 100;
                Column6.ReadOnly = true;
                Column6.Visible = chkRdc.Checked;
                Column6.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column6.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column7.HeaderText = "편차(Ω)";
                Column7.Name = "DC_Deviation";
                Column7.Width = 80;
                Column7.ReadOnly = true;
                Column7.Visible = chkRdc.Checked;
                Column7.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column7.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column8.HeaderText = "AC 저항(Ω)";
                Column8.Name = "AC_ResistanceValue";
                Column8.Width = 100;
                Column8.ReadOnly = true;
                Column8.Visible = chkRac.Checked;
                Column8.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column8.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column9.HeaderText = "편차(Ω)";
                Column9.Name = "AC_Deviation";
                Column9.Width = 80;
                Column9.ReadOnly = true;
                Column9.Visible = chkRac.Checked;
                Column9.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column9.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column10.HeaderText = "인덕턴스(mH)";
                Column10.Name = "L_InductanceValue";
                Column10.Width = 100;
                Column10.ReadOnly = true;
                Column10.Visible = chkL.Checked;
                Column10.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column10.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column11.HeaderText = "편차(mH)";
                Column11.Name = "L_Deviation";
                Column11.Width = 80;
                Column11.ReadOnly = true;
                Column11.Visible = chkL.Checked;
                Column11.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column11.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column12.HeaderText = "캐패시턴스(uF)";
                Column12.Name = "C_CapacitanceValue";
                Column12.Width = 100;
                Column12.ReadOnly = true;
                Column12.Visible = chkC.Checked;
                Column12.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column12.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column13.HeaderText = "편차(uF)";
                Column13.Name = "C_Deviation";
                Column13.Width = 80;
                Column13.ReadOnly = true;
                Column13.Visible = chkC.Checked;
                Column13.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column13.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column14.HeaderText = "Q Factor";
                Column14.Name = "Q_FactorValue";
                Column14.Width = 100;
                Column14.ReadOnly = true;
                Column14.Visible = chkQ.Checked;
                Column14.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column14.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column15.HeaderText = "편차";
                Column15.Name = "Q_Deviation";
                Column15.Width = 80;
                Column15.ReadOnly = true;
                Column15.Visible = chkQ.Checked;
                Column15.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column15.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column16.HeaderText = "온도(℃)";
                Column16.Name = "Temperature_ReferenceValue";
                Column16.Width = 80;
                Column16.ReadOnly = true;
                Column16.Visible = true;
                Column16.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column16.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column17.HeaderText = "주파수";
                Column17.Name = "Frequency";
                Column17.Width = 80;
                Column17.ReadOnly = true;
                Column17.Visible = true;
                Column17.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column17.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column18.HeaderText = "전압레벨(mV)";
                Column18.Name = "VoltageLevel";
                Column18.Width = 80;
                Column18.ReadOnly = true;
                Column18.Visible = true;
                Column18.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column18.SortMode = DataGridViewColumnSortMode.NotSortable;

                dgvMeasurement.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] { Column1, Column2, Column3, Column4, Column5, Column6
                    , Column7, Column8, Column9, Column10, Column11, Column12, Column13, Column14, Column15, Column16, Column17, Column18 });

                dgvMeasurement.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing;
                dgvMeasurement.ColumnHeadersHeight = 40;
                dgvMeasurement.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dgvMeasurement.RowHeadersWidth = 40;

                // DataGridView 기타 세팅
                dgvMeasurement.AllowUserToAddRows = false; // Row 추가 기능 미사용
                dgvMeasurement.AllowUserToDeleteRows = false; // Row 삭제 기능 미사용
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
        }

        private void cboType_SelectedIndexChanged(object sender, EventArgs e)
        {
            // 그리드와 그래프 초기화
            SetDataGridViewAndGraphInitialize();

            string strDRPIControlItem = Gini.GetValue("DRPI", "Control_Item").Trim();
            string strDRPIStopItem = Gini.GetValue("DRPI", "Stop_Item").Trim();
            string[] arrayDRPIControlAndStop;

            if (cboDRPIType.SelectedItem.ToString().Trim() == "전체" || cboDRPIType.SelectedItem.ToString().Trim() == "제어용")
            {
                arrayDRPIControlAndStop = strDRPIControlItem.Split(',');
            }
            else
            {
                arrayDRPIControlAndStop = strDRPIStopItem.Split(',');
            }

            // 코일명
            cboCoilName.Items.Clear();

            cboCoilName.Items.Add("전체");

            for (int i = 0; i < arrayDRPIControlAndStop.Length; i++)
            {
                if (arrayDRPIControlAndStop[i].Trim() != "N/A")
                    cboCoilName.Items.Add(arrayDRPIControlAndStop[i].Trim());
            }

            cboCoilName.SelectedIndex = 0;

            if (cboDRPIType.SelectedItem.ToString().Trim() == "제어용")
            {
                gbControlRod52.Enabled = true;
                chkStopRod52.Checked = false;
                gbStopRod52.Enabled = false;
            }
            else if (cboDRPIType.SelectedItem.ToString().Trim() == "정지용")
            {
                chkControlRod52.Checked = false;
                gbControlRod52.Enabled = false;
                gbStopRod52.Enabled = true;
            }
            else
            {
                gbControlRod52.Enabled = true;
                gbStopRod52.Enabled = true;
            }
        }

        /// <summary>
        /// 측정대상(호기) Selected Index Changed Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cboHogi_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                // 그리드와 그래프 초기화
                SetDataGridViewAndGraphInitialize();

                string strSelectHogi = cboHogi.SelectedItem == null ? "" : cboHogi.SelectedItem.ToString().Trim();

                // 호기별 OH 차수 설정
                DataTable dt = new DataTable();
                dt = m_db.GetDRPIDiagnosisHeaderOhDegreeInfo(strPlantName, strSelectHogi);

                cboOhDegree.Items.Clear();

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (dt.Rows[i]["Oh_Degree"] != null)
                    {
                        cboOhDegree.Items.Add(dt.Rows[i]["Oh_Degree"].ToString().Trim());
                    }
                }

                if (Gini.GetValue("DRPI", "SelectDRPI_OHDegree").Trim() == "")
                    cboOhDegree.SelectedIndex = dt.Rows.Count - 1;
                else
                    cboOhDegree.SelectedItem = "제 " + Gini.GetValue("DRPI", "SelectDRPI_OHDegree").Trim() + " 차";

                if (cboOhDegree.SelectedItem == null) cboOhDegree.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
        }

        /// <summary>
        /// 차수 Selected Index Changed Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cboOhDegree_SelectedIndexChanged(object sender, EventArgs e)
        {
            // 그리드와 그래프 초기화
            SetDataGridViewAndGraphInitialize();
        }

        /// <summary>
        /// 이력조회타입 Selected Index Changed Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cboHistorySearchType_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                // 그리드와 그래프 초기화
                SetDataGridViewAndGraphInitialize();

                if (cboHistorySearchType.SelectedItem.ToString().Trim() == "기간별")
                {
                    // 코일명 ComboBox 바인딩
                    SetCoilNameComboBoxBinding();
                }
                else
                {
                    // 코일명 ComboBox 바인딩
                    SetCoilNameComboBoxBinding();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
        }

        /// <summary>
        /// 코일명 ComboBox 바인딩
        /// </summary>
        private void SetCoilNameComboBoxBinding()
        {
            try
            {
                string strDRPIControlItem = Gini.GetValue("DRPI", "Control_Item").Trim();
                string strDRPIStopItem = Gini.GetValue("DRPI", "Stop_Item").Trim();
                string[] arrayDRPIControlAndStop;

                if (cboHistorySearchType.SelectedItem.ToString().Trim() == "기간별")
                {
                    if (cboDRPIType.SelectedItem.ToString().Trim() == "전체" || cboDRPIType.SelectedItem.ToString().Trim() == "제어용")
                    {
                        arrayDRPIControlAndStop = strDRPIControlItem.Split(',');
                    }
                    else
                    {
                        arrayDRPIControlAndStop = strDRPIStopItem.Split(',');
                    }

                    // 코일명
                    cboCoilName.Items.Clear();

                    for (int i = 0; i < arrayDRPIControlAndStop.Length; i++)
                    {
                        if (arrayDRPIControlAndStop[i].Trim() != "N/A")
                            cboCoilName.Items.Add(arrayDRPIControlAndStop[i].Trim());
                    }

                    cboCoilName.SelectedIndex = 0;
                }
                else
                {
                    if (cboDRPIType.SelectedItem.ToString().Trim() == "전체" || cboDRPIType.SelectedItem.ToString().Trim() == "제어용")
                    {
                        arrayDRPIControlAndStop = strDRPIControlItem.Split(',');
                    }
                    else
                    {
                        arrayDRPIControlAndStop = strDRPIStopItem.Split(',');
                    }

                    // 코일명
                    cboCoilName.Items.Clear();

                    cboCoilName.Items.Add("전체");

                    for (int i = 0; i < arrayDRPIControlAndStop.Length; i++)
                    {
                        if (arrayDRPIControlAndStop[i].Trim() != "N/A")
                            cboCoilName.Items.Add(arrayDRPIControlAndStop[i].Trim());
                    }

                    cboCoilName.SelectedIndex = 0;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
        }

        /// <summary>
        /// 그룹 Selected Index Changed Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cboGroup_SelectedIndexChanged(object sender, EventArgs e)
        {
            // 그리드와 그래프 초기화
            SetDataGridViewAndGraphInitialize();
        }

        /// <summary>
        /// 측정타입 Selected Index Changed Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cboMeasurementType_SelectedIndexChanged(object sender, EventArgs e)
        {
            // 그리드와 그래프 초기화
            SetDataGridViewAndGraphInitialize();

            // 타입 ComboBox Index 0번으로 설정
            cboDRPIType.SelectedIndex = 0;
        }

        /// <summary>
        /// 주파수 Selected Index Changed Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cboFrequency_SelectedIndexChanged(object sender, EventArgs e)
        {
            // 그리드와 그래프 초기화
            SetDataGridViewAndGraphInitialize();
        }

        /// <summary>
        /// 코일명 Selected Index Changed Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cboCoilName_SelectedIndexChanged(object sender, EventArgs e)
        {
            // 그리드와 그래프 초기화
            SetDataGridViewAndGraphInitialize();
        }

        /// <summary>
        /// 비교대상 Selected Index Changed Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cboComparisonTarget_SelectedIndexChanged(object sender, EventArgs e)
        {
            // 그리드와 그래프 초기화
            SetDataGridViewAndGraphInitialize();
        }

        /// <summary>
        /// 전체선택 Button Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCheckBoxAllSelect_Click(object sender, EventArgs e)
        {
            if (gbStopRod52.Enabled)
                chkStopRod52.Checked = true;

            if (gbControlRod52.Enabled)
                chkControlRod52.Checked = true;
        }

        /// <summary>
        /// 전체해제 Button Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCheckBoxAllClear_Click(object sender, EventArgs e)
        {
            if (gbStopRod52.Enabled)
                chkStopRod52.Checked = false;
            
            if (gbControlRod52.Enabled)
                chkControlRod52.Checked = false;
        }

        /// <summary>
        /// 닫기 Button Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// 폴더 바로가기
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnFolderMove_Click(object sender, EventArgs e)
        {
            string path = Application.StartupPath + @"\Report";
            System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(path);
            
            if (!di.Exists)
            {
                di.Create();
            }
            
            System.Diagnostics.Process.Start(path);
        }

        /// <summary>
        /// 확인 Button Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnConfirm_Click(object sender, EventArgs e)
		{
			System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;

            try
			{
				if (!chkRdc.Checked && !chkRac.Checked && !chkL.Checked && !chkC.Checked && !chkQ.Checked)
				{
					frmMB.lblMessage.Text = "Item 선택을 하십시오";
					frmMB.TopMost = true;
					frmMB.ShowDialog();

					System.Windows.Forms.Cursor.Current = Cursors.Default;
					return;
				}

                btnConfirm.Enabled = false;

                // 그리드와 그래프 초기화
                SetDataGridViewAndGraphInitialize();

                string strHogi = cboHogi.SelectedItem == null ? "" : cboHogi.SelectedItem.ToString().Trim();
                string strOhDegree = cboOhDegree.SelectedItem == null ? "" : cboOhDegree.SelectedItem.ToString().Trim();
                string strDRPIGroup = cboGroup.SelectedItem == null ? "" : cboGroup.SelectedItem.ToString().Trim();
                string strMeasurementType = cboMeasurementType.SelectedItem == null ? "" : cboMeasurementType.SelectedItem.ToString().Trim();
                string strDRPIType = cboDRPIType.SelectedItem == null ? "" : cboDRPIType.SelectedItem.ToString().Trim();
                string strCoilName = cboCoilName.SelectedItem == null ? "" : cboCoilName.SelectedItem.ToString().Trim();
                string strFrequency = cboFrequency.SelectedItem == null ? "" : cboFrequency.SelectedItem.ToString().Trim();
                string strComparisonTarget = cboComparisonTarget.SelectedItem == null ? "" : cboComparisonTarget.SelectedItem.ToString().Trim();
                string strStop = "", strControlRod = "", strControlName = "";

                if (strDRPIGroup.Trim() != "전체" && strCoilName.Trim() != "" && strCoilName.Trim() != "전체")
                {
                    strCoilName = string.Format("'{0}{1}'", strDRPIGroup.Trim(), Regex.Replace(strCoilName, @"\D", ""));
                }
                else if (strDRPIGroup.Trim() == "전체" && strCoilName.Trim() != "" && strCoilName.Trim() != "전체")
                {
                    strCoilName = string.Format("'A{0}', 'B{0}'", Regex.Replace(strCoilName, @"\D", ""));
                }

                strControlName = GetCheckBoxCheckedTrueItem(pnlRod);

                if (strControlName.Length == 0)
                {
                    frmMB.lblMessage.Text = "제어봉을 선택을 하십시오";
                    frmMB.TopMost = true;
                    frmMB.ShowDialog();

                    System.Windows.Forms.Cursor.Current = Cursors.Default;
                    return;
                }

                // 기준 값(DRPI) 가져오기
                GetDRPIReferenceValueSelected();

                DataTable dtDB = new DataTable();
                DataTable dtAverage = new DataTable();

                // 측정 타입이 평균치일 경우
                if (cboMeasurementType.SelectedItem != null && cboMeasurementType.SelectedItem.ToString().Trim() == "평균치")
                {
                    dtDB = m_db.GetDRPIDiagnosisDetailDataGridViewDataAverage(strPlantName.Trim(), strHogi, strOhDegree, strDRPIGroup
                            , strDRPIType, strControlName, strCoilName, strFrequency, strReferenceHogi, strReferenceOHDegree);

                    dtAverage = dtDB;
                }
                else
                {
                    dtDB = m_db.GetDRPIDiagnosisDetailDataGridViewDataMeasure(strPlantName.Trim(), strHogi, strOhDegree, strDRPIGroup
                            , strDRPIType, strControlName, strCoilName, strFrequency, strReferenceHogi, strReferenceOHDegree);

                    dtAverage = m_db.GetDRPIDiagnosisDetailDataGridViewDataAverage(strPlantName.Trim(), strHogi, strOhDegree, strDRPIGroup
                            , strDRPIType, strControlName, strCoilName, strFrequency, strReferenceHogi, strReferenceOHDegree);
                }

                //dtDB = m_db.GetDRPIReportDataInfo(strPlantName.Trim(), strHogi, strOhDegree, strDRPIGroup, strDRPIType
                //        , strFrequency, strCoilName, strControlName);

                if (dtDB == null || dtDB.Rows.Count <= 0)
                {
                    frmMB.lblMessage.Text = "보고서 데이터가 없습니다.";
                    frmMB.TopMost = true;
                    frmMB.ShowDialog();
                }
                else
                {
                    // 작성일자(최종 측정일)
                    DateTime dtMeasurementDate = new DateTime();
                    DateTime dtSelectMeasurementDate = new DateTime();

                    for (int i = 0; i < dtDB.Rows.Count; i++)
                    {
                        dtMeasurementDate = dtDB.Rows[i]["MeasurementDate"] == null || dtDB.Rows[i]["MeasurementDate"].ToString().Trim() == ""
                            ? DateTime.Now : Convert.ToDateTime(dtDB.Rows[i]["MeasurementDate"].ToString().Trim());

                        if (dtSelectMeasurementDate < dtMeasurementDate)
                            dtSelectMeasurementDate = dtMeasurementDate;
                    }

                    strMeasurementDate = dtSelectMeasurementDate.ToString("yyyy-MM-dd HH:mm:ss");
                    if (strMeasurementDate.Trim() == "0001-01-01 12:00:00")
                        strMeasurementDate = "";

                    // Data 가져오기
                    GetDRPIMeasurementDataBinding(dtDB, dtAverage, strComparisonTarget, strHogi, strOhDegree, strDRPIGroup
                        , strDRPIType, strControlName, strCoilName, strFrequency);

					if (dgvMeasurement.RowCount > 0)
					{
						// 엑셀 보고서 보내기
						if (ExcelReportExport(strHogi, strOhDegree, strDRPIType, strCoilName, strComparisonTarget))
						{
							frmMB.lblMessage.Text = "보고서 생성이 완료되었습니다.";
							frmMB.TopMost = true;
							frmMB.ShowDialog();
						}
						else
						{
							frmMB.lblMessage.Text = "보고서 생성이 실패하였습니다.";
							frmMB.TopMost = true;
							frmMB.ShowDialog();
						}
					}
					else
					{
						frmMB.lblMessage.Text = "보고서 생성 데이터가 없습니다.";
						frmMB.TopMost = true;
						frmMB.ShowDialog();
					}
                }
            }
            catch (Exception ex)
            {
                frmMB.lblMessage.Text = "보고서 생성이 실패";
                frmMB.TopMost = true;
                frmMB.ShowDialog();
                System.Diagnostics.Debug.Print(ex.Message);
            }
            finally
            {
                btnConfirm.Enabled = true;
            }

			System.Windows.Forms.Cursor.Current = Cursors.Default;
        }

        /// <summary>
        /// 그리드와 그래프 초기화
        /// </summary>
        private void SetDataGridViewAndGraphInitialize()
        {
            // 그리드 초기화
            dgvMeasurement.Rows.Clear();

            // 그래프 초기화
            chartMeasurementValueDC.Series.Clear();
            chartMeasurementValueAC.Series.Clear();
            chartMeasurementValueL.Series.Clear();
            chartMeasurementValueC.Series.Clear();
            chartMeasurementValueQ.Series.Clear();
            chartAverageValueDC.Series.Clear();
            chartAverageValueAC.Series.Clear();
            chartAverageValueL.Series.Clear();
            chartAverageValueC.Series.Clear();
            chartAverageValueQ.Series.Clear();
        }

        /// <summary>
        /// Data 가져오기
        /// </summary>
        private void GetDRPIMeasurementDataBinding(DataTable _dtDB, DataTable _dtAverage, string _strComparisonTarget
            , string strHogi, string strOhDegree, string strDRPIGroup, string strDRPIType, string strControlRod
            , string strCoilName, string strFrequency)
        {
            try
            {
                if (_dtDB != null && _dtDB.Rows.Count > 0)
                {
                    // 코일 그룹별 평균치 산출
                    SetControlRodCoilGroupAverageCalculation(_dtAverage, strHogi, strOhDegree, strDRPIGroup
                        , strDRPIType, strControlRod, strCoilName, strFrequency);

                    // 표 형식
                    // 측정 타입이 평균치일 경우
                    if (cboMeasurementType.SelectedItem != null && cboMeasurementType.SelectedItem.ToString().Trim() == "평균치")
                    {
                        // 측정타입이 측정치일 경우
                        if (_dtAverage != null && _dtAverage.Rows.Count > 0)
                            SetDRPIReferenceValueTableMark(ref _dtAverage, _strComparisonTarget);
                    }
                    else
                    {
                        // 측정타입이 측정치일 경우
                        if (_dtDB != null && _dtDB.Rows.Count > 0)
                            SetDRPIReferenceValueTableMark(ref _dtDB, _strComparisonTarget);

                        // 평균 구하기
                        SetDataAverage(ref _dtAverage, _dtDB);
                    }

                    if (_dtAverage != null && _dtAverage.Rows.Count > 0)
                    {
                        string strXLabelValue = "", strXLabelCheck = "", strDataCheck = "";

                        for (int i = 0; i < _dtAverage.Rows.Count; i++)
                        {
                            strDataCheck = _dtAverage.Rows[i]["ControlName"].ToString().Trim() + " " + _dtAverage.Rows[i]["DRPIGroup"].ToString().Trim();

                            if (strXLabelCheck.Trim() != strDataCheck.Trim())
                            {
                                strXLabelCheck = _dtAverage.Rows[i]["ControlName"].ToString().Trim() + " " + _dtAverage.Rows[i]["DRPIGroup"].ToString().Trim();

                                if (strXLabelValue.Trim() == "")
                                    strXLabelValue = _dtAverage.Rows[i]["ControlName"].ToString().Trim() + " " + _dtAverage.Rows[i]["DRPIGroup"].ToString().Trim();
                                else
                                    strXLabelValue = strXLabelValue + "," + _dtAverage.Rows[i]["ControlName"].ToString().Trim() + " " + _dtAverage.Rows[i]["DRPIGroup"].ToString().Trim();
                            }
                        }

                        string[] arrayXLabelValue = strXLabelValue.Split(',');
                        int intXLabelValueCount = arrayXLabelValue.Length;

                        // 측정값 및 편차 그래프 그리기
                        SetDRPIReferenceValueMeasurement(_dtAverage, arrayXLabelValue, intXLabelValueCount);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
        }

        /// <summary>
        /// 평균 구하기
        /// </summary>
        /// <param name="dtAverage"></param>
        /// <param name="dtDB"></param>
        private void SetDataAverage(ref DataTable dtAverage, DataTable dtDB)
        {
            decimal dcmSumDCDeviation = 0.000M, dcmSumACDeviation = 0.000M, dcmSumLDeviation = 0.000M, dcmSumCDeviation = 0.000M
                , dcmSumQDeviation = 0.000M;
            string strDB_PlantNeme = "", strDB_Hogi = "", strDB_Oh_Degree = "", strDB_DRPIGroup = "", strDB_DRPIType = ""
                , strDB_ControlName = "", strDB_CoilName = "";
            int intCoilCount = 0;

            for (int i = 0; i < dtAverage.Rows.Count; i++)
            {
                if (strDB_PlantNeme.Trim() != dtAverage.Rows[i]["PlantName"].ToString().Trim()
                    || strDB_Hogi.Trim() != dtAverage.Rows[i]["Hogi"].ToString().Trim()
                    || strDB_Oh_Degree.Trim() != dtAverage.Rows[i]["Oh_Degree"].ToString().Trim()
                    || strDB_DRPIGroup.Trim() != dtAverage.Rows[i]["DRPIGroup"].ToString().Trim()
                    || strDB_DRPIType.Trim() != dtAverage.Rows[i]["DRPIType"].ToString().Trim()
                    || strDB_ControlName.Trim() != dtAverage.Rows[i]["ControlName"].ToString().Trim()
                    || strDB_CoilName.Trim() != dtAverage.Rows[i]["CoilName"].ToString().Trim())
                {
                    strDB_PlantNeme = dtAverage.Rows[i]["PlantName"].ToString().Trim();
                    strDB_Hogi = dtAverage.Rows[i]["Hogi"].ToString().Trim();
                    strDB_Oh_Degree = dtAverage.Rows[i]["Oh_Degree"].ToString().Trim();
                    strDB_DRPIGroup = dtAverage.Rows[i]["DRPIGroup"].ToString().Trim();
                    strDB_DRPIType = dtAverage.Rows[i]["DRPIType"].ToString().Trim();
                    strDB_ControlName = dtAverage.Rows[i]["ControlName"].ToString().Trim();
                    strDB_CoilName = dtAverage.Rows[i]["CoilName"].ToString().Trim();
                    intCoilCount = 0;

                    dcmSumDCDeviation = 0.000M;
                    dcmSumACDeviation = 0.000M;
                    dcmSumLDeviation = 0.000M;
                    dcmSumCDeviation = 0.000M;
                    dcmSumQDeviation = 0.000M;

                    for (int j = 0; j < dtDB.Rows.Count; j++)
                    {
                        if (strDB_PlantNeme.Trim() == dtDB.Rows[j]["PlantName"].ToString().Trim()
                            && strDB_Hogi.Trim() == dtDB.Rows[j]["Hogi"].ToString().Trim()
                            && strDB_Oh_Degree.Trim() == dtDB.Rows[j]["Oh_Degree"].ToString().Trim()
                            && strDB_DRPIGroup.Trim() == dtDB.Rows[j]["DRPIGroup"].ToString().Trim()
                            && strDB_DRPIType.Trim() == dtDB.Rows[j]["DRPIType"].ToString().Trim()
                            && strDB_ControlName.Trim() == dtDB.Rows[j]["ControlName"].ToString().Trim()
                            && strDB_CoilName.Trim() == dtDB.Rows[j]["CoilName"].ToString().Trim())
                        {
                            dcmSumDCDeviation += dtDB.Rows[j]["DC_Deviation"] == null || dtDB.Rows[j]["DC_Deviation"].ToString().Trim() == ""
                                ? 0.000M : Convert.ToDecimal(dtDB.Rows[j]["DC_Deviation"].ToString().Trim());
                            dcmSumACDeviation += dtDB.Rows[j]["AC_Deviation"] == null || dtDB.Rows[j]["AC_Deviation"].ToString().Trim() == ""
                                ? 0.000M : Convert.ToDecimal(dtDB.Rows[j]["AC_Deviation"].ToString().Trim());
                            dcmSumLDeviation += dtDB.Rows[j]["L_Deviation"] == null || dtDB.Rows[j]["L_Deviation"].ToString().Trim() == ""
                                ? 0.000M : Convert.ToDecimal(dtDB.Rows[j]["L_Deviation"].ToString().Trim());
                            dcmSumCDeviation += dtDB.Rows[j]["C_Deviation"] == null || dtDB.Rows[j]["C_Deviation"].ToString().Trim() == ""
                                ? 0.000000M : Convert.ToDecimal(dtDB.Rows[j]["C_Deviation"].ToString().Trim());
                            dcmSumQDeviation += dtDB.Rows[j]["Q_Deviation"] == null || dtDB.Rows[j]["Q_Deviation"].ToString().Trim() == ""
                                ? 0.000M : Convert.ToDecimal(dtDB.Rows[j]["Q_Deviation"].ToString().Trim());

                            intCoilCount++;
                        }
                    }

                    dtAverage.Rows[i]["DC_Deviation"] = (dcmSumDCDeviation / intCoilCount).ToString().Trim();
                    dtAverage.Rows[i]["AC_Deviation"] = (dcmSumACDeviation / intCoilCount).ToString().Trim();
                    dtAverage.Rows[i]["L_Deviation"] = (dcmSumLDeviation / intCoilCount).ToString().Trim();
                    dtAverage.Rows[i]["C_Deviation"] = (dcmSumCDeviation / intCoilCount).ToString().Trim();
                    dtAverage.Rows[i]["Q_Deviation"] = (dcmSumQDeviation / intCoilCount).ToString().Trim();
                }
            }
        }

        /// <summary>
        /// 기준 값(DRPI) 가져오기
        /// </summary>
        private void GetDRPIReferenceValueSelected()
        {
            DataTable dt;
            string strSelectHogi = "";
            string strSelectOhDegree = "";

            try
            {
                // DRPI 기준값 Data 가져오기
                strSelectHogi = cboHogi.SelectedItem == null ? "초기값" : cboHogi.SelectedItem.ToString().Trim();
                strSelectOhDegree = cboOhDegree.SelectedItem == null ? "제 0 차" : cboOhDegree.SelectedItem.ToString().Trim();

                dt = new DataTable();

                if ((m_db.GetDRPIReferenceValueDataCount(strPlantName.Trim(), strSelectHogi.Trim(), strSelectOhDegree.Trim())) > 0)
                {
                    // 해당 차수의 기준 값을 가져온다.
                    strReferenceHogi = strSelectHogi.Trim();
                    strReferenceOHDegree = strSelectOhDegree.Trim();
                    dt = m_db.GetDRPIReferenceValueDataInfo(strPlantName, strReferenceHogi.Trim(), strReferenceOHDegree.Trim());
                }
                else
                {
                    // MAX OH 차수를 가져온다.
                    int intMaxOhDegree = m_db.GetDRPIReferenceValueDataMaxOhDegree(strPlantName.Trim(), strSelectHogi.Trim());

                    if (intMaxOhDegree > 0)
                    {
                        // MAX 차수의 기준 값을 가져온다.
                        strReferenceHogi = "초기값";
                        strReferenceOHDegree = "제 " + intMaxOhDegree.ToString().Trim() + " 차";
                        dt = m_db.GetDRPIReferenceValueDataInfo(strPlantName, strReferenceHogi, strReferenceOHDegree.Trim());
                    }
                    else
                    {
                        // 초기값 0 차수의 기준 값을 가져온다.
                        strReferenceHogi = "초기값";
                        strReferenceOHDegree = "제 0 차";
                        dt = m_db.GetDRPIReferenceValueDataInfo(strPlantName.Trim(), strReferenceHogi.Trim(), strReferenceOHDegree.Trim());
                    }
                }

                if (dt != null && dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        if (dt.Rows[i]["CoilCabinetType"].ToString().Trim() == "제어용" && dt.Rows[i]["CoilCabinetName"].ToString().Trim() == "A1")
                        {
                            dRefA_ControlRodRdc[0] = dt.Rows[i]["Rdc_DRPIReferenceValue"] == null || dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim());
                            dRefA_ControlRodRac[0] = dt.Rows[i]["Rac_DRPIReferenceValue"] == null || dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim());
                            dRefA_ControlRodL[0] = dt.Rows[i]["L_DRPIReferenceValue"] == null || dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim());
                            dRefA_ControlRodC[0] = dt.Rows[i]["C_DRPIReferenceValue"] == null || dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000000M : Convert.ToDecimal(dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim());
                            dRefA_ControlRodQ[0] = dt.Rows[i]["Q_DRPIReferenceValue"] == null || dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim());
                        }
                        else if (dt.Rows[i]["CoilCabinetType"].ToString().Trim() == "제어용" && dt.Rows[i]["CoilCabinetName"].ToString().Trim() == "A2")
                        {
                            dRefA_ControlRodRdc[1] = dt.Rows[i]["Rdc_DRPIReferenceValue"] == null || dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim());
                            dRefA_ControlRodRac[1] = dt.Rows[i]["Rac_DRPIReferenceValue"] == null || dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim());
                            dRefA_ControlRodL[1] = dt.Rows[i]["L_DRPIReferenceValue"] == null || dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim());
                            dRefA_ControlRodC[1] = dt.Rows[i]["C_DRPIReferenceValue"] == null || dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000000M : Convert.ToDecimal(dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim());
                            dRefA_ControlRodQ[1] = dt.Rows[i]["Q_DRPIReferenceValue"] == null || dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim());
                        }
                        else if (dt.Rows[i]["CoilCabinetType"].ToString().Trim() == "제어용" && dt.Rows[i]["CoilCabinetName"].ToString().Substring(0, 1) == "A")
                        {
                            dRefA_ControlRodRdc[2] = dt.Rows[i]["Rdc_DRPIReferenceValue"] == null || dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim());
                            dRefA_ControlRodRac[2] = dt.Rows[i]["Rac_DRPIReferenceValue"] == null || dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim());
                            dRefA_ControlRodL[2] = dt.Rows[i]["L_DRPIReferenceValue"] == null || dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim());
                            dRefA_ControlRodC[2] = dt.Rows[i]["C_DRPIReferenceValue"] == null || dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000000M : Convert.ToDecimal(dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim());
                            dRefA_ControlRodQ[2] = dt.Rows[i]["Q_DRPIReferenceValue"] == null || dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim());
                        }
                        else if (dt.Rows[i]["CoilCabinetType"].ToString().Trim() == "제어용" && dt.Rows[i]["CoilCabinetName"].ToString().Trim() == "B1")
                        {
                            dRefB_ControlRodRdc[0] = dt.Rows[i]["Rdc_DRPIReferenceValue"] == null || dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim());
                            dRefB_ControlRodRac[0] = dt.Rows[i]["Rac_DRPIReferenceValue"] == null || dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim());
                            dRefB_ControlRodL[0] = dt.Rows[i]["L_DRPIReferenceValue"] == null || dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim());
                            dRefB_ControlRodC[0] = dt.Rows[i]["C_DRPIReferenceValue"] == null || dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000000M : Convert.ToDecimal(dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim());
                            dRefB_ControlRodQ[0] = dt.Rows[i]["Q_DRPIReferenceValue"] == null || dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim());
                        }
                        else if (dt.Rows[i]["CoilCabinetType"].ToString().Trim() == "제어용" && dt.Rows[i]["CoilCabinetName"].ToString().Trim() == "B2")
                        {
                            dRefB_ControlRodRdc[1] = dt.Rows[i]["Rdc_DRPIReferenceValue"] == null || dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim());
                            dRefB_ControlRodRac[1] = dt.Rows[i]["Rac_DRPIReferenceValue"] == null || dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim());
                            dRefB_ControlRodL[1] = dt.Rows[i]["L_DRPIReferenceValue"] == null || dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim());
                            dRefB_ControlRodC[1] = dt.Rows[i]["C_DRPIReferenceValue"] == null || dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000000M : Convert.ToDecimal(dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim());
                            dRefB_ControlRodQ[1] = dt.Rows[i]["Q_DRPIReferenceValue"] == null || dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim());
                        }
                        else if (dt.Rows[i]["CoilCabinetType"].ToString().Trim() == "제어용" && dt.Rows[i]["CoilCabinetName"].ToString().Substring(0, 1) == "B")
                        {
                            dRefB_ControlRodRdc[2] = dt.Rows[i]["Rdc_DRPIReferenceValue"] == null || dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim());
                            dRefB_ControlRodRac[2] = dt.Rows[i]["Rac_DRPIReferenceValue"] == null || dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim());
                            dRefB_ControlRodL[2] = dt.Rows[i]["L_DRPIReferenceValue"] == null || dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim());
                            dRefB_ControlRodC[2] = dt.Rows[i]["C_DRPIReferenceValue"] == null || dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000000M : Convert.ToDecimal(dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim());
                            dRefB_ControlRodQ[2] = dt.Rows[i]["Q_DRPIReferenceValue"] == null || dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim());
                        }
                        else if (dt.Rows[i]["CoilCabinetType"].ToString().Trim() == "정지용" & dt.Rows[i]["CoilCabinetName"].ToString().Trim() == "A1")
                        {
                            dRefA_StopRdc[0] = dt.Rows[i]["Rdc_DRPIReferenceValue"] == null || dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim());
                            dRefA_StopRac[0] = dt.Rows[i]["Rac_DRPIReferenceValue"] == null || dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim());
                            dRefA_StopL[0] = dt.Rows[i]["L_DRPIReferenceValue"] == null || dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim());
                            dRefA_StopC[0] = dt.Rows[i]["C_DRPIReferenceValue"] == null || dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000000M : Convert.ToDecimal(dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim());
                            dRefA_StopQ[0] = dt.Rows[i]["Q_DRPIReferenceValue"] == null || dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim());
                        }
                        else if (dt.Rows[i]["CoilCabinetType"].ToString().Trim() == "정지용" && dt.Rows[i]["CoilCabinetName"].ToString().Trim() == "A2")
                        {
                            dRefA_StopRdc[1] = dt.Rows[i]["Rdc_DRPIReferenceValue"] == null || dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim());
                            dRefA_StopRac[1] = dt.Rows[i]["Rac_DRPIReferenceValue"] == null || dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim());
                            dRefA_StopL[1] = dt.Rows[i]["L_DRPIReferenceValue"] == null || dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim());
                            dRefA_StopC[1] = dt.Rows[i]["C_DRPIReferenceValue"] == null || dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000000M : Convert.ToDecimal(dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim());
                            dRefA_StopQ[1] = dt.Rows[i]["Q_DRPIReferenceValue"] == null || dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim());
                        }
                        else if (dt.Rows[i]["CoilCabinetType"].ToString().Trim() == "정지용" && dt.Rows[i]["CoilCabinetName"].ToString().Substring(0, 1) == "A")
                        {
                            dRefA_StopRdc[2] = dt.Rows[i]["Rdc_DRPIReferenceValue"] == null || dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim());
                            dRefA_StopRac[2] = dt.Rows[i]["Rac_DRPIReferenceValue"] == null || dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim());
                            dRefA_StopL[2] = dt.Rows[i]["L_DRPIReferenceValue"] == null || dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim());
                            dRefA_StopC[2] = dt.Rows[i]["C_DRPIReferenceValue"] == null || dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000000M : Convert.ToDecimal(dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim());
                            dRefA_StopQ[2] = dt.Rows[i]["Q_DRPIReferenceValue"] == null || dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim());
                        }
                        else if (dt.Rows[i]["CoilCabinetType"].ToString().Trim() == "정지용" && dt.Rows[i]["CoilCabinetName"].ToString().Trim() == "B1")
                        {
                            dRefB_StopRdc[0] = dt.Rows[i]["Rdc_DRPIReferenceValue"] == null || dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim());
                            dRefB_StopRac[0] = dt.Rows[i]["Rac_DRPIReferenceValue"] == null || dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim());
                            dRefB_StopL[0] = dt.Rows[i]["L_DRPIReferenceValue"] == null || dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim());
                            dRefB_StopC[0] = dt.Rows[i]["C_DRPIReferenceValue"] == null || dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000000M : Convert.ToDecimal(dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim());
                            dRefB_StopQ[0] = dt.Rows[i]["Q_DRPIReferenceValue"] == null || dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim());
                        }
                        else if (dt.Rows[i]["CoilCabinetType"].ToString().Trim() == "정지용" && dt.Rows[i]["CoilCabinetName"].ToString().Trim() == "B2")
                        {
                            dRefB_StopRdc[1] = dt.Rows[i]["Rdc_DRPIReferenceValue"] == null || dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim());
                            dRefB_StopRac[1] = dt.Rows[i]["Rac_DRPIReferenceValue"] == null || dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim());
                            dRefB_StopL[1] = dt.Rows[i]["L_DRPIReferenceValue"] == null || dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim());
                            dRefB_StopC[1] = dt.Rows[i]["C_DRPIReferenceValue"] == null || dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000000M : Convert.ToDecimal(dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim());
                            dRefB_StopQ[1] = dt.Rows[i]["Q_DRPIReferenceValue"] == null || dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim());
                        }
                        else if (dt.Rows[i]["CoilCabinetType"].ToString().Trim() == "정지용" && dt.Rows[i]["CoilCabinetName"].ToString().Substring(0, 1) == "B")
                        {
                            dRefB_StopRdc[2] = dt.Rows[i]["Rdc_DRPIReferenceValue"] == null || dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim());
                            dRefB_StopRac[2] = dt.Rows[i]["Rac_DRPIReferenceValue"] == null || dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim());
                            dRefB_StopL[2] = dt.Rows[i]["L_DRPIReferenceValue"] == null || dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim());
                            dRefB_StopC[2] = dt.Rows[i]["C_DRPIReferenceValue"] == null || dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000000M : Convert.ToDecimal(dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim());
                            dRefB_StopQ[2] = dt.Rows[i]["Q_DRPIReferenceValue"] == null || dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim());
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
        }

        /// <summary>
        /// 제어용 / 정지용 코일 그룹별 평균치 산출
        /// </summary>
        /// <param name="_dt"></param>
		private void SetControlRodCoilGroupAverageCalculation(DataTable _dtAverage, string strHogi, string strOhDegree, string strDRPIGroup
                        , string strDRPIType, string strControlRod, string strCoilName, string strFrequency)
        {
            try
            {
                decimal[] dcmControlARdcSum = new decimal[3];
                decimal[] dcmControlARacSum =  new decimal[3];
                decimal[] dcmControlALSum =  new decimal[3];
                decimal[] dcmControlACSum =  new decimal[3];
                decimal[] dcmControlAQSum =  new decimal[3];
                decimal[] dcmStopARdcSum =  new decimal[3];
                decimal[] dcmStopARacSum =  new decimal[3];
                decimal[] dcmStopALSum =  new decimal[3];
                decimal[] dcmStopACSum =  new decimal[3];
                decimal[] dcmStopAQSum =  new decimal[3];
                decimal[] dcmControlBRdcSum =  new decimal[3];
                decimal[] dcmControlBRacSum =  new decimal[3];
                decimal[] dcmControlBLSum =  new decimal[3];
                decimal[] dcmControlBCSum =  new decimal[3];
                decimal[] dcmControlBQSum =  new decimal[3];
                decimal[] dcmStopBRdcSum =  new decimal[3];
                decimal[] dcmStopBRacSum =  new decimal[3];
                decimal[] dcmStopBLSum =  new decimal[3];
                decimal[] dcmStopBCSum = new decimal[3];
                decimal[] dcmStopBQSum = new decimal[3];
                int[] intControlARdcCount = new int[3];
                int[] intControlARacCount = new int[3];
                int[] intControlALCount = new int[3];
                int[] intControlACCount = new int[3];
                int[] intControlAQCount = new int[3];
                int[] intStopARdcCount = new int[3];
                int[] intStopARacCount = new int[3];
                int[] intStopALCount = new int[3];
                int[] intStopACCount = new int[3];
                int[] intStopAQCount = new int[3];
                int[] intControlBRdcCount = new int[3];
                int[] intControlBRacCount = new int[3];
                int[] intControlBLCount = new int[3];
                int[] intControlBCCount = new int[3];
                int[] intControlBQCount = new int[3];
                int[] intStopBRdcCount = new int[3];
                int[] intStopBRacCount = new int[3];
                int[] intStopBLCount = new int[3];
                int[] intStopBCCount = new int[3];
                int[] intStopBQCount = new int[3];
                
                // 그룹 A, B의 각각 제어용, 정지용의 평균가져오기
                DataTable dtAvg = new DataTable();
                dtAvg = m_db.GetDRPIDiagnosisDetailDataGridViewDataAverage(strPlantName.Trim(), strHogi, strOhDegree, strDRPIGroup
                        , strDRPIType, strControlRod, strCoilName, strFrequency);

                decimal dcmGroupA1_ControlDC = 0M, dcmGroupA1_ControlAC = 0M, dcmGroupA1_ControlL = 0M, dcmGroupA1_ControlC = 0M, dcmGroupA1_ControlQ = 0M
                       , dcmGroupA2_ControlDC = 0M, dcmGroupA2_ControlAC = 0M, dcmGroupA2_ControlL = 0M, dcmGroupA2_ControlC = 0M, dcmGroupA2_ControlQ = 0M
                       , dcmGroupA_ControlDC = 0M, dcmGroupA_ControlAC = 0M, dcmGroupA_ControlL = 0M, dcmGroupA_ControlC = 0M, dcmGroupA_ControlQ = 0M
                       , dcmGroupB1_ControlDC = 0M, dcmGroupB1_ControlAC = 0M, dcmGroupB1_ControlL = 0M, dcmGroupB1_ControlC = 0M, dcmGroupB1_ControlQ = 0M
                       , dcmGroupB2_ControlDC = 0M, dcmGroupB2_ControlAC = 0M, dcmGroupB2_ControlL = 0M, dcmGroupB2_ControlC = 0M, dcmGroupB2_ControlQ = 0M
                       , dcmGroupB_ControlDC = 0M, dcmGroupB_ControlAC = 0M, dcmGroupB_ControlL = 0M, dcmGroupB_ControlC = 0M, dcmGroupB_ControlQ = 0M
                       , dcmGroupA1_StopDC = 0M, dcmGroupA1_StopAC = 0M, dcmGroupA1_StopL = 0M, dcmGroupA1_StopC = 0M, dcmGroupA1_StopQ = 0M
                       , dcmGroupA2_StopDC = 0M, dcmGroupA2_StopAC = 0M, dcmGroupA2_StopL = 0M, dcmGroupA2_StopC = 0M, dcmGroupA2_StopQ = 0M
                       , dcmGroupA_StopDC = 0M, dcmGroupA_StopAC = 0M, dcmGroupA_StopL = 0M, dcmGroupA_StopC = 0M, dcmGroupA_StopQ = 0M
                       , dcmGroupB1_StopDC = 0M, dcmGroupB1_StopAC = 0M, dcmGroupB1_StopL = 0M, dcmGroupB1_StopC = 0M, dcmGroupB1_StopQ = 0M
                       , dcmGroupB2_StopDC = 0M, dcmGroupB2_StopAC = 0M, dcmGroupB2_StopL = 0M, dcmGroupB2_StopC = 0M, dcmGroupB2_StopQ = 0M
                       , dcmGroupB_StopDC = 0M, dcmGroupB_StopAC = 0M, dcmGroupB_StopL = 0M, dcmGroupB_StopC = 0M, dcmGroupB_StopQ = 0M;

                // 그룹 A, B의 각각 제어용, 정지용의 가져온 평균값을 변수에 설정
                GetGeoupItemAverage(dtAvg, ref dcmGroupA1_ControlDC, ref dcmGroupA1_ControlAC, ref dcmGroupA1_ControlL, ref dcmGroupA1_ControlC, ref dcmGroupA1_ControlQ
                                        , ref dcmGroupA2_ControlDC, ref dcmGroupA2_ControlAC, ref dcmGroupA2_ControlL, ref dcmGroupA2_ControlC, ref dcmGroupA2_ControlQ
                                        , ref dcmGroupA_ControlDC, ref dcmGroupA_ControlAC, ref dcmGroupA_ControlL, ref dcmGroupA_ControlC, ref dcmGroupA_ControlQ
                                        , ref dcmGroupB1_ControlDC, ref dcmGroupB1_ControlAC, ref dcmGroupB1_ControlL, ref dcmGroupB1_ControlC, ref dcmGroupB1_ControlQ
                                        , ref dcmGroupB2_ControlDC, ref dcmGroupB2_ControlAC, ref dcmGroupB2_ControlL, ref dcmGroupB2_ControlC, ref dcmGroupB2_ControlQ
                                        , ref dcmGroupB_ControlDC, ref dcmGroupB_ControlAC, ref dcmGroupB_ControlL, ref dcmGroupB_ControlC, ref dcmGroupB_ControlQ
                                        , ref dcmGroupA1_StopDC, ref dcmGroupA1_StopAC, ref dcmGroupA1_StopL, ref dcmGroupA1_StopC, ref dcmGroupA1_StopQ
                                        , ref dcmGroupA2_StopDC, ref dcmGroupA2_StopAC, ref dcmGroupA2_StopL, ref dcmGroupA2_StopC, ref dcmGroupA2_StopQ
                                        , ref dcmGroupA_StopDC, ref dcmGroupA_StopAC, ref dcmGroupA_StopL, ref dcmGroupA_StopC, ref dcmGroupA_StopQ
                                        , ref dcmGroupB1_StopDC, ref dcmGroupB1_StopAC, ref dcmGroupB1_StopL, ref dcmGroupB1_StopC, ref dcmGroupB1_StopQ
                                        , ref dcmGroupB2_StopDC, ref dcmGroupB2_StopAC, ref dcmGroupB2_StopL, ref dcmGroupB2_StopC, ref dcmGroupB2_StopQ
                                        , ref dcmGroupB_StopDC, ref dcmGroupB_StopAC, ref dcmGroupB_StopL, ref dcmGroupB_StopC, ref dcmGroupB_StopQ);

				// 코일 그룹A 제어용 평균치 산출
				string filterExpression = string.Format("DRPIGroup = '{0}' AND DRPIType = '{1}'", "A", "제어용");
				string sort = "DRPIGroup, DRPIType, SeqNumber";
				if (_dtAverage.Select(filterExpression, sort) != null && _dtAverage.Select(filterExpression, sort).Length > 0)
					SetControlRodCoilGroupA_AverageCalculation(_dtAverage.Select(filterExpression, sort).CopyToDataTable()
						, dcmGroupA1_ControlDC, dcmGroupA1_ControlAC, dcmGroupA1_ControlL, dcmGroupA1_ControlC, dcmGroupA1_ControlQ
						, dcmGroupA2_ControlDC, dcmGroupA2_ControlAC, dcmGroupA2_ControlL, dcmGroupA2_ControlC, dcmGroupA2_ControlQ
						, dcmGroupA_ControlDC, dcmGroupA_ControlAC, dcmGroupA_ControlL, dcmGroupA_ControlC, dcmGroupA_ControlQ);
				// 코일 그룹B 제어용 평균치 산출
				filterExpression = string.Format("DRPIGroup = '{0}' AND DRPIType = '{1}'", "B", "제어용");
				if (_dtAverage.Select(filterExpression, sort) != null && _dtAverage.Select(filterExpression, sort).Length > 0)
					SetControlRodCoilGroupB_AverageCalculation(_dtAverage.Select(filterExpression, sort).CopyToDataTable()
						, dcmGroupB1_ControlDC, dcmGroupB1_ControlAC, dcmGroupB1_ControlL, dcmGroupB1_ControlC, dcmGroupB1_ControlQ
						, dcmGroupB2_ControlDC, dcmGroupB2_ControlAC, dcmGroupB2_ControlL, dcmGroupB2_ControlC, dcmGroupB2_ControlQ
						, dcmGroupB_ControlDC, dcmGroupB_ControlAC, dcmGroupB_ControlL, dcmGroupB_ControlC, dcmGroupB_ControlQ);
				// 코일 그룹A 정지용 평균치 산출
				filterExpression = string.Format("DRPIGroup = '{0}' AND DRPIType = '{1}'", "A", "정지용");
				if (_dtAverage.Select(filterExpression, sort) != null && _dtAverage.Select(filterExpression, sort).Length > 0)
					SetStopCoilGroupA_AverageCalculation(_dtAverage.Select(filterExpression, sort).CopyToDataTable()
						, dcmGroupA1_StopDC, dcmGroupA1_StopAC, dcmGroupA1_StopL, dcmGroupA1_StopC, dcmGroupA1_StopQ
						, dcmGroupA2_StopDC, dcmGroupA2_StopAC, dcmGroupA2_StopL, dcmGroupA2_StopC, dcmGroupA2_StopQ
						, dcmGroupA_StopDC, dcmGroupA_StopAC, dcmGroupA_StopL, dcmGroupA_StopC, dcmGroupA_StopQ);
				// 코일 그룹B 정지용 평균치 산출
				filterExpression = string.Format("DRPIGroup = '{0}' AND DRPIType = '{1}'", "B", "정지용");
				if (_dtAverage.Select(filterExpression, sort) != null && _dtAverage.Select(filterExpression, sort).Length > 0)
					SetStopCoilGroupB_AverageCalculation(_dtAverage.Select(filterExpression, sort).CopyToDataTable()
						, dcmGroupB1_StopDC, dcmGroupB1_StopAC, dcmGroupB1_StopL, dcmGroupB1_StopC, dcmGroupB1_StopQ
						, dcmGroupB2_StopDC, dcmGroupB2_StopAC, dcmGroupB2_StopL, dcmGroupB2_StopC, dcmGroupB2_StopQ
						, dcmGroupB_StopDC, dcmGroupB_StopAC, dcmGroupB_StopL, dcmGroupB_StopC, dcmGroupB_StopQ);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
        }

        /// <summary>
        /// 그룹 A, B의 각각 제어용, 정지용의 가져온 평균값을 변수에 설정
        /// </summary>
        /// <param name="dtAvg"></param>
        /// <param name="dcmGroupA1_ControlDC"></param>
        /// <param name="dcmGroupA1_ControlAC"></param>
        /// <param name="dcmGroupA1_ControlL"></param>
        /// <param name="dcmGroupA1_ControlC"></param>
        /// <param name="dcmGroupA1_ControlQ"></param>
        /// <param name="dcmGroupA2_ControlDC"></param>
        /// <param name="dcmGroupA2_ControlAC"></param>
        /// <param name="dcmGroupA2_ControlL"></param>
        /// <param name="dcmGroupA2_ControlC"></param>
        /// <param name="dcmGroupA2_ControlQ"></param>
        /// <param name="dcmGroupA_ControlDC"></param>
        /// <param name="dcmGroupA_ControlAC"></param>
        /// <param name="dcmGroupA_ControlL"></param>
        /// <param name="dcmGroupA_ControlC"></param>
        /// <param name="dcmGroupA_ControlQ"></param>
        /// <param name="dcmGroupB1_ControlDC"></param>
        /// <param name="dcmGroupB1_ControlAC"></param>
        /// <param name="dcmGroupB1_ControlL"></param>
        /// <param name="dcmGroupB1_ControlC"></param>
        /// <param name="dcmGroupB1_ControlQ"></param>
        /// <param name="dcmGroupB2_ControlDC"></param>
        /// <param name="dcmGroupB2_ControlAC"></param>
        /// <param name="dcmGroupB2_ControlL"></param>
        /// <param name="dcmGroupB2_ControlC"></param>
        /// <param name="dcmGroupB2_ControlQ"></param>
        /// <param name="dcmGroupB_ControlDC"></param>
        /// <param name="dcmGroupB_ControlAC"></param>
        /// <param name="dcmGroupB_ControlL"></param>
        /// <param name="dcmGroupB_ControlC"></param>
        /// <param name="dcmGroupB_ControlQ"></param>
        /// <param name="dcmGroupA1_StopDC"></param>
        /// <param name="dcmGroupA1_StopAC"></param>
        /// <param name="dcmGroupA1_StopL"></param>
        /// <param name="dcmGroupA1_StopC"></param>
        /// <param name="dcmGroupA1_StopQ"></param>
        /// <param name="dcmGroupA2_StopDC"></param>
        /// <param name="dcmGroupA2_StopAC"></param>
        /// <param name="dcmGroupA2_StopL"></param>
        /// <param name="dcmGroupA2_StopC"></param>
        /// <param name="dcmGroupA2_StopQ"></param>
        /// <param name="dcmGroupA_StopDC"></param>
        /// <param name="dcmGroupA_StopAC"></param>
        /// <param name="dcmGroupA_StopL"></param>
        /// <param name="dcmGroupA_StopC"></param>
        /// <param name="dcmGroupA_StopQ"></param>
        /// <param name="dcmGroupB1_StopDC"></param>
        /// <param name="dcmGroupB1_StopAC"></param>
        /// <param name="dcmGroupB1_StopL"></param>
        /// <param name="dcmGroupB1_StopC"></param>
        /// <param name="dcmGroupB1_StopQ"></param>
        /// <param name="dcmGroupB2_StopDC"></param>
        /// <param name="dcmGroupB2_StopAC"></param>
        /// <param name="dcmGroupB2_StopL"></param>
        /// <param name="dcmGroupB2_StopC"></param>
        /// <param name="dcmGroupB2_StopQ"></param>
        /// <param name="dcmGroupB_StopDC"></param>
        /// <param name="dcmGroupB_StopAC"></param>
        /// <param name="dcmGroupB_StopL"></param>
        /// <param name="dcmGroupB_StopC"></param>
        /// <param name="dcmGroupB_StopQ"></param>
        private void GetGeoupItemAverage(DataTable dtAvg
            , ref decimal dcmGroupA1_ControlDC, ref decimal dcmGroupA1_ControlAC, ref decimal dcmGroupA1_ControlL, ref decimal dcmGroupA1_ControlC, ref decimal dcmGroupA1_ControlQ
            , ref decimal dcmGroupA2_ControlDC, ref decimal dcmGroupA2_ControlAC, ref decimal dcmGroupA2_ControlL, ref decimal dcmGroupA2_ControlC, ref decimal dcmGroupA2_ControlQ
            , ref decimal dcmGroupA_ControlDC, ref decimal dcmGroupA_ControlAC, ref decimal dcmGroupA_ControlL, ref decimal dcmGroupA_ControlC, ref decimal dcmGroupA_ControlQ
            , ref decimal dcmGroupB1_ControlDC, ref decimal dcmGroupB1_ControlAC, ref decimal dcmGroupB1_ControlL, ref decimal dcmGroupB1_ControlC, ref decimal dcmGroupB1_ControlQ
            , ref decimal dcmGroupB2_ControlDC, ref decimal dcmGroupB2_ControlAC, ref decimal dcmGroupB2_ControlL, ref decimal dcmGroupB2_ControlC, ref decimal dcmGroupB2_ControlQ
            , ref decimal dcmGroupB_ControlDC, ref decimal dcmGroupB_ControlAC, ref decimal dcmGroupB_ControlL, ref decimal dcmGroupB_ControlC, ref decimal dcmGroupB_ControlQ
            , ref decimal dcmGroupA1_StopDC, ref decimal dcmGroupA1_StopAC, ref decimal dcmGroupA1_StopL, ref decimal dcmGroupA1_StopC, ref decimal dcmGroupA1_StopQ
            , ref decimal dcmGroupA2_StopDC, ref decimal dcmGroupA2_StopAC, ref decimal dcmGroupA2_StopL, ref decimal dcmGroupA2_StopC, ref decimal dcmGroupA2_StopQ
            , ref decimal dcmGroupA_StopDC, ref decimal dcmGroupA_StopAC, ref decimal dcmGroupA_StopL, ref decimal dcmGroupA_StopC, ref decimal dcmGroupA_StopQ
            , ref decimal dcmGroupB1_StopDC, ref decimal dcmGroupB1_StopAC, ref decimal dcmGroupB1_StopL, ref decimal dcmGroupB1_StopC, ref decimal dcmGroupB1_StopQ
            , ref decimal dcmGroupB2_StopDC, ref decimal dcmGroupB2_StopAC, ref decimal dcmGroupB2_StopL, ref decimal dcmGroupB2_StopC, ref decimal dcmGroupB2_StopQ
            , ref decimal dcmGroupB_StopDC, ref decimal dcmGroupB_StopAC, ref decimal dcmGroupB_StopL, ref decimal dcmGroupB_StopC, ref decimal dcmGroupB_StopQ)
        {
            try
            {
                for (int i = 0; i < dtAvg.Rows.Count; i++)
                {
                    if (dtAvg.Rows[i]["DRPIGroup"] == null) continue;
                    if (dtAvg.Rows[i]["DRPIType"] == null) continue;

                    if (dtAvg.Rows[i]["DRPIGroup"].ToString().Trim() == "A" && dtAvg.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                    {
                        if (dtAvg.Rows[i]["SeqNumber"].ToString().Trim() == "1")
                        {
                            dcmGroupA1_ControlDC = dtAvg.Rows[i]["DC_ResistanceValue"] == null || dtAvg.Rows[i]["DC_ResistanceValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["DC_ResistanceValue"].ToString().Trim());
                            dcmGroupA1_ControlAC = dtAvg.Rows[i]["AC_ResistanceValue"] == null || dtAvg.Rows[i]["AC_ResistanceValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["AC_ResistanceValue"].ToString().Trim());
                            dcmGroupA1_ControlL = dtAvg.Rows[i]["L_InductanceValue"] == null || dtAvg.Rows[i]["L_InductanceValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["L_InductanceValue"].ToString().Trim());
                            dcmGroupA1_ControlC = dtAvg.Rows[i]["C_CapacitanceValue"] == null || dtAvg.Rows[i]["C_CapacitanceValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["C_CapacitanceValue"].ToString().Trim());
                            dcmGroupA1_ControlQ = dtAvg.Rows[i]["Q_FactorValue"] == null || dtAvg.Rows[i]["Q_FactorValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["Q_FactorValue"].ToString().Trim());
                        }
                        else if (dtAvg.Rows[i]["SeqNumber"].ToString().Trim() == "2")
                        {
                            dcmGroupA2_ControlDC = dtAvg.Rows[i]["DC_ResistanceValue"] == null || dtAvg.Rows[i]["DC_ResistanceValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["DC_ResistanceValue"].ToString().Trim());
                            dcmGroupA2_ControlAC = dtAvg.Rows[i]["AC_ResistanceValue"] == null || dtAvg.Rows[i]["AC_ResistanceValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["AC_ResistanceValue"].ToString().Trim());
                            dcmGroupA2_ControlL = dtAvg.Rows[i]["L_InductanceValue"] == null || dtAvg.Rows[i]["L_InductanceValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["L_InductanceValue"].ToString().Trim());
                            dcmGroupA2_ControlC = dtAvg.Rows[i]["C_CapacitanceValue"] == null || dtAvg.Rows[i]["C_CapacitanceValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["C_CapacitanceValue"].ToString().Trim());
                            dcmGroupA2_ControlQ = dtAvg.Rows[i]["Q_FactorValue"] == null || dtAvg.Rows[i]["Q_FactorValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["Q_FactorValue"].ToString().Trim());
                        }
                        else
                        {
                            dcmGroupA_ControlDC = dtAvg.Rows[i]["DC_ResistanceValue"] == null || dtAvg.Rows[i]["DC_ResistanceValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["DC_ResistanceValue"].ToString().Trim());
                            dcmGroupA_ControlAC = dtAvg.Rows[i]["AC_ResistanceValue"] == null || dtAvg.Rows[i]["AC_ResistanceValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["AC_ResistanceValue"].ToString().Trim());
                            dcmGroupA_ControlL = dtAvg.Rows[i]["L_InductanceValue"] == null || dtAvg.Rows[i]["L_InductanceValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["L_InductanceValue"].ToString().Trim());
                            dcmGroupA_ControlC = dtAvg.Rows[i]["C_CapacitanceValue"] == null || dtAvg.Rows[i]["C_CapacitanceValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["C_CapacitanceValue"].ToString().Trim());
                            dcmGroupA_ControlQ = dtAvg.Rows[i]["Q_FactorValue"] == null || dtAvg.Rows[i]["Q_FactorValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["Q_FactorValue"].ToString().Trim());
                        }
                    }
                    else if (dtAvg.Rows[i]["DRPIGroup"].ToString().Trim() == "A" && dtAvg.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                    {
                        if (dtAvg.Rows[i]["SeqNumber"].ToString().Trim() == "1")
                        {
                            dcmGroupA1_StopDC = dtAvg.Rows[i]["DC_ResistanceValue"] == null || dtAvg.Rows[i]["DC_ResistanceValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["DC_ResistanceValue"].ToString().Trim());
                            dcmGroupA1_StopAC = dtAvg.Rows[i]["AC_ResistanceValue"] == null || dtAvg.Rows[i]["AC_ResistanceValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["AC_ResistanceValue"].ToString().Trim());
                            dcmGroupA1_StopL = dtAvg.Rows[i]["L_InductanceValue"] == null || dtAvg.Rows[i]["L_InductanceValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["L_InductanceValue"].ToString().Trim());
                            dcmGroupA1_StopC = dtAvg.Rows[i]["C_CapacitanceValue"] == null || dtAvg.Rows[i]["C_CapacitanceValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["C_CapacitanceValue"].ToString().Trim());
                            dcmGroupA1_StopQ = dtAvg.Rows[i]["Q_FactorValue"] == null || dtAvg.Rows[i]["Q_FactorValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["Q_FactorValue"].ToString().Trim());
                        }
                        else if (dtAvg.Rows[i]["SeqNumber"].ToString().Trim() == "2")
                        {
                            dcmGroupA2_StopDC = dtAvg.Rows[i]["DC_ResistanceValue"] == null || dtAvg.Rows[i]["DC_ResistanceValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["DC_ResistanceValue"].ToString().Trim());
                            dcmGroupA2_StopAC = dtAvg.Rows[i]["AC_ResistanceValue"] == null || dtAvg.Rows[i]["AC_ResistanceValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["AC_ResistanceValue"].ToString().Trim());
                            dcmGroupA2_StopL = dtAvg.Rows[i]["L_InductanceValue"] == null || dtAvg.Rows[i]["L_InductanceValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["L_InductanceValue"].ToString().Trim());
                            dcmGroupA2_StopC = dtAvg.Rows[i]["C_CapacitanceValue"] == null || dtAvg.Rows[i]["C_CapacitanceValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["C_CapacitanceValue"].ToString().Trim());
                            dcmGroupA2_StopQ = dtAvg.Rows[i]["Q_FactorValue"] == null || dtAvg.Rows[i]["Q_FactorValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["Q_FactorValue"].ToString().Trim());
                        }
                        else
                        {
                            dcmGroupA_StopDC = dtAvg.Rows[i]["DC_ResistanceValue"] == null || dtAvg.Rows[i]["DC_ResistanceValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["DC_ResistanceValue"].ToString().Trim());
                            dcmGroupA_StopAC = dtAvg.Rows[i]["AC_ResistanceValue"] == null || dtAvg.Rows[i]["AC_ResistanceValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["AC_ResistanceValue"].ToString().Trim());
                            dcmGroupA_StopL = dtAvg.Rows[i]["L_InductanceValue"] == null || dtAvg.Rows[i]["L_InductanceValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["L_InductanceValue"].ToString().Trim());
                            dcmGroupA_StopC = dtAvg.Rows[i]["C_CapacitanceValue"] == null || dtAvg.Rows[i]["C_CapacitanceValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["C_CapacitanceValue"].ToString().Trim());
                            dcmGroupA_StopQ = dtAvg.Rows[i]["Q_FactorValue"] == null || dtAvg.Rows[i]["Q_FactorValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["Q_FactorValue"].ToString().Trim());
                        }
                    }
                    else if (dtAvg.Rows[i]["DRPIGroup"].ToString().Trim() == "B" && dtAvg.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                    {
                        if (dtAvg.Rows[i]["SeqNumber"].ToString().Trim() == "1")
                        {
                            dcmGroupB1_ControlDC = dtAvg.Rows[i]["DC_ResistanceValue"] == null || dtAvg.Rows[i]["DC_ResistanceValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["DC_ResistanceValue"].ToString().Trim());
                            dcmGroupB1_ControlAC = dtAvg.Rows[i]["AC_ResistanceValue"] == null || dtAvg.Rows[i]["AC_ResistanceValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["AC_ResistanceValue"].ToString().Trim());
                            dcmGroupB1_ControlL = dtAvg.Rows[i]["L_InductanceValue"] == null || dtAvg.Rows[i]["L_InductanceValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["L_InductanceValue"].ToString().Trim());
                            dcmGroupB1_ControlC = dtAvg.Rows[i]["C_CapacitanceValue"] == null || dtAvg.Rows[i]["C_CapacitanceValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["C_CapacitanceValue"].ToString().Trim());
                            dcmGroupB1_ControlQ = dtAvg.Rows[i]["Q_FactorValue"] == null || dtAvg.Rows[i]["Q_FactorValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["Q_FactorValue"].ToString().Trim());
                        }
                        else if (dtAvg.Rows[i]["SeqNumber"].ToString().Trim() == "2")
                        {
                            dcmGroupB2_ControlDC = dtAvg.Rows[i]["DC_ResistanceValue"] == null || dtAvg.Rows[i]["DC_ResistanceValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["DC_ResistanceValue"].ToString().Trim());
                            dcmGroupB2_ControlAC = dtAvg.Rows[i]["AC_ResistanceValue"] == null || dtAvg.Rows[i]["AC_ResistanceValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["AC_ResistanceValue"].ToString().Trim());
                            dcmGroupB2_ControlL = dtAvg.Rows[i]["L_InductanceValue"] == null || dtAvg.Rows[i]["L_InductanceValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["L_InductanceValue"].ToString().Trim());
                            dcmGroupB2_ControlC = dtAvg.Rows[i]["C_CapacitanceValue"] == null || dtAvg.Rows[i]["C_CapacitanceValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["C_CapacitanceValue"].ToString().Trim());
                            dcmGroupB2_ControlQ = dtAvg.Rows[i]["Q_FactorValue"] == null || dtAvg.Rows[i]["Q_FactorValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["Q_FactorValue"].ToString().Trim());
                        }
                        else
                        {
                            dcmGroupB_ControlDC = dtAvg.Rows[i]["DC_ResistanceValue"] == null || dtAvg.Rows[i]["DC_ResistanceValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["DC_ResistanceValue"].ToString().Trim());
                            dcmGroupB_ControlAC = dtAvg.Rows[i]["AC_ResistanceValue"] == null || dtAvg.Rows[i]["AC_ResistanceValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["AC_ResistanceValue"].ToString().Trim());
                            dcmGroupB_ControlL = dtAvg.Rows[i]["L_InductanceValue"] == null || dtAvg.Rows[i]["L_InductanceValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["L_InductanceValue"].ToString().Trim());
                            dcmGroupB_ControlC = dtAvg.Rows[i]["C_CapacitanceValue"] == null || dtAvg.Rows[i]["C_CapacitanceValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["C_CapacitanceValue"].ToString().Trim());
                            dcmGroupB_ControlQ = dtAvg.Rows[i]["Q_FactorValue"] == null || dtAvg.Rows[i]["Q_FactorValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["Q_FactorValue"].ToString().Trim());
                        }
                    }
                    else if (dtAvg.Rows[i]["DRPIGroup"].ToString().Trim() == "B" && dtAvg.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                    {
                        if (dtAvg.Rows[i]["SeqNumber"].ToString().Trim() == "1")
                        {
                            dcmGroupB1_StopDC = dtAvg.Rows[i]["DC_ResistanceValue"] == null || dtAvg.Rows[i]["DC_ResistanceValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["DC_ResistanceValue"].ToString().Trim());
                            dcmGroupB1_StopAC = dtAvg.Rows[i]["AC_ResistanceValue"] == null || dtAvg.Rows[i]["AC_ResistanceValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["AC_ResistanceValue"].ToString().Trim());
                            dcmGroupB1_StopL = dtAvg.Rows[i]["L_InductanceValue"] == null || dtAvg.Rows[i]["L_InductanceValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["L_InductanceValue"].ToString().Trim());
                            dcmGroupB1_StopC = dtAvg.Rows[i]["C_CapacitanceValue"] == null || dtAvg.Rows[i]["C_CapacitanceValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["C_CapacitanceValue"].ToString().Trim());
                            dcmGroupB1_StopQ = dtAvg.Rows[i]["Q_FactorValue"] == null || dtAvg.Rows[i]["Q_FactorValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["Q_FactorValue"].ToString().Trim());
                        }
                        else if (dtAvg.Rows[i]["SeqNumber"].ToString().Trim() == "2")
                        {
                            dcmGroupB2_StopDC = dtAvg.Rows[i]["DC_ResistanceValue"] == null || dtAvg.Rows[i]["DC_ResistanceValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["DC_ResistanceValue"].ToString().Trim());
                            dcmGroupB2_StopAC = dtAvg.Rows[i]["AC_ResistanceValue"] == null || dtAvg.Rows[i]["AC_ResistanceValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["AC_ResistanceValue"].ToString().Trim());
                            dcmGroupB2_StopL = dtAvg.Rows[i]["L_InductanceValue"] == null || dtAvg.Rows[i]["L_InductanceValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["L_InductanceValue"].ToString().Trim());
                            dcmGroupB2_StopC = dtAvg.Rows[i]["C_CapacitanceValue"] == null || dtAvg.Rows[i]["C_CapacitanceValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["C_CapacitanceValue"].ToString().Trim());
                            dcmGroupB2_StopQ = dtAvg.Rows[i]["Q_FactorValue"] == null || dtAvg.Rows[i]["Q_FactorValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["Q_FactorValue"].ToString().Trim());
                        }
                        else
                        {
                            dcmGroupB_StopDC = dtAvg.Rows[i]["DC_ResistanceValue"] == null || dtAvg.Rows[i]["DC_ResistanceValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["DC_ResistanceValue"].ToString().Trim());
                            dcmGroupB_StopAC = dtAvg.Rows[i]["AC_ResistanceValue"] == null || dtAvg.Rows[i]["AC_ResistanceValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["AC_ResistanceValue"].ToString().Trim());
                            dcmGroupB_StopL = dtAvg.Rows[i]["L_InductanceValue"] == null || dtAvg.Rows[i]["L_InductanceValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["L_InductanceValue"].ToString().Trim());
                            dcmGroupB_StopC = dtAvg.Rows[i]["C_CapacitanceValue"] == null || dtAvg.Rows[i]["C_CapacitanceValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["C_CapacitanceValue"].ToString().Trim());
                            dcmGroupB_StopQ = dtAvg.Rows[i]["Q_FactorValue"] == null || dtAvg.Rows[i]["Q_FactorValue"].ToString().Trim() == ""
                                ? 0M : Convert.ToDecimal(dtAvg.Rows[i]["Q_FactorValue"].ToString().Trim());
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
        }

		/// <summary>
		/// 코일 그룹A 제어용 평균치 산출
		/// </summary>
		/// <param name="_dt"></param>
		/// <param name="_dcmGroupA_ControlDC"></param>
		/// <param name="_dcmGroupA_ControlAC"></param>
		/// <param name="_dcmGroupA_ControlL"></param>
		/// <param name="_dcmGroupA_ControlC"></param>
		/// <param name="_dcmGroupA_ControlQ"></param>
		private void SetControlRodCoilGroupA_AverageCalculation(DataTable _dt
			, decimal _dcmGroupA1_ControlDC, decimal _dcmGroupA1_ControlAC, decimal _dcmGroupA1_ControlL, decimal _dcmGroupA1_ControlC, decimal _dcmGroupA1_ControlQ
			, decimal _dcmGroupA2_ControlDC, decimal _dcmGroupA2_ControlAC, decimal _dcmGroupA2_ControlL, decimal _dcmGroupA2_ControlC, decimal _dcmGroupA2_ControlQ
			, decimal _dcmGroupA_ControlDC, decimal _dcmGroupA_ControlAC, decimal _dcmGroupA_ControlL, decimal _dcmGroupA_ControlC, decimal _dcmGroupA_ControlQ)
		{
			try
			{
                string[] arrayCoilItem = Gini.GetValue("DRPI", "Control_Item").Split(',');

                decimal dcmASumControlRodRdc = 0.000M, dcmASumControlRodRac = 0.000M, dcmASumControlRodL = 0.000M, dcmASumControlRodC = 0.000000M, dcmASumControlRodQ = 0.000M;
                decimal dcmReferenceValue = Convert.ToDecimal(Gini.GetValue("ReferenceValue", "RdcDecisionRange_ReferenceValue"));
                decimal dcmMeasureValue = 0M, dcmChkValue = 0M;
				int intAAvgRdcCount = 0, intAAvgRacCount = 0, intAAvgLCount = 0, intAAvgCCount = 0, intAAvgQCount = 0;

				DataTable dtAverage = new DataTable();
				dtAverage.Columns.Add("CoilCabinetType");
				dtAverage.Columns.Add("CoilCabinetName");
				dtAverage.Columns.Add("Rdc_DRPIReferenceValue");
				dtAverage.Columns.Add("Rac_DRPIReferenceValue");
				dtAverage.Columns.Add("L_DRPIReferenceValue");
				dtAverage.Columns.Add("C_DRPIReferenceValue");
				dtAverage.Columns.Add("Q_DRPIReferenceValue");

				for (int i = 0; i < arrayCoilItem.Length; i++)
				{
					string strCoilName = arrayCoilItem[i].Trim();
					dcmASumControlRodRdc = 0.000M; dcmASumControlRodRac = 0.000M; dcmASumControlRodL = 0.000M; dcmASumControlRodC = 0.000000M; dcmASumControlRodQ = 0.000M;
					dcmMeasureValue = 0M; dcmChkValue = 0M;
					intAAvgRdcCount = 0; intAAvgRacCount = 0; intAAvgLCount = 0; intAAvgCCount = 0; intAAvgQCount = 0;

					for (int iRow = 0; iRow < _dt.Rows.Count; iRow++)
					{
						if (_dt.Rows[iRow]["DRPIGroup"].ToString().Trim() != "A") continue;
						if (_dt.Rows[iRow]["DRPIType"].ToString().Trim() != "제어용") continue;
						if (_dt.Rows[iRow]["CoilName"].ToString().Trim() != ("A" + Regex.Replace(strCoilName.Trim(), @"\D", ""))) continue;

						dcmMeasureValue = _dt.Rows[iRow]["DC_ResistanceValue"] == null || _dt.Rows[iRow]["DC_ResistanceValue"].ToString().Trim() == ""
							? 0.000M : Convert.ToDecimal(_dt.Rows[iRow]["DC_ResistanceValue"].ToString().Trim());

						if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "1")
							dcmMeasureValue = dcmMeasureValue - _dcmGroupA1_ControlDC;
						else if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "2")
							dcmMeasureValue = dcmMeasureValue - _dcmGroupA2_ControlDC;
						else
							dcmMeasureValue = dcmMeasureValue - _dcmGroupA_ControlDC;

						if (dcmMeasureValue < 0) dcmMeasureValue = -dcmMeasureValue;

						if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "1")
                            dcmChkValue = _dcmGroupA1_ControlDC * (dcmReferenceValue / 100);
						else if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "2")
                            dcmChkValue = _dcmGroupA2_ControlDC * (dcmReferenceValue / 100);
						else
                            dcmChkValue = _dcmGroupA_ControlDC * (dcmReferenceValue / 100);

						if (dcmMeasureValue <= dcmChkValue)
						{
							dcmASumControlRodRdc += _dt.Rows[iRow]["DC_ResistanceValue"] == null || _dt.Rows[iRow]["DC_ResistanceValue"].ToString().Trim() == "" ? 0.000M : Convert.ToDecimal(_dt.Rows[iRow]["DC_ResistanceValue"].ToString().Trim());
							intAAvgRdcCount++;
						}

						dcmMeasureValue = _dt.Rows[iRow]["AC_ResistanceValue"] == null || _dt.Rows[iRow]["AC_ResistanceValue"].ToString().Trim() == ""
							? 0.000M : Convert.ToDecimal(_dt.Rows[iRow]["AC_ResistanceValue"].ToString().Trim());

						if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "1")
							dcmMeasureValue = dcmMeasureValue - _dcmGroupA1_ControlAC;
						else if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "2")
							dcmMeasureValue = dcmMeasureValue - _dcmGroupA2_ControlAC;
						else
							dcmMeasureValue = dcmMeasureValue - _dcmGroupA_ControlAC;

						if (dcmMeasureValue < 0) dcmMeasureValue = -dcmMeasureValue;

						if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "1")
                            dcmChkValue = _dcmGroupA1_ControlAC * (dcmReferenceValue / 100);
						else if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "2")
                            dcmChkValue = _dcmGroupA2_ControlAC * (dcmReferenceValue / 100);
						else
                            dcmChkValue = _dcmGroupA_ControlAC * (dcmReferenceValue / 100);

						if (dcmMeasureValue <= dcmChkValue)
						{
							dcmASumControlRodRac += _dt.Rows[iRow]["AC_ResistanceValue"] == null || _dt.Rows[iRow]["AC_ResistanceValue"].ToString().Trim() == "" ? 0.000M : Convert.ToDecimal(_dt.Rows[iRow]["AC_ResistanceValue"].ToString().Trim());
							intAAvgRacCount++;
						}

						dcmMeasureValue = _dt.Rows[iRow]["L_InductanceValue"] == null || _dt.Rows[iRow]["L_InductanceValue"].ToString().Trim() == ""
							? 0.000M : Convert.ToDecimal(_dt.Rows[iRow]["L_InductanceValue"].ToString().Trim());

						if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "1")
							dcmMeasureValue = dcmMeasureValue - _dcmGroupA1_ControlL;
						else if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "2")
							dcmMeasureValue = dcmMeasureValue - _dcmGroupA2_ControlL;
						else
							dcmMeasureValue = dcmMeasureValue - _dcmGroupA_ControlL;

						if (dcmMeasureValue < 0) dcmMeasureValue = -dcmMeasureValue;

						if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "1")
                            dcmChkValue = _dcmGroupA1_ControlL * (dcmReferenceValue / 100);
						else if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "2")
                            dcmChkValue = _dcmGroupA2_ControlL * (dcmReferenceValue / 100);
						else
                            dcmChkValue = _dcmGroupA_ControlL * (dcmReferenceValue / 100);

						if (dcmMeasureValue <= dcmChkValue)
						{
							dcmASumControlRodL += _dt.Rows[iRow]["L_InductanceValue"] == null || _dt.Rows[iRow]["L_InductanceValue"].ToString().Trim() == "" ? 0.000M : Convert.ToDecimal(_dt.Rows[iRow]["L_InductanceValue"].ToString().Trim());
							intAAvgLCount++;
						}

						dcmMeasureValue = _dt.Rows[iRow]["C_CapacitanceValue"] == null || _dt.Rows[iRow]["C_CapacitanceValue"].ToString().Trim() == ""
							? 0.000M : Convert.ToDecimal(_dt.Rows[iRow]["C_CapacitanceValue"].ToString().Trim());

						if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "1")
							dcmMeasureValue = dcmMeasureValue - _dcmGroupA1_ControlC;
						else if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "2")
							dcmMeasureValue = dcmMeasureValue - _dcmGroupA2_ControlC;
						else
							dcmMeasureValue = dcmMeasureValue - _dcmGroupA_ControlC;

						if (dcmMeasureValue < 0) dcmMeasureValue = -dcmMeasureValue;

						if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "1")
							dcmChkValue = _dcmGroupA1_ControlC * (dcmReferenceValue / 100);
						else if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "2")
							dcmChkValue = _dcmGroupA2_ControlC * (dcmReferenceValue / 100);
						else
							dcmChkValue = _dcmGroupA_ControlC * (dcmReferenceValue / 100);

						if (dcmMeasureValue <= dcmChkValue)
						{
							dcmASumControlRodC += _dt.Rows[iRow]["C_CapacitanceValue"] == null || _dt.Rows[iRow]["C_CapacitanceValue"].ToString().Trim() == "" ? 0.000000M : Convert.ToDecimal(_dt.Rows[iRow]["C_CapacitanceValue"].ToString().Trim());
							intAAvgCCount++;
						}

						dcmMeasureValue = _dt.Rows[iRow]["Q_FactorValue"] == null || _dt.Rows[iRow]["Q_FactorValue"].ToString().Trim() == ""
							? 0.000M : Convert.ToDecimal(_dt.Rows[iRow]["Q_FactorValue"].ToString().Trim());

						if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "1")
							dcmMeasureValue = dcmMeasureValue - _dcmGroupA1_ControlQ;
						else if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "2")
							dcmMeasureValue = dcmMeasureValue - _dcmGroupA2_ControlQ;
						else
							dcmMeasureValue = dcmMeasureValue - _dcmGroupA_ControlQ;

						if (dcmMeasureValue < 0) dcmMeasureValue = -dcmMeasureValue;

						if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "1")
							dcmChkValue = _dcmGroupA1_ControlQ * (dcmReferenceValue / 100);
						else if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "2")
							dcmChkValue = _dcmGroupA2_ControlQ * (dcmReferenceValue / 100);
						else
							dcmChkValue = _dcmGroupA_ControlQ * (dcmReferenceValue / 100);

						if (dcmMeasureValue <= dcmChkValue)
						{
							dcmASumControlRodQ += _dt.Rows[iRow]["Q_FactorValue"] == null || _dt.Rows[iRow]["Q_FactorValue"].ToString().Trim() == "" ? 0.000M : Convert.ToDecimal(_dt.Rows[iRow]["Q_FactorValue"].ToString().Trim());
							intAAvgQCount++;
						}
					}

					if (intAAvgRdcCount == 0 && intAAvgRacCount == 0 && intAAvgLCount == 0 && intAAvgCCount == 0 && intAAvgQCount == 0)
						continue;

					// 그룹 A 제어용 평균치
					DataRow dr = dtAverage.NewRow();
					dr[0] = "제어용";
					dr[1] = "A" + Regex.Replace(strCoilName.Trim(), @"\D", "");
					
					if (intAAvgRdcCount > 0)
						dr[2] = (dcmASumControlRodRdc / intAAvgRdcCount).ToString("F3").Trim();
					else
						dr[2] = "0.000";

					if (intAAvgRacCount > 0)
						dr[3] = (dcmASumControlRodRac / intAAvgRacCount).ToString("F3").Trim();
					else
						dr[3] = "0.000";

					if (intAAvgLCount > 0)
						dr[4] = (dcmASumControlRodL / intAAvgLCount).ToString("F3").Trim();
					else
						dr[4] = "0.000";

					if (intAAvgCCount > 0)
						dr[5] = (dcmASumControlRodC / intAAvgCCount).ToString("F6").Trim();
					else
						dr[5] = "0.000";

					if (intAAvgQCount > 0)
						dr[6] = (dcmASumControlRodQ / intAAvgQCount).ToString("F3").Trim();
					else
						dr[6] = "0.000";

					dtAverage.Rows.Add(dr);
				}

				decimal dcmA3RdcSum = 0M, dcmA3RacSum = 0M, dcmA3LSum = 0M, dcmA3CSum = 0M, dcmA3QSum = 0M;
				int intA3SCount = 0;

				for (int i = 0; i < dtAverage.Rows.Count; i++)
				{
					if (dtAverage.Rows[i]["CoilCabinetName"].ToString().Trim() == "A1")
					{
						dAveA_ControlRodRdc[0] = dtAverage.Rows[i]["Rdc_DRPIReferenceValue"] == null || dtAverage.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.000M : Convert.ToDecimal(dtAverage.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim());
						dAveA_ControlRodRac[0] = dtAverage.Rows[i]["Rac_DRPIReferenceValue"] == null || dtAverage.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.000M : Convert.ToDecimal(dtAverage.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim());
						dAveA_ControlRodL[0] = dtAverage.Rows[i]["L_DRPIReferenceValue"] == null || dtAverage.Rows[i]["L_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.000M : Convert.ToDecimal(dtAverage.Rows[i]["L_DRPIReferenceValue"].ToString().Trim());
						dAveA_ControlRodC[0] = dtAverage.Rows[i]["C_DRPIReferenceValue"] == null || dtAverage.Rows[i]["C_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.000000M : Convert.ToDecimal(dtAverage.Rows[i]["C_DRPIReferenceValue"].ToString().Trim());
						dAveA_ControlRodQ[0] = dtAverage.Rows[i]["Q_DRPIReferenceValue"] == null || dtAverage.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.0M : Convert.ToDecimal(dtAverage.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim());
					}
					else if (dtAverage.Rows[i]["CoilCabinetName"].ToString().Trim() == "A2")
					{
						dAveA_ControlRodRdc[1] = dtAverage.Rows[i]["Rdc_DRPIReferenceValue"] == null || dtAverage.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.000M : Convert.ToDecimal(dtAverage.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim());
						dAveA_ControlRodRac[1] = dtAverage.Rows[i]["Rac_DRPIReferenceValue"] == null || dtAverage.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.000M : Convert.ToDecimal(dtAverage.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim());
						dAveA_ControlRodL[1] = dtAverage.Rows[i]["L_DRPIReferenceValue"] == null || dtAverage.Rows[i]["L_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.000M : Convert.ToDecimal(dtAverage.Rows[i]["L_DRPIReferenceValue"].ToString().Trim());
						dAveA_ControlRodC[1] = dtAverage.Rows[i]["C_DRPIReferenceValue"] == null || dtAverage.Rows[i]["C_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.000000M : Convert.ToDecimal(dtAverage.Rows[i]["C_DRPIReferenceValue"].ToString().Trim());
						dAveA_ControlRodQ[1] = dtAverage.Rows[i]["Q_DRPIReferenceValue"] == null || dtAverage.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.0M : Convert.ToDecimal(dtAverage.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim());
					}
					else
					{
						dcmA3RdcSum += dtAverage.Rows[i]["Rdc_DRPIReferenceValue"] == null || dtAverage.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.000M : Convert.ToDecimal(dtAverage.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim());
						dcmA3RacSum += dtAverage.Rows[i]["Rac_DRPIReferenceValue"] == null || dtAverage.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.000M : Convert.ToDecimal(dtAverage.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim());
						dcmA3LSum += dtAverage.Rows[i]["L_DRPIReferenceValue"] == null || dtAverage.Rows[i]["L_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.000M : Convert.ToDecimal(dtAverage.Rows[i]["L_DRPIReferenceValue"].ToString().Trim());
						dcmA3CSum += dtAverage.Rows[i]["C_DRPIReferenceValue"] == null || dtAverage.Rows[i]["C_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.000000M : Convert.ToDecimal(dtAverage.Rows[i]["C_DRPIReferenceValue"].ToString().Trim());
						dcmA3QSum += dtAverage.Rows[i]["Q_DRPIReferenceValue"] == null || dtAverage.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.0M : Convert.ToDecimal(dtAverage.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim());
						intA3SCount++;
					}
				}

				if (intA3SCount > 0)
				{
					dAveA_ControlRodRdc[2] = dcmA3RdcSum / intA3SCount;
					dAveA_ControlRodRac[2] = dcmA3RacSum / intA3SCount;
					dAveA_ControlRodL[2] = dcmA3LSum / intA3SCount;
					dAveA_ControlRodC[2] = dcmA3CSum / intA3SCount;
					dAveA_ControlRodQ[2] = dcmA3QSum / intA3SCount;
				}
				else
				{
					dAveA_ControlRodRdc[2] = 0.000M;
					dAveA_ControlRodRac[2] = 0.000M;
					dAveA_ControlRodL[2] = 0.000M;
					dAveA_ControlRodC[2] = 0.000000M;
					dAveA_ControlRodQ[2] = 0.000M;
				}
			}
			catch (Exception ex)
			{
				System.Diagnostics.Debug.Print(ex.Message);
			}
		}

		/// <summary>
		/// 코일 그룹B 제어용 평균치 산출
		/// </summary>
		/// <param name="_dt"></param>
		/// <param name="_dcmGroupB_ControlDC"></param>
		/// <param name="_dcmGroupB_ControlAC"></param>
		/// <param name="_dcmGroupB_ControlL"></param>
		/// <param name="_dcmGroupB_ControlC"></param>
		/// <param name="_dcmGroupB_ControlQ"></param>
		private void SetControlRodCoilGroupB_AverageCalculation(DataTable _dt
			, decimal _dcmGroupB1_ControlDC, decimal _dcmGroupB1_ControlAC, decimal _dcmGroupB1_ControlL, decimal _dcmGroupB1_ControlC, decimal _dcmGroupB1_ControlQ
			, decimal _dcmGroupB2_ControlDC, decimal _dcmGroupB2_ControlAC, decimal _dcmGroupB2_ControlL, decimal _dcmGroupB2_ControlC, decimal _dcmGroupB2_ControlQ
			, decimal _dcmGroupB_ControlDC, decimal _dcmGroupB_ControlAC, decimal _dcmGroupB_ControlL, decimal _dcmGroupB_ControlC, decimal _dcmGroupB_ControlQ)
		{
			try
			{
                string[] arrayCoilItem = Gini.GetValue("DRPI", "Control_Item").Split(',');

                decimal dcmBSumControlRodRdc = 0.000M, dcmBSumControlRodRac = 0.000M, dcmBSumControlRodL = 0.000M, dcmBSumControlRodC = 0.000000M, dcmBSumControlRodQ = 0.000M;
                decimal dcmReferenceValue = Convert.ToDecimal(Gini.GetValue("ReferenceValue", "RdcDecisionRange_ReferenceValue"));
                decimal dcmMeasureValue = 0M, dcmChkValue = 0M;
				int intBAvgRdcCount = 0, intBAvgRacCount = 0, intBAvgLCount = 0, intBAvgCCount = 0, intBAvgQCount = 0;

				DataTable dtAverage = new DataTable();
				dtAverage.Columns.Add("CoilCabinetType");
				dtAverage.Columns.Add("CoilCabinetName");
				dtAverage.Columns.Add("Rdc_DRPIReferenceValue");
				dtAverage.Columns.Add("Rac_DRPIReferenceValue");
				dtAverage.Columns.Add("L_DRPIReferenceValue");
				dtAverage.Columns.Add("C_DRPIReferenceValue");
				dtAverage.Columns.Add("Q_DRPIReferenceValue");

				for (int i = 0; i < arrayCoilItem.Length; i++)
				{
					string strCoilName = arrayCoilItem[i].Trim();
					dcmBSumControlRodRdc = 0.000M; dcmBSumControlRodRac = 0.000M; dcmBSumControlRodL = 0.000M; dcmBSumControlRodC = 0.000000M; dcmBSumControlRodQ = 0.000M;
					dcmMeasureValue = 0M; dcmChkValue = 0M;
					intBAvgRdcCount = 0; intBAvgRacCount = 0; intBAvgLCount = 0; intBAvgCCount = 0; intBAvgQCount = 0;

					for (int iRow = 0; iRow < _dt.Rows.Count; iRow++)
					{
						if (_dt.Rows[iRow]["DRPIGroup"].ToString().Trim() != "B") continue;
						if (_dt.Rows[iRow]["DRPIType"].ToString().Trim() != "제어용") continue;
						if (_dt.Rows[iRow]["CoilName"].ToString().Trim() != ("B" + Regex.Replace(strCoilName.Trim(), @"\D", ""))) continue;

						dcmMeasureValue = _dt.Rows[iRow]["DC_ResistanceValue"] == null || _dt.Rows[iRow]["DC_ResistanceValue"].ToString().Trim() == ""
							? 0.000M : Convert.ToDecimal(_dt.Rows[iRow]["DC_ResistanceValue"].ToString().Trim());

						if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "1")
							dcmMeasureValue = dcmMeasureValue - _dcmGroupB1_ControlDC;
						else if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "2")
							dcmMeasureValue = dcmMeasureValue - _dcmGroupB2_ControlDC;
						else
							dcmMeasureValue = dcmMeasureValue - _dcmGroupB_ControlDC;

						if (dcmMeasureValue < 0) dcmMeasureValue = -dcmMeasureValue;

						if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "1")
							dcmChkValue = _dcmGroupB1_ControlDC * (dcmReferenceValue / 100);
						else if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "2")
							dcmChkValue = _dcmGroupB2_ControlDC * (dcmReferenceValue / 100);
						else
							dcmChkValue = _dcmGroupB_ControlDC * (dcmReferenceValue / 100);

						if (dcmMeasureValue <= dcmChkValue)
						{
							dcmBSumControlRodRdc += _dt.Rows[iRow]["DC_ResistanceValue"] == null || _dt.Rows[iRow]["DC_ResistanceValue"].ToString().Trim() == "" ? 0.000M : Convert.ToDecimal(_dt.Rows[iRow]["DC_ResistanceValue"].ToString().Trim());
							intBAvgRdcCount++;
						}

						dcmMeasureValue = _dt.Rows[iRow]["AC_ResistanceValue"] == null || _dt.Rows[iRow]["AC_ResistanceValue"].ToString().Trim() == ""
							? 0.000M : Convert.ToDecimal(_dt.Rows[iRow]["AC_ResistanceValue"].ToString().Trim());

						if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "1")
							dcmMeasureValue = dcmMeasureValue - _dcmGroupB1_ControlAC;
						else if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "2")
							dcmMeasureValue = dcmMeasureValue - _dcmGroupB2_ControlAC;
						else
							dcmMeasureValue = dcmMeasureValue - _dcmGroupB_ControlAC;

						if (dcmMeasureValue < 0) dcmMeasureValue = -dcmMeasureValue;

						if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "1")
							dcmChkValue = _dcmGroupB1_ControlAC * (dcmReferenceValue / 100);
						else if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "2")
							dcmChkValue = _dcmGroupB2_ControlAC * (dcmReferenceValue / 100);
						else
							dcmChkValue = _dcmGroupB_ControlAC * (dcmReferenceValue / 100);

						if (dcmMeasureValue <= dcmChkValue)
						{
							dcmBSumControlRodRac += _dt.Rows[iRow]["AC_ResistanceValue"] == null || _dt.Rows[iRow]["AC_ResistanceValue"].ToString().Trim() == "" ? 0.000M : Convert.ToDecimal(_dt.Rows[iRow]["AC_ResistanceValue"].ToString().Trim());
							intBAvgRacCount++;
						}

						dcmMeasureValue = _dt.Rows[iRow]["L_InductanceValue"] == null || _dt.Rows[iRow]["L_InductanceValue"].ToString().Trim() == ""
							? 0.000M : Convert.ToDecimal(_dt.Rows[iRow]["L_InductanceValue"].ToString().Trim());

						if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "1")
							dcmMeasureValue = dcmMeasureValue - _dcmGroupB1_ControlL;
						else if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "2")
							dcmMeasureValue = dcmMeasureValue - _dcmGroupB2_ControlL;
						else
							dcmMeasureValue = dcmMeasureValue - _dcmGroupB_ControlL;

						if (dcmMeasureValue < 0) dcmMeasureValue = -dcmMeasureValue;

						if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "1")
							dcmChkValue = _dcmGroupB1_ControlL * (dcmReferenceValue / 100);
						else if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "2")
							dcmChkValue = _dcmGroupB2_ControlL * (dcmReferenceValue / 100);
						else
							dcmChkValue = _dcmGroupB_ControlL * (dcmReferenceValue / 100);

						if (dcmMeasureValue <= dcmChkValue)
						{
							dcmBSumControlRodL += _dt.Rows[iRow]["L_InductanceValue"] == null || _dt.Rows[iRow]["L_InductanceValue"].ToString().Trim() == "" ? 0.000M : Convert.ToDecimal(_dt.Rows[iRow]["L_InductanceValue"].ToString().Trim());
							intBAvgLCount++;
						}

						dcmMeasureValue = _dt.Rows[iRow]["C_CapacitanceValue"] == null || _dt.Rows[iRow]["C_CapacitanceValue"].ToString().Trim() == ""
							? 0.000M : Convert.ToDecimal(_dt.Rows[iRow]["C_CapacitanceValue"].ToString().Trim());

						if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "1")
							dcmMeasureValue = dcmMeasureValue - _dcmGroupB1_ControlC;
						else if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "2")
							dcmMeasureValue = dcmMeasureValue - _dcmGroupB2_ControlC;
						else
							dcmMeasureValue = dcmMeasureValue - _dcmGroupB_ControlC;

						if (dcmMeasureValue < 0) dcmMeasureValue = -dcmMeasureValue;

						if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "1")
							dcmChkValue = _dcmGroupB1_ControlC * (dcmReferenceValue / 100);
						else if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "2")
							dcmChkValue = _dcmGroupB2_ControlC * (dcmReferenceValue / 100);
						else
							dcmChkValue = _dcmGroupB_ControlC * (dcmReferenceValue / 100);

						if (dcmMeasureValue <= dcmChkValue)
						{
							dcmBSumControlRodC += _dt.Rows[iRow]["C_CapacitanceValue"] == null || _dt.Rows[iRow]["C_CapacitanceValue"].ToString().Trim() == "" ? 0.000000M : Convert.ToDecimal(_dt.Rows[iRow]["C_CapacitanceValue"].ToString().Trim());
							intBAvgCCount++;
						}

						dcmMeasureValue = _dt.Rows[iRow]["Q_FactorValue"] == null || _dt.Rows[iRow]["Q_FactorValue"].ToString().Trim() == ""
							? 0.000M : Convert.ToDecimal(_dt.Rows[iRow]["Q_FactorValue"].ToString().Trim());

						if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "1")
							dcmMeasureValue = dcmMeasureValue - _dcmGroupB1_ControlQ;
						else if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "2")
							dcmMeasureValue = dcmMeasureValue - _dcmGroupB2_ControlQ;
						else
							dcmMeasureValue = dcmMeasureValue - _dcmGroupB_ControlQ;

						if (dcmMeasureValue < 0) dcmMeasureValue = -dcmMeasureValue;

						if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "1")
							dcmChkValue = _dcmGroupB1_ControlQ * (dcmReferenceValue / 100);
						else if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "2")
							dcmChkValue = _dcmGroupB2_ControlQ * (dcmReferenceValue / 100);
						else
							dcmChkValue = _dcmGroupB_ControlQ * (dcmReferenceValue / 100);

						if (dcmMeasureValue <= dcmChkValue)
						{
							dcmBSumControlRodQ += _dt.Rows[iRow]["Q_FactorValue"] == null || _dt.Rows[iRow]["Q_FactorValue"].ToString().Trim() == "" ? 0.000M : Convert.ToDecimal(_dt.Rows[iRow]["Q_FactorValue"].ToString().Trim());
							intBAvgQCount++;
						}
					}

					if (intBAvgRdcCount == 0 && intBAvgRacCount == 0 && intBAvgLCount == 0 && intBAvgCCount == 0 && intBAvgQCount == 0)
						continue;

					// 그룹 B 제어용 평균치
					DataRow dr = dtAverage.NewRow();
					dr[0] = "제어용";
					dr[1] = "B" + Regex.Replace(strCoilName.Trim(), @"\D", "");

					if (intBAvgRdcCount > 0)
						dr[2] = (dcmBSumControlRodRdc / intBAvgRdcCount).ToString("F3").Trim();
					else
						dr[2] = "0.000";

					if (intBAvgRacCount > 0)
						dr[3] = (dcmBSumControlRodRac / intBAvgRacCount).ToString("F3").Trim();
					else
						dr[3] = "0.000";

					if (intBAvgLCount > 0)
						dr[4] = (dcmBSumControlRodL / intBAvgLCount).ToString("F3").Trim();
					else
						dr[4] = "0.000";

					if (intBAvgCCount > 0)
						dr[5] = (dcmBSumControlRodC / intBAvgCCount).ToString("F6").Trim();
					else
						dr[5] = "0.000";

					if (intBAvgQCount > 0)
						dr[6] = (dcmBSumControlRodQ / intBAvgQCount).ToString("F3").Trim();
					else
						dr[6] = "0.000";

					dtAverage.Rows.Add(dr);
				}

				decimal dcmB3RdcSum = 0M, dcmB3RacSum = 0M, dcmB3LSum = 0M, dcmB3CSum = 0M, dcmB3QSum = 0M;
				int intB3SCount = 0;

				for (int i = 0; i < dtAverage.Rows.Count; i++)
				{
					if (dtAverage.Rows[i]["CoilCabinetName"].ToString().Trim() == "B1")
					{
						dAveB_ControlRodRdc[0] = dtAverage.Rows[i]["Rdc_DRPIReferenceValue"] == null || dtAverage.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.000M : Convert.ToDecimal(dtAverage.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim());
						dAveB_ControlRodRac[0] = dtAverage.Rows[i]["Rac_DRPIReferenceValue"] == null || dtAverage.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.000M : Convert.ToDecimal(dtAverage.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim());
						dAveB_ControlRodL[0] = dtAverage.Rows[i]["L_DRPIReferenceValue"] == null || dtAverage.Rows[i]["L_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.000M : Convert.ToDecimal(dtAverage.Rows[i]["L_DRPIReferenceValue"].ToString().Trim());
						dAveB_ControlRodC[0] = dtAverage.Rows[i]["C_DRPIReferenceValue"] == null || dtAverage.Rows[i]["C_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.000000M : Convert.ToDecimal(dtAverage.Rows[i]["C_DRPIReferenceValue"].ToString().Trim());
						dAveB_ControlRodQ[0] = dtAverage.Rows[i]["Q_DRPIReferenceValue"] == null || dtAverage.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.0M : Convert.ToDecimal(dtAverage.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim());
					}
					else if (dtAverage.Rows[i]["CoilCabinetName"].ToString().Trim() == "B2")
					{
						dAveB_ControlRodRdc[1] = dtAverage.Rows[i]["Rdc_DRPIReferenceValue"] == null || dtAverage.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.000M : Convert.ToDecimal(dtAverage.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim());
						dAveB_ControlRodRac[1] = dtAverage.Rows[i]["Rac_DRPIReferenceValue"] == null || dtAverage.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.000M : Convert.ToDecimal(dtAverage.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim());
						dAveB_ControlRodL[1] = dtAverage.Rows[i]["L_DRPIReferenceValue"] == null || dtAverage.Rows[i]["L_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.000M : Convert.ToDecimal(dtAverage.Rows[i]["L_DRPIReferenceValue"].ToString().Trim());
						dAveB_ControlRodC[1] = dtAverage.Rows[i]["C_DRPIReferenceValue"] == null || dtAverage.Rows[i]["C_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.000000M : Convert.ToDecimal(dtAverage.Rows[i]["C_DRPIReferenceValue"].ToString().Trim());
						dAveB_ControlRodQ[1] = dtAverage.Rows[i]["Q_DRPIReferenceValue"] == null || dtAverage.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.0M : Convert.ToDecimal(dtAverage.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim());
					}
					else
					{
						dcmB3RdcSum += dtAverage.Rows[i]["Rdc_DRPIReferenceValue"] == null || dtAverage.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.000M : Convert.ToDecimal(dtAverage.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim());
						dcmB3RacSum += dtAverage.Rows[i]["Rac_DRPIReferenceValue"] == null || dtAverage.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.000M : Convert.ToDecimal(dtAverage.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim());
						dcmB3LSum += dtAverage.Rows[i]["L_DRPIReferenceValue"] == null || dtAverage.Rows[i]["L_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.000M : Convert.ToDecimal(dtAverage.Rows[i]["L_DRPIReferenceValue"].ToString().Trim());
						dcmB3CSum += dtAverage.Rows[i]["C_DRPIReferenceValue"] == null || dtAverage.Rows[i]["C_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.000000M : Convert.ToDecimal(dtAverage.Rows[i]["C_DRPIReferenceValue"].ToString().Trim());
						dcmB3QSum += dtAverage.Rows[i]["Q_DRPIReferenceValue"] == null || dtAverage.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.0M : Convert.ToDecimal(dtAverage.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim());
						intB3SCount++;
					}
				}

				if (intB3SCount > 0)
				{
					dAveB_ControlRodRdc[2] = dcmB3RdcSum / intB3SCount;
					dAveB_ControlRodRac[2] = dcmB3RacSum / intB3SCount;
					dAveB_ControlRodL[2] = dcmB3LSum / intB3SCount;
					dAveB_ControlRodC[2] = dcmB3CSum / intB3SCount;
					dAveB_ControlRodQ[2] = dcmB3QSum / intB3SCount;
				}
				else
				{
					dAveB_ControlRodRdc[2] = 0.000M;
					dAveB_ControlRodRac[2] = 0.000M;
					dAveB_ControlRodL[2] = 0.000M;
					dAveB_ControlRodC[2] = 0.000000M;
					dAveB_ControlRodQ[2] = 0.000M;
				}
			}
			catch (Exception ex)
			{
				System.Diagnostics.Debug.Print(ex.Message);
			}
		}

		/// <summary>
		/// 코일 그룹A 정지용 평균치 산출
		/// </summary>
		/// <param name="_dt"></param>
		/// <param name="_dcmGroupA_ControlDC"></param>
		/// <param name="_dcmGroupA_ControlAC"></param>
		/// <param name="_dcmGroupA_ControlL"></param>
		/// <param name="_dcmGroupA_ControlC"></param>
		/// <param name="_dcmGroupA_ControlQ"></param>
		private void SetStopCoilGroupA_AverageCalculation(DataTable _dt
			, decimal _dcmGroupA1_StopDC, decimal _dcmGroupA1_StopAC, decimal _dcmGroupA1_StopL, decimal _dcmGroupA1_StopC, decimal _dcmGroupA1_StopQ
			, decimal _dcmGroupA2_StopDC, decimal _dcmGroupA2_StopAC, decimal _dcmGroupA2_StopL, decimal _dcmGroupA2_StopC, decimal _dcmGroupA2_StopQ
			, decimal _dcmGroupA_StopDC, decimal _dcmGroupA_StopAC, decimal _dcmGroupA_StopL, decimal _dcmGroupA_StopC, decimal _dcmGroupA_StopQ)
		{
			try
			{
                string[] arrayCoilItem = Gini.GetValue("DRPI", "Stop_Item").Split(',');

                decimal dcmASumStopRdc = 0.000M, dcmASumStopRac = 0.000M, dcmASumStopL = 0.000M, dcmASumStopC = 0.000000M, dcmASumStopQ = 0.000M;
                decimal dcmReferenceValue = Convert.ToDecimal(Gini.GetValue("ReferenceValue", "RdcDecisionRange_ReferenceValue"));
                decimal dcmMeasureValue = 0M, dcmChkValue = 0M;
				int intAAvgRdcCount = 0, intAAvgRacCount = 0, intAAvgLCount = 0, intAAvgCCount = 0, intAAvgQCount = 0;

				DataTable dtAverage = new DataTable();
				dtAverage.Columns.Add("CoilCabinetType");
				dtAverage.Columns.Add("CoilCabinetName");
				dtAverage.Columns.Add("Rdc_DRPIReferenceValue");
				dtAverage.Columns.Add("Rac_DRPIReferenceValue");
				dtAverage.Columns.Add("L_DRPIReferenceValue");
				dtAverage.Columns.Add("C_DRPIReferenceValue");
				dtAverage.Columns.Add("Q_DRPIReferenceValue");

				for (int i = 0; i < arrayCoilItem.Length; i++)
				{
					string strCoilName = arrayCoilItem[i].Trim();
					dcmASumStopRdc = 0.000M; dcmASumStopRac = 0.000M; dcmASumStopL = 0.000M; dcmASumStopC = 0.000000M; dcmASumStopQ = 0.000M;
					dcmMeasureValue = 0M; dcmChkValue = 0M;
					intAAvgRdcCount = 0; intAAvgRacCount = 0; intAAvgLCount = 0; intAAvgCCount = 0; intAAvgQCount = 0;

					for (int iRow = 0; iRow < _dt.Rows.Count; iRow++)
					{
						if (_dt.Rows[iRow]["DRPIGroup"].ToString().Trim() != "A") continue;
						if (_dt.Rows[iRow]["DRPIType"].ToString().Trim() != "정지용") continue;
						if (_dt.Rows[iRow]["CoilName"].ToString().Trim() != ("A" + Regex.Replace(strCoilName.Trim(), @"\D", ""))) continue;

						dcmMeasureValue = _dt.Rows[iRow]["DC_ResistanceValue"] == null || _dt.Rows[iRow]["DC_ResistanceValue"].ToString().Trim() == ""
							? 0.000M : Convert.ToDecimal(_dt.Rows[iRow]["DC_ResistanceValue"].ToString().Trim());

						if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "1")
							dcmMeasureValue = dcmMeasureValue - _dcmGroupA1_StopDC;
						else if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "2")
							dcmMeasureValue = dcmMeasureValue - _dcmGroupA2_StopDC;
						else
							dcmMeasureValue = dcmMeasureValue - _dcmGroupA_StopDC;

						if (dcmMeasureValue < 0) dcmMeasureValue = -dcmMeasureValue;

						if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "1")
							dcmChkValue = _dcmGroupA1_StopDC * (dcmReferenceValue / 100);
						else if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "2")
							dcmChkValue = _dcmGroupA2_StopDC * (dcmReferenceValue / 100);
						else
							dcmChkValue = _dcmGroupA_StopDC * (dcmReferenceValue / 100);

						if (dcmMeasureValue <= dcmChkValue)
						{
							dcmASumStopRdc += _dt.Rows[iRow]["DC_ResistanceValue"] == null || _dt.Rows[iRow]["DC_ResistanceValue"].ToString().Trim() == "" ? 0.000M : Convert.ToDecimal(_dt.Rows[iRow]["DC_ResistanceValue"].ToString().Trim());
							intAAvgRdcCount++;
						}

						dcmMeasureValue = _dt.Rows[iRow]["AC_ResistanceValue"] == null || _dt.Rows[iRow]["AC_ResistanceValue"].ToString().Trim() == ""
							? 0.000M : Convert.ToDecimal(_dt.Rows[iRow]["AC_ResistanceValue"].ToString().Trim());

						if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "1")
							dcmMeasureValue = dcmMeasureValue - _dcmGroupA1_StopAC;
						else if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "2")
							dcmMeasureValue = dcmMeasureValue - _dcmGroupA2_StopAC;
						else
							dcmMeasureValue = dcmMeasureValue - _dcmGroupA_StopAC;

						if (dcmMeasureValue < 0) dcmMeasureValue = -dcmMeasureValue;

						if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "1")
							dcmChkValue = _dcmGroupA1_StopAC * (dcmReferenceValue / 100);
						else if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "2")
							dcmChkValue = _dcmGroupA2_StopAC * (dcmReferenceValue / 100);
						else
							dcmChkValue = _dcmGroupA_StopAC * (dcmReferenceValue / 100);

						if (dcmMeasureValue <= dcmChkValue)
						{
							dcmASumStopRac += _dt.Rows[iRow]["AC_ResistanceValue"] == null || _dt.Rows[iRow]["AC_ResistanceValue"].ToString().Trim() == "" ? 0.000M : Convert.ToDecimal(_dt.Rows[iRow]["AC_ResistanceValue"].ToString().Trim());
							intAAvgRacCount++;
						}

						dcmMeasureValue = _dt.Rows[iRow]["L_InductanceValue"] == null || _dt.Rows[iRow]["L_InductanceValue"].ToString().Trim() == ""
							? 0.000M : Convert.ToDecimal(_dt.Rows[iRow]["L_InductanceValue"].ToString().Trim());

						if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "1")
							dcmMeasureValue = dcmMeasureValue - _dcmGroupA1_StopL;
						else if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "2")
							dcmMeasureValue = dcmMeasureValue - _dcmGroupA2_StopL;
						else
							dcmMeasureValue = dcmMeasureValue - _dcmGroupA_StopL;

						if (dcmMeasureValue < 0) dcmMeasureValue = -dcmMeasureValue;

						if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "1")
							dcmChkValue = _dcmGroupA1_StopL * (dcmReferenceValue / 100);
						else if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "2")
							dcmChkValue = _dcmGroupA2_StopL * (dcmReferenceValue / 100);
						else
							dcmChkValue = _dcmGroupA_StopL * (dcmReferenceValue / 100);

						if (dcmMeasureValue <= dcmChkValue)
						{
							dcmASumStopL += _dt.Rows[iRow]["L_InductanceValue"] == null || _dt.Rows[iRow]["L_InductanceValue"].ToString().Trim() == "" ? 0.000M : Convert.ToDecimal(_dt.Rows[iRow]["L_InductanceValue"].ToString().Trim());
							intAAvgLCount++;
						}

						dcmMeasureValue = _dt.Rows[iRow]["C_CapacitanceValue"] == null || _dt.Rows[iRow]["C_CapacitanceValue"].ToString().Trim() == ""
							? 0.000M : Convert.ToDecimal(_dt.Rows[iRow]["C_CapacitanceValue"].ToString().Trim());

						if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "1")
							dcmMeasureValue = dcmMeasureValue - _dcmGroupA1_StopC;
						else if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "2")
							dcmMeasureValue = dcmMeasureValue - _dcmGroupA2_StopC;
						else
							dcmMeasureValue = dcmMeasureValue - _dcmGroupA_StopC;

						if (dcmMeasureValue < 0) dcmMeasureValue = -dcmMeasureValue;

						if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "1")
							dcmChkValue = _dcmGroupA1_StopC * (dcmReferenceValue / 100);
						else if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "2")
							dcmChkValue = _dcmGroupA2_StopC * (dcmReferenceValue / 100);
						else
							dcmChkValue = _dcmGroupA_StopC * (dcmReferenceValue / 100);

						if (dcmMeasureValue <= dcmChkValue)
						{
							dcmASumStopC += _dt.Rows[iRow]["C_CapacitanceValue"] == null || _dt.Rows[iRow]["C_CapacitanceValue"].ToString().Trim() == "" ? 0.000000M : Convert.ToDecimal(_dt.Rows[iRow]["C_CapacitanceValue"].ToString().Trim());
							intAAvgCCount++;
						}

						dcmMeasureValue = _dt.Rows[iRow]["Q_FactorValue"] == null || _dt.Rows[iRow]["Q_FactorValue"].ToString().Trim() == ""
							? 0.000M : Convert.ToDecimal(_dt.Rows[iRow]["Q_FactorValue"].ToString().Trim());

						if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "1")
							dcmMeasureValue = dcmMeasureValue - _dcmGroupA1_StopQ;
						else if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "2")
							dcmMeasureValue = dcmMeasureValue - _dcmGroupA2_StopQ;
						else
							dcmMeasureValue = dcmMeasureValue - _dcmGroupA_StopQ;

						if (dcmMeasureValue < 0) dcmMeasureValue = -dcmMeasureValue;

						if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "1")
							dcmChkValue = _dcmGroupA1_StopQ * (dcmReferenceValue / 100);
						else if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "2")
							dcmChkValue = _dcmGroupA2_StopQ * (dcmReferenceValue / 100);
						else
							dcmChkValue = _dcmGroupA_StopQ * (dcmReferenceValue / 100);

						if (dcmMeasureValue <= dcmChkValue)
						{
							dcmASumStopQ += _dt.Rows[iRow]["Q_FactorValue"] == null || _dt.Rows[iRow]["Q_FactorValue"].ToString().Trim() == "" ? 0.000M : Convert.ToDecimal(_dt.Rows[iRow]["Q_FactorValue"].ToString().Trim());
							intAAvgQCount++;
						}
					}

					if (intAAvgRdcCount == 0 && intAAvgRacCount == 0 && intAAvgLCount == 0 && intAAvgCCount == 0 && intAAvgQCount == 0)
						continue;

					// 그룹 A 정지용 평균치
					DataRow dr = dtAverage.NewRow();
					dr[0] = "정지용";
					dr[1] = "A" + Regex.Replace(strCoilName.Trim(), @"\D", "");

					if (intAAvgRdcCount > 0)
						dr[2] = (dcmASumStopRdc / intAAvgRdcCount).ToString("F3").Trim();
					else
						dr[2] = "0.000";

					if (intAAvgRacCount > 0)
						dr[3] = (dcmASumStopRac / intAAvgRacCount).ToString("F3").Trim();
					else
						dr[3] = "0.000";

					if (intAAvgLCount > 0)
						dr[4] = (dcmASumStopL / intAAvgLCount).ToString("F3").Trim();
					else
						dr[4] = "0.000";

					if (intAAvgCCount > 0)
						dr[5] = (dcmASumStopC / intAAvgCCount).ToString("F6").Trim();
					else
						dr[5] = "0.000";

					if (intAAvgQCount > 0)
						dr[6] = (dcmASumStopQ / intAAvgQCount).ToString("F3").Trim();
					else
						dr[6] = "0.000";

					dtAverage.Rows.Add(dr);
				}

				decimal dcmA3RdcSum = 0M, dcmA3RacSum = 0M, dcmA3LSum = 0M, dcmA3CSum = 0M, dcmA3QSum = 0M;
				int intA3SCount = 0;

				for (int i = 0; i < dtAverage.Rows.Count; i++)
				{
					if (dtAverage.Rows[i]["CoilCabinetName"].ToString().Trim() == "A1")
					{
						dAveA_StopRdc[0] = dtAverage.Rows[i]["Rdc_DRPIReferenceValue"] == null || dtAverage.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.000M : Convert.ToDecimal(dtAverage.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim());
						dAveA_StopRac[0] = dtAverage.Rows[i]["Rac_DRPIReferenceValue"] == null || dtAverage.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.000M : Convert.ToDecimal(dtAverage.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim());
						dAveA_StopL[0] = dtAverage.Rows[i]["L_DRPIReferenceValue"] == null || dtAverage.Rows[i]["L_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.000M : Convert.ToDecimal(dtAverage.Rows[i]["L_DRPIReferenceValue"].ToString().Trim());
						dAveA_StopC[0] = dtAverage.Rows[i]["C_DRPIReferenceValue"] == null || dtAverage.Rows[i]["C_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.000000M : Convert.ToDecimal(dtAverage.Rows[i]["C_DRPIReferenceValue"].ToString().Trim());
						dAveA_StopQ[0] = dtAverage.Rows[i]["Q_DRPIReferenceValue"] == null || dtAverage.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.0M : Convert.ToDecimal(dtAverage.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim());
					}
					else if (dtAverage.Rows[i]["CoilCabinetName"].ToString().Trim() == "A2")
					{
						dAveA_StopRdc[1] = dtAverage.Rows[i]["Rdc_DRPIReferenceValue"] == null || dtAverage.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.000M : Convert.ToDecimal(dtAverage.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim());
						dAveA_StopRac[1] = dtAverage.Rows[i]["Rac_DRPIReferenceValue"] == null || dtAverage.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.000M : Convert.ToDecimal(dtAverage.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim());
						dAveA_StopL[1] = dtAverage.Rows[i]["L_DRPIReferenceValue"] == null || dtAverage.Rows[i]["L_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.000M : Convert.ToDecimal(dtAverage.Rows[i]["L_DRPIReferenceValue"].ToString().Trim());
						dAveA_StopC[1] = dtAverage.Rows[i]["C_DRPIReferenceValue"] == null || dtAverage.Rows[i]["C_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.000000M : Convert.ToDecimal(dtAverage.Rows[i]["C_DRPIReferenceValue"].ToString().Trim());
						dAveA_StopQ[1] = dtAverage.Rows[i]["Q_DRPIReferenceValue"] == null || dtAverage.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.0M : Convert.ToDecimal(dtAverage.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim());
					}
					else
					{
						dcmA3RdcSum += dtAverage.Rows[i]["Rdc_DRPIReferenceValue"] == null || dtAverage.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.000M : Convert.ToDecimal(dtAverage.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim());
						dcmA3RacSum += dtAverage.Rows[i]["Rac_DRPIReferenceValue"] == null || dtAverage.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.000M : Convert.ToDecimal(dtAverage.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim());
						dcmA3LSum += dtAverage.Rows[i]["L_DRPIReferenceValue"] == null || dtAverage.Rows[i]["L_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.000M : Convert.ToDecimal(dtAverage.Rows[i]["L_DRPIReferenceValue"].ToString().Trim());
						dcmA3CSum += dtAverage.Rows[i]["C_DRPIReferenceValue"] == null || dtAverage.Rows[i]["C_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.000000M : Convert.ToDecimal(dtAverage.Rows[i]["C_DRPIReferenceValue"].ToString().Trim());
						dcmA3QSum += dtAverage.Rows[i]["Q_DRPIReferenceValue"] == null || dtAverage.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.0M : Convert.ToDecimal(dtAverage.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim());
						intA3SCount++;
					}
				}

				if (intA3SCount > 0)
				{
					dAveA_StopRdc[2] = dcmA3RdcSum / intA3SCount;
					dAveA_StopRac[2] = dcmA3RacSum / intA3SCount;
					dAveA_StopL[2] = dcmA3LSum / intA3SCount;
					dAveA_StopC[2] = dcmA3CSum / intA3SCount;
					dAveA_StopQ[2] = dcmA3QSum / intA3SCount;
				}
				else
				{
					dAveA_StopRdc[2] = 0.000M;
					dAveA_StopRac[2] = 0.000M;
					dAveA_StopL[2] = 0.000M;
					dAveA_StopC[2] = 0.000000M;
					dAveA_StopQ[2] = 0.000M;
				}
			}
			catch (Exception ex)
			{
				System.Diagnostics.Debug.Print(ex.Message);
			}
		}

		/// <summary>
		/// 코일 그룹B 정지용 평균치 산출
		/// </summary>
		/// <param name="_dt"></param>
		/// <param name="_dcmGroupB_ControlDC"></param>
		/// <param name="_dcmGroupB_ControlAC"></param>
		/// <param name="_dcmGroupB_ControlL"></param>
		/// <param name="_dcmGroupB_ControlC"></param>
		/// <param name="_dcmGroupB_ControlQ"></param>
		private void SetStopCoilGroupB_AverageCalculation(DataTable _dt
			, decimal _dcmGroupB1_StopDC, decimal _dcmGroupB1_StopAC, decimal _dcmGroupB1_StopL, decimal _dcmGroupB1_StopC, decimal _dcmGroupB1_StopQ
			, decimal _dcmGroupB2_StopDC, decimal _dcmGroupB2_StopAC, decimal _dcmGroupB2_StopL, decimal _dcmGroupB2_StopC, decimal _dcmGroupB2_StopQ
			, decimal _dcmGroupB_StopDC, decimal _dcmGroupB_StopAC, decimal _dcmGroupB_StopL, decimal _dcmGroupB_StopC, decimal _dcmGroupB_StopQ)
		{
			try
			{
                string[] arrayCoilItem = Gini.GetValue("DRPI", "Stop_Item").Split(',');

                decimal dcmBSumStopRdc = 0.000M, dcmBSumStopRac = 0.000M, dcmBSumStopL = 0.000M, dcmBSumStopC = 0.000000M, dcmBSumStopQ = 0.000M;
                decimal dcmReferenceValue = Convert.ToDecimal(Gini.GetValue("ReferenceValue", "RdcDecisionRange_ReferenceValue"));
                decimal dcmMeasureValue = 0M, dcmChkValue = 0M;
				int intBAvgRdcCount = 0, intBAvgRacCount = 0, intBAvgLCount = 0, intBAvgCCount = 0, intBAvgQCount = 0;

				DataTable dtAverage = new DataTable();
				dtAverage.Columns.Add("CoilCabinetType");
				dtAverage.Columns.Add("CoilCabinetName");
				dtAverage.Columns.Add("Rdc_DRPIReferenceValue");
				dtAverage.Columns.Add("Rac_DRPIReferenceValue");
				dtAverage.Columns.Add("L_DRPIReferenceValue");
				dtAverage.Columns.Add("C_DRPIReferenceValue");
				dtAverage.Columns.Add("Q_DRPIReferenceValue");

				for (int i = 0; i < arrayCoilItem.Length; i++)
				{
					string strCoilName = arrayCoilItem[i].Trim();
					dcmBSumStopRdc = 0.000M; dcmBSumStopRac = 0.000M; dcmBSumStopL = 0.000M; dcmBSumStopC = 0.000000M; dcmBSumStopQ = 0.000M;
					dcmMeasureValue = 0M; dcmChkValue = 0M;
					intBAvgRdcCount = 0; intBAvgRacCount = 0; intBAvgLCount = 0; intBAvgCCount = 0; intBAvgQCount = 0;

					for (int iRow = 0; iRow < _dt.Rows.Count; iRow++)
					{
						if (_dt.Rows[iRow]["DRPIGroup"].ToString().Trim() != "B") continue;
						if (_dt.Rows[iRow]["DRPIType"].ToString().Trim() != "정지용") continue;
						if (_dt.Rows[iRow]["CoilName"].ToString().Trim() != ("B" + Regex.Replace(strCoilName.Trim(), @"\D", ""))) continue;

						dcmMeasureValue = _dt.Rows[iRow]["DC_ResistanceValue"] == null || _dt.Rows[iRow]["DC_ResistanceValue"].ToString().Trim() == ""
							? 0.000M : Convert.ToDecimal(_dt.Rows[iRow]["DC_ResistanceValue"].ToString().Trim());

						if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "1")
							dcmMeasureValue = dcmMeasureValue - _dcmGroupB1_StopDC;
						else if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "2")
							dcmMeasureValue = dcmMeasureValue - _dcmGroupB2_StopDC;
						else
							dcmMeasureValue = dcmMeasureValue - _dcmGroupB_StopDC;

						if (dcmMeasureValue < 0) dcmMeasureValue = -dcmMeasureValue;

						if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "1")
							dcmChkValue = _dcmGroupB1_StopDC * (dcmReferenceValue / 100);
						else if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "2")
							dcmChkValue = _dcmGroupB2_StopDC * (dcmReferenceValue / 100);
						else
							dcmChkValue = _dcmGroupB_StopDC * (dcmReferenceValue / 100);

						if (dcmMeasureValue <= dcmChkValue)
						{
							dcmBSumStopRdc += _dt.Rows[iRow]["DC_ResistanceValue"] == null || _dt.Rows[iRow]["DC_ResistanceValue"].ToString().Trim() == "" ? 0.000M : Convert.ToDecimal(_dt.Rows[iRow]["DC_ResistanceValue"].ToString().Trim());
							intBAvgRdcCount++;
						}

						dcmMeasureValue = _dt.Rows[iRow]["AC_ResistanceValue"] == null || _dt.Rows[iRow]["AC_ResistanceValue"].ToString().Trim() == ""
							? 0.000M : Convert.ToDecimal(_dt.Rows[iRow]["AC_ResistanceValue"].ToString().Trim());

						if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "1")
							dcmMeasureValue = dcmMeasureValue - _dcmGroupB1_StopAC;
						else if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "2")
							dcmMeasureValue = dcmMeasureValue - _dcmGroupB2_StopAC;
						else
							dcmMeasureValue = dcmMeasureValue - _dcmGroupB_StopAC;

						if (dcmMeasureValue < 0) dcmMeasureValue = -dcmMeasureValue;

						if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "1")
							dcmChkValue = _dcmGroupB1_StopAC * (dcmReferenceValue / 100);
						else if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "2")
							dcmChkValue = _dcmGroupB2_StopAC * (dcmReferenceValue / 100);
						else
							dcmChkValue = _dcmGroupB_StopAC * (dcmReferenceValue / 100);

						if (dcmMeasureValue <= dcmChkValue)
						{
							dcmBSumStopRac += _dt.Rows[iRow]["AC_ResistanceValue"] == null || _dt.Rows[iRow]["AC_ResistanceValue"].ToString().Trim() == "" ? 0.000M : Convert.ToDecimal(_dt.Rows[iRow]["AC_ResistanceValue"].ToString().Trim());
							intBAvgRacCount++;
						}

						dcmMeasureValue = _dt.Rows[iRow]["L_InductanceValue"] == null || _dt.Rows[iRow]["L_InductanceValue"].ToString().Trim() == ""
							? 0.000M : Convert.ToDecimal(_dt.Rows[iRow]["L_InductanceValue"].ToString().Trim());

						if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "1")
							dcmMeasureValue = dcmMeasureValue - _dcmGroupB1_StopL;
						else if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "2")
							dcmMeasureValue = dcmMeasureValue - _dcmGroupB2_StopL;
						else
							dcmMeasureValue = dcmMeasureValue - _dcmGroupB_StopL;

						if (dcmMeasureValue < 0) dcmMeasureValue = -dcmMeasureValue;

						if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "1")
							dcmChkValue = _dcmGroupB1_StopL * (dcmReferenceValue / 100);
						else if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "2")
							dcmChkValue = _dcmGroupB2_StopL * (dcmReferenceValue / 100);
						else
							dcmChkValue = _dcmGroupB_StopL * (dcmReferenceValue / 100);

						if (dcmMeasureValue <= dcmChkValue)
						{
							dcmBSumStopL += _dt.Rows[iRow]["L_InductanceValue"] == null || _dt.Rows[iRow]["L_InductanceValue"].ToString().Trim() == "" ? 0.000M : Convert.ToDecimal(_dt.Rows[iRow]["L_InductanceValue"].ToString().Trim());
							intBAvgLCount++;
						}

						dcmMeasureValue = _dt.Rows[iRow]["C_CapacitanceValue"] == null || _dt.Rows[iRow]["C_CapacitanceValue"].ToString().Trim() == ""
							? 0.000M : Convert.ToDecimal(_dt.Rows[iRow]["C_CapacitanceValue"].ToString().Trim());

						if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "1")
							dcmMeasureValue = dcmMeasureValue - _dcmGroupB1_StopC;
						else if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "2")
							dcmMeasureValue = dcmMeasureValue - _dcmGroupB2_StopC;
						else
							dcmMeasureValue = dcmMeasureValue - _dcmGroupB_StopC;

						if (dcmMeasureValue < 0) dcmMeasureValue = -dcmMeasureValue;

						if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "1")
							dcmChkValue = _dcmGroupB1_StopC * (dcmReferenceValue / 100);
						else if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "2")
							dcmChkValue = _dcmGroupB2_StopC * (dcmReferenceValue / 100);
						else
							dcmChkValue = _dcmGroupB_StopC * (dcmReferenceValue / 100);

						if (dcmMeasureValue <= dcmChkValue)
						{
							dcmBSumStopC += _dt.Rows[iRow]["C_CapacitanceValue"] == null || _dt.Rows[iRow]["C_CapacitanceValue"].ToString().Trim() == "" ? 0.000000M : Convert.ToDecimal(_dt.Rows[iRow]["C_CapacitanceValue"].ToString().Trim());
							intBAvgCCount++;
						}

						dcmMeasureValue = _dt.Rows[iRow]["Q_FactorValue"] == null || _dt.Rows[iRow]["Q_FactorValue"].ToString().Trim() == ""
							? 0.000M : Convert.ToDecimal(_dt.Rows[iRow]["Q_FactorValue"].ToString().Trim());

						if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "1")
							dcmMeasureValue = dcmMeasureValue - _dcmGroupB1_StopQ;
						else if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "2")
							dcmMeasureValue = dcmMeasureValue - _dcmGroupB2_StopQ;
						else
							dcmMeasureValue = dcmMeasureValue - _dcmGroupB_StopQ;

						if (dcmMeasureValue < 0) dcmMeasureValue = -dcmMeasureValue;

						if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "1")
							dcmChkValue = _dcmGroupB1_StopQ * (dcmReferenceValue / 100);
						else if (_dt.Rows[iRow]["SeqNumber"].ToString().Trim() == "2")
							dcmChkValue = _dcmGroupB2_StopQ * (dcmReferenceValue / 100);
						else
							dcmChkValue = _dcmGroupB_StopQ * (dcmReferenceValue / 100);

						if (dcmMeasureValue <= dcmChkValue)
						{
							dcmBSumStopQ += _dt.Rows[iRow]["Q_FactorValue"] == null || _dt.Rows[iRow]["Q_FactorValue"].ToString().Trim() == "" ? 0.000M : Convert.ToDecimal(_dt.Rows[iRow]["Q_FactorValue"].ToString().Trim());
							intBAvgQCount++;
						}
					}

					if (intBAvgRdcCount == 0 && intBAvgRacCount == 0 && intBAvgLCount == 0 && intBAvgCCount == 0 && intBAvgQCount == 0)
						continue;

					// 그룹 B 정지용 평균치
					DataRow dr = dtAverage.NewRow();
					dr[0] = "정지용";
					dr[1] = "B" + Regex.Replace(strCoilName.Trim(), @"\D", "");

					if (intBAvgRdcCount > 0)
						dr[2] = (dcmBSumStopRdc / intBAvgRdcCount).ToString("F3").Trim();
					else
						dr[2] = "0.000";

					if (intBAvgRacCount > 0)
						dr[3] = (dcmBSumStopRac / intBAvgRacCount).ToString("F3").Trim();
					else
						dr[3] = "0.000";

					if (intBAvgLCount > 0)
						dr[4] = (dcmBSumStopL / intBAvgLCount).ToString("F3").Trim();
					else
						dr[4] = "0.000";

					if (intBAvgCCount > 0)
						dr[5] = (dcmBSumStopC / intBAvgCCount).ToString("F6").Trim();
					else
						dr[5] = "0.000";

					if (intBAvgQCount > 0)
						dr[6] = (dcmBSumStopQ / intBAvgQCount).ToString("F3").Trim();
					else
						dr[6] = "0.000";

					dtAverage.Rows.Add(dr);
				}

				decimal dcmB3RdcSum = 0M, dcmB3RacSum = 0M, dcmB3LSum = 0M, dcmB3CSum = 0M, dcmB3QSum = 0M;
				int intB3SCount = 0;

				for (int i = 0; i < dtAverage.Rows.Count; i++)
				{
					if (dtAverage.Rows[i]["CoilCabinetName"].ToString().Trim() == "B1")
					{
						dAveB_StopRdc[0] = dtAverage.Rows[i]["Rdc_DRPIReferenceValue"] == null || dtAverage.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.000M : Convert.ToDecimal(dtAverage.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim());
						dAveB_StopRac[0] = dtAverage.Rows[i]["Rac_DRPIReferenceValue"] == null || dtAverage.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.000M : Convert.ToDecimal(dtAverage.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim());
						dAveB_StopL[0] = dtAverage.Rows[i]["L_DRPIReferenceValue"] == null || dtAverage.Rows[i]["L_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.000M : Convert.ToDecimal(dtAverage.Rows[i]["L_DRPIReferenceValue"].ToString().Trim());
						dAveB_StopC[0] = dtAverage.Rows[i]["C_DRPIReferenceValue"] == null || dtAverage.Rows[i]["C_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.000000M : Convert.ToDecimal(dtAverage.Rows[i]["C_DRPIReferenceValue"].ToString().Trim());
						dAveB_StopQ[0] = dtAverage.Rows[i]["Q_DRPIReferenceValue"] == null || dtAverage.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.0M : Convert.ToDecimal(dtAverage.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim());
					}
					else if (dtAverage.Rows[i]["CoilCabinetName"].ToString().Trim() == "B2")
					{
						dAveB_StopRdc[1] = dtAverage.Rows[i]["Rdc_DRPIReferenceValue"] == null || dtAverage.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.000M : Convert.ToDecimal(dtAverage.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim());
						dAveB_StopRac[1] = dtAverage.Rows[i]["Rac_DRPIReferenceValue"] == null || dtAverage.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.000M : Convert.ToDecimal(dtAverage.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim());
						dAveB_StopL[1] = dtAverage.Rows[i]["L_DRPIReferenceValue"] == null || dtAverage.Rows[i]["L_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.000M : Convert.ToDecimal(dtAverage.Rows[i]["L_DRPIReferenceValue"].ToString().Trim());
						dAveB_StopC[1] = dtAverage.Rows[i]["C_DRPIReferenceValue"] == null || dtAverage.Rows[i]["C_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.000000M : Convert.ToDecimal(dtAverage.Rows[i]["C_DRPIReferenceValue"].ToString().Trim());
						dAveB_StopQ[1] = dtAverage.Rows[i]["Q_DRPIReferenceValue"] == null || dtAverage.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.0M : Convert.ToDecimal(dtAverage.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim());
					}
					else
					{
						dcmB3RdcSum += dtAverage.Rows[i]["Rdc_DRPIReferenceValue"] == null || dtAverage.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.000M : Convert.ToDecimal(dtAverage.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim());
						dcmB3RacSum += dtAverage.Rows[i]["Rac_DRPIReferenceValue"] == null || dtAverage.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.000M : Convert.ToDecimal(dtAverage.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim());
						dcmB3LSum += dtAverage.Rows[i]["L_DRPIReferenceValue"] == null || dtAverage.Rows[i]["L_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.000M : Convert.ToDecimal(dtAverage.Rows[i]["L_DRPIReferenceValue"].ToString().Trim());
						dcmB3CSum += dtAverage.Rows[i]["C_DRPIReferenceValue"] == null || dtAverage.Rows[i]["C_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.000000M : Convert.ToDecimal(dtAverage.Rows[i]["C_DRPIReferenceValue"].ToString().Trim());
						dcmB3QSum += dtAverage.Rows[i]["Q_DRPIReferenceValue"] == null || dtAverage.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim() == ""
							? 0.0M : Convert.ToDecimal(dtAverage.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim());
						intB3SCount++;
					}
				}

				if (intB3SCount > 0)
				{
					dAveB_StopRdc[2] = dcmB3RdcSum / intB3SCount;
					dAveB_StopRac[2] = dcmB3RacSum / intB3SCount;
					dAveB_StopL[2] = dcmB3LSum / intB3SCount;
					dAveB_StopC[2] = dcmB3CSum / intB3SCount;
					dAveB_StopQ[2] = dcmB3QSum / intB3SCount;
				}
				else
				{
					dAveB_StopRdc[2] = 0.000M;
					dAveB_StopRac[2] = 0.000M;
					dAveB_StopL[2] = 0.000M;
					dAveB_StopC[2] = 0.000000M;
					dAveB_StopQ[2] = 0.000M;
				}
			}
			catch (Exception ex)
			{
				System.Diagnostics.Debug.Print(ex.Message);
			}
		}

        /// <summary>
        /// 평균치 기준 값 가져오기
        /// </summary>
        /// <param name="_strDRPIGroup"></param>
        /// <param name="_strDRPIType"></param>
        /// <param name="_strCoilName"></param>
        /// <param name="strColumns"></param>
        /// <returns></returns>
        private decimal GetAverageValue(string _strDRPIGroup, string _strDRPIType, string _strCoilName, string strColumns)
        {
            decimal dcmResultValue = 0.000M;

            switch (_strDRPIGroup)
            {
                case "A":
                    if (_strDRPIType.Trim() == "제어용")
                    {
                        if (strColumns.Trim() == "Rdc" && _strCoilName.Trim() == "A1")
                            dcmResultValue = dAveA_ControlRodRdc[0];
                        else if (strColumns.Trim() == "Rdc" && _strCoilName.Trim() == "A2")
                            dcmResultValue = dAveA_ControlRodRdc[1];
                        else if (strColumns.Trim() == "Rdc")
                            dcmResultValue = dAveA_ControlRodRdc[2];
                        else if (strColumns.Trim() == "Rac" && _strCoilName.Trim() == "A1")
                            dcmResultValue = dAveA_ControlRodRac[0];
                        else if (strColumns.Trim() == "Rac" && _strCoilName.Trim() == "A2")
                            dcmResultValue = dAveA_ControlRodRac[1];
                        else if (strColumns.Trim() == "Rac")
                            dcmResultValue = dAveA_ControlRodRac[2];
                        else if (strColumns.Trim() == "L" && _strCoilName.Trim() == "A1")
                            dcmResultValue = dAveA_ControlRodL[0];
                        else if (strColumns.Trim() == "L" && _strCoilName.Trim() == "A2")
                            dcmResultValue = dAveA_ControlRodL[1];
                        else if (strColumns.Trim() == "L")
                            dcmResultValue = dAveA_ControlRodL[2];
                        else if (strColumns.Trim() == "C" && _strCoilName.Trim() == "A1")
                            dcmResultValue = dAveA_ControlRodC[0];
                        else if (strColumns.Trim() == "C" && _strCoilName.Trim() == "A2")
                            dcmResultValue = dAveA_ControlRodC[1];
                        else if (strColumns.Trim() == "C")
                            dcmResultValue = dAveA_ControlRodC[2];
                        else if (strColumns.Trim() == "Q" && _strCoilName.Trim() == "A1")
                            dcmResultValue = dAveA_ControlRodQ[0];
                        else if (strColumns.Trim() == "Q" && _strCoilName.Trim() == "A2")
                            dcmResultValue = dAveA_ControlRodQ[1];
                        else if (strColumns.Trim() == "Q")
                            dcmResultValue = dAveA_ControlRodQ[2];
                    }
                    else
                    {
                        if (strColumns.Trim() == "Rdc" && _strCoilName.Trim() == "A1")
                            dcmResultValue = dAveA_StopRdc[0];
                        else if (strColumns.Trim() == "Rdc" && _strCoilName.Trim() == "A2")
                            dcmResultValue = dAveA_StopRdc[1];
                        else if (strColumns.Trim() == "Rdc")
                            dcmResultValue = dAveA_StopRdc[2];
                        else if (strColumns.Trim() == "Rac" && _strCoilName.Trim() == "A1")
                            dcmResultValue = dAveA_StopRac[0];
                        else if (strColumns.Trim() == "Rac" && _strCoilName.Trim() == "A2")
                            dcmResultValue = dAveA_StopRac[1];
                        else if (strColumns.Trim() == "Rac")
                            dcmResultValue = dAveA_StopRac[2];
                        else if (strColumns.Trim() == "L" && _strCoilName.Trim() == "A1")
                            dcmResultValue = dAveA_StopL[0];
                        else if (strColumns.Trim() == "L" && _strCoilName.Trim() == "A2")
                            dcmResultValue = dAveA_StopL[1];
                        else if (strColumns.Trim() == "L")
                            dcmResultValue = dAveA_StopL[2];
                        else if (strColumns.Trim() == "C" && _strCoilName.Trim() == "A1")
                            dcmResultValue = dAveA_StopC[0];
                        else if (strColumns.Trim() == "C" && _strCoilName.Trim() == "A2")
                            dcmResultValue = dAveA_StopC[1];
                        else if (strColumns.Trim() == "C")
                            dcmResultValue = dAveA_StopC[2];
                        else if (strColumns.Trim() == "Q" && _strCoilName.Trim() == "A1")
                            dcmResultValue = dAveA_StopQ[0];
                        else if (strColumns.Trim() == "Q" && _strCoilName.Trim() == "A2")
                            dcmResultValue = dAveA_StopQ[1];
                        else if (strColumns.Trim() == "Q")
                            dcmResultValue = dAveA_StopQ[2];
                    }
                    break;
                case "B":
                    if (_strDRPIType.Trim() == "제어용")
                    {
                        if (strColumns.Trim() == "Rdc" && _strCoilName.Trim() == "B1")
                            dcmResultValue = dAveB_ControlRodRdc[0];
                        else if (strColumns.Trim() == "Rdc" && _strCoilName.Trim() == "B2")
                            dcmResultValue = dAveB_ControlRodRdc[1];
                        else if (strColumns.Trim() == "Rdc")
                            dcmResultValue = dAveB_ControlRodRdc[2];
                        else if (strColumns.Trim() == "Rac" && _strCoilName.Trim() == "B1")
                            dcmResultValue = dAveB_ControlRodRac[0];
                        else if (strColumns.Trim() == "Rac" && _strCoilName.Trim() == "B2")
                            dcmResultValue = dAveB_ControlRodRac[1];
                        else if (strColumns.Trim() == "Rac")
                            dcmResultValue = dAveB_ControlRodRac[2];
                        else if (strColumns.Trim() == "L" && _strCoilName.Trim() == "B1")
                            dcmResultValue = dAveB_ControlRodL[0];
                        else if (strColumns.Trim() == "L" && _strCoilName.Trim() == "B2")
                            dcmResultValue = dAveB_ControlRodL[1];
                        else if (strColumns.Trim() == "L")
                            dcmResultValue = dAveB_ControlRodL[2];
                        else if (strColumns.Trim() == "C" && _strCoilName.Trim() == "B1")
                            dcmResultValue = dAveB_ControlRodC[0];
                        else if (strColumns.Trim() == "C" && _strCoilName.Trim() == "B2")
                            dcmResultValue = dAveB_ControlRodC[1];
                        else if (strColumns.Trim() == "C")
                            dcmResultValue = dAveB_ControlRodC[2];
                        else if (strColumns.Trim() == "Q" && _strCoilName.Trim() == "B1")
                            dcmResultValue = dAveB_ControlRodQ[0];
                        else if (strColumns.Trim() == "Q" && _strCoilName.Trim() == "B2")
                            dcmResultValue = dAveB_ControlRodQ[1];
                        else if (strColumns.Trim() == "Q")
                            dcmResultValue = dAveB_ControlRodQ[2];
                    }
                    else
                    {
                        if (strColumns.Trim() == "Rdc" && _strCoilName.Trim() == "B1")
                            dcmResultValue = dAveB_StopRdc[0];
                        else if (strColumns.Trim() == "Rdc" && _strCoilName.Trim() == "B2")
                            dcmResultValue = dAveB_StopRdc[1];
                        else if (strColumns.Trim() == "Rdc")
                            dcmResultValue = dAveB_StopRdc[2];
                        else if (strColumns.Trim() == "Rac" && _strCoilName.Trim() == "B1")
                            dcmResultValue = dAveB_StopRac[0];
                        else if (strColumns.Trim() == "Rac" && _strCoilName.Trim() == "B2")
                            dcmResultValue = dAveB_StopRac[1];
                        else if (strColumns.Trim() == "Rac")
                            dcmResultValue = dAveB_StopRac[2];
                        else if (strColumns.Trim() == "L" && _strCoilName.Trim() == "B1")
                            dcmResultValue = dAveB_StopL[0];
                        else if (strColumns.Trim() == "L" && _strCoilName.Trim() == "B2")
                            dcmResultValue = dAveB_StopL[1];
                        else if (strColumns.Trim() == "L")
                            dcmResultValue = dAveB_StopL[2];
                        else if (strColumns.Trim() == "C" && _strCoilName.Trim() == "B1")
                            dcmResultValue = dAveB_StopC[0];
                        else if (strColumns.Trim() == "C" && _strCoilName.Trim() == "B2")
                            dcmResultValue = dAveB_StopC[1];
                        else if (strColumns.Trim() == "C")
                            dcmResultValue = dAveB_StopC[2];
                        else if (strColumns.Trim() == "Q" && _strCoilName.Trim() == "B1")
                            dcmResultValue = dAveB_StopQ[0];
                        else if (strColumns.Trim() == "Q" && _strCoilName.Trim() == "B2")
                            dcmResultValue = dAveB_StopQ[1];
                        else if (strColumns.Trim() == "Q")
                            dcmResultValue = dAveB_StopQ[2];
                    }
                    break;
            }

            return dcmResultValue;
        }

        /// <summary>
        /// 기준치 기준 값 가져오기
        /// </summary>
        /// <param name="_strDRPIGroup"></param>
        /// <param name="_strDRPIType"></param>
        /// <param name="_strCoilName"></param>
        /// <param name="strColumns"></param>
        /// <returns></returns>
        private decimal GetDRPIReferenceValue(string _strDRPIGroup, string _strDRPIType, string _strCoilName, string strColumns)
        {
            decimal dcmResultValue = 0.000M;

            switch (_strDRPIGroup)
            {
                case "A":
                    if (_strDRPIType.Trim() == "제어용")
                    {
                        if (strColumns.Trim() == "Rdc" && _strCoilName.Trim() == "A1")
                            dcmResultValue = dRefA_ControlRodRdc[0];
                        else if (strColumns.Trim() == "Rdc" && _strCoilName.Trim() == "A2")
                            dcmResultValue = dRefA_ControlRodRdc[1];
                        else if (strColumns.Trim() == "Rdc")
                            dcmResultValue = dRefA_ControlRodRdc[2];
                        else if (strColumns.Trim() == "Rac" && _strCoilName.Trim() == "A1")
                            dcmResultValue = dRefA_ControlRodRac[0];
                        else if (strColumns.Trim() == "Rac" && _strCoilName.Trim() == "A2")
                            dcmResultValue = dRefA_ControlRodRac[1];
                        else if (strColumns.Trim() == "Rac")
                            dcmResultValue = dRefA_ControlRodRac[2];
                        else if (strColumns.Trim() == "L" && _strCoilName.Trim() == "A1")
                            dcmResultValue = dRefA_ControlRodL[0];
                        else if (strColumns.Trim() == "L" && _strCoilName.Trim() == "A2")
                            dcmResultValue = dRefA_ControlRodL[1];
                        else if (strColumns.Trim() == "L")
                            dcmResultValue = dRefA_ControlRodL[2];
                        else if (strColumns.Trim() == "C" && _strCoilName.Trim() == "A1")
                            dcmResultValue = dRefA_ControlRodC[0];
                        else if (strColumns.Trim() == "C" && _strCoilName.Trim() == "A2")
                            dcmResultValue = dRefA_ControlRodC[1];
                        else if (strColumns.Trim() == "C")
                            dcmResultValue = dRefA_ControlRodC[2];
                        else if (strColumns.Trim() == "Q" && _strCoilName.Trim() == "A1")
                            dcmResultValue = dRefA_ControlRodQ[0];
                        else if (strColumns.Trim() == "Q" && _strCoilName.Trim() == "A2")
                            dcmResultValue = dRefA_ControlRodQ[1];
                        else if (strColumns.Trim() == "Q")
                            dcmResultValue = dRefA_ControlRodQ[2];
                    }
                    else
                    {
                        if (strColumns.Trim() == "Rdc" && _strCoilName.Trim() == "A1")
                            dcmResultValue = dRefA_StopRdc[0];
                        else if (strColumns.Trim() == "Rdc" && _strCoilName.Trim() == "A2")
                            dcmResultValue = dRefA_StopRdc[1];
                        else if (strColumns.Trim() == "Rdc")
                            dcmResultValue = dRefA_StopRdc[2];
                        else if (strColumns.Trim() == "Rac" && _strCoilName.Trim() == "A1")
                            dcmResultValue = dRefA_StopRac[0];
                        else if (strColumns.Trim() == "Rac" && _strCoilName.Trim() == "A2")
                            dcmResultValue = dRefA_StopRac[1];
                        else if (strColumns.Trim() == "Rac")
                            dcmResultValue = dRefA_StopRac[2];
                        else if (strColumns.Trim() == "L" && _strCoilName.Trim() == "A1")
                            dcmResultValue = dRefA_StopL[0];
                        else if (strColumns.Trim() == "L" && _strCoilName.Trim() == "A2")
                            dcmResultValue = dRefA_StopL[1];
                        else if (strColumns.Trim() == "L")
                            dcmResultValue = dRefA_StopL[2];
                        else if (strColumns.Trim() == "C" && _strCoilName.Trim() == "A1")
                            dcmResultValue = dRefA_StopC[0];
                        else if (strColumns.Trim() == "C" && _strCoilName.Trim() == "A2")
                            dcmResultValue = dRefA_StopC[1];
                        else if (strColumns.Trim() == "C")
                            dcmResultValue = dRefA_StopC[2];
                        else if (strColumns.Trim() == "Q" && _strCoilName.Trim() == "A1")
                            dcmResultValue = dRefA_StopQ[0];
                        else if (strColumns.Trim() == "Q" && _strCoilName.Trim() == "A2")
                            dcmResultValue = dRefA_StopQ[1];
                        else if (strColumns.Trim() == "Q")
                            dcmResultValue = dRefA_StopQ[2];
                    }
                    break;
                case "B":
                    if (_strDRPIType.Trim() == "제어용")
                    {
                        if (strColumns.Trim() == "Rdc" && _strCoilName.Trim() == "B1")
                            dcmResultValue = dRefB_ControlRodRdc[0];
                        else if (strColumns.Trim() == "Rdc" && _strCoilName.Trim() == "B2")
                            dcmResultValue = dRefB_ControlRodRdc[1];
                        else if (strColumns.Trim() == "Rdc")
                            dcmResultValue = dRefB_ControlRodRdc[2];
                        else if (strColumns.Trim() == "Rac" && _strCoilName.Trim() == "B1")
                            dcmResultValue = dRefB_ControlRodRac[0];
                        else if (strColumns.Trim() == "Rac" && _strCoilName.Trim() == "B2")
                            dcmResultValue = dRefB_ControlRodRac[1];
                        else if (strColumns.Trim() == "Rac")
                            dcmResultValue = dRefB_ControlRodRac[2];
                        else if (strColumns.Trim() == "L" && _strCoilName.Trim() == "B1")
                            dcmResultValue = dRefB_ControlRodL[0];
                        else if (strColumns.Trim() == "L" && _strCoilName.Trim() == "B2")
                            dcmResultValue = dRefB_ControlRodL[1];
                        else if (strColumns.Trim() == "L")
                            dcmResultValue = dRefB_ControlRodL[2];
                        else if (strColumns.Trim() == "C" && _strCoilName.Trim() == "B1")
                            dcmResultValue = dRefB_ControlRodC[0];
                        else if (strColumns.Trim() == "C" && _strCoilName.Trim() == "B2")
                            dcmResultValue = dRefB_ControlRodC[1];
                        else if (strColumns.Trim() == "C")
                            dcmResultValue = dRefB_ControlRodC[2];
                        else if (strColumns.Trim() == "Q" && _strCoilName.Trim() == "B1")
                            dcmResultValue = dRefB_ControlRodQ[0];
                        else if (strColumns.Trim() == "Q" && _strCoilName.Trim() == "B2")
                            dcmResultValue = dRefB_ControlRodQ[1];
                        else if (strColumns.Trim() == "Q")
                            dcmResultValue = dRefB_ControlRodQ[2];
                    }
                    else
                    {
                        if (strColumns.Trim() == "Rdc" && _strCoilName.Trim() == "B1")
                            dcmResultValue = dRefB_StopRdc[0];
                        else if (strColumns.Trim() == "Rdc" && _strCoilName.Trim() == "B2")
                            dcmResultValue = dRefB_StopRdc[1];
                        else if (strColumns.Trim() == "Rdc")
                            dcmResultValue = dRefB_StopRdc[2];
                        else if (strColumns.Trim() == "Rac" && _strCoilName.Trim() == "B1")
                            dcmResultValue = dRefB_StopRac[0];
                        else if (strColumns.Trim() == "Rac" && _strCoilName.Trim() == "B2")
                            dcmResultValue = dRefB_StopRac[1];
                        else if (strColumns.Trim() == "Rac")
                            dcmResultValue = dRefB_StopRac[2];
                        else if (strColumns.Trim() == "L" && _strCoilName.Trim() == "B1")
                            dcmResultValue = dRefB_StopL[0];
                        else if (strColumns.Trim() == "L" && _strCoilName.Trim() == "B2")
                            dcmResultValue = dRefB_StopL[1];
                        else if (strColumns.Trim() == "L")
                            dcmResultValue = dRefB_StopL[2];
                        else if (strColumns.Trim() == "C" && _strCoilName.Trim() == "B1")
                            dcmResultValue = dRefB_StopC[0];
                        else if (strColumns.Trim() == "C" && _strCoilName.Trim() == "B2")
                            dcmResultValue = dRefB_StopC[1];
                        else if (strColumns.Trim() == "C")
                            dcmResultValue = dRefB_StopC[2];
                        else if (strColumns.Trim() == "Q" && _strCoilName.Trim() == "B1")
                            dcmResultValue = dRefB_StopQ[0];
                        else if (strColumns.Trim() == "Q" && _strCoilName.Trim() == "B2")
                            dcmResultValue = dRefB_StopQ[1];
                        else if (strColumns.Trim() == "Q")
                            dcmResultValue = dRefB_StopQ[2];
                    }
                    break;
            }

            return dcmResultValue;
        }

        /// <summary>
        /// 표 형식 / 평균 값(DRPI) 계산
        /// </summary>
        /// <param name="_dt"></param>
        private void SetDRPIReferenceValueTableMark(ref DataTable _dt, string _strComparisonTarget)
        {
            try
            {
                decimal dcmDecisionRange_Rdc = 0M, dcmDecisionRange_Rac = 0M, dcmDecisionRange_L = 0M, dcmDecisionRange_C = 0M
                    , dcmDecisionRange_Q = 0M, dcmEffectiveStandardRange = 0M, dcmDecision = 0M;
                string strRdcResult = "", strRacResult = "", strLResult = "", strCResult = "", strQResult = "";
                int intPASS = 0, intFAIL = 0, intDoubt = 0;

                dcmDecisionRange_Rdc = Convert.ToDecimal(Gini.GetValue("ReferenceValue", "RdcDecisionRange_ReferenceValue"));
                dcmDecisionRange_Rac = Convert.ToDecimal(Gini.GetValue("ReferenceValue", "RacDecisionRange_ReferenceValue"));
                dcmDecisionRange_L = Convert.ToDecimal(Gini.GetValue("ReferenceValue", "LDecisionRange_ReferenceValue"));
                dcmDecisionRange_C = Convert.ToDecimal(Gini.GetValue("ReferenceValue", "CDecisionRange_ReferenceValue"));
                dcmDecisionRange_Q = Convert.ToDecimal(Gini.GetValue("ReferenceValue", "QDecisionRange_ReferenceValue"));
                dcmEffectiveStandardRange = Convert.ToDecimal(Gini.GetValue("ReferenceValue", "EffectiveStandardRangeOfVariation"));

                for (int i = 0; i < _dt.Rows.Count; i++)
                {
                    int iRow = dgvMeasurement.Rows.Add();

                    dgvMeasurement.Rows[iRow].Cells["ControlName"].Value = _dt.Rows[i]["ControlName"].ToString().Trim();
                    dgvMeasurement.Rows[iRow].Cells["CoilName"].Value = _dt.Rows[i]["CoilName"].ToString().Trim();
                    dgvMeasurement.Rows[iRow].Cells["CoilNumber"].Value = _dt.Rows[i]["CoilNumber"].ToString().Trim();
                    dgvMeasurement.Rows[iRow].Cells["DRPIGroup"].Value = _dt.Rows[i]["DRPIGroup"].ToString().Trim();
                    dgvMeasurement.Rows[iRow].Cells["Result"].Value = "";

                    #region Rdc 체크
                    decimal dcmRdcValue = _dt.Rows[i]["DC_ResistanceValue"] == null || _dt.Rows[i]["DC_ResistanceValue"].ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(_dt.Rows[i]["DC_ResistanceValue"].ToString().Trim());
                    dcmDecision = 0.000M;

                    if (_dt.Rows[i]["DC_ResistanceValue"] == null || _dt.Rows[i]["DC_ResistanceValue"].ToString().Trim() == "")
                    {
                        strRdcResult = "부적합";
                    }
                    else
                    {
                        if (_strComparisonTarget.Trim() == "평균값")
                        {
                            if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dAveA_ControlRodRdc[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dAveA_ControlRodRdc[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "A")
                                dcmDecision = dAveA_ControlRodRdc[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dAveA_StopRdc[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dAveA_StopRdc[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "A")
                                dcmDecision = dAveA_StopRdc[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dAveB_ControlRodRdc[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dAveB_ControlRodRdc[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "B")
                                dcmDecision = dAveB_ControlRodRdc[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dAveB_StopRdc[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dAveB_StopRdc[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "B")
                                dcmDecision = dAveB_StopRdc[2];

                            strRdcResult = m_MeasureProcess.GetMesurmentValueDecision(dcmRdcValue, dcmDecision, dcmDecisionRange_Rdc, dcmEffectiveStandardRange, "DC");
                        }
                        else if (_strComparisonTarget.Trim() == "기준값")
                        {
                            if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dRefA_ControlRodRdc[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dRefA_ControlRodRdc[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "A")
                                dcmDecision = dRefA_ControlRodRdc[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dRefA_StopRdc[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dRefA_StopRdc[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "A")
                                dcmDecision = dRefA_StopRdc[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dRefB_ControlRodRdc[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dRefB_ControlRodRdc[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "B")
                                dcmDecision = dRefB_ControlRodRdc[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dRefB_StopRdc[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dRefB_StopRdc[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "B")
                                dcmDecision = dRefB_StopRdc[2];

                            strRdcResult = m_MeasureProcess.GetMesurmentValueDecision(dcmRdcValue, dcmDecision, dcmDecisionRange_Rdc, dcmEffectiveStandardRange, "DC");
                        }
                        else
                            strRdcResult = "부적합";
                    }

                    dgvMeasurement.Rows[iRow].Cells["DC_ResistanceValue"].Value = dcmRdcValue.ToString("F3").Trim();
                    dgvMeasurement.Rows[iRow].Cells["DC_Deviation"].Value = (dcmRdcValue - dcmDecision).ToString("F3").Trim();
                    _dt.Rows[i]["DC_Deviation"] = (dcmRdcValue - dcmDecision).ToString("F3").Trim();

                    if ((dcmRdcValue - dcmDecision) == 0)
                        dgvMeasurement.Rows[iRow].Cells["Q_FactorValue"].Style.ForeColor = System.Drawing.Color.Black;
                    else
                    {
                        if (strRdcResult.Trim() == "부적합")
                        {
                            dgvMeasurement.Rows[iRow].Cells["DC_ResistanceValue"].Style.ForeColor = System.Drawing.Color.Red;
                        }
                        else if (strRdcResult.Trim() == "의심")
                        {
                            dgvMeasurement.Rows[iRow].Cells["DC_ResistanceValue"].Style.ForeColor = System.Drawing.Color.Blue;
                        }
                        else if (strRdcResult.Trim() == "적합")
                        {
                            dgvMeasurement.Rows[iRow].Cells["DC_ResistanceValue"].Style.ForeColor = System.Drawing.Color.Black;
                        }
                    }
                    #endregion

                    #region Rac 체크
                    decimal dcmRacValue = _dt.Rows[i]["AC_ResistanceValue"] == null || _dt.Rows[i]["AC_ResistanceValue"].ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(_dt.Rows[i]["AC_ResistanceValue"].ToString().Trim());
                    dcmDecision = 0.000M;

                    if (_dt.Rows[i]["AC_ResistanceValue"] == null || _dt.Rows[i]["AC_ResistanceValue"].ToString().Trim() == "")
                    {
                        strRacResult = "부적합";
                    }
                    else
                    {
                        if (_strComparisonTarget.Trim() == "평균값")
                        {
                            if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dAveA_ControlRodRac[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dAveA_ControlRodRac[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "A")
                                dcmDecision = dAveA_ControlRodRac[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dAveA_StopRac[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dAveA_StopRac[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "A")
                                dcmDecision = dAveA_StopRac[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dAveB_ControlRodRac[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dAveB_ControlRodRac[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "B")
                                dcmDecision = dAveB_ControlRodRac[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dAveB_StopRac[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dAveB_StopRac[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "B")
                                dcmDecision = dAveB_StopRac[2];

                            strRacResult = m_MeasureProcess.GetMesurmentValueDecision(dcmRacValue, dcmDecision, dcmDecisionRange_Rac, dcmEffectiveStandardRange, "AC");
                        }
                        else if (_strComparisonTarget.Trim() == "기준값")
                        {
                            if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dRefA_ControlRodRac[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dRefA_ControlRodRac[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "A")
                                dcmDecision = dRefA_ControlRodRac[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dRefA_StopRac[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dRefA_StopRac[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "A")
                                dcmDecision = dRefA_StopRac[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dRefB_ControlRodRac[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dRefB_ControlRodRac[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "B")
                                dcmDecision = dRefB_ControlRodRac[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dRefB_StopRac[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dRefB_StopRac[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "B")
                                dcmDecision = dRefB_StopRac[2];

                            strRacResult = m_MeasureProcess.GetMesurmentValueDecision(dcmRacValue, dcmDecision, dcmDecisionRange_Rac, dcmEffectiveStandardRange, "AC");
                        }
                        else
                            strRacResult = "부적합";
                    }

                    dgvMeasurement.Rows[iRow].Cells["AC_ResistanceValue"].Value = dcmRacValue.ToString("F3").Trim();
                    dgvMeasurement.Rows[iRow].Cells["AC_Deviation"].Value = (dcmRacValue - dcmDecision).ToString("F3").Trim();
                    _dt.Rows[i]["AC_Deviation"] = (dcmRacValue - dcmDecision).ToString("F3").Trim();

                    if ((dcmRacValue - dcmDecision) == 0)
                        dgvMeasurement.Rows[iRow].Cells["Q_FactorValue"].Style.ForeColor = System.Drawing.Color.Black;
                    else
                    {
                        if (strRacResult.Trim() == "부적합")
                        {
                            dgvMeasurement.Rows[iRow].Cells["AC_ResistanceValue"].Style.ForeColor = System.Drawing.Color.Red;
                        }
                        else if (strRacResult.Trim() == "의심")
                        {
                            dgvMeasurement.Rows[iRow].Cells["AC_ResistanceValue"].Style.ForeColor = System.Drawing.Color.Blue;
                        }
                        else if (strRacResult.Trim() == "적합")
                        {
                            dgvMeasurement.Rows[iRow].Cells["AC_ResistanceValue"].Style.ForeColor = System.Drawing.Color.Black;
                        }
                    }
                    #endregion

                    #region L 체크
                    decimal dcmLValue = _dt.Rows[i]["L_InductanceValue"] == null || _dt.Rows[i]["L_InductanceValue"].ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(_dt.Rows[i]["L_InductanceValue"].ToString().Trim());
                    dcmDecision = 0.000M;

                    if (_dt.Rows[i]["L_InductanceValue"] == null || _dt.Rows[i]["L_InductanceValue"].ToString().Trim() == "")
                    {
                        strLResult = "부적합";
                    }
                    else
                    {
                        if (_strComparisonTarget.Trim() == "평균값")
                        {
                            if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dAveA_ControlRodL[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dAveA_ControlRodL[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "A")
                                dcmDecision = dAveA_ControlRodL[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dAveA_StopL[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dAveA_StopL[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "A")
                                dcmDecision = dAveA_StopL[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dAveB_ControlRodL[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dAveB_ControlRodL[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "B")
                                dcmDecision = dAveB_ControlRodL[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dAveB_StopL[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dAveB_StopL[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "B")
                                dcmDecision = dAveB_StopL[2];

                            strLResult = m_MeasureProcess.GetMesurmentValueDecision(dcmLValue, dcmDecision, dcmDecisionRange_L, dcmEffectiveStandardRange, "L");
                        }
                        else if (_strComparisonTarget.Trim() == "기준값")
                        {
                            if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dRefA_ControlRodL[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dRefA_ControlRodL[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "A")
                                dcmDecision = dRefA_ControlRodL[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dRefA_StopL[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dRefA_StopL[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "A")
                                dcmDecision = dRefA_StopL[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dRefB_ControlRodL[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dRefB_ControlRodL[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "B")
                                dcmDecision = dRefB_ControlRodL[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dRefB_StopL[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dRefB_StopL[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "B")
                                dcmDecision = dRefB_StopL[2];

                            strLResult = m_MeasureProcess.GetMesurmentValueDecision(dcmLValue, dcmDecision, dcmDecisionRange_L, dcmEffectiveStandardRange, "L");
                        }
                        else
                            strLResult = "부적합";
                    }

                    dgvMeasurement.Rows[iRow].Cells["L_InductanceValue"].Value = dcmLValue.ToString("F3").Trim();
                    dgvMeasurement.Rows[iRow].Cells["L_Deviation"].Value = (dcmLValue - dcmDecision).ToString("F3").Trim();
                    _dt.Rows[i]["L_Deviation"] = (dcmLValue - dcmDecision).ToString("F3").Trim();

                    if ((dcmLValue - dcmDecision) == 0)
                        dgvMeasurement.Rows[iRow].Cells["Q_FactorValue"].Style.ForeColor = System.Drawing.Color.Black;
                    else
                    {
                        if (strLResult.Trim() == "부적합")
                        {
                            dgvMeasurement.Rows[iRow].Cells["L_InductanceValue"].Style.ForeColor = System.Drawing.Color.Red;
                        }
                        else if (strLResult.Trim() == "의심")
                        {
                            dgvMeasurement.Rows[iRow].Cells["L_InductanceValue"].Style.ForeColor = System.Drawing.Color.Blue;
                        }
                        else if (strLResult.Trim() == "적합")
                        {
                            dgvMeasurement.Rows[iRow].Cells["L_InductanceValue"].Style.ForeColor = System.Drawing.Color.Black;
                        }
                    }
                    #endregion

                    #region C 체크
                    decimal dcmCValue = _dt.Rows[i]["C_CapacitanceValue"] == null || _dt.Rows[i]["C_CapacitanceValue"].ToString().Trim() == ""
                        ? 0.000000M : Convert.ToDecimal(_dt.Rows[i]["C_CapacitanceValue"].ToString().Trim());
                    dcmDecision = 0.000000M;

                    if (_dt.Rows[i]["C_CapacitanceValue"] == null || _dt.Rows[i]["C_CapacitanceValue"].ToString().Trim() == "")
                    {
                        strCResult = "부적합";
                    }
                    else
                    {
                        if (_strComparisonTarget.Trim() == "평균값")
                        {
                            if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dAveA_ControlRodC[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dAveA_ControlRodC[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "A")
                                dcmDecision = dAveA_ControlRodC[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dAveA_StopC[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dAveA_StopC[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "A")
                                dcmDecision = dAveA_StopC[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dAveB_ControlRodC[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dAveB_ControlRodC[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "B")
                                dcmDecision = dAveB_ControlRodC[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dAveB_StopC[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dAveB_StopC[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "B")
                                dcmDecision = dAveB_StopC[2];

                            strCResult = m_MeasureProcess.GetMesurmentValueDecision(dcmCValue, dcmDecision, dcmDecisionRange_C, dcmEffectiveStandardRange, "C");
                        }
                        else if (_strComparisonTarget.Trim() == "기준값")
                        {
                            if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dRefA_ControlRodC[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dRefA_ControlRodC[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "A")
                                dcmDecision = dRefA_ControlRodC[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dRefA_StopC[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dRefA_StopC[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "A")
                                dcmDecision = dRefA_StopC[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dRefB_ControlRodC[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dRefB_ControlRodC[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "B")
                                dcmDecision = dRefB_ControlRodC[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dRefB_StopC[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dRefB_StopC[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "B")
                                dcmDecision = dRefB_StopC[2];

                            strCResult = m_MeasureProcess.GetMesurmentValueDecision(dcmCValue, dcmDecision, dcmDecisionRange_C, dcmEffectiveStandardRange, "C");
                        }
                        else
                            strCResult = "부적합";
                    }

                    dgvMeasurement.Rows[iRow].Cells["C_CapacitanceValue"].Value = dcmCValue.ToString("F6").Trim();
                    dgvMeasurement.Rows[iRow].Cells["C_Deviation"].Value = (dcmCValue - dcmDecision).ToString("F6").Trim();
                    _dt.Rows[i]["C_Deviation"] = (dcmCValue - dcmDecision).ToString("F6").Trim();

                    if ((dcmCValue - dcmDecision) == 0)
                        dgvMeasurement.Rows[iRow].Cells["Q_FactorValue"].Style.ForeColor = System.Drawing.Color.Black;
                    else
                    {
                        if (strCResult.Trim() == "부적합")
                        {
                            dgvMeasurement.Rows[iRow].Cells["C_CapacitanceValue"].Style.ForeColor = System.Drawing.Color.Red;
                        }
                        else if (strCResult.Trim() == "의심")
                        {
                            dgvMeasurement.Rows[iRow].Cells["C_CapacitanceValue"].Style.ForeColor = System.Drawing.Color.Blue;
                        }
                        else if (strCResult.Trim() == "적합")
                        {
                            dgvMeasurement.Rows[iRow].Cells["C_CapacitanceValue"].Style.ForeColor = System.Drawing.Color.Black;
                        }
                    }
                    #endregion

                    #region Q 체크
                    decimal dcmQValue = _dt.Rows[i]["Q_FactorValue"] == null || _dt.Rows[i]["Q_FactorValue"].ToString().Trim() == ""
                        ? 0.0M : Convert.ToDecimal(_dt.Rows[i]["Q_FactorValue"].ToString().Trim());
                    dcmDecision = 0.000M;

                    if (_dt.Rows[i]["Q_FactorValue"] == null || _dt.Rows[i]["Q_FactorValue"].ToString().Trim() == "")
                    {
                        strQResult = "부적합";
                    }
                    else
                    {
                        if (_strComparisonTarget.Trim() == "평균값")
                        {
                            if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dAveA_ControlRodQ[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dAveA_ControlRodQ[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "A")
                                dcmDecision = dAveA_ControlRodQ[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dAveA_StopQ[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dAveA_StopQ[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "A")
                                dcmDecision = dAveA_StopQ[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dAveB_ControlRodQ[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dAveB_ControlRodQ[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "B")
                                dcmDecision = dAveB_ControlRodQ[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dAveB_StopQ[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dAveB_StopQ[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "B")
                                dcmDecision = dAveB_StopQ[2];

                            strQResult = m_MeasureProcess.GetMesurmentValueDecision(dcmQValue, dcmDecision, dcmDecisionRange_Q, dcmEffectiveStandardRange, "Q");
                        }
                        else if (_strComparisonTarget.Trim() == "기준값")
                        {
                            if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dRefA_ControlRodQ[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dRefA_ControlRodQ[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "A")
                                dcmDecision = dRefA_ControlRodQ[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dRefA_StopQ[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dRefA_StopQ[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "A")
                                dcmDecision = dRefA_StopQ[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dRefB_ControlRodQ[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dRefB_ControlRodQ[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "B")
                                dcmDecision = dRefB_ControlRodQ[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dRefB_StopQ[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dRefB_StopQ[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "B")
                                dcmDecision = dRefB_StopQ[2];

                            strQResult = m_MeasureProcess.GetMesurmentValueDecision(dcmQValue, dcmDecision, dcmDecisionRange_Q, dcmEffectiveStandardRange, "Q");
                        }
                        else
                            strRdcResult = "부적합";
                    }

                    dgvMeasurement.Rows[iRow].Cells["Q_FactorValue"].Value = dcmQValue.ToString("F3").Trim();
                    dgvMeasurement.Rows[iRow].Cells["Q_Deviation"].Value = (dcmQValue - dcmDecision).ToString("F3").Trim();
                    _dt.Rows[i]["Q_Deviation"] = (dcmQValue - dcmDecision).ToString("F3").Trim();

                    if ((dcmQValue - dcmDecision) == 0)
                        dgvMeasurement.Rows[iRow].Cells["Q_FactorValue"].Style.ForeColor = System.Drawing.Color.Black;
                    else
                    {
                        if (strQResult.Trim() == "부적합")
                        {
                            dgvMeasurement.Rows[iRow].Cells["Q_FactorValue"].Style.ForeColor = System.Drawing.Color.Red;
                        }
                        else if (strQResult.Trim() == "의심")
                        {
                            dgvMeasurement.Rows[iRow].Cells["Q_FactorValue"].Style.ForeColor = System.Drawing.Color.Blue;
                        }
                        else if (strQResult.Trim() == "적합")
                        {
                            dgvMeasurement.Rows[iRow].Cells["Q_FactorValue"].Style.ForeColor = System.Drawing.Color.Black;
                        }
                    }
                    #endregion

                    if (dgvMeasurement.Rows[iRow].Cells["DC_ResistanceValue"].Style.ForeColor == System.Drawing.Color.Red
                        || dgvMeasurement.Rows[iRow].Cells["AC_ResistanceValue"].Style.ForeColor == System.Drawing.Color.Red
                        || dgvMeasurement.Rows[iRow].Cells["L_InductanceValue"].Style.ForeColor == System.Drawing.Color.Red
                        || dgvMeasurement.Rows[iRow].Cells["C_CapacitanceValue"].Style.ForeColor == System.Drawing.Color.Red
                        || dgvMeasurement.Rows[iRow].Cells["Q_FactorValue"].Style.ForeColor == System.Drawing.Color.Red)
                    {
                        intFAIL++;
                        dgvMeasurement.Rows[iRow].Cells["Result"].Value = "부적합";
                        dgvMeasurement.Rows[iRow].Cells["Result"].Style.ForeColor = System.Drawing.Color.Red;
                    }
                    else if (dgvMeasurement.Rows[iRow].Cells["DC_ResistanceValue"].Style.ForeColor == System.Drawing.Color.Blue
                        || dgvMeasurement.Rows[iRow].Cells["AC_ResistanceValue"].Style.ForeColor == System.Drawing.Color.Blue
                        || dgvMeasurement.Rows[iRow].Cells["L_InductanceValue"].Style.ForeColor == System.Drawing.Color.Blue
                        || dgvMeasurement.Rows[iRow].Cells["C_CapacitanceValue"].Style.ForeColor == System.Drawing.Color.Blue
                        || dgvMeasurement.Rows[iRow].Cells["Q_FactorValue"].Style.ForeColor == System.Drawing.Color.Blue)
                    {
                        intDoubt++;
                        dgvMeasurement.Rows[iRow].Cells["Result"].Value = "의심";
                        dgvMeasurement.Rows[iRow].Cells["Result"].Style.ForeColor = System.Drawing.Color.Blue;
                    }
                    else
                    {
                        intPASS++;
                        dgvMeasurement.Rows[iRow].Cells["Result"].Value = "적합";
                        dgvMeasurement.Rows[iRow].Cells["Result"].Style.ForeColor = System.Drawing.Color.Black;
                    }

					if (i == 0)
					{
						strTemperature_ReferenceValue = _dt.Rows[i]["Temperature_ReferenceValue"].ToString().Trim();
						strFrequency = _dt.Rows[i]["Frequency"].ToString().Trim();
						strVoltageLevel = _dt.Rows[i]["VoltageLevel"].ToString().Trim();
						//strMeasurementDate = _dt.Rows[i]["MeasurementDate"].ToString().Trim();
					}

                    dgvMeasurement.Rows[iRow].Cells["Temperature_ReferenceValue"].Value = _dt.Rows[i]["Temperature_ReferenceValue"].ToString().Trim();
                    dgvMeasurement.Rows[iRow].Cells["Frequency"].Value = _dt.Rows[i]["Frequency"].ToString().Trim();
                    dgvMeasurement.Rows[iRow].Cells["VoltageLevel"].Value = _dt.Rows[i]["VoltageLevel"].ToString().Trim();
                }

                if (intFAIL > 0)
                    strMeasurementResult = "부적합";
                else if (intDoubt > 0)
                    strMeasurementResult = "의심";
                else
                    strMeasurementResult = "적합";
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
        }

        /// <summary>
        /// 측정값 및 편차 그래프 그리기
        /// </summary>
        /// <param name="_dt"></param>
        /// <param name="_arrayXLabelValue"></param>
        /// <param name="_intXLabelValueCount"></param>
        public void SetDRPIReferenceValueMeasurement(DataTable _dt, string[] _arrayXLabelValue, int _intXLabelValueCount)
        {
            try
            {
                double[] dA01Rdc = new double[_intXLabelValueCount];
                double[] dA02Rdc = new double[_intXLabelValueCount];
                double[] dA03Rdc = new double[_intXLabelValueCount];
                double[] dA04Rdc = new double[_intXLabelValueCount];
                double[] dA05Rdc = new double[_intXLabelValueCount];
                double[] dA06Rdc = new double[_intXLabelValueCount];
                double[] dA07Rdc = new double[_intXLabelValueCount];
                double[] dA08Rdc = new double[_intXLabelValueCount];
                double[] dA09Rdc = new double[_intXLabelValueCount];
                double[] dA10Rdc = new double[_intXLabelValueCount];
                double[] dA11Rdc = new double[_intXLabelValueCount];
                double[] dA12Rdc = new double[_intXLabelValueCount];
                double[] dA13Rdc = new double[_intXLabelValueCount];
                double[] dA14Rdc = new double[_intXLabelValueCount];
                double[] dA15Rdc = new double[_intXLabelValueCount];
                double[] dA16Rdc = new double[_intXLabelValueCount];
                double[] dA17Rdc = new double[_intXLabelValueCount];
                double[] dA18Rdc = new double[_intXLabelValueCount];
                double[] dA19Rdc = new double[_intXLabelValueCount];
                double[] dA20Rdc = new double[_intXLabelValueCount];
                double[] dA21Rdc = new double[_intXLabelValueCount];

                double[] dA01Rac = new double[_intXLabelValueCount];
                double[] dA02Rac = new double[_intXLabelValueCount];
                double[] dA03Rac = new double[_intXLabelValueCount];
                double[] dA04Rac = new double[_intXLabelValueCount];
                double[] dA05Rac = new double[_intXLabelValueCount];
                double[] dA06Rac = new double[_intXLabelValueCount];
                double[] dA07Rac = new double[_intXLabelValueCount];
                double[] dA08Rac = new double[_intXLabelValueCount];
                double[] dA09Rac = new double[_intXLabelValueCount];
                double[] dA10Rac = new double[_intXLabelValueCount];
                double[] dA11Rac = new double[_intXLabelValueCount];
                double[] dA12Rac = new double[_intXLabelValueCount];
                double[] dA13Rac = new double[_intXLabelValueCount];
                double[] dA14Rac = new double[_intXLabelValueCount];
                double[] dA15Rac = new double[_intXLabelValueCount];
                double[] dA16Rac = new double[_intXLabelValueCount];
                double[] dA17Rac = new double[_intXLabelValueCount];
                double[] dA18Rac = new double[_intXLabelValueCount];
                double[] dA19Rac = new double[_intXLabelValueCount];
                double[] dA20Rac = new double[_intXLabelValueCount];
                double[] dA21Rac = new double[_intXLabelValueCount];

                double[] dA01L = new double[_intXLabelValueCount];
                double[] dA02L = new double[_intXLabelValueCount];
                double[] dA03L = new double[_intXLabelValueCount];
                double[] dA04L = new double[_intXLabelValueCount];
                double[] dA05L = new double[_intXLabelValueCount];
                double[] dA06L = new double[_intXLabelValueCount];
                double[] dA07L = new double[_intXLabelValueCount];
                double[] dA08L = new double[_intXLabelValueCount];
                double[] dA09L = new double[_intXLabelValueCount];
                double[] dA10L = new double[_intXLabelValueCount];
                double[] dA11L = new double[_intXLabelValueCount];
                double[] dA12L = new double[_intXLabelValueCount];
                double[] dA13L = new double[_intXLabelValueCount];
                double[] dA14L = new double[_intXLabelValueCount];
                double[] dA15L = new double[_intXLabelValueCount];
                double[] dA16L = new double[_intXLabelValueCount];
                double[] dA17L = new double[_intXLabelValueCount];
                double[] dA18L = new double[_intXLabelValueCount];
                double[] dA19L = new double[_intXLabelValueCount];
                double[] dA20L = new double[_intXLabelValueCount];
                double[] dA21L = new double[_intXLabelValueCount];

                double[] dA01C = new double[_intXLabelValueCount];
                double[] dA02C = new double[_intXLabelValueCount];
                double[] dA03C = new double[_intXLabelValueCount];
                double[] dA04C = new double[_intXLabelValueCount];
                double[] dA05C = new double[_intXLabelValueCount];
                double[] dA06C = new double[_intXLabelValueCount];
                double[] dA07C = new double[_intXLabelValueCount];
                double[] dA08C = new double[_intXLabelValueCount];
                double[] dA09C = new double[_intXLabelValueCount];
                double[] dA10C = new double[_intXLabelValueCount];
                double[] dA11C = new double[_intXLabelValueCount];
                double[] dA12C = new double[_intXLabelValueCount];
                double[] dA13C = new double[_intXLabelValueCount];
                double[] dA14C = new double[_intXLabelValueCount];
                double[] dA15C = new double[_intXLabelValueCount];
                double[] dA16C = new double[_intXLabelValueCount];
                double[] dA17C = new double[_intXLabelValueCount];
                double[] dA18C = new double[_intXLabelValueCount];
                double[] dA19C = new double[_intXLabelValueCount];
                double[] dA20C = new double[_intXLabelValueCount];
                double[] dA21C = new double[_intXLabelValueCount];

                double[] dA01Q = new double[_intXLabelValueCount];
                double[] dA02Q = new double[_intXLabelValueCount];
                double[] dA03Q = new double[_intXLabelValueCount];
                double[] dA04Q = new double[_intXLabelValueCount];
                double[] dA05Q = new double[_intXLabelValueCount];
                double[] dA06Q = new double[_intXLabelValueCount];
                double[] dA07Q = new double[_intXLabelValueCount];
                double[] dA08Q = new double[_intXLabelValueCount];
                double[] dA09Q = new double[_intXLabelValueCount];
                double[] dA10Q = new double[_intXLabelValueCount];
                double[] dA11Q = new double[_intXLabelValueCount];
                double[] dA12Q = new double[_intXLabelValueCount];
                double[] dA13Q = new double[_intXLabelValueCount];
                double[] dA14Q = new double[_intXLabelValueCount];
                double[] dA15Q = new double[_intXLabelValueCount];
                double[] dA16Q = new double[_intXLabelValueCount];
                double[] dA17Q = new double[_intXLabelValueCount];
                double[] dA18Q = new double[_intXLabelValueCount];
                double[] dA19Q = new double[_intXLabelValueCount];
                double[] dA20Q = new double[_intXLabelValueCount];
                double[] dA21Q = new double[_intXLabelValueCount];

                double[] dDevA01Rdc = new double[_intXLabelValueCount];
                double[] dDevA02Rdc = new double[_intXLabelValueCount];
                double[] dDevA03Rdc = new double[_intXLabelValueCount];
                double[] dDevA04Rdc = new double[_intXLabelValueCount];
                double[] dDevA05Rdc = new double[_intXLabelValueCount];
                double[] dDevA06Rdc = new double[_intXLabelValueCount];
                double[] dDevA07Rdc = new double[_intXLabelValueCount];
                double[] dDevA08Rdc = new double[_intXLabelValueCount];
                double[] dDevA09Rdc = new double[_intXLabelValueCount];
                double[] dDevA10Rdc = new double[_intXLabelValueCount];
                double[] dDevA11Rdc = new double[_intXLabelValueCount];
                double[] dDevA12Rdc = new double[_intXLabelValueCount];
                double[] dDevA13Rdc = new double[_intXLabelValueCount];
                double[] dDevA14Rdc = new double[_intXLabelValueCount];
                double[] dDevA15Rdc = new double[_intXLabelValueCount];
                double[] dDevA16Rdc = new double[_intXLabelValueCount];
                double[] dDevA17Rdc = new double[_intXLabelValueCount];
                double[] dDevA18Rdc = new double[_intXLabelValueCount];
                double[] dDevA19Rdc = new double[_intXLabelValueCount];
                double[] dDevA20Rdc = new double[_intXLabelValueCount];
                double[] dDevA21Rdc = new double[_intXLabelValueCount];

                double[] dDevA01Rac = new double[_intXLabelValueCount];
                double[] dDevA02Rac = new double[_intXLabelValueCount];
                double[] dDevA03Rac = new double[_intXLabelValueCount];
                double[] dDevA04Rac = new double[_intXLabelValueCount];
                double[] dDevA05Rac = new double[_intXLabelValueCount];
                double[] dDevA06Rac = new double[_intXLabelValueCount];
                double[] dDevA07Rac = new double[_intXLabelValueCount];
                double[] dDevA08Rac = new double[_intXLabelValueCount];
                double[] dDevA09Rac = new double[_intXLabelValueCount];
                double[] dDevA10Rac = new double[_intXLabelValueCount];
                double[] dDevA11Rac = new double[_intXLabelValueCount];
                double[] dDevA12Rac = new double[_intXLabelValueCount];
                double[] dDevA13Rac = new double[_intXLabelValueCount];
                double[] dDevA14Rac = new double[_intXLabelValueCount];
                double[] dDevA15Rac = new double[_intXLabelValueCount];
                double[] dDevA16Rac = new double[_intXLabelValueCount];
                double[] dDevA17Rac = new double[_intXLabelValueCount];
                double[] dDevA18Rac = new double[_intXLabelValueCount];
                double[] dDevA19Rac = new double[_intXLabelValueCount];
                double[] dDevA20Rac = new double[_intXLabelValueCount];
                double[] dDevA21Rac = new double[_intXLabelValueCount];

                double[] dDevA01L = new double[_intXLabelValueCount];
                double[] dDevA02L = new double[_intXLabelValueCount];
                double[] dDevA03L = new double[_intXLabelValueCount];
                double[] dDevA04L = new double[_intXLabelValueCount];
                double[] dDevA05L = new double[_intXLabelValueCount];
                double[] dDevA06L = new double[_intXLabelValueCount];
                double[] dDevA07L = new double[_intXLabelValueCount];
                double[] dDevA08L = new double[_intXLabelValueCount];
                double[] dDevA09L = new double[_intXLabelValueCount];
                double[] dDevA10L = new double[_intXLabelValueCount];
                double[] dDevA11L = new double[_intXLabelValueCount];
                double[] dDevA12L = new double[_intXLabelValueCount];
                double[] dDevA13L = new double[_intXLabelValueCount];
                double[] dDevA14L = new double[_intXLabelValueCount];
                double[] dDevA15L = new double[_intXLabelValueCount];
                double[] dDevA16L = new double[_intXLabelValueCount];
                double[] dDevA17L = new double[_intXLabelValueCount];
                double[] dDevA18L = new double[_intXLabelValueCount];
                double[] dDevA19L = new double[_intXLabelValueCount];
                double[] dDevA20L = new double[_intXLabelValueCount];
                double[] dDevA21L = new double[_intXLabelValueCount];

                double[] dDevA01C = new double[_intXLabelValueCount];
                double[] dDevA02C = new double[_intXLabelValueCount];
                double[] dDevA03C = new double[_intXLabelValueCount];
                double[] dDevA04C = new double[_intXLabelValueCount];
                double[] dDevA05C = new double[_intXLabelValueCount];
                double[] dDevA06C = new double[_intXLabelValueCount];
                double[] dDevA07C = new double[_intXLabelValueCount];
                double[] dDevA08C = new double[_intXLabelValueCount];
                double[] dDevA09C = new double[_intXLabelValueCount];
                double[] dDevA10C = new double[_intXLabelValueCount];
                double[] dDevA11C = new double[_intXLabelValueCount];
                double[] dDevA12C = new double[_intXLabelValueCount];
                double[] dDevA13C = new double[_intXLabelValueCount];
                double[] dDevA14C = new double[_intXLabelValueCount];
                double[] dDevA15C = new double[_intXLabelValueCount];
                double[] dDevA16C = new double[_intXLabelValueCount];
                double[] dDevA17C = new double[_intXLabelValueCount];
                double[] dDevA18C = new double[_intXLabelValueCount];
                double[] dDevA19C = new double[_intXLabelValueCount];
                double[] dDevA20C = new double[_intXLabelValueCount];
                double[] dDevA21C = new double[_intXLabelValueCount];

                double[] dDevA01Q = new double[_intXLabelValueCount];
                double[] dDevA02Q = new double[_intXLabelValueCount];
                double[] dDevA03Q = new double[_intXLabelValueCount];
                double[] dDevA04Q = new double[_intXLabelValueCount];
                double[] dDevA05Q = new double[_intXLabelValueCount];
                double[] dDevA06Q = new double[_intXLabelValueCount];
                double[] dDevA07Q = new double[_intXLabelValueCount];
                double[] dDevA08Q = new double[_intXLabelValueCount];
                double[] dDevA09Q = new double[_intXLabelValueCount];
                double[] dDevA10Q = new double[_intXLabelValueCount];
                double[] dDevA11Q = new double[_intXLabelValueCount];
                double[] dDevA12Q = new double[_intXLabelValueCount];
                double[] dDevA13Q = new double[_intXLabelValueCount];
                double[] dDevA14Q = new double[_intXLabelValueCount];
                double[] dDevA15Q = new double[_intXLabelValueCount];
                double[] dDevA16Q = new double[_intXLabelValueCount];
                double[] dDevA17Q = new double[_intXLabelValueCount];
                double[] dDevA18Q = new double[_intXLabelValueCount];
                double[] dDevA19Q = new double[_intXLabelValueCount];
                double[] dDevA20Q = new double[_intXLabelValueCount];
                double[] dDevA21Q = new double[_intXLabelValueCount];

                double dRdcValue = 0d, dRacValue = 0d, dLValue = 0d, dCValue = 0d, dQValue = 0d
                    , dDevRdcValue = 0d, dDevRacValue = 0d, dDevLValue = 0d, dDevCValue = 0d, dDevQValue = 0d;
                int intRdcIndex = 0, intRacIndex = 0, intLIndex = 0, intCIndex = 0, intQIndex = 0;
                int intA1Index = 0, intA2Index = 0, intA3Index = 0, intA4Index = 0, intA5Index = 0, intA6Index = 0, intA7Index = 0
                    , intA8Index = 0, intA9Index = 0, intA10Index = 0, intA11Index = 0, intA12Index = 0, intA13Index = 0, intA14Index = 0
                    , intA15Index = 0, intA16Index = 0, intA17Index = 0, intA18Index = 0, intA19Index = 0, intA20Index = 0, intA21Index = 0;

                string[] arrayItemName = new string[_intXLabelValueCount];
                string strItemName = "", strSelectItemName = "";
                int intItemIndex = 0;

                for (int i = 0; i < _dt.Rows.Count; i++)
                {
                    strSelectItemName = _dt.Rows[i]["ControlName"].ToString().Trim() + "/" + _dt.Rows[i]["DRPIGroup"].ToString().Trim();
                    if (strItemName.Trim() != strSelectItemName)
                    {
                        strItemName = _dt.Rows[i]["ControlName"].ToString().Trim() + "/" + _dt.Rows[i]["DRPIGroup"].ToString().Trim();
                        arrayItemName[intItemIndex] = strItemName.Trim();
                        intItemIndex++;
                    }
                }

                string strSelectControlName = "", strSelectDRPIGroup = "";

                for (int ii = 0; ii < arrayItemName.Length; ii++)
                {
                    if (arrayItemName[ii] == null) continue;

                    string[] arrayData = arrayItemName[ii].Split('/');

                    if (arrayData.Length < 2) continue;

                    strSelectControlName = arrayData[0];
                    strSelectDRPIGroup = arrayData[1];

                    for (int i = 0; i < _dt.Rows.Count; i++)
                    {
                        if (strSelectControlName != _dt.Rows[i]["ControlName"].ToString().Trim()) continue;
                        if (strSelectDRPIGroup != _dt.Rows[i]["DRPIGroup"].ToString().Trim()) continue;

                        if (Regex.Replace(_dt.Rows[i]["CoilName"].ToString().Trim(), @"\D", "") == "1")
                        {
                            dRdcValue = _dt.Rows[i]["DC_ResistanceValue"] == null || _dt.Rows[i]["DC_ResistanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["DC_ResistanceValue"].ToString().Trim());
                            dA01Rdc[ii] = dRdcValue;

                            dRacValue = _dt.Rows[i]["AC_ResistanceValue"] == null || _dt.Rows[i]["AC_ResistanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["AC_ResistanceValue"].ToString().Trim());
                            dA01Rac[ii] = dRacValue;

                            dLValue = _dt.Rows[i]["L_InductanceValue"] == null || _dt.Rows[i]["L_InductanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["L_InductanceValue"].ToString().Trim());
                            dA01L[ii] = dLValue;

                            dCValue = _dt.Rows[i]["C_CapacitanceValue"] == null || _dt.Rows[i]["C_CapacitanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["C_CapacitanceValue"].ToString().Trim());
                            dA01C[ii] = dCValue;

                            dQValue = _dt.Rows[i]["Q_FactorValue"] == null || _dt.Rows[i]["Q_FactorValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["Q_FactorValue"].ToString().Trim());
                            dA01Q[ii] = dQValue;


                            dDevRdcValue = _dt.Rows[i]["DC_Deviation"] == null || _dt.Rows[i]["DC_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["DC_Deviation"].ToString().Trim());
                            dDevA01Rdc[ii] = dDevRdcValue;

                            dDevRacValue = _dt.Rows[i]["AC_Deviation"] == null || _dt.Rows[i]["AC_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["AC_Deviation"].ToString().Trim());
                            dDevA01Rac[ii] = dDevRacValue;

                            dDevLValue = _dt.Rows[i]["L_Deviation"] == null || _dt.Rows[i]["L_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["L_Deviation"].ToString().Trim());
                            dDevA01L[ii] = dDevLValue;

                            dDevCValue = _dt.Rows[i]["C_Deviation"] == null || _dt.Rows[i]["C_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["C_Deviation"].ToString().Trim());
                            dDevA01C[ii] = dDevCValue;

                            dDevQValue = _dt.Rows[i]["Q_Deviation"] == null || _dt.Rows[i]["Q_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["Q_Deviation"].ToString().Trim());
                            dDevA01Q[ii] = dDevQValue;

                            intA1Index++;
                        }
                        else if (Regex.Replace(_dt.Rows[i]["CoilName"].ToString().Trim(), @"\D", "") == "2")
                        {
                            dRdcValue = _dt.Rows[i]["DC_ResistanceValue"] == null || _dt.Rows[i]["DC_ResistanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["DC_ResistanceValue"].ToString().Trim());
                            dA02Rdc[ii] = dRdcValue;

                            dRacValue = _dt.Rows[i]["AC_ResistanceValue"] == null || _dt.Rows[i]["AC_ResistanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["AC_ResistanceValue"].ToString().Trim());
                            dA02Rac[ii] = dRacValue;

                            dLValue = _dt.Rows[i]["L_InductanceValue"] == null || _dt.Rows[i]["L_InductanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["L_InductanceValue"].ToString().Trim());
                            dA02L[ii] = dLValue;

                            dCValue = _dt.Rows[i]["C_CapacitanceValue"] == null || _dt.Rows[i]["C_CapacitanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["C_CapacitanceValue"].ToString().Trim());
                            dA02C[ii] = dCValue;

                            dQValue = _dt.Rows[i]["Q_FactorValue"] == null || _dt.Rows[i]["Q_FactorValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["Q_FactorValue"].ToString().Trim());
                            dA02Q[ii] = dQValue;


                            dDevRdcValue = _dt.Rows[i]["DC_Deviation"] == null || _dt.Rows[i]["DC_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["DC_Deviation"].ToString().Trim());
                            dDevA02Rdc[ii] = dDevRdcValue;

                            dDevRacValue = _dt.Rows[i]["AC_Deviation"] == null || _dt.Rows[i]["AC_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["AC_Deviation"].ToString().Trim());
                            dDevA02Rac[ii] = dDevRacValue;

                            dDevLValue = _dt.Rows[i]["L_Deviation"] == null || _dt.Rows[i]["L_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["L_Deviation"].ToString().Trim());
                            dDevA02L[ii] = dDevLValue;

                            dDevCValue = _dt.Rows[i]["C_Deviation"] == null || _dt.Rows[i]["C_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["C_Deviation"].ToString().Trim());
                            dDevA02C[ii] = dDevCValue;

                            dDevQValue = _dt.Rows[i]["Q_Deviation"] == null || _dt.Rows[i]["Q_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["Q_Deviation"].ToString().Trim());
                            dDevA02Q[ii] = dDevQValue;

                            intA2Index++;
                        }
                        else if (Regex.Replace(_dt.Rows[i]["CoilName"].ToString().Trim(), @"\D", "") == "3")
                        {
                            dRdcValue = _dt.Rows[i]["DC_ResistanceValue"] == null || _dt.Rows[i]["DC_ResistanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["DC_ResistanceValue"].ToString().Trim());
                            dA03Rdc[ii] = dRdcValue;

                            dRacValue = _dt.Rows[i]["AC_ResistanceValue"] == null || _dt.Rows[i]["AC_ResistanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["AC_ResistanceValue"].ToString().Trim());
                            dA03Rac[ii] = dRacValue;

                            dLValue = _dt.Rows[i]["L_InductanceValue"] == null || _dt.Rows[i]["L_InductanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["L_InductanceValue"].ToString().Trim());
                            dA03L[ii] = dLValue;

                            dCValue = _dt.Rows[i]["C_CapacitanceValue"] == null || _dt.Rows[i]["C_CapacitanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["C_CapacitanceValue"].ToString().Trim());
                            dA03C[ii] = dCValue;

                            dQValue = _dt.Rows[i]["Q_FactorValue"] == null || _dt.Rows[i]["Q_FactorValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["Q_FactorValue"].ToString().Trim());
                            dA03Q[ii] = dQValue;


                            dDevRdcValue = _dt.Rows[i]["DC_Deviation"] == null || _dt.Rows[i]["DC_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["DC_Deviation"].ToString().Trim());
                            dDevA03Rdc[ii] = dDevRdcValue;

                            dDevRacValue = _dt.Rows[i]["AC_Deviation"] == null || _dt.Rows[i]["AC_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["AC_Deviation"].ToString().Trim());
                            dDevA03Rac[ii] = dDevRacValue;

                            dDevLValue = _dt.Rows[i]["L_Deviation"] == null || _dt.Rows[i]["L_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["L_Deviation"].ToString().Trim());
                            dDevA03L[ii] = dDevLValue;

                            dDevCValue = _dt.Rows[i]["C_Deviation"] == null || _dt.Rows[i]["C_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["C_Deviation"].ToString().Trim());
                            dDevA03C[ii] = dDevCValue;

                            dDevQValue = _dt.Rows[i]["Q_Deviation"] == null || _dt.Rows[i]["Q_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["Q_Deviation"].ToString().Trim());
                            dDevA03Q[ii] = dDevQValue;

                            intA3Index++;
                        }
                        else if (Regex.Replace(_dt.Rows[i]["CoilName"].ToString().Trim(), @"\D", "") == "4")
                        {
                            dRdcValue = _dt.Rows[i]["DC_ResistanceValue"] == null || _dt.Rows[i]["DC_ResistanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["DC_ResistanceValue"].ToString().Trim());
                            dA04Rdc[ii] = dRdcValue;

                            dRacValue = _dt.Rows[i]["AC_ResistanceValue"] == null || _dt.Rows[i]["AC_ResistanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["AC_ResistanceValue"].ToString().Trim());
                            dA04Rac[ii] = dRacValue;

                            dLValue = _dt.Rows[i]["L_InductanceValue"] == null || _dt.Rows[i]["L_InductanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["L_InductanceValue"].ToString().Trim());
                            dA04L[ii] = dLValue;

                            dCValue = _dt.Rows[i]["C_CapacitanceValue"] == null || _dt.Rows[i]["C_CapacitanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["C_CapacitanceValue"].ToString().Trim());
                            dA04C[ii] = dCValue;

                            dQValue = _dt.Rows[i]["Q_FactorValue"] == null || _dt.Rows[i]["Q_FactorValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["Q_FactorValue"].ToString().Trim());
                            dA04Q[ii] = dQValue;


                            dDevRdcValue = _dt.Rows[i]["DC_Deviation"] == null || _dt.Rows[i]["DC_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["DC_Deviation"].ToString().Trim());
                            dDevA04Rdc[ii] = dDevRdcValue;

                            dDevRacValue = _dt.Rows[i]["AC_Deviation"] == null || _dt.Rows[i]["AC_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["AC_Deviation"].ToString().Trim());
                            dDevA04Rac[ii] = dDevRacValue;

                            dDevLValue = _dt.Rows[i]["L_Deviation"] == null || _dt.Rows[i]["L_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["L_Deviation"].ToString().Trim());
                            dDevA04L[ii] = dDevLValue;

                            dDevCValue = _dt.Rows[i]["C_Deviation"] == null || _dt.Rows[i]["C_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["C_Deviation"].ToString().Trim());
                            dDevA04C[ii] = dDevCValue;

                            dDevQValue = _dt.Rows[i]["Q_Deviation"] == null || _dt.Rows[i]["Q_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["Q_Deviation"].ToString().Trim());
                            dDevA04Q[ii] = dDevQValue;

                            intA4Index++;
                        }
                        else if (Regex.Replace(_dt.Rows[i]["CoilName"].ToString().Trim(), @"\D", "") == "5")
                        {
                            dRdcValue = _dt.Rows[i]["DC_ResistanceValue"] == null || _dt.Rows[i]["DC_ResistanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["DC_ResistanceValue"].ToString().Trim());
                            dA05Rdc[ii] = dRdcValue;

                            dRacValue = _dt.Rows[i]["AC_ResistanceValue"] == null || _dt.Rows[i]["AC_ResistanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["AC_ResistanceValue"].ToString().Trim());
                            dA05Rac[ii] = dRacValue;

                            dLValue = _dt.Rows[i]["L_InductanceValue"] == null || _dt.Rows[i]["L_InductanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["L_InductanceValue"].ToString().Trim());
                            dA05L[ii] = dLValue;

                            dCValue = _dt.Rows[i]["C_CapacitanceValue"] == null || _dt.Rows[i]["C_CapacitanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["C_CapacitanceValue"].ToString().Trim());
                            dA05C[ii] = dCValue;

                            dQValue = _dt.Rows[i]["Q_FactorValue"] == null || _dt.Rows[i]["Q_FactorValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["Q_FactorValue"].ToString().Trim());
                            dA05Q[ii] = dQValue;


                            dDevRdcValue = _dt.Rows[i]["DC_Deviation"] == null || _dt.Rows[i]["DC_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["DC_Deviation"].ToString().Trim());
                            dDevA05Rdc[ii] = dDevRdcValue;

                            dDevRacValue = _dt.Rows[i]["AC_Deviation"] == null || _dt.Rows[i]["AC_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["AC_Deviation"].ToString().Trim());
                            dDevA05Rac[ii] = dDevRacValue;

                            dDevLValue = _dt.Rows[i]["L_Deviation"] == null || _dt.Rows[i]["L_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["L_Deviation"].ToString().Trim());
                            dDevA05L[ii] = dDevLValue;

                            dDevCValue = _dt.Rows[i]["C_Deviation"] == null || _dt.Rows[i]["C_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["C_Deviation"].ToString().Trim());
                            dDevA05C[ii] = dDevCValue;

                            dDevQValue = _dt.Rows[i]["Q_Deviation"] == null || _dt.Rows[i]["Q_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["Q_Deviation"].ToString().Trim());
                            dDevA05Q[ii] = dDevQValue;

                            intA5Index++;
                        }
                        else if (Regex.Replace(_dt.Rows[i]["CoilName"].ToString().Trim(), @"\D", "") == "6")
                        {
                            dRdcValue = _dt.Rows[i]["DC_ResistanceValue"] == null || _dt.Rows[i]["DC_ResistanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["DC_ResistanceValue"].ToString().Trim());
                            dA06Rdc[ii] = dRdcValue;

                            dRacValue = _dt.Rows[i]["AC_ResistanceValue"] == null || _dt.Rows[i]["AC_ResistanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["AC_ResistanceValue"].ToString().Trim());
                            dA06Rac[ii] = dRacValue;

                            dLValue = _dt.Rows[i]["L_InductanceValue"] == null || _dt.Rows[i]["L_InductanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["L_InductanceValue"].ToString().Trim());
                            dA06L[ii] = dLValue;

                            dCValue = _dt.Rows[i]["C_CapacitanceValue"] == null || _dt.Rows[i]["C_CapacitanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["C_CapacitanceValue"].ToString().Trim());
                            dA06C[ii] = dCValue;

                            dQValue = _dt.Rows[i]["Q_FactorValue"] == null || _dt.Rows[i]["Q_FactorValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["Q_FactorValue"].ToString().Trim());
                            dA06Q[ii] = dQValue;


                            dDevRdcValue = _dt.Rows[i]["DC_Deviation"] == null || _dt.Rows[i]["DC_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["DC_Deviation"].ToString().Trim());
                            dDevA06Rdc[ii] = dDevRdcValue;

                            dDevRacValue = _dt.Rows[i]["AC_Deviation"] == null || _dt.Rows[i]["AC_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["AC_Deviation"].ToString().Trim());
                            dDevA06Rac[ii] = dDevRacValue;

                            dDevLValue = _dt.Rows[i]["L_Deviation"] == null || _dt.Rows[i]["L_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["L_Deviation"].ToString().Trim());
                            dDevA06L[ii] = dDevLValue;

                            dDevCValue = _dt.Rows[i]["C_Deviation"] == null || _dt.Rows[i]["C_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["C_Deviation"].ToString().Trim());
                            dDevA06C[ii] = dDevCValue;

                            dDevQValue = _dt.Rows[i]["Q_Deviation"] == null || _dt.Rows[i]["Q_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["Q_Deviation"].ToString().Trim());
                            dDevA06Q[ii] = dDevQValue;

                            intA6Index++;
                        }
                        else if (Regex.Replace(_dt.Rows[i]["CoilName"].ToString().Trim(), @"\D", "") == "7")
                        {
                            dRdcValue = _dt.Rows[i]["DC_ResistanceValue"] == null || _dt.Rows[i]["DC_ResistanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["DC_ResistanceValue"].ToString().Trim());
                            dA07Rdc[ii] = dRdcValue;

                            dRacValue = _dt.Rows[i]["AC_ResistanceValue"] == null || _dt.Rows[i]["AC_ResistanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["AC_ResistanceValue"].ToString().Trim());
                            dA07Rac[ii] = dRacValue;

                            dLValue = _dt.Rows[i]["L_InductanceValue"] == null || _dt.Rows[i]["L_InductanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["L_InductanceValue"].ToString().Trim());
                            dA07L[ii] = dLValue;

                            dCValue = _dt.Rows[i]["C_CapacitanceValue"] == null || _dt.Rows[i]["C_CapacitanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["C_CapacitanceValue"].ToString().Trim());
                            dA07C[ii] = dCValue;

                            dQValue = _dt.Rows[i]["Q_FactorValue"] == null || _dt.Rows[i]["Q_FactorValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["Q_FactorValue"].ToString().Trim());
                            dA07Q[ii] = dQValue;


                            dDevRdcValue = _dt.Rows[i]["DC_Deviation"] == null || _dt.Rows[i]["DC_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["DC_Deviation"].ToString().Trim());
                            dDevA07Rdc[ii] = dDevRdcValue;

                            dDevRacValue = _dt.Rows[i]["AC_Deviation"] == null || _dt.Rows[i]["AC_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["AC_Deviation"].ToString().Trim());
                            dDevA07Rac[ii] = dDevRacValue;

                            dDevLValue = _dt.Rows[i]["L_Deviation"] == null || _dt.Rows[i]["L_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["L_Deviation"].ToString().Trim());
                            dDevA07L[ii] = dDevLValue;

                            dDevCValue = _dt.Rows[i]["C_Deviation"] == null || _dt.Rows[i]["C_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["C_Deviation"].ToString().Trim());
                            dDevA07C[ii] = dDevCValue;

                            dDevQValue = _dt.Rows[i]["Q_Deviation"] == null || _dt.Rows[i]["Q_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["Q_Deviation"].ToString().Trim());
                            dDevA07Q[ii] = dDevQValue;

                            intA7Index++;
                        }
                        else if (Regex.Replace(_dt.Rows[i]["CoilName"].ToString().Trim(), @"\D", "") == "8")
                        {
                            dRdcValue = _dt.Rows[i]["DC_ResistanceValue"] == null || _dt.Rows[i]["DC_ResistanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["DC_ResistanceValue"].ToString().Trim());
                            dA08Rdc[ii] = dRdcValue;

                            dRacValue = _dt.Rows[i]["AC_ResistanceValue"] == null || _dt.Rows[i]["AC_ResistanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["AC_ResistanceValue"].ToString().Trim());
                            dA08Rac[ii] = dRacValue;

                            dLValue = _dt.Rows[i]["L_InductanceValue"] == null || _dt.Rows[i]["L_InductanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["L_InductanceValue"].ToString().Trim());
                            dA08L[ii] = dLValue;

                            dCValue = _dt.Rows[i]["C_CapacitanceValue"] == null || _dt.Rows[i]["C_CapacitanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["C_CapacitanceValue"].ToString().Trim());
                            dA08C[ii] = dCValue;

                            dQValue = _dt.Rows[i]["Q_FactorValue"] == null || _dt.Rows[i]["Q_FactorValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["Q_FactorValue"].ToString().Trim());
                            dA08Q[ii] = dQValue;


                            dDevRdcValue = _dt.Rows[i]["DC_Deviation"] == null || _dt.Rows[i]["DC_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["DC_Deviation"].ToString().Trim());
                            dDevA08Rdc[ii] = dDevRdcValue;

                            dDevRacValue = _dt.Rows[i]["AC_Deviation"] == null || _dt.Rows[i]["AC_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["AC_Deviation"].ToString().Trim());
                            dDevA08Rac[ii] = dDevRacValue;

                            dDevLValue = _dt.Rows[i]["L_Deviation"] == null || _dt.Rows[i]["L_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["L_Deviation"].ToString().Trim());
                            dDevA08L[ii] = dDevLValue;

                            dDevCValue = _dt.Rows[i]["C_Deviation"] == null || _dt.Rows[i]["C_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["C_Deviation"].ToString().Trim());
                            dDevA08C[ii] = dDevCValue;

                            dDevQValue = _dt.Rows[i]["Q_Deviation"] == null || _dt.Rows[i]["Q_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["Q_Deviation"].ToString().Trim());
                            dDevA08Q[ii] = dDevQValue;

                            intA8Index++;
                        }
                        else if (Regex.Replace(_dt.Rows[i]["CoilName"].ToString().Trim(), @"\D", "") == "9")
                        {
                            dRdcValue = _dt.Rows[i]["DC_ResistanceValue"] == null || _dt.Rows[i]["DC_ResistanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["DC_ResistanceValue"].ToString().Trim());
                            dA09Rdc[ii] = dRdcValue;

                            dRacValue = _dt.Rows[i]["AC_ResistanceValue"] == null || _dt.Rows[i]["AC_ResistanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["AC_ResistanceValue"].ToString().Trim());
                            dA09Rac[ii] = dRacValue;

                            dLValue = _dt.Rows[i]["L_InductanceValue"] == null || _dt.Rows[i]["L_InductanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["L_InductanceValue"].ToString().Trim());
                            dA09L[ii] = dLValue;

                            dCValue = _dt.Rows[i]["C_CapacitanceValue"] == null || _dt.Rows[i]["C_CapacitanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["C_CapacitanceValue"].ToString().Trim());
                            dA09C[ii] = dCValue;

                            dQValue = _dt.Rows[i]["Q_FactorValue"] == null || _dt.Rows[i]["Q_FactorValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["Q_FactorValue"].ToString().Trim());
                            dA09Q[ii] = dQValue;


                            dDevRdcValue = _dt.Rows[i]["DC_Deviation"] == null || _dt.Rows[i]["DC_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["DC_Deviation"].ToString().Trim());
                            dDevA09Rdc[ii] = dDevRdcValue;

                            dDevRacValue = _dt.Rows[i]["AC_Deviation"] == null || _dt.Rows[i]["AC_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["AC_Deviation"].ToString().Trim());
                            dDevA09Rac[ii] = dDevRacValue;

                            dDevLValue = _dt.Rows[i]["L_Deviation"] == null || _dt.Rows[i]["L_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["L_Deviation"].ToString().Trim());
                            dDevA09L[ii] = dDevLValue;

                            dDevCValue = _dt.Rows[i]["C_Deviation"] == null || _dt.Rows[i]["C_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["C_Deviation"].ToString().Trim());
                            dDevA09C[ii] = dDevCValue;

                            dDevQValue = _dt.Rows[i]["Q_Deviation"] == null || _dt.Rows[i]["Q_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["Q_Deviation"].ToString().Trim());
                            dDevA09Q[ii] = dDevQValue;

                            intA9Index++;
                        }
                        else if (Regex.Replace(_dt.Rows[i]["CoilName"].ToString().Trim(), @"\D", "") == "10")
                        {
                            dRdcValue = _dt.Rows[i]["DC_ResistanceValue"] == null || _dt.Rows[i]["DC_ResistanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["DC_ResistanceValue"].ToString().Trim());
                            dA10Rdc[ii] = dRdcValue;

                            dRacValue = _dt.Rows[i]["AC_ResistanceValue"] == null || _dt.Rows[i]["AC_ResistanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["AC_ResistanceValue"].ToString().Trim());
                            dA10Rac[ii] = dRacValue;

                            dLValue = _dt.Rows[i]["L_InductanceValue"] == null || _dt.Rows[i]["L_InductanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["L_InductanceValue"].ToString().Trim());
                            dA10L[ii] = dLValue;

                            dCValue = _dt.Rows[i]["C_CapacitanceValue"] == null || _dt.Rows[i]["C_CapacitanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["C_CapacitanceValue"].ToString().Trim());
                            dA10C[ii] = dCValue;

                            dQValue = _dt.Rows[i]["Q_FactorValue"] == null || _dt.Rows[i]["Q_FactorValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["Q_FactorValue"].ToString().Trim());
                            dA10Q[ii] = dQValue;


                            dDevRdcValue = _dt.Rows[i]["DC_Deviation"] == null || _dt.Rows[i]["DC_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["DC_Deviation"].ToString().Trim());
                            dDevA10Rdc[ii] = dDevRdcValue;

                            dDevRacValue = _dt.Rows[i]["AC_Deviation"] == null || _dt.Rows[i]["AC_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["AC_Deviation"].ToString().Trim());
                            dDevA10Rac[ii] = dDevRacValue;

                            dDevLValue = _dt.Rows[i]["L_Deviation"] == null || _dt.Rows[i]["L_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["L_Deviation"].ToString().Trim());
                            dDevA10L[ii] = dDevLValue;

                            dDevCValue = _dt.Rows[i]["C_Deviation"] == null || _dt.Rows[i]["C_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["C_Deviation"].ToString().Trim());
                            dDevA10C[ii] = dDevCValue;

                            dDevQValue = _dt.Rows[i]["Q_Deviation"] == null || _dt.Rows[i]["Q_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["Q_Deviation"].ToString().Trim());
                            dDevA10Q[ii] = dDevQValue;

                            intA10Index++;
                        }
                        else if (Regex.Replace(_dt.Rows[i]["CoilName"].ToString().Trim(), @"\D", "") == "11")
                        {
                            dRdcValue = _dt.Rows[i]["DC_ResistanceValue"] == null || _dt.Rows[i]["DC_ResistanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["DC_ResistanceValue"].ToString().Trim());
                            dA11Rdc[ii] = dRdcValue;

                            dRacValue = _dt.Rows[i]["AC_ResistanceValue"] == null || _dt.Rows[i]["AC_ResistanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["AC_ResistanceValue"].ToString().Trim());
                            dA11Rac[ii] = dRacValue;

                            dLValue = _dt.Rows[i]["L_InductanceValue"] == null || _dt.Rows[i]["L_InductanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["L_InductanceValue"].ToString().Trim());
                            dA11L[ii] = dLValue;

                            dCValue = _dt.Rows[i]["C_CapacitanceValue"] == null || _dt.Rows[i]["C_CapacitanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["C_CapacitanceValue"].ToString().Trim());
                            dA11C[ii] = dCValue;

                            dQValue = _dt.Rows[i]["Q_FactorValue"] == null || _dt.Rows[i]["Q_FactorValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["Q_FactorValue"].ToString().Trim());
                            dA11Q[ii] = dQValue;


                            dDevRdcValue = _dt.Rows[i]["DC_Deviation"] == null || _dt.Rows[i]["DC_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["DC_Deviation"].ToString().Trim());
                            dDevA11Rdc[ii] = dDevRdcValue;

                            dDevRacValue = _dt.Rows[i]["AC_Deviation"] == null || _dt.Rows[i]["AC_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["AC_Deviation"].ToString().Trim());
                            dDevA11Rac[ii] = dDevRacValue;

                            dDevLValue = _dt.Rows[i]["L_Deviation"] == null || _dt.Rows[i]["L_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["L_Deviation"].ToString().Trim());
                            dDevA11L[ii] = dDevLValue;

                            dDevCValue = _dt.Rows[i]["C_Deviation"] == null || _dt.Rows[i]["C_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["C_Deviation"].ToString().Trim());
                            dDevA11C[ii] = dDevCValue;

                            dDevQValue = _dt.Rows[i]["Q_Deviation"] == null || _dt.Rows[i]["Q_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["Q_Deviation"].ToString().Trim());
                            dDevA11Q[ii] = dDevQValue;

                            intA11Index++;
                        }
                        else if (Regex.Replace(_dt.Rows[i]["CoilName"].ToString().Trim(), @"\D", "") == "12")
                        {
                            dRdcValue = _dt.Rows[i]["DC_ResistanceValue"] == null || _dt.Rows[i]["DC_ResistanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["DC_ResistanceValue"].ToString().Trim());
                            dA12Rdc[ii] = dRdcValue;

                            dRacValue = _dt.Rows[i]["AC_ResistanceValue"] == null || _dt.Rows[i]["AC_ResistanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["AC_ResistanceValue"].ToString().Trim());
                            dA12Rac[ii] = dRacValue;

                            dLValue = _dt.Rows[i]["L_InductanceValue"] == null || _dt.Rows[i]["L_InductanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["L_InductanceValue"].ToString().Trim());
                            dA12L[ii] = dLValue;

                            dCValue = _dt.Rows[i]["C_CapacitanceValue"] == null || _dt.Rows[i]["C_CapacitanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["C_CapacitanceValue"].ToString().Trim());
                            dA12C[ii] = dCValue;

                            dQValue = _dt.Rows[i]["Q_FactorValue"] == null || _dt.Rows[i]["Q_FactorValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["Q_FactorValue"].ToString().Trim());
                            dA12Q[ii] = dQValue;


                            dDevRdcValue = _dt.Rows[i]["DC_Deviation"] == null || _dt.Rows[i]["DC_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["DC_Deviation"].ToString().Trim());
                            dDevA12Rdc[ii] = dDevRdcValue;

                            dDevRacValue = _dt.Rows[i]["AC_Deviation"] == null || _dt.Rows[i]["AC_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["AC_Deviation"].ToString().Trim());
                            dDevA12Rac[ii] = dDevRacValue;

                            dDevLValue = _dt.Rows[i]["L_Deviation"] == null || _dt.Rows[i]["L_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["L_Deviation"].ToString().Trim());
                            dDevA12L[ii] = dDevLValue;

                            dDevCValue = _dt.Rows[i]["C_Deviation"] == null || _dt.Rows[i]["C_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["C_Deviation"].ToString().Trim());
                            dDevA12C[ii] = dDevCValue;

                            dDevQValue = _dt.Rows[i]["Q_Deviation"] == null || _dt.Rows[i]["Q_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["Q_Deviation"].ToString().Trim());
                            dDevA12Q[ii] = dDevQValue;

                            intA12Index++;
                        }
                        else if (Regex.Replace(_dt.Rows[i]["CoilName"].ToString().Trim(), @"\D", "") == "13")
                        {
                            dRdcValue = _dt.Rows[i]["DC_ResistanceValue"] == null || _dt.Rows[i]["DC_ResistanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["DC_ResistanceValue"].ToString().Trim());
                            dA13Rdc[ii] = dRdcValue;

                            dRacValue = _dt.Rows[i]["AC_ResistanceValue"] == null || _dt.Rows[i]["AC_ResistanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["AC_ResistanceValue"].ToString().Trim());
                            dA13Rac[ii] = dRacValue;

                            dLValue = _dt.Rows[i]["L_InductanceValue"] == null || _dt.Rows[i]["L_InductanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["L_InductanceValue"].ToString().Trim());
                            dA13L[ii] = dLValue;

                            dCValue = _dt.Rows[i]["C_CapacitanceValue"] == null || _dt.Rows[i]["C_CapacitanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["C_CapacitanceValue"].ToString().Trim());
                            dA13C[ii] = dCValue;

                            dQValue = _dt.Rows[i]["Q_FactorValue"] == null || _dt.Rows[i]["Q_FactorValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["Q_FactorValue"].ToString().Trim());
                            dA13Q[ii] = dQValue;


                            dDevRdcValue = _dt.Rows[i]["DC_Deviation"] == null || _dt.Rows[i]["DC_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["DC_Deviation"].ToString().Trim());
                            dDevA13Rdc[ii] = dDevRdcValue;

                            dDevRacValue = _dt.Rows[i]["AC_Deviation"] == null || _dt.Rows[i]["AC_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["AC_Deviation"].ToString().Trim());
                            dDevA13Rac[ii] = dDevRacValue;

                            dDevLValue = _dt.Rows[i]["L_Deviation"] == null || _dt.Rows[i]["L_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["L_Deviation"].ToString().Trim());
                            dDevA13L[ii] = dDevLValue;

                            dDevCValue = _dt.Rows[i]["C_Deviation"] == null || _dt.Rows[i]["C_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["C_Deviation"].ToString().Trim());
                            dDevA13C[ii] = dDevCValue;

                            dDevQValue = _dt.Rows[i]["Q_Deviation"] == null || _dt.Rows[i]["Q_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["Q_Deviation"].ToString().Trim());
                            dDevA13Q[ii] = dDevQValue;

                            intA13Index++;
                        }
                        else if (Regex.Replace(_dt.Rows[i]["CoilName"].ToString().Trim(), @"\D", "") == "14")
                        {
                            dRdcValue = _dt.Rows[i]["DC_ResistanceValue"] == null || _dt.Rows[i]["DC_ResistanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["DC_ResistanceValue"].ToString().Trim());
                            dA14Rdc[ii] = dRdcValue;

                            dRacValue = _dt.Rows[i]["AC_ResistanceValue"] == null || _dt.Rows[i]["AC_ResistanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["AC_ResistanceValue"].ToString().Trim());
                            dA14Rac[ii] = dRacValue;

                            dLValue = _dt.Rows[i]["L_InductanceValue"] == null || _dt.Rows[i]["L_InductanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["L_InductanceValue"].ToString().Trim());
                            dA14L[ii] = dLValue;

                            dCValue = _dt.Rows[i]["C_CapacitanceValue"] == null || _dt.Rows[i]["C_CapacitanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["C_CapacitanceValue"].ToString().Trim());
                            dA14C[ii] = dCValue;

                            dQValue = _dt.Rows[i]["Q_FactorValue"] == null || _dt.Rows[i]["Q_FactorValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["Q_FactorValue"].ToString().Trim());
                            dA14Q[ii] = dQValue;


                            dDevRdcValue = _dt.Rows[i]["DC_Deviation"] == null || _dt.Rows[i]["DC_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["DC_Deviation"].ToString().Trim());
                            dDevA14Rdc[ii] = dDevRdcValue;

                            dDevRacValue = _dt.Rows[i]["AC_Deviation"] == null || _dt.Rows[i]["AC_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["AC_Deviation"].ToString().Trim());
                            dDevA14Rac[ii] = dDevRacValue;

                            dDevLValue = _dt.Rows[i]["L_Deviation"] == null || _dt.Rows[i]["L_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["L_Deviation"].ToString().Trim());
                            dDevA14L[ii] = dDevLValue;

                            dDevCValue = _dt.Rows[i]["C_Deviation"] == null || _dt.Rows[i]["C_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["C_Deviation"].ToString().Trim());
                            dDevA14C[ii] = dDevCValue;

                            dDevQValue = _dt.Rows[i]["Q_Deviation"] == null || _dt.Rows[i]["Q_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["Q_Deviation"].ToString().Trim());
                            dDevA14Q[ii] = dDevQValue;

                            intA14Index++;
                        }
                        else if (Regex.Replace(_dt.Rows[i]["CoilName"].ToString().Trim(), @"\D", "") == "15")
                        {
                            dRdcValue = _dt.Rows[i]["DC_ResistanceValue"] == null || _dt.Rows[i]["DC_ResistanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["DC_ResistanceValue"].ToString().Trim());
                            dA15Rdc[ii] = dRdcValue;

                            dRacValue = _dt.Rows[i]["AC_ResistanceValue"] == null || _dt.Rows[i]["AC_ResistanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["AC_ResistanceValue"].ToString().Trim());
                            dA15Rac[ii] = dRacValue;

                            dLValue = _dt.Rows[i]["L_InductanceValue"] == null || _dt.Rows[i]["L_InductanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["L_InductanceValue"].ToString().Trim());
                            dA15L[ii] = dLValue;

                            dCValue = _dt.Rows[i]["C_CapacitanceValue"] == null || _dt.Rows[i]["C_CapacitanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["C_CapacitanceValue"].ToString().Trim());
                            dA15C[ii] = dCValue;

                            dQValue = _dt.Rows[i]["Q_FactorValue"] == null || _dt.Rows[i]["Q_FactorValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["Q_FactorValue"].ToString().Trim());
                            dA15Q[ii] = dQValue;


                            dDevRdcValue = _dt.Rows[i]["DC_Deviation"] == null || _dt.Rows[i]["DC_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["DC_Deviation"].ToString().Trim());
                            dDevA15Rdc[ii] = dDevRdcValue;

                            dDevRacValue = _dt.Rows[i]["AC_Deviation"] == null || _dt.Rows[i]["AC_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["AC_Deviation"].ToString().Trim());
                            dDevA15Rac[ii] = dDevRacValue;

                            dDevLValue = _dt.Rows[i]["L_Deviation"] == null || _dt.Rows[i]["L_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["L_Deviation"].ToString().Trim());
                            dDevA15L[ii] = dDevLValue;

                            dDevCValue = _dt.Rows[i]["C_Deviation"] == null || _dt.Rows[i]["C_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["C_Deviation"].ToString().Trim());
                            dDevA15C[ii] = dDevCValue;

                            dDevQValue = _dt.Rows[i]["Q_Deviation"] == null || _dt.Rows[i]["Q_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["Q_Deviation"].ToString().Trim());
                            dDevA15Q[ii] = dDevQValue;

                            intA15Index++;
                        }
                        else if (Regex.Replace(_dt.Rows[i]["CoilName"].ToString().Trim(), @"\D", "") == "16")
                        {
                            dRdcValue = _dt.Rows[i]["DC_ResistanceValue"] == null || _dt.Rows[i]["DC_ResistanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["DC_ResistanceValue"].ToString().Trim());
                            dA16Rdc[ii] = dRdcValue;

                            dRacValue = _dt.Rows[i]["AC_ResistanceValue"] == null || _dt.Rows[i]["AC_ResistanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["AC_ResistanceValue"].ToString().Trim());
                            dA16Rac[ii] = dRacValue;

                            dLValue = _dt.Rows[i]["L_InductanceValue"] == null || _dt.Rows[i]["L_InductanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["L_InductanceValue"].ToString().Trim());
                            dA16L[ii] = dLValue;

                            dCValue = _dt.Rows[i]["C_CapacitanceValue"] == null || _dt.Rows[i]["C_CapacitanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["C_CapacitanceValue"].ToString().Trim());
                            dA16C[ii] = dCValue;

                            dQValue = _dt.Rows[i]["Q_FactorValue"] == null || _dt.Rows[i]["Q_FactorValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["Q_FactorValue"].ToString().Trim());
                            dA16Q[ii] = dQValue;


                            dDevRdcValue = _dt.Rows[i]["DC_Deviation"] == null || _dt.Rows[i]["DC_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["DC_Deviation"].ToString().Trim());
                            dDevA16Rdc[ii] = dDevRdcValue;

                            dDevRacValue = _dt.Rows[i]["AC_Deviation"] == null || _dt.Rows[i]["AC_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["AC_Deviation"].ToString().Trim());
                            dDevA16Rac[ii] = dDevRacValue;

                            dDevLValue = _dt.Rows[i]["L_Deviation"] == null || _dt.Rows[i]["L_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["L_Deviation"].ToString().Trim());
                            dDevA16L[ii] = dDevLValue;

                            dDevCValue = _dt.Rows[i]["C_Deviation"] == null || _dt.Rows[i]["C_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["C_Deviation"].ToString().Trim());
                            dDevA16C[ii] = dDevCValue;

                            dDevQValue = _dt.Rows[i]["Q_Deviation"] == null || _dt.Rows[i]["Q_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["Q_Deviation"].ToString().Trim());
                            dDevA16Q[ii] = dDevQValue;

                            intA16Index++;
                        }
                        else if (Regex.Replace(_dt.Rows[i]["CoilName"].ToString().Trim(), @"\D", "") == "17")
                        {
                            dRdcValue = _dt.Rows[i]["DC_ResistanceValue"] == null || _dt.Rows[i]["DC_ResistanceValue"].ToString().Trim() == ""
                            ? 0.000d : Convert.ToDouble(_dt.Rows[i]["DC_ResistanceValue"].ToString().Trim());
                            dA17Rdc[ii] = dRdcValue;

                            dRacValue = _dt.Rows[i]["AC_ResistanceValue"] == null || _dt.Rows[i]["AC_ResistanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["AC_ResistanceValue"].ToString().Trim());
                            dA17Rac[ii] = dRacValue;

                            dLValue = _dt.Rows[i]["L_InductanceValue"] == null || _dt.Rows[i]["L_InductanceValue"].ToString().Trim() == ""
                            ? 0.000d : Convert.ToDouble(_dt.Rows[i]["L_InductanceValue"].ToString().Trim());
                            dA17L[ii] = dLValue;

                            dCValue = _dt.Rows[i]["C_CapacitanceValue"] == null || _dt.Rows[i]["C_CapacitanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["C_CapacitanceValue"].ToString().Trim());
                            dA17C[ii] = dCValue;

                            dQValue = _dt.Rows[i]["Q_FactorValue"] == null || _dt.Rows[i]["Q_FactorValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["Q_FactorValue"].ToString().Trim());
                            dA17Q[ii] = dQValue;


                            dDevRdcValue = _dt.Rows[i]["DC_Deviation"] == null || _dt.Rows[i]["DC_Deviation"].ToString().Trim() == ""
                            ? 0.000d : Convert.ToDouble(_dt.Rows[i]["DC_Deviation"].ToString().Trim());
                            dDevA17Rdc[ii] = dDevRdcValue;

                            dDevRacValue = _dt.Rows[i]["AC_Deviation"] == null || _dt.Rows[i]["AC_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["AC_Deviation"].ToString().Trim());
                            dDevA17Rac[ii] = dDevRacValue;

                            dDevLValue = _dt.Rows[i]["L_Deviation"] == null || _dt.Rows[i]["L_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["L_Deviation"].ToString().Trim());
                            dDevA17L[ii] = dDevLValue;

                            dDevCValue = _dt.Rows[i]["C_Deviation"] == null || _dt.Rows[i]["C_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["C_Deviation"].ToString().Trim());
                            dDevA17C[ii] = dDevCValue;

                            dDevQValue = _dt.Rows[i]["Q_Deviation"] == null || _dt.Rows[i]["Q_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["Q_Deviation"].ToString().Trim());
                            dDevA17Q[ii] = dDevQValue;

                            intA17Index++;
                        }
                        else if (Regex.Replace(_dt.Rows[i]["CoilName"].ToString().Trim(), @"\D", "") == "18")
                        {
                            dRdcValue = _dt.Rows[i]["DC_ResistanceValue"] == null || _dt.Rows[i]["DC_ResistanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["DC_ResistanceValue"].ToString().Trim());
                            dA18Rdc[ii] = dRdcValue;

                            dRacValue = _dt.Rows[i]["AC_ResistanceValue"] == null || _dt.Rows[i]["AC_ResistanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["AC_ResistanceValue"].ToString().Trim());
                            dA18Rac[ii] = dRacValue;

                            dLValue = _dt.Rows[i]["L_InductanceValue"] == null || _dt.Rows[i]["L_InductanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["L_InductanceValue"].ToString().Trim());
                            dA18L[ii] = dLValue;

                            dCValue = _dt.Rows[i]["C_CapacitanceValue"] == null || _dt.Rows[i]["C_CapacitanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["C_CapacitanceValue"].ToString().Trim());
                            dA18C[ii] = dCValue;

                            dQValue = _dt.Rows[i]["Q_FactorValue"] == null || _dt.Rows[i]["Q_FactorValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["Q_FactorValue"].ToString().Trim());
                            dA18Q[ii] = dQValue;


                            dDevRdcValue = _dt.Rows[i]["DC_Deviation"] == null || _dt.Rows[i]["DC_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["DC_Deviation"].ToString().Trim());
                            dDevA18Rdc[ii] = dDevRdcValue;

                            dDevRacValue = _dt.Rows[i]["AC_Deviation"] == null || _dt.Rows[i]["AC_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["AC_Deviation"].ToString().Trim());
                            dDevA18Rac[ii] = dDevRacValue;

                            dDevLValue = _dt.Rows[i]["L_Deviation"] == null || _dt.Rows[i]["L_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["L_Deviation"].ToString().Trim());
                            dDevA18L[ii] = dDevLValue;

                            dDevCValue = _dt.Rows[i]["C_Deviation"] == null || _dt.Rows[i]["C_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["C_Deviation"].ToString().Trim());
                            dDevA18C[ii] = dDevCValue;

                            dDevQValue = _dt.Rows[i]["Q_Deviation"] == null || _dt.Rows[i]["Q_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["Q_Deviation"].ToString().Trim());
                            dDevA18Q[ii] = dDevQValue;

                            intA18Index++;
                        }
                        else if (Regex.Replace(_dt.Rows[i]["CoilName"].ToString().Trim(), @"\D", "") == "19")
                        {
                            dRdcValue = _dt.Rows[i]["DC_ResistanceValue"] == null || _dt.Rows[i]["DC_ResistanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["DC_ResistanceValue"].ToString().Trim());
                            dA19Rdc[ii] = dRdcValue;

                            dRacValue = _dt.Rows[i]["AC_ResistanceValue"] == null || _dt.Rows[i]["AC_ResistanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["AC_ResistanceValue"].ToString().Trim());
                            dA19Rac[ii] = dRacValue;

                            dLValue = _dt.Rows[i]["L_InductanceValue"] == null || _dt.Rows[i]["L_InductanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["L_InductanceValue"].ToString().Trim());
                            dA19L[ii] = dLValue;

                            dCValue = _dt.Rows[i]["C_CapacitanceValue"] == null || _dt.Rows[i]["C_CapacitanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["C_CapacitanceValue"].ToString().Trim());
                            dA19C[ii] = dCValue;

                            dQValue = _dt.Rows[i]["Q_FactorValue"] == null || _dt.Rows[i]["Q_FactorValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["Q_FactorValue"].ToString().Trim());
                            dA19Q[ii] = dQValue;


                            dDevRdcValue = _dt.Rows[i]["DC_Deviation"] == null || _dt.Rows[i]["DC_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["DC_Deviation"].ToString().Trim());
                            dDevA19Rdc[ii] = dDevRdcValue;

                            dDevRacValue = _dt.Rows[i]["AC_Deviation"] == null || _dt.Rows[i]["AC_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["AC_Deviation"].ToString().Trim());
                            dDevA19Rac[ii] = dDevRacValue;

                            dDevLValue = _dt.Rows[i]["L_Deviation"] == null || _dt.Rows[i]["L_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["L_Deviation"].ToString().Trim());
                            dDevA19L[ii] = dDevLValue;

                            dDevCValue = _dt.Rows[i]["C_Deviation"] == null || _dt.Rows[i]["C_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["C_Deviation"].ToString().Trim());
                            dDevA19C[ii] = dDevCValue;

                            dDevQValue = _dt.Rows[i]["Q_Deviation"] == null || _dt.Rows[i]["Q_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["Q_Deviation"].ToString().Trim());
                            dDevA19Q[ii] = dDevQValue;

                            intA19Index++;
                        }
                        else if (Regex.Replace(_dt.Rows[i]["CoilName"].ToString().Trim(), @"\D", "") == "20")
                        {
                            dRdcValue = _dt.Rows[i]["DC_ResistanceValue"] == null || _dt.Rows[i]["DC_ResistanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["DC_ResistanceValue"].ToString().Trim());
                            dA20Rdc[ii] = dRdcValue;

                            dRacValue = _dt.Rows[i]["AC_ResistanceValue"] == null || _dt.Rows[i]["AC_ResistanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["AC_ResistanceValue"].ToString().Trim());
                            dA20Rac[ii] = dRacValue;

                            dLValue = _dt.Rows[i]["L_InductanceValue"] == null || _dt.Rows[i]["L_InductanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["L_InductanceValue"].ToString().Trim());
                            dA20L[ii] = dLValue;

                            dCValue = _dt.Rows[i]["C_CapacitanceValue"] == null || _dt.Rows[i]["C_CapacitanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["C_CapacitanceValue"].ToString().Trim());
                            dA20C[ii] = dCValue;

                            dQValue = _dt.Rows[i]["Q_FactorValue"] == null || _dt.Rows[i]["Q_FactorValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["Q_FactorValue"].ToString().Trim());
                            dA20Q[ii] = dQValue;


                            dDevRdcValue = _dt.Rows[i]["DC_Deviation"] == null || _dt.Rows[i]["DC_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["DC_Deviation"].ToString().Trim());
                            dDevA20Rdc[ii] = dDevRdcValue;

                            dDevRacValue = _dt.Rows[i]["AC_Deviation"] == null || _dt.Rows[i]["AC_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["AC_Deviation"].ToString().Trim());
                            dDevA20Rac[ii] = dDevRacValue;

                            dDevLValue = _dt.Rows[i]["L_Deviation"] == null || _dt.Rows[i]["L_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["L_Deviation"].ToString().Trim());
                            dDevA20L[ii] = dDevLValue;

                            dDevCValue = _dt.Rows[i]["C_Deviation"] == null || _dt.Rows[i]["C_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["C_Deviation"].ToString().Trim());
                            dDevA20C[ii] = dDevCValue;

                            dDevQValue = _dt.Rows[i]["Q_Deviation"] == null || _dt.Rows[i]["Q_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["Q_Deviation"].ToString().Trim());
                            dDevA20Q[ii] = dDevQValue;

                            intA20Index++;
                        }
                        else if (Regex.Replace(_dt.Rows[i]["CoilName"].ToString().Trim(), @"\D", "") == "21")
                        {
                            dRdcValue = _dt.Rows[i]["DC_ResistanceValue"] == null || _dt.Rows[i]["DC_ResistanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["DC_ResistanceValue"].ToString().Trim());
                            dA21Rdc[ii] = dRdcValue;

                            dRacValue = _dt.Rows[i]["AC_ResistanceValue"] == null || _dt.Rows[i]["AC_ResistanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["AC_ResistanceValue"].ToString().Trim());
                            dA21Rac[ii] = dRacValue;

                            dLValue = _dt.Rows[i]["L_InductanceValue"] == null || _dt.Rows[i]["L_InductanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["L_InductanceValue"].ToString().Trim());
                            dA21L[ii] = dLValue;

                            dCValue = _dt.Rows[i]["C_CapacitanceValue"] == null || _dt.Rows[i]["C_CapacitanceValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["C_CapacitanceValue"].ToString().Trim());
                            dA21C[ii] = dCValue;

                            dQValue = _dt.Rows[i]["Q_FactorValue"] == null || _dt.Rows[i]["Q_FactorValue"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["Q_FactorValue"].ToString().Trim());
                            dA21Q[ii] = dQValue;


                            dDevRdcValue = _dt.Rows[i]["DC_Deviation"] == null || _dt.Rows[i]["DC_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["DC_Deviation"].ToString().Trim());
                            dDevA21Rdc[ii] = dDevRdcValue;

                            dDevRacValue = _dt.Rows[i]["AC_Deviation"] == null || _dt.Rows[i]["AC_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["AC_Deviation"].ToString().Trim());
                            dDevA21Rac[ii] = dDevRacValue;

                            dDevLValue = _dt.Rows[i]["L_Deviation"] == null || _dt.Rows[i]["L_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["L_Deviation"].ToString().Trim());
                            dDevA21L[ii] = dDevLValue;

                            dDevCValue = _dt.Rows[i]["C_Deviation"] == null || _dt.Rows[i]["C_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["C_Deviation"].ToString().Trim());
                            dDevA21C[ii] = dDevCValue;

                            dDevQValue = _dt.Rows[i]["Q_Deviation"] == null || _dt.Rows[i]["Q_Deviation"].ToString().Trim() == ""
                                ? 0.000d : Convert.ToDouble(_dt.Rows[i]["Q_Deviation"].ToString().Trim());
                            dDevA21Q[ii] = dDevQValue;

                            intA21Index++;
                        }
                    }
                }

                // DC 저항 Chart 그리기  
                SetChart_Depict(chartMeasurementValueDC, "DC 저항", dA01Rdc, dA02Rdc, dA03Rdc, dA04Rdc, dA05Rdc, dA06Rdc, dA07Rdc, dA08Rdc, dA09Rdc, dA10Rdc
                    , dA11Rdc, dA12Rdc, dA13Rdc, dA14Rdc, dA15Rdc, dA16Rdc, dA17Rdc, dA18Rdc, dA19Rdc, dA20Rdc, dA21Rdc, _arrayXLabelValue, 0, "", "");

                // AC 저항 Chart 그리기  
                SetChart_Depict(chartMeasurementValueAC, "AC 저항", dA01Rac, dA02Rac, dA03Rac, dA04Rac, dA05Rac, dA06Rac, dA07Rac, dA08Rac, dA09Rac, dA10Rac
                    , dA11Rac, dA12Rac, dA13Rac, dA14Rac, dA15Rac, dA16Rac, dA17Rac, dA18Rac, dA19Rac, dA20Rac, dA21Rac, _arrayXLabelValue, 0, "", "");

                // L 저항 Chart 그리기  
                SetChart_Depict(chartMeasurementValueL, "인덕턴스", dA01L, dA02L, dA03L, dA04L, dA05L, dA06L, dA07L, dA08L, dA09L, dA10L
                    , dA11L, dA12L, dA13L, dA14L, dA15L, dA16L, dA17L, dA18L, dA19L, dA20L, dA21L, _arrayXLabelValue, 0, "", "");

                // C 저항 Chart 그리기  
                SetChart_Depict(chartMeasurementValueC, "캐패시턴스", dA01C, dA02C, dA03C, dA04C, dA05C, dA06C, dA07C, dA08C, dA09C, dA10C
                    , dA11C, dA12C, dA13C, dA14C, dA15C, dA16C, dA17C, dA18C, dA19C, dA20C, dA21C, _arrayXLabelValue, 0, "", "");

                // Q 저항 Chart 그리기  
                SetChart_Depict(chartMeasurementValueQ, "Q-Factor", dA01Q, dA02Q, dA03Q, dA04Q, dA05Q, dA06Q, dA07Q, dA08Q, dA09Q, dA10Q
                    , dA11Q, dA12Q, dA13Q, dA14Q, dA15Q, dA16Q, dA17Q, dA18Q, dA19Q, dA20Q, dA21Q, _arrayXLabelValue, 0, "", "");

                // DC 편차 Chart 그리기  
                SetChart_Depict(chartAverageValueDC, "DC 편차", dDevA01Rdc, dDevA02Rdc, dDevA03Rdc, dDevA04Rdc, dDevA05Rdc, dDevA06Rdc, dDevA07Rdc, dDevA08Rdc, dDevA09Rdc, dDevA10Rdc
                    , dDevA11Rdc, dDevA12Rdc, dDevA13Rdc, dDevA14Rdc, dDevA15Rdc, dDevA16Rdc, dDevA17Rdc, dDevA18Rdc, dDevA19Rdc, dDevA20Rdc, dDevA21Rdc, _arrayXLabelValue, 0, "", "");

                // AC 편차 Chart 그리기  
                SetChart_Depict(chartAverageValueAC, "AC 편차", dDevA01Rac, dDevA02Rac, dDevA03Rac, dDevA04Rac, dDevA05Rac, dDevA06Rac, dDevA07Rac, dDevA08Rac, dDevA09Rac, dDevA10Rac
                    , dDevA11Rac, dDevA12Rac, dDevA13Rac, dDevA14Rac, dDevA15Rac, dDevA16Rac, dDevA17Rac, dDevA18Rac, dDevA19Rac, dDevA20Rac, dDevA21Rac, _arrayXLabelValue, 0, "", "");

                // L 편차 Chart 그리기  
                SetChart_Depict(chartAverageValueL, "인덕턴스 편차", dDevA01L, dDevA02L, dDevA03L, dDevA04L, dDevA05L, dDevA06L, dDevA07L, dDevA08L, dDevA09L, dDevA10L
                    , dDevA11L, dDevA12L, dDevA13L, dDevA14L, dDevA15L, dDevA16L, dDevA17L, dDevA18L, dDevA19L, dDevA20L, dDevA21L, _arrayXLabelValue, 0, "", "");

                // C 편차 Chart 그리기  
                SetChart_Depict(chartAverageValueC, "캐패시턴스 편차", dDevA01C, dDevA02C, dDevA03C, dDevA04C, dDevA05C, dDevA06C, dDevA07C, dDevA08C, dDevA09C, dDevA10C
                    , dDevA11C, dDevA12C, dDevA13C, dDevA14C, dDevA15C, dDevA16C, dDevA17C, dDevA18C, dDevA19C, dDevA20C, dDevA21C, _arrayXLabelValue, 0, "", "");

                // Q 편차 Chart 그리기  
                SetChart_Depict(chartAverageValueQ, "Q-Factor 편차", dDevA01Q, dDevA02Q, dDevA03Q, dDevA04Q, dDevA05Q, dDevA06Q, dDevA07Q, dDevA08Q, dDevA09Q, dDevA10Q
                    , dDevA11Q, dDevA12Q, dDevA13Q, dDevA14Q, dDevA15Q, dDevA16Q, dDevA17Q, dDevA18Q, dDevA19Q, dDevA20Q, dDevA21Q, _arrayXLabelValue, 0, "", "");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
        }

        // 20240329 한인석
        private string GetCheckBoxCheckedTrueItem(Panel p)
        {
            string returnVal = "";
            if (p == null) return returnVal;

            foreach (Control grv in p.Controls)
            {
                if (grv.GetType().Name.Trim() == "GroupBox")
                {
                    string powerCabinetName = $"chk{grv.Name.Substring(2)}";
                    foreach (Control c in grv.Controls)
                    {
                        if (((CheckBox)c).Checked && c.Name != powerCabinetName)
                        {
                            if (returnVal.Trim() == "")
                                returnVal = "'" + c.Text.Trim() + "'";
                            else
                                returnVal = returnVal.Trim() + ",'" + c.Text.Trim() + "'";
                        }
                    }
                }
            }
            return returnVal;
        }

        private void SetChart_Depict(System.Windows.Forms.DataVisualization.Charting.Chart cht, string chartName
            , double[] dA1Value, double[] dA2Value, double[] dA3Value, double[] dA4Value, double[] dA5Value, double[] dA6Value, double[] dA7Value
            , double[] dA8Value, double[] dA9Value, double[] dA10Value, double[] dA11Value, double[] dA12Value, double[] dA13Value, double[] dA14Value
            , double[] dA15Value, double[] dA16Value, double[] dA17Value, double[] dA18Value, double[] dA19Value, double[] dA20Value, double[] dA21Value
            , string[] arrayXLabelValue, int chartOption, string coilName, string chartType)
        {
            for (int i = 0; i < 21; i++)
            {
                sSeries = cht.Series.Add(string.Format("A{0}", i.ToString().Trim()));
            }

            sSeries = cht.Series.Add("제어용 기준값");
            sSeries = cht.Series.Add("정지용 기준값");
            sSeries = cht.Series.Add("제어용 평균값");
            sSeries = cht.Series.Add("정지용 평균값");

            sSeries.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
            sSeries.XValueType = System.Windows.Forms.DataVisualization.Charting.ChartValueType.Int32;
            sSeries.YValueType = System.Windows.Forms.DataVisualization.Charting.ChartValueType.Double;

            byte chartAreaNo;
            chartAreaNo = 0; //차트 영역 번호.
            m_Chart.InitialChart(ref cht, chartName, 1);

            m_Chart.MakeLegend(ref cht, "A1", "right", chartAreaNo);
            m_Chart.AddChartColumn(ref cht, chartAreaNo, chartAreaNo, "A1", ref arrayXLabelValue, ref dA1Value, true, System.Drawing.Color.Lime, 0, true);

            m_Chart.MakeLegend(ref cht, "A2", "right", chartAreaNo);
            m_Chart.AddChartColumn(ref cht, chartAreaNo, chartAreaNo, "A2", ref arrayXLabelValue, ref dA2Value, true, System.Drawing.Color.Red, 0, true);

            m_Chart.MakeLegend(ref cht, "A3", "right", chartAreaNo);
            m_Chart.AddChartColumn(ref cht, chartAreaNo, chartAreaNo, "A3", ref arrayXLabelValue, ref dA3Value, true, System.Drawing.Color.Blue, 0, true);

            m_Chart.MakeLegend(ref cht, "A4", "right", chartAreaNo);
            m_Chart.AddChartColumn(ref cht, chartAreaNo, chartAreaNo, "A4", ref arrayXLabelValue, ref dA4Value, true, System.Drawing.Color.Cyan, 0, true);

            m_Chart.MakeLegend(ref cht, "A5", "right", chartAreaNo);
            m_Chart.AddChartColumn(ref cht, chartAreaNo, chartAreaNo, "A5", ref arrayXLabelValue, ref dA5Value, true, System.Drawing.Color.Blue, 0, true);

            m_Chart.MakeLegend(ref cht, "A6", "right", chartAreaNo);
            m_Chart.AddChartColumn(ref cht, chartAreaNo, chartAreaNo, "A6", ref arrayXLabelValue, ref dA6Value, true, System.Drawing.Color.Fuchsia, 0, true);

            m_Chart.MakeLegend(ref cht, "A7", "right", chartAreaNo);
            m_Chart.AddChartColumn(ref cht, chartAreaNo, chartAreaNo, "A7", ref arrayXLabelValue, ref dA7Value, true, System.Drawing.Color.Gray, 0, true);

            m_Chart.MakeLegend(ref cht, "A8", "right", chartAreaNo);
            m_Chart.AddChartColumn(ref cht, chartAreaNo, chartAreaNo, "A8", ref arrayXLabelValue, ref dA8Value, true, System.Drawing.Color.DarkBlue, 0, true);

            m_Chart.MakeLegend(ref cht, "A9", "right", chartAreaNo);
            m_Chart.AddChartColumn(ref cht, chartAreaNo, chartAreaNo, "A9", ref arrayXLabelValue, ref dA9Value, true, System.Drawing.Color.Green, 0, true);

            m_Chart.MakeLegend(ref cht, "A10", "right", chartAreaNo);
            m_Chart.AddChartColumn(ref cht, chartAreaNo, chartAreaNo, "A10", ref arrayXLabelValue, ref dA10Value, true, System.Drawing.Color.Violet, 0, true);

            m_Chart.MakeLegend(ref cht, "A11", "right", chartAreaNo);
            m_Chart.AddChartColumn(ref cht, chartAreaNo, chartAreaNo, "A11", ref arrayXLabelValue, ref dA11Value, true, System.Drawing.Color.DarkViolet, 0, true);

            m_Chart.MakeLegend(ref cht, "A12", "right", chartAreaNo);
            m_Chart.AddChartColumn(ref cht, chartAreaNo, chartAreaNo, "A12", ref arrayXLabelValue, ref dA12Value, true, System.Drawing.Color.DarkGoldenrod, 0, true);

            m_Chart.MakeLegend(ref cht, "A13", "right", chartAreaNo);
            m_Chart.AddChartColumn(ref cht, chartAreaNo, chartAreaNo, "A13", ref arrayXLabelValue, ref dA13Value, true, System.Drawing.Color.DodgerBlue, 0, true);

            m_Chart.MakeLegend(ref cht, "A14", "right", chartAreaNo);
            m_Chart.AddChartColumn(ref cht, chartAreaNo, chartAreaNo, "A14", ref arrayXLabelValue, ref dA14Value, true, System.Drawing.Color.DarkCyan, 0, true);

            m_Chart.MakeLegend(ref cht, "A15", "right", chartAreaNo);
            m_Chart.AddChartColumn(ref cht, chartAreaNo, chartAreaNo, "A15", ref arrayXLabelValue, ref dA15Value, true, System.Drawing.Color.Chocolate, 0, true);

            m_Chart.MakeLegend(ref cht, "A16", "right", chartAreaNo);
            m_Chart.AddChartColumn(ref cht, chartAreaNo, chartAreaNo, "A16", ref arrayXLabelValue, ref dA16Value, true, System.Drawing.Color.DarkMagenta, 0, true);

            m_Chart.MakeLegend(ref cht, "A17", "right", chartAreaNo);
            m_Chart.AddChartColumn(ref cht, chartAreaNo, chartAreaNo, "A17", ref arrayXLabelValue, ref dA17Value, true, System.Drawing.Color.MediumAquamarine, 0, true);

            m_Chart.MakeLegend(ref cht, "A18", "right", chartAreaNo);
            m_Chart.AddChartColumn(ref cht, chartAreaNo, chartAreaNo, "A18", ref arrayXLabelValue, ref dA18Value, true, System.Drawing.Color.DarkOrange, 0, true);

            m_Chart.MakeLegend(ref cht, "A19", "right", chartAreaNo);
            m_Chart.AddChartColumn(ref cht, chartAreaNo, chartAreaNo, "A19", ref arrayXLabelValue, ref dA19Value, true, System.Drawing.Color.SkyBlue, 0, true);

            m_Chart.MakeLegend(ref cht, "A20", "right", chartAreaNo);
            m_Chart.AddChartColumn(ref cht, chartAreaNo, chartAreaNo, "A20", ref arrayXLabelValue, ref dA20Value, true, System.Drawing.Color.DeepPink, 0, true);

            m_Chart.MakeLegend(ref cht, "A21", "right", chartAreaNo);
            m_Chart.AddChartColumn(ref cht, chartAreaNo, chartAreaNo, "A21", ref arrayXLabelValue, ref dA21Value, true, System.Drawing.Color.IndianRed, 0, true);

            if (arrayXLabelValue.Length < 30)
                cht.ChartAreas[0].AxisX.Interval = 1; // 차수별일 경우는 X 축 1씩 증가 값 설정
            else
                cht.ChartAreas[0].AxisX.Interval = 10; // 제어봉일 경우는 X 축 10씩 증가 값 설정
		}

		private bool ExcelReportExport(string _strHogi, string _strOhDegree, string _strDRPIType, string _strCoilName, string _strComparisonTarget)
		{
			bool boolResult = false;

			// 보고서 생성 전 Process에 생성된 EXCEL.EXE 목록 추출
			System.Diagnostics.Process[] excelProcessID;
			excelProcessID = System.Diagnostics.Process.GetProcessesByName("EXCEL");
			System.Collections.ArrayList aryProcessID = new System.Collections.ArrayList();

			for (int i = 0; i < excelProcessID.Length; i++)
			{
				aryProcessID.Add(excelProcessID[i].Id);
			}

			// 엑셀 클래스, 어플리케이션, 워크북, 워크시트 정의.
			Excel.Application xlsxApp = new Excel.Application();
			Excel.Workbook xlsxBook = null;
			Excel.Worksheet xlsxSheet1 = null;
			Excel.Worksheet xlsxSheet2 = null;
			string newfileName = "";
			string strTitle = string.Format("DRPI 이력 보고서");
			string strStartPath = Application.StartupPath;


			try
			{
                string fileName = strStartPath + @"\코일보고서Form_Rv1.xlsx";

				DateTime nowDataTime = System.DateTime.Now;

				if (_strCoilName.Trim() == "" && _strCoilName.Trim() == "전체")
					newfileName = strStartPath + @"\Report\" + "DRPI 보고서_" + _strHogi + "_" + _strOhDegree + "_" + _strDRPIType + "_(" + nowDataTime.ToString("yyyyMMddHHmmss") + ").xlsx";
				else
					newfileName = strStartPath + @"\Report\" + "DRPI 보고서_" + _strHogi + "_" + _strOhDegree + "_" + _strDRPIType + "_" + _strCoilName + "_(" + nowDataTime.ToString("yyyyMMddHHmmss") + ").xlsx";

				// 원본파일 존재 확인.
				if (!System.IO.File.Exists(fileName))
				{
					frmMB.lblMessage.Text = "원본 파일이 존재하지 않습니다.";
					frmMB.TopMost = true;
					frmMB.ShowDialog();
					return false;
				}

				xlsxApp.Visible = false;
				xlsxBook = xlsxApp.Workbooks.Open(fileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
				xlsxSheet1 = (Excel.Worksheet)xlsxBook.Sheets.get_Item(1);
				xlsxSheet2 = (Excel.Worksheet)xlsxBook.Sheets.get_Item(2);


				Excel.Range oRng;

				// Header Data Binding
				oRng = xlsxSheet1.get_Range("B1", "O1"); //해당 범위의 셀 획득
				oRng.MergeCells = true; //머지
				oRng.Value = strTitle.Trim();

				// 측정 대상
				oRng = xlsxSheet1.get_Range("C3", "C3"); //해당 범위의 셀 획득
				oRng.MergeCells = true; //머지
				oRng.Value = _strHogi.Trim();

				// 차수
				oRng = xlsxSheet1.get_Range("E3", "E3"); //해당 범위의 셀 획득
				oRng.MergeCells = true; //머지
				oRng.Value = _strOhDegree.Trim();

				// 타입
				oRng = xlsxSheet1.get_Range("G3", "G3"); //해당 범위의 셀 획득
				oRng.MergeCells = true; //머지
				oRng.Value = _strDRPIType.Trim();

				// 코일
				oRng = xlsxSheet1.get_Range("I3", "J3"); //해당 범위의 셀 획득
				oRng.MergeCells = true; //머지
				oRng.Value = _strCoilName.Trim().Replace("'", "");

				// 온도
				oRng = xlsxSheet1.get_Range("L3", "L3"); //해당 범위의 셀 획득
				oRng.MergeCells = true; //머지
				oRng.Value = strTemperature_ReferenceValue.Trim() + "℃";

				// 검사자 
				oRng = xlsxSheet1.get_Range("N3", "O3"); //해당 범위의 셀 획득
				oRng.MergeCells = true; //머지
				oRng.Value = "";

				// 주파수
				oRng = xlsxSheet1.get_Range("C4", "C4"); //해당 범위의 셀 획득
				oRng.MergeCells = true; //머지
				oRng.Value = strFrequency.Trim();

				// 전압라벨
				oRng = xlsxSheet1.get_Range("E4", "E4"); //해당 범위의 셀 획득
				oRng.MergeCells = true; //머지
				oRng.Value = strVoltageLevel.Trim();

				// 비교대상
				oRng = xlsxSheet1.get_Range("G4", "G4"); //해당 범위의 셀 획득
				oRng.MergeCells = true; //머지
				oRng.Value = _strComparisonTarget.Trim();

				// 점검 결과
				oRng = xlsxSheet1.get_Range("I4", "I4"); //해당 범위의 셀 획득
				oRng.MergeCells = true; //머지
				oRng.Value = strMeasurementResult.Trim();

				// 작성일
				oRng = xlsxSheet1.get_Range("K4", "L4"); //해당 범위의 셀 획득
				oRng.MergeCells = true; //머지
				oRng.Value = strMeasurementDate.Trim();

				// 승인자
				oRng = xlsxSheet1.get_Range("N4", "O4"); //해당 범위의 셀 획득
				oRng.MergeCells = true; //머지
				oRng.Value = "";


                // 항목별 타이틀
                string strItemRefDC = "D", strItemAveDC = "E", strItemRefAC = "F", strItemAveAC = "G", strItemRefL = "H", strItemAveL = "I"
                    , strItemRefC = "J", strItemAveC = "K", strItemRefQ = "L", strItemAveQ = "N", strItemResultA = "N", strItemResultB = "O";

                // 항목별 타이틀 위치
                GetItemTitleName(ref strItemRefDC, ref strItemAveDC, ref strItemRefAC, ref strItemAveAC, ref strItemRefL, ref strItemAveL
                    , ref strItemRefC, ref strItemAveC, ref strItemRefQ, ref strItemAveQ, ref strItemResultA, ref strItemResultB);

                // 제어봉
                oRng = xlsxSheet1.get_Range("B6", "B6"); //해당 범위의 셀 획득
                oRng.MergeCells = true; //머지
                oRng.Value = "제어봉";

                // 코일명
                oRng = xlsxSheet1.get_Range("C6", "C6"); //해당 범위의 셀 획득
                oRng.MergeCells = true; //머지
                oRng.Value = "코일명";

                if (chkRdc.Checked)
                {
                    // DC저항(Ω)
                    oRng = xlsxSheet1.get_Range(string.Format("{0}6", strItemRefDC), string.Format("{0}6", strItemRefDC)); //해당 범위의 셀 획득
                    oRng.MergeCells = true; //머지
                    oRng.Value = "DC저항(Ω)";

                    // 편차(Ω)
                    oRng = xlsxSheet1.get_Range(string.Format("{0}6", strItemAveDC), string.Format("{0}6", strItemAveDC)); //해당 범위의 셀 획득
                    oRng.MergeCells = true; //머지
                    oRng.Value = "편차(Ω)";
                }

                if (chkRac.Checked)
                {
                    // AC저항(Ω)
                    oRng = xlsxSheet1.get_Range(string.Format("{0}6", strItemRefAC), string.Format("{0}6", strItemRefAC)); //해당 범위의 셀 획득
                    oRng.MergeCells = true; //머지
                    oRng.Value = "AC저항(Ω)";

                    // 편차(Ω)
                    oRng = xlsxSheet1.get_Range(string.Format("{0}6", strItemAveAC), string.Format("{0}6", strItemAveAC)); //해당 범위의 셀 획득
                    oRng.MergeCells = true; //머지
                    oRng.Value = "편차(Ω)";
                }

                if (chkL.Checked)
                {
                    // 인덕턴스(mH)
                    oRng = xlsxSheet1.get_Range(string.Format("{0}6", strItemRefL), string.Format("{0}6", strItemRefL)); //해당 범위의 셀 획득
                    oRng.MergeCells = true; //머지
                    oRng.Value = "인덕턴스(mH)";

                    // 편차(mH)
                    oRng = xlsxSheet1.get_Range(string.Format("{0}6", strItemAveL), string.Format("{0}6", strItemAveL)); //해당 범위의 셀 획득
                    oRng.MergeCells = true; //머지
                    oRng.Value = "편차(mH)";
                }

                if (chkC.Checked)
                {
                    // 캐패시턴스(uF)
                    oRng = xlsxSheet1.get_Range(string.Format("{0}6", strItemRefC), string.Format("{0}6", strItemRefC)); //해당 범위의 셀 획득
                    oRng.MergeCells = true; //머지
                    oRng.Value = "캐패시턴스(uF)";

                    // 편차
                    oRng = xlsxSheet1.get_Range(string.Format("{0}6", strItemAveC), string.Format("{0}6", strItemAveC)); //해당 범위의 셀 획득
                    oRng.MergeCells = true; //머지
                    oRng.Value = "편차(uF)";
                }

                if (chkQ.Checked)
                {
                    // Q_factor
                    oRng = xlsxSheet1.get_Range(string.Format("{0}6", strItemRefQ), string.Format("{0}6", strItemRefQ)); //해당 범위의 셀 획득
                    oRng.MergeCells = true; //머지
                    oRng.Value = "Q_factor";

                    // 편차
                    oRng = xlsxSheet1.get_Range(string.Format("{0}6", strItemAveQ), string.Format("{0}6", strItemAveQ)); //해당 범위의 셀 획득
                    oRng.MergeCells = true; //머지
                    oRng.Value = "편차";
                }

                // 비고
                oRng = xlsxSheet1.get_Range(string.Format("{0}6", strItemResultA), string.Format("{0}6", strItemResultB)); //해당 범위의 셀 획득
                oRng.MergeCells = true; //머지
                oRng.Value = "제어봉";

				// 보고서 생성
				int intReportPoint = 7;
				for (int i = 0; i < dgvMeasurement.RowCount; i++)
				{
					// Header Data Binding
					oRng = xlsxSheet1.get_Range("B" + (i + intReportPoint), "B" + (i + intReportPoint)); //해당 범위의 셀 획득
					oRng.MergeCells = true; //머지
					oRng.Value = dgvMeasurement.Rows[i].Cells["ControlName"].Value == null ? "" : dgvMeasurement.Rows[i].Cells["ControlName"].Value.ToString().Trim();
					oRng.Borders.LineStyle = 1;
					oRng.Font.Color = dgvMeasurement.Rows[i].Cells["ControlName"].Style.ForeColor;

					oRng = xlsxSheet1.get_Range("C" + (i + intReportPoint), "C" + (i + intReportPoint));
					oRng.MergeCells = true; //머지
					oRng.Value = dgvMeasurement.Rows[i].Cells["CoilName"].Value == null ? "" : dgvMeasurement.Rows[i].Cells["CoilName"].Value.ToString().Trim();
                    oRng.Borders.LineStyle = 1;
					oRng.Font.Color = dgvMeasurement.Rows[i].Cells["CoilName"].Style.ForeColor;

                    if (chkRdc.Checked)
                    {
                        oRng = xlsxSheet1.get_Range(string.Format("{0}{1}", strItemRefDC, (i + intReportPoint), string.Format("{0}{1}", strItemRefDC, (i + intReportPoint))));
                        oRng.MergeCells = true; //머지
                        oRng.Value = dgvMeasurement.Rows[i].Cells["DC_ResistanceValue"].Value == null ? "" : dgvMeasurement.Rows[i].Cells["DC_ResistanceValue"].Value.ToString().Trim();
                        oRng.Borders.LineStyle = 1;
                        oRng.Font.Color = dgvMeasurement.Rows[i].Cells["DC_ResistanceValue"].Style.ForeColor;

						oRng = xlsxSheet1.get_Range(string.Format("{0}{1}", strItemAveDC, (i + intReportPoint), string.Format("{0}{1}", strItemAveDC, (i + intReportPoint))));
                        oRng.MergeCells = true; //머지
                        oRng.Value = dgvMeasurement.Rows[i].Cells["DC_Deviation"].Value == null ? "" : dgvMeasurement.Rows[i].Cells["DC_Deviation"].Value.ToString().Trim();
                        oRng.Borders.LineStyle = 1;
                        oRng.Font.Color = dgvMeasurement.Rows[i].Cells["DC_Deviation"].Style.ForeColor;
                    }

                    if (chkRac.Checked)
                    {
                        oRng = xlsxSheet1.get_Range(string.Format("{0}{1}", strItemRefAC, (i + intReportPoint), string.Format("{0}{1}", strItemRefAC, (i + intReportPoint))));
                        oRng.MergeCells = true; //머지
                        oRng.Value = dgvMeasurement.Rows[i].Cells["AC_ResistanceValue"].Value == null ? "" : dgvMeasurement.Rows[i].Cells["AC_ResistanceValue"].Value.ToString().Trim();
                        oRng.Borders.LineStyle = 1;
                        oRng.Font.Color = dgvMeasurement.Rows[i].Cells["AC_ResistanceValue"].Style.ForeColor;

                        oRng = xlsxSheet1.get_Range(string.Format("{0}{1}", strItemAveAC, (i + intReportPoint), string.Format("{0}{1}", strItemAveAC, (i + intReportPoint))));
                        oRng.MergeCells = true; //머지
                        oRng.Value = dgvMeasurement.Rows[i].Cells["AC_Deviation"].Value == null ? "" : dgvMeasurement.Rows[i].Cells["AC_Deviation"].Value.ToString().Trim();
                        oRng.Borders.LineStyle = 1;
                        oRng.Font.Color = dgvMeasurement.Rows[i].Cells["AC_Deviation"].Style.ForeColor;
                    }

                    if (chkL.Checked)
                    {
                        oRng = xlsxSheet1.get_Range(string.Format("{0}{1}", strItemRefL, (i + intReportPoint), string.Format("{0}{1}", strItemRefL, (i + intReportPoint))));
                        oRng.MergeCells = true; //머지
                        oRng.Value = dgvMeasurement.Rows[i].Cells["L_InductanceValue"].Value == null ? "" : dgvMeasurement.Rows[i].Cells["L_InductanceValue"].Value.ToString().Trim();
                        oRng.Borders.LineStyle = 1;
                        oRng.Font.Color = dgvMeasurement.Rows[i].Cells["L_InductanceValue"].Style.ForeColor;

                        oRng = xlsxSheet1.get_Range(string.Format("{0}{1}", strItemAveL, (i + intReportPoint), string.Format("{0}{1}", strItemAveL, (i + intReportPoint))));
                        oRng.MergeCells = true; //머지
                        oRng.Value = dgvMeasurement.Rows[i].Cells["L_Deviation"].Value == null ? "" : dgvMeasurement.Rows[i].Cells["L_Deviation"].Value.ToString().Trim();
                        oRng.Borders.LineStyle = 1;
                        oRng.Font.Color = dgvMeasurement.Rows[i].Cells["L_Deviation"].Style.ForeColor;
                    }

                    if (chkC.Checked)
                    {
                        oRng = xlsxSheet1.get_Range(string.Format("{0}{1}", strItemRefC, (i + intReportPoint), string.Format("{0}{1}", strItemRefC, (i + intReportPoint))));
                        oRng.MergeCells = true; //머지
                        oRng.Value = dgvMeasurement.Rows[i].Cells["C_CapacitanceValue"].Value == null ? "" : dgvMeasurement.Rows[i].Cells["C_CapacitanceValue"].Value.ToString().Trim();
                        oRng.Borders.LineStyle = 1;
                        oRng.Font.Color = dgvMeasurement.Rows[i].Cells["C_CapacitanceValue"].Style.ForeColor;

                        oRng = xlsxSheet1.get_Range(string.Format("{0}{1}", strItemAveC, (i + intReportPoint), string.Format("{0}{1}", strItemAveC, (i + intReportPoint))));
                        oRng.MergeCells = true; //머지
                        oRng.Value = dgvMeasurement.Rows[i].Cells["C_Deviation"].Value == null ? "" : dgvMeasurement.Rows[i].Cells["C_Deviation"].Value.ToString().Trim();
                        oRng.Borders.LineStyle = 1;
                        oRng.Font.Color = dgvMeasurement.Rows[i].Cells["C_Deviation"].Style.ForeColor;
                    }

                    if (chkQ.Checked)
                    {
                        oRng = xlsxSheet1.get_Range(string.Format("{0}{1}", strItemRefQ, (i + intReportPoint), string.Format("{0}{1}", strItemRefQ, (i + intReportPoint))));
                        oRng.MergeCells = true; //머지
                        oRng.Value = dgvMeasurement.Rows[i].Cells["Q_FactorValue"].Value == null ? "" : dgvMeasurement.Rows[i].Cells["Q_FactorValue"].Value.ToString().Trim();
                        oRng.Borders.LineStyle = 1;
                        oRng.Font.Color = dgvMeasurement.Rows[i].Cells["Q_FactorValue"].Style.ForeColor;

                        oRng = xlsxSheet1.get_Range(string.Format("{0}{1}", strItemAveQ, (i + intReportPoint), string.Format("{0}{1}", strItemAveQ, (i + intReportPoint))));
                        oRng.MergeCells = true; //머지
                        oRng.Value = dgvMeasurement.Rows[i].Cells["Q_Deviation"].Value == null ? "" : dgvMeasurement.Rows[i].Cells["Q_Deviation"].Value.ToString().Trim();
                        oRng.Borders.LineStyle = 1;
                        oRng.Font.Color = dgvMeasurement.Rows[i].Cells["Q_Deviation"].Style.ForeColor;
                    }

                    oRng = xlsxSheet1.get_Range(string.Format("{0}{1}", strItemResultA, (i + intReportPoint)), string.Format("{0}{1}", strItemResultB, (i + intReportPoint)));
					oRng.MergeCells = true; //머지
					oRng.Value = dgvMeasurement.Rows[i].Cells["Result"].Value == null ? "" : dgvMeasurement.Rows[i].Cells["Result"].Value.ToString().Trim();
                    oRng.Borders.LineStyle = 1;
                    oRng.Font.Color = dgvMeasurement.Rows[i].Cells["Result"].Style.ForeColor;
				}

				int intLeft = 2, intTop = 1, intWidth = 915, intHeight = 295;

				if (chkRdc.Checked)
				{
					chartMeasurementValueDC.SaveImage(strStartPath + @"\\Chart1.bmp", ChartImageFormat.Bmp);
					xlsxSheet2.Shapes.AddPicture(strStartPath + @"\\Chart1.bmp", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, intLeft, intTop, intWidth, intHeight); // Left, Top, Width, Height

					intTop = intTop + 296;
				}

				if (chkRac.Checked)
				{
					chartMeasurementValueAC.SaveImage(strStartPath + @"\\Chart1.bmp", ChartImageFormat.Bmp);
					xlsxSheet2.Shapes.AddPicture(strStartPath + @"\\Chart1.bmp", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, intLeft, intTop, intWidth, intHeight); // Left, Top, Width, Height

					if (chkRdc.Checked)
						intTop = intTop + 298;
					else
						intTop = intTop + 296;
				}

				if (chkL.Checked)
				{
					chartMeasurementValueL.SaveImage(strStartPath + @"\\Chart1.bmp", ChartImageFormat.Bmp);
					xlsxSheet2.Shapes.AddPicture(strStartPath + @"\\Chart1.bmp", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, intLeft, intTop, intWidth, intHeight); // Left, Top, Width, Height

					if (chkRdc.Checked && chkRac.Checked)
						intTop = intTop + 296;
					else if (chkRdc.Checked && !chkRac.Checked)
						intTop = intTop + 298;
					else if (!chkRdc.Checked && chkRac.Checked)
						intTop = intTop + 298;
					else
						intTop = intTop + 296;
				}

				if (chkC.Checked)
				{
					chartMeasurementValueC.SaveImage(strStartPath + @"\\Chart1.bmp", ChartImageFormat.Bmp);
					xlsxSheet2.Shapes.AddPicture(strStartPath + @"\\Chart1.bmp", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, intLeft, intTop, intWidth, intHeight); // Left, Top, Width, Height

					if (chkRdc.Checked && chkRac.Checked && chkL.Checked)
						intTop = intTop + 298;
					else if (chkRdc.Checked && chkRac.Checked && !chkL.Checked)
						intTop = intTop + 296;
					else if (chkRdc.Checked && !chkRac.Checked && chkL.Checked)
						intTop = intTop + 296;
					else if (!chkRdc.Checked && chkRac.Checked && chkL.Checked)
						intTop = intTop + 296;
					else if (chkRdc.Checked && !chkRac.Checked && !chkL.Checked)
						intTop = intTop + 298;
					else if (!chkRdc.Checked && chkRac.Checked && !chkL.Checked)
						intTop = intTop + 298;
					else if (!chkRdc.Checked && !chkRac.Checked && chkL.Checked)
						intTop = intTop + 298;
					else
						intTop = intTop + 296;
				}

				if (chkQ.Checked)
				{
					chartMeasurementValueQ.SaveImage(strStartPath + @"\\Chart1.bmp", ChartImageFormat.Bmp);
					xlsxSheet2.Shapes.AddPicture(strStartPath + @"\\Chart1.bmp", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, intLeft, intTop, intWidth, intHeight); // Left, Top, Width, Height

					if (chkRdc.Checked && chkRac.Checked && chkL.Checked && chkC.Checked)
						intTop = intTop + 296;
					else if (chkRdc.Checked && chkRac.Checked && chkL.Checked && !chkC.Checked)
						intTop = intTop + 298;
					else if (chkRdc.Checked && chkRac.Checked && !chkL.Checked && chkC.Checked)
						intTop = intTop + 298;
					else if (chkRdc.Checked && !chkRac.Checked && chkL.Checked && chkC.Checked)
						intTop = intTop + 299;
					else if (!chkRdc.Checked && chkRac.Checked && chkL.Checked && chkC.Checked)
						intTop = intTop + 299;
					else if (chkRdc.Checked && chkRac.Checked && !chkL.Checked && !chkC.Checked)
						intTop = intTop + 296;
					else if (chkRdc.Checked && !chkRac.Checked && chkL.Checked && !chkC.Checked)
						intTop = intTop + 296;
					else if (!chkRdc.Checked && chkRac.Checked && chkL.Checked && !chkC.Checked)
						intTop = intTop + 296;
					else if (chkRdc.Checked && !chkRac.Checked && !chkL.Checked && chkC.Checked)
						intTop = intTop + 296;
					else if (!chkRdc.Checked && chkRac.Checked && !chkL.Checked && chkC.Checked)
						intTop = intTop + 296;
					else if (!chkRdc.Checked && !chkRac.Checked && chkL.Checked && chkC.Checked)
						intTop = intTop + 296;
					else if (chkRdc.Checked && !chkRac.Checked && !chkL.Checked && !chkC.Checked)
						intTop = intTop + 298;
					else if (!chkRdc.Checked && chkRac.Checked && !chkL.Checked && !chkC.Checked)
						intTop = intTop + 298;
					else if (!chkRdc.Checked && !chkRac.Checked && chkL.Checked && !chkC.Checked)
						intTop = intTop + 298;
					else if (!chkRdc.Checked && !chkRac.Checked && !chkL.Checked && chkC.Checked)
						intTop = intTop + 298;
					else
						intTop = intTop + 296;
				}

				if (chkRdc.Checked)
				{
					chartAverageValueDC.SaveImage(strStartPath + @"\\Chart1.bmp", ChartImageFormat.Bmp);
					xlsxSheet2.Shapes.AddPicture(strStartPath + @"\\Chart1.bmp", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, intLeft, intTop, intWidth, intHeight); // Left, Top, Width, Height

					if (chkRdc.Checked && chkRac.Checked && chkL.Checked && chkC.Checked && chkQ.Checked)
						intTop = intTop + 298;
					else if (chkRdc.Checked && chkRac.Checked && chkL.Checked && chkC.Checked && !chkQ.Checked)
						intTop = intTop + 296;
					else if (chkRdc.Checked && chkRac.Checked && chkL.Checked && !chkC.Checked && chkQ.Checked)
						intTop = intTop + 296;
					else if (chkRdc.Checked && chkRac.Checked && !chkL.Checked && chkC.Checked && chkQ.Checked)
						intTop = intTop + 296;
					else if (chkRdc.Checked && !chkRac.Checked && chkL.Checked && chkC.Checked && chkQ.Checked)
						intTop = intTop + 296;
					else if (!chkRdc.Checked && chkRac.Checked && chkL.Checked && chkC.Checked && chkQ.Checked)
						intTop = intTop + 296;
					else if (chkRdc.Checked && chkRac.Checked && chkL.Checked && !chkC.Checked && !chkQ.Checked)
						intTop = intTop + 298;
					else if (chkRdc.Checked && chkRac.Checked && !chkL.Checked && chkC.Checked && !chkQ.Checked)
						intTop = intTop + 298;
					else if (chkRdc.Checked && !chkRac.Checked && chkL.Checked && chkC.Checked && !chkQ.Checked)
						intTop = intTop + 298;
					else if (!chkRdc.Checked && chkRac.Checked && chkL.Checked && chkC.Checked && !chkQ.Checked)
						intTop = intTop + 298;
					else if (chkRdc.Checked && chkRac.Checked && !chkL.Checked && !chkC.Checked && chkQ.Checked)
						intTop = intTop + 298;
					else if (chkRdc.Checked && !chkRac.Checked && chkL.Checked && !chkC.Checked && chkQ.Checked)
						intTop = intTop + 298;
					else if (!chkRdc.Checked && chkRac.Checked && chkL.Checked && !chkC.Checked && chkQ.Checked)
						intTop = intTop + 298;
					else if (chkRdc.Checked && !chkRac.Checked && !chkL.Checked && chkC.Checked && chkQ.Checked)
						intTop = intTop + 298;
					else if (!chkRdc.Checked && chkRac.Checked && !chkL.Checked && chkC.Checked && chkQ.Checked)
						intTop = intTop + 298;
					else if (!chkRdc.Checked && !chkRac.Checked && chkL.Checked && chkC.Checked && chkQ.Checked)
						intTop = intTop + 298;
					else if (chkRdc.Checked && chkRac.Checked && !chkL.Checked && !chkC.Checked && !chkQ.Checked)
						intTop = intTop + 296;
					else if (chkRdc.Checked && !chkRac.Checked && chkL.Checked && !chkC.Checked && !chkQ.Checked)
						intTop = intTop + 296;
					else if (!chkRdc.Checked && chkRac.Checked && chkL.Checked && !chkC.Checked && !chkQ.Checked)
						intTop = intTop + 296;
					else if (chkRdc.Checked && !chkRac.Checked && !chkL.Checked && chkC.Checked && !chkQ.Checked)
						intTop = intTop + 296;
					else if (!chkRdc.Checked && chkRac.Checked && !chkL.Checked && chkC.Checked && !chkQ.Checked)
						intTop = intTop + 296;
					else if (!chkRdc.Checked && !chkRac.Checked && chkL.Checked && chkC.Checked && !chkQ.Checked)
						intTop = intTop + 296;
					else if (chkRdc.Checked && !chkRac.Checked && !chkL.Checked && !chkC.Checked && chkQ.Checked)
						intTop = intTop + 296;
					else if (!chkRdc.Checked && chkRac.Checked && !chkL.Checked && !chkC.Checked && chkQ.Checked)
						intTop = intTop + 296;
					else if (!chkRdc.Checked && !chkRac.Checked && chkL.Checked && !chkC.Checked && chkQ.Checked)
						intTop = intTop + 296;
					else if (!chkRdc.Checked && chkRac.Checked && !chkL.Checked && !chkC.Checked && chkQ.Checked)
						intTop = intTop + 296;
					else if (chkRdc.Checked && !chkRac.Checked && !chkL.Checked && !chkC.Checked && !chkQ.Checked)
						intTop = intTop + 298;
					else if (!chkRdc.Checked && chkRac.Checked && !chkL.Checked && !chkC.Checked && !chkQ.Checked)
						intTop = intTop + 298;
					else if (!chkRdc.Checked && !chkRac.Checked && chkL.Checked && !chkC.Checked && !chkQ.Checked)
						intTop = intTop + 298;
					else if (!chkRdc.Checked && !chkRac.Checked && !chkL.Checked && chkC.Checked && !chkQ.Checked)
						intTop = intTop + 298;
					else if (!chkRdc.Checked && !chkRac.Checked && !chkL.Checked && !chkC.Checked && chkQ.Checked)
						intTop = intTop + 298;
					else
						intTop = intTop + 296;
				}

				if (chkRac.Checked)
				{
					chartAverageValueAC.SaveImage(strStartPath + @"\\Chart1.bmp", ChartImageFormat.Bmp);
					xlsxSheet2.Shapes.AddPicture(strStartPath + @"\\Chart1.bmp", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, intLeft, intTop, intWidth, intHeight); // Left, Top, Width, Height

					if (chkRdc.Checked && chkRac.Checked && chkL.Checked && chkC.Checked && chkQ.Checked)
						intTop = intTop + 296;
					else if (chkRdc.Checked && chkRac.Checked && chkL.Checked && chkC.Checked && !chkQ.Checked)
						intTop = intTop + 298;
					else if (chkRdc.Checked && chkRac.Checked && chkL.Checked && !chkC.Checked && chkQ.Checked)
						intTop = intTop + 298;
					else if (chkRdc.Checked && chkRac.Checked && !chkL.Checked && chkC.Checked && chkQ.Checked)
						intTop = intTop + 298;
					else if (chkRdc.Checked && !chkRac.Checked && chkL.Checked && chkC.Checked && chkQ.Checked)
						intTop = intTop + 298;
					else if (!chkRdc.Checked && chkRac.Checked && chkL.Checked && chkC.Checked && chkQ.Checked)
						intTop = intTop + 296;
					else if (chkRdc.Checked && chkRac.Checked && chkL.Checked && !chkC.Checked && !chkQ.Checked)
						intTop = intTop + 296;
					else if (chkRdc.Checked && chkRac.Checked && !chkL.Checked && chkC.Checked && !chkQ.Checked)
						intTop = intTop + 296;
					else if (chkRdc.Checked && !chkRac.Checked && chkL.Checked && chkC.Checked && !chkQ.Checked)
						intTop = intTop + 296;
					else if (!chkRdc.Checked && chkRac.Checked && chkL.Checked && chkC.Checked && !chkQ.Checked)
						intTop = intTop + 296;
					else if (chkRdc.Checked && chkRac.Checked && !chkL.Checked && !chkC.Checked && chkQ.Checked)
						intTop = intTop + 296;
					else if (chkRdc.Checked && !chkRac.Checked && chkL.Checked && !chkC.Checked && chkQ.Checked)
						intTop = intTop + 296;
					else if (!chkRdc.Checked && chkRac.Checked && chkL.Checked && !chkC.Checked && chkQ.Checked)
						intTop = intTop + 296;
					else if (chkRdc.Checked && !chkRac.Checked && !chkL.Checked && chkC.Checked && chkQ.Checked)
						intTop = intTop + 296;
					else if (!chkRdc.Checked && chkRac.Checked && !chkL.Checked && chkC.Checked && chkQ.Checked)
						intTop = intTop + 298;
					else if (!chkRdc.Checked && !chkRac.Checked && chkL.Checked && chkC.Checked && chkQ.Checked)
						intTop = intTop + 296;
					else if (chkRdc.Checked && chkRac.Checked && !chkL.Checked && !chkC.Checked && !chkQ.Checked)
						intTop = intTop + 298;
					else if (chkRdc.Checked && !chkRac.Checked && chkL.Checked && !chkC.Checked && !chkQ.Checked)
						intTop = intTop + 298;
					else if (!chkRdc.Checked && chkRac.Checked && chkL.Checked && !chkC.Checked && !chkQ.Checked)
						intTop = intTop + 296;
					else if (chkRdc.Checked && !chkRac.Checked && !chkL.Checked && chkC.Checked && !chkQ.Checked)
						intTop = intTop + 298;
					else if (!chkRdc.Checked && chkRac.Checked && !chkL.Checked && chkC.Checked && !chkQ.Checked)
						intTop = intTop + 296;
					else if (!chkRdc.Checked && !chkRac.Checked && chkL.Checked && chkC.Checked && !chkQ.Checked)
						intTop = intTop + 298;
					else if (chkRdc.Checked && !chkRac.Checked && !chkL.Checked && !chkC.Checked && chkQ.Checked)
						intTop = intTop + 298;
					else if (!chkRdc.Checked && chkRac.Checked && !chkL.Checked && !chkC.Checked && chkQ.Checked)
						intTop = intTop + 296;
					else if (!chkRdc.Checked && !chkRac.Checked && chkL.Checked && !chkC.Checked && chkQ.Checked)
						intTop = intTop + 298;
					else if (!chkRdc.Checked && chkRac.Checked && !chkL.Checked && !chkC.Checked && chkQ.Checked)
						intTop = intTop + 298;
					else if (chkRdc.Checked && !chkRac.Checked && !chkL.Checked && !chkC.Checked && !chkQ.Checked)
						intTop = intTop + 296;
					else if (!chkRdc.Checked && chkRac.Checked && !chkL.Checked && !chkC.Checked && !chkQ.Checked)
						intTop = intTop + 296;
					else if (!chkRdc.Checked && !chkRac.Checked && chkL.Checked && !chkC.Checked && !chkQ.Checked)
						intTop = intTop + 296;
					else if (!chkRdc.Checked && !chkRac.Checked && !chkL.Checked && chkC.Checked && !chkQ.Checked)
						intTop = intTop + 296;
					else if (!chkRdc.Checked && !chkRac.Checked && !chkL.Checked && !chkC.Checked && chkQ.Checked)
						intTop = intTop + 296;
					else
						intTop = intTop + 296;
				}

				if (chkL.Checked)
				{
					chartAverageValueL.SaveImage(strStartPath + @"\\Chart1.bmp", ChartImageFormat.Bmp);
					xlsxSheet2.Shapes.AddPicture(strStartPath + @"\\Chart1.bmp", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, intLeft, intTop, intWidth, intHeight); // Left, Top, Width, Height

					if (chkRdc.Checked && chkRac.Checked && chkL.Checked && chkC.Checked && chkQ.Checked)
						intTop = intTop + 298;
					else if (chkRdc.Checked && chkRac.Checked && chkL.Checked && chkC.Checked && !chkQ.Checked)
						intTop = intTop + 296;
					else if (chkRdc.Checked && chkRac.Checked && chkL.Checked && !chkC.Checked && chkQ.Checked)
						intTop = intTop + 296;
					else if (chkRdc.Checked && chkRac.Checked && !chkL.Checked && chkC.Checked && chkQ.Checked)
						intTop = intTop + 296;
					else if (chkRdc.Checked && !chkRac.Checked && chkL.Checked && chkC.Checked && chkQ.Checked)
						intTop = intTop + 296;
					else if (!chkRdc.Checked && chkRac.Checked && chkL.Checked && chkC.Checked && chkQ.Checked)
						intTop = intTop + 296;
					else if (chkRdc.Checked && chkRac.Checked && chkL.Checked && !chkC.Checked && !chkQ.Checked)
						intTop = intTop + 296;
					else if (chkRdc.Checked && chkRac.Checked && !chkL.Checked && chkC.Checked && !chkQ.Checked)
						intTop = intTop + 296;
					else if (chkRdc.Checked && !chkRac.Checked && chkL.Checked && chkC.Checked && !chkQ.Checked)
						intTop = intTop + 296;
					else if (!chkRdc.Checked && chkRac.Checked && chkL.Checked && chkC.Checked && !chkQ.Checked)
						intTop = intTop + 296;
					else if (chkRdc.Checked && chkRac.Checked && !chkL.Checked && !chkC.Checked && chkQ.Checked)
						intTop = intTop + 296;
					else if (chkRdc.Checked && !chkRac.Checked && chkL.Checked && !chkC.Checked && chkQ.Checked)
						intTop = intTop + 296;
					else if (!chkRdc.Checked && chkRac.Checked && chkL.Checked && !chkC.Checked && chkQ.Checked)
						intTop = intTop + 296;
					else if (chkRdc.Checked && !chkRac.Checked && !chkL.Checked && chkC.Checked && chkQ.Checked)
						intTop = intTop + 296;
					else if (!chkRdc.Checked && chkRac.Checked && !chkL.Checked && chkC.Checked && chkQ.Checked)
						intTop = intTop + 296;
					else if (!chkRdc.Checked && !chkRac.Checked && chkL.Checked && chkC.Checked && chkQ.Checked)
						intTop = intTop + 298;
					else if (chkRdc.Checked && chkRac.Checked && !chkL.Checked && !chkC.Checked && !chkQ.Checked)
						intTop = intTop + 298;
					else if (chkRdc.Checked && !chkRac.Checked && chkL.Checked && !chkC.Checked && !chkQ.Checked)
						intTop = intTop + 298;
					else if (!chkRdc.Checked && chkRac.Checked && chkL.Checked && !chkC.Checked && !chkQ.Checked)
						intTop = intTop + 298;
					else if (chkRdc.Checked && !chkRac.Checked && !chkL.Checked && chkC.Checked && !chkQ.Checked)
						intTop = intTop + 298;
					else if (!chkRdc.Checked && chkRac.Checked && !chkL.Checked && chkC.Checked && !chkQ.Checked)
						intTop = intTop + 298;
					else if (!chkRdc.Checked && !chkRac.Checked && chkL.Checked && chkC.Checked && !chkQ.Checked)
						intTop = intTop + 296;
					else if (chkRdc.Checked && !chkRac.Checked && !chkL.Checked && !chkC.Checked && chkQ.Checked)
						intTop = intTop + 298;
					else if (!chkRdc.Checked && chkRac.Checked && !chkL.Checked && !chkC.Checked && chkQ.Checked)
						intTop = intTop + 298;
					else if (!chkRdc.Checked && !chkRac.Checked && chkL.Checked && !chkC.Checked && chkQ.Checked)
						intTop = intTop + 296;
					else if (!chkRdc.Checked && chkRac.Checked && !chkL.Checked && !chkC.Checked && chkQ.Checked)
						intTop = intTop + 298;
					else if (chkRdc.Checked && !chkRac.Checked && !chkL.Checked && !chkC.Checked && !chkQ.Checked)
						intTop = intTop + 296;
					else if (!chkRdc.Checked && chkRac.Checked && !chkL.Checked && !chkC.Checked && !chkQ.Checked)
						intTop = intTop + 296;
					else if (!chkRdc.Checked && !chkRac.Checked && chkL.Checked && !chkC.Checked && !chkQ.Checked)
						intTop = intTop + 296;
					else if (!chkRdc.Checked && !chkRac.Checked && !chkL.Checked && chkC.Checked && !chkQ.Checked)
						intTop = intTop + 296;
					else if (!chkRdc.Checked && !chkRac.Checked && !chkL.Checked && !chkC.Checked && chkQ.Checked)
						intTop = intTop + 296;
					else
						intTop = intTop + 296;
				}

				if (chkC.Checked)
				{
					chartAverageValueC.SaveImage(strStartPath + @"\\Chart1.bmp", ChartImageFormat.Bmp);
					xlsxSheet2.Shapes.AddPicture(strStartPath + @"\\Chart1.bmp", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, intLeft, intTop, intWidth, intHeight); // Left, Top, Width, Height

					if (chkRdc.Checked && chkRac.Checked && chkL.Checked && chkC.Checked && chkQ.Checked)
						intTop = intTop + 296;
					else if (chkRdc.Checked && chkRac.Checked && chkL.Checked && chkC.Checked && !chkQ.Checked)
						intTop = intTop + 296;
					else if (chkRdc.Checked && chkRac.Checked && chkL.Checked && !chkC.Checked && chkQ.Checked)
						intTop = intTop + 296;
					else if (chkRdc.Checked && chkRac.Checked && !chkL.Checked && chkC.Checked && chkQ.Checked)
						intTop = intTop + 296;
					else if (chkRdc.Checked && !chkRac.Checked && chkL.Checked && chkC.Checked && chkQ.Checked)
						intTop = intTop + 296;
					else if (!chkRdc.Checked && chkRac.Checked && chkL.Checked && chkC.Checked && chkQ.Checked)
						intTop = intTop + 296;
					else if (chkRdc.Checked && chkRac.Checked && chkL.Checked && !chkC.Checked && !chkQ.Checked)
						intTop = intTop + 298;
					else if (chkRdc.Checked && chkRac.Checked && !chkL.Checked && chkC.Checked && !chkQ.Checked)
						intTop = intTop + 298;
					else if (chkRdc.Checked && !chkRac.Checked && chkL.Checked && chkC.Checked && !chkQ.Checked)
						intTop = intTop + 298;
					else if (!chkRdc.Checked && chkRac.Checked && chkL.Checked && chkC.Checked && !chkQ.Checked)
						intTop = intTop + 298;
					else if (chkRdc.Checked && chkRac.Checked && !chkL.Checked && !chkC.Checked && chkQ.Checked)
						intTop = intTop + 298;
					else if (chkRdc.Checked && !chkRac.Checked && chkL.Checked && !chkC.Checked && chkQ.Checked)
						intTop = intTop + 298;
					else if (!chkRdc.Checked && chkRac.Checked && chkL.Checked && !chkC.Checked && chkQ.Checked)
						intTop = intTop + 298;
					else if (chkRdc.Checked && !chkRac.Checked && !chkL.Checked && chkC.Checked && chkQ.Checked)
						intTop = intTop + 296;
					else if (!chkRdc.Checked && chkRac.Checked && !chkL.Checked && chkC.Checked && chkQ.Checked)
						intTop = intTop + 296;
					else if (!chkRdc.Checked && !chkRac.Checked && chkL.Checked && chkC.Checked && chkQ.Checked)
						intTop = intTop + 296;
					else if (chkRdc.Checked && chkRac.Checked && !chkL.Checked && !chkC.Checked && !chkQ.Checked)
						intTop = intTop + 296;
					else if (chkRdc.Checked && !chkRac.Checked && chkL.Checked && !chkC.Checked && !chkQ.Checked)
						intTop = intTop + 296;
					else if (!chkRdc.Checked && chkRac.Checked && chkL.Checked && !chkC.Checked && !chkQ.Checked)
						intTop = intTop + 296;
					else if (chkRdc.Checked && !chkRac.Checked && !chkL.Checked && chkC.Checked && !chkQ.Checked)
						intTop = intTop + 296;
					else if (!chkRdc.Checked && chkRac.Checked && !chkL.Checked && chkC.Checked && !chkQ.Checked)
						intTop = intTop + 296;
					else if (!chkRdc.Checked && !chkRac.Checked && chkL.Checked && chkC.Checked && !chkQ.Checked)
						intTop = intTop + 296;
					else if (chkRdc.Checked && !chkRac.Checked && !chkL.Checked && !chkC.Checked && chkQ.Checked)
						intTop = intTop + 296;
					else if (!chkRdc.Checked && chkRac.Checked && !chkL.Checked && !chkC.Checked && chkQ.Checked)
						intTop = intTop + 296;
					else if (!chkRdc.Checked && !chkRac.Checked && chkL.Checked && !chkC.Checked && chkQ.Checked)
						intTop = intTop + 296;
					else if (!chkRdc.Checked && chkRac.Checked && !chkL.Checked && !chkC.Checked && chkQ.Checked)
						intTop = intTop + 296;
					else if (chkRdc.Checked && !chkRac.Checked && !chkL.Checked && !chkC.Checked && !chkQ.Checked)
						intTop = intTop + 298;
					else if (!chkRdc.Checked && chkRac.Checked && !chkL.Checked && !chkC.Checked && !chkQ.Checked)
						intTop = intTop + 298;
					else if (!chkRdc.Checked && !chkRac.Checked && chkL.Checked && !chkC.Checked && !chkQ.Checked)
						intTop = intTop + 298;
					else if (!chkRdc.Checked && !chkRac.Checked && !chkL.Checked && chkC.Checked && !chkQ.Checked)
						intTop = intTop + 298;
					else if (!chkRdc.Checked && !chkRac.Checked && !chkL.Checked && !chkC.Checked && chkQ.Checked)
						intTop = intTop + 298;
					else
						intTop = intTop + 296;
				}

                if (chkQ.Checked)
                {
                    chartAverageValueQ.SaveImage(strStartPath + @"\\Chart1.bmp", ChartImageFormat.Bmp);
                    xlsxSheet2.Shapes.AddPicture(strStartPath + @"\\Chart1.bmp", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, intLeft, intTop, intWidth, intHeight); // Left, Top, Width, Height
                }

				boolResult = true;
			}
			catch (Exception ex)
			{
				frmMB.lblMessage.Text = "보고서 생성 중 오류가 발생하였습니다.";
				frmMB.TopMost = true;
				frmMB.ShowDialog();
				System.Diagnostics.Debug.Print(ex.Message);
			}
			finally
			{
				xlsxBook.Close(true, newfileName, false);
				xlsxApp.Quit();

				releaseObject(xlsxApp);
                releaseObject(xlsxSheet1);
                releaseObject(xlsxSheet2);
				releaseObject(xlsxBook);

				xlsxApp = null;
				xlsxBook = null;

				// 보고서 생성 시 Process에 생성된 EXCEL.EXE을 Kill 함.
				System.Diagnostics.Process[] process = System.Diagnostics.Process.GetProcessesByName("EXCEL");
				for (int i = 0; i < process.Length; i++)
				{
					if (!aryProcessID.Contains(process[i].Id))
						process[i].Kill();
				}
			}

			return boolResult;
		}

        /// <summary>
        /// 항목별 타이틀 위치
        /// </summary>
        /// <param name="strItemRefDC"></param>
        /// <param name="strItemAveDC"></param>
        /// <param name="strItemRefAC"></param>
        /// <param name="strItemAveAC"></param>
        /// <param name="strItemRefL"></param>
        /// <param name="strItemAveL"></param>
        /// <param name="strItemRefC"></param>
        /// <param name="strItemAveC"></param>
        /// <param name="strItemRefQ"></param>
        /// <param name="strItemAveQ"></param>
        /// <param name="strItemResultA"></param>
        /// <param name="strItemResultB"></param>
        private void GetItemTitleName(ref string strItemRefDC, ref string strItemAveDC, ref string strItemRefAC, ref string strItemAveAC
            , ref string strItemRefL, ref string strItemAveL, ref string strItemRefC, ref string strItemAveC, ref string strItemRefQ
            , ref string strItemAveQ, ref string strItemResultA, ref string strItemResultB)
        {
            // Rdc
            if (chkRdc.Checked)
            {
                strItemRefDC = "D";
                strItemAveDC = "E";
            }
            else
            {
                strItemRefDC = "";
                strItemAveDC = "";
            }

            // Rac
            if (chkRac.Checked)
            {
                if (strItemRefDC != "")
                {
                    strItemRefAC = "F";
                    strItemAveAC = "G";
                }
                else
                {
                    strItemRefAC = "D";
                    strItemAveAC = "E";
                }
            }
            else
            {
                strItemRefAC = "";
                strItemAveAC = "";
            }

            // L
            if (chkL.Checked)
            {
                if (strItemRefDC != "" && strItemRefAC != "")
                {
                    strItemRefL = "H";
                    strItemAveL = "I";
                }
                else if ((strItemRefDC == "" && strItemRefAC != "") || (strItemRefDC != "" && strItemRefAC == ""))
                {
                    strItemRefL = "F";
                    strItemAveL = "G";
                }
                else
                {
                    strItemRefL = "D";
                    strItemAveL = "E";
                }
            }
            else
            {
                strItemRefL = "";
                strItemAveL = "";
            }

            // C
            if (chkC.Checked)
            {
                if (strItemRefDC != "" && strItemRefAC != "" && strItemRefL != "")
                {
                    strItemRefC = "J";
                    strItemAveC = "K";
                }
                else if ((strItemRefDC == "" && strItemRefAC != "" && strItemRefL != "") || (strItemRefDC != "" && strItemRefAC == "" && strItemRefL != "")
                    || (strItemRefDC != "" && strItemRefAC != "" && strItemRefL == ""))
                {
                    strItemRefC = "H";
                    strItemAveC = "I";
                }
                else if ((strItemRefDC == "" && strItemRefAC == "" && strItemRefL != "") || (strItemRefDC == "" && strItemRefAC != "" && strItemRefL == "")
                    || (strItemRefDC != "" && strItemRefAC == "" && strItemRefL == ""))
                {
                    strItemRefC = "F";
                    strItemAveC = "G";
                }
                else
                {
                    strItemRefC = "D";
                    strItemAveC = "E";
                }
            }
            else
            {
                strItemRefC = "";
                strItemAveC = "";
            }

            // Q
            if (chkQ.Checked)
            {
                if (strItemRefDC != "" && strItemRefAC != "" && strItemRefL != "" && strItemRefC != "")
                {
                    strItemRefQ = "L";
                    strItemAveQ = "M";
                }
                else if ((strItemRefDC == "" && strItemRefAC != "" && strItemRefL != "" && strItemRefC != "")
                    || (strItemRefDC != "" && strItemRefAC == "" && strItemRefL != "" && strItemRefC != "")
                    || (strItemRefDC != "" && strItemRefAC != "" && strItemRefL == "" && strItemRefC != "")
                    || (strItemRefDC != "" && strItemRefAC != "" && strItemRefL != "" && strItemRefC == ""))
                {
                    strItemRefQ = "J";
                    strItemAveQ = "K";
                }
                else if ((strItemRefDC == "" && strItemRefAC == "" && strItemRefL != "" && strItemRefC != "")
                    || (strItemRefDC == "" && strItemRefAC != "" && strItemRefL == "" && strItemRefC != "")
                    || (strItemRefDC == "" && strItemRefAC != "" && strItemRefL != "" && strItemRefC == "")
                    || (strItemRefDC != "" && strItemRefAC == "" && strItemRefL == "" && strItemRefC != "")
                    || (strItemRefDC != "" && strItemRefAC == "" && strItemRefL != "" && strItemRefC == "")
                    || (strItemRefDC != "" && strItemRefAC != "" && strItemRefL == "" && strItemRefC == ""))
                {
                    strItemRefQ = "H";
                    strItemAveQ = "I";
                }
                else if ((strItemRefDC == "" && strItemRefAC == "" && strItemRefL == "" && strItemRefC != "")
                    || (strItemRefDC == "" && strItemRefAC == "" && strItemRefL != "" && strItemRefC == "")
                    || (strItemRefDC == "" && strItemRefAC != "" && strItemRefL == "" && strItemRefC == "")
                    || (strItemRefDC != "" && strItemRefAC == "" && strItemRefL == "" && strItemRefC == ""))
                {
                    strItemRefQ = "F";
                    strItemAveQ = "G";
                }
                else
                {
                    strItemRefQ = "D";
                    strItemAveQ = "E";
                }
            }
            else
            {
                strItemRefQ = "";
                strItemAveQ = "";
            }

            // 기타
            if (strItemRefDC != "" && strItemRefAC != "" && strItemRefL != "" && strItemRefC != "" && strItemRefQ != "")
            {
                strItemResultA = "N";
                strItemResultB = "O";
            }
            else if ((strItemRefDC == "" && strItemRefAC != "" && strItemRefL != "" && strItemRefC != "" && strItemRefQ != "")
                || (strItemRefDC != "" && strItemRefAC == "" && strItemRefL != "" && strItemRefC != "" && strItemRefQ != "")
                || (strItemRefDC != "" && strItemRefAC != "" && strItemRefL == "" && strItemRefC != "" && strItemRefQ != "")
                || (strItemRefDC != "" && strItemRefAC != "" && strItemRefL != "" && strItemRefC == "" && strItemRefQ != "")
                || (strItemRefDC != "" && strItemRefAC != "" && strItemRefL != "" && strItemRefC != "" && strItemRefQ == ""))
            {
                strItemResultA = "L";
                strItemResultB = "O";
            }
            else if ((strItemRefDC == "" && strItemRefAC == "" && strItemRefL != "" && strItemRefC != "" && strItemRefQ != "")
                || (strItemRefDC == "" && strItemRefAC != "" && strItemRefL == "" && strItemRefC != "" && strItemRefQ != "")
                || (strItemRefDC == "" && strItemRefAC != "" && strItemRefL != "" && strItemRefC == "" && strItemRefQ != "")
                || (strItemRefDC == "" && strItemRefAC != "" && strItemRefL != "" && strItemRefC != "" && strItemRefQ == "")
                || (strItemRefDC != "" && strItemRefAC == "" && strItemRefL == "" && strItemRefC != "" && strItemRefQ != "")
                || (strItemRefDC != "" && strItemRefAC == "" && strItemRefL != "" && strItemRefC == "" && strItemRefQ != "")
                || (strItemRefDC != "" && strItemRefAC == "" && strItemRefL != "" && strItemRefC != "" && strItemRefQ == "")
                || (strItemRefDC != "" && strItemRefAC != "" && strItemRefL == "" && strItemRefC == "" && strItemRefQ != "")
                || (strItemRefDC != "" && strItemRefAC != "" && strItemRefL == "" && strItemRefC != "" && strItemRefQ == "")
                || (strItemRefDC != "" && strItemRefAC != "" && strItemRefL != "" && strItemRefC == "" && strItemRefQ == ""))
            {
                strItemResultA = "J";
                strItemResultB = "O";
            }
            else if ((strItemRefDC == "" && strItemRefAC == "" && strItemRefL == "" && strItemRefC != "" && strItemRefQ != "")
                || (strItemRefDC == "" && strItemRefAC == "" && strItemRefL != "" && strItemRefC == "" && strItemRefQ != "")
                || (strItemRefDC == "" && strItemRefAC == "" && strItemRefL != "" && strItemRefC != "" && strItemRefQ == "")
                || (strItemRefDC == "" && strItemRefAC != "" && strItemRefL == "" && strItemRefC == "" && strItemRefQ != "")
                || (strItemRefDC == "" && strItemRefAC != "" && strItemRefL == "" && strItemRefC != "" && strItemRefQ == "")
                || (strItemRefDC == "" && strItemRefAC != "" && strItemRefL != "" && strItemRefC == "" && strItemRefQ == "")
                || (strItemRefDC != "" && strItemRefAC == "" && strItemRefL == "" && strItemRefC == "" && strItemRefQ != "")
                || (strItemRefDC != "" && strItemRefAC == "" && strItemRefL == "" && strItemRefC != "" && strItemRefQ == "")
                || (strItemRefDC != "" && strItemRefAC == "" && strItemRefL != "" && strItemRefC == "" && strItemRefQ == "")
                || (strItemRefDC != "" && strItemRefAC != "" && strItemRefL == "" && strItemRefC == "" && strItemRefQ == ""))
            {
                strItemResultA = "H";
                strItemResultB = "O";
            }
            else if ((strItemRefDC == "" && strItemRefAC == "" && strItemRefL == "" && strItemRefC == "" && strItemRefQ != "")
                || (strItemRefDC == "" && strItemRefAC == "" && strItemRefL == "" && strItemRefC != "" && strItemRefQ == "")
                || (strItemRefDC == "" && strItemRefAC == "" && strItemRefL != "" && strItemRefC == "" && strItemRefQ == "")
                || (strItemRefDC == "" && strItemRefAC != "" && strItemRefL == "" && strItemRefC == "" && strItemRefQ == "")
                || (strItemRefDC != "" && strItemRefAC == "" && strItemRefL == "" && strItemRefC == "" && strItemRefQ == ""))
            {
                strItemResultA = "F";
                strItemResultB = "O";
            }
            else
            {
                strItemResultA = "D";
                strItemResultB = "O";
            }
        }

		#region 메모리해제

		private static void releaseObject(object obj)
		{
			try
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
				obj = null;
			}
			catch (Exception ex)
			{
				obj = null;
				System.Diagnostics.Debug.Print(ex.Message);
			}
			finally
			{
				GC.Collect();
			}
		}

        #endregion

        /// <summary>
        /// 측정 대상 Checked Changed Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void chkMeasurementTarget_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                switch (((CheckBox)sender).Name.Trim())
                {
                    case "chkRdc":
                        if (((CheckBox)sender).Checked)
                        {
                            dgvMeasurement.Columns["DC_ResistanceValue"].Visible = true;
                            dgvMeasurement.Columns["DC_Deviation"].Visible = true;
                            chartMeasurementValueDC.Visible = true;
                            chartAverageValueDC.Visible = true;
                        }
                        else
                        {
                            dgvMeasurement.Columns["DC_ResistanceValue"].Visible = false;
                            dgvMeasurement.Columns["DC_Deviation"].Visible = false;
                            chartMeasurementValueDC.Visible = false;
                            chartAverageValueDC.Visible = false;

                        }
                        break;
                    case "chkRac":
                        if (((CheckBox)sender).Checked)
                        {
                            dgvMeasurement.Columns["AC_ResistanceValue"].Visible = true;
                            dgvMeasurement.Columns["AC_Deviation"].Visible = true;
                            chartMeasurementValueAC.Visible = true;
                            chartAverageValueAC.Visible = true;
                        }
                        else
                        {
                            dgvMeasurement.Columns["AC_ResistanceValue"].Visible = false;
                            dgvMeasurement.Columns["AC_Deviation"].Visible = false;
                            chartMeasurementValueAC.Visible = false;
                            chartAverageValueAC.Visible = false;
                        }
                        break;
                    case "chkL":
                        if (((CheckBox)sender).Checked)
                        {
                            dgvMeasurement.Columns["L_InductanceValue"].Visible = true;
                            dgvMeasurement.Columns["L_Deviation"].Visible = true;
                            chartMeasurementValueL.Visible = true;
                            chartAverageValueL.Visible = true;
                        }
                        else
                        {
                            dgvMeasurement.Columns["L_InductanceValue"].Visible = false;
                            dgvMeasurement.Columns["L_Deviation"].Visible = false;
                            chartMeasurementValueL.Visible = false;
                            chartAverageValueL.Visible = false;
                        }
                        break;
                    case "chkC":
                        if (((CheckBox)sender).Checked)
                        {
                            dgvMeasurement.Columns["C_CapacitanceValue"].Visible = true;
                            dgvMeasurement.Columns["C_Deviation"].Visible = true;
                            chartMeasurementValueC.Visible = true;
                            chartAverageValueC.Visible = true;
                        }
                        else
                        {
                            dgvMeasurement.Columns["C_CapacitanceValue"].Visible = false;
                            dgvMeasurement.Columns["C_Deviation"].Visible = false;
                            chartMeasurementValueC.Visible = false;
                            chartAverageValueC.Visible = false;
                        }
                        break;
                    case "chkQ":
                        if (((CheckBox)sender).Checked)
                        {
                            dgvMeasurement.Columns["Q_FactorValue"].Visible = true;
                            dgvMeasurement.Columns["Q_Deviation"].Visible = true;
                            chartMeasurementValueQ.Visible = true;
                            chartAverageValueQ.Visible = true;
                        }
                        else
                        {
                            dgvMeasurement.Columns["Q_FactorValue"].Visible = false;
                            dgvMeasurement.Columns["Q_Deviation"].Visible = false;
                            chartMeasurementValueQ.Visible = false;
                            chartAverageValueQ.Visible = false;
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
        }

        private void btnCheckBoxAllSelect33_Click(object sender, EventArgs e)
        {
            if (gbStopRod33.Enabled)
                chkStopRod33.Checked = true;

            if (gbControlRod33.Enabled)
                chkControlRod33.Checked = true;
        }

        private void btnCheckBoxAllClear33_Click(object sender, EventArgs e)
        {
            if (gbStopRod33.Enabled)
                chkStopRod33.Checked = false;

            if (gbControlRod33.Enabled)
                chkControlRod33.Checked = false;
        }

        private void chkControlRod_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cb = sender as CheckBox;
            string grBoxName = $"gb{cb.Name.Substring(3)}";
            var gb = this.Controls.Find(grBoxName, true).FirstOrDefault();

            foreach (Control rod in gb.Controls)
            {
                if (rod is CheckBox && rod.Name != cb.Name && rod.Visible)
                {
                    ((CheckBox)rod).Checked = cb.Checked;
                }
            }
        }
    }
}
