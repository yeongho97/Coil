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

namespace Coil_Diagnostor
{
    public partial class frmDRPIHistory : Form
    {
        protected Function.FunctionMeasureProcess m_MeasureProcess = new Function.FunctionMeasureProcess();
        protected Function.FunctionDataControl m_db = new Function.FunctionDataControl();
        protected Function.FunctionChart m_Chart = new Function.FunctionChart();
        protected frmMessageBox frmMB = new frmMessageBox();
        protected string strPlantName = "";
        protected bool boolFormLoad = false;
        protected string strReferenceHogi = "초기값";
        protected string strReferenceOHDegree = "제 0 차";

        protected decimal[] dA_ControlRodRdc = new decimal[3];    // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dA_ControlRodRac = new decimal[3];    // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dA_ControlRodL = new decimal[3];      // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dA_ControlRodC = new decimal[3];      // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dA_ControlRodQ = new decimal[3];      // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dB_ControlRodRdc = new decimal[3];    // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dB_ControlRodRac = new decimal[3];    // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dB_ControlRodL = new decimal[3];      // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dB_ControlRodC = new decimal[3];      // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dB_ControlRodQ = new decimal[3];      // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dA_StopRdc = new decimal[3];          // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dA_StopRac = new decimal[3];          // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dA_StopL = new decimal[3];            // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dA_StopC = new decimal[3];            // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dA_StopQ = new decimal[3];            // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dB_StopRdc = new decimal[3];          // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dB_StopRac = new decimal[3];          // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dB_StopL = new decimal[3];            // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dB_StopC = new decimal[3];            // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dB_StopQ = new decimal[3];            // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)

        protected decimal[] dA_ControlRodAveRdc = new decimal[3];    // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dA_ControlRodAveRac = new decimal[3];    // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dA_ControlRodAveL = new decimal[3];      // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dA_ControlRodAveC = new decimal[3];      // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dA_ControlRodAveQ = new decimal[3];      // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dB_ControlRodAveRdc = new decimal[3];    // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dB_ControlRodAveRac = new decimal[3];    // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dB_ControlRodAveL = new decimal[3];      // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dB_ControlRodAveC = new decimal[3];      // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dB_ControlRodAveQ = new decimal[3];      // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dA_StopAveRdc = new decimal[3];          // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dA_StopAveRac = new decimal[3];          // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dA_StopAveL = new decimal[3];            // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dA_StopAveC = new decimal[3];            // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dA_StopAveQ = new decimal[3];            // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dB_StopAveRdc = new decimal[3];          // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dB_StopAveRac = new decimal[3];          // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dB_StopAveL = new decimal[3];            // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dB_StopAveC = new decimal[3];            // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        protected decimal[] dB_StopAveQ = new decimal[3];            // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)

        System.Windows.Forms.DataVisualization.Charting.Series sSeries = new System.Windows.Forms.DataVisualization.Charting.Series();

        public frmDRPIHistory()
        {
            CheckForIllegalCrossThreadCalls = false;
            InitializeComponent();

            for (int i = 0; i < 3; i++)
            {
                dA_ControlRodRdc[i] = 0.000M;
                dA_ControlRodRac[i] = 0.000M;
                dA_ControlRodL[i] = 0.000M;
                dA_ControlRodC[i] = 0.000000M;
                dA_ControlRodQ[i] = 0.000M;

                dB_ControlRodRdc[i] = 0.000M;
                dB_ControlRodRac[i] = 0.000M;
                dB_ControlRodL[i] = 0.000M;
                dB_ControlRodC[i] = 0.000000M;
                dB_ControlRodQ[i] = 0.000M;

                dA_StopRdc[i] = 0.000M;
                dA_StopRac[i] = 0.000M;
                dA_StopL[i] = 0.000M;
                dA_StopC[i] = 0.000000M;
                dA_StopQ[i] = 0.000M;

                dB_StopRdc[i] = 0.000M;
                dB_StopRac[i] = 0.000M;
                dB_StopL[i] = 0.000M;
                dB_StopC[i] = 0.000000M;
                dB_StopQ[i] = 0.000M;

                dA_ControlRodAveRdc[i] = 0.000M;
                dA_ControlRodAveRac[i] = 0.000M;
                dA_ControlRodAveL[i] = 0.000M;
                dA_ControlRodAveC[i] = 0.000000M;
                dA_ControlRodAveQ[i] = 0.000M;

                dB_ControlRodAveRdc[i] = 0.000M;
                dB_ControlRodAveRac[i] = 0.000M;
                dB_ControlRodAveL[i] = 0.000M;
                dB_ControlRodAveC[i] = 0.000000M;
                dB_ControlRodAveQ[i] = 0.000M;

                dA_StopAveRdc[i] = 0.000M;
                dA_StopAveRac[i] = 0.000M;
                dA_StopAveL[i] = 0.000M;
                dA_StopAveC[i] = 0.000000M;
                dA_StopAveQ[i] = 0.000M;

                dB_StopAveRdc[i] = 0.000M;
                dB_StopAveRac[i] = 0.000M;
                dB_StopAveL[i] = 0.000M;
                dB_StopAveC[i] = 0.000000M;
                dB_StopAveQ[i] = 0.000M;
            }

            teA1RdcControlRod_DRPIReferenceValue.Text = "0.000";
            teA1RacControlRod_DRPIReferenceValue.Text = "0.000";
            teA1LControlRod_DRPIReferenceValue.Text = "0.000";
            teA1CControlRod_DRPIReferenceValue.Text = "0.000000";
            teA1QControlRod_DRPIReferenceValue.Text = "0.000";

            teA1RdcStop_DRPIReferenceValue.Text = "0.000";
            teA1RacStop_DRPIReferenceValue.Text = "0.000";
            teA1LStop_DRPIReferenceValue.Text = "0.000";
            teA1CStop_DRPIReferenceValue.Text = "0.000000";
            teA1QStop_DRPIReferenceValue.Text = "0.000";

            teA1RdcControlRod_DRPIAverageValue.Text = "0.000";
            teA1RacControlRod_DRPIAverageValue.Text = "0.000";
            teA1LControlRod_DRPIAverageValue.Text = "0.000";
            teA1CControlRod_DRPIAverageValue.Text = "0.000000";
            teA1QControlRod_DRPIAverageValue.Text = "0.000";

            teA1RdcStop_DRPIAverageValue.Text = "0.000";
            teA1RacStop_DRPIAverageValue.Text = "0.000";
            teA1LStop_DRPIAverageValue.Text = "0.000";
            teA1CStop_DRPIAverageValue.Text = "0.000000";
            teA1QStop_DRPIAverageValue.Text = "0.000";

            teB2RdcControlRod_DRPIReferenceValue.Text = "0.000";
            teB2RacControlRod_DRPIReferenceValue.Text = "0.000";
            teB2LControlRod_DRPIReferenceValue.Text = "0.000";
            teB2CControlRod_DRPIReferenceValue.Text = "0.000000";
            teB2QControlRod_DRPIReferenceValue.Text = "0.000";

            teB2RdcStop_DRPIReferenceValue.Text = "0.000";
            teB2RacStop_DRPIReferenceValue.Text = "0.000";
            teB2LStop_DRPIReferenceValue.Text = "0.000";
            teB2CStop_DRPIReferenceValue.Text = "0.000000";
            teB2QStop_DRPIReferenceValue.Text = "0.000";

            teB2RdcControlRod_DRPIAverageValue.Text = "0.000";
            teB2RacControlRod_DRPIAverageValue.Text = "0.000";
            teB2LControlRod_DRPIAverageValue.Text = "0.000";
            teB2CControlRod_DRPIAverageValue.Text = "0.000000";
            teB2QControlRod_DRPIAverageValue.Text = "0.000";

            teB2RdcStop_DRPIAverageValue.Text = "0.000";
            teB2RacStop_DRPIAverageValue.Text = "0.000";
            teB2LStop_DRPIAverageValue.Text = "0.000";
            teB2CStop_DRPIAverageValue.Text = "0.000000";
            teB2QStop_DRPIAverageValue.Text = "0.000";

            teB3RdcControlRod_DRPIReferenceValue.Text = "0.000";
            teB3RacControlRod_DRPIReferenceValue.Text = "0.000";
            teB3LControlRod_DRPIReferenceValue.Text = "0.000";
            teB3CControlRod_DRPIReferenceValue.Text = "0.000000";
            teB3QControlRod_DRPIReferenceValue.Text = "0.000";

            teB3RdcStop_DRPIReferenceValue.Text = "0.000";
            teB3RacStop_DRPIReferenceValue.Text = "0.000";
            teB3LStop_DRPIReferenceValue.Text = "0.000";
            teB3CStop_DRPIReferenceValue.Text = "0.000000";
            teB3QStop_DRPIReferenceValue.Text = "0.000";

            teB3RdcControlRod_DRPIAverageValue.Text = "0.000";
            teB3RacControlRod_DRPIAverageValue.Text = "0.000";
            teB3LControlRod_DRPIAverageValue.Text = "0.000";
            teB3CControlRod_DRPIAverageValue.Text = "0.000000";
            teB3QControlRod_DRPIAverageValue.Text = "0.000";

            teB3RdcStop_DRPIAverageValue.Text = "0.000";
            teB3RacStop_DRPIAverageValue.Text = "0.000";
            teB3LStop_DRPIAverageValue.Text = "0.000";
            teB3CStop_DRPIAverageValue.Text = "0.000000";
            teB3QStop_DRPIAverageValue.Text = "0.000";

            teB1RdcControlRod_DRPIReferenceValue.Text = "0.000";
            teB1RacControlRod_DRPIReferenceValue.Text = "0.000";
            teB1LControlRod_DRPIReferenceValue.Text = "0.000";
            teB1CControlRod_DRPIReferenceValue.Text = "0.000000";
            teB1QControlRod_DRPIReferenceValue.Text = "0.000";

            teB1RdcStop_DRPIReferenceValue.Text = "0.000";
            teB1RacStop_DRPIReferenceValue.Text = "0.000";
            teB1LStop_DRPIReferenceValue.Text = "0.000";
            teB1CStop_DRPIReferenceValue.Text = "0.000000";
            teB1QStop_DRPIReferenceValue.Text = "0.000";

            teB1RdcControlRod_DRPIAverageValue.Text = "0.000";
            teB1RacControlRod_DRPIAverageValue.Text = "0.000";
            teB1LControlRod_DRPIAverageValue.Text = "0.000";
            teB1CControlRod_DRPIAverageValue.Text = "0.000000";
            teB1QControlRod_DRPIAverageValue.Text = "0.000";

            teB1RdcStop_DRPIAverageValue.Text = "0.000";
            teB1RacStop_DRPIAverageValue.Text = "0.000";
            teB1LStop_DRPIAverageValue.Text = "0.000";
            teB1CStop_DRPIAverageValue.Text = "0.000000";
            teB1QStop_DRPIAverageValue.Text = "0.000";

            teB2RdcControlRod_DRPIReferenceValue.Text = "0.000";
            teB2RacControlRod_DRPIReferenceValue.Text = "0.000";
            teB2LControlRod_DRPIReferenceValue.Text = "0.000";
            teB2CControlRod_DRPIReferenceValue.Text = "0.000000";
            teB2QControlRod_DRPIReferenceValue.Text = "0.000";

            teB2RdcStop_DRPIReferenceValue.Text = "0.000";
            teB2RacStop_DRPIReferenceValue.Text = "0.000";
            teB2LStop_DRPIReferenceValue.Text = "0.000";
            teB2CStop_DRPIReferenceValue.Text = "0.000000";
            teB2QStop_DRPIReferenceValue.Text = "0.000";

            teB2RdcControlRod_DRPIAverageValue.Text = "0.000";
            teB2RacControlRod_DRPIAverageValue.Text = "0.000";
            teB2LControlRod_DRPIAverageValue.Text = "0.000";
            teB2CControlRod_DRPIAverageValue.Text = "0.000000";
            teB2QControlRod_DRPIAverageValue.Text = "0.000";

            teB2RdcStop_DRPIAverageValue.Text = "0.000";
            teB2RacStop_DRPIAverageValue.Text = "0.000";
            teB2LStop_DRPIAverageValue.Text = "0.000";
            teB2CStop_DRPIAverageValue.Text = "0.000000";
            teB2QStop_DRPIAverageValue.Text = "0.000";

            teB3RdcControlRod_DRPIReferenceValue.Text = "0.000";
            teB3RacControlRod_DRPIReferenceValue.Text = "0.000";
            teB3LControlRod_DRPIReferenceValue.Text = "0.000";
            teB3CControlRod_DRPIReferenceValue.Text = "0.000000";
            teB3QControlRod_DRPIReferenceValue.Text = "0.000";

            teB3RdcStop_DRPIReferenceValue.Text = "0.000";
            teB3RacStop_DRPIReferenceValue.Text = "0.000";
            teB3LStop_DRPIReferenceValue.Text = "0.000";
            teB3CStop_DRPIReferenceValue.Text = "0.000000";
            teB3QStop_DRPIReferenceValue.Text = "0.000";

            teB3RdcControlRod_DRPIAverageValue.Text = "0.000";
            teB3RacControlRod_DRPIAverageValue.Text = "0.000";
            teB3LControlRod_DRPIAverageValue.Text = "0.000";
            teB3CControlRod_DRPIAverageValue.Text = "0.000000";
            teB3QControlRod_DRPIAverageValue.Text = "0.000";

            teB3RdcStop_DRPIAverageValue.Text = "0.000";
            teB3RacStop_DRPIAverageValue.Text = "0.000";
            teB3LStop_DRPIAverageValue.Text = "0.000";
            teB3CStop_DRPIAverageValue.Text = "0.000000";
            teB3QStop_DRPIAverageValue.Text = "0.000";
        }

        /// <summary>
        /// 폼 단축키 지정
        /// </summary>
        protected override bool ProcessCmdKey(ref System.Windows.Forms.Message msg, Keys keyData)
        {
            Keys key = keyData & ~(Keys.Shift | Keys.Control);

            switch (key)
            {
                case Keys.F2: // 조회 버튼
                    btnSearch.PerformClick();
                    break;
                case Keys.F3: // 기준값 표시/숨김 버튼
                    btnReferenceValue.PerformClick();                    
                    break;
                case Keys.F4: // 평균값 표시/숨김 버튼
                    btnAverageValue.PerformClick();
                    break;
                case Keys.F5: // 기준 값 설정 버튼
                    btnSearch.PerformClick();
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
        private void frmDRPIHistory_Load(object sender, EventArgs e)
        {
            strPlantName = Gini.GetValue("Device", "PlantName").Trim();

            // ComboBox 데이터 설정
            SetComboBoxDataBinding();

            // 표형식 그리드 초기 설정
            SetDataGridViewInitialize();

            // A 그룹 제어용 평균치 그리드 초기 설정
            SetControlAverageDataGridViewInitialize(dgvControlAverageA);

            // B 그룹 제어용 평균치 그리드 초기 설정
            SetControlAverageDataGridViewInitialize(dgvControlAverageB);

            // A 그룹 정지용 평균치 그리드 초기 설정
            SetStopAverageDataGridViewInitialize(dgvStopAverageA);

            // B 그룹 정지용 평균치 그리드 초기 설정
            SetStopAverageDataGridViewInitialize(dgvStopAverageB);

            // 정지,올림,이동 label 기준값 Visible 설정
            SetReferenceValueLabelVisible(false);

            // 정지,올림,이동 label 기준값 Visible 설정
            SetReferenceValueLabelVisible(false);

            tabControl.SelectedIndex = 0;
            if (tabControl.SelectedIndex == 1 || tabControl.SelectedIndex == 2)
            {
                btnReferenceValue.Enabled = true;
                btnAverageValue.Enabled = true;
                btnScalingSetting.Enabled = false;
            }
            else
            {
                btnReferenceValue.Enabled = false;
                btnAverageValue.Enabled = false;

                // 기준값 저장(기준 값 Setting) 버튼 활성화/비활성화 설정
                if (tabControl.SelectedIndex == 5 || tabControl.SelectedIndex == 6)
                    btnScalingSetting.Enabled = true;
                else
                    btnScalingSetting.Enabled = false;
            }

            // Rdc, Rac 활성화 및 비활성화
            SetTabControlDCAC();

            // Rdc, Rac 활성화 및 비활성화
            SetTabControlLCQ();

            chkMeasurementTarget_CheckedChanged(chkC, null);
        }

        /// <summary>
        /// ComboBox 데이터 설정
        /// </summary>
        private void SetComboBoxDataBinding()
        {
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

            // 타입
            string[] strDRPITypeItem = Gini.GetValue("Combo", "DRPIReferenceValueLabelType").Split(',');

            cboType.Items.Clear();

            cboType.Items.Add("전체");

            for (int i = 0; i < strDRPITypeItem.Length; i++)
            {
                cboType.Items.Add(strDRPITypeItem[i].Trim());
            }

            cboType.SelectedIndex = 0;

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
                DataGridViewTextBoxColumn Column19 = new DataGridViewTextBoxColumn();

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
                Column10.Visible = true;
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

                Column19.HeaderText = "측정타입";
                Column19.Name = "DRPIType";
                Column19.Width = 60;
                Column19.ReadOnly = true;
                Column19.Visible = true;
                Column19.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column19.SortMode = DataGridViewColumnSortMode.NotSortable;

                dgvMeasurement.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] { Column1, Column2, Column3, Column4, Column5, Column6
                    , Column7, Column8, Column9, Column10, Column11, Column12, Column13, Column14, Column15, Column16, Column17, Column18, Column19 });

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

        /// <summary>
        /// 제어용 평균치 그리드 초기 설정
        /// </summary>
        private void SetControlAverageDataGridViewInitialize(DataGridView _dgv)
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

                Column1.HeaderText = "구분";
                Column1.Name = "CoilCabinetType";
                Column1.Width = 60;
                Column1.ReadOnly = true;
                Column1.Visible = true;
                Column1.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column1.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column2.HeaderText = "코일명";
                Column2.Name = "CoilCabinetName";
                Column2.Width = 60;
                Column2.ReadOnly = true;
                Column2.Visible = true;
                Column2.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column2.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column3.HeaderText = "DC (Ω)";
                Column3.Name = "Rdc_DRPIReferenceValue";
                Column3.Width = 100;
                Column3.ReadOnly = true;
                Column3.Visible = true;
                Column3.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column3.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column4.HeaderText = "AC (Ω)";
                Column4.Name = "Rac_DRPIReferenceValue";
                Column4.Width = 100;
                Column4.ReadOnly = true;
                Column4.Visible = true;
                Column4.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column4.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column5.HeaderText = "L (mH)";
                Column5.Name = "L_DRPIReferenceValue";
                Column5.Width = 100;
                Column5.ReadOnly = true;
                Column5.Visible = true;
                Column5.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column5.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column6.HeaderText = "C (uF)";
                Column6.Name = "C_DRPIReferenceValue";
                Column6.Width = 100;
                Column6.ReadOnly = true;
                Column6.Visible = true;
                Column6.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column6.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column7.HeaderText = "Q";
                Column7.Name = "Q_DRPIReferenceValue";
                Column7.Width = 100;
                Column7.ReadOnly = true;
                Column7.Visible = true;
                Column7.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column7.SortMode = DataGridViewColumnSortMode.NotSortable;

                _dgv.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] { Column1, Column2, Column3, Column4, Column5, Column6
                    , Column7 });

                _dgv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing;
                _dgv.ColumnHeadersHeight = 40;
                _dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                _dgv.RowHeadersWidth = 40;

                // DataGridView 기타 세팅
                _dgv.AllowUserToAddRows = false; // Row 추가 기능 미사용
                _dgv.AllowUserToDeleteRows = false; // Row 삭제 기능 미사용
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
        }

        /// <summary>
        /// 정지용 평균치 그리드 초기 설정
        /// </summary>
        private void SetStopAverageDataGridViewInitialize(DataGridView _dgv)
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

                Column1.HeaderText = "구분";
                Column1.Name = "CoilCabinetType";
                Column1.Width = 60;
                Column1.ReadOnly = true;
                Column1.Visible = true;
                Column1.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column1.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column2.HeaderText = "코일명";
                Column2.Name = "CoilCabinetName";
                Column2.Width = 60;
                Column2.ReadOnly = true;
                Column2.Visible = true;
                Column2.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column2.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column3.HeaderText = "DC (Ω)";
                Column3.Name = "Rdc_DRPIReferenceValue";
                Column3.Width = 97;
                Column3.ReadOnly = true;
                Column3.Visible = true;
                Column3.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column3.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column4.HeaderText = "AC (Ω)";
                Column4.Name = "Rac_DRPIReferenceValue";
                Column4.Width = 97;
                Column4.ReadOnly = true;
                Column4.Visible = true;
                Column4.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column4.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column5.HeaderText = "L (mH)";
                Column5.Name = "L_DRPIReferenceValue";
                Column5.Width = 97;
                Column5.ReadOnly = true;
                Column5.Visible = true;
                Column5.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column5.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column6.HeaderText = "C (uF)";
                Column6.Name = "C_DRPIReferenceValue";
                Column6.Width = 97;
                Column6.ReadOnly = true;
                Column6.Visible = true;
                Column6.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column6.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column7.HeaderText = "Q";
                Column7.Name = "Q_DRPIReferenceValue";
                Column7.Width = 97;
                Column7.ReadOnly = true;
                Column7.Visible = true;
                Column7.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column7.SortMode = DataGridViewColumnSortMode.NotSortable;

                _dgv.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] { Column1, Column2, Column3, Column4, Column5, Column6
                    , Column7 });

                _dgv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing;
                _dgv.ColumnHeadersHeight = 40;
                _dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                _dgv.RowHeadersWidth = 40;

                // DataGridView 기타 세팅
                _dgv.AllowUserToAddRows = false; // Row 추가 기능 미사용
                _dgv.AllowUserToDeleteRows = false; // Row 삭제 기능 미사용
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
        }

        /// <summary>
        /// Tab Control Selected Index Changed Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void tabControl_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl.SelectedIndex == 1 || tabControl.SelectedIndex == 2)
            {
                btnReferenceValue.Enabled = true;
                btnAverageValue.Enabled = true;
                btnScalingSetting.Enabled = false;
            }
            else
            {
                btnReferenceValue.Enabled = false;
                btnAverageValue.Enabled = false;
                
                //// 기준값 저장(기준 값 Setting) 버튼 활성화/비활성화 설정
                //string strMeasurementType = cboMeasurementType.SelectedItem == null ? "" : cboMeasurementType.SelectedItem.ToString().Trim();

                //if (strMeasurementType.Trim() == "측정치")
                //{
                    if (tabControl.SelectedIndex == 5 || tabControl.SelectedIndex == 6)
                        btnScalingSetting.Enabled = true;
                    else
                        btnScalingSetting.Enabled = false;
                //}
                //else
                //{
                //    btnScalingSetting.Enabled = false;
                //}
            }
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
                    // 제어봉 ComboBox 바인딩
                    SetControlRodComboBoxBinding();

                    // 코일명 ComboBox 바인딩
                    SetCoilNameComboBoxBinding();
                }
                else
                {
                    // 제어봉 ComboBox 바인딩
                    SetControlRodComboBoxBinding();

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
                    cboOhDegree.SelectedItem = "제 " + Gini.GetValue("DRPI", "SelectDRPI_OHDegree") + " 차";

                if (cboOhDegree.SelectedItem == null) cboOhDegree.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
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
            cboType.SelectedIndex = 0;
        }

        /// <summary>
        /// 타입 Selected Index Changed Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cboType_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                // 그리드와 그래프 초기화
                SetDataGridViewAndGraphInitialize();

                // 제어봉 ComboBox 바인딩
                SetControlRodComboBoxBinding();

                // 코일명 ComboBox 바인딩
                SetCoilNameComboBoxBinding();
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
            if (cboType.SelectedItem == null || cboHistorySearchType.SelectedItem == null)
                return;

            try
            {
                string strDRPIControlItem = Gini.GetValue("DRPI", "Control_Item").Trim();
                string strDRPIStopItem = Gini.GetValue("DRPI", "Stop_Item").Trim();
                string[] arrayDRPIControlAndStop;

                if (cboHistorySearchType.SelectedItem.ToString().Trim() == "기간별")
                {
                    if (cboType.SelectedItem.ToString().Trim() == "전체" || cboType.SelectedItem.ToString().Trim() == "제어용")
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
                    if (cboType.SelectedItem.ToString().Trim() == "전체" || cboType.SelectedItem.ToString().Trim() == "제어용")
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
        /// 제어봉 ComboBox 바인딩
        /// </summary>
        private void SetControlRodComboBoxBinding()
        {
            if (cboType.SelectedItem == null || cboHistorySearchType.SelectedItem == null)
                return;

            try
            {
                string strDRPIControlRodColilNameItem = Gini.GetValue("Combo", "DRPIControlRodColilName_Item").Trim();
                string strDRPIControlRodStopNameItem = Gini.GetValue("Combo", "DRPIControlRodStopName_Item").Trim();
                string[] arrayDRPIControlRod;

                if (cboHistorySearchType.SelectedItem.ToString().Trim() == "기간별")
                {
                    if (cboType.SelectedItem.ToString().Trim() == "전체")
                    {
                        arrayDRPIControlRod = (strDRPIControlRodColilNameItem.Trim() + "," + strDRPIControlRodStopNameItem.Trim()).Split(',');
                    }
                    else if (cboType.SelectedItem.ToString().Trim() == "제어용")
                    {
                        arrayDRPIControlRod = strDRPIControlRodColilNameItem.Split(',');
                    }
                    else
                    {
                        arrayDRPIControlRod = strDRPIControlRodStopNameItem.Split(',');
                    }

                    // 제어봉
                    cboControlRod.Items.Clear();

                    for (int i = 0; i < arrayDRPIControlRod.Length; i++)
                    {
                        if (arrayDRPIControlRod[i].Trim() != "N/A")
                            cboControlRod.Items.Add(arrayDRPIControlRod[i].Trim());
                    }

                    cboControlRod.SelectedIndex = 0;
                }
                else
                {
                    if (cboType.SelectedItem.ToString().Trim() == "전체")
                    {
                        arrayDRPIControlRod = (strDRPIControlRodColilNameItem.Trim() + "," + strDRPIControlRodStopNameItem.Trim()).Split(',');
                    }
                    else if (cboType.SelectedItem.ToString().Trim() == "제어용")
                    {
                        arrayDRPIControlRod = strDRPIControlRodColilNameItem.Split(',');
                    }
                    else
                    {
                        arrayDRPIControlRod = strDRPIControlRodStopNameItem.Split(',');
                    }

                    // 제어봉
                    cboControlRod.Items.Clear();

                    cboControlRod.Items.Add("전체");

                    for (int i = 0; i < arrayDRPIControlRod.Length; i++)
                    {
                        if (arrayDRPIControlRod[i].Trim() != "N/A")
                            cboControlRod.Items.Add(arrayDRPIControlRod[i].Trim());
                    }

                    cboControlRod.SelectedIndex = 0;
                }
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
        /// 제어봉 Selected Index Changed Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cboControlRod_SelectedIndexChanged(object sender, EventArgs e)
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
        /// 기준값 표시/숨김 Button Click Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnReferenceValue_Click(object sender, EventArgs e)
        {
            try
            {
                if (btnReferenceValue.Text.Trim() == "기준값 표시 (F3)")
                {
                    btnReferenceValue.Text = string.Format("기준값 {0} (F3)", "숨김");

                    // 정지,올림,이동 label 기준값 Visible 설정
                    SetReferenceValueLabelVisible(true);
                }
                else
                {
                    btnReferenceValue.Text = string.Format("기준값 {0} (F3)", "표시");

                    // 정지,올림,이동 label 기준값 Visible 설정
                    SetReferenceValueLabelVisible(false);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
        }

        /// <summary>
        /// 평균값 표시/숨김 Button Click Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnAverageValue_Click(object sender, EventArgs e)
        {
            try
            {
                if (btnAverageValue.Text.Trim() == "평균값 표시 (F4)")
                {
                    btnAverageValue.Text = string.Format("평균값 {0} (F4)", "숨김");

                    // 정지,올림,이동 label 평균값 Visible 설정
                    SetAverageValueLabelVisible(true);
                }
                else
                {
                    btnAverageValue.Text = string.Format("평균값 {0} (F4)", "표시");

                    // 정지,올림,이동 label 평균값 Visible 설정
                    SetAverageValueLabelVisible(false);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
        }

        /// <summary>
        /// 정지,올림,이동 label 기준값 Visible 설정
        /// </summary>
        private void SetReferenceValueLabelVisible(bool _boolVisible)
        {
            if (chartMeasurementValueDC.Series.Count > 21)
                chartMeasurementValueDC.Series[21].Enabled = _boolVisible;

            if (chartMeasurementValueDC.Series.Count > 22)
                chartMeasurementValueDC.Series[22].Enabled = _boolVisible;

            if (chartMeasurementValueAC.Series.Count > 21)
                chartMeasurementValueAC.Series[21].Enabled = _boolVisible;

            if (chartMeasurementValueAC.Series.Count > 22)
                chartMeasurementValueAC.Series[22].Enabled = _boolVisible;

            if (chartMeasurementValueL.Series.Count > 21)
                chartMeasurementValueL.Series[21].Enabled = _boolVisible;

            if (chartMeasurementValueL.Series.Count > 22)
                chartMeasurementValueL.Series[22].Enabled = _boolVisible;

            if (chartMeasurementValueC.Series.Count > 21)
                chartMeasurementValueC.Series[21].Enabled = _boolVisible;

            if (chartMeasurementValueC.Series.Count > 22)
                chartMeasurementValueC.Series[22].Enabled = _boolVisible;

            if (chartMeasurementValueQ.Series.Count > 21)
                chartMeasurementValueQ.Series[21].Enabled = _boolVisible;

            if (chartMeasurementValueQ.Series.Count > 22)
                chartMeasurementValueQ.Series[22].Enabled = _boolVisible;
        }

        /// <summary>
        /// 정지,올림,이동 label 평균값 Visible 설정
        /// </summary>
        private void SetAverageValueLabelVisible(bool _boolVisible)
        {
            if (chartMeasurementValueDC.Series.Count > 23)
                chartMeasurementValueDC.Series[23].Enabled = _boolVisible;

            if (chartMeasurementValueDC.Series.Count > 24)
                chartMeasurementValueDC.Series[24].Enabled = _boolVisible;

            if (chartMeasurementValueAC.Series.Count > 23)
                chartMeasurementValueAC.Series[23].Enabled = _boolVisible;

            if (chartMeasurementValueAC.Series.Count > 24)
                chartMeasurementValueAC.Series[24].Enabled = _boolVisible;

            if (chartMeasurementValueL.Series.Count > 23)
                chartMeasurementValueL.Series[23].Enabled = _boolVisible;

            if (chartMeasurementValueL.Series.Count > 24)
                chartMeasurementValueL.Series[24].Enabled = _boolVisible;

            if (chartMeasurementValueC.Series.Count > 23)
                chartMeasurementValueC.Series[23].Enabled = _boolVisible;

            if (chartMeasurementValueC.Series.Count > 24)
                chartMeasurementValueC.Series[24].Enabled = _boolVisible;

            if (chartMeasurementValueQ.Series.Count > 23)
                chartMeasurementValueQ.Series[23].Enabled = _boolVisible;

            if (chartMeasurementValueQ.Series.Count > 24)
                chartMeasurementValueQ.Series[24].Enabled = _boolVisible;
        }

        /// <summary>
        /// 기준 값 Setting Button Evene
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnScalingSetting_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;

            bool boolSave = false;

            try
            {
                string strHogi = cboHogi.SelectedItem == null ? "" : cboHogi.SelectedItem.ToString().Trim();
                string strOhDegree = cboOhDegree.SelectedItem == null ? "" : cboOhDegree.SelectedItem.ToString().Trim();

                DataGridView dgvControl;
                DataGridView dgvStop;

                if (tabControl.SelectedIndex == 5)  // 그룹 A
                {
                    dgvControl = dgvControlAverageA;
                    dgvStop = dgvStopAverageA;

                    // 제어용 기준 값 저장
                    if (dgvControlAverageA.Rows.Count > 0)
                    {
                        boolSave = SaveControlReferenceValue(dgvControlAverageA, strHogi, strOhDegree, "제어용", "A");
                    }

                    // 정지용 기준 값 저장
                    if (dgvStopAverageA.Rows.Count > 0)
                    {
                        boolSave = SaveControlReferenceValue(dgvStopAverageA, strHogi, strOhDegree, "정지용", "A");
                    }
                }
                else  // 그룹 B
                {
                    dgvControl = dgvControlAverageB;
                    dgvStop = dgvStopAverageB;

                    // 제어용 기준 값 저장
                    if (dgvControlAverageB.Rows.Count > 0)
                    {
                        boolSave = SaveControlReferenceValue(dgvControlAverageB, strHogi, strOhDegree, "제어용", "B");
                    }

                    // 정지용 기준 값 저장
                    if (dgvStopAverageB.Rows.Count > 0)
                    {
                        boolSave = SaveControlReferenceValue(dgvStopAverageB, strHogi, strOhDegree, "정지용", "B");
                    }
                }

                if (boolSave)
                {
                    // 기준 값(DRPI) 가져오기
                    GetDRPIReferenceValueSelected();

                    // 표 형식 결과 재판단 및 업데이터
                    SetDataGridViewUpdate();

                    frmMB.lblMessage.Text = "기준 값 설정이 완료";
                    frmMB.TopMost = true;
                    frmMB.ShowDialog();
                }
                else if (dgvControl.Rows.Count < 1 && dgvStop.Rows.Count < 1)
                {
                    frmMB.lblMessage.Text = "기준 값 설정할 데이터가 없습니다.";
                    frmMB.TopMost = true;
                    frmMB.ShowDialog();
                }
                else
                {
                    frmMB.lblMessage.Text = "기준 값 설정 중 실패";
                    frmMB.TopMost = true;
                    frmMB.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                frmMB.lblMessage.Text = "기준 값 설정 중 오류";
                frmMB.TopMost = true;
                frmMB.ShowDialog();
                System.Diagnostics.Debug.Print(ex.Message);
            }

            System.Windows.Forms.Cursor.Current = Cursors.Default;
        }

        /// <summary>
        /// 표 형식 결과 재판단 및 업데이터
        /// </summary>
        private void SetDataGridViewUpdate()
        {
            try
            {
                decimal dRdc = 0, dRac = 0M, dL = 0M, dC = 0M, dQ = 0M;
                decimal dcmDecisionRangeMaxValue = 0M, dcmEffectiveStandardRangeMaxValue = 0M, dcmDecisionRangeMinValue = 0M
                    , dcmEffectiveStandardRangeMinValue = 0M;

                decimal dcmRefRdcValue = 0M, dcmRefRacValue = 0M, dcmRefLValue = 0M, dcmRefCValue = 0M, dcmRefQValue = 0M;

                decimal dcmDecisionRange_Rdc = 0M, dcmDecisionRange_Rac = 0M, dcmDecisionRange_L = 0M, dcmDecisionRange_C = 0M
                    , dcmDecisionRange_Q = 0M, dcmEffectiveStandardRange = 0M;

                dcmDecisionRange_Rdc = Convert.ToDecimal(Gini.GetValue("ReferenceValue", "RdcDecisionRange_ReferenceValue"));
                dcmDecisionRange_Rac = Convert.ToDecimal(Gini.GetValue("ReferenceValue", "RacDecisionRange_ReferenceValue"));
                dcmDecisionRange_L = Convert.ToDecimal(Gini.GetValue("ReferenceValue", "LDecisionRange_ReferenceValue"));
                dcmDecisionRange_C = Convert.ToDecimal(Gini.GetValue("ReferenceValue", "CDecisionRange_ReferenceValue"));
                dcmDecisionRange_Q = Convert.ToDecimal(Gini.GetValue("ReferenceValue", "QDecisionRange_ReferenceValue"));
                dcmEffectiveStandardRange = Convert.ToDecimal(Gini.GetValue("ReferenceValue", "EffectiveStandardRangeOfVariation"));

                string strResult = "";

                for (int i = 0; i < dgvMeasurement.RowCount; i++)
                {
                    string strHogi = cboHogi.SelectedItem == null ? "" : cboHogi.SelectedItem.ToString().Trim();
                    string strOhDegree = cboOhDegree.SelectedItem == null ? "" : cboOhDegree.SelectedItem.ToString().Trim();
                    string strDRPIGroup = dgvMeasurement.Rows[i].Cells["DRPIGroup"].Value == null ? "" : dgvMeasurement.Rows[i].Cells["DRPIGroup"].Value.ToString().Trim();
                    string strDRPIType = dgvMeasurement.Rows[i].Cells["DRPIType"].Value == null ? "" : dgvMeasurement.Rows[i].Cells["DRPIType"].Value.ToString().Trim();
                    string strControlName = dgvMeasurement.Rows[i].Cells["ControlName"].Value == null ? "" : dgvMeasurement.Rows[i].Cells["ControlName"].Value.ToString().Trim();
                    string strCoilName = dgvMeasurement.Rows[i].Cells["CoilName"].Value == null ? "" : dgvMeasurement.Rows[i].Cells["CoilName"].Value.ToString().Trim();
                    int intCoilNumber = dgvMeasurement.Rows[i].Cells["CoilNumber"].Value == null || dgvMeasurement.Rows[i].Cells["CoilNumber"].Value.ToString().Trim() == ""
                        ? 0 : Convert.ToInt32(dgvMeasurement.Rows[i].Cells["CoilNumber"].Value.ToString().Trim());

                    if (tabControl.SelectedIndex == 5)
                    {
                        if (strDRPIGroup.Trim() == "B") continue;
                        if (dgvControlAverageA.RowCount < 1) continue;
                        if (dgvStopAverageA.RowCount < 1) continue;
                    }
                    else if (tabControl.SelectedIndex == 6)
                    {
                        if (strDRPIGroup.Trim() == "A") continue;
                        if (dgvControlAverageB.RowCount < 1) continue;
                        if (dgvStopAverageB.RowCount < 1) continue;
                    }

                    dRdc = dgvMeasurement.Rows[i].Cells["DC_ResistanceValue"].Value == null || dgvMeasurement.Rows[i].Cells["DC_ResistanceValue"].Value.ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(dgvMeasurement.Rows[i].Cells["DC_ResistanceValue"].Value.ToString().Trim());
                    dRac = dgvMeasurement.Rows[i].Cells["AC_ResistanceValue"].Value == null || dgvMeasurement.Rows[i].Cells["AC_ResistanceValue"].Value.ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(dgvMeasurement.Rows[i].Cells["AC_ResistanceValue"].Value.ToString().Trim());
                    dL = dgvMeasurement.Rows[i].Cells["L_InductanceValue"].Value == null || dgvMeasurement.Rows[i].Cells["L_InductanceValue"].Value.ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(dgvMeasurement.Rows[i].Cells["L_InductanceValue"].Value.ToString().Trim());
                    dC = dgvMeasurement.Rows[i].Cells["C_CapacitanceValue"].Value == null || dgvMeasurement.Rows[i].Cells["C_CapacitanceValue"].Value.ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(dgvMeasurement.Rows[i].Cells["C_CapacitanceValue"].Value.ToString().Trim());
                    dQ = dgvMeasurement.Rows[i].Cells["Q_FactorValue"].Value == null || dgvMeasurement.Rows[i].Cells["Q_FactorValue"].Value.ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(dgvMeasurement.Rows[i].Cells["Q_FactorValue"].Value.ToString().Trim());

                    if (strDRPIGroup.Trim() == "A" && strDRPIType.Trim() == "제어용")
                    {
                        if (strCoilName.Trim() == "A1")
                        {
                            if (cboComparisonTarget.SelectedItem.ToString().Trim() == "평균값")
                            {
                                dgvMeasurement.Rows[i].Cells["DC_Deviation"].Value = (dA_ControlRodAveRdc[0] - dRdc).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["AC_Deviation"].Value = (dA_ControlRodAveRac[0] - dRac).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["L_Deviation"].Value = (dA_ControlRodAveL[0] - dL).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["C_Deviation"].Value = (dA_ControlRodAveC[0] - dC).ToString("F6");
                                dgvMeasurement.Rows[i].Cells["Q_Deviation"].Value = (dA_ControlRodAveQ[0] - dQ).ToString("F3");

                                dcmRefRdcValue = dA_ControlRodAveRdc[0];
                                dcmRefRacValue = dA_ControlRodAveRac[0];
                                dcmRefLValue = dA_ControlRodAveL[0];
                                dcmRefCValue = dA_ControlRodAveC[0];
                                dcmRefQValue = dA_ControlRodAveQ[0];
                            }
                            else
                            {
                                dgvMeasurement.Rows[i].Cells["DC_Deviation"].Value = (dA_ControlRodRdc[0] - dRdc).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["AC_Deviation"].Value = (dA_ControlRodRac[0] - dRac).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["L_Deviation"].Value = (dA_ControlRodL[0] - dL).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["C_Deviation"].Value = (dA_ControlRodC[0] - dC).ToString("F6");
                                dgvMeasurement.Rows[i].Cells["Q_Deviation"].Value = (dA_ControlRodQ[0] - dQ).ToString("F3");

                                dcmRefRdcValue = dA_ControlRodRdc[0];
                                dcmRefRacValue = dA_ControlRodRac[0];
                                dcmRefLValue = dA_ControlRodL[0];
                                dcmRefCValue = dA_ControlRodC[0];
                                dcmRefQValue = dA_ControlRodQ[0];
                            }
                        }
                        else if (strCoilName.Trim() == "A2")
                        {
                            if (cboComparisonTarget.SelectedItem.ToString().Trim() == "평균값")
                            {
                                dgvMeasurement.Rows[i].Cells["DC_Deviation"].Value = (dA_ControlRodAveRdc[1] - dRdc).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["AC_Deviation"].Value = (dA_ControlRodAveRac[1] - dRac).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["L_Deviation"].Value = (dA_ControlRodAveL[1] - dL).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["C_Deviation"].Value = (dA_ControlRodAveC[1] - dC).ToString("F6");
                                dgvMeasurement.Rows[i].Cells["Q_Deviation"].Value = (dA_ControlRodAveQ[1] - dQ).ToString("F3");

                                dcmRefRdcValue = dA_ControlRodAveRdc[1];
                                dcmRefRacValue = dA_ControlRodAveRac[1];
                                dcmRefLValue = dA_ControlRodAveL[1];
                                dcmRefCValue = dA_ControlRodAveC[1];
                                dcmRefQValue = dA_ControlRodAveQ[1];
                            }
                            else
                            {
                                dgvMeasurement.Rows[i].Cells["DC_Deviation"].Value = (dA_ControlRodRdc[1] - dRdc).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["AC_Deviation"].Value = (dA_ControlRodRac[1] - dRac).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["L_Deviation"].Value = (dA_ControlRodL[1] - dL).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["C_Deviation"].Value = (dA_ControlRodC[1] - dC).ToString("F6");
                                dgvMeasurement.Rows[i].Cells["Q_Deviation"].Value = (dA_ControlRodQ[1] - dQ).ToString("F3");

                                dcmRefRdcValue = dA_ControlRodRdc[1];
                                dcmRefRacValue = dA_ControlRodRac[1];
                                dcmRefLValue = dA_ControlRodL[1];
                                dcmRefCValue = dA_ControlRodC[1];
                                dcmRefQValue = dA_ControlRodQ[1];
                            }
                        }
                        else
                        {
                            if (cboComparisonTarget.SelectedItem.ToString().Trim() == "평균값")
                            {
                                dgvMeasurement.Rows[i].Cells["DC_Deviation"].Value = (dA_ControlRodAveRdc[2] - dRdc).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["AC_Deviation"].Value = (dA_ControlRodAveRac[2] - dRac).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["L_Deviation"].Value = (dA_ControlRodAveL[2] - dL).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["C_Deviation"].Value = (dA_ControlRodAveC[2] - dC).ToString("F6");
                                dgvMeasurement.Rows[i].Cells["Q_Deviation"].Value = (dA_ControlRodAveQ[2] - dQ).ToString("F3");

                                dcmRefRdcValue = dA_ControlRodAveRdc[2];
                                dcmRefRacValue = dA_ControlRodAveRac[2];
                                dcmRefLValue = dA_ControlRodAveL[2];
                                dcmRefCValue = dA_ControlRodAveC[2];
                                dcmRefQValue = dA_ControlRodAveQ[2];
                            }
                            else
                            {
                                dgvMeasurement.Rows[i].Cells["DC_Deviation"].Value = (dA_ControlRodRdc[2] - dRdc).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["AC_Deviation"].Value = (dA_ControlRodRac[2] - dRac).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["L_Deviation"].Value = (dA_ControlRodL[2] - dL).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["C_Deviation"].Value = (dA_ControlRodC[2] - dC).ToString("F6");
                                dgvMeasurement.Rows[i].Cells["Q_Deviation"].Value = (dA_ControlRodQ[2] - dQ).ToString("F3");

                                dcmRefRdcValue = dA_ControlRodRdc[2];
                                dcmRefRacValue = dA_ControlRodRac[2];
                                dcmRefLValue = dA_ControlRodL[2];
                                dcmRefCValue = dA_ControlRodC[2];
                                dcmRefQValue = dA_ControlRodQ[2];
                            }
                        }
                    }
                    else if (strDRPIGroup.Trim() == "A" && strDRPIType.Trim() == "정지용")
                    {
                        if (strCoilName.Trim() == "A1")
                        {
                            if (cboComparisonTarget.SelectedItem.ToString().Trim() == "평균값")
                            {
                                dgvMeasurement.Rows[i].Cells["DC_Deviation"].Value = (dA_StopAveRdc[0] - dRdc).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["AC_Deviation"].Value = (dA_StopAveRac[0] - dRac).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["L_Deviation"].Value = (dA_StopAveL[0] - dL).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["C_Deviation"].Value = (dA_StopAveC[0] - dC).ToString("F6");
                                dgvMeasurement.Rows[i].Cells["Q_Deviation"].Value = (dA_StopAveQ[0] - dQ).ToString("F3");

                                dcmRefRdcValue = dA_StopAveRdc[0];
                                dcmRefRacValue = dA_StopAveRac[0];
                                dcmRefLValue = dA_StopAveL[0];
                                dcmRefCValue = dA_StopAveC[0];
                                dcmRefQValue = dA_StopAveQ[0];
                            }
                            else
                            {
                                dgvMeasurement.Rows[i].Cells["DC_Deviation"].Value = (dA_StopRdc[0] - dRdc).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["AC_Deviation"].Value = (dA_StopRac[0] - dRac).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["L_Deviation"].Value = (dA_StopL[0] - dL).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["C_Deviation"].Value = (dA_StopC[0] - dC).ToString("F6");
                                dgvMeasurement.Rows[i].Cells["Q_Deviation"].Value = (dA_StopQ[0] - dQ).ToString("F3");

                                dcmRefRdcValue = dA_StopRdc[0];
                                dcmRefRacValue = dA_StopRac[0];
                                dcmRefLValue = dA_StopL[0];
                                dcmRefCValue = dA_StopC[0];
                                dcmRefQValue = dA_StopQ[0];
                            }
                        }
                        else if (strCoilName.Trim() == "A2")
                        {
                            if (cboComparisonTarget.SelectedItem.ToString().Trim() == "평균값")
                            {
                                dgvMeasurement.Rows[i].Cells["DC_Deviation"].Value = (dA_StopAveRdc[1] - dRdc).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["AC_Deviation"].Value = (dA_StopAveRac[1] - dRac).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["L_Deviation"].Value = (dA_StopAveL[1] - dL).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["C_Deviation"].Value = (dA_StopAveC[1] - dC).ToString("F6");
                                dgvMeasurement.Rows[i].Cells["Q_Deviation"].Value = (dA_StopAveQ[1] - dQ).ToString("F3");

                                dcmRefRdcValue = dA_StopAveRdc[1];
                                dcmRefRacValue = dA_StopAveRac[1];
                                dcmRefLValue = dA_StopAveL[1];
                                dcmRefCValue = dA_StopAveC[1];
                                dcmRefQValue = dA_StopAveQ[1];
                            }
                            else
                            {
                                dgvMeasurement.Rows[i].Cells["DC_Deviation"].Value = (dA_StopRdc[1] - dRdc).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["AC_Deviation"].Value = (dA_StopRac[1] - dRac).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["L_Deviation"].Value = (dA_StopL[1] - dL).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["C_Deviation"].Value = (dA_StopC[1] - dC).ToString("F6");
                                dgvMeasurement.Rows[i].Cells["Q_Deviation"].Value = (dA_StopQ[1] - dQ).ToString("F3");

                                dcmRefRdcValue = dA_StopRdc[1];
                                dcmRefRacValue = dA_StopRac[1];
                                dcmRefLValue = dA_StopL[1];
                                dcmRefCValue = dA_StopC[1];
                                dcmRefQValue = dA_StopQ[1];
                            }
                        }
                        else
                        {
                            if (cboComparisonTarget.SelectedItem.ToString().Trim() == "평균값")
                            {
                                dgvMeasurement.Rows[i].Cells["DC_Deviation"].Value = (dA_StopAveRdc[2] - dRdc).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["AC_Deviation"].Value = (dA_StopAveRac[2] - dRac).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["L_Deviation"].Value = (dA_StopAveL[2] - dL).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["C_Deviation"].Value = (dA_StopAveC[2] - dC).ToString("F6");
                                dgvMeasurement.Rows[i].Cells["Q_Deviation"].Value = (dA_StopAveQ[2] - dQ).ToString("F3");

                                dcmRefRdcValue = dA_StopAveRdc[2];
                                dcmRefRacValue = dA_StopAveRac[2];
                                dcmRefLValue = dA_StopAveL[2];
                                dcmRefCValue = dA_StopAveC[2];
                                dcmRefQValue = dA_StopAveQ[2];
                            }
                            else
                            {
                                dgvMeasurement.Rows[i].Cells["DC_Deviation"].Value = (dA_StopRdc[2] - dRdc).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["AC_Deviation"].Value = (dA_StopRac[2] - dRac).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["L_Deviation"].Value = (dA_StopL[2] - dL).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["C_Deviation"].Value = (dA_StopC[2] - dC).ToString("F6");
                                dgvMeasurement.Rows[i].Cells["Q_Deviation"].Value = (dA_StopQ[2] - dQ).ToString("F3");

                                dcmRefRdcValue = dA_StopRdc[2];
                                dcmRefRacValue = dA_StopRac[2];
                                dcmRefLValue = dA_StopL[2];
                                dcmRefCValue = dA_StopC[2];
                                dcmRefQValue = dA_StopQ[2];
                            }
                        }
                    }
                    else if (strDRPIGroup.Trim() == "B" && strDRPIType.Trim() == "제어용")
                    {
                        if (strCoilName.Trim() == "B1")
                        {
                            if (cboComparisonTarget.SelectedItem.ToString().Trim() == "평균값")
                            {
                                dgvMeasurement.Rows[i].Cells["DC_Deviation"].Value = (dB_ControlRodAveRdc[0] - dRdc).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["AC_Deviation"].Value = (dB_ControlRodAveRac[0] - dRac).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["L_Deviation"].Value = (dB_ControlRodAveL[0] - dL).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["C_Deviation"].Value = (dB_ControlRodAveC[0] - dC).ToString("F6");
                                dgvMeasurement.Rows[i].Cells["Q_Deviation"].Value = (dB_ControlRodAveQ[0] - dQ).ToString("F3");

                                dcmRefRdcValue = dB_ControlRodAveRdc[0];
                                dcmRefRacValue = dB_ControlRodAveRac[0];
                                dcmRefLValue = dB_ControlRodAveL[0];
                                dcmRefCValue = dB_ControlRodAveC[0];
                                dcmRefQValue = dB_ControlRodAveQ[0];
                            }
                            else
                            {
                                dgvMeasurement.Rows[i].Cells["DC_Deviation"].Value = (dB_ControlRodRdc[0] - dRdc).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["AC_Deviation"].Value = (dB_ControlRodRac[0] - dRac).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["L_Deviation"].Value = (dB_ControlRodL[0] - dL).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["C_Deviation"].Value = (dB_ControlRodC[0] - dC).ToString("F6");
                                dgvMeasurement.Rows[i].Cells["Q_Deviation"].Value = (dB_ControlRodQ[0] - dQ).ToString("F3");

                                dcmRefRdcValue = dB_ControlRodRdc[0];
                                dcmRefRacValue = dB_ControlRodRac[0];
                                dcmRefLValue = dB_ControlRodL[0];
                                dcmRefCValue = dB_ControlRodC[0];
                                dcmRefQValue = dB_ControlRodQ[0];
                            }
                        }
                        else if (strCoilName.Trim() == "B2")
                        {
                            if (cboComparisonTarget.SelectedItem.ToString().Trim() == "평균값")
                            {
                                dgvMeasurement.Rows[i].Cells["DC_Deviation"].Value = (dB_ControlRodAveRdc[1] - dRdc).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["AC_Deviation"].Value = (dB_ControlRodAveRac[1] - dRac).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["L_Deviation"].Value = (dB_ControlRodAveL[1] - dL).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["C_Deviation"].Value = (dB_ControlRodAveC[1] - dC).ToString("F6");
                                dgvMeasurement.Rows[i].Cells["Q_Deviation"].Value = (dB_ControlRodAveQ[1] - dQ).ToString("F3");

                                dcmRefRdcValue = dB_ControlRodAveRdc[1];
                                dcmRefRacValue = dB_ControlRodAveRac[1];
                                dcmRefLValue = dB_ControlRodAveL[1];
                                dcmRefCValue = dB_ControlRodAveC[1];
                                dcmRefQValue = dB_ControlRodAveQ[1];
                            }
                            else
                            {
                                dgvMeasurement.Rows[i].Cells["DC_Deviation"].Value = (dB_ControlRodRdc[1] - dRdc).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["AC_Deviation"].Value = (dB_ControlRodRac[1] - dRac).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["L_Deviation"].Value = (dB_ControlRodL[1] - dL).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["C_Deviation"].Value = (dB_ControlRodC[1] - dC).ToString("F6");
                                dgvMeasurement.Rows[i].Cells["Q_Deviation"].Value = (dB_ControlRodQ[1] - dQ).ToString("F3");

                                dcmRefRdcValue = dB_ControlRodRdc[1];
                                dcmRefRacValue = dB_ControlRodRac[1];
                                dcmRefLValue = dB_ControlRodL[1];
                                dcmRefCValue = dB_ControlRodC[1];
                                dcmRefQValue = dB_ControlRodQ[1];
                            }
                        }
                        else
                        {
                            if (cboComparisonTarget.SelectedItem.ToString().Trim() == "평균값")
                            {
                                dgvMeasurement.Rows[i].Cells["DC_Deviation"].Value = (dB_ControlRodAveRdc[2] - dRdc).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["AC_Deviation"].Value = (dB_ControlRodAveRac[2] - dRac).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["L_Deviation"].Value = (dB_ControlRodAveL[2] - dL).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["C_Deviation"].Value = (dB_ControlRodAveC[2] - dC).ToString("F6");
                                dgvMeasurement.Rows[i].Cells["Q_Deviation"].Value = (dB_ControlRodAveQ[2] - dQ).ToString("F3");

                                dcmRefRdcValue = dB_ControlRodAveRdc[2];
                                dcmRefRacValue = dB_ControlRodAveRac[2];
                                dcmRefLValue = dB_ControlRodAveL[2];
                                dcmRefCValue = dB_ControlRodAveC[2];
                                dcmRefQValue = dB_ControlRodAveQ[2];
                            }
                            else
                            {
                                dgvMeasurement.Rows[i].Cells["DC_Deviation"].Value = (dB_ControlRodRdc[2] - dRdc).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["AC_Deviation"].Value = (dB_ControlRodRac[2] - dRac).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["L_Deviation"].Value = (dB_ControlRodL[2] - dL).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["C_Deviation"].Value = (dB_ControlRodC[2] - dC).ToString("F6");
                                dgvMeasurement.Rows[i].Cells["Q_Deviation"].Value = (dB_ControlRodQ[2] - dQ).ToString("F3");

                                dcmRefRdcValue = dB_ControlRodRdc[2];
                                dcmRefRacValue = dB_ControlRodRac[2];
                                dcmRefLValue = dB_ControlRodL[2];
                                dcmRefCValue = dB_ControlRodC[2];
                                dcmRefQValue = dB_ControlRodQ[2];
                            }
                        }
                    }
                    else if (strDRPIGroup.Trim() == "B" && strDRPIType.Trim() == "정지용")
                    {
                        if (strCoilName.Trim() == "B1")
                        {
                            if (cboComparisonTarget.SelectedItem.ToString().Trim() == "평균값")
                            {
                                dgvMeasurement.Rows[i].Cells["DC_Deviation"].Value = (dB_StopAveRdc[0] - dRdc).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["AC_Deviation"].Value = (dB_StopAveRac[0] - dRac).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["L_Deviation"].Value = (dB_StopAveL[0] - dL).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["C_Deviation"].Value = (dB_StopAveC[0] - dC).ToString("F6");
                                dgvMeasurement.Rows[i].Cells["Q_Deviation"].Value = (dB_StopAveQ[0] - dQ).ToString("F3");

                                dcmRefRdcValue = dB_StopAveRdc[0];
                                dcmRefRacValue = dB_StopAveRac[0];
                                dcmRefLValue = dB_StopAveL[0];
                                dcmRefCValue = dB_StopAveC[0];
                                dcmRefQValue = dB_StopAveQ[0];
                            }
                            else
                            {
                                dgvMeasurement.Rows[i].Cells["DC_Deviation"].Value = (dB_StopRdc[0] - dRdc).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["AC_Deviation"].Value = (dB_StopRac[0] - dRac).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["L_Deviation"].Value = (dB_StopL[0] - dL).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["C_Deviation"].Value = (dB_StopC[0] - dC).ToString("F6");
                                dgvMeasurement.Rows[i].Cells["Q_Deviation"].Value = (dB_StopQ[0] - dQ).ToString("F3");

                                dcmRefRdcValue = dB_StopRdc[0];
                                dcmRefRacValue = dB_StopRac[0];
                                dcmRefLValue = dB_StopL[0];
                                dcmRefCValue = dB_StopC[0];
                                dcmRefQValue = dB_StopQ[0];
                            }
                        }
                        else if (strCoilName.Trim() == "B2")
                        {
                            if (cboComparisonTarget.SelectedItem.ToString().Trim() == "평균값")
                            {
                                dgvMeasurement.Rows[i].Cells["DC_Deviation"].Value = (dB_StopAveRdc[1] - dRdc).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["AC_Deviation"].Value = (dB_StopAveRac[1] - dRac).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["L_Deviation"].Value = (dB_StopAveL[1] - dL).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["C_Deviation"].Value = (dB_StopAveC[1] - dC).ToString("F6");
                                dgvMeasurement.Rows[i].Cells["Q_Deviation"].Value = (dB_StopAveQ[1] - dQ).ToString("F3");

                                dcmRefRdcValue = dB_StopAveRdc[1];
                                dcmRefRacValue = dB_StopAveRac[1];
                                dcmRefLValue = dB_StopAveL[1];
                                dcmRefCValue = dB_StopAveC[1];
                                dcmRefQValue = dB_StopAveQ[1];
                            }
                            else
                            {
                                dgvMeasurement.Rows[i].Cells["DC_Deviation"].Value = (dB_StopRdc[1] - dRdc).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["AC_Deviation"].Value = (dB_StopRac[1] - dRac).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["L_Deviation"].Value = (dB_StopL[1] - dL).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["C_Deviation"].Value = (dB_StopC[1] - dC).ToString("F6");
                                dgvMeasurement.Rows[i].Cells["Q_Deviation"].Value = (dB_StopQ[1] - dQ).ToString("F3");

                                dcmRefRdcValue = dB_StopRdc[1];
                                dcmRefRacValue = dB_StopRac[1];
                                dcmRefLValue = dB_StopL[1];
                                dcmRefCValue = dB_StopC[1];
                                dcmRefQValue = dB_StopQ[1];
                            }
                        }
                        else
                        {
                            if (cboComparisonTarget.SelectedItem.ToString().Trim() == "평균값")
                            {
                                dgvMeasurement.Rows[i].Cells["DC_Deviation"].Value = (dB_StopAveRdc[2] - dRdc).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["AC_Deviation"].Value = (dB_StopAveRac[2] - dRac).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["L_Deviation"].Value = (dB_StopAveL[2] - dL).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["C_Deviation"].Value = (dB_StopAveC[2] - dC).ToString("F6");
                                dgvMeasurement.Rows[i].Cells["Q_Deviation"].Value = (dB_StopAveQ[2] - dQ).ToString("F3");

                                dcmRefRdcValue = dB_StopAveRdc[2];
                                dcmRefRacValue = dB_StopAveRac[2];
                                dcmRefLValue = dB_StopAveL[2];
                                dcmRefCValue = dB_StopAveC[2];
                                dcmRefQValue = dB_StopAveQ[2];
                            }
                            else
                            {
                                dgvMeasurement.Rows[i].Cells["DC_Deviation"].Value = (dB_StopRdc[2] - dRdc).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["AC_Deviation"].Value = (dB_StopRac[2] - dRac).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["L_Deviation"].Value = (dB_StopL[2] - dL).ToString("F3");
                                dgvMeasurement.Rows[i].Cells["C_Deviation"].Value = (dB_StopC[2] - dC).ToString("F6");
                                dgvMeasurement.Rows[i].Cells["Q_Deviation"].Value = (dB_StopQ[2] - dQ).ToString("F3");

                                dcmRefRdcValue = dB_StopRdc[2];
                                dcmRefRacValue = dB_StopRac[2];
                                dcmRefLValue = dB_StopL[2];
                                dcmRefCValue = dB_StopC[2];
                                dcmRefQValue = dB_StopQ[2];
                            }
                        }
                    }

                    // DC
                    dcmDecisionRangeMaxValue = dcmRefRdcValue + (dcmRefRdcValue * (dcmDecisionRange_Rdc / 100));
                    dcmEffectiveStandardRangeMaxValue = dcmRefRdcValue + (dcmRefRdcValue * (dcmEffectiveStandardRange / 100));

                    dcmDecisionRangeMinValue = dcmRefRdcValue - (dcmRefRdcValue * (dcmDecisionRange_Rdc / 100));
                    dcmEffectiveStandardRangeMinValue = dcmRefRdcValue - (dcmRefRdcValue * (dcmEffectiveStandardRange / 100));

                    // 결과 체크
                    if ((dRdc <= dcmDecisionRangeMaxValue && dRdc > dcmEffectiveStandardRangeMaxValue)
                        || (dRdc >= dcmDecisionRangeMinValue && dRdc < dcmEffectiveStandardRangeMinValue))
                    {
                        dgvMeasurement.Rows[i].Cells["DC_ResistanceValue"].Style.ForeColor = System.Drawing.Color.Blue;
                    }
                    else if (dRdc > dcmDecisionRangeMaxValue || dRdc < dcmDecisionRangeMinValue)
                    {
                        dgvMeasurement.Rows[i].Cells["DC_ResistanceValue"].Style.ForeColor = System.Drawing.Color.Red;
                    }
                    else
                    {
                        dgvMeasurement.Rows[i].Cells["DC_ResistanceValue"].Style.ForeColor = System.Drawing.Color.Black;
                    }

                    // AC
                    dcmDecisionRangeMaxValue = dcmRefRacValue + (dcmRefRacValue * (dcmDecisionRange_Rac / 100));
                    dcmEffectiveStandardRangeMaxValue = dcmRefRacValue + (dcmRefRacValue * (dcmEffectiveStandardRange / 100));

                    dcmDecisionRangeMinValue = dcmRefRacValue - (dcmRefRacValue * (dcmDecisionRange_Rac / 100));
                    dcmEffectiveStandardRangeMinValue = dcmRefRacValue - (dcmRefRacValue * (dcmEffectiveStandardRange / 100));

                    // 결과 체크
                    if ((dRac <= dcmDecisionRangeMaxValue && dRac > dcmEffectiveStandardRangeMaxValue)
                        || (dRac >= dcmDecisionRangeMinValue && dRac < dcmEffectiveStandardRangeMinValue))
                    {
                        dgvMeasurement.Rows[i].Cells["AC_ResistanceValue"].Style.ForeColor = System.Drawing.Color.Blue;
                    }
                    else if (dRac > dcmDecisionRangeMaxValue || dRac < dcmDecisionRangeMinValue)
                    {
                        dgvMeasurement.Rows[i].Cells["AC_ResistanceValue"].Style.ForeColor = System.Drawing.Color.Red;
                    }
                    else
                    {
                        dgvMeasurement.Rows[i].Cells["AC_ResistanceValue"].Style.ForeColor = System.Drawing.Color.Black;
                    }

                    // L
                    dcmDecisionRangeMaxValue = dcmRefLValue + (dcmRefLValue * (dcmDecisionRange_L / 100));
                    dcmEffectiveStandardRangeMaxValue = dcmRefLValue + (dcmRefLValue * (dcmEffectiveStandardRange / 100));

                    dcmDecisionRangeMinValue = dcmRefLValue - (dcmRefLValue * (dcmDecisionRange_L / 100));
                    dcmEffectiveStandardRangeMinValue = dcmRefLValue - (dcmRefLValue * (dcmEffectiveStandardRange / 100));

                    // 결과 체크
                    if ((dL <= dcmDecisionRangeMaxValue && dL > dcmEffectiveStandardRangeMaxValue)
                        || (dL >= dcmDecisionRangeMinValue && dL < dcmEffectiveStandardRangeMinValue))
                    {
                        dgvMeasurement.Rows[i].Cells["L_InductanceValue"].Style.ForeColor = System.Drawing.Color.Blue;
                    }
                    else if (dL > dcmDecisionRangeMaxValue || dL < dcmDecisionRangeMinValue)
                    {
                        dgvMeasurement.Rows[i].Cells["L_InductanceValue"].Style.ForeColor = System.Drawing.Color.Red;
                    }
                    else
                    {
                        dgvMeasurement.Rows[i].Cells["L_InductanceValue"].Style.ForeColor = System.Drawing.Color.Black;
                    }

                    // C
                    dcmDecisionRangeMaxValue = dcmRefCValue + (dcmRefCValue * (dcmDecisionRange_C / 100));
                    dcmEffectiveStandardRangeMaxValue = dcmRefCValue + (dcmRefCValue * (dcmEffectiveStandardRange / 100));

                    dcmDecisionRangeMinValue = dcmRefCValue - (dcmRefCValue * (dcmDecisionRange_C / 100));
                    dcmEffectiveStandardRangeMinValue = dcmRefCValue - (dcmRefCValue * (dcmEffectiveStandardRange / 100));

                    // 결과 체크
                    if ((dC <= dcmDecisionRangeMaxValue && dC > dcmEffectiveStandardRangeMaxValue)
                        || (dC >= dcmDecisionRangeMinValue && dC < dcmEffectiveStandardRangeMinValue))
                    {
                        dgvMeasurement.Rows[i].Cells["C_CapacitanceValue"].Style.ForeColor = System.Drawing.Color.Blue;
                    }
                    else if (dC > dcmDecisionRangeMaxValue || dC < dcmDecisionRangeMinValue)
                    {
                        dgvMeasurement.Rows[i].Cells["C_CapacitanceValue"].Style.ForeColor = System.Drawing.Color.Red;
                    }
                    else
                    {
                        dgvMeasurement.Rows[i].Cells["C_CapacitanceValue"].Style.ForeColor = System.Drawing.Color.Black;
                    }

                    // Q
                    dcmDecisionRangeMaxValue = dcmRefQValue + (dcmRefQValue * (dcmDecisionRange_Q / 100));
                    dcmEffectiveStandardRangeMaxValue = dcmRefQValue + (dcmRefQValue * (dcmEffectiveStandardRange / 100));

                    dcmDecisionRangeMinValue = dcmRefQValue - (dcmRefQValue * (dcmDecisionRange_Q / 100));
                    dcmEffectiveStandardRangeMinValue = dcmRefQValue - (dcmRefQValue * (dcmEffectiveStandardRange / 100));

                    // 결과 체크
                    if ((dQ <= dcmDecisionRangeMaxValue && dQ > dcmEffectiveStandardRangeMaxValue)
                        || (dQ >= dcmDecisionRangeMinValue && dQ < dcmEffectiveStandardRangeMinValue))
                    {
                        dgvMeasurement.Rows[i].Cells["Q_FactorValue"].Style.ForeColor = System.Drawing.Color.Blue;
                    }
                    else if (dQ > dcmDecisionRangeMaxValue || dQ < dcmDecisionRangeMinValue)
                    {
                        dgvMeasurement.Rows[i].Cells["Q_FactorValue"].Style.ForeColor = System.Drawing.Color.Red;
                    }
                    else
                    {
                        dgvMeasurement.Rows[i].Cells["Q_FactorValue"].Style.ForeColor = System.Drawing.Color.Black;
                    }

                    // 결과 판단
                    if (dgvMeasurement.Rows[i].Cells["DC_ResistanceValue"].Style.ForeColor == System.Drawing.Color.Red
                        || dgvMeasurement.Rows[i].Cells["AC_ResistanceValue"].Style.ForeColor == System.Drawing.Color.Red
                        || dgvMeasurement.Rows[i].Cells["L_InductanceValue"].Style.ForeColor == System.Drawing.Color.Red
                        || dgvMeasurement.Rows[i].Cells["C_CapacitanceValue"].Style.ForeColor == System.Drawing.Color.Red
                        || dgvMeasurement.Rows[i].Cells["Q_FactorValue"].Style.ForeColor == System.Drawing.Color.Red)
                    {
                        strResult = "부적합";
                        dgvMeasurement.Rows[i].Cells["Result"].Value = strResult;
                        dgvMeasurement.Rows[i].Cells["Result"].Style.ForeColor = System.Drawing.Color.Red;
                    }
                    else if (dgvMeasurement.Rows[i].Cells["DC_ResistanceValue"].Style.ForeColor == System.Drawing.Color.Blue
                        || dgvMeasurement.Rows[i].Cells["AC_ResistanceValue"].Style.ForeColor == System.Drawing.Color.Blue
                        || dgvMeasurement.Rows[i].Cells["L_InductanceValue"].Style.ForeColor == System.Drawing.Color.Blue
                        || dgvMeasurement.Rows[i].Cells["C_CapacitanceValue"].Style.ForeColor == System.Drawing.Color.Blue
                        || dgvMeasurement.Rows[i].Cells["Q_FactorValue"].Style.ForeColor == System.Drawing.Color.Blue)
                    {
                        strResult = "의심";
                        dgvMeasurement.Rows[i].Cells["Result"].Value = strResult;
                        dgvMeasurement.Rows[i].Cells["Result"].Style.ForeColor = System.Drawing.Color.Blue;
                    }
                    else
                    {
                        strResult = "적합";
                        dgvMeasurement.Rows[i].Cells["Result"].Value = strResult;
                        dgvMeasurement.Rows[i].Cells["Result"].Style.ForeColor = System.Drawing.Color.Black;
                    }

                    bool boolDetail = false;

                    decimal dcmRdc_Deviation = dgvMeasurement.Rows[i].Cells["DC_Deviation"].Value == null || dgvMeasurement.Rows[i].Cells["DC_Deviation"].Value.ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(dgvMeasurement.Rows[i].Cells["DC_Deviation"].Value.ToString().Trim());
                    decimal dcmRac_Deviation = dgvMeasurement.Rows[i].Cells["AC_Deviation"].Value == null || dgvMeasurement.Rows[i].Cells["AC_Deviation"].Value.ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(dgvMeasurement.Rows[i].Cells["AC_Deviation"].Value.ToString().Trim());
                    decimal dcmL_Deviation = dgvMeasurement.Rows[i].Cells["L_Deviation"].Value == null || dgvMeasurement.Rows[i].Cells["L_Deviation"].Value.ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(dgvMeasurement.Rows[i].Cells["L_Deviation"].Value.ToString().Trim());
                    decimal dcmC_Deviation = dgvMeasurement.Rows[i].Cells["C_Deviation"].Value == null || dgvMeasurement.Rows[i].Cells["C_Deviation"].Value.ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(dgvMeasurement.Rows[i].Cells["C_Deviation"].Value.ToString().Trim());
                    decimal dcmQ_Deviation = dgvMeasurement.Rows[i].Cells["Q_Deviation"].Value == null || dgvMeasurement.Rows[i].Cells["Q_Deviation"].Value.ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(dgvMeasurement.Rows[i].Cells["Q_Deviation"].Value.ToString().Trim());

                    if ((m_db.GetDRPIDiagnosisDetailDataGridViewDataCount(strPlantName.Trim(), strHogi, strOhDegree, strDRPIGroup, strDRPIType, strControlName, strCoilName, intCoilNumber)) > 0)
                    {
                        // 기존 데이터 Update
                        boolDetail = m_db.UpdateDRPIDiagnosisDetailDataGridViewDataInfo(strPlantName.Trim(), strHogi, strOhDegree, strDRPIGroup, strDRPIType
                            , strControlName, strCoilName, intCoilNumber, dcmRdc_Deviation, dcmRac_Deviation, dcmL_Deviation, dcmC_Deviation
                            , dcmQ_Deviation, strResult);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
        }

        /// <summary>
        /// 제어용/정지용의 그룹 A/B 기준 값 저장
        /// </summary>
        /// <param name="dgv"></param>
        /// <param name="strHogi"></param>
        /// <param name="strOhDegree"></param>
        /// <param name="strColiCabinetType"></param>
        /// <param name="strCoilCabinetName"></param>
        /// <returns></returns>
        private bool SaveControlReferenceValue(DataGridView dgv, string strHogi, string strOhDegree, string strColiCabinetType, string strCoilCabinetName)
        {
            bool boolResultSave = false;
            string[] arrayCoilCabinetType = new string[3];
            string[] arrayCoilCabinetName = new string[3];
            double[] dRdc_DRPIReferenceValue = new double[3];
            double[] dRac_DRPIReferenceValue = new double[3];
            double[] dL_DRPIReferenceValue = new double[3];
            double[] dC_DRPIReferenceValue = new double[3];
            double[] dQ_DRPIReferenceValue = new double[3];
            double dSumRdc_DRPIReferenceValue = 0.000d;
            double dSumRac_DRPIReferenceValue = 0.000d;
            double dSumL_DRPIReferenceValue = 0.000d;
            double dSumC_DRPIReferenceValue = 0.000000d;
            double dSumQ_DRPIReferenceValue = 0.000d;
            string[] arrayData = new string[5];
            int intSumCount = 0;

            for (int i = 0; i < dgv.Rows.Count; i++)
            {
                if (dgv.Rows[i].Cells["CoilCabinetName"].Value.ToString().Trim() == (strCoilCabinetName.Trim() + "1"))
                {
                    arrayCoilCabinetType[0] = dgv.Rows[i].Cells["CoilCabinetType"].Value == null ? strColiCabinetType.Trim() : dgv.Rows[i].Cells["CoilCabinetType"].Value.ToString().Trim();
                    arrayCoilCabinetName[0] = dgv.Rows[i].Cells["CoilCabinetName"].Value == null ? strCoilCabinetName.Trim() + "1" : dgv.Rows[i].Cells["CoilCabinetName"].Value.ToString().Trim();
                    dRdc_DRPIReferenceValue[0] = dgv.Rows[i].Cells["Rdc_DRPIReferenceValue"].Value == null
                        || dgv.Rows[i].Cells["Rdc_DRPIReferenceValue"].Value.ToString().Trim() == ""
                        ? 0.000d : Convert.ToDouble(dgv.Rows[i].Cells["Rdc_DRPIReferenceValue"].Value.ToString().Trim());
                    dRac_DRPIReferenceValue[0] = dgv.Rows[i].Cells["Rac_DRPIReferenceValue"].Value == null
                        || dgv.Rows[i].Cells["Rac_DRPIReferenceValue"].Value.ToString().Trim() == ""
                        ? 0.000d : Convert.ToDouble(dgv.Rows[i].Cells["Rac_DRPIReferenceValue"].Value.ToString().Trim());
                    dL_DRPIReferenceValue[0] = dgv.Rows[i].Cells["L_DRPIReferenceValue"].Value == null
                        || dgv.Rows[i].Cells["L_DRPIReferenceValue"].Value.ToString().Trim() == ""
                        ? 0.000d : Convert.ToDouble(dgv.Rows[i].Cells["L_DRPIReferenceValue"].Value.ToString().Trim());
                    dC_DRPIReferenceValue[0] = dgv.Rows[i].Cells["C_DRPIReferenceValue"].Value == null
                        || dgv.Rows[i].Cells["C_DRPIReferenceValue"].Value.ToString().Trim() == ""
                        ? 0.000d : Convert.ToDouble(dgv.Rows[i].Cells["C_DRPIReferenceValue"].Value.ToString().Trim());
                    dQ_DRPIReferenceValue[0] = dgv.Rows[i].Cells["Q_DRPIReferenceValue"].Value == null
                        || dgv.Rows[i].Cells["Q_DRPIReferenceValue"].Value.ToString().Trim() == ""
                        ? 0.000d : Convert.ToDouble(dgv.Rows[i].Cells["Q_DRPIReferenceValue"].Value.ToString().Trim());
                }
                else if (dgv.Rows[i].Cells["CoilCabinetName"].Value.ToString().Trim() == (strCoilCabinetName.Trim() + "2"))
                {
                    arrayCoilCabinetType[1] = dgv.Rows[i].Cells["CoilCabinetType"].Value == null ? strColiCabinetType.Trim() : dgv.Rows[i].Cells["CoilCabinetType"].Value.ToString().Trim();
                    arrayCoilCabinetName[1] = dgv.Rows[i].Cells["CoilCabinetName"].Value == null ? strCoilCabinetName.Trim() + "2" : dgv.Rows[i].Cells["CoilCabinetName"].Value.ToString().Trim();
                    dRdc_DRPIReferenceValue[1] = dgv.Rows[i].Cells["Rdc_DRPIReferenceValue"].Value == null
                        || dgv.Rows[i].Cells["Rdc_DRPIReferenceValue"].Value.ToString().Trim() == ""
                        ? 0.000d : Convert.ToDouble(dgv.Rows[i].Cells["Rdc_DRPIReferenceValue"].Value.ToString().Trim());
                    dRac_DRPIReferenceValue[1] = dgv.Rows[i].Cells["Rac_DRPIReferenceValue"].Value == null
                        || dgv.Rows[i].Cells["Rac_DRPIReferenceValue"].Value.ToString().Trim() == ""
                        ? 0.000d : Convert.ToDouble(dgv.Rows[i].Cells["Rac_DRPIReferenceValue"].Value.ToString().Trim());
                    dL_DRPIReferenceValue[1] = dgv.Rows[i].Cells["L_DRPIReferenceValue"].Value == null
                        || dgv.Rows[i].Cells["L_DRPIReferenceValue"].Value.ToString().Trim() == ""
                        ? 0.000d : Convert.ToDouble(dgv.Rows[i].Cells["L_DRPIReferenceValue"].Value.ToString().Trim());
                    dC_DRPIReferenceValue[1] = dgv.Rows[i].Cells["C_DRPIReferenceValue"].Value == null
                        || dgv.Rows[i].Cells["C_DRPIReferenceValue"].Value.ToString().Trim() == ""
                        ? 0.000d : Convert.ToDouble(dgv.Rows[i].Cells["C_DRPIReferenceValue"].Value.ToString().Trim());
                    dQ_DRPIReferenceValue[1] = dgv.Rows[i].Cells["Q_DRPIReferenceValue"].Value == null
                        || dgv.Rows[i].Cells["Q_DRPIReferenceValue"].Value.ToString().Trim() == ""
                        ? 0.000d : Convert.ToDouble(dgv.Rows[i].Cells["Q_DRPIReferenceValue"].Value.ToString().Trim());
                }
                else
                {
                    dSumRdc_DRPIReferenceValue += dgv.Rows[i].Cells["Rdc_DRPIReferenceValue"].Value == null
                        || dgv.Rows[i].Cells["Rdc_DRPIReferenceValue"].Value.ToString().Trim() == ""
                        ? 0.000d : Convert.ToDouble(dgv.Rows[i].Cells["Rdc_DRPIReferenceValue"].Value.ToString().Trim());
                    dSumRac_DRPIReferenceValue += dgv.Rows[i].Cells["Rac_DRPIReferenceValue"].Value == null
                        || dgv.Rows[i].Cells["Rac_DRPIReferenceValue"].Value.ToString().Trim() == ""
                        ? 0.000d : Convert.ToDouble(dgv.Rows[i].Cells["Rac_DRPIReferenceValue"].Value.ToString().Trim());
                    dSumL_DRPIReferenceValue += dgv.Rows[i].Cells["L_DRPIReferenceValue"].Value == null
                        || dgv.Rows[i].Cells["L_DRPIReferenceValue"].Value.ToString().Trim() == ""
                        ? 0.000d : Convert.ToDouble(dgv.Rows[i].Cells["L_DRPIReferenceValue"].Value.ToString().Trim());
                    dSumC_DRPIReferenceValue += dgv.Rows[i].Cells["C_DRPIReferenceValue"].Value == null
                        || dgv.Rows[i].Cells["C_DRPIReferenceValue"].Value.ToString().Trim() == ""
                        ? 0.000d : Convert.ToDouble(dgv.Rows[i].Cells["C_DRPIReferenceValue"].Value.ToString().Trim());
                    dSumQ_DRPIReferenceValue += dgv.Rows[i].Cells["Q_DRPIReferenceValue"].Value == null
                        || dgv.Rows[i].Cells["Q_DRPIReferenceValue"].Value.ToString().Trim() == ""
                        ? 0.000d : Convert.ToDouble(dgv.Rows[i].Cells["Q_DRPIReferenceValue"].Value.ToString().Trim());

                    intSumCount++;
                }
            }

            if (intSumCount < 1) intSumCount = 1;

            arrayCoilCabinetType[2] = strColiCabinetType.Trim();
            arrayCoilCabinetName[2] = strCoilCabinetName + (strColiCabinetType.Trim() == "제어용" ? "3 ~ 21" : "3~6,18~21");
            dRdc_DRPIReferenceValue[2] = dSumRdc_DRPIReferenceValue / intSumCount;
            dRac_DRPIReferenceValue[2] = dSumRac_DRPIReferenceValue / intSumCount;
            dL_DRPIReferenceValue[2] = dSumL_DRPIReferenceValue / intSumCount;
            dC_DRPIReferenceValue[2] = dSumC_DRPIReferenceValue / intSumCount;
            dQ_DRPIReferenceValue[2] = dSumQ_DRPIReferenceValue / intSumCount;

            for (int i = 0; i < arrayCoilCabinetType.Length; i++)
            {
                arrayData[0] = dRdc_DRPIReferenceValue[i].ToString("F3").Trim();
                arrayData[1] = dRac_DRPIReferenceValue[i].ToString("F3").Trim();
                arrayData[2] = dL_DRPIReferenceValue[i].ToString("F3").Trim();
                arrayData[3] = dC_DRPIReferenceValue[i].ToString("F6").Trim();
                arrayData[4] = dQ_DRPIReferenceValue[i].ToString("F3").Trim();

                if ((m_db.GetDRPIReferenceValueDataCount(strPlantName.Trim(), strHogi.Trim(), strOhDegree.Trim(), arrayCoilCabinetType[i].Trim(), arrayCoilCabinetName[i].Trim())) > 0)
                {
                    boolResultSave = m_db.UpDateDRPIReferenceValueDataInfo(strPlantName.Trim(), strHogi.Trim(), strOhDegree.Trim(), arrayCoilCabinetType[i].Trim(), arrayCoilCabinetName[i].Trim(), arrayData);
                }
                else
                {
                    boolResultSave = m_db.InsertDRPIReferenceValueDataInfo(strPlantName.Trim(), strHogi.Trim(), strOhDegree.Trim(), arrayCoilCabinetType[i].Trim(), arrayCoilCabinetName[i].Trim(), arrayData);
                }
            }

            return boolResultSave;
        }

        /// <summary>
        /// 닫기 Button Evene
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// 조회 Button Evene
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSearch_Click(object sender, EventArgs e)
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

                btnSearch.Enabled = false;

                // 그리드와 그래프 초기화
                SetDataGridViewAndGraphInitialize();

                // 기준 값(DRPI) 초기화
                SetDRPIReferenceValueInitialize();

                // 평균 값(DRPI) 초기화
                SetDRPIAverageValueInitialize();

                // 기준 값(DRPI) 가져오기
                GetDRPIReferenceValueSelected();

                // Data 가져오기
                GetDRPIMeasurementDataBinding();

                if (btnReferenceValue.Text.Trim() == "기준값 표시 (F3)")
                {
                    // 정지,올림,이동 label 기준값 Visible 설정
                    SetReferenceValueLabelVisible(false);
                }
                else
                {
                    // 정지,올림,이동 label 기준값 Visible 설정
                    SetReferenceValueLabelVisible(true);
                }

                if (btnAverageValue.Text.Trim() == "평균값 표시 (F4)")
                {
                    // 정지,올림,이동 label 평균값 Visible 설정
                    SetAverageValueLabelVisible(false);
                }
                else
                {
                    // 정지,올림,이동 label 평균값 Visible 설정
                    SetAverageValueLabelVisible(true);
                }

                frmMB.lblMessage.Text = "조회 완료되었습니다.";
                frmMB.TopMost = true;
                frmMB.ShowDialog();
            }
            catch (Exception ex)
            {
                frmMB.lblMessage.Text = "조회 중 실패";
                frmMB.TopMost = true;
                frmMB.ShowDialog();
                System.Diagnostics.Debug.Print(ex.Message);
            }
            finally
            {
                btnSearch.Enabled = true;
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
            dgvControlAverageA.Rows.Clear();
            dgvControlAverageB.Rows.Clear();
            dgvStopAverageA.Rows.Clear();
            dgvStopAverageB.Rows.Clear();

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
        private void GetDRPIMeasurementDataBinding()
        {
            try
            {
                string strHogi = cboHogi.SelectedItem == null ? "" : cboHogi.SelectedItem.ToString().Trim();
                string strOhDegree = cboOhDegree.SelectedItem == null ? "" : cboOhDegree.SelectedItem.ToString().Trim();
                string strDRPIGroup = cboGroup.SelectedItem == null ? "" : cboGroup.SelectedItem.ToString().Trim();            
                string strMeasurementType = cboMeasurementType.SelectedItem == null ? "" : cboMeasurementType.SelectedItem.ToString().Trim();
                string strType = cboType.SelectedItem == null ? "" : cboType.SelectedItem.ToString().Trim();
                string strControlRod = cboControlRod.SelectedItem == null 
                    ? "" : cboControlRod.SelectedItem.ToString().Trim() == "전체" 
                        ? cboControlRod.SelectedItem.ToString().Trim() : string.Format("'{0}'", cboControlRod.SelectedItem.ToString().Trim());
                string strCoilName = cboCoilName.SelectedItem == null ? "" : cboCoilName.SelectedItem.ToString().Trim();
                string strFrequency = cboFrequency.SelectedItem == null ? "" : cboFrequency.SelectedItem.ToString().Trim();
                string strComparisonTarget = cboComparisonTarget.SelectedItem == null ? "" : cboComparisonTarget.SelectedItem.ToString().Trim();

                if (strDRPIGroup.Trim() != "전체" && strCoilName.Trim() != "" && strCoilName.Trim() != "전체")
                {
                    strCoilName = string.Format("'{0}{1}'", strDRPIGroup.Trim(), Regex.Replace(strCoilName, @"\D", ""));
                }
                else if (strDRPIGroup.Trim() == "전체" && strCoilName.Trim() != "" && strCoilName.Trim() != "전체")
                {
                    strCoilName = string.Format("'A{0}', 'B{0}'", Regex.Replace(strCoilName, @"\D", ""));
                }

                DataTable dtDB = new DataTable();
                DataTable dtAverage = new DataTable();
                
                    // 측정 타입이 평균치일 경우
                if (cboMeasurementType.SelectedItem != null && cboMeasurementType.SelectedItem.ToString().Trim() == "평균치")
                {
                    dtDB = m_db.GetDRPIDiagnosisDetailDataGridViewDataAverage(strPlantName.Trim(), strHogi, strOhDegree, strDRPIGroup
                            , strType, strControlRod, strCoilName, strFrequency, strReferenceHogi, strReferenceOHDegree);

                    dtAverage = dtDB;
                }
                else
                {
                    dtDB = m_db.GetDRPIDiagnosisDetailDataGridViewDataMeasure(strPlantName.Trim(), strHogi, strOhDegree, strDRPIGroup
                            , strType, strControlRod, strCoilName, strFrequency, strReferenceHogi, strReferenceOHDegree);

                    dtAverage = m_db.GetDRPIDiagnosisDetailDataGridViewDataAverage(strPlantName.Trim(), strHogi, strOhDegree, strDRPIGroup
                            , strType, strControlRod, strCoilName, strFrequency, strReferenceHogi, strReferenceOHDegree);
                }

                if (dtDB == null || dtDB.Rows.Count <= 0)
                {
                    frmMB.lblMessage.Text = "데이터가 없습니다.";
                    frmMB.TopMost = true;
                    frmMB.ShowDialog();
                }
                else
                {
                    // 그룹 A, B의 각각 제어용, 정지용의 평균가져오기
                    DataTable dtAvg = new DataTable();
                    dtAvg = m_db.GetDRPIDiagnosisDetailDataGridViewDataAverage(strPlantName.Trim(), strHogi, strOhDegree, strDRPIGroup
                            , strType, strControlRod, strCoilName, strFrequency);

                    decimal dGroupA1_ControlDC = 0M, dGroupA1_ControlAC = 0M, dGroupA1_ControlL = 0M, dGroupA1_ControlC = 0M, dGroupA1_ControlQ = 0M
                           , dGroupA2_ControlDC = 0M, dGroupA2_ControlAC = 0M, dGroupA2_ControlL = 0M, dGroupA2_ControlC = 0M, dGroupA2_ControlQ = 0M
                           , dGroupA_ControlDC = 0M, dGroupA_ControlAC = 0M, dGroupA_ControlL = 0M, dGroupA_ControlC = 0M, dGroupA_ControlQ = 0M
                           , dGroupB1_ControlDC = 0M, dGroupB1_ControlAC = 0M, dGroupB1_ControlL = 0M, dGroupB1_ControlC = 0M, dGroupB1_ControlQ = 0M
                           , dGroupB2_ControlDC = 0M, dGroupB2_ControlAC = 0M, dGroupB2_ControlL = 0M, dGroupB2_ControlC = 0M, dGroupB2_ControlQ = 0M
                           , dGroupB_ControlDC = 0M, dGroupB_ControlAC = 0M, dGroupB_ControlL = 0M, dGroupB_ControlC = 0M, dGroupB_ControlQ = 0M
                           , dGroupA1_StopDC = 0M, dGroupA1_StopAC = 0M, dGroupA1_StopL = 0M, dGroupA1_StopC = 0M, dGroupA1_StopQ = 0M
                           , dGroupA2_StopDC = 0M, dGroupA2_StopAC = 0M, dGroupA2_StopL = 0M, dGroupA2_StopC = 0M, dGroupA2_StopQ = 0M
                           , dGroupA_StopDC = 0M, dGroupA_StopAC = 0M, dGroupA_StopL = 0M, dGroupA_StopC = 0M, dGroupA_StopQ = 0M
                           , dGroupB1_StopDC = 0M, dGroupB1_StopAC = 0M, dGroupB1_StopL = 0M, dGroupB1_StopC = 0M, dGroupB1_StopQ = 0M
                           , dGroupB2_StopDC = 0M, dGroupB2_StopAC = 0M, dGroupB2_StopL = 0M, dGroupB2_StopC = 0M, dGroupB2_StopQ = 0M
                           , dGroupB_StopDC = 0M, dGroupB_StopAC = 0M, dGroupB_StopL = 0M, dGroupB_StopC = 0M, dGroupB_StopQ = 0M;

                    // 그룹 A, B의 각각 제어용, 정지용의 가져온 평균값을 변수에 설정
                    GetGeoupItemAverage(dtAvg, ref dGroupA1_ControlDC, ref dGroupA1_ControlAC, ref dGroupA1_ControlL, ref dGroupA1_ControlC, ref dGroupA1_ControlQ
                                            , ref dGroupA2_ControlDC, ref dGroupA2_ControlAC, ref dGroupA2_ControlL, ref dGroupA2_ControlC, ref dGroupA2_ControlQ
                                            , ref dGroupA_ControlDC, ref dGroupA_ControlAC, ref dGroupA_ControlL, ref dGroupA_ControlC, ref dGroupA_ControlQ
                                            , ref dGroupB1_ControlDC, ref dGroupB1_ControlAC, ref dGroupB1_ControlL, ref dGroupB1_ControlC, ref dGroupB1_ControlQ
                                            , ref dGroupB2_ControlDC, ref dGroupB2_ControlAC, ref dGroupB2_ControlL, ref dGroupB2_ControlC, ref dGroupB2_ControlQ
                                            , ref dGroupB_ControlDC, ref dGroupB_ControlAC, ref dGroupB_ControlL, ref dGroupB_ControlC, ref dGroupB_ControlQ
                                            , ref dGroupA1_StopDC, ref dGroupA1_StopAC, ref dGroupA1_StopL, ref dGroupA1_StopC, ref dGroupA1_StopQ
                                            , ref dGroupA2_StopDC, ref dGroupA2_StopAC, ref dGroupA2_StopL, ref dGroupA2_StopC, ref dGroupA2_StopQ
                                            , ref dGroupA_StopDC, ref dGroupA_StopAC, ref dGroupA_StopL, ref dGroupA_StopC, ref dGroupA_StopQ
                                            , ref dGroupB1_StopDC, ref dGroupB1_StopAC, ref dGroupB1_StopL, ref dGroupB1_StopC, ref dGroupB1_StopQ
                                            , ref dGroupB2_StopDC, ref dGroupB2_StopAC, ref dGroupB2_StopL, ref dGroupB2_StopC, ref dGroupB2_StopQ
                                            , ref dGroupB_StopDC, ref dGroupB_StopAC, ref dGroupB_StopL, ref dGroupB_StopC, ref dGroupB_StopQ);

                    // 코일 그룹A 제어용 평균치 산출
                    string filterExpression = string.Format("DRPIGroup = '{0}' AND DRPIType = '{1}'", "A", "제어용");
                    string sort = "DRPIGroup, DRPIType, SeqNumber";
                    if (dtAverage.Select(filterExpression, sort) != null && dtAverage.Select(filterExpression, sort).Length > 0)
                        SetControlRodCoilGroupA_AverageCalculation(dtAverage.Select(filterExpression, sort).CopyToDataTable()
                            , dGroupA1_ControlDC, dGroupA1_ControlAC, dGroupA1_ControlL, dGroupA1_ControlC, dGroupA1_ControlQ
                            , dGroupA2_ControlDC, dGroupA2_ControlAC, dGroupA2_ControlL, dGroupA2_ControlC, dGroupA2_ControlQ
                            , dGroupA_ControlDC, dGroupA_ControlAC, dGroupA_ControlL, dGroupA_ControlC, dGroupA_ControlQ);
                    // 코일 그룹B 제어용 평균치 산출
                    filterExpression = string.Format("DRPIGroup = '{0}' AND DRPIType = '{1}'", "B", "제어용");
                    if (dtAverage.Select(filterExpression, sort) != null && dtAverage.Select(filterExpression, sort).Length > 0)
                        SetControlRodCoilGroupB_AverageCalculation(dtAverage.Select(filterExpression, sort).CopyToDataTable()
                            , dGroupB1_ControlDC, dGroupB1_ControlAC, dGroupB1_ControlL, dGroupB1_ControlC, dGroupB1_ControlQ
                            , dGroupB2_ControlDC, dGroupB2_ControlAC, dGroupB2_ControlL, dGroupB2_ControlC, dGroupB2_ControlQ
                            , dGroupB_ControlDC, dGroupB_ControlAC, dGroupB_ControlL, dGroupB_ControlC, dGroupB_ControlQ);
                    // 코일 그룹A 정지용 평균치 산출
                    filterExpression = string.Format("DRPIGroup = '{0}' AND DRPIType = '{1}'", "A", "정지용");
                    if (dtAverage.Select(filterExpression, sort) != null && dtAverage.Select(filterExpression, sort).Length > 0)
                        SetStopCoilGroupA_AverageCalculation(dtAverage.Select(filterExpression, sort).CopyToDataTable()
                            , dGroupA1_StopDC, dGroupA1_StopAC, dGroupA1_StopL, dGroupA1_StopC, dGroupA1_StopQ
                            , dGroupA2_StopDC, dGroupA2_StopAC, dGroupA2_StopL, dGroupA2_StopC, dGroupA2_StopQ
                            , dGroupA_StopDC, dGroupA_StopAC, dGroupA_StopL, dGroupA_StopC, dGroupA_StopQ);
                    // 코일 그룹B 정지용 평균치 산출
                    filterExpression = string.Format("DRPIGroup = '{0}' AND DRPIType = '{1}'", "B", "정지용");
                    if (dtAverage.Select(filterExpression, sort) != null && dtAverage.Select(filterExpression, sort).Length > 0)
                        SetStopCoilGroupB_AverageCalculation(dtAverage.Select(filterExpression, sort).CopyToDataTable()
                            , dGroupB1_StopDC, dGroupB1_StopAC, dGroupB1_StopL, dGroupB1_StopC, dGroupB1_StopQ
                            , dGroupB2_StopDC, dGroupB2_StopAC, dGroupB2_StopL, dGroupB2_StopC, dGroupB2_StopQ
                            , dGroupB_StopDC, dGroupB_StopAC, dGroupB_StopL, dGroupB_StopC, dGroupB_StopQ);

                    // 표 형식
                    SetDRPIReferenceValueTableMark(ref dtDB, strComparisonTarget);

                    // 평균 값 Text Box 설정
                    SetAverageValueTextBox();

                    if (dtAverage != null && dtAverage.Rows.Count > 0)
                    {
                        string strXLabelValue = "", strXLabelCheck = "", strDataCheck = "";

                        for (int i = 0; i < dtAverage.Rows.Count; i++)
                        {
                            strDataCheck = dtAverage.Rows[i]["ControlName"].ToString().Trim() + " " + dtAverage.Rows[i]["DRPIGroup"].ToString().Trim();

                            if (strXLabelCheck.Trim() != strDataCheck.Trim())
                            {
                                strXLabelCheck = dtAverage.Rows[i]["ControlName"].ToString().Trim() + " " + dtAverage.Rows[i]["DRPIGroup"].ToString().Trim();

                                if (strXLabelValue.Trim() == "")
                                    strXLabelValue = dtAverage.Rows[i]["ControlName"].ToString().Trim() + " " + dtAverage.Rows[i]["DRPIGroup"].ToString().Trim();
                                else
                                    strXLabelValue = strXLabelValue + "," + dtAverage.Rows[i]["ControlName"].ToString().Trim() + " " + dtAverage.Rows[i]["DRPIGroup"].ToString().Trim();
                            }
                        }

                        string[] arrayXLabelValue = strXLabelValue.Split(',');
                        int intXLabelValueCount = arrayXLabelValue.Length;

                        // 측정값 및 편차 그래프 그리기
                        SetDRPIReferenceValueMeasurement(dtAverage, arrayXLabelValue, intXLabelValueCount);
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
                    int intControlRowA = 0;
                    if (intAAvgRdcCount > 0 || intAAvgRacCount > 0 || intAAvgLCount > 0 || intAAvgCCount > 0 || intAAvgQCount > 0)
                    {
                        intControlRowA = dgvControlAverageA.Rows.Add();

                        dgvControlAverageA.Rows[intControlRowA].Cells["CoilCabinetType"].Value = "제어용";
                        dgvControlAverageA.Rows[intControlRowA].Cells["CoilCabinetName"].Value = "A" + Regex.Replace(strCoilName.Trim(), @"\D", "");
                    }

                    if (intAAvgRdcCount > 0)
                        dgvControlAverageA.Rows[intControlRowA].Cells["Rdc_DRPIReferenceValue"].Value = (dcmASumControlRodRdc / intAAvgRdcCount).ToString("F3").Trim();
                    else
                        dgvControlAverageA.Rows[intControlRowA].Cells["Rdc_DRPIReferenceValue"].Value = "0.000";

                    if (intAAvgRacCount > 0)
                        dgvControlAverageA.Rows[intControlRowA].Cells["Rac_DRPIReferenceValue"].Value = (dcmASumControlRodRac / intAAvgRacCount).ToString("F3").Trim();
                    else
                        dgvControlAverageA.Rows[intControlRowA].Cells["Rac_DRPIReferenceValue"].Value = "0.000";

                    if (intAAvgLCount > 0)
                        dgvControlAverageA.Rows[intControlRowA].Cells["L_DRPIReferenceValue"].Value = (dcmASumControlRodL / intAAvgLCount).ToString("F3").Trim();
                    else
                        dgvControlAverageA.Rows[intControlRowA].Cells["L_DRPIReferenceValue"].Value = "0.000";

                    if (intAAvgCCount > 0)
                        dgvControlAverageA.Rows[intControlRowA].Cells["C_DRPIReferenceValue"].Value = (dcmASumControlRodC / intAAvgCCount).ToString("F6").Trim();
                    else
                        dgvControlAverageA.Rows[intControlRowA].Cells["C_DRPIReferenceValue"].Value = "0.000";

                    if (intAAvgQCount > 0)
                        dgvControlAverageA.Rows[intControlRowA].Cells["Q_DRPIReferenceValue"].Value = (dcmASumControlRodQ / intAAvgQCount).ToString("F3").Trim();
                    else
                        dgvControlAverageA.Rows[intControlRowA].Cells["Q_DRPIReferenceValue"].Value = "0.000";
                }

                decimal dcmA3RdcSum = 0M, dcmA3RacSum = 0M, dcmA3LSum = 0M, dcmA3CSum = 0M, dcmA3QSum = 0M;
                int intA3SCount = 0;

                for (int i = 0; i < dgvControlAverageA.RowCount; i++)
                {
                    if (dgvControlAverageA.Rows[i].Cells["CoilCabinetName"].Value.ToString().Trim() == "A1")
                    {
                        dA_ControlRodAveRdc[0] = dgvControlAverageA.Rows[i].Cells["Rdc_DRPIReferenceValue"].Value == null || dgvControlAverageA.Rows[i].Cells["Rdc_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(dgvControlAverageA.Rows[i].Cells["Rdc_DRPIReferenceValue"].Value.ToString().Trim());
                        dA_ControlRodAveRac[0] = dgvControlAverageA.Rows[i].Cells["Rac_DRPIReferenceValue"].Value == null || dgvControlAverageA.Rows[i].Cells["Rac_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(dgvControlAverageA.Rows[i].Cells["Rac_DRPIReferenceValue"].Value.ToString().Trim());
                        dA_ControlRodAveL[0] = dgvControlAverageA.Rows[i].Cells["L_DRPIReferenceValue"].Value == null || dgvControlAverageA.Rows[i].Cells["L_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(dgvControlAverageA.Rows[i].Cells["L_DRPIReferenceValue"].Value.ToString().Trim());
                        dA_ControlRodAveC[0] = dgvControlAverageA.Rows[i].Cells["C_DRPIReferenceValue"].Value == null || dgvControlAverageA.Rows[i].Cells["C_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.000000M : Convert.ToDecimal(dgvControlAverageA.Rows[i].Cells["C_DRPIReferenceValue"].Value.ToString().Trim());
                        dA_ControlRodAveQ[0] = dgvControlAverageA.Rows[i].Cells["Q_DRPIReferenceValue"].Value == null || dgvControlAverageA.Rows[i].Cells["Q_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.0M : Convert.ToDecimal(dgvControlAverageA.Rows[i].Cells["Q_DRPIReferenceValue"].Value.ToString().Trim());
                    }
                    else if (dgvControlAverageA.Rows[i].Cells["CoilCabinetName"].Value.ToString().Trim() == "A2")
                    {
                        dA_ControlRodAveRdc[1] = dgvControlAverageA.Rows[i].Cells["Rdc_DRPIReferenceValue"].Value == null || dgvControlAverageA.Rows[i].Cells["Rdc_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(dgvControlAverageA.Rows[i].Cells["Rdc_DRPIReferenceValue"].Value.ToString().Trim());
                        dA_ControlRodAveRac[1] = dgvControlAverageA.Rows[i].Cells["Rac_DRPIReferenceValue"].Value == null || dgvControlAverageA.Rows[i].Cells["Rac_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(dgvControlAverageA.Rows[i].Cells["Rac_DRPIReferenceValue"].Value.ToString().Trim());
                        dA_ControlRodAveL[1] = dgvControlAverageA.Rows[i].Cells["L_DRPIReferenceValue"].Value == null || dgvControlAverageA.Rows[i].Cells["L_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(dgvControlAverageA.Rows[i].Cells["L_DRPIReferenceValue"].Value.ToString().Trim());
                        dA_ControlRodAveC[1] = dgvControlAverageA.Rows[i].Cells["C_DRPIReferenceValue"].Value == null || dgvControlAverageA.Rows[i].Cells["C_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.000000M : Convert.ToDecimal(dgvControlAverageA.Rows[i].Cells["C_DRPIReferenceValue"].Value.ToString().Trim());
                        dA_ControlRodAveQ[1] = dgvControlAverageA.Rows[i].Cells["Q_DRPIReferenceValue"].Value == null || dgvControlAverageA.Rows[i].Cells["Q_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.0M : Convert.ToDecimal(dgvControlAverageA.Rows[i].Cells["Q_DRPIReferenceValue"].Value.ToString().Trim());
                    }
                    else
                    {
                        dcmA3RdcSum += dgvControlAverageA.Rows[i].Cells["Rdc_DRPIReferenceValue"].Value == null || dgvControlAverageA.Rows[i].Cells["Rdc_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(dgvControlAverageA.Rows[i].Cells["Rdc_DRPIReferenceValue"].Value.ToString().Trim());
                        dcmA3RacSum += dgvControlAverageA.Rows[i].Cells["Rac_DRPIReferenceValue"].Value == null || dgvControlAverageA.Rows[i].Cells["Rac_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(dgvControlAverageA.Rows[i].Cells["Rac_DRPIReferenceValue"].Value.ToString().Trim());
                        dcmA3LSum += dgvControlAverageA.Rows[i].Cells["L_DRPIReferenceValue"].Value == null || dgvControlAverageA.Rows[i].Cells["L_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(dgvControlAverageA.Rows[i].Cells["L_DRPIReferenceValue"].Value.ToString().Trim());
                        dcmA3CSum += dgvControlAverageA.Rows[i].Cells["C_DRPIReferenceValue"].Value == null || dgvControlAverageA.Rows[i].Cells["C_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.000000M : Convert.ToDecimal(dgvControlAverageA.Rows[i].Cells["C_DRPIReferenceValue"].Value.ToString().Trim());
                        dcmA3QSum += dgvControlAverageA.Rows[i].Cells["Q_DRPIReferenceValue"].Value == null || dgvControlAverageA.Rows[i].Cells["Q_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.0M : Convert.ToDecimal(dgvControlAverageA.Rows[i].Cells["Q_DRPIReferenceValue"].Value.ToString().Trim());
                        intA3SCount++;
                    }
                }

                if (intA3SCount > 0)
                {
                    dA_ControlRodAveRdc[2] = dcmA3RdcSum / intA3SCount;
                    dA_ControlRodAveRac[2] = dcmA3RacSum / intA3SCount;
                    dA_ControlRodAveL[2] = dcmA3LSum / intA3SCount;
                    dA_ControlRodAveC[2] = dcmA3CSum / intA3SCount;
                    dA_ControlRodAveQ[2] = dcmA3QSum / intA3SCount;
                }
                else
                {
                    dA_ControlRodAveRdc[2] = 0.000M;
                    dA_ControlRodAveRac[2] = 0.000M;
                    dA_ControlRodAveL[2] = 0.000M;
                    dA_ControlRodAveC[2] = 0.000000M;
                    dA_ControlRodAveQ[2] = 0.000M;
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
                    int intControlRowB = 0;
                    if (intBAvgRdcCount > 0 || intBAvgRacCount > 0 || intBAvgLCount > 0 || intBAvgCCount > 0 || intBAvgQCount > 0)
                    {
                        intControlRowB = dgvControlAverageB.Rows.Add();

                        dgvControlAverageB.Rows[intControlRowB].Cells["CoilCabinetType"].Value = "제어용";
                        dgvControlAverageB.Rows[intControlRowB].Cells["CoilCabinetName"].Value = "B" + Regex.Replace(strCoilName.Trim(), @"\D", "");
                    }

                    if (intBAvgRdcCount > 0)
                        dgvControlAverageB.Rows[intControlRowB].Cells["Rdc_DRPIReferenceValue"].Value = (dcmBSumControlRodRdc / intBAvgRdcCount).ToString("F3").Trim();
                    else
                        dgvControlAverageB.Rows[intControlRowB].Cells["Rdc_DRPIReferenceValue"].Value = "0.000";

                    if (intBAvgRacCount > 0)
                        dgvControlAverageB.Rows[intControlRowB].Cells["Rac_DRPIReferenceValue"].Value = (dcmBSumControlRodRac / intBAvgRacCount).ToString("F3").Trim();
                    else
                        dgvControlAverageB.Rows[intControlRowB].Cells["Rac_DRPIReferenceValue"].Value = "0.000";

                    if (intBAvgLCount > 0)
                        dgvControlAverageB.Rows[intControlRowB].Cells["L_DRPIReferenceValue"].Value = (dcmBSumControlRodL / intBAvgLCount).ToString("F3").Trim();
                    else
                        dgvControlAverageB.Rows[intControlRowB].Cells["L_DRPIReferenceValue"].Value = "0.000";

                    if (intBAvgCCount > 0)
                        dgvControlAverageB.Rows[intControlRowB].Cells["C_DRPIReferenceValue"].Value = (dcmBSumControlRodC / intBAvgCCount).ToString("F6").Trim();
                    else
                        dgvControlAverageB.Rows[intControlRowB].Cells["C_DRPIReferenceValue"].Value = "0.000";

                    if (intBAvgQCount > 0)
                        dgvControlAverageB.Rows[intControlRowB].Cells["Q_DRPIReferenceValue"].Value = (dcmBSumControlRodQ / intBAvgQCount).ToString("F3").Trim();
                    else
                        dgvControlAverageB.Rows[intControlRowB].Cells["Q_DRPIReferenceValue"].Value = "0.000";
                }

                decimal dcmB3RdcSum = 0M, dcmB3RacSum = 0M, dcmB3LSum = 0M, dcmB3CSum = 0M, dcmB3QSum = 0M;
                int intB3SCount = 0;

                for (int i = 0; i < dgvControlAverageB.RowCount; i++)
                {
                    if (dgvControlAverageB.Rows[i].Cells["CoilCabinetName"].Value.ToString().Trim() == "B1")
                    {
                        dB_ControlRodAveRdc[0] = dgvControlAverageB.Rows[i].Cells["Rdc_DRPIReferenceValue"].Value == null || dgvControlAverageB.Rows[i].Cells["Rdc_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(dgvControlAverageB.Rows[i].Cells["Rdc_DRPIReferenceValue"].Value.ToString().Trim());
                        dB_ControlRodAveRac[0] = dgvControlAverageB.Rows[i].Cells["Rac_DRPIReferenceValue"].Value == null || dgvControlAverageB.Rows[i].Cells["Rac_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(dgvControlAverageB.Rows[i].Cells["Rac_DRPIReferenceValue"].Value.ToString().Trim());
                        dB_ControlRodAveL[0] = dgvControlAverageB.Rows[i].Cells["L_DRPIReferenceValue"].Value == null || dgvControlAverageB.Rows[i].Cells["L_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(dgvControlAverageB.Rows[i].Cells["L_DRPIReferenceValue"].Value.ToString().Trim());
                        dB_ControlRodAveC[0] = dgvControlAverageB.Rows[i].Cells["C_DRPIReferenceValue"].Value == null || dgvControlAverageB.Rows[i].Cells["C_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.000000M : Convert.ToDecimal(dgvControlAverageB.Rows[i].Cells["C_DRPIReferenceValue"].Value.ToString().Trim());
                        dB_ControlRodAveQ[0] = dgvControlAverageB.Rows[i].Cells["Q_DRPIReferenceValue"].Value == null || dgvControlAverageB.Rows[i].Cells["Q_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.0M : Convert.ToDecimal(dgvControlAverageB.Rows[i].Cells["Q_DRPIReferenceValue"].Value.ToString().Trim());
                    }
                    else if (dgvControlAverageB.Rows[i].Cells["CoilCabinetName"].Value.ToString().Trim() == "B2")
                    {
                        dB_ControlRodAveRdc[1] = dgvControlAverageB.Rows[i].Cells["Rdc_DRPIReferenceValue"].Value == null || dgvControlAverageB.Rows[i].Cells["Rdc_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(dgvControlAverageB.Rows[i].Cells["Rdc_DRPIReferenceValue"].Value.ToString().Trim());
                        dB_ControlRodAveRac[1] = dgvControlAverageB.Rows[i].Cells["Rac_DRPIReferenceValue"].Value == null || dgvControlAverageB.Rows[i].Cells["Rac_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(dgvControlAverageB.Rows[i].Cells["Rac_DRPIReferenceValue"].Value.ToString().Trim());
                        dB_ControlRodAveL[1] = dgvControlAverageB.Rows[i].Cells["L_DRPIReferenceValue"].Value == null || dgvControlAverageB.Rows[i].Cells["L_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(dgvControlAverageB.Rows[i].Cells["L_DRPIReferenceValue"].Value.ToString().Trim());
                        dB_ControlRodAveC[1] = dgvControlAverageB.Rows[i].Cells["C_DRPIReferenceValue"].Value == null || dgvControlAverageB.Rows[i].Cells["C_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.000000M : Convert.ToDecimal(dgvControlAverageB.Rows[i].Cells["C_DRPIReferenceValue"].Value.ToString().Trim());
                        dB_ControlRodAveQ[1] = dgvControlAverageB.Rows[i].Cells["Q_DRPIReferenceValue"].Value == null || dgvControlAverageB.Rows[i].Cells["Q_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.0M : Convert.ToDecimal(dgvControlAverageB.Rows[i].Cells["Q_DRPIReferenceValue"].Value.ToString().Trim());
                    }
                    else
                    {
                        dcmB3RdcSum += dgvControlAverageB.Rows[i].Cells["Rdc_DRPIReferenceValue"].Value == null || dgvControlAverageB.Rows[i].Cells["Rdc_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(dgvControlAverageB.Rows[i].Cells["Rdc_DRPIReferenceValue"].Value.ToString().Trim());
                        dcmB3RacSum += dgvControlAverageB.Rows[i].Cells["Rac_DRPIReferenceValue"].Value == null || dgvControlAverageB.Rows[i].Cells["Rac_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(dgvControlAverageB.Rows[i].Cells["Rac_DRPIReferenceValue"].Value.ToString().Trim());
                        dcmB3LSum += dgvControlAverageB.Rows[i].Cells["L_DRPIReferenceValue"].Value == null || dgvControlAverageB.Rows[i].Cells["L_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(dgvControlAverageB.Rows[i].Cells["L_DRPIReferenceValue"].Value.ToString().Trim());
                        dcmB3CSum += dgvControlAverageB.Rows[i].Cells["C_DRPIReferenceValue"].Value == null || dgvControlAverageB.Rows[i].Cells["C_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.000000M : Convert.ToDecimal(dgvControlAverageB.Rows[i].Cells["C_DRPIReferenceValue"].Value.ToString().Trim());
                        dcmB3QSum += dgvControlAverageB.Rows[i].Cells["Q_DRPIReferenceValue"].Value == null || dgvControlAverageB.Rows[i].Cells["Q_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.0M : Convert.ToDecimal(dgvControlAverageB.Rows[i].Cells["Q_DRPIReferenceValue"].Value.ToString().Trim());
                        intB3SCount++;
                    }
                }

                if (intB3SCount > 0)
                {
                    dB_ControlRodAveRdc[2] = dcmB3RdcSum / intB3SCount;
                    dB_ControlRodAveRac[2] = dcmB3RacSum / intB3SCount;
                    dB_ControlRodAveL[2] = dcmB3LSum / intB3SCount;
                    dB_ControlRodAveC[2] = dcmB3CSum / intB3SCount;
                    dB_ControlRodAveQ[2] = dcmB3QSum / intB3SCount;
                }
                else
                {
                    dB_ControlRodAveRdc[2] = 0.000M;
                    dB_ControlRodAveRac[2] = 0.000M;
                    dB_ControlRodAveL[2] = 0.000M;
                    dB_ControlRodAveC[2] = 0.000000M;
                    dB_ControlRodAveQ[2] = 0.000M;
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
                    int intStopRowA = 0;
                    if (intAAvgRdcCount > 0 || intAAvgRacCount > 0 || intAAvgLCount > 0 || intAAvgCCount > 0 || intAAvgQCount > 0)
                    {
                        intStopRowA = dgvStopAverageA.Rows.Add();

                        dgvStopAverageA.Rows[intStopRowA].Cells["CoilCabinetType"].Value = "정지용";
                        dgvStopAverageA.Rows[intStopRowA].Cells["CoilCabinetName"].Value = "A" + Regex.Replace(strCoilName.Trim(), @"\D", "");
                    }

                    if (intAAvgRdcCount > 0)
                        dgvStopAverageA.Rows[intStopRowA].Cells["Rdc_DRPIReferenceValue"].Value = (dcmASumStopRdc / intAAvgRdcCount).ToString("F3").Trim();
                    else
                        dgvStopAverageA.Rows[intStopRowA].Cells["Rdc_DRPIReferenceValue"].Value = "0.000";

                    if (intAAvgRacCount > 0)
                        dgvStopAverageA.Rows[intStopRowA].Cells["Rac_DRPIReferenceValue"].Value = (dcmASumStopRac / intAAvgRacCount).ToString("F3").Trim();
                    else
                        dgvStopAverageA.Rows[intStopRowA].Cells["Rac_DRPIReferenceValue"].Value = "0.000";

                    if (intAAvgLCount > 0)
                        dgvStopAverageA.Rows[intStopRowA].Cells["L_DRPIReferenceValue"].Value = (dcmASumStopL / intAAvgLCount).ToString("F3").Trim();
                    else
                        dgvStopAverageA.Rows[intStopRowA].Cells["L_DRPIReferenceValue"].Value = "0.000";

                    if (intAAvgCCount > 0)
                        dgvStopAverageA.Rows[intStopRowA].Cells["C_DRPIReferenceValue"].Value = (dcmASumStopC / intAAvgCCount).ToString("F6").Trim();
                    else
                        dgvStopAverageA.Rows[intStopRowA].Cells["C_DRPIReferenceValue"].Value = "0.000";

                    if (intAAvgQCount > 0)
                        dgvStopAverageA.Rows[intStopRowA].Cells["Q_DRPIReferenceValue"].Value = (dcmASumStopQ / intAAvgQCount).ToString("F3").Trim();
                    else
                        dgvStopAverageA.Rows[intStopRowA].Cells["Q_DRPIReferenceValue"].Value = "0.000";
                }

                decimal dcmA3RdcSum = 0M, dcmA3RacSum = 0M, dcmA3LSum = 0M, dcmA3CSum = 0M, dcmA3QSum = 0M;
                int intA3SCount = 0;

                for (int i = 0; i < dgvStopAverageA.RowCount; i++)
                {
                    if (dgvStopAverageA.Rows[i].Cells["CoilCabinetName"].Value.ToString().Trim() == "A1")
                    {
                        dA_StopAveRdc[0] = dgvStopAverageA.Rows[i].Cells["Rdc_DRPIReferenceValue"].Value == null || dgvStopAverageA.Rows[i].Cells["Rdc_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(dgvStopAverageA.Rows[i].Cells["Rdc_DRPIReferenceValue"].Value.ToString().Trim());
                        dA_StopAveRac[0] = dgvStopAverageA.Rows[i].Cells["Rac_DRPIReferenceValue"].Value == null || dgvStopAverageA.Rows[i].Cells["Rac_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(dgvStopAverageA.Rows[i].Cells["Rac_DRPIReferenceValue"].Value.ToString().Trim());
                        dA_StopAveL[0] = dgvStopAverageA.Rows[i].Cells["L_DRPIReferenceValue"].Value == null || dgvStopAverageA.Rows[i].Cells["L_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(dgvStopAverageA.Rows[i].Cells["L_DRPIReferenceValue"].Value.ToString().Trim());
                        dA_StopAveC[0] = dgvStopAverageA.Rows[i].Cells["C_DRPIReferenceValue"].Value == null || dgvStopAverageA.Rows[i].Cells["C_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.000000M : Convert.ToDecimal(dgvStopAverageA.Rows[i].Cells["C_DRPIReferenceValue"].Value.ToString().Trim());
                        dA_StopAveQ[0] = dgvStopAverageA.Rows[i].Cells["Q_DRPIReferenceValue"].Value == null || dgvStopAverageA.Rows[i].Cells["Q_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.0M : Convert.ToDecimal(dgvStopAverageA.Rows[i].Cells["Q_DRPIReferenceValue"].Value.ToString().Trim());
                    }
                    else if (dgvStopAverageA.Rows[i].Cells["CoilCabinetName"].Value.ToString().Trim() == "A2")
                    {
                        dA_StopAveRdc[1] = dgvStopAverageA.Rows[i].Cells["Rdc_DRPIReferenceValue"].Value == null || dgvStopAverageA.Rows[i].Cells["Rdc_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(dgvStopAverageA.Rows[i].Cells["Rdc_DRPIReferenceValue"].Value.ToString().Trim());
                        dA_StopAveRac[1] = dgvStopAverageA.Rows[i].Cells["Rac_DRPIReferenceValue"].Value == null || dgvStopAverageA.Rows[i].Cells["Rac_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(dgvStopAverageA.Rows[i].Cells["Rac_DRPIReferenceValue"].Value.ToString().Trim());
                        dA_StopAveL[1] = dgvStopAverageA.Rows[i].Cells["L_DRPIReferenceValue"].Value == null || dgvStopAverageA.Rows[i].Cells["L_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(dgvStopAverageA.Rows[i].Cells["L_DRPIReferenceValue"].Value.ToString().Trim());
                        dA_StopAveC[1] = dgvStopAverageA.Rows[i].Cells["C_DRPIReferenceValue"].Value == null || dgvStopAverageA.Rows[i].Cells["C_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.000000M : Convert.ToDecimal(dgvStopAverageA.Rows[i].Cells["C_DRPIReferenceValue"].Value.ToString().Trim());
                        dA_StopAveQ[1] = dgvStopAverageA.Rows[i].Cells["Q_DRPIReferenceValue"].Value == null || dgvStopAverageA.Rows[i].Cells["Q_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.0M : Convert.ToDecimal(dgvStopAverageA.Rows[i].Cells["Q_DRPIReferenceValue"].Value.ToString().Trim());
                    }
                    else
                    {
                        dcmA3RdcSum += dgvStopAverageA.Rows[i].Cells["Rdc_DRPIReferenceValue"].Value == null || dgvStopAverageA.Rows[i].Cells["Rdc_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(dgvStopAverageA.Rows[i].Cells["Rdc_DRPIReferenceValue"].Value.ToString().Trim());
                        dcmA3RacSum += dgvStopAverageA.Rows[i].Cells["Rac_DRPIReferenceValue"].Value == null || dgvStopAverageA.Rows[i].Cells["Rac_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(dgvStopAverageA.Rows[i].Cells["Rac_DRPIReferenceValue"].Value.ToString().Trim());
                        dcmA3LSum += dgvStopAverageA.Rows[i].Cells["L_DRPIReferenceValue"].Value == null || dgvStopAverageA.Rows[i].Cells["L_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(dgvStopAverageA.Rows[i].Cells["L_DRPIReferenceValue"].Value.ToString().Trim());
                        dcmA3CSum += dgvStopAverageA.Rows[i].Cells["C_DRPIReferenceValue"].Value == null || dgvStopAverageA.Rows[i].Cells["C_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.000000M : Convert.ToDecimal(dgvStopAverageA.Rows[i].Cells["C_DRPIReferenceValue"].Value.ToString().Trim());
                        dcmA3QSum += dgvStopAverageA.Rows[i].Cells["Q_DRPIReferenceValue"].Value == null || dgvStopAverageA.Rows[i].Cells["Q_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.0M : Convert.ToDecimal(dgvStopAverageA.Rows[i].Cells["Q_DRPIReferenceValue"].Value.ToString().Trim());
                        intA3SCount++;
                    }
                }

                if (intA3SCount > 0)
                {
                    dA_StopAveRdc[2] = dcmA3RdcSum / intA3SCount;
                    dA_StopAveRac[2] = dcmA3RacSum / intA3SCount;
                    dA_StopAveL[2] = dcmA3LSum / intA3SCount;
                    dA_StopAveC[2] = dcmA3CSum / intA3SCount;
                    dA_StopAveQ[2] = dcmA3QSum / intA3SCount;
                }
                else
                {
                    dA_StopAveRdc[2] = 0.000M;
                    dA_StopAveRac[2] = 0.000M;
                    dA_StopAveL[2] = 0.000M;
                    dA_StopAveC[2] = 0.000000M;
                    dA_StopAveQ[2] = 0.000M;
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
                    int intStopRowB = 0;
                    if (intBAvgRdcCount > 0 || intBAvgRacCount > 0 || intBAvgLCount > 0 || intBAvgCCount > 0 || intBAvgQCount > 0)
                    {
                        intStopRowB = dgvStopAverageB.Rows.Add();

                        dgvStopAverageB.Rows[intStopRowB].Cells["CoilCabinetType"].Value = "정지용";
                        dgvStopAverageB.Rows[intStopRowB].Cells["CoilCabinetName"].Value = "B" + Regex.Replace(strCoilName.Trim(), @"\D", "");
                    }

                    if (intBAvgRdcCount > 0)
                        dgvStopAverageB.Rows[intStopRowB].Cells["Rdc_DRPIReferenceValue"].Value = (dcmBSumStopRdc / intBAvgRdcCount).ToString("F3").Trim();
                    else
                        dgvStopAverageB.Rows[intStopRowB].Cells["Rdc_DRPIReferenceValue"].Value = "0.000";

                    if (intBAvgRacCount > 0)
                        dgvStopAverageB.Rows[intStopRowB].Cells["Rac_DRPIReferenceValue"].Value = (dcmBSumStopRac / intBAvgRacCount).ToString("F3").Trim();
                    else
                        dgvStopAverageB.Rows[intStopRowB].Cells["Rac_DRPIReferenceValue"].Value = "0.000";

                    if (intBAvgLCount > 0)
                        dgvStopAverageB.Rows[intStopRowB].Cells["L_DRPIReferenceValue"].Value = (dcmBSumStopL / intBAvgLCount).ToString("F3").Trim();
                    else
                        dgvStopAverageB.Rows[intStopRowB].Cells["L_DRPIReferenceValue"].Value = "0.000";

                    if (intBAvgCCount > 0)
                        dgvStopAverageB.Rows[intStopRowB].Cells["C_DRPIReferenceValue"].Value = (dcmBSumStopC / intBAvgCCount).ToString("F6").Trim();
                    else
                        dgvStopAverageB.Rows[intStopRowB].Cells["C_DRPIReferenceValue"].Value = "0.000";

                    if (intBAvgQCount > 0)
                        dgvStopAverageB.Rows[intStopRowB].Cells["Q_DRPIReferenceValue"].Value = (dcmBSumStopQ / intBAvgQCount).ToString("F3").Trim();
                    else
                        dgvStopAverageB.Rows[intStopRowB].Cells["Q_DRPIReferenceValue"].Value = "0.000";
                }

                decimal dcmB3RdcSum = 0M, dcmB3RacSum = 0M, dcmB3LSum = 0M, dcmB3CSum = 0M, dcmB3QSum = 0M;
                int intB3SCount = 0;

                for (int i = 0; i < dgvStopAverageB.RowCount; i++)
                {
                    if (dgvStopAverageB.Rows[i].Cells["CoilCabinetName"].Value.ToString().Trim() == "B1")
                    {
                        dB_StopAveRdc[0] = dgvStopAverageB.Rows[i].Cells["Rdc_DRPIReferenceValue"].Value == null || dgvStopAverageB.Rows[i].Cells["Rdc_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(dgvStopAverageB.Rows[i].Cells["Rdc_DRPIReferenceValue"].Value.ToString().Trim());
                        dB_StopAveRac[0] = dgvStopAverageB.Rows[i].Cells["Rac_DRPIReferenceValue"].Value == null || dgvStopAverageB.Rows[i].Cells["Rac_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(dgvStopAverageB.Rows[i].Cells["Rac_DRPIReferenceValue"].Value.ToString().Trim());
                        dB_StopAveL[0] = dgvStopAverageB.Rows[i].Cells["L_DRPIReferenceValue"].Value == null || dgvStopAverageB.Rows[i].Cells["L_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(dgvStopAverageB.Rows[i].Cells["L_DRPIReferenceValue"].Value.ToString().Trim());
                        dB_StopAveC[0] = dgvStopAverageB.Rows[i].Cells["C_DRPIReferenceValue"].Value == null || dgvStopAverageB.Rows[i].Cells["C_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.000000M : Convert.ToDecimal(dgvStopAverageB.Rows[i].Cells["C_DRPIReferenceValue"].Value.ToString().Trim());
                        dB_StopAveQ[0] = dgvStopAverageB.Rows[i].Cells["Q_DRPIReferenceValue"].Value == null || dgvStopAverageB.Rows[i].Cells["Q_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.0M : Convert.ToDecimal(dgvStopAverageB.Rows[i].Cells["Q_DRPIReferenceValue"].Value.ToString().Trim());
                    }
                    else if (dgvStopAverageB.Rows[i].Cells["CoilCabinetName"].Value.ToString().Trim() == "B2")
                    {
                        dB_StopAveRdc[1] = dgvStopAverageB.Rows[i].Cells["Rdc_DRPIReferenceValue"].Value == null || dgvStopAverageB.Rows[i].Cells["Rdc_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(dgvStopAverageB.Rows[i].Cells["Rdc_DRPIReferenceValue"].Value.ToString().Trim());
                        dB_StopAveRac[1] = dgvStopAverageB.Rows[i].Cells["Rac_DRPIReferenceValue"].Value == null || dgvStopAverageB.Rows[i].Cells["Rac_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(dgvStopAverageB.Rows[i].Cells["Rac_DRPIReferenceValue"].Value.ToString().Trim());
                        dB_StopAveL[1] = dgvStopAverageB.Rows[i].Cells["L_DRPIReferenceValue"].Value == null || dgvStopAverageB.Rows[i].Cells["L_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(dgvStopAverageB.Rows[i].Cells["L_DRPIReferenceValue"].Value.ToString().Trim());
                        dB_StopAveC[1] = dgvStopAverageB.Rows[i].Cells["C_DRPIReferenceValue"].Value == null || dgvStopAverageB.Rows[i].Cells["C_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.000000M : Convert.ToDecimal(dgvStopAverageB.Rows[i].Cells["C_DRPIReferenceValue"].Value.ToString().Trim());
                        dB_StopAveQ[1] = dgvStopAverageB.Rows[i].Cells["Q_DRPIReferenceValue"].Value == null || dgvStopAverageB.Rows[i].Cells["Q_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.0M : Convert.ToDecimal(dgvStopAverageB.Rows[i].Cells["Q_DRPIReferenceValue"].Value.ToString().Trim());
                    }
                    else
                    {
                        dcmB3RdcSum += dgvStopAverageB.Rows[i].Cells["Rdc_DRPIReferenceValue"].Value == null || dgvStopAverageB.Rows[i].Cells["Rdc_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(dgvStopAverageB.Rows[i].Cells["Rdc_DRPIReferenceValue"].Value.ToString().Trim());
                        dcmB3RacSum += dgvStopAverageB.Rows[i].Cells["Rac_DRPIReferenceValue"].Value == null || dgvStopAverageB.Rows[i].Cells["Rac_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(dgvStopAverageB.Rows[i].Cells["Rac_DRPIReferenceValue"].Value.ToString().Trim());
                        dcmB3LSum += dgvStopAverageB.Rows[i].Cells["L_DRPIReferenceValue"].Value == null || dgvStopAverageB.Rows[i].Cells["L_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(dgvStopAverageB.Rows[i].Cells["L_DRPIReferenceValue"].Value.ToString().Trim());
                        dcmB3CSum += dgvStopAverageB.Rows[i].Cells["C_DRPIReferenceValue"].Value == null || dgvStopAverageB.Rows[i].Cells["C_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.000000M : Convert.ToDecimal(dgvStopAverageB.Rows[i].Cells["C_DRPIReferenceValue"].Value.ToString().Trim());
                        dcmB3QSum += dgvStopAverageB.Rows[i].Cells["Q_DRPIReferenceValue"].Value == null || dgvStopAverageB.Rows[i].Cells["Q_DRPIReferenceValue"].Value.ToString().Trim() == ""
                            ? 0.0M : Convert.ToDecimal(dgvStopAverageB.Rows[i].Cells["Q_DRPIReferenceValue"].Value.ToString().Trim());
                        intB3SCount++;
                    }
                }

                if (intB3SCount > 0)
                {
                    dB_StopAveRdc[2] = dcmB3RdcSum / intB3SCount;
                    dB_StopAveRac[2] = dcmB3RacSum / intB3SCount;
                    dB_StopAveL[2] = dcmB3LSum / intB3SCount;
                    dB_StopAveC[2] = dcmB3CSum / intB3SCount;
                    dB_StopAveQ[2] = dcmB3QSum / intB3SCount;
                }
                else
                {
                    dB_StopAveRdc[2] = 0.000M;
                    dB_StopAveRac[2] = 0.000M;
                    dB_StopAveL[2] = 0.000M;
                    dB_StopAveC[2] = 0.000000M;
                    dB_StopAveQ[2] = 0.000M;
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
        /// <param name="_dgv"></param>
        /// <param name="_strCoilName"></param>
        /// <returns></returns>
        private decimal GetAverageValue(DataGridView _dgv, string _strCoilName, string strColumns)
        {
            decimal dcmResultValue = 0.000M;

            for (int i = 0; i < _dgv.RowCount; i++)
            {
                if (_dgv.Rows[i].Cells["CoilCabinetName"].Value.ToString().Trim() == _strCoilName.Trim() && strColumns.Trim() == "Rdc")
                { 
                    dcmResultValue = _dgv.Rows[i].Cells["Rdc_DRPIReferenceValue"].Value == null || _dgv.Rows[i].Cells["Rdc_DRPIReferenceValue"].Value.ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(_dgv.Rows[i].Cells["Rdc_DRPIReferenceValue"].Value.ToString().Trim());
                }
                else if (_dgv.Rows[i].Cells["CoilCabinetName"].Value.ToString().Trim() == _strCoilName.Trim() && strColumns.Trim() == "Rac")
                {
                    dcmResultValue = _dgv.Rows[i].Cells["Rac_DRPIReferenceValue"].Value == null || _dgv.Rows[i].Cells["Rac_DRPIReferenceValue"].Value.ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(_dgv.Rows[i].Cells["Rac_DRPIReferenceValue"].Value.ToString().Trim());
                }
                else if (_dgv.Rows[i].Cells["CoilCabinetName"].Value.ToString().Trim() == _strCoilName.Trim() && strColumns.Trim() == "L")
                {
                    dcmResultValue = _dgv.Rows[i].Cells["L_DRPIReferenceValue"].Value == null || _dgv.Rows[i].Cells["L_DRPIReferenceValue"].Value.ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(_dgv.Rows[i].Cells["L_DRPIReferenceValue"].Value.ToString().Trim());
                }
                else if (_dgv.Rows[i].Cells["CoilCabinetName"].Value.ToString().Trim() == _strCoilName.Trim() && strColumns.Trim() == "C")
                {
                    dcmResultValue = _dgv.Rows[i].Cells["C_DRPIReferenceValue"].Value == null || _dgv.Rows[i].Cells["C_DRPIReferenceValue"].Value.ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(_dgv.Rows[i].Cells["C_DRPIReferenceValue"].Value.ToString().Trim());
                }
                else if (_dgv.Rows[i].Cells["CoilCabinetName"].Value.ToString().Trim() == _strCoilName.Trim() && strColumns.Trim() == "Q")
                {
                    dcmResultValue = _dgv.Rows[i].Cells["Q_DRPIReferenceValue"].Value == null || _dgv.Rows[i].Cells["Q_DRPIReferenceValue"].Value.ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(_dgv.Rows[i].Cells["Q_DRPIReferenceValue"].Value.ToString().Trim());
                }
            }

            return dcmResultValue;
        }

        /// <summary>
        /// 평균치 기준 값 가져오기
        /// </summary>
        /// <param name="_dgv"></param>
        /// <param name="_strCoilName"></param>
        /// <returns></returns>
        private decimal GetDRPIReferenceValue(decimal[] _dcm, string _strCoilName)
        {
            decimal dcmResultValue = 0.000M;

            if (_strCoilName.Trim() == "A1")
                dcmResultValue = _dcm[0];
            else if (_strCoilName.Trim() == "A2")
                dcmResultValue = _dcm[1];
            else 
                dcmResultValue = _dcm[2];

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
                    dgvMeasurement.Rows[iRow].Cells["DRPIType"].Value = _dt.Rows[i]["DRPIType"].ToString().Trim();
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
                                dcmDecision = dA_ControlRodAveRdc[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dA_ControlRodAveRdc[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "A")
                                dcmDecision = dA_ControlRodAveRdc[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dA_StopAveRdc[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dA_StopAveRdc[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "A")
                                dcmDecision = dA_StopAveRdc[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dB_ControlRodAveRdc[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dB_ControlRodAveRdc[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "B")
                                dcmDecision = dB_ControlRodAveRdc[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dB_StopAveRdc[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dB_StopAveRdc[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "B")
                                dcmDecision = dB_StopAveRdc[2];

                            strRdcResult = m_MeasureProcess.GetMesurmentValueDecision(dcmRdcValue, dcmDecision, dcmDecisionRange_Rdc, dcmEffectiveStandardRange, "DC");
                        }
                        else if (_strComparisonTarget.Trim() == "기준값")
                        {
                            if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dA_ControlRodRdc[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dA_ControlRodRdc[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "A")
                                dcmDecision = dA_ControlRodRdc[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dA_StopRdc[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dA_StopRdc[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "A")
                                dcmDecision = dA_StopRdc[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dB_ControlRodRdc[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dB_ControlRodRdc[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "B")
                                dcmDecision = dB_ControlRodRdc[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dB_StopRdc[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dB_StopRdc[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "B")
                                dcmDecision = dB_StopRdc[2];

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
                                dcmDecision = dA_ControlRodAveRac[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dA_ControlRodAveRac[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "A")
                                dcmDecision = dA_ControlRodAveRac[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dA_StopAveRac[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dA_StopAveRac[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "A")
                                dcmDecision = dA_StopAveRac[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dB_ControlRodAveRac[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dB_ControlRodAveRac[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "B")
                                dcmDecision = dB_ControlRodAveRac[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dB_StopAveRac[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dB_StopAveRac[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "B")
                                dcmDecision = dB_StopAveRac[2];

                            strRacResult = m_MeasureProcess.GetMesurmentValueDecision(dcmRacValue, dcmDecision, dcmDecisionRange_Rac, dcmEffectiveStandardRange, "AC");
                        }
                        else if (_strComparisonTarget.Trim() == "기준값")
                        {
                            if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dA_ControlRodRac[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dA_ControlRodRac[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "A")
                                dcmDecision = dA_ControlRodRac[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dA_StopRac[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dA_StopRac[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "A")
                                dcmDecision = dA_StopRac[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dB_ControlRodRac[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dB_ControlRodRac[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "B")
                                dcmDecision = dB_ControlRodRac[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dB_StopRac[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dB_StopRac[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "B")
                                dcmDecision = dB_StopRac[2];

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
                                dcmDecision = dA_ControlRodAveL[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dA_ControlRodAveL[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "A")
                                dcmDecision = dA_ControlRodAveL[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dA_StopAveL[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dA_StopAveL[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "A")
                                dcmDecision = dA_StopAveL[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dB_ControlRodAveL[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dB_ControlRodAveL[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "B")
                                dcmDecision = dB_ControlRodAveL[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dB_StopAveL[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dB_StopAveL[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "B")
                                dcmDecision = dB_StopAveL[2];

                            strLResult = m_MeasureProcess.GetMesurmentValueDecision(dcmLValue, dcmDecision, dcmDecisionRange_L, dcmEffectiveStandardRange, "L");
                        }
                        else if (_strComparisonTarget.Trim() == "기준값")
                        {
                            if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dA_ControlRodL[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dA_ControlRodL[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "A")
                                dcmDecision = dA_ControlRodL[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dA_StopL[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dA_StopL[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "A")
                                dcmDecision = dA_StopL[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dB_ControlRodL[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dB_ControlRodL[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "B")
                                dcmDecision = dB_ControlRodL[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dB_StopL[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dB_StopL[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "B")
                                dcmDecision = dB_StopL[2];

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
                    dcmDecision = 0.000M;

                    if (_dt.Rows[i]["C_CapacitanceValue"] == null || _dt.Rows[i]["C_CapacitanceValue"].ToString().Trim() == "")
                    {
                        strCResult = "부적합";
                    }
                    else
                    {
                        if (_strComparisonTarget.Trim() == "평균값")
                        {
                            if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dA_ControlRodAveC[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dA_ControlRodAveC[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "A")
                                dcmDecision = dA_ControlRodAveC[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dA_StopAveC[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dA_StopAveC[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "A")
                                dcmDecision = dA_StopAveC[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dB_ControlRodAveC[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dB_ControlRodAveC[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "B")
                                dcmDecision = dB_ControlRodAveC[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dB_StopAveC[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dB_StopAveC[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "B")
                                dcmDecision = dB_StopAveC[2];

                            strCResult = m_MeasureProcess.GetMesurmentValueDecision(dcmCValue, dcmDecision, dcmDecisionRange_Q, dcmEffectiveStandardRange, "C");
                        }
                        else if (_strComparisonTarget.Trim() == "기준값")
                        {
                            if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dA_ControlRodC[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dA_ControlRodC[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "A")
                                dcmDecision = dA_ControlRodC[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dA_StopC[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dA_StopC[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "A")
                                dcmDecision = dA_StopC[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dB_ControlRodC[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dB_ControlRodC[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "B")
                                dcmDecision = dB_ControlRodC[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dB_StopC[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dB_StopC[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "B")
                                dcmDecision = dB_StopC[2];

                            strCResult = m_MeasureProcess.GetMesurmentValueDecision(dcmCValue, dcmDecision, dcmDecisionRange_Q, dcmEffectiveStandardRange, "C");
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
                                dcmDecision = dA_ControlRodAveQ[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dA_ControlRodAveQ[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "A")
                                dcmDecision = dA_ControlRodAveQ[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dA_StopAveQ[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dA_StopAveQ[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "A")
                                dcmDecision = dA_StopAveQ[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dB_ControlRodAveQ[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dB_ControlRodAveQ[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "B")
                                dcmDecision = dB_ControlRodAveQ[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dB_StopAveQ[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dB_StopAveQ[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "B")
                                dcmDecision = dB_StopAveQ[2];

                            strQResult = m_MeasureProcess.GetMesurmentValueDecision(dcmQValue, dcmDecision, dcmDecisionRange_Q, dcmEffectiveStandardRange, "Q");
                        }
                        else if (_strComparisonTarget.Trim() == "기준값")
                        {
                            if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dA_ControlRodQ[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dA_ControlRodQ[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "A")
                                dcmDecision = dA_ControlRodQ[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dA_StopQ[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "A2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dA_StopQ[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "A")
                                dcmDecision = dA_StopQ[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dB_ControlRodQ[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용")
                                dcmDecision = dB_ControlRodQ[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "제어용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "B")
                                dcmDecision = dB_ControlRodQ[2];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B1" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dB_StopQ[0];
                            else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "B2" && _dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용")
                                dcmDecision = dB_StopQ[1];
                            else if (_dt.Rows[i]["DRPIType"].ToString().Trim() == "정지용" && _dt.Rows[i]["DRPIGroup"].ToString().Trim() == "B")
                                dcmDecision = dB_StopQ[2];

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
                        strRdcResult = "부적합";
                        dgvMeasurement.Rows[iRow].Cells["Result"].Value = strRdcResult.Trim();
                        dgvMeasurement.Rows[iRow].Cells["Result"].Style.ForeColor = System.Drawing.Color.Red;
                    }
                    else if (dgvMeasurement.Rows[iRow].Cells["DC_ResistanceValue"].Style.ForeColor == System.Drawing.Color.Blue
                        || dgvMeasurement.Rows[iRow].Cells["AC_ResistanceValue"].Style.ForeColor == System.Drawing.Color.Blue
                        || dgvMeasurement.Rows[iRow].Cells["L_InductanceValue"].Style.ForeColor == System.Drawing.Color.Blue
                        || dgvMeasurement.Rows[iRow].Cells["C_CapacitanceValue"].Style.ForeColor == System.Drawing.Color.Blue
                        || dgvMeasurement.Rows[iRow].Cells["Q_FactorValue"].Style.ForeColor == System.Drawing.Color.Blue)
                    {
                        strRdcResult = "의심";
                        dgvMeasurement.Rows[iRow].Cells["Result"].Value = strRdcResult.Trim();
                        dgvMeasurement.Rows[iRow].Cells["Result"].Style.ForeColor = System.Drawing.Color.Blue;
                    }
                    else
                    {
                        strRdcResult = "적합";
                        dgvMeasurement.Rows[iRow].Cells["Result"].Value = strRdcResult.Trim();
                        dgvMeasurement.Rows[iRow].Cells["Result"].Style.ForeColor = System.Drawing.Color.Black;
                    }

                    dgvMeasurement.Rows[iRow].Cells["Temperature_ReferenceValue"].Value = _dt.Rows[i]["Temperature_ReferenceValue"].ToString().Trim();
                    dgvMeasurement.Rows[iRow].Cells["Frequency"].Value = _dt.Rows[i]["Frequency"].ToString().Trim();
                    dgvMeasurement.Rows[iRow].Cells["VoltageLevel"].Value = _dt.Rows[i]["VoltageLevel"].ToString().Trim();
                }
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
                #region 변수 선언
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
                double[] dRefControlRodRdc = new double[_intXLabelValueCount];
                double[] dRefStopRdc = new double[_intXLabelValueCount];
                double[] dAveControlRodRdc = new double[_intXLabelValueCount];
                double[] dAveStopRdc = new double[_intXLabelValueCount];

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
                double[] dRefControlRodRac = new double[_intXLabelValueCount];
                double[] dRefStopRac = new double[_intXLabelValueCount];
                double[] dAveControlRodRac = new double[_intXLabelValueCount];
                double[] dAveStopRac = new double[_intXLabelValueCount];

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
                double[] dRefControlRodL = new double[_intXLabelValueCount];
                double[] dRefStopL = new double[_intXLabelValueCount];
                double[] dAveControlRodL = new double[_intXLabelValueCount];
                double[] dAveStopL = new double[_intXLabelValueCount];

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
                double[] dRefControlRodC = new double[_intXLabelValueCount];
                double[] dRefStopC = new double[_intXLabelValueCount];
                double[] dAveControlRodC = new double[_intXLabelValueCount];
                double[] dAveStopC = new double[_intXLabelValueCount];

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
                double[] dRefControlRodQ = new double[_intXLabelValueCount];
                double[] dRefStopQ = new double[_intXLabelValueCount];
                double[] dAveControlRodQ = new double[_intXLabelValueCount];
                double[] dAveStopQ = new double[_intXLabelValueCount];

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
                int intA1Index = 0, intA2Index = 0, intA3Index = 0, intA4Index = 0, intA5Index = 0, intA6Index = 0, intA7Index = 0
                    , intA8Index = 0, intA9Index = 0, intA10Index = 0, intA11Index = 0, intA12Index = 0, intA13Index = 0, intA14Index = 0
                    , intA15Index = 0, intA16Index = 0, intA17Index = 0, intA18Index = 0, intA19Index = 0, intA20Index = 0, intA21Index = 0;

                string[] arrayItemName = new string[_intXLabelValueCount];
                string strItemName = "", strSelectItemName = "";
                int intItemIndex = 0;
                #endregion

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

                double dRefGroupA_CtlDC = 0d, dRefGroupA_CtlAC = 0d, dRefGroupA_CtlL = 0d, dRefGroupA_CtlC = 0d, dRefGroupA_CtlQ = 0d
                    , dRefGroupB_CtlDC = 0d, dRefGroupB_CtlAC = 0d, dRefGroupB_CtlL = 0d, dRefGroupB_CtlC = 0d, dRefGroupB_CtlQ = 0d
                    , dRefGroupA_StopDC = 0d, dRefGroupA_StopAC = 0d, dRefGroupA_StopL = 0d, dRefGroupA_StopC = 0d, dRefGroupA_StopQ = 0d
                    , dRefGroupB_StopDC = 0d, dRefGroupB_StopAC = 0d, dRefGroupB_StopL = 0d, dRefGroupB_StopC = 0d, dRefGroupB_StopQ = 0d

                    , dAvgGroupA_CtlDC = 0d, dAvgGroupA_CtlAC = 0d, dAvgGroupA_CtlL = 0d, dAvgGroupA_CtlC = 0d, dAvgGroupA_CtlQ = 0d
                    , dAvgGroupB_CtlDC = 0d, dAvgGroupB_CtlAC = 0d, dAvgGroupB_CtlL = 0d, dAvgGroupB_CtlC = 0d, dAvgGroupB_CtlQ = 0d
                    , dAvgGroupA_StopDC = 0d, dAvgGroupA_StopAC = 0d, dAvgGroupA_StopL = 0d, dAvgGroupA_StopC = 0d, dAvgGroupA_StopQ = 0d
                    , dAvgGroupB_StopDC = 0d, dAvgGroupB_StopAC = 0d, dAvgGroupB_StopL = 0d, dAvgGroupB_StopC = 0d, dAvgGroupB_StopQ = 0d;
                
                #region 기준 값
                #region 그룹 A - 제어용
                dRefGroupA_CtlDC = teA1RdcControlRod_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA1RdcControlRod_DRPIReferenceValue.Text.Trim());
                dRefGroupA_CtlDC = dRefGroupA_CtlDC + (teA2RdcControlRod_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA2RdcControlRod_DRPIReferenceValue.Text.Trim()));
                dRefGroupA_CtlDC = dRefGroupA_CtlDC + (teA3RdcControlRod_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA3RdcControlRod_DRPIReferenceValue.Text.Trim()));
                dRefGroupA_CtlDC = dRefGroupA_CtlDC / 3;

                dRefGroupA_CtlAC = teA1RacControlRod_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA1RacControlRod_DRPIReferenceValue.Text.Trim());
                dRefGroupA_CtlAC = dRefGroupA_CtlAC + (teA2RacControlRod_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA2RacControlRod_DRPIReferenceValue.Text.Trim()));
                dRefGroupA_CtlAC = dRefGroupA_CtlAC + (teA3RacControlRod_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA3RacControlRod_DRPIReferenceValue.Text.Trim()));
                dRefGroupA_CtlAC = dRefGroupA_CtlAC / 3;

                dRefGroupA_CtlL = teA1LControlRod_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA1LControlRod_DRPIReferenceValue.Text.Trim());
                dRefGroupA_CtlL = dRefGroupA_CtlL + (teA2LControlRod_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA2LControlRod_DRPIReferenceValue.Text.Trim()));
                dRefGroupA_CtlL = dRefGroupA_CtlL + (teA3LControlRod_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA3LControlRod_DRPIReferenceValue.Text.Trim()));
                dRefGroupA_CtlL = dRefGroupA_CtlL / 3;

                dRefGroupA_CtlC = teA1CControlRod_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA1CControlRod_DRPIReferenceValue.Text.Trim());
                dRefGroupA_CtlC = dRefGroupA_CtlC + (teA2CControlRod_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA2CControlRod_DRPIReferenceValue.Text.Trim()));
                dRefGroupA_CtlC = dRefGroupA_CtlC + (teA3CControlRod_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA3CControlRod_DRPIReferenceValue.Text.Trim()));
                dRefGroupA_CtlC = dRefGroupA_CtlC / 3;

                dRefGroupA_CtlQ = teA1QControlRod_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA1QControlRod_DRPIReferenceValue.Text.Trim());
                dRefGroupA_CtlQ = dRefGroupA_CtlQ + (teA2QControlRod_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA2QControlRod_DRPIReferenceValue.Text.Trim()));
                dRefGroupA_CtlQ = dRefGroupA_CtlQ + (teA3QControlRod_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA3QControlRod_DRPIReferenceValue.Text.Trim()));
                dRefGroupA_CtlQ = dRefGroupA_CtlQ / 3;
                #endregion

                #region 그룹 B - 제어용
                dRefGroupB_CtlDC = teB1RdcControlRod_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB1RdcControlRod_DRPIReferenceValue.Text.Trim());
                dRefGroupB_CtlDC = dRefGroupB_CtlDC + (teB2RdcControlRod_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB2RdcControlRod_DRPIReferenceValue.Text.Trim()));
                dRefGroupB_CtlDC = dRefGroupB_CtlDC + (teB3RdcControlRod_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB3RdcControlRod_DRPIReferenceValue.Text.Trim()));
                dRefGroupB_CtlDC = dRefGroupB_CtlDC / 3;

                dRefGroupB_CtlAC = teB1RacControlRod_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB1RacControlRod_DRPIReferenceValue.Text.Trim());
                dRefGroupB_CtlAC = dRefGroupB_CtlAC + (teB2RacControlRod_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB2RacControlRod_DRPIReferenceValue.Text.Trim()));
                dRefGroupB_CtlAC = dRefGroupB_CtlAC + (teB3RacControlRod_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB3RacControlRod_DRPIReferenceValue.Text.Trim()));
                dRefGroupB_CtlAC = dRefGroupB_CtlAC / 3;

                dRefGroupB_CtlL = teB1LControlRod_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB1LControlRod_DRPIReferenceValue.Text.Trim());
                dRefGroupB_CtlL = dRefGroupB_CtlL + (teB2LControlRod_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB2LControlRod_DRPIReferenceValue.Text.Trim()));
                dRefGroupB_CtlL = dRefGroupB_CtlL + (teB3LControlRod_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB3LControlRod_DRPIReferenceValue.Text.Trim()));
                dRefGroupB_CtlL = dRefGroupB_CtlL / 3;

                dRefGroupB_CtlC = teB1CControlRod_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB1CControlRod_DRPIReferenceValue.Text.Trim());
                dRefGroupB_CtlC = dRefGroupB_CtlC + (teB2CControlRod_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB2CControlRod_DRPIReferenceValue.Text.Trim()));
                dRefGroupB_CtlC = dRefGroupB_CtlC + (teB3CControlRod_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB3CControlRod_DRPIReferenceValue.Text.Trim()));
                dRefGroupB_CtlC = dRefGroupB_CtlC / 3;

                dRefGroupB_CtlQ = teB1QControlRod_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB1QControlRod_DRPIReferenceValue.Text.Trim());
                dRefGroupB_CtlQ = dRefGroupB_CtlQ + (teB2QControlRod_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB2QControlRod_DRPIReferenceValue.Text.Trim()));
                dRefGroupB_CtlQ = dRefGroupB_CtlQ + (teB3QControlRod_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB3QControlRod_DRPIReferenceValue.Text.Trim()));
                dRefGroupB_CtlQ = dRefGroupB_CtlQ / 3;
                #endregion

                #region 그룹 A - 정지용
                dRefGroupA_StopDC = teA1RdcStop_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA1RdcStop_DRPIReferenceValue.Text.Trim());
                dRefGroupA_StopDC = dRefGroupA_StopDC + (teA2RdcStop_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA2RdcStop_DRPIReferenceValue.Text.Trim()));
                dRefGroupA_StopDC = dRefGroupA_StopDC + (teA3RdcStop_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA3RdcStop_DRPIReferenceValue.Text.Trim()));
                dRefGroupA_StopDC = dRefGroupA_StopDC / 3;

                dRefGroupA_StopAC = teA1RacStop_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA1RacStop_DRPIReferenceValue.Text.Trim());
                dRefGroupA_StopAC = dRefGroupA_StopAC + (teA2RacStop_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA2RacStop_DRPIReferenceValue.Text.Trim()));
                dRefGroupA_StopAC = dRefGroupA_StopAC + (teA3RacStop_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA3RacStop_DRPIReferenceValue.Text.Trim()));
                dRefGroupA_StopAC = dRefGroupA_StopAC / 3;

                dRefGroupA_StopL = teA1LStop_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA1LStop_DRPIReferenceValue.Text.Trim());
                dRefGroupA_StopL = dRefGroupA_StopL + (teA2LStop_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA2LStop_DRPIReferenceValue.Text.Trim()));
                dRefGroupA_StopL = dRefGroupA_StopL + (teA3LStop_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA3LStop_DRPIReferenceValue.Text.Trim()));
                dRefGroupA_StopL = dRefGroupA_StopL / 3;

                dRefGroupA_StopC = teA1CStop_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA1CStop_DRPIReferenceValue.Text.Trim());
                dRefGroupA_StopC = dRefGroupA_StopC + (teA2CStop_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA2CStop_DRPIReferenceValue.Text.Trim()));
                dRefGroupA_StopC = dRefGroupA_StopC + (teA3CStop_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA3CStop_DRPIReferenceValue.Text.Trim()));
                dRefGroupA_StopC = dRefGroupA_StopC / 3;

                dRefGroupA_StopQ = teA1QStop_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA1QStop_DRPIReferenceValue.Text.Trim());
                dRefGroupA_StopQ = dRefGroupA_StopQ + (teA2QStop_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA2QStop_DRPIReferenceValue.Text.Trim()));
                dRefGroupA_StopQ = dRefGroupA_StopQ + (teA3QStop_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA3QStop_DRPIReferenceValue.Text.Trim()));
                dRefGroupA_StopQ = dRefGroupA_StopQ / 3;
                #endregion

                #region 그룹 B - 정지용
                dRefGroupB_StopDC = teB1RdcStop_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB1RdcStop_DRPIReferenceValue.Text.Trim());
                dRefGroupB_StopDC = dRefGroupB_StopDC + (teB2RdcStop_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB2RdcStop_DRPIReferenceValue.Text.Trim()));
                dRefGroupB_StopDC = dRefGroupB_StopDC + (teB3RdcStop_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB3RdcStop_DRPIReferenceValue.Text.Trim()));
                dRefGroupB_StopDC = dRefGroupB_StopDC / 3;

                dRefGroupB_StopAC = teB1RacStop_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB1RacStop_DRPIReferenceValue.Text.Trim());
                dRefGroupB_StopAC = dRefGroupB_StopAC + (teB2RacStop_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB2RacStop_DRPIReferenceValue.Text.Trim()));
                dRefGroupB_StopAC = dRefGroupB_StopAC + (teB3RacStop_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB3RacStop_DRPIReferenceValue.Text.Trim()));
                dRefGroupB_StopAC = dRefGroupB_StopAC / 3;

                dRefGroupB_StopL = teB1LStop_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB1LStop_DRPIReferenceValue.Text.Trim());
                dRefGroupB_StopL = dRefGroupB_StopL + (teB2LStop_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB2LStop_DRPIReferenceValue.Text.Trim()));
                dRefGroupB_StopL = dRefGroupB_StopL + (teB3LStop_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB3LStop_DRPIReferenceValue.Text.Trim()));
                dRefGroupB_StopL = dRefGroupB_StopL / 3;

                dRefGroupB_StopC = teB1CStop_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB1CStop_DRPIReferenceValue.Text.Trim());
                dRefGroupB_StopC = dRefGroupB_StopC + (teB2CStop_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB2CStop_DRPIReferenceValue.Text.Trim()));
                dRefGroupB_StopC = dRefGroupB_StopC + (teB3CStop_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB3CStop_DRPIReferenceValue.Text.Trim()));
                dRefGroupB_StopC = dRefGroupB_StopC / 3;

                dRefGroupB_StopQ = teB1QStop_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB1QStop_DRPIReferenceValue.Text.Trim());
                dRefGroupB_StopQ = dRefGroupB_StopQ + (teB2QStop_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB2QStop_DRPIReferenceValue.Text.Trim()));
                dRefGroupB_StopQ = dRefGroupB_StopQ + (teB3QStop_DRPIReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB3QStop_DRPIReferenceValue.Text.Trim()));
                dRefGroupB_StopQ = dRefGroupB_StopQ / 3;
                #endregion
                #endregion

                #region 평균 값
                #region 그룹 A - 제어용
                dAvgGroupA_CtlDC = teA1RdcControlRod_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA1RdcControlRod_DRPIAverageValue.Text.Trim());
                dAvgGroupA_CtlDC = dAvgGroupA_CtlDC + (teA2RdcControlRod_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA2RdcControlRod_DRPIAverageValue.Text.Trim()));
                dAvgGroupA_CtlDC = dAvgGroupA_CtlDC + (teA3RdcControlRod_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA3RdcControlRod_DRPIAverageValue.Text.Trim()));
                dAvgGroupA_CtlDC = dAvgGroupA_CtlDC / 3;

                dAvgGroupA_CtlAC = teA1RacControlRod_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA1RacControlRod_DRPIAverageValue.Text.Trim());
                dAvgGroupA_CtlAC = dAvgGroupA_CtlAC + (teA2RacControlRod_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA2RacControlRod_DRPIAverageValue.Text.Trim()));
                dAvgGroupA_CtlAC = dAvgGroupA_CtlAC + (teA3RacControlRod_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA3RacControlRod_DRPIAverageValue.Text.Trim()));
                dAvgGroupA_CtlAC = dAvgGroupA_CtlAC / 3;

                dAvgGroupA_CtlL = teA1LControlRod_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA1LControlRod_DRPIAverageValue.Text.Trim());
                dAvgGroupA_CtlL = dAvgGroupA_CtlL + (teA2LControlRod_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA2LControlRod_DRPIAverageValue.Text.Trim()));
                dAvgGroupA_CtlL = dAvgGroupA_CtlL + (teA3LControlRod_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA3LControlRod_DRPIAverageValue.Text.Trim()));
                dAvgGroupA_CtlL = dAvgGroupA_CtlL / 3;

                dAvgGroupA_CtlC = teA1CControlRod_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA1CControlRod_DRPIAverageValue.Text.Trim());
                dAvgGroupA_CtlC = dAvgGroupA_CtlC + (teA2CControlRod_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA2CControlRod_DRPIAverageValue.Text.Trim()));
                dAvgGroupA_CtlC = dAvgGroupA_CtlC + (teA3CControlRod_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA3CControlRod_DRPIAverageValue.Text.Trim()));
                dAvgGroupA_CtlC = dAvgGroupA_CtlC / 3;

                dAvgGroupA_CtlQ = teA1QControlRod_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA1QControlRod_DRPIAverageValue.Text.Trim());
                dAvgGroupA_CtlQ = dAvgGroupA_CtlQ + (teA2QControlRod_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA2QControlRod_DRPIAverageValue.Text.Trim()));
                dAvgGroupA_CtlQ = dAvgGroupA_CtlQ + (teA3QControlRod_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA3QControlRod_DRPIAverageValue.Text.Trim()));
                dAvgGroupA_CtlQ = dAvgGroupA_CtlQ / 3;
                #endregion

                #region 그룹 B - 제어용
                dAvgGroupB_CtlDC = teB1RdcControlRod_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB1RdcControlRod_DRPIAverageValue.Text.Trim());
                dAvgGroupB_CtlDC = dAvgGroupB_CtlDC + (teB2RdcControlRod_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB2RdcControlRod_DRPIAverageValue.Text.Trim()));
                dAvgGroupB_CtlDC = dAvgGroupB_CtlDC + (teB3RdcControlRod_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB3RdcControlRod_DRPIAverageValue.Text.Trim()));
                dAvgGroupB_CtlDC = dAvgGroupB_CtlDC / 3;

                dAvgGroupB_CtlAC = teB1RacControlRod_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB1RacControlRod_DRPIAverageValue.Text.Trim());
                dAvgGroupB_CtlAC = dAvgGroupB_CtlAC + (teB2RacControlRod_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB2RacControlRod_DRPIAverageValue.Text.Trim()));
                dAvgGroupB_CtlAC = dAvgGroupB_CtlAC + (teB3RacControlRod_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB3RacControlRod_DRPIAverageValue.Text.Trim()));
                dAvgGroupB_CtlAC = dAvgGroupB_CtlAC / 3;

                dAvgGroupB_CtlL = teB1LControlRod_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB1LControlRod_DRPIAverageValue.Text.Trim());
                dAvgGroupB_CtlL = dAvgGroupB_CtlL + (teB2LControlRod_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB2LControlRod_DRPIAverageValue.Text.Trim()));
                dAvgGroupB_CtlL = dAvgGroupB_CtlL + (teB3LControlRod_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB3LControlRod_DRPIAverageValue.Text.Trim()));
                dAvgGroupB_CtlL = dAvgGroupB_CtlL / 3;

                dAvgGroupB_CtlC = teB1CControlRod_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB1CControlRod_DRPIAverageValue.Text.Trim());
                dAvgGroupB_CtlC = dAvgGroupB_CtlC + (teB2CControlRod_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB2CControlRod_DRPIAverageValue.Text.Trim()));
                dAvgGroupB_CtlC = dAvgGroupB_CtlC + (teB3CControlRod_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB3CControlRod_DRPIAverageValue.Text.Trim()));
                dAvgGroupB_CtlC = dAvgGroupB_CtlC / 3;

                dAvgGroupB_CtlQ = teB1QControlRod_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB1QControlRod_DRPIAverageValue.Text.Trim());
                dAvgGroupB_CtlQ = dAvgGroupB_CtlQ + (teB2QControlRod_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB2QControlRod_DRPIAverageValue.Text.Trim()));
                dAvgGroupB_CtlQ = dAvgGroupB_CtlQ + (teB3QControlRod_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB3QControlRod_DRPIAverageValue.Text.Trim()));
                dAvgGroupB_CtlQ = dAvgGroupB_CtlQ / 3;
                #endregion

                #region 그룹 A - 정지용
                dAvgGroupA_StopDC = teA1RdcStop_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA1RdcStop_DRPIAverageValue.Text.Trim());
                dAvgGroupA_StopDC = dAvgGroupA_StopDC + (teA2RdcStop_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA2RdcStop_DRPIAverageValue.Text.Trim()));
                dAvgGroupA_StopDC = dAvgGroupA_StopDC + (teA3RdcStop_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA3RdcStop_DRPIAverageValue.Text.Trim()));
                dAvgGroupA_StopDC = dAvgGroupA_StopDC / 3;

                dAvgGroupA_StopAC = teA1RacStop_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA1RacStop_DRPIAverageValue.Text.Trim());
                dAvgGroupA_StopAC = dAvgGroupA_StopAC + (teA2RacStop_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA2RacStop_DRPIAverageValue.Text.Trim()));
                dAvgGroupA_StopAC = dAvgGroupA_StopAC + (teA3RacStop_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA3RacStop_DRPIAverageValue.Text.Trim()));
                dAvgGroupA_StopAC = dAvgGroupA_StopAC / 3;

                dAvgGroupA_StopL = teA1LStop_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA1LStop_DRPIAverageValue.Text.Trim());
                dAvgGroupA_StopL = dAvgGroupA_StopL + (teA2LStop_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA2LStop_DRPIAverageValue.Text.Trim()));
                dAvgGroupA_StopL = dAvgGroupA_StopL + (teA3LStop_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA3LStop_DRPIAverageValue.Text.Trim()));
                dAvgGroupA_StopL = dAvgGroupA_StopL / 3;

                dAvgGroupA_StopC = teA1CStop_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA1CStop_DRPIAverageValue.Text.Trim());
                dAvgGroupA_StopC = dAvgGroupA_StopC + (teA2CStop_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA2CStop_DRPIAverageValue.Text.Trim()));
                dAvgGroupA_StopC = dAvgGroupA_StopC + (teA3CStop_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA3CStop_DRPIAverageValue.Text.Trim()));
                dAvgGroupA_StopC = dAvgGroupA_StopC / 3;

                dAvgGroupA_StopQ = teA1QStop_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA1QStop_DRPIAverageValue.Text.Trim());
                dAvgGroupA_StopQ = dAvgGroupA_StopQ + (teA2QStop_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA2QStop_DRPIAverageValue.Text.Trim()));
                dAvgGroupA_StopQ = dAvgGroupA_StopQ + (teA3QStop_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teA3QStop_DRPIAverageValue.Text.Trim()));
                dAvgGroupA_StopQ = dAvgGroupA_StopQ / 3;
                #endregion

                #region 그룹 B - 정지용
                dAvgGroupB_StopDC = teB1RdcStop_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB1RdcStop_DRPIAverageValue.Text.Trim());
                dAvgGroupB_StopDC = dAvgGroupB_StopDC + (teB2RdcStop_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB2RdcStop_DRPIAverageValue.Text.Trim()));
                dAvgGroupB_StopDC = dAvgGroupB_StopDC + (teB3RdcStop_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB3RdcStop_DRPIAverageValue.Text.Trim()));
                dAvgGroupB_StopDC = dAvgGroupB_StopDC / 3;

                dAvgGroupB_StopAC = teB1RacStop_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB1RacStop_DRPIAverageValue.Text.Trim());
                dAvgGroupB_StopAC = dAvgGroupB_StopAC + (teB2RacStop_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB2RacStop_DRPIAverageValue.Text.Trim()));
                dAvgGroupB_StopAC = dAvgGroupB_StopAC + (teB3RacStop_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB3RacStop_DRPIAverageValue.Text.Trim()));
                dAvgGroupB_StopAC = dAvgGroupB_StopAC / 3;

                dAvgGroupB_StopL = teB1LStop_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB1LStop_DRPIAverageValue.Text.Trim());
                dAvgGroupB_StopL = dAvgGroupB_StopL + (teB2LStop_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB2LStop_DRPIAverageValue.Text.Trim()));
                dAvgGroupB_StopL = dAvgGroupB_StopL + (teB3LStop_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB3LStop_DRPIAverageValue.Text.Trim()));
                dAvgGroupB_StopL = dAvgGroupB_StopL / 3;

                dAvgGroupB_StopC = teB1CStop_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB1CStop_DRPIAverageValue.Text.Trim());
                dAvgGroupB_StopC = dAvgGroupB_StopC + (teB2CStop_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB2CStop_DRPIAverageValue.Text.Trim()));
                dAvgGroupB_StopC = dAvgGroupB_StopC + (teB3CStop_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB3CStop_DRPIAverageValue.Text.Trim()));
                dAvgGroupB_StopC = dAvgGroupB_StopC / 3;

                dAvgGroupB_StopQ = teB1QStop_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB1QStop_DRPIAverageValue.Text.Trim());
                dAvgGroupB_StopQ = dAvgGroupB_StopQ + (teB2QStop_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB2QStop_DRPIAverageValue.Text.Trim()));
                dAvgGroupB_StopQ = dAvgGroupB_StopQ + (teB3QStop_DRPIAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teB3QStop_DRPIAverageValue.Text.Trim()));
                dAvgGroupB_StopQ = dAvgGroupB_StopQ / 3;
                #endregion
                #endregion

                for (int ii = 0; ii < arrayItemName.Length; ii++)
                {
                    if (arrayItemName[ii] == null) continue;

                    string[] arrayData = arrayItemName[ii].Split('/');

                    if (arrayData.Length < 2) continue;

                    strSelectControlName = arrayData[0];
                    strSelectDRPIGroup = arrayData[1];

                    // 측정 및 편차 값 설정
                    for (int i = 0; i < _dt.Rows.Count; i++)
                    {
                        if (strSelectControlName != _dt.Rows[i]["ControlName"].ToString().Trim()) continue;
                        if (strSelectDRPIGroup != _dt.Rows[i]["DRPIGroup"].ToString().Trim()) continue;

                        #region 측정 및 편차 값 설정
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
                        #endregion
                    }
                    
                    #region 기준치 및 평균치 그래프 데이터 설정
                    if (strSelectDRPIGroup.Trim() == "A")
                    {
                        // 기준치 그래프 데이터로 변환
                        dRefControlRodRdc[ii] = dRefGroupA_CtlDC;
                        dRefStopRdc[ii] = dRefGroupA_StopDC;

                        // 평균치 그래프 데이터로 변환
                        dAveControlRodRdc[ii] = dAvgGroupA_CtlDC;
                        dAveStopRdc[ii] = dAvgGroupA_StopDC;

                        // 기준치 그래프 데이터로 변환
                        dRefControlRodRac[ii] = dRefGroupA_CtlAC;
                        dRefStopRac[ii] = dRefGroupA_StopAC;

                        // 평균치 그래프 데이터로 변환
                        dAveControlRodRac[ii] = dAvgGroupA_CtlAC;
                        dAveStopRac[ii] = dAvgGroupA_StopAC;

                        // 기준치 그래프 데이터로 변환
                        dRefControlRodL[ii] = dRefGroupA_CtlL;
                        dRefStopL[ii] = dRefGroupA_StopL;

                        // 평균치 그래프 데이터로 변환
                        dAveControlRodL[ii] = dAvgGroupA_CtlL;
                        dAveStopL[ii] = dAvgGroupA_StopL;

                        // 기준치 그래프 데이터로 변환
                        dRefControlRodC[ii] = dRefGroupA_CtlC;
                        dRefStopC[ii] = dRefGroupA_StopC;

                        // 평균치 그래프 데이터로 변환
                        dAveControlRodC[ii] = dAvgGroupA_CtlC;
                        dAveStopC[ii] = dAvgGroupA_StopC;

                        // 기준치 그래프 데이터로 변환
                        dRefControlRodQ[ii] = dRefGroupA_CtlQ;
                        dRefStopQ[ii] = dRefGroupA_StopQ;

                        // 평균치 그래프 데이터로 변환
                        dAveControlRodQ[ii] = dAvgGroupA_CtlQ;
                        dAveStopQ[ii] = dAvgGroupA_StopQ;
                    }
                    else
                    {
                        // 기준치 그래프 데이터로 변환
                        dRefControlRodRdc[ii] = dRefGroupB_CtlDC;
                        dRefStopRdc[ii] = dRefGroupB_StopDC;

                        // 평균치 그래프 데이터로 변환
                        dAveControlRodRdc[ii] = dAvgGroupB_CtlDC;
                        dAveStopRdc[ii] = dAvgGroupB_StopDC;

                        // 기준치 그래프 데이터로 변환
                        dRefControlRodRac[ii] = dRefGroupB_CtlAC;
                        dRefStopRac[ii] = dRefGroupB_StopAC;

                        // 평균치 그래프 데이터로 변환
                        dAveControlRodRac[ii] = dAvgGroupB_CtlAC;
                        dAveStopRac[ii] = dAvgGroupB_StopAC;

                        // 기준치 그래프 데이터로 변환
                        dRefControlRodL[ii] = dRefGroupB_CtlL;
                        dRefStopL[ii] = dRefGroupB_StopL;

                        // 평균치 그래프 데이터로 변환
                        dAveControlRodL[ii] = dAvgGroupB_CtlL;
                        dAveStopL[ii] = dAvgGroupB_StopL;

                        // 기준치 그래프 데이터로 변환
                        dRefControlRodC[ii] = dRefGroupB_CtlC;
                        dRefStopC[ii] = dRefGroupB_StopC;

                        // 평균치 그래프 데이터로 변환
                        dAveControlRodC[ii] = dAvgGroupB_CtlC;
                        dAveStopC[ii] = dAvgGroupB_StopC;

                        // 기준치 그래프 데이터로 변환
                        dRefControlRodQ[ii] = dRefGroupB_CtlQ;
                        dRefStopQ[ii] = dRefGroupB_StopQ;

                        // 평균치 그래프 데이터로 변환
                        dAveControlRodQ[ii] = dAvgGroupB_CtlQ;
                        dAveStopQ[ii] = dAvgGroupB_StopQ;
                    }
                    #endregion

                }

                // DC 저항 Chart 그리기  
                SetChart_Depict(chartMeasurementValueDC, "DC 저항", dA01Rdc, dA02Rdc, dA03Rdc, dA04Rdc, dA05Rdc, dA06Rdc, dA07Rdc, dA08Rdc, dA09Rdc, dA10Rdc
                    , dA11Rdc, dA12Rdc, dA13Rdc, dA14Rdc, dA15Rdc, dA16Rdc, dA17Rdc, dA18Rdc, dA19Rdc, dA20Rdc, dA21Rdc, dRefControlRodRdc, dRefStopRdc
                    , dAveControlRodRdc, dAveStopRdc, _arrayXLabelValue, 0, "", "");

                // AC 저항 Chart 그리기  
                SetChart_Depict(chartMeasurementValueAC, "AC 저항", dA01Rac, dA02Rac, dA03Rac, dA04Rac, dA05Rac, dA06Rac, dA07Rac, dA08Rac, dA09Rac, dA10Rac
                    , dA11Rac, dA12Rac, dA13Rac, dA14Rac, dA15Rac, dA16Rac, dA17Rac, dA18Rac, dA19Rac, dA20Rac, dA21Rac, dRefControlRodRac, dRefStopRac
                    , dAveControlRodRac, dAveStopRac, _arrayXLabelValue, 0, "", "");

                // L 저항 Chart 그리기  
                SetChart_Depict(chartMeasurementValueL, "인덕턴스", dA01L, dA02L, dA03L, dA04L, dA05L, dA06L, dA07L, dA08L, dA09L, dA10L
                    , dA11L, dA12L, dA13L, dA14L, dA15L, dA16L, dA17L, dA18L, dA19L, dA20L, dA21L, dRefControlRodL, dRefStopL
                    , dAveControlRodL, dAveStopL, _arrayXLabelValue, 0, "", "");

                // C 저항 Chart 그리기  
                SetChart_Depict(chartMeasurementValueC, "캐패시턴스", dA01C, dA02C, dA03C, dA04C, dA05C, dA06C, dA07C, dA08C, dA09C, dA10C
                    , dA11C, dA12C, dA13C, dA14C, dA15C, dA16C, dA17C, dA18C, dA19C, dA20C, dA21C, dRefControlRodC, dRefStopC
                    , dAveControlRodC, dAveStopC, _arrayXLabelValue, 0, "", "");

                // Q 저항 Chart 그리기  
                SetChart_Depict(chartMeasurementValueQ, "Q-Factor", dA01Q, dA02Q, dA03Q, dA04Q, dA05Q, dA06Q, dA07Q, dA08Q, dA09Q, dA10Q
                    , dA11Q, dA12Q, dA13Q, dA14Q, dA15Q, dA16Q, dA17Q, dA18Q, dA19Q, dA20Q, dA21Q, dRefControlRodQ, dRefStopQ
                    , dAveControlRodQ, dAveStopQ, _arrayXLabelValue, 0, "", "");

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

        /// <summary>
        /// 기준 값(DRPI) 초기화
        /// </summary>
        private void SetDRPIReferenceValueInitialize()
        {
            teA1RdcControlRod_DRPIReferenceValue.Text = "0.000";
            teA1RacControlRod_DRPIReferenceValue.Text = "0.000";
            teA1LControlRod_DRPIReferenceValue.Text = "0.000";
            teA1CControlRod_DRPIReferenceValue.Text = "0.000000";
            teA1QControlRod_DRPIReferenceValue.Text = "0.000";

            teA1RdcStop_DRPIReferenceValue.Text = "0.000";
            teA1RacStop_DRPIReferenceValue.Text = "0.000";
            teA1LStop_DRPIReferenceValue.Text = "0.000";
            teA1CStop_DRPIReferenceValue.Text = "0.000000";
            teA1QStop_DRPIReferenceValue.Text = "0.000";

            teA2RdcControlRod_DRPIReferenceValue.Text = "0.000";
            teA2RacControlRod_DRPIReferenceValue.Text = "0.000";
            teA2LControlRod_DRPIReferenceValue.Text = "0.000";
            teA2CControlRod_DRPIReferenceValue.Text = "0.000000";
            teA2QControlRod_DRPIReferenceValue.Text = "0.000";

            teA2RdcStop_DRPIReferenceValue.Text = "0.000";
            teA2RacStop_DRPIReferenceValue.Text = "0.000";
            teA2LStop_DRPIReferenceValue.Text = "0.000";
            teA2CStop_DRPIReferenceValue.Text = "0.000000";
            teA2QStop_DRPIReferenceValue.Text = "0.000";

            teA3RdcControlRod_DRPIReferenceValue.Text = "0.000";
            teA3RacControlRod_DRPIReferenceValue.Text = "0.000";
            teA3LControlRod_DRPIReferenceValue.Text = "0.000";
            teA3CControlRod_DRPIReferenceValue.Text = "0.000000";
            teA3QControlRod_DRPIReferenceValue.Text = "0.000";

            teA3RdcStop_DRPIReferenceValue.Text = "0.000";
            teA3RacStop_DRPIReferenceValue.Text = "0.000";
            teA3LStop_DRPIReferenceValue.Text = "0.000";
            teA3CStop_DRPIReferenceValue.Text = "0.000000";
            teA3QStop_DRPIReferenceValue.Text = "0.000";

            teB1RdcControlRod_DRPIReferenceValue.Text = "0.000";
            teB1RacControlRod_DRPIReferenceValue.Text = "0.000";
            teB1LControlRod_DRPIReferenceValue.Text = "0.000";
            teB1CControlRod_DRPIReferenceValue.Text = "0.000000";
            teB1QControlRod_DRPIReferenceValue.Text = "0.000";

            teB1RdcStop_DRPIReferenceValue.Text = "0.000";
            teB1RacStop_DRPIReferenceValue.Text = "0.000";
            teB1LStop_DRPIReferenceValue.Text = "0.000";
            teB1CStop_DRPIReferenceValue.Text = "0.000000";
            teB1QStop_DRPIReferenceValue.Text = "0.000";

            teB2RdcControlRod_DRPIReferenceValue.Text = "0.000";
            teB2RacControlRod_DRPIReferenceValue.Text = "0.000";
            teB2LControlRod_DRPIReferenceValue.Text = "0.000";
            teB2CControlRod_DRPIReferenceValue.Text = "0.000000";
            teB2QControlRod_DRPIReferenceValue.Text = "0.000";

            teB2RdcStop_DRPIReferenceValue.Text = "0.000";
            teB2RacStop_DRPIReferenceValue.Text = "0.000";
            teB2LStop_DRPIReferenceValue.Text = "0.000";
            teB2CStop_DRPIReferenceValue.Text = "0.000000";
            teB2QStop_DRPIReferenceValue.Text = "0.000";

            teB3RdcControlRod_DRPIReferenceValue.Text = "0.000";
            teB3RacControlRod_DRPIReferenceValue.Text = "0.000";
            teB3LControlRod_DRPIReferenceValue.Text = "0.000";
            teB3CControlRod_DRPIReferenceValue.Text = "0.000000";
            teB3QControlRod_DRPIReferenceValue.Text = "0.000";

            teB3RdcStop_DRPIReferenceValue.Text = "0.000";
            teB3RacStop_DRPIReferenceValue.Text = "0.000";
            teB3LStop_DRPIReferenceValue.Text = "0.000";
            teB3CStop_DRPIReferenceValue.Text = "0.000000";
            teB3QStop_DRPIReferenceValue.Text = "0.000";
        }

        /// <summary>
        /// 평균 값(DRPI) 초기화
        /// </summary>
        private void SetDRPIAverageValueInitialize()
        {
            teA1RdcControlRod_DRPIAverageValue.Text = "0.000";
            teA1RacControlRod_DRPIAverageValue.Text = "0.000";
            teA1LControlRod_DRPIAverageValue.Text = "0.000";
            teA1CControlRod_DRPIAverageValue.Text = "0.000000";
            teA1QControlRod_DRPIAverageValue.Text = "0.000";

            teA1RdcStop_DRPIAverageValue.Text = "0.000";
            teA1RacStop_DRPIAverageValue.Text = "0.000";
            teA1LStop_DRPIAverageValue.Text = "0.000";
            teA1CStop_DRPIAverageValue.Text = "0.000000";
            teA1QStop_DRPIAverageValue.Text = "0.000";

            teA2RdcControlRod_DRPIAverageValue.Text = "0.000";
            teA2RacControlRod_DRPIAverageValue.Text = "0.000";
            teA2LControlRod_DRPIAverageValue.Text = "0.000";
            teA2CControlRod_DRPIAverageValue.Text = "0.000000";
            teA2QControlRod_DRPIAverageValue.Text = "0.000";

            teA2RdcStop_DRPIAverageValue.Text = "0.000";
            teA2RacStop_DRPIAverageValue.Text = "0.000";
            teA2LStop_DRPIAverageValue.Text = "0.000";
            teA2CStop_DRPIAverageValue.Text = "0.000000";
            teA2QStop_DRPIAverageValue.Text = "0.000";

            teA3RdcControlRod_DRPIAverageValue.Text = "0.000";
            teA3RacControlRod_DRPIAverageValue.Text = "0.000";
            teA3LControlRod_DRPIAverageValue.Text = "0.000";
            teA3CControlRod_DRPIAverageValue.Text = "0.000000";
            teA3QControlRod_DRPIAverageValue.Text = "0.000";

            teA3RdcStop_DRPIAverageValue.Text = "0.000";
            teA3RacStop_DRPIAverageValue.Text = "0.000";
            teA3LStop_DRPIAverageValue.Text = "0.000";
            teA3CStop_DRPIAverageValue.Text = "0.000000";
            teA3QStop_DRPIAverageValue.Text = "0.000";

            teB1RdcControlRod_DRPIAverageValue.Text = "0.000";
            teB1RacControlRod_DRPIAverageValue.Text = "0.000";
            teB1LControlRod_DRPIAverageValue.Text = "0.000";
            teB1CControlRod_DRPIAverageValue.Text = "0.000000";
            teB1QControlRod_DRPIAverageValue.Text = "0.000";

            teB1RdcStop_DRPIAverageValue.Text = "0.000";
            teB1RacStop_DRPIAverageValue.Text = "0.000";
            teB1LStop_DRPIAverageValue.Text = "0.000";
            teB1CStop_DRPIAverageValue.Text = "0.000000";
            teB1QStop_DRPIAverageValue.Text = "0.000";

            teB2RdcControlRod_DRPIAverageValue.Text = "0.000";
            teB2RacControlRod_DRPIAverageValue.Text = "0.000";
            teB2LControlRod_DRPIAverageValue.Text = "0.000";
            teB2CControlRod_DRPIAverageValue.Text = "0.000000";
            teB2QControlRod_DRPIAverageValue.Text = "0.000";

            teB2RdcStop_DRPIAverageValue.Text = "0.000";
            teB2RacStop_DRPIAverageValue.Text = "0.000";
            teB2LStop_DRPIAverageValue.Text = "0.000";
            teB2CStop_DRPIAverageValue.Text = "0.000000";
            teB2QStop_DRPIAverageValue.Text = "0.000";

            teB3RdcControlRod_DRPIAverageValue.Text = "0.000";
            teB3RacControlRod_DRPIAverageValue.Text = "0.000";
            teB3LControlRod_DRPIAverageValue.Text = "0.000";
            teB3CControlRod_DRPIAverageValue.Text = "0.000000";
            teB3QControlRod_DRPIAverageValue.Text = "0.000";

            teB3RdcStop_DRPIAverageValue.Text = "0.000";
            teB3RacStop_DRPIAverageValue.Text = "0.000";
            teB3LStop_DRPIAverageValue.Text = "0.000";
            teB3CStop_DRPIAverageValue.Text = "0.000000";
            teB3QStop_DRPIAverageValue.Text = "0.000";
            
            dA_ControlRodAveRdc[0] = 0.000M;
            dA_ControlRodAveRac[0] = 0.000M;
            dA_ControlRodAveL[0] = 0.000M;
            dA_ControlRodAveC[0] = 0.000000M;
            dA_ControlRodAveQ[0] = 0.000M;

            dA_ControlRodAveRdc[1] = 0.000M;
            dA_ControlRodAveRac[1] = 0.000M;
            dA_ControlRodAveL[1] = 0.000M;
            dA_ControlRodAveC[1] = 0.000000M;
            dA_ControlRodAveQ[1] = 0.000M;

            dA_ControlRodAveRdc[2] = 0.000M;
            dA_ControlRodAveRac[2] = 0.000M;
            dA_ControlRodAveL[2] = 0.000M;
            dA_ControlRodAveC[2] = 0.000000M;
            dA_ControlRodAveQ[2] = 0.000M;

            dA_StopAveRdc[0] = 0.000M;
            dA_StopAveRac[0] = 0.000M;
            dA_StopAveL[0] = 0.000M;
            dA_StopAveC[0] = 0.000000M;
            dA_StopAveQ[0] = 0.000M;

            dA_StopAveRdc[1] = 0.000M;
            dA_StopAveRac[1] = 0.000M;
            dA_StopAveL[1] = 0.000M;
            dA_StopAveC[1] = 0.000000M;
            dA_StopAveQ[1] = 0.000M;

            dA_StopAveRdc[2] = 0.000M;
            dA_StopAveRac[2] = 0.000M;
            dA_StopAveL[2] = 0.000M;
            dA_StopAveC[2] = 0.000000M;
            dA_StopAveQ[2] = 0.000M;

            dB_ControlRodAveRdc[0] = 0.000M;
            dB_ControlRodAveRac[0] = 0.000M;
            dB_ControlRodAveL[0] = 0.000M;
            dB_ControlRodAveC[0] = 0.000000M;
            dB_ControlRodAveQ[0] = 0.000M;

            dB_ControlRodAveRdc[1] = 0.000M;
            dB_ControlRodAveRac[1] = 0.000M;
            dB_ControlRodAveL[1] = 0.000M;
            dB_ControlRodAveC[1] = 0.000000M;
            dB_ControlRodAveQ[1] = 0.000M;

            dB_ControlRodAveRdc[2] = 0.000M;
            dB_ControlRodAveRac[2] = 0.000M;
            dB_ControlRodAveL[2] = 0.000M;
            dB_ControlRodAveC[2] = 0.000000M;
            dB_ControlRodAveQ[2] = 0.000M;

            dB_StopAveRdc[0] = 0.000M;
            dB_StopAveRac[0] = 0.000M;
            dB_StopAveL[0] = 0.000M;
            dB_StopAveC[0] = 0.000000M;
            dB_StopAveQ[0] = 0.000M;

            dB_StopAveRdc[1] = 0.000M;
            dB_StopAveRac[1] = 0.000M;
            dB_StopAveL[1] = 0.000M;
            dB_StopAveC[1] = 0.000000M;
            dB_StopAveQ[1] = 0.000M;

            dB_StopAveRdc[2] = 0.000M;
            dB_StopAveRac[2] = 0.000M;
            dB_StopAveL[2] = 0.000M;
            dB_StopAveC[2] = 0.000000M;
            dB_StopAveQ[2] = 0.000M;
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
                            dA_ControlRodRdc[0] = dt.Rows[i]["Rdc_DRPIReferenceValue"] == null || dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim());
                            dA_ControlRodRac[0] = dt.Rows[i]["Rac_DRPIReferenceValue"] == null || dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim());
                            dA_ControlRodL[0] = dt.Rows[i]["L_DRPIReferenceValue"] == null || dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim());
                            dA_ControlRodC[0] = dt.Rows[i]["C_DRPIReferenceValue"] == null || dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000000M : Convert.ToDecimal(dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim());
                            dA_ControlRodQ[0] = dt.Rows[i]["Q_DRPIReferenceValue"] == null || dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim());
                        }
                        else if (dt.Rows[i]["CoilCabinetType"].ToString().Trim() == "제어용" && dt.Rows[i]["CoilCabinetName"].ToString().Trim() == "A2")
                        {
                            dA_ControlRodRdc[1] = dt.Rows[i]["Rdc_DRPIReferenceValue"] == null || dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim());
                            dA_ControlRodRac[1] = dt.Rows[i]["Rac_DRPIReferenceValue"] == null || dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim());
                            dA_ControlRodL[1] = dt.Rows[i]["L_DRPIReferenceValue"] == null || dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim());
                            dA_ControlRodC[1] = dt.Rows[i]["C_DRPIReferenceValue"] == null || dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000000M : Convert.ToDecimal(dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim());
                            dA_ControlRodQ[1] = dt.Rows[i]["Q_DRPIReferenceValue"] == null || dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim());
                        }
                        else if (dt.Rows[i]["CoilCabinetType"].ToString().Trim() == "제어용" && dt.Rows[i]["CoilCabinetName"].ToString().Substring(0, 1) == "A")
                        {
                            dA_ControlRodRdc[2] = dt.Rows[i]["Rdc_DRPIReferenceValue"] == null || dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim());
                            dA_ControlRodRac[2] = dt.Rows[i]["Rac_DRPIReferenceValue"] == null || dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim());
                            dA_ControlRodL[2] = dt.Rows[i]["L_DRPIReferenceValue"] == null || dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim());
                            dA_ControlRodC[2] = dt.Rows[i]["C_DRPIReferenceValue"] == null || dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000000M : Convert.ToDecimal(dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim());
                            dA_ControlRodQ[2] = dt.Rows[i]["Q_DRPIReferenceValue"] == null || dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim());
                        }
                        else if (dt.Rows[i]["CoilCabinetType"].ToString().Trim() == "제어용" && dt.Rows[i]["CoilCabinetName"].ToString().Trim() == "B1")
                        {
                            dB_ControlRodRdc[0] = dt.Rows[i]["Rdc_DRPIReferenceValue"] == null || dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim());
                            dB_ControlRodRac[0] = dt.Rows[i]["Rac_DRPIReferenceValue"] == null || dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim());
                            dB_ControlRodL[0] = dt.Rows[i]["L_DRPIReferenceValue"] == null || dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim());
                            dB_ControlRodC[0] = dt.Rows[i]["C_DRPIReferenceValue"] == null || dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000000M : Convert.ToDecimal(dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim());
                            dB_ControlRodQ[0] = dt.Rows[i]["Q_DRPIReferenceValue"] == null || dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim());
                        }
                        else if (dt.Rows[i]["CoilCabinetType"].ToString().Trim() == "제어용" && dt.Rows[i]["CoilCabinetName"].ToString().Trim() == "B2")
                        {
                            dB_ControlRodRdc[1] = dt.Rows[i]["Rdc_DRPIReferenceValue"] == null || dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim());
                            dB_ControlRodRac[1] = dt.Rows[i]["Rac_DRPIReferenceValue"] == null || dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim());
                            dB_ControlRodL[1] = dt.Rows[i]["L_DRPIReferenceValue"] == null || dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim());
                            dB_ControlRodC[1] = dt.Rows[i]["C_DRPIReferenceValue"] == null || dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000000M : Convert.ToDecimal(dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim());
                            dB_ControlRodQ[1] = dt.Rows[i]["Q_DRPIReferenceValue"] == null || dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim());
                        }
                        else if (dt.Rows[i]["CoilCabinetType"].ToString().Trim() == "제어용" && dt.Rows[i]["CoilCabinetName"].ToString().Substring(0, 1) == "B")
                        {
                            dB_ControlRodRdc[2] = dt.Rows[i]["Rdc_DRPIReferenceValue"] == null || dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim());
                            dB_ControlRodRac[2] = dt.Rows[i]["Rac_DRPIReferenceValue"] == null || dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim());
                            dB_ControlRodL[2] = dt.Rows[i]["L_DRPIReferenceValue"] == null || dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim());
                            dB_ControlRodC[2] = dt.Rows[i]["C_DRPIReferenceValue"] == null || dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000000M : Convert.ToDecimal(dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim());
                            dB_ControlRodQ[2] = dt.Rows[i]["Q_DRPIReferenceValue"] == null || dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim());
                        }
                        else if (dt.Rows[i]["CoilCabinetType"].ToString().Trim() == "정지용" & dt.Rows[i]["CoilCabinetName"].ToString().Trim() == "A1")
                        {
                            dA_StopRdc[0] = dt.Rows[i]["Rdc_DRPIReferenceValue"] == null || dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim());
                            dA_StopRac[0] = dt.Rows[i]["Rac_DRPIReferenceValue"] == null || dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim());
                            dA_StopL[0] = dt.Rows[i]["L_DRPIReferenceValue"] == null || dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim());
                            dA_StopC[0] = dt.Rows[i]["C_DRPIReferenceValue"] == null || dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000000M : Convert.ToDecimal(dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim());
                            dA_StopQ[0] = dt.Rows[i]["Q_DRPIReferenceValue"] == null || dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim());
                        }
                        else if (dt.Rows[i]["CoilCabinetType"].ToString().Trim() == "정지용" && dt.Rows[i]["CoilCabinetName"].ToString().Trim() == "A2")
                        {
                            dA_StopRdc[1] = dt.Rows[i]["Rdc_DRPIReferenceValue"] == null || dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim());
                            dA_StopRac[1] = dt.Rows[i]["Rac_DRPIReferenceValue"] == null || dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim());
                            dA_StopL[1] = dt.Rows[i]["L_DRPIReferenceValue"] == null || dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim());
                            dA_StopC[1] = dt.Rows[i]["C_DRPIReferenceValue"] == null || dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000000M : Convert.ToDecimal(dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim());
                            dA_StopQ[1] = dt.Rows[i]["Q_DRPIReferenceValue"] == null || dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim());
                        }
                        else if (dt.Rows[i]["CoilCabinetType"].ToString().Trim() == "정지용" && dt.Rows[i]["CoilCabinetName"].ToString().Substring(0, 1) == "A")
                        {
                            dA_StopRdc[2] = dt.Rows[i]["Rdc_DRPIReferenceValue"] == null || dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim());
                            dA_StopRac[2] = dt.Rows[i]["Rac_DRPIReferenceValue"] == null || dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim());
                            dA_StopL[2] = dt.Rows[i]["L_DRPIReferenceValue"] == null || dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim());
                            dA_StopC[2] = dt.Rows[i]["C_DRPIReferenceValue"] == null || dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000000M : Convert.ToDecimal(dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim());
                            dA_StopQ[2] = dt.Rows[i]["Q_DRPIReferenceValue"] == null || dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim());
                        }
                        else if (dt.Rows[i]["CoilCabinetType"].ToString().Trim() == "정지용" && dt.Rows[i]["CoilCabinetName"].ToString().Trim() == "B1")
                        {
                            dB_StopRdc[0] = dt.Rows[i]["Rdc_DRPIReferenceValue"] == null || dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim());
                            dB_StopRac[0] = dt.Rows[i]["Rac_DRPIReferenceValue"] == null || dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim());
                            dB_StopL[0] = dt.Rows[i]["L_DRPIReferenceValue"] == null || dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim());
                            dB_StopC[0] = dt.Rows[i]["C_DRPIReferenceValue"] == null || dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000000M : Convert.ToDecimal(dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim());
                            dB_StopQ[0] = dt.Rows[i]["Q_DRPIReferenceValue"] == null || dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim());
                        }
                        else if (dt.Rows[i]["CoilCabinetType"].ToString().Trim() == "정지용" && dt.Rows[i]["CoilCabinetName"].ToString().Trim() == "B2")
                        {
                            dB_StopRdc[1] = dt.Rows[i]["Rdc_DRPIReferenceValue"] == null || dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim());
                            dB_StopRac[1] = dt.Rows[i]["Rac_DRPIReferenceValue"] == null || dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim());
                            dB_StopL[1] = dt.Rows[i]["L_DRPIReferenceValue"] == null || dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim());
                            dB_StopC[1] = dt.Rows[i]["C_DRPIReferenceValue"] == null || dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000000M : Convert.ToDecimal(dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim());
                            dB_StopQ[1] = dt.Rows[i]["Q_DRPIReferenceValue"] == null || dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim());
                        }
                        else if (dt.Rows[i]["CoilCabinetType"].ToString().Trim() == "정지용" && dt.Rows[i]["CoilCabinetName"].ToString().Substring(0, 1) == "B")
                        {
                            dB_StopRdc[2] = dt.Rows[i]["Rdc_DRPIReferenceValue"] == null || dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim());
                            dB_StopRac[2] = dt.Rows[i]["Rac_DRPIReferenceValue"] == null || dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim());
                            dB_StopL[2] = dt.Rows[i]["L_DRPIReferenceValue"] == null || dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim());
                            dB_StopC[2] = dt.Rows[i]["C_DRPIReferenceValue"] == null || dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000000M : Convert.ToDecimal(dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim());
                            dB_StopQ[2] = dt.Rows[i]["Q_DRPIReferenceValue"] == null || dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim() == ""
                                    ? 0.000M : Convert.ToDecimal(dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim());
                        }
                    }

                    // 기준 값 Text Box 설정
                    SetReferenceValueTextBox();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
        }

        /// <summary>
        /// 기준 값 Text Box 설정
        /// </summary>
        private void SetReferenceValueTextBox()
        {
            teA1RdcControlRod_DRPIReferenceValue.Text = dA_ControlRodRdc[0].ToString("F3").Trim();
            teA1RacControlRod_DRPIReferenceValue.Text = dA_ControlRodRac[0].ToString("F3").Trim();
            teA1LControlRod_DRPIReferenceValue.Text = dA_ControlRodL[0].ToString("F3").Trim();
            teA1CControlRod_DRPIReferenceValue.Text = dA_ControlRodC[0].ToString("F6").Trim();
            teA1QControlRod_DRPIReferenceValue.Text = dA_ControlRodQ[0].ToString("F3").Trim();

            teA1RdcStop_DRPIReferenceValue.Text = dA_StopRdc[0].ToString("F3").Trim();
            teA1RacStop_DRPIReferenceValue.Text = dA_StopRac[0].ToString("F3").Trim();
            teA1LStop_DRPIReferenceValue.Text = dA_StopL[0].ToString("F3").Trim();
            teA1CStop_DRPIReferenceValue.Text = dA_StopC[0].ToString("F6").Trim();
            teA1QStop_DRPIReferenceValue.Text = dA_StopQ[0].ToString("F3").Trim();

            teA2RdcControlRod_DRPIReferenceValue.Text = dA_ControlRodRdc[1].ToString("F3").Trim();
            teA2RacControlRod_DRPIReferenceValue.Text = dA_ControlRodRac[1].ToString("F3").Trim();
            teA2LControlRod_DRPIReferenceValue.Text = dA_ControlRodL[1].ToString("F3").Trim();
            teA2CControlRod_DRPIReferenceValue.Text = dA_ControlRodC[1].ToString("F6").Trim();
            teA2QControlRod_DRPIReferenceValue.Text = dA_ControlRodQ[1].ToString("F3").Trim();

            teA2RdcStop_DRPIReferenceValue.Text = dA_StopRdc[1].ToString("F3").Trim();
            teA2RacStop_DRPIReferenceValue.Text = dA_StopRac[1].ToString("F3").Trim();
            teA2LStop_DRPIReferenceValue.Text = dA_StopL[1].ToString("F3").Trim();
            teA2CStop_DRPIReferenceValue.Text = dA_StopC[1].ToString("F6").Trim();
            teA2QStop_DRPIReferenceValue.Text = dA_StopQ[1].ToString("F3").Trim();

            teA3RdcControlRod_DRPIReferenceValue.Text = dA_ControlRodRdc[2].ToString("F3").Trim();
            teA3RacControlRod_DRPIReferenceValue.Text = dA_ControlRodRac[2].ToString("F3").Trim();
            teA3LControlRod_DRPIReferenceValue.Text = dA_ControlRodL[2].ToString("F3").Trim();
            teA3CControlRod_DRPIReferenceValue.Text = dA_ControlRodC[2].ToString("F6").Trim();
            teA3QControlRod_DRPIReferenceValue.Text = dA_ControlRodQ[2].ToString("F3").Trim();

            teA3RdcStop_DRPIReferenceValue.Text = dA_StopRdc[2].ToString("F3").Trim();
            teA3RacStop_DRPIReferenceValue.Text = dA_StopRac[2].ToString("F3").Trim();
            teA3LStop_DRPIReferenceValue.Text = dA_StopL[2].ToString("F3").Trim();
            teA3CStop_DRPIReferenceValue.Text = dA_StopC[2].ToString("F6").Trim();
            teA3QStop_DRPIReferenceValue.Text = dA_StopQ[2].ToString("F3").Trim();

            teB1RdcControlRod_DRPIReferenceValue.Text = dB_ControlRodRdc[0].ToString("F3").Trim();
            teB1RacControlRod_DRPIReferenceValue.Text = dB_ControlRodRac[0].ToString("F3").Trim();
            teB1LControlRod_DRPIReferenceValue.Text = dB_ControlRodL[0].ToString("F3").Trim();
            teB1CControlRod_DRPIReferenceValue.Text = dB_ControlRodC[0].ToString("F6").Trim();
            teB1QControlRod_DRPIReferenceValue.Text = dB_ControlRodQ[0].ToString("F3").Trim();

            teB1RdcStop_DRPIReferenceValue.Text = dB_StopRdc[0].ToString("F3").Trim();
            teB1RacStop_DRPIReferenceValue.Text = dB_StopRac[0].ToString("F3").Trim();
            teB1LStop_DRPIReferenceValue.Text = dB_StopL[0].ToString("F3").Trim();
            teB1CStop_DRPIReferenceValue.Text = dB_StopC[0].ToString("F6").Trim();
            teB1QStop_DRPIReferenceValue.Text = dB_StopQ[0].ToString("F3").Trim();

            teB2RdcControlRod_DRPIReferenceValue.Text = dB_ControlRodRdc[1].ToString("F3").Trim();
            teB2RacControlRod_DRPIReferenceValue.Text = dB_ControlRodRac[1].ToString("F3").Trim();
            teB2LControlRod_DRPIReferenceValue.Text = dB_ControlRodL[1].ToString("F3").Trim();
            teB2CControlRod_DRPIReferenceValue.Text = dB_ControlRodC[1].ToString("F6").Trim();
            teB2QControlRod_DRPIReferenceValue.Text = dB_ControlRodQ[1].ToString("F3").Trim();

            teB2RdcStop_DRPIReferenceValue.Text = dB_StopRdc[1].ToString("F3").Trim();
            teB2RacStop_DRPIReferenceValue.Text = dB_StopRac[1].ToString("F3").Trim();
            teB2LStop_DRPIReferenceValue.Text = dB_StopL[1].ToString("F3").Trim();
            teB2CStop_DRPIReferenceValue.Text = dB_StopC[1].ToString("F6").Trim();
            teB2QStop_DRPIReferenceValue.Text = dB_StopQ[1].ToString("F3").Trim();

            teB3RdcControlRod_DRPIReferenceValue.Text = dB_ControlRodRdc[2].ToString("F3").Trim();
            teB3RacControlRod_DRPIReferenceValue.Text = dB_ControlRodRac[2].ToString("F3").Trim();
            teB3LControlRod_DRPIReferenceValue.Text = dB_ControlRodL[2].ToString("F3").Trim();
            teB3CControlRod_DRPIReferenceValue.Text = dB_ControlRodC[2].ToString("F6").Trim();
            teB3QControlRod_DRPIReferenceValue.Text = dB_ControlRodQ[2].ToString("F3").Trim();

            teB3RdcStop_DRPIReferenceValue.Text = dB_StopRdc[2].ToString("F3").Trim();
            teB3RacStop_DRPIReferenceValue.Text = dB_StopRac[2].ToString("F3").Trim();
            teB3LStop_DRPIReferenceValue.Text = dB_StopL[2].ToString("F3").Trim();
            teB3CStop_DRPIReferenceValue.Text = dB_StopC[2].ToString("F6").Trim();
            teB3QStop_DRPIReferenceValue.Text = dB_StopQ[2].ToString("F3").Trim();
        }

        /// <summary>
        /// 평균 값 Text Box 설정
        /// </summary>
        private void SetAverageValueTextBox()
        {
            teA1RdcControlRod_DRPIAverageValue.Text = dA_ControlRodAveRdc[0].ToString("F3").Trim();
            teA1RacControlRod_DRPIAverageValue.Text = dA_ControlRodAveRac[0].ToString("F3").Trim();
            teA1LControlRod_DRPIAverageValue.Text = dA_ControlRodAveL[0].ToString("F3").Trim();
            teA1CControlRod_DRPIAverageValue.Text = dA_ControlRodAveC[0].ToString("F6").Trim();
            teA1QControlRod_DRPIAverageValue.Text = dA_ControlRodAveQ[0].ToString("F3").Trim();

            teA1RdcStop_DRPIAverageValue.Text = dA_StopAveRdc[0].ToString("F3").Trim();
            teA1RacStop_DRPIAverageValue.Text = dA_StopAveRac[0].ToString("F3").Trim();
            teA1LStop_DRPIAverageValue.Text = dA_StopAveL[0].ToString("F3").Trim();
            teA1CStop_DRPIAverageValue.Text = dA_StopAveC[0].ToString("F6").Trim();
            teA1QStop_DRPIAverageValue.Text = dA_StopAveQ[0].ToString("F3").Trim();

            teA2RdcControlRod_DRPIAverageValue.Text = dA_ControlRodAveRdc[1].ToString("F3").Trim();
            teA2RacControlRod_DRPIAverageValue.Text = dA_ControlRodAveRac[1].ToString("F3").Trim();
            teA2LControlRod_DRPIAverageValue.Text = dA_ControlRodAveL[1].ToString("F3").Trim();
            teA2CControlRod_DRPIAverageValue.Text = dA_ControlRodAveC[1].ToString("F6").Trim();
            teA2QControlRod_DRPIAverageValue.Text = dA_ControlRodAveQ[1].ToString("F3").Trim();

            teA2RdcStop_DRPIAverageValue.Text = dA_StopAveRdc[1].ToString("F3").Trim();
            teA2RacStop_DRPIAverageValue.Text = dA_StopAveRac[1].ToString("F3").Trim();
            teA2LStop_DRPIAverageValue.Text = dA_StopAveL[1].ToString("F3").Trim();
            teA2CStop_DRPIAverageValue.Text = dA_StopAveC[1].ToString("F6").Trim();
            teA2QStop_DRPIAverageValue.Text = dA_StopAveQ[1].ToString("F3").Trim();

            teA3RdcControlRod_DRPIAverageValue.Text = dA_ControlRodAveRdc[2].ToString("F3").Trim();
            teA3RacControlRod_DRPIAverageValue.Text = dA_ControlRodAveRac[2].ToString("F3").Trim();
            teA3LControlRod_DRPIAverageValue.Text = dA_ControlRodAveL[2].ToString("F3").Trim();
            teA3CControlRod_DRPIAverageValue.Text = dA_ControlRodAveC[2].ToString("F6").Trim();
            teA3QControlRod_DRPIAverageValue.Text = dA_ControlRodAveQ[2].ToString("F3").Trim();

            teA3RdcStop_DRPIAverageValue.Text = dA_StopAveRdc[2].ToString("F3").Trim();
            teA3RacStop_DRPIAverageValue.Text = dA_StopAveRac[2].ToString("F3").Trim();
            teA3LStop_DRPIAverageValue.Text = dA_StopAveL[2].ToString("F3").Trim();
            teA3CStop_DRPIAverageValue.Text = dA_StopAveC[2].ToString("F6").Trim();
            teA3QStop_DRPIAverageValue.Text = dA_StopAveQ[2].ToString("F3").Trim();

            teB1RdcControlRod_DRPIAverageValue.Text = dB_ControlRodAveRdc[0].ToString("F3").Trim();
            teB1RacControlRod_DRPIAverageValue.Text = dB_ControlRodAveRac[0].ToString("F3").Trim();
            teB1LControlRod_DRPIAverageValue.Text = dB_ControlRodAveL[0].ToString("F3").Trim();
            teB1CControlRod_DRPIAverageValue.Text = dB_ControlRodAveC[0].ToString("F6").Trim();
            teB1QControlRod_DRPIAverageValue.Text = dB_ControlRodAveQ[0].ToString("F3").Trim();

            teB1RdcStop_DRPIAverageValue.Text = dB_StopAveRdc[0].ToString("F3").Trim();
            teB1RacStop_DRPIAverageValue.Text = dB_StopAveRac[0].ToString("F3").Trim();
            teB1LStop_DRPIAverageValue.Text = dB_StopAveL[0].ToString("F3").Trim();
            teB1CStop_DRPIAverageValue.Text = dB_StopAveC[0].ToString("F6").Trim();
            teB1QStop_DRPIAverageValue.Text = dB_StopAveQ[0].ToString("F3").Trim();

            teB2RdcControlRod_DRPIAverageValue.Text = dB_ControlRodAveRdc[1].ToString("F3").Trim();
            teB2RacControlRod_DRPIAverageValue.Text = dB_ControlRodAveRac[1].ToString("F3").Trim();
            teB2LControlRod_DRPIAverageValue.Text = dB_ControlRodAveL[1].ToString("F3").Trim();
            teB2CControlRod_DRPIAverageValue.Text = dB_ControlRodAveC[1].ToString("F6").Trim();
            teB2QControlRod_DRPIAverageValue.Text = dB_ControlRodAveQ[1].ToString("F3").Trim();

            teB2RdcStop_DRPIAverageValue.Text = dB_StopAveRdc[1].ToString("F3").Trim();
            teB2RacStop_DRPIAverageValue.Text = dB_StopAveRac[1].ToString("F3").Trim();
            teB2LStop_DRPIAverageValue.Text = dB_StopAveL[1].ToString("F3").Trim();
            teB2CStop_DRPIAverageValue.Text = dB_StopAveC[1].ToString("F6").Trim();
            teB2QStop_DRPIAverageValue.Text = dB_StopAveQ[1].ToString("F3").Trim();

            teB3RdcControlRod_DRPIAverageValue.Text = dB_ControlRodAveRdc[2].ToString("F3").Trim();
            teB3RacControlRod_DRPIAverageValue.Text = dB_ControlRodAveRac[2].ToString("F3").Trim();
            teB3LControlRod_DRPIAverageValue.Text = dB_ControlRodAveL[2].ToString("F3").Trim();
            teB3CControlRod_DRPIAverageValue.Text = dB_ControlRodAveC[2].ToString("F6").Trim();
            teB3QControlRod_DRPIAverageValue.Text = dB_ControlRodAveQ[2].ToString("F3").Trim();

            teB3RdcStop_DRPIAverageValue.Text = dB_StopAveRdc[2].ToString("F3").Trim();
            teB3RacStop_DRPIAverageValue.Text = dB_StopAveRac[2].ToString("F3").Trim();
            teB3LStop_DRPIAverageValue.Text = dB_StopAveL[2].ToString("F3").Trim();
            teB3CStop_DRPIAverageValue.Text = dB_StopAveC[2].ToString("F6").Trim();
            teB3QStop_DRPIAverageValue.Text = dB_StopAveQ[2].ToString("F3").Trim();
        }

        private void SetChart_Depict(System.Windows.Forms.DataVisualization.Charting.Chart cht, string chartName
            , double[] dA1Value, double[] dA2Value, double[] dA3Value, double[] dA4Value, double[] dA5Value, double[] dA6Value, double[] dA7Value
            , double[] dA8Value, double[] dA9Value, double[] dA10Value, double[] dA11Value, double[] dA12Value, double[] dA13Value, double[] dA14Value
            , double[] dA15Value, double[] dA16Value, double[] dA17Value, double[] dA18Value, double[] dA19Value, double[] dA20Value, double[] dA21Value
            , double[] dRefControlRodValue, double[] dRefStopValue, double[] dAveControlRodValue, double[] dAveStopValue
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

            m_Chart.MakeLegend(ref cht, "제어용 기준값", "right", chartAreaNo);
            m_Chart.AddChartColumn(ref cht, chartAreaNo, chartAreaNo, "제어용 기준값", ref arrayXLabelValue, ref dRefControlRodValue, true, System.Drawing.Color.Lime, 1, false);

            m_Chart.MakeLegend(ref cht, "정지용 기준값", "right", chartAreaNo);
            m_Chart.AddChartColumn(ref cht, chartAreaNo, chartAreaNo, "정지용 기준값", ref arrayXLabelValue, ref dRefStopValue, true, System.Drawing.Color.Blue, 1, false);

            m_Chart.MakeLegend(ref cht, "제어용 평균값", "right", chartAreaNo);
            m_Chart.AddChartColumn(ref cht, chartAreaNo, chartAreaNo, "제어용 평균값", ref arrayXLabelValue, ref dAveControlRodValue, true, System.Drawing.Color.Lime, 2, false);

            m_Chart.MakeLegend(ref cht, "정지용 평균값", "right", chartAreaNo);
            m_Chart.AddChartColumn(ref cht, chartAreaNo, chartAreaNo, "정지용 평균값", ref arrayXLabelValue, ref dAveStopValue, true, System.Drawing.Color.Blue, 2, false);

            if (arrayXLabelValue.Length < 30)
                cht.ChartAreas[0].AxisX.Interval = 1; // 차수별일 경우는 X 축 1씩 증가 값 설정
            else
                cht.ChartAreas[0].AxisX.Interval = 10; // 제어봉일 경우는 X 축 10씩 증가 값 설정
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

                        }
                        else
                        {
                            dgvMeasurement.Columns["DC_ResistanceValue"].Visible = false;
                            dgvMeasurement.Columns["DC_Deviation"].Visible = false;

                        }
                        break;
                    case "chkRac":
                        if (((CheckBox)sender).Checked)
                        {
                            dgvMeasurement.Columns["AC_ResistanceValue"].Visible = true;
                            dgvMeasurement.Columns["AC_Deviation"].Visible = true;
                        }
                        else
                        {
                            dgvMeasurement.Columns["AC_ResistanceValue"].Visible = false;
                            dgvMeasurement.Columns["AC_Deviation"].Visible = false;
                        }
                        break;
                    case "chkL":
                        if (((CheckBox)sender).Checked)
                        {
                            dgvMeasurement.Columns["L_InductanceValue"].Visible = true;
                            dgvMeasurement.Columns["L_Deviation"].Visible = true;
                        }
                        else
                        {
                            dgvMeasurement.Columns["L_InductanceValue"].Visible = false;
                            dgvMeasurement.Columns["L_Deviation"].Visible = false;
                        }
                        break;
                    case "chkC":
                        if (((CheckBox)sender).Checked)
                        {
                            dgvMeasurement.Columns["C_CapacitanceValue"].Visible = true;
                            dgvMeasurement.Columns["C_Deviation"].Visible = true;
                        }
                        else
                        {
                            dgvMeasurement.Columns["C_CapacitanceValue"].Visible = false;
                            dgvMeasurement.Columns["C_Deviation"].Visible = false;
                        }
                        break;
                    case "chkQ":
                        if (((CheckBox)sender).Checked)
                        {
                            dgvMeasurement.Columns["Q_FactorValue"].Visible = true;
                            dgvMeasurement.Columns[12].Visible = true;
                        }
                        else
                        {
                            dgvMeasurement.Columns["Q_FactorValue"].Visible = false;
                            dgvMeasurement.Columns["Q_Deviation"].Visible = false;
                        }
                        break;
                }

                // Rdc, Rac 활성화 및 비활성화
                SetTabControlDCAC();

                // Rdc, Rac 활성화 및 비활성화
                SetTabControlLCQ();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
        }

        /// <summary>
        /// L, C, Q 활성화 및 비활성화
        /// </summary>
        private void SetTabControlDCAC()
        {
            if (chkRdc.Checked && chkRac.Checked)
            {
                chartMeasurementValueDC.Visible = true;
                chartAverageValueDC.Visible = true;

                chartMeasurementValueAC.Visible = true;
                chartAverageValueAC.Visible = true;

                tlpRefDCAC.RowCount = 2;
                tlpRefDCAC.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
                tlpRefDCAC.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));

                tlpRefDCAC.SetRowSpan(panel7, 2);

                tlpRefDCAC.Controls.Add(this.chartMeasurementValueAC, 0, 1);
                tlpRefDCAC.SetRowSpan(chartMeasurementValueAC, 1);
                tlpRefDCAC.Controls.Add(this.chartMeasurementValueDC, 0, 0);
                tlpRefDCAC.SetRowSpan(chartMeasurementValueDC, 1);

                tlpAveDCAC.RowCount = 2;
                tlpAveDCAC.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
                tlpAveDCAC.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));

                tlpAveDCAC.SetRowSpan(panel6, 2);

                tlpAveDCAC.Controls.Add(this.chartAverageValueAC, 0, 1);
                tlpRefDCAC.SetRowSpan(chartAverageValueAC, 1);
                tlpAveDCAC.Controls.Add(this.chartAverageValueDC, 0, 0);
                tlpAveDCAC.SetRowSpan(chartAverageValueDC, 1);
            }
            else if (chkRdc.Checked && !chkRac.Checked)
            {
                chartMeasurementValueDC.Visible = true;
                chartAverageValueDC.Visible = true;

                chartMeasurementValueAC.Visible = false;
                chartAverageValueAC.Visible = false;

                tlpRefDCAC.RowCount = 1;
                tlpRefDCAC.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));

                tlpRefDCAC.SetRowSpan(panel7, 1);

                tlpRefDCAC.Controls.Add(this.chartMeasurementValueDC, 0, 0);
                tlpRefDCAC.SetRowSpan(chartMeasurementValueDC, 1);

                tlpAveDCAC.RowCount = 1;
                tlpAveDCAC.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));

                tlpAveDCAC.SetRowSpan(panel6, 1);

                tlpAveDCAC.Controls.Add(this.chartAverageValueDC, 0, 0);
                tlpAveDCAC.SetRowSpan(chartAverageValueDC, 1);
            }
            else if (!chkRdc.Checked && chkRac.Checked)
            {
                chartMeasurementValueDC.Visible = false;
                chartAverageValueDC.Visible = false;

                chartMeasurementValueAC.Visible = true;
                chartAverageValueAC.Visible = true;

                tlpRefDCAC.RowCount = 1;
                tlpRefDCAC.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));

                tlpRefDCAC.SetRowSpan(panel7, 1);

                tlpRefDCAC.Controls.Add(this.chartMeasurementValueAC, 0, 0);
                tlpRefDCAC.SetRowSpan(chartMeasurementValueAC, 1);

                tlpAveDCAC.RowCount = 1;
                tlpAveDCAC.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));

                tlpAveDCAC.SetRowSpan(panel6, 1);

                tlpAveDCAC.Controls.Add(this.chartAverageValueAC, 0, 0);
                tlpAveDCAC.SetRowSpan(chartAverageValueAC, 1);
            }
            else if (!chkRdc.Checked && !chkRac.Checked)
            {
                chartMeasurementValueDC.Visible = false;
                chartAverageValueDC.Visible = false;

                chartMeasurementValueAC.Visible = false;
                chartAverageValueAC.Visible = false;

            }
        }

        /// <summary>
        /// L, C, Q 활성화 및 비활성화
        /// </summary>
        private void SetTabControlLCQ()
        {
            if (chkL.Checked && chkC.Checked && chkQ.Checked)
            {
                chartMeasurementValueL.Visible = true;
                chartAverageValueL.Visible = true;

                chartMeasurementValueC.Visible = true;
                chartAverageValueC.Visible = true;

                chartMeasurementValueQ.Visible = true;
                chartAverageValueQ.Visible = true;

                tlpRefLCQ.RowCount = 3;
                tlpRefLCQ.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.33333F));
                tlpRefLCQ.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.33333F));
                tlpRefLCQ.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.33333F));

                tlpRefLCQ.SetRowSpan(panel5, 3);

                tlpRefLCQ.Controls.Add(this.chartMeasurementValueQ, 0, 2);
                tlpRefLCQ.SetRowSpan(chartMeasurementValueQ, 1);
                tlpRefLCQ.Controls.Add(this.chartMeasurementValueC, 0, 1);
                tlpRefLCQ.SetRowSpan(chartMeasurementValueC, 1);
                tlpRefLCQ.Controls.Add(this.chartMeasurementValueL, 0, 0);
                tlpRefLCQ.SetRowSpan(chartMeasurementValueL, 1);

                tlpAveLCQ.RowCount = 3;
                tlpAveLCQ.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.33333F));
                tlpAveLCQ.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.33333F));
                tlpAveLCQ.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.33333F));

                tlpRefLCQ.SetRowSpan(panel4, 3);

                tlpAveLCQ.Controls.Add(this.chartAverageValueQ, 0, 2);
                tlpAveLCQ.SetRowSpan(chartAverageValueQ, 1);
                tlpAveLCQ.Controls.Add(this.chartAverageValueC, 0, 1);
                tlpAveLCQ.SetRowSpan(chartAverageValueC, 1);
                tlpAveLCQ.Controls.Add(this.chartAverageValueL, 0, 0);
                tlpAveLCQ.SetRowSpan(chartAverageValueL, 1);
            }
            else if (chkL.Checked && chkC.Checked && !chkQ.Checked)
            {
                chartMeasurementValueL.Visible = true;
                chartAverageValueL.Visible = true;

                chartMeasurementValueC.Visible = true;
                chartAverageValueC.Visible = true;

                chartMeasurementValueQ.Visible = false;
                chartAverageValueQ.Visible = false;

                tlpRefLCQ.RowCount = 2;
                tlpRefLCQ.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
                tlpRefLCQ.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));

                tlpRefLCQ.SetRowSpan(panel5, 2);

                tlpRefLCQ.Controls.Add(this.chartMeasurementValueC, 0, 1);
                tlpRefLCQ.SetRowSpan(chartMeasurementValueC, 1);
                tlpRefLCQ.Controls.Add(this.chartMeasurementValueL, 0, 0);
                tlpRefLCQ.SetRowSpan(chartMeasurementValueL, 1);

                tlpAveLCQ.RowCount = 2;
                tlpAveLCQ.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
                tlpAveLCQ.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));

                tlpRefLCQ.SetRowSpan(panel4, 2);

                tlpAveLCQ.Controls.Add(this.chartAverageValueC, 0, 1);
                tlpAveLCQ.SetRowSpan(chartAverageValueC, 1);
                tlpAveLCQ.Controls.Add(this.chartAverageValueL, 0, 0);
                tlpAveLCQ.SetRowSpan(chartAverageValueL, 1);
            }
            else if (chkL.Checked && !chkC.Checked && chkQ.Checked)
            {
                chartMeasurementValueL.Visible = true;
                chartAverageValueL.Visible = true;

                chartMeasurementValueC.Visible = false;
                chartAverageValueC.Visible = false;

                chartMeasurementValueQ.Visible = true;
                chartAverageValueQ.Visible = true;

                tlpRefLCQ.RowCount = 2;
                tlpRefLCQ.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
                tlpRefLCQ.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));

                tlpRefLCQ.SetRowSpan(panel5, 2);

                tlpRefLCQ.Controls.Add(this.chartMeasurementValueQ, 0, 1);
                tlpRefLCQ.SetRowSpan(chartMeasurementValueQ, 1);
                tlpRefLCQ.Controls.Add(this.chartMeasurementValueL, 0, 0);
                tlpRefLCQ.SetRowSpan(chartMeasurementValueL, 1);

                tlpAveLCQ.RowCount = 2;
                tlpAveLCQ.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
                tlpAveLCQ.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));

                tlpRefLCQ.SetRowSpan(panel4, 2);

                tlpAveLCQ.Controls.Add(this.chartAverageValueQ, 0, 1);
                tlpAveLCQ.SetRowSpan(chartAverageValueQ, 1);
                tlpAveLCQ.Controls.Add(this.chartAverageValueL, 0, 0);
                tlpAveLCQ.SetRowSpan(chartAverageValueL, 1);
            }
            else if (!chkL.Checked && chkC.Checked && chkQ.Checked)
            {
                chartMeasurementValueL.Visible = false;
                chartAverageValueL.Visible = false;

                chartMeasurementValueC.Visible = true;
                chartAverageValueC.Visible = true;

                chartMeasurementValueQ.Visible = true;
                chartAverageValueQ.Visible = true;

                tlpRefLCQ.RowCount = 2;
                tlpRefLCQ.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
                tlpRefLCQ.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));

                tlpRefLCQ.SetRowSpan(panel5, 2);

                tlpRefLCQ.Controls.Add(this.chartMeasurementValueQ, 0, 1);
                tlpRefLCQ.SetRowSpan(chartMeasurementValueQ, 1);
                tlpRefLCQ.Controls.Add(this.chartMeasurementValueC, 0, 0);
                tlpRefLCQ.SetRowSpan(chartMeasurementValueC, 1);

                tlpAveLCQ.RowCount = 2;
                tlpAveLCQ.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
                tlpAveLCQ.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));

                tlpRefLCQ.SetRowSpan(panel4, 2);

                tlpAveLCQ.Controls.Add(this.chartAverageValueQ, 0, 1);
                tlpAveLCQ.SetRowSpan(chartAverageValueQ, 1);
                tlpAveLCQ.Controls.Add(this.chartAverageValueC, 0, 0);
                tlpAveLCQ.SetRowSpan(chartAverageValueC, 1);
            }
            else if (chkL.Checked && !chkC.Checked && !chkQ.Checked)
            {
                chartMeasurementValueL.Visible = true;
                chartAverageValueL.Visible = true;

                chartMeasurementValueC.Visible = false;
                chartAverageValueC.Visible = false;

                chartMeasurementValueQ.Visible = false;
                chartAverageValueQ.Visible = false;

                tlpRefLCQ.RowCount = 1;
                tlpRefLCQ.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));

                tlpRefLCQ.SetRowSpan(panel5, 1);

                tlpRefLCQ.Controls.Add(this.chartMeasurementValueL, 0, 0);
                tlpRefLCQ.SetRowSpan(chartMeasurementValueL, 1);

                tlpAveLCQ.RowCount = 1;
                tlpAveLCQ.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));

                tlpRefLCQ.SetRowSpan(panel4, 1);

                tlpAveLCQ.Controls.Add(this.chartAverageValueL, 0, 0);
                tlpAveLCQ.SetRowSpan(chartAverageValueL, 1);
            }
            else if (!chkL.Checked && chkC.Checked && !chkQ.Checked)
            {
                chartMeasurementValueL.Visible = false;
                chartAverageValueL.Visible = false;

                chartMeasurementValueC.Visible = true;
                chartAverageValueC.Visible = true;

                chartMeasurementValueQ.Visible = false;
                chartAverageValueQ.Visible = false;

                tlpRefLCQ.RowCount = 1;
                tlpRefLCQ.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));

                tlpRefLCQ.SetRowSpan(panel5, 1);

                tlpRefLCQ.Controls.Add(this.chartMeasurementValueC, 0, 0);
                tlpRefLCQ.SetRowSpan(chartMeasurementValueC, 1);

                tlpAveLCQ.RowCount = 1;
                tlpAveLCQ.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));

                tlpRefLCQ.SetRowSpan(panel4, 1);

                tlpAveLCQ.Controls.Add(this.chartAverageValueC, 0, 0);
                tlpAveLCQ.SetRowSpan(chartAverageValueC, 1);
            }
            else if (!chkL.Checked && !chkC.Checked && chkQ.Checked)
            {
                chartMeasurementValueL.Visible = false;
                chartAverageValueL.Visible = false;

                chartMeasurementValueC.Visible = false;
                chartAverageValueC.Visible = false;

                chartMeasurementValueQ.Visible = true;
                chartAverageValueQ.Visible = true;

                tlpRefLCQ.RowCount = 1;
                tlpRefLCQ.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));

                tlpRefLCQ.SetRowSpan(panel5, 1);

                tlpRefLCQ.Controls.Add(this.chartMeasurementValueQ, 0, 0);
                tlpRefLCQ.SetRowSpan(chartMeasurementValueQ, 1);

                tlpAveLCQ.RowCount = 1;
                tlpAveLCQ.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));

                tlpRefLCQ.SetRowSpan(panel4, 1);

                tlpAveLCQ.Controls.Add(this.chartAverageValueQ, 0, 0);
                tlpAveLCQ.SetRowSpan(chartAverageValueQ, 1);
            }
            else if (!chkL.Checked && !chkC.Checked && !chkQ.Checked)
            {
                chartMeasurementValueL.Visible = false;
                chartAverageValueL.Visible = false;

                chartMeasurementValueC.Visible = false;
                chartAverageValueC.Visible = false;

                chartMeasurementValueQ.Visible = false;
                chartAverageValueQ.Visible = false;
            }
        }
    }
}
