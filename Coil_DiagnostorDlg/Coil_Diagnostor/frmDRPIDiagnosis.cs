using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using Coil_Diagnostor.Function;

namespace Coil_Diagnostor
{
    public partial class frmDRPIDiagnosis : Form
    {
        protected Function.FunctionMeasureProcess m_MeasureProcess = new Function.FunctionMeasureProcess();
        protected Function.FunctionDataControl m_db = new Function.FunctionDataControl();
        protected frmMessageBox frmMB = new frmMessageBox();
        protected string strPlantName = "";
        protected bool boolMessageOK = false;
        protected bool boolFormLoad = false;
        protected bool boollInitialize = false;
        public bool boolMeasurementStart = false;
        public bool boolMeasurementStop = false;
        public bool boolNormalMode = false;
        public bool boolDAQ = false;
        public bool boolLCRMeter = false;
        public static Thread threadMeasurementStart;
        public static Thread threadMeasurementStop;

        public decimal[] dA_ControlRodRdc = new decimal[3];    // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)   //?
        public decimal[] dA_ControlRodRac = new decimal[3];    // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        public decimal[] dA_ControlRodL = new decimal[3];      // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        public decimal[] dA_ControlRodC = new decimal[3];      // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        public decimal[] dA_ControlRodQ = new decimal[3];      // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        public decimal[] dB_ControlRodRdc = new decimal[3];    // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        public decimal[] dB_ControlRodRac = new decimal[3];    // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        public decimal[] dB_ControlRodL = new decimal[3];      // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        public decimal[] dB_ControlRodC = new decimal[3];      // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        public decimal[] dB_ControlRodQ = new decimal[3];      // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        public decimal[] dA_StopRdc = new decimal[3];          // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        public decimal[] dA_StopRac = new decimal[3];          // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        public decimal[] dA_StopL = new decimal[3];            // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        public decimal[] dA_StopC = new decimal[3];            // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        public decimal[] dA_StopQ = new decimal[3];            // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        public decimal[] dB_StopRdc = new decimal[3];          // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        public decimal[] dB_StopRac = new decimal[3];          // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        public decimal[] dB_StopL = new decimal[3];            // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        public decimal[] dB_StopC = new decimal[3];            // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)
        public decimal[] dB_StopQ = new decimal[3];            // 기준값 (Index 0 : A1, Index 1 : A2, Index 2 : A3 ~ 21)

        public string strMeasurementDate = "";
        public string strMeasurementResult = "";
        public string strDAQDeviceName = "";

        public decimal dcmFrequency = 120M;
        public decimal dcmVoltageLevel = 1000M;
        public decimal dcmTemperature_ReferenceValue = 0M;
        public decimal dcmTemperature_ChangeValue = 0M;
        public decimal dcmTemperature_Measurement = 0M;

        public string[] arrayCard01_NormalModePinMap = new string[21];
        public string[] arrayCard02_NormalModePinMap = new string[21];
        public string[] arrayCard03_NormalModePinMap = new string[21];
        public string[] arrayCard04_NormalModePinMap = new string[21];
        public string[] arrayCard05_NormalModePinMap = new string[21];
        public string[] arrayCard06_NormalModePinMap = new string[21];
        public string[] arrayCard07_NormalModePinMap = new string[21];
        public string[] arrayCard08_NormalModePinMap = new string[21];
        public string[] arrayCard09_NormalModePinMap = new string[21];
        public string[] arrayCard10_NormalModePinMap = new string[21];

        public string[] arrayCard01_WheatstoneModePinMap = new string[21];
        public string[] arrayCard02_WheatstoneModePinMap = new string[21];
        public string[] arrayCard03_WheatstoneModePinMap = new string[21];
        public string[] arrayCard04_WheatstoneModePinMap = new string[21];
        public string[] arrayCard05_WheatstoneModePinMap = new string[21];
        public string[] arrayCard06_WheatstoneModePinMap = new string[21];
        public string[] arrayCard07_WheatstoneModePinMap = new string[21];
        public string[] arrayCard08_WheatstoneModePinMap = new string[21];
        public string[] arrayCard09_WheatstoneModePinMap = new string[21];
        public string[] arrayCard10_WheatstoneModePinMap = new string[21];

        protected CheckBox allCheck = new CheckBox();
        protected bool isCheck = true;

        private Panel DRPIrodPanel;
        private GroupBox gbStop;
        private GroupBox gbControl;
        private GroupBox gbCardFrame;
        private RadioButton rboCoilA;
        private RadioButton rboCoilB;

        public frmDRPIDiagnosis()
        {
            CheckForIllegalCrossThreadCalls = false;
            InitializeComponent();

            // 카드 핀맵 초기화
            CardPinMapArrayInitialize();
        }

        /// <summary>
        /// 카드 핀맵 초기화
        /// </summary>
        private void CardPinMapArrayInitialize()
        {
            for (int i = 0; i < arrayCard01_NormalModePinMap.Length; i++)
            {
                // 일반모드 Relay Card별 핀맵 설정
                arrayCard01_NormalModePinMap[i] = GetNormalModeDAMPinMap(Gini.GetValue("DRPI", "DRPINormalMode_SelectDAMRelay1"), i);
                arrayCard02_NormalModePinMap[i] = GetNormalModeDAMPinMap(Gini.GetValue("DRPI", "DRPINormalMode_SelectDAMRelay2"), i);
                arrayCard03_NormalModePinMap[i] = GetNormalModeDAMPinMap(Gini.GetValue("DRPI", "DRPINormalMode_SelectDAMRelay3"), i);
                arrayCard04_NormalModePinMap[i] = GetNormalModeDAMPinMap(Gini.GetValue("DRPI", "DRPINormalMode_SelectDAMRelay4"), i);
                arrayCard05_NormalModePinMap[i] = GetNormalModeDAMPinMap(Gini.GetValue("DRPI", "DRPINormalMode_SelectDAMRelay5"), i);
                arrayCard06_NormalModePinMap[i] = GetNormalModeDAMPinMap(Gini.GetValue("DRPI", "DRPINormalMode_SelectDAMRelay6"), i);
                arrayCard07_NormalModePinMap[i] = GetNormalModeDAMPinMap(Gini.GetValue("DRPI", "DRPINormalMode_SelectDAMRelay7"), i);
                arrayCard08_NormalModePinMap[i] = GetNormalModeDAMPinMap(Gini.GetValue("DRPI", "DRPINormalMode_SelectDAMRelay8"), i);
                arrayCard09_NormalModePinMap[i] = GetNormalModeDAMPinMap(Gini.GetValue("DRPI", "DRPINormalMode_SelectDAMRelay9"), i);
                arrayCard10_NormalModePinMap[i] = GetNormalModeDAMPinMap(Gini.GetValue("DRPI", "DRPINormalMode_SelectDAMRelay10"), i);

                // 휘스톤모드 Relay Card별 핀맵 설정
                arrayCard01_WheatstoneModePinMap[i] = GetWheatstoneModeDAMPinMap(Gini.GetValue("DRPI", "DRPIWheatstoneMode_SelectDAMRelay1"), i);
                arrayCard02_WheatstoneModePinMap[i] = GetWheatstoneModeDAMPinMap(Gini.GetValue("DRPI", "DRPIWheatstoneMode_SelectDAMRelay2"), i);
                arrayCard03_WheatstoneModePinMap[i] = GetWheatstoneModeDAMPinMap(Gini.GetValue("DRPI", "DRPIWheatstoneMode_SelectDAMRelay3"), i);
                arrayCard04_WheatstoneModePinMap[i] = GetWheatstoneModeDAMPinMap(Gini.GetValue("DRPI", "DRPIWheatstoneMode_SelectDAMRelay4"), i);
                arrayCard05_WheatstoneModePinMap[i] = GetWheatstoneModeDAMPinMap(Gini.GetValue("DRPI", "DRPIWheatstoneMode_SelectDAMRelay5"), i);
                arrayCard06_WheatstoneModePinMap[i] = GetWheatstoneModeDAMPinMap(Gini.GetValue("DRPI", "DRPIWheatstoneMode_SelectDAMRelay6"), i);
                arrayCard07_WheatstoneModePinMap[i] = GetWheatstoneModeDAMPinMap(Gini.GetValue("DRPI", "DRPIWheatstoneMode_SelectDAMRelay7"), i);
                arrayCard08_WheatstoneModePinMap[i] = GetWheatstoneModeDAMPinMap(Gini.GetValue("DRPI", "DRPIWheatstoneMode_SelectDAMRelay8"), i);
                arrayCard09_WheatstoneModePinMap[i] = GetWheatstoneModeDAMPinMap(Gini.GetValue("DRPI", "DRPIWheatstoneMode_SelectDAMRelay9"), i);
                arrayCard10_WheatstoneModePinMap[i] = GetWheatstoneModeDAMPinMap(Gini.GetValue("DRPI", "DRPIWheatstoneMode_SelectDAMRelay10"), i);
            }
        }

        /// <summary>
        /// 일반모드 Relay Card별 핀맵 설정
        /// </summary>
        /// <param name="_strSelectDAMRelay"></param>
        /// <param name="intSelectIndex"></param>
        /// <returns></returns>
        private string GetNormalModeDAMPinMap(string _strSelectDAMRelay, int intSelectIndex)
        {
            string strResult = "";

            //_strSelectDAMRelay
            strResult = string.Format(Gini.GetValue("DAM_PinName", $"{_strSelectDAMRelay.Replace(" ", "")}_RelayCardNumberDAQPinName".Trim()), Gini.GetValue("Device", "DAQDeviceName").Trim(), strResult.Trim());
            //intSelectIndex
            strResult = string.Format(Gini.GetValue("Channel_PinName", $"NormalModeCh{intSelectIndex + 1}_DAQPinName").Trim(), Gini.GetValue("Device", "DAQDeviceName").Trim(), strResult.Trim());
            return strResult.Trim();
        }

        /// <summary>
        /// 휘스톤모드 Relay Card별 핀맵 설정
        /// </summary>
        /// <param name="_strSelectDAMRelay"></param>
        /// <param name="intSelectIndex"></param>
        /// <returns></returns>
        private string GetWheatstoneModeDAMPinMap(string _strSelectDAMRelay, int intSelectIndex)
        {
            string strResult = "";

            //_strSelectDAMRelay
            strResult = string.Format(Gini.GetValue("DAM_PinName", $"{_strSelectDAMRelay.Replace(" ", "")}_RelayCardNumberDAQPinName".Trim()), Gini.GetValue("Device", "DAQDeviceName").Trim(), strResult.Trim());
            //intSelectIndex
            strResult = string.Format(Gini.GetValue("Channel_PinName", $"WheatstoneModeCh{intSelectIndex + 1}_DAQPinName").Trim(), Gini.GetValue("Device", "DAQDeviceName").Trim(), strResult.Trim());
            return strResult.Trim();
        }

        /// <summary>
        /// 폼 단축키 지정
        /// </summary>
        protected override bool ProcessCmdKey(ref System.Windows.Forms.Message msg, Keys keyData)
        {
            Keys key = keyData & ~(Keys.Shift | Keys.Control);

            switch (key)
            {
                case Keys.F2: // 측정 버튼
                    btnMeasurementStart.PerformClick();
                    break;
                case Keys.F3: // 정지 버튼
                    btnMeasurementStop.PerformClick();
                    break;
                case Keys.F4: // 저장 버튼
                    btnSave.PerformClick();
                    break;
                case Keys.F5: // 카드 선택 버튼
                    btnSelectCard.PerformClick();
                    break;
                case Keys.F8: // PinMap Column 활성화/비활성화 설정
                    if (dgvMeasurement.Columns["DAQPinMap"].Visible)
                    {
                        dgvMeasurement.Columns["DAMName"].Visible = false;
                        dgvMeasurement.Columns["DAQPinMap"].Visible = false;
                        dgvMeasurement.Columns["Channel"].Visible = false;
                    }
                    else
                    {
                        dgvMeasurement.Columns["DAMName"].Visible = true;
                        dgvMeasurement.Columns["DAQPinMap"].Visible = true;
                        dgvMeasurement.Columns["Channel"].Visible = true;
                    }
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
        private void frmDRPIDiagnosis_Load(object sender, EventArgs e)
        {
            //boolFormLoad = true;

            strPlantName = Gini.GetValue("Device", "PlantName").Trim();
            string plantName = Gini.GetValue("Device", "PlantName");
            int rodCount = SetRodPanel(plantName);
            // ComboBox 데이터 설정
            SetComboBoxDataBinding();

            // 초기화
            SetControlInitialize();

            SetRodButtonDataBinding(DRPIrodPanel, rodCount);

            GetReferenceValue();
            //Button 초기화
            SetStopAndControlButtonAllClear(gbStop, gbControl, gbCardFrame);
            SetStopAndControlButtonAllInitialize();

            // 그리드 초기 설정
            SetDataGridViewInitialize();

            boolFormLoad = true;

            // 전력함 표시 Button 중 초기 선택은 1AC로 함
            //btnCardFrame1.PerformClick();

            // LCR-Meter 접속 체크
            if (m_MeasureProcess.DaqAndLCRMeterCommunication(Gini.GetValue("Device", "DAQDeviceName").Trim(), Gini.GetValue("Device", "LCRMeter_Addr").Trim(), ref boolDAQ, ref boolLCRMeter))
            {
                //측정 시작 중단에 따라 버튼 활성화/비활성화 설정
                SetMeasurementStartButtonEnabled(true);
            }
            else
            {
                btnMeasurementStart.Enabled = false;
                btnMeasurementStop.Enabled = false;
                btnSave.Enabled = false;
            }

            //ledRDAQ.Value = boolDAQ;
            //ledLCRMeter.Value = boolLCRMeter;
            ledRDAQ.On = boolDAQ;
            ledLCRMeter.On = boolLCRMeter;

            // 포인트
            cboHogi.Focus();

            //boolFormLoad = false;
        }

        // 20240401 한인석
        private int SetRodPanel(string plantName)
        {
            int returnVal = 0;
            switch (plantName)
            {
                case "고리 1발전소":
                    DRPIrodPanel = pnlRod33;
                    gbStop = gbStop33;
                    gbControl = gbControl33;
                    gbCardFrame = gbCardFrame33;
                    rboCoilA = rboCoilA33;
                    rboCoilB = rboCoilB33;
                    returnVal = 33;
                    break;
                case "한빛 1발전소":
                    DRPIrodPanel = pnlRod52;
                    gbStop = gbStop52;
                    gbControl = gbControl52;
                    gbCardFrame = gbCardFrame52;
                    rboCoilA = rboCoilA52;
                    rboCoilB = rboCoilB52;
                    returnVal = 52;
                    break;
                default:
                    DRPIrodPanel = pnlRod52;
                    gbStop = gbStop52;
                    gbControl = gbControl52;
                    gbCardFrame = gbCardFrame52;
                    rboCoilA = rboCoilA52;
                    rboCoilB = rboCoilB52;
                    returnVal = 52;
                    break;
            }
            DRPIrodPanel.Visible = true;
            return returnVal;
        }

        // 20240329 한인석
        private void SetRodButtonDataBinding(Panel p, int rodCount)
        {
            if (p == null) return;
            string[] rodKind = { "Shutdown", "Control" };
            for (int i = 0; i < rodKind.Length; i++)
            {
                var shutdownPanel = p.Controls.Find($"{p.Name}{rodKind[i]}", false).FirstOrDefault();
                if (shutdownPanel != null)
                {
                    GroupBox gb = shutdownPanel.Controls[0] as GroupBox;
                    if (gb != null)
                    {
                        string[] rodNames = Gini.GetValue("DRPI", $"DRPI{rodKind[i]}Item").Split(',');
                        string buttonNameTamplate = $"btn{rodKind[i]}Rod{rodCount}_";
                        foreach (Control c in gb.Controls)
                        {
                            Button b = c as Button;
                            if (b != null)
                            {
                                int rodIndex = Convert.ToInt32(b.Name.Substring(buttonNameTamplate.Length));
                                if (rodIndex < rodNames.Length)
                                {
                                    string name = rodNames[rodIndex];
                                    b.Text = name;
                                }
                                else
                                {
                                    b.Text = "N/A";
                                    b.Visible = false;
                                }
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Form Closing Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmDRPIDiagnosis_FormClosing(object sender, FormClosingEventArgs e)
        {
            Gini.SetValue("DRPI", "SelectDRPI_Hogi", cboHogi.SelectedItem.ToString().Trim());
            Gini.SetValue("DRPI", "SelectDRPI_OHDegree", teOhDegree2.Text.Trim());
        }

        /// <summary>
        /// ComboBox 데이터 설정
        /// </summary>
        private void SetComboBoxDataBinding()
        {
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

            for (int i = 0; i < strFrequencyItem.Length; i++)
            {
                cboFrequency.Items.Add(strFrequencyItem[i].Trim());
            }

            // 전압레벨
            string[] strVoltageLevelItem;
            cboVoltageLevel.Items.Clear();

            strVoltageLevelItem = Gini.GetValue("Combo", "VoltageLevelItem1").Split(',');
            for (int i = 0; i < strVoltageLevelItem.Length; i++)
            {
                cboVoltageLevel.Items.Add(strVoltageLevelItem[i].Trim());
            }
            strVoltageLevelItem = Gini.GetValue("Combo", "VoltageLevelItem2").Split(',');
            for (int i = 0; i < strVoltageLevelItem.Length; i++)
            {
                cboVoltageLevel.Items.Add(strVoltageLevelItem[i].Trim());
            }

            // 측정 횟수
            string[] strMeasurementCountItem = Gini.GetValue("Combo", "MeasurementCountItem").Split(',');

            cboMeasurementCount.Items.Clear();

            for (int i = 0; i < strMeasurementCountItem.Length; i++)
            {
                cboMeasurementCount.Items.Add(strMeasurementCountItem[i].Trim());
            }
        }

        /// <summary>
        /// 초기화
        /// </summary>
        private void SetControlInitialize()
        {
            // 온도 및 온도 증감 값
            dcmTemperature_ReferenceValue = Convert.ToDecimal(Gini.GetValue("ReferenceValue", "Temperature_ReferenceValue"));
            teTemperature_ReferenceValue.Text = Gini.GetValue("DeviReferenceValuece", "Temperature_ReferenceValue").Trim();
            teTemperatureUpDown_ReferenceValue.Text = Gini.GetValue("ReferenceValue", "TemperatureUpDown_ReferenceValue").Trim();

            // 주파수 기본 선택
            cboFrequency.SelectedIndex = 1;

            // 전압레벨 기본 선택
            cboVoltageLevel.SelectedItem = "1000mV";

            // 측정 횟수 기본 선택 (1회로 기본 설정)
            cboMeasurementCount.SelectedIndex = 0;

            // 측정 모드 
            rboNormalMode.Checked = true;
            rboWheatstoneMode.Checked = false;

            if (rboNormalMode.Checked)
                boolNormalMode = true;
            else
                boolNormalMode = false;

            // 측정 대상
            chkRdc.Checked = true;
            chkRac.Checked = true;
            chkL.Checked = true;
            chkC.Checked = false;
            chkQ.Checked = true;
        }

        /// <summary>
        /// 온도 변경 이벤트
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void teTemperature_ReferenceValue_TextChanged(object sender, EventArgs e)
        {
            dcmTemperature_ChangeValue = teTemperature_ReferenceValue.Text.Trim() == "" ? 0M : Convert.ToDecimal(teTemperature_ReferenceValue.Text.Trim());
            decimal dcmTemperatureUpDown_ReferenceValue = teTemperatureUpDown_ReferenceValue.Text.Trim() == "" ? 0M : Convert.ToDecimal(teTemperatureUpDown_ReferenceValue.Text.Trim());

            dcmTemperature_Measurement = (dcmTemperature_ChangeValue - dcmTemperature_ReferenceValue) * dcmTemperatureUpDown_ReferenceValue;
        }

        /// <summary>
        /// 온도 증/감 변경 이벤트
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void teTemperatureUpDown_ReferenceValue_TextChanged(object sender, EventArgs e)
        {
            decimal dcmTemperatureUpDown_ReferenceValue = teTemperatureUpDown_ReferenceValue.Text.Trim() == "" ? 0M : Convert.ToDecimal(teTemperatureUpDown_ReferenceValue.Text.Trim());

            dcmTemperature_Measurement = (dcmTemperature_ChangeValue - dcmTemperature_ReferenceValue) * dcmTemperatureUpDown_ReferenceValue;
        }

        private void SetStopAndControlButtonAllClear(params GroupBox[] arrayGb)
        {
            foreach (GroupBox gb in arrayGb)
            {
                if (gb != null)
                { 
                    SetStopAndControlButtonClear(gb);
                }
            }
        }

        private void SetStopAndControlButtonAllInitialize()
        {
            SetStopAndControlButtonInitialize(gbStop);
            SetStopAndControlButtonInitialize(gbControl);
        }

        private void SetStopAndControlButtonClear(GroupBox _gb)
        {
            if (_gb != null)
            {
                foreach (Control c in _gb.Controls)
                {
                    if (c is Button)
                    {
                        c.ForeColor = Color.Black;
                        c.BackgroundImage = Properties.Resources.버튼_4;
                        c.Tag = "FALSE";
                        c.Enabled = true;
                    }
                }
            }
        }

        /// <summary>
        /// 정지용/제어용 Button 초기화
        /// </summary>
        private void SetStopAndControlButtonInitialize(GroupBox _gb)
        {
            string strHog = cboHogi.SelectedItem == null ? "" : cboHogi.SelectedItem.ToString().Trim();
            string strOh_Degree = teOhDegree2.Text.Trim() == "" ? "" : string.Format("제 {0} 차", teOhDegree2.Text.Trim());
            string strDRPIGroup = rboCoilA.Checked ? rboCoilA.Text.Trim() : rboCoilB.Text.Trim();
            string strDRPIType = _gb.Text.Trim();

            //foreach (Control c in _gb.Controls)
            //{
            //    if (c.GetType().Name.Trim() == "Button")
            //    {
            //        string strControlName = c.Text.Trim();
            //        int intCount = m_db.GetDRPIDiagnosisDetailDataGridViewDataCount(strPlantName.Trim(), strHog.Trim(), strOh_Degree.Trim(), strDRPIGroup.Trim(), strDRPIType.Trim(), strControlName.Trim());

            //        if (intCount > 0) c.ForeColor = System.Drawing.Color.Red;
            //        else c.ForeColor = System.Drawing.Color.Black;

            //        c.BackgroundImage = global::Coil_Diagnostor.Properties.Resources.버튼_4;
            //        c.Tag = "FALSE";
            //        c.Enabled = true;
            //    }
            //}
            string[] diagnosisList = m_db.GetDRPIDiagnosisDetailDataGridViewExist(strPlantName.Trim(), strHog.Trim(), strOh_Degree.Trim(), strDRPIGroup.Trim(), strDRPIType.Trim()).ToString().Split(',');
            foreach (Control c in _gb.Controls)
            {
                if(c is Button)
                {
                    c.ForeColor = (diagnosisList.Contains(c.Text)) ? Color.Red : Color.Black;
                }
            }
        }

        /// <summary>
        /// 그리드 초기 설정
        /// </summary>
        private void SetDataGridViewInitialize()
        {
            try
            {
                //----- DataGridView Column 추가
                DataGridViewCheckBoxColumn Column = new DataGridViewCheckBoxColumn();
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

                Column.HeaderText = "";
                Column.Name = "Select";
                Column.Width = 40;
                Column.ReadOnly = true;
                Column.Visible = true;
                Column.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column.SortMode = DataGridViewColumnSortMode.NotSortable;
                Column.DefaultCellStyle.BackColor = System.Drawing.Color.Silver;

                Column1.HeaderText = "제어용";
                Column1.Name = "ControlName";
                Column1.Width = 60;
                Column1.ReadOnly = true;
                Column1.Visible = true;
                Column1.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column1.SortMode = DataGridViewColumnSortMode.NotSortable;
                Column1.DefaultCellStyle.BackColor = System.Drawing.Color.Silver;

                Column2.HeaderText = "코일명";
                Column2.Name = "CoilName";
                Column2.Width = 80;
                Column2.ReadOnly = true;
                Column2.Visible = true;
                Column2.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column2.SortMode = DataGridViewColumnSortMode.NotSortable;
                Column2.DefaultCellStyle.BackColor = System.Drawing.Color.Silver;

                Column3.HeaderText = "Index";
                Column3.Name = "CoilNumber";
                Column3.Width = 80;
                Column3.ReadOnly = true;
                Column3.Visible = false;
                Column3.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column3.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column4.HeaderText = "DC 저항(Ω)";
                Column4.Name = "DC_ResistanceValue";
                Column4.Width = 90;
                Column4.ReadOnly = true;
                Column4.Visible = chkRdc.Checked;
                Column4.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column4.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column5.HeaderText = "편차(Ω)";
                Column5.Name = "DC_Deviation";
                Column5.Width = 80;
                Column5.ReadOnly = true;
                Column5.Visible = chkRdc.Checked;
                Column5.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column5.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column6.HeaderText = "AC 저항(Ω)";
                Column6.Name = "AC_ResistanceValue";
                Column6.Width = 90;
                Column6.ReadOnly = true;
                Column6.Visible = chkRac.Checked;
                Column6.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column6.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column7.HeaderText = "편차(Ω)";
                Column7.Name = "AC_Deviation";
                Column7.Width = 80;
                Column7.ReadOnly = true;
                Column7.Visible = chkRac.Checked;
                Column7.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column7.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column8.HeaderText = "인덕턴스(mH)";
                Column8.Name = "L_InductanceValue";
                Column8.Width = 90;
                Column8.ReadOnly = true;
                Column8.Visible = chkL.Checked;
                Column8.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column8.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column9.HeaderText = "편차(mH)";
                Column9.Name = "L_Deviation";
                Column9.Width = 80;
                Column9.ReadOnly = true;
                Column9.Visible = chkL.Checked;
                Column9.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column9.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column10.HeaderText = "캐패시턴스(uF)";
                Column10.Name = "C_CapacitanceValue";
                Column10.Width = 100;
                Column10.ReadOnly = true;
                Column10.Visible = chkC.Checked;
                Column10.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column10.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column11.HeaderText = "편차(uF)";
                Column11.Name = "C_Deviation";
                Column11.Width = 80;
                Column11.ReadOnly = true;
                Column11.Visible = chkC.Checked;
                Column11.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column11.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column12.HeaderText = "Q Factor";
                Column12.Name = "Q_FactorValue";
                Column12.Width = 90;
                Column12.ReadOnly = true;
                Column12.Visible = chkQ.Checked;
                Column12.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column12.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column13.HeaderText = "편차";
                Column13.Name = "Q_Deviation";
                Column13.Width = 80;
                Column13.ReadOnly = true;
                Column13.Visible = chkQ.Checked;
                Column13.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column13.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column14.HeaderText = "결과";
                Column14.Name = "Result";
                Column14.Width = 70;
                Column14.ReadOnly = true;
                Column14.Visible = true;
                Column14.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column14.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column15.HeaderText = "그룹";
                Column15.Name = "DRPIGroup";
                Column15.Width = 80;
                Column15.ReadOnly = true;
                Column15.Visible = false;
                Column15.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column15.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column16.HeaderText = "Type";
                Column16.Name = "DRPIType";
                Column16.Width = 80;
                Column16.ReadOnly = true;
                Column16.Visible = false;
                Column16.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column16.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column17.HeaderText = "DAM";
                Column17.Name = "DAMName";
                Column17.Width = 80;
                Column17.ReadOnly = true;
                Column17.Visible = false;
                Column17.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column17.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column18.HeaderText = "DAQPinMap";
                Column18.Name = "DAQPinMap";
                Column18.Width = 500;
                Column18.ReadOnly = true;
                Column18.Visible = false;
                Column18.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                Column18.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column19.HeaderText = "채널";
                Column19.Name = "Channel";
                Column19.Width = 60;
                Column19.ReadOnly = true;
                Column19.Visible = false;
                Column19.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                Column19.SortMode = DataGridViewColumnSortMode.NotSortable;

                dgvMeasurement.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] { Column, Column1, Column2, Column3
                    , Column4, Column5, Column6, Column7, Column8, Column9, Column10, Column11, Column12, Column13, Column14
                    , Column15, Column16, Column17, Column18, Column19 });

                dgvMeasurement.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing;
                dgvMeasurement.ColumnHeadersHeight = 40;
                dgvMeasurement.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dgvMeasurement.RowHeadersWidth = 40;

                // DataGridView 기타 세팅
                dgvMeasurement.AllowUserToAddRows = false; // Row 추가 기능 미사용
                dgvMeasurement.AllowUserToDeleteRows = false; // Row 삭제 기능 미사용

                // CheckBox 세팅
                allCheck.Name = "allCheck";
                allCheck.CheckedChanged += new EventHandler(AllCheckClick);
                allCheck.Size = new Size(13, 13);
                allCheck.Location = new Point(((dgvMeasurement.Columns[0].Width / 2) - (allCheck.Width / 2)), (dgvMeasurement.ColumnHeadersHeight / 2) - (allCheck.Height / 2));
                dgvMeasurement.Controls.Add(allCheck); // DataGridView에 CheckBox 추가( 헤더용 )
                allCheck.Checked = true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
        }

        /// <summary>
        /// DataGridView의 모든 CheckBoxCell Checked값 적용
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void AllCheckClick(object sender, EventArgs e)
        {
            if (isCheck)
            {
                if (allCheck.Checked)
                    for (int i = 0; i < dgvMeasurement.Rows.Count; i++)
                        dgvMeasurement.Rows[i].Cells[0].Value = true;
                else
                    for (int i = 0; i < dgvMeasurement.Rows.Count; i++)
                        dgvMeasurement.Rows[i].Cells[0].Value = false;

                dgvMeasurement.EndEdit(DataGridViewDataErrorContexts.Commit); // << 이거 안할경우 선택된 Cell이 CheckBoxCell일 경우 변화가 없는것처럼 보임
            }
        }

        /// <summary>
        /// 측정 시작 중단에 따라 버튼 활성화/비활성화 설정
        /// </summary>
        /// <param name="boolEnabled"></param>
        public void SetMeasurementStartButtonEnabled(bool boolEnabled)
        {
            btnMeasurementStart.Enabled = boolEnabled;
            btnMeasurementStop.Enabled = !boolEnabled;
            btnSave.Enabled = boolEnabled;

            if (boolEnabled)
            {
                btnMeasurementStart.ForeColor = System.Drawing.Color.White;
                btnSave.ForeColor = System.Drawing.Color.White;

                btnMeasurementStop.ForeColor = System.Drawing.Color.Silver;
            }
            else
            {
                btnMeasurementStart.ForeColor = System.Drawing.Color.Silver;
                btnSave.ForeColor = System.Drawing.Color.Silver;

                btnMeasurementStop.ForeColor = System.Drawing.Color.White;
            }
        }

        /// <summary>
        /// 카드 프레임 Button Click Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCardFrame_Click(object sender, EventArgs e)
        {
            if (boolMeasurementStart) return;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;

            try
            {
                // Button 선택 여부에 따라 글자 색변경 (선택 시 White, 선택 취소 시 Black)
                if (((Button)sender).ForeColor == System.Drawing.Color.Black)
                {
                    if (((Button)sender).ForeColor != System.Drawing.Color.Red)
                        ((Button)sender).ForeColor = System.Drawing.Color.White;

                    ((Button)sender).BackgroundImage = global::Coil_Diagnostor.Properties.Resources.버튼_3;

                    // 선택 Card Frame외 Card Frame은 기본 Button Color 및 배경이미지 설정
                    SetCardFrameButtonColor(((Button)sender));
                }
                else
                {
                    // 20240401 한인석
                    // 존재 이유 없는 코드
                    //if (((Button)sender).ForeColor != System.Drawing.Color.Red)
                    //    ((Button)sender).ForeColor = System.Drawing.Color.Black;

                    //((Button)sender).BackgroundImage = global::Coil_Diagnostor.Properties.Resources.버튼_4;

                    // 그리드 초기화
                    dgvMeasurement.Rows.Clear();

                    SetStopAndControlButtonAllClear(gbStop, gbControl, gbCardFrame);
                    SetStopAndControlButtonAllInitialize();
                    return;
                }

                // 그리드 초기화
                dgvMeasurement.Rows.Clear();

                // 정지용 Button 초기화
                SetStopAndControlButtonAllClear(gbStop, gbControl);

                SelectStopAndControlButtonSetting(((Button)sender));

            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            System.Windows.Forms.Cursor.Current = Cursors.Default;
        }

        /// <summary>
        /// 선택 Card Frame외 Card Frame은 기본 Button Color 및 배경이미지 설정
        /// </summary>
        private void SetCardFrameButtonColor(Button _btn)
        {
            foreach (Control c in gbCardFrame.Controls)
            {
                if (c.GetType().Name.Trim() == "Button")
                {
                    if (c.Name.Trim() != _btn.Name.Trim())
                    {
                        c.ForeColor = System.Drawing.Color.Black;
                        c.BackgroundImage = global::Coil_Diagnostor.Properties.Resources.버튼_4;
                    }
                }
            }
        }

        /// <summary>
        /// Card Frame 선택에 따라 정지용 또는 제어용 자동 선택
        /// </summary>
        /// <param name="_gb"></param>
        //private void SelectStopAndControlButtonSetting(GroupBox _gb, string[] arraySelectButtonList)
        private void SelectStopAndControlButtonSetting(Button button)
        {
            string framebtnNameTamplate = "btnCardFrame00_";
            string frameName = $"DRPIButtonCardFrame{((Button)button).Name.Substring(framebtnNameTamplate.Length)}";
            string[] arrayRod = Gini.GetValue("DRPI", frameName).Split('/',',');
            GroupBox[] arrayGb = {gbStop, gbControl};
            foreach (GroupBox gb in arrayGb)
            {
                if (gb != null)
                {
                    foreach (Control c in gb.Controls)
                    {
                        if (c is Button)
                        {
                            Button b = c as Button;
                            if (arrayRod.Contains(b.Text.Trim()))
                            {
                                b.ForeColor = Color.White;
                                b.BackgroundImage = Properties.Resources.버튼_3;
                                b.PerformClick();
                            }
                            else
                            {
                                b.ForeColor = Color.Black;
                                b.BackgroundImage = Properties.Resources.버튼_4;
                            }
                            c.Enabled = false;
                        }
                    }
                }
            }
        }


        /// <summary>
        /// 정지용/제어용 Button Click Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnStopControl_Click(object sender, EventArgs e)
        {
            if (boolMeasurementStart) return;

            try
            {
                // Button 선택 여부에 따라 글자 색변경 (선택 시 White, 선택 취소 시 Black)
                if (((Button)sender).Tag.ToString() == "FALSE")
                {
                    if (((Button)sender).ForeColor != System.Drawing.Color.Red)
                        ((Button)sender).ForeColor = System.Drawing.Color.White;

                    ((Button)sender).BackgroundImage = global::Coil_Diagnostor.Properties.Resources.버튼_3;

                    ((Button)sender).Tag = "TRUE";
                }
                else
                {
                    if (((Button)sender).ForeColor != System.Drawing.Color.Red)
                        ((Button)sender).ForeColor = System.Drawing.Color.Black;

                    ((Button)sender).BackgroundImage = global::Coil_Diagnostor.Properties.Resources.버튼_4;

                    ((Button)sender).Tag = "FALSE";
                }

                //int intRowAddIndex = 0;

                // 제어용 선택에 따라 그리드 행 추가
               // if (Convert.ToInt32(Regex.Replace(((Button)sender).Name.Trim(), @"\D", "")) < 100)
                    SetDataGridViewRowAdd(0, ((Button)sender));
                //else
                //    SetDataGridViewRowAdd(0, ((Button)sender), gbControl);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
        }

        /// <summary>
        /// 제어용 선택에 따라 그리드 행 추가
        /// </summary>
        /// <param name="_intRowStartIndex"></param>
        /// <param name="_strControlName"></param>
        private void SetDataGridViewRowAdd(int _intRowStartIndex, Button _btn)
        {
            try
            {
                if (_btn == null || _btn.Parent == null) return;
                GroupBox gb = _btn.Parent as GroupBox;

                int intRowStartIndex = 0, intRowEndCount = 0, intRow = 0, intCoilNameCount = 1, intCoilNameItemIndex = 0;//, intNextControlRodNameNumber = 0,
                int intControlButtonCount = 0, intShutdownButtonCount = 0; ;
                string strCoilNameItme = "", strCoilName = "", strNextControlRodText = "";
                string[] arrayCoilNameItme;

                // 그리드 행 추가 시작 Index 구하기
                int intItemCount = cboMeasurementCount.SelectedItem == null || cboMeasurementCount.SelectedItem.ToString().Trim() == "" ? 1 : (Convert.ToInt32(Regex.Replace(cboMeasurementCount.SelectedItem.ToString().Trim(), @"\D", "")));
                int intButtonSelectNumber = Convert.ToInt32(Regex.Replace(_btn.Name.Trim(), @"\D", ""));
                int intButtonSelectBeforeCount = 0, intButtonSelectAfterCount = 0;

                if(gb.Text.Trim() == "정지용")
                {
                    foreach (Control c in gbControl.Controls)
                    {
                        if (c is Button)
                        {
                            if (c.Tag.ToString().Trim() == "TRUE")
                                intControlButtonCount++;
                        }
                    }
                }
                else //if (gb.Text.Trim() == "제어용")
                {
                    foreach (Control c in gbStop.Controls)
                    {
                        if (c is Button)
                        {
                            if (c.Tag.ToString().Trim() == "TRUE")
                                intShutdownButtonCount++;
                        }
                    }
                }

                // 정지용 / 제어용 그룹박스에 선택된 항목들 체크
                foreach (Control c in gb.Controls)
                {
                    if (c.GetType().Name.Trim() == "Button")
                    {
                        if (Convert.ToInt32(Regex.Replace(c.Name.Trim(), @"\D", "")) < intButtonSelectNumber && c.Tag.ToString().Trim() == "TRUE")
                            intButtonSelectBeforeCount++;
                        else if (Convert.ToInt32(Regex.Replace(c.Name.Trim(), @"\D", "")) > intButtonSelectNumber && c.Tag.ToString().Trim() == "TRUE")
                            intButtonSelectAfterCount++;
                    }
                }
                if (gb.Text.Trim() == "정지용")
                {
                    intRowEndCount = intItemCount * 10;
                    strCoilNameItme = Gini.GetValue("DRPI", "DRPIStopCoilName_Item").Trim();
                }
                else
                {
                    intRowEndCount = intItemCount * 21;
                    strCoilNameItme = Gini.GetValue("DRPI", "DRPIControlCoilName_Item").Trim();
                }

                // 가져온 코일명 배열로 담기
                arrayCoilNameItme = strCoilNameItme.Split(',');

                intRow = intControlButtonCount * 21 * intItemCount + intRowEndCount * intButtonSelectBeforeCount;

                // 시작 코일명 가져온다.
                strCoilName = rboCoilA.Checked ? rboCoilA.Text.Trim() : rboCoilB.Text.Trim();

                if (_btn.Tag.ToString().Trim() == "TRUE")
                {
                    for (int i = 0; i < intRowEndCount; i++)
                    {
                        if (intButtonSelectAfterCount > 0 || intShutdownButtonCount > 0)
                        {
                            // 그리드 행 삽입
                            dgvMeasurement.Rows.Insert(intRow, 1);

                            dgvMeasurement.Rows[intRow].Cells["Select"].Value = allCheck.Checked;
                            dgvMeasurement.Rows[intRow].Cells["ControlName"].Value = _btn.Text.Trim();
                            dgvMeasurement.Rows[intRow].Cells["DRPIGroup"].Value = rboCoilA.Checked ? "A" : "B";
                            dgvMeasurement.Rows[intRow].Cells["DRPIType"].Value = gb.Text.Trim();

                            if (intCoilNameCount < intItemCount)
                            {
                                dgvMeasurement.Rows[intRow].Cells["CoilName"].Value = strCoilName.Trim() + arrayCoilNameItme[intCoilNameItemIndex].Trim();
                                dgvMeasurement.Rows[intRow].Cells["CoilNumber"].Value = intCoilNameCount.ToString().Trim();
                                intCoilNameCount++;
                            }
                            else
                            {
                                dgvMeasurement.Rows[intRow].Cells["CoilName"].Value = strCoilName.Trim() + arrayCoilNameItme[intCoilNameItemIndex].Trim();
                                dgvMeasurement.Rows[intRow].Cells["CoilNumber"].Value = intCoilNameCount.ToString().Trim();

                                intCoilNameCount = 1;
                                intCoilNameItemIndex++;
                            }

                            dgvMeasurement.Rows[intRow].Cells["DC_ResistanceValue"].Value = "";
                            dgvMeasurement.Rows[intRow].Cells["DC_Deviation"].Value = "";
                            dgvMeasurement.Rows[intRow].Cells["AC_ResistanceValue"].Value = "";
                            dgvMeasurement.Rows[intRow].Cells["AC_Deviation"].Value = "";
                            dgvMeasurement.Rows[intRow].Cells["L_InductanceValue"].Value = "";
                            dgvMeasurement.Rows[intRow].Cells["L_Deviation"].Value = "";
                            dgvMeasurement.Rows[intRow].Cells["C_CapacitanceValue"].Value = "";
                            dgvMeasurement.Rows[intRow].Cells["C_Deviation"].Value = "";
                            dgvMeasurement.Rows[intRow].Cells["Q_FactorValue"].Value = "";
                            dgvMeasurement.Rows[intRow].Cells["Q_Deviation"].Value = "";
                            dgvMeasurement.Rows[intRow].Cells["Result"].Value = "";
                            intRow++;
                        }
                        else
                        {
                            // 그리드 행 추가
                            intRow = dgvMeasurement.Rows.Add();

                            dgvMeasurement.Rows[intRow].Cells["Select"].Value = allCheck.Checked;
                            dgvMeasurement.Rows[intRow].Cells["ControlName"].Value = _btn.Text.Trim();
                            dgvMeasurement.Rows[intRow].Cells["DRPIGroup"].Value = rboCoilA.Checked ? "A" : "B";
                            dgvMeasurement.Rows[intRow].Cells["DRPIType"].Value = gb.Text.Trim();

                            if (intCoilNameCount < intItemCount)
                            {
                                dgvMeasurement.Rows[intRow].Cells["CoilName"].Value = strCoilName.Trim() + arrayCoilNameItme[intCoilNameItemIndex].Trim();
                                dgvMeasurement.Rows[intRow].Cells["CoilNumber"].Value = intCoilNameCount.ToString().Trim();
                                intCoilNameCount++;
                            }
                            else
                            {
                                dgvMeasurement.Rows[intRow].Cells["CoilName"].Value = strCoilName.Trim() + arrayCoilNameItme[intCoilNameItemIndex].Trim();
                                dgvMeasurement.Rows[intRow].Cells["CoilNumber"].Value = intCoilNameCount.ToString().Trim();

                                intCoilNameCount = 1;
                                intCoilNameItemIndex++;
                            }

                            dgvMeasurement.Rows[intRow].Cells["DC_ResistanceValue"].Value = "";
                            dgvMeasurement.Rows[intRow].Cells["DC_Deviation"].Value = "";
                            dgvMeasurement.Rows[intRow].Cells["AC_ResistanceValue"].Value = "";
                            dgvMeasurement.Rows[intRow].Cells["AC_Deviation"].Value = "";
                            dgvMeasurement.Rows[intRow].Cells["L_InductanceValue"].Value = "";
                            dgvMeasurement.Rows[intRow].Cells["L_Deviation"].Value = "";
                            dgvMeasurement.Rows[intRow].Cells["C_CapacitanceValue"].Value = "";
                            dgvMeasurement.Rows[intRow].Cells["C_Deviation"].Value = "";
                            dgvMeasurement.Rows[intRow].Cells["Q_FactorValue"].Value = "";
                            dgvMeasurement.Rows[intRow].Cells["Q_Deviation"].Value = "";
                            dgvMeasurement.Rows[intRow].Cells["Result"].Value = "";
                        }
                    }
                }
                else
                {
                    // 그리드 행 삭제
                    for (int i = dgvMeasurement.RowCount - 1; i >= 0; i--)
                    {
                        if (dgvMeasurement.Rows[i].Cells["ControlName"].Value.ToString().Trim() == _btn.Text.Trim())
                            dgvMeasurement.Rows.RemoveAt(i);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
        }

        /// <summary>
        /// 주파수 Selelcted Index Changed Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cboFrequency_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (boolMeasurementStart) return;

            if (cboFrequency.SelectedItem == null || cboFrequency.SelectedItem.ToString().Trim() == "")
                dcmFrequency = 0M;
            else
            {
                switch (cboFrequency.SelectedItem.ToString().Trim())
                {
                    case "100 Hz":
                        dcmFrequency = 100M;
                        break;
                    case "120 Hz":
                        dcmFrequency = 120M;
                        break;
                    case "1 kHz":
                        dcmFrequency = 1000M;
                        break;
                    case "10 kHz":
                        dcmFrequency = 10000M;
                        break;
                    case "100 kHz":
                        dcmFrequency = 100000M;
                        break;
                }
            }
        }

        /// <summary>
        /// 전압레벨 Selelcted Index Changed Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cboVoltageLevel_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (boolMeasurementStart) return;

            if (cboFrequency.SelectedItem == null || cboFrequency.SelectedItem.ToString().Trim() == "")
                dcmVoltageLevel = 1;
            else
            {
                decimal dcmSelectVoltageLevel = Convert.ToDecimal(Regex.Replace(cboVoltageLevel.SelectedItem.ToString().Trim(), @"\D", ""));
                dcmVoltageLevel = dcmSelectVoltageLevel / 1000M;
            }
        }

        /// <summary>
        /// 호기 Selected Index Changed Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cboHogi_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (boolMeasurementStart || !boolFormLoad) return;

            /*
            string strSelectHogi = cboHogi.SelectedItem == null ? "" : cboHogi.SelectedItem.ToString().Trim();

            string[] arrayMaxOhDegree = m_db.GetDRPIDiagnosisHogiMaxOhDegree(strPlantName.Trim(), strSelectHogi.Trim()).Split(' ');

            for (int i = 0; i < arrayMaxOhDegree.Length; i++)
            {
                if (arrayMaxOhDegree[i].Trim() != "제" && arrayMaxOhDegree[i].Trim() != "차")
                    teOhDegree2.Text = arrayMaxOhDegree[i].Trim();
            }
             * */
            teOhDegree2.Text = Gini.GetValue("DRPI", "SelectDRPI_OHDegree").Trim();

            // 기준값 가져온다
            GetReferenceValue();

            // 그리드 행 초기화
            dgvMeasurement.Rows.Clear();

            SetStopAndControlButtonAllClear(gbStop, gbControl, gbCardFrame);
            SetStopAndControlButtonAllInitialize();
        }

        /// <summary>
        /// Oh Degree 변경 후 Control에서 이동 시 Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void teOhDegree2_Leave(object sender, EventArgs e)
        {
            if (boolMeasurementStart || !boolFormLoad) return;

            // 기준값 가져온다
            GetReferenceValue();

            // 그리드 행 초기화
            dgvMeasurement.Rows.Clear();

            SetStopAndControlButtonAllClear(gbStop, gbControl, gbCardFrame);
            SetStopAndControlButtonAllInitialize();
        }

        /// <summary>
        /// 기준값 가져온다
        /// </summary>
        private void GetReferenceValue()
        {
            string strSelectHogi = cboHogi.SelectedItem == null ? "" : cboHogi.SelectedItem.ToString().Trim();
            string strSelectOhDegree = lblOhDegree1.Text.Trim() + " " + teOhDegree2.Text.Trim() + " " + lblOhDegree3.Text.Trim();

            DataTable dt = new DataTable();

            // 해당 차수의 기준 값 유무 체크
			if ((m_db.GetDRPIReferenceValueDataCount(strPlantName.Trim(), strSelectHogi.Trim(), strSelectOhDegree.Trim())) > 0)
            {
                // 해당 차수의 기준 값을 가져온다.
				dt = m_db.GetDRPIReferenceValueDataInfo(strPlantName, strSelectHogi, strSelectOhDegree);
			}
			else
            {
                // MAX OH 차수를 가져온다.
                int intMaxOhDegree = m_db.GetDRPIReferenceValueDataMaxOhDegree(strPlantName.Trim(), strSelectHogi.Trim());

                if (intMaxOhDegree > 0)
                {
                    // MAX 차수의 기준 값을 가져온다.
                    dt = m_db.GetDRPIReferenceValueDataInfo(strPlantName, strSelectHogi, intMaxOhDegree.ToString().Trim());
                }
                else
                {
                    // 초기값 0 차수의 기준 값을 가져온다.
                    dt = m_db.GetDRPIReferenceValueDataInfo(strPlantName.Trim(), "초기값", "제 0 차");
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
            }
        }

        /// <summary>
        /// 측정 횟수 Selected Index Changed Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cboMeasurementCount_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (boolMeasurementStart) return;

            try
            {
                // 그리드 행 초기화
                dgvMeasurement.Rows.Clear();

                SetStopAndControlButtonAllClear(gbStop, gbControl, gbCardFrame);
                SetStopAndControlButtonAllInitialize();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
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
                            dgvMeasurement.Columns[3].Visible = true;
                            dgvMeasurement.Columns[4].Visible = true;
                        }
                        else
                        {
                            dgvMeasurement.Columns[3].Visible = false;
                            dgvMeasurement.Columns[4].Visible = false;
                        }
                        break;
                    case "chkRac":
                        if (((CheckBox)sender).Checked)
                        {
                            dgvMeasurement.Columns[5].Visible = true;
                            dgvMeasurement.Columns[6].Visible = true;
                        }
                        else
                        {
                            dgvMeasurement.Columns[5].Visible = false;
                            dgvMeasurement.Columns[6].Visible = false;
                        }
                        break;
                    case "chkL":
                        if (((CheckBox)sender).Checked)
                        {
                            dgvMeasurement.Columns[7].Visible = true;
                            dgvMeasurement.Columns[8].Visible = true;
                        }
                        else
                        {
                            dgvMeasurement.Columns[7].Visible = false;
                            dgvMeasurement.Columns[8].Visible = false;
                        }
                        break;
                    case "chkC":
                        if (((CheckBox)sender).Checked)
                        {
                            dgvMeasurement.Columns[9].Visible = true;
                            dgvMeasurement.Columns[10].Visible = true;
                        }
                        else
                        {
                            dgvMeasurement.Columns[9].Visible = false;
                            dgvMeasurement.Columns[10].Visible = false;
                        }
                        break;
                    case "chkQ":
                        if (((CheckBox)sender).Checked)
                        {
                            dgvMeasurement.Columns[11].Visible = true;
                            dgvMeasurement.Columns[12].Visible = true;
                        }
                        else
                        {
                            dgvMeasurement.Columns[11].Visible = false;
                            dgvMeasurement.Columns[12].Visible = false;
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
        }

        /// <summary>
        /// 일반모드 Checked Changed Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void rboNormalMode_CheckedChanged(object sender, EventArgs e)
        {
            if (rboNormalMode.Checked)
                boolNormalMode = true;
            else
                boolNormalMode = false;
        }

        /// <summary>
        /// 휘스톤 Checked Changed Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void rboWheatstoneMode_CheckedChanged(object sender, EventArgs e)
        {
            if (rboNormalMode.Checked)
                boolNormalMode = true;
            else
                boolNormalMode = false;
        }

        /// <summary>
        /// 닫기 Button Evnet
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnClose_Click(object sender, EventArgs e)
        {
            try
            {
                btnClose.ForeColor = System.Drawing.Color.Blue;

                if (!boolMeasurementStart)
                    this.Close();
                else
                {
                    frmMB.lblMessage.Text = "측정 중이므로 측정 종료 또는 정지 후 닫기하십시오.";
                    frmMB.TopMost = true;
                    frmMB.ShowDialog();
                }

                btnClose.ForeColor = System.Drawing.Color.White;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
        }

        /// <summary>
        /// 저장 Button Evnet
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSave_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;

            bool boolSave = false;
            bool boolHeader = false;
            bool boolDetail = false;

            try
            {
                // 그리드 행 확인
                if (dgvMeasurement.RowCount <= 0)
                {
                    frmMB.lblMessage.Text = "저장할 데이터가 없습니다.";
                    frmMB.TopMost = true;
                    frmMB.ShowDialog();
                    return;
                }

                string strHogi = cboHogi.SelectedItem == null ? "" : cboHogi.SelectedItem.ToString().Trim();
                string strOhDegree = lblOhDegree1.Text.Trim() + " " + teOhDegree2.Text.Trim() + " " + lblOhDegree3.Text.Trim();                
                string strDRPIGroup = rboCoilA.Checked ? "A" : "B";
                string strDRPIType = "";

                // 측정 데이터 Header Table 저장 처리
                decimal dTemperature_ReferenceValue = teTemperature_ReferenceValue.Text.Trim() == "" ? 0.00M : Convert.ToDecimal(teTemperature_ReferenceValue.Text.Trim());
                decimal dTemperatureUpDown_ReferenceValue = teTemperatureUpDown_ReferenceValue.Text.Trim() == "" ? 0.00M : Convert.ToDecimal(teTemperatureUpDown_ReferenceValue.Text.Trim());
                string strFrequency = cboFrequency.SelectedItem == null ? "" : cboFrequency.SelectedItem.ToString().Trim();
                string stVoltageLevel = cboVoltageLevel.SelectedItem == null ? "" : cboVoltageLevel.SelectedItem.ToString().Trim();
                int intMeasurementCount = cboMeasurementCount.SelectedItem == null ? 0 : Convert.ToInt32(Regex.Replace(cboMeasurementCount.SelectedItem.ToString().Trim(), @"\D", ""));

                int intItem_Rdc = 0, intItem_Rac = 0, intItem_L = 0, intItem_C = 0, intItem_Q = 0;

                if (chkRdc.Checked) intItem_Rdc = 0; else intItem_Rdc = 1;
                if (chkRac.Checked) intItem_Rac = 0; else intItem_Rac = 1;
                if (chkL.Checked) intItem_L = 0; else intItem_L = 1;
                if (chkC.Checked) intItem_C = 0; else intItem_C = 1;
                if (chkQ.Checked) intItem_Q = 0; else intItem_Q = 1;

                string strMeasurementMode = rboNormalMode.Checked ? "일반모드" : "휘스톤모드";

                for (int i = 0; i < dgvMeasurement.RowCount; i++)
                {
                    if (dgvMeasurement.Rows[i].Cells["Result"].Value.ToString().Trim() == "부적합"
                        && (strMeasurementResult.Trim() == "" || strMeasurementResult.Trim() == "의심" || strMeasurementResult.Trim() == "적합"))
                        strMeasurementResult = "부적합";
                    else if (dgvMeasurement.Rows[i].Cells["Result"].Value.ToString().Trim() == "의심"
                        && (strMeasurementResult.Trim() == "" || strMeasurementResult.Trim() == "적합"))
                        strMeasurementResult = "의심";
                    else if (dgvMeasurement.Rows[i].Cells["Result"].Value.ToString().Trim() == "적합" && strMeasurementResult.Trim() == "")
                        strMeasurementResult = "적합";
                }

                // 제어용 저장
                foreach (Control c in gbControl.Controls)
                {
                    if (c.GetType().Name.Trim() == "Button" && (c.ForeColor == System.Drawing.Color.White || c.ForeColor == System.Drawing.Color.Red))
                    {
                        if ((m_db.GetDRPIDiagnosisHeaderDataGridViewDataCount(strPlantName.Trim(), strHogi, strOhDegree, strDRPIGroup, gbControl.Text.Trim(), c.Text.Trim())) > 0)
                        {
                            // 기존 데이터 저장
                            boolHeader = m_db.DeleteDRPIDiagnosisHeaderDataGridViewDataInfo(strPlantName.Trim(), strHogi, strOhDegree, strDRPIGroup, gbControl.Text.Trim(), c.Text.Trim());
                        }

                        // 데이터 저장
                        boolHeader = m_db.InsertDRPIDiagnosisHeaderDataGridViewDataInfo(strPlantName.Trim(), strHogi, strOhDegree, strDRPIGroup, gbControl.Text.Trim()
                            , c.Text.Trim(), dTemperature_ReferenceValue, dTemperatureUpDown_ReferenceValue, strFrequency, stVoltageLevel
                            , intMeasurementCount, intItem_Rdc, intItem_Rac, intItem_L, intItem_C, intItem_Q, strMeasurementMode, strMeasurementDate.Trim()
                            , strMeasurementResult.Trim());

                        c.ForeColor = System.Drawing.Color.Red;
                    }
                }

                // 정지용 저장
                foreach (Control c in gbStop.Controls)
                {
                    if (c.GetType().Name.Trim() == "Button" && (c.ForeColor == System.Drawing.Color.White || c.ForeColor == System.Drawing.Color.Red))
                    {
                        if ((m_db.GetDRPIDiagnosisHeaderDataGridViewDataCount(strPlantName.Trim(), strHogi, strOhDegree, strDRPIGroup, gbStop.Text.Trim(), c.Text.Trim())) > 0)
                        {
                            // 기존 데이터 저장
                            boolHeader = m_db.DeleteDRPIDiagnosisHeaderDataGridViewDataInfo(strPlantName.Trim(), strHogi, strOhDegree, strDRPIGroup, gbStop.Text.Trim(), c.Text.Trim());
                        }

                        // 데이터 저장
                        boolHeader = m_db.InsertDRPIDiagnosisHeaderDataGridViewDataInfo(strPlantName.Trim(), strHogi, strOhDegree, strDRPIGroup, gbStop.Text.Trim()
                            , c.Text.Trim(), dTemperature_ReferenceValue, dTemperatureUpDown_ReferenceValue, strFrequency, stVoltageLevel
                            , intMeasurementCount, intItem_Rdc, intItem_Rac, intItem_L, intItem_C, intItem_Q, strMeasurementMode, strMeasurementDate.Trim()
                            , strMeasurementResult.Trim());

                        c.ForeColor = System.Drawing.Color.Red;
                    }
                }

                boolSave = boolHeader;

                // 측정 데이터 Detail Table 저장 처리
                string strControlName = "", strCoilName = "", strResult = "";
				int intCoilNumber = 0, intSeqNumber = 0, intControlSeqNumber = 0;
                decimal dcmDC_ResistanceValue = 0.000M, dcmDC_Deviation = 0.000M, dcmAC_ResistanceValue = 0.000M, dcmAC_Deviation = 0.000M
                    , dcmL_Inductace = 0.000M, dcmL_Deviation = 0.000M, dcmC_CapacitanceValue = 0.000M, dcmC_Deviation = 0.000M
                    , dcmQ_FactorValue = 0.000M, dcmQ_Deviation = 0.000M;

                string[] controlRod = Gini.GetValue("DRPI", "DRPIControlItem").Split(',');
                string[] shutdownRod = Gini.GetValue("DRPI", "DRPIShutdownItem").Split(',');

                for (int i = 0; i < dgvMeasurement.RowCount; i++)
                {
                    strControlName = dgvMeasurement.Rows[i].Cells["ControlName"].Value.ToString();
                    strCoilName = dgvMeasurement.Rows[i].Cells["CoilName"].Value.ToString();
                    intCoilNumber = Convert.ToInt32(dgvMeasurement.Rows[i].Cells["CoilNumber"].Value.ToString());
                    strResult = dgvMeasurement.Rows[i].Cells["Result"].Value.ToString();
                    strDRPIType = dgvMeasurement.Rows[i].Cells["DRPIType"].Value.ToString();

					// 코일 순번 설정
					if (strCoilName.Trim() != "")
						intSeqNumber = Convert.ToInt32(Regex.Replace(strCoilName, @"\D", ""));

					// 제어봉 순번 가져오기
					intControlSeqNumber = GetControlSeqNumber(ref controlRod, ref shutdownRod, strDRPIType, strControlName);

                    dcmDC_ResistanceValue = dgvMeasurement.Rows[i].Cells["DC_ResistanceValue"].Value == null || dgvMeasurement.Rows[i].Cells["DC_ResistanceValue"].Value.ToString() == ""
                        ? 0.000M : Convert.ToDecimal(dgvMeasurement.Rows[i].Cells["DC_ResistanceValue"].Value.ToString());
                    dcmDC_Deviation = dgvMeasurement.Rows[i].Cells["DC_Deviation"].Value == null || dgvMeasurement.Rows[i].Cells["DC_Deviation"].Value.ToString() == ""
                        ? 0.000M : Convert.ToDecimal(dgvMeasurement.Rows[i].Cells["DC_Deviation"].Value.ToString());
                    dcmAC_ResistanceValue = dgvMeasurement.Rows[i].Cells["AC_ResistanceValue"].Value == null || dgvMeasurement.Rows[i].Cells["AC_ResistanceValue"].Value.ToString() == ""
                        ? 0.000M : Convert.ToDecimal(dgvMeasurement.Rows[i].Cells["AC_ResistanceValue"].Value.ToString());
                    dcmAC_Deviation = dgvMeasurement.Rows[i].Cells["AC_Deviation"].Value == null || dgvMeasurement.Rows[i].Cells["AC_Deviation"].Value.ToString() == ""
                        ? 0.000M : Convert.ToDecimal(dgvMeasurement.Rows[i].Cells["AC_Deviation"].Value.ToString());
                    dcmL_Inductace = dgvMeasurement.Rows[i].Cells["L_InductanceValue"].Value == null || dgvMeasurement.Rows[i].Cells["L_InductanceValue"].Value.ToString() == ""
                        ? 0.000M : Convert.ToDecimal(dgvMeasurement.Rows[i].Cells["L_InductanceValue"].Value.ToString());
                    dcmL_Deviation = dgvMeasurement.Rows[i].Cells["L_Deviation"].Value == null || dgvMeasurement.Rows[i].Cells["L_Deviation"].Value.ToString() == ""
                        ? 0.000M : Convert.ToDecimal(dgvMeasurement.Rows[i].Cells["L_Deviation"].Value.ToString());
                    dcmC_CapacitanceValue = dgvMeasurement.Rows[i].Cells["C_CapacitanceValue"].Value == null || dgvMeasurement.Rows[i].Cells["C_CapacitanceValue"].Value.ToString() == ""
                        ? 0.000000M : Convert.ToDecimal(dgvMeasurement.Rows[i].Cells["C_CapacitanceValue"].Value.ToString());
                    dcmC_Deviation = dgvMeasurement.Rows[i].Cells["C_Deviation"].Value == null || dgvMeasurement.Rows[i].Cells["C_Deviation"].Value.ToString() == ""
                        ? 0.000000M : Convert.ToDecimal(dgvMeasurement.Rows[i].Cells["C_Deviation"].Value.ToString());
                    dcmQ_FactorValue = dgvMeasurement.Rows[i].Cells["Q_FactorValue"].Value == null || dgvMeasurement.Rows[i].Cells["Q_FactorValue"].Value.ToString() == ""
                        ? 0.000M : Convert.ToDecimal(dgvMeasurement.Rows[i].Cells["Q_FactorValue"].Value.ToString());
                    dcmQ_Deviation = dgvMeasurement.Rows[i].Cells["Q_Deviation"].Value == null || dgvMeasurement.Rows[i].Cells["Q_Deviation"].Value.ToString() == ""
                        ? 0.000M : Convert.ToDecimal(dgvMeasurement.Rows[i].Cells["Q_Deviation"].Value.ToString());

                    if ((m_db.GetDRPIDiagnosisDetailDataGridViewDataCount(strPlantName.Trim(), strHogi, strOhDegree, strDRPIGroup, strDRPIType, strControlName, strCoilName, intCoilNumber)) > 0)
                    {
                        // 기존 데이터 저장
                        boolDetail = m_db.DeleteDRPIDiagnosisDetailDataGridViewDataInfo(strPlantName.Trim(), strHogi, strOhDegree, strDRPIGroup, strDRPIType
                            , strControlName, strCoilName, intCoilNumber);
                    }

                    // 데이터 저장
                    boolDetail = m_db.InsertDRPIDiagnosisDetailDataGridViewDataInfo(strPlantName.Trim(), strHogi, strOhDegree, strDRPIGroup, strDRPIType
                        , strControlName, strCoilName, intCoilNumber, strMeasurementMode, dcmDC_ResistanceValue, dcmDC_Deviation, dcmAC_ResistanceValue
                        , dcmAC_Deviation, dcmL_Inductace, dcmL_Deviation, dcmC_CapacitanceValue, dcmC_Deviation, dcmQ_FactorValue, dcmQ_Deviation, strResult
						, strMeasurementDate, intSeqNumber, intControlSeqNumber);

                    boolSave = boolDetail;
                }

                if (boolSave)
                {
                    frmMB.lblMessage.Text = "저장 완료";
                    frmMB.TopMost = true;
                    frmMB.ShowDialog();
                }
                else
                {
                    if (!boolHeader)
                    {
                        frmMB.lblMessage.Text = "Header Table 저장 중 실패";
                        frmMB.TopMost = true;
                        frmMB.ShowDialog();
                    }
                    else
                    {
                        frmMB.lblMessage.Text = "Detail Table 저장 중 실패";
                        frmMB.TopMost = true;
                        frmMB.ShowDialog();
                    }
                }
            }
            catch (Exception ex)
            {
                frmMB.lblMessage.Text = "저장 중 오류";
                frmMB.TopMost = true;
                frmMB.ShowDialog();
                System.Diagnostics.Debug.Print(ex.Message);
            }

            System.Windows.Forms.Cursor.Current = Cursors.Default;
        }

		/// <summary>
		/// 제어봉 순번 가져오기
		/// </summary>
		/// <param name="_strControlName"></param>
		/// <returns></returns>
		private int GetControlSeqNumber(ref string[] controlRod, ref string[] shutdownRod, string strDRPIType, string _strControlName)
		{
			int intControlSeqNumber = 0;
            if(strDRPIType.Trim() == "제어용")
            {
                for (; intControlSeqNumber < controlRod.Length; intControlSeqNumber++)
                {
                    if (_strControlName == controlRod[intControlSeqNumber])
                    {
                        return intControlSeqNumber + 1;
                    }
                }
            }

            else //if (strDRPIType.Trim() == "정지용")
            {
                for (; intControlSeqNumber < shutdownRod.Length; intControlSeqNumber++)
                {
                    if (_strControlName == shutdownRod[intControlSeqNumber])
                    {
                        return intControlSeqNumber + controlRod.Length + 1;
                    }
                }
            }
                
            return -1; // not found 
		}

        /// <summary>
        /// 정지 Button Evnet
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnMeasurementStop_Click(object sender, EventArgs e)
        {
            try
            {
                if (!boolMeasurementStart)
                {
                    frmMB.lblMessage.Text = "측정 중이 아닙니다.";
                    frmMB.TopMost = true;
                    frmMB.ShowDialog();
                    return;
                }

                boolMeasurementStop = true;

                threadMeasurementStop = new Thread(new ThreadStart(threadStopTesterMeasurment));
                threadMeasurementStop.Start();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
        }

        /// <summary>
        /// 측정 정지 
        /// </summary>
        private void threadStopTesterMeasurment()
        {
            try
            {
                while (boolMeasurementStart)
                {
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
            finally
            {
                boolMeasurementStop = false;

                if (threadMeasurementStop != null)
                    threadMeasurementStop.Abort();
            }
        }

        /// <summary>
        /// 측정 Button Evnet
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnMeasurementStart_Click(object sender, EventArgs e)
        {
            if (boolMeasurementStart) return;

            try
            {
                // DAQ 핀 초기화
                m_MeasureProcess.DigitalDAQ_CloseChannel();

                // 측정 대상 호기 체크
                if (cboHogi.Items.Count <= 0 || cboHogi.SelectedItem == null)
                {
                    frmMB.lblMessage.Text = "호기를 선택하십시오.";
                    frmMB.TopMost = true;
                    frmMB.ShowDialog();
                    return;
                }

                // 측정 대상 차수 체크
                if (teOhDegree2.Text.Trim() == "")
                {
                    frmMB.lblMessage.Text = "차수를 입력하십시오.";
                    frmMB.TopMost = true;
                    frmMB.ShowDialog();
                    return;
                }

                // 측정 대상 항목 체크
                if (dgvMeasurement.RowCount <= 0)
                {
                    frmMB.lblMessage.Text = "측정할 대상이 없습니다.";
                    frmMB.TopMost = true;
                    frmMB.ShowDialog();
                    return;
                }

                //측정 시작 중단에 따라 버튼 활성화/비활성화 설정
                SetMeasurementStartButtonEnabled(false);

                // 카드 선택 버튼 비활성화 설정
                btnSelectCard.Enabled = false;

                // 측정 시 Control 활성화/비활성화 설정 
                SetControlEnabled(false);

                // 그리드에 DAM와 Channel Column PinMap 설정
                SetDataGridViewCellsDMMChannel();

                // 측정 일시
                strMeasurementDate = System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                boolMeasurementStop = false;
                boolMeasurementStart = true;

                // 일반모드 선택된 OI Card 측정 Thread 시작 
                threadMeasurementStart = new Thread(new ThreadStart(threadStartTesterMeasurment));
                threadMeasurementStart.Start();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
        }

        /// <summary>
        /// 측정 시 Control 활성화/비활성화 설정 
        /// </summary>
        /// <param name="_boolEnabled"></param>
        private void SetControlEnabled(bool _boolEnabled)
        {
            rboCoilA.Enabled = _boolEnabled;
            rboCoilB.Enabled = _boolEnabled;

            cboHogi.Enabled = _boolEnabled;
            teOhDegree2.Enabled = _boolEnabled;
            //cboFrequency.Enabled = _boolEnabled;
            cboVoltageLevel.Enabled = _boolEnabled;
            cboMeasurementCount.Enabled = _boolEnabled;

            teTemperature_ReferenceValue.Enabled = _boolEnabled;
            teTemperatureUpDown_ReferenceValue.Enabled = _boolEnabled;

            rboNormalMode.Enabled = _boolEnabled;
            rboWheatstoneMode.Enabled = _boolEnabled;

            chkRdc.Enabled = _boolEnabled;
            chkRac.Enabled = _boolEnabled;
            chkL.Enabled = _boolEnabled;
            chkC.Enabled = _boolEnabled;
            chkQ.Enabled = _boolEnabled;
        }

        /// <summary>
        /// 그리드에 DAM와 Channel Column PinMap 설정
        /// </summary>
        private void SetDataGridViewCellsDMMChannel()
        {
            int intDAMCount = 0, intChannelCount = 0;
            string strControlName = "", strChannelNumber = "1", strChannelPinMap = "";

            for (int i = 0; i < dgvMeasurement.RowCount; i++)
            {
                strChannelNumber = dgvMeasurement.Rows[i].Cells["CoilName"].Value == null ? "1" : Regex.Replace(dgvMeasurement.Rows[i].Cells["CoilName"].Value.ToString().Trim(), @"\D", "");

                if (dgvMeasurement.Rows[i].Cells["DRPIType"].Value.ToString().Trim() == "정지용")
                {
                    switch (strChannelNumber.Trim())
                    {
                        case "1":
                            intChannelCount = 1;
                            break;
                        case "2":
                            intChannelCount = 2;
                            break;
                        case "3":
                            intChannelCount = 3;
                            break;
                        case "4":
                            intChannelCount = 7;
                            break;
                        case "5":
                            intChannelCount = 8;
                            break;
                        case "6":
                            intChannelCount = 9;
                            break;
                        case "18":
                            intChannelCount = 10;
                            break;
                        case "19":
                            intChannelCount = 4;
                            break;
                        case "20":
                            intChannelCount = 5;
                            break;
                        case "21":
                            intChannelCount = 6;
                            break;
                    }
                }
                else
                {
                    intChannelCount = strChannelNumber == "" ? 0 : Convert.ToInt32(strChannelNumber.Trim());
                }

                if (strControlName.Trim() != dgvMeasurement.Rows[i].Cells["ControlName"].Value.ToString().Trim())
                {
                    if (intChannelCount == 1)
                        intDAMCount++;

                    if (boolNormalMode)
                    {
            
                        dgvMeasurement.Rows[i].Cells["DAMName"].Value = Gini.GetValue("DRPI", $"DRPINormalMode_SelectDAMRelay{intDAMCount}").Trim();

                        #region Normal Mode PinMap 설정
                        if (intDAMCount == 1)
                        {
                            if (dgvMeasurement.Rows[i].Cells["DRPIType"].Value.ToString().Trim() == "제어용")
                                strChannelPinMap = arrayCard01_NormalModePinMap[intChannelCount - 1].Trim();
                            else
                                strChannelPinMap = arrayCard01_NormalModePinMap[intChannelCount - 1].Trim();
                        }
                        else if (intDAMCount == 2)
                        {
                            if (dgvMeasurement.Rows[i].Cells["DRPIType"].Value.ToString().Trim() == "제어용")
                                strChannelPinMap = arrayCard02_NormalModePinMap[intChannelCount - 1].Trim();
                            else
                                strChannelPinMap = arrayCard02_NormalModePinMap[intChannelCount - 1].Trim();
                        }
                        else if (intDAMCount == 3)
                        {
                            if (dgvMeasurement.Rows[i].Cells["DRPIType"].Value.ToString().Trim() == "제어용")
                                strChannelPinMap = arrayCard03_NormalModePinMap[intChannelCount - 1].Trim();
                            else
                                strChannelPinMap = arrayCard03_NormalModePinMap[intChannelCount - 1].Trim();
                        }
                        else if (intDAMCount == 4)
                        {
                            if (dgvMeasurement.Rows[i].Cells["DRPIType"].Value.ToString().Trim() == "제어용")
                                strChannelPinMap = arrayCard04_NormalModePinMap[intChannelCount - 1].Trim();
                            else
                                strChannelPinMap = arrayCard04_NormalModePinMap[intChannelCount - 1].Trim();
                        }
                        else if (intDAMCount == 5)
                        {
                            if (dgvMeasurement.Rows[i].Cells["DRPIType"].Value.ToString().Trim() == "제어용")
                                strChannelPinMap = arrayCard05_NormalModePinMap[intChannelCount - 1].Trim();
                            else
                                strChannelPinMap = arrayCard05_NormalModePinMap[intChannelCount - 1].Trim();
                        }
                        else if (intDAMCount == 6)
                        {
                            if (dgvMeasurement.Rows[i].Cells["DRPIType"].Value.ToString().Trim() == "제어용")
                                strChannelPinMap = arrayCard06_NormalModePinMap[intChannelCount - 1].Trim();
                            else
                                strChannelPinMap = arrayCard06_NormalModePinMap[intChannelCount - 1].Trim();
                        }
                        else if (intDAMCount == 7)
                        {
                            if (dgvMeasurement.Rows[i].Cells["DRPIType"].Value.ToString().Trim() == "제어용")
                                strChannelPinMap = arrayCard07_NormalModePinMap[intChannelCount - 1].Trim();
                            else
                                strChannelPinMap = arrayCard07_NormalModePinMap[intChannelCount - 1].Trim();
                        }
                        else if (intDAMCount == 8)
                        {
                            if (dgvMeasurement.Rows[i].Cells["DRPIType"].Value.ToString().Trim() == "제어용")
                                strChannelPinMap = arrayCard08_NormalModePinMap[intChannelCount - 1].Trim();
                            else
                                strChannelPinMap = arrayCard08_NormalModePinMap[intChannelCount - 1].Trim();
                        }
                        else if (intDAMCount == 9)
                        {
                            if (dgvMeasurement.Rows[i].Cells["DRPIType"].Value.ToString().Trim() == "제어용")
                                strChannelPinMap = arrayCard09_NormalModePinMap[intChannelCount - 1].Trim();
                            else
                                strChannelPinMap = arrayCard09_NormalModePinMap[intChannelCount - 1].Trim();
                        }
                        else if (intDAMCount == 10)
                        {
                            if (dgvMeasurement.Rows[i].Cells["DRPIType"].Value.ToString().Trim() == "제어용")
                                strChannelPinMap = arrayCard10_NormalModePinMap[intChannelCount - 1].Trim();
                            else
                                strChannelPinMap = arrayCard10_NormalModePinMap[intChannelCount - 1].Trim();
                        }
                        #endregion
                    }
                    else
                    {
                        dgvMeasurement.Rows[i].Cells["DAMName"].Value = Gini.GetValue("DRPI", $"DRPIWheatstoneMode_SelectDAMRelay{intDAMCount}").Trim();

                        #region Wheatstone Mode PinMap 설정
                        if (intDAMCount == 1)
                        {
                            if (dgvMeasurement.Rows[i].Cells["DRPIType"].Value.ToString().Trim() == "제어용")
                                strChannelPinMap = arrayCard01_WheatstoneModePinMap[intChannelCount - 1].Trim();
                            else
                                strChannelPinMap = arrayCard01_WheatstoneModePinMap[intChannelCount - 1].Trim();
                        }
                        else if (intDAMCount == 2)
                        {
                            if (dgvMeasurement.Rows[i].Cells["DRPIType"].Value.ToString().Trim() == "제어용")
                                strChannelPinMap = arrayCard02_WheatstoneModePinMap[intChannelCount - 1].Trim();
                            else
                                strChannelPinMap = arrayCard02_WheatstoneModePinMap[intChannelCount - 1].Trim();
                        }
                        else if (intDAMCount == 3)
                        {
                            if (dgvMeasurement.Rows[i].Cells["DRPIType"].Value.ToString().Trim() == "제어용")
                                strChannelPinMap = arrayCard03_WheatstoneModePinMap[intChannelCount - 1].Trim();
                            else
                                strChannelPinMap = arrayCard03_WheatstoneModePinMap[intChannelCount - 1].Trim();
                        }
                        else if (intDAMCount == 4)
                        {
                            if (dgvMeasurement.Rows[i].Cells["DRPIType"].Value.ToString().Trim() == "제어용")
                                strChannelPinMap = arrayCard04_WheatstoneModePinMap[intChannelCount - 1].Trim();
                            else
                                strChannelPinMap = arrayCard04_WheatstoneModePinMap[intChannelCount - 1].Trim();
                        }
                        else if (intDAMCount == 5)
                        {
                            if (dgvMeasurement.Rows[i].Cells["DRPIType"].Value.ToString().Trim() == "제어용")
                                strChannelPinMap = arrayCard05_WheatstoneModePinMap[intChannelCount - 1].Trim();
                            else
                                strChannelPinMap = arrayCard05_WheatstoneModePinMap[intChannelCount - 1].Trim();
                        }
                        else if (intDAMCount == 6)
                        {
                            if (dgvMeasurement.Rows[i].Cells["DRPIType"].Value.ToString().Trim() == "제어용")
                                strChannelPinMap = arrayCard06_WheatstoneModePinMap[intChannelCount - 1].Trim();
                            else
                                strChannelPinMap = arrayCard06_WheatstoneModePinMap[intChannelCount - 1].Trim();
                        }
                        else if (intDAMCount == 7)
                        {
                            if (dgvMeasurement.Rows[i].Cells["DRPIType"].Value.ToString().Trim() == "제어용")
                                strChannelPinMap = arrayCard07_WheatstoneModePinMap[intChannelCount - 1].Trim();
                            else
                                strChannelPinMap = arrayCard07_WheatstoneModePinMap[intChannelCount - 1].Trim();
                        }
                        else if (intDAMCount == 8)
                        {
                            if (dgvMeasurement.Rows[i].Cells["DRPIType"].Value.ToString().Trim() == "제어용")
                                strChannelPinMap = arrayCard08_WheatstoneModePinMap[intChannelCount - 1].Trim();
                            else
                                strChannelPinMap = arrayCard08_WheatstoneModePinMap[intChannelCount - 1].Trim();
                        }
                        else if (intDAMCount == 9)
                        {
                            if (dgvMeasurement.Rows[i].Cells["DRPIType"].Value.ToString().Trim() == "제어용")
                                strChannelPinMap = arrayCard09_WheatstoneModePinMap[intChannelCount - 1].Trim();
                            else
                                strChannelPinMap = arrayCard09_WheatstoneModePinMap[intChannelCount - 1].Trim();
                        }
                        else if (intDAMCount == 10)
                        {
                            if (dgvMeasurement.Rows[i].Cells["DRPIType"].Value.ToString().Trim() == "제어용")
                                strChannelPinMap = arrayCard10_WheatstoneModePinMap[intChannelCount - 1].Trim();
                            else
                                strChannelPinMap = arrayCard10_WheatstoneModePinMap[intChannelCount - 1].Trim();
                        }
                        #endregion
                    }

                    strControlName = dgvMeasurement.Rows[i].Cells["ControlName"].Value.ToString().Trim();
                }
                else
                {
                    if (boolNormalMode)
                    {
                        dgvMeasurement.Rows[i].Cells["DAMName"].Value = Gini.GetValue("DRPI", $"DRPINormalMode_SelectDAMRelay{intDAMCount}").Trim();

                        #region Normal Mode PinMap 설정
                        if (intDAMCount == 1)
                        {
                            if (dgvMeasurement.Rows[i].Cells["DRPIType"].Value.ToString().Trim() == "제어용")
                                strChannelPinMap = arrayCard01_NormalModePinMap[intChannelCount - 1].Trim();
                            else
                                strChannelPinMap = arrayCard01_NormalModePinMap[intChannelCount - 1].Trim();
                        }
                        else if (intDAMCount == 2)
                        {
                            if (dgvMeasurement.Rows[i].Cells["DRPIType"].Value.ToString().Trim() == "제어용")
                                strChannelPinMap = arrayCard02_NormalModePinMap[intChannelCount - 1].Trim();
                            else
                                strChannelPinMap = arrayCard02_NormalModePinMap[intChannelCount - 1].Trim();
                        }
                        else if (intDAMCount == 3)
                        {
                            if (dgvMeasurement.Rows[i].Cells["DRPIType"].Value.ToString().Trim() == "제어용")
                                strChannelPinMap = arrayCard03_NormalModePinMap[intChannelCount - 1].Trim();
                            else
                                strChannelPinMap = arrayCard03_NormalModePinMap[intChannelCount - 1].Trim();
                        }
                        else if (intDAMCount == 4)
                        {
                            if (dgvMeasurement.Rows[i].Cells["DRPIType"].Value.ToString().Trim() == "제어용")
                                strChannelPinMap = arrayCard04_NormalModePinMap[intChannelCount - 1].Trim();
                            else
                                strChannelPinMap = arrayCard04_NormalModePinMap[intChannelCount - 1].Trim();
                        }
                        else if (intDAMCount == 5)
                        {
                            if (dgvMeasurement.Rows[i].Cells["DRPIType"].Value.ToString().Trim() == "제어용")
                                strChannelPinMap = arrayCard05_NormalModePinMap[intChannelCount - 1].Trim();
                            else
                                strChannelPinMap = arrayCard05_NormalModePinMap[intChannelCount - 1].Trim();
                        }
                        else if (intDAMCount == 6)
                        {
                            if (dgvMeasurement.Rows[i].Cells["DRPIType"].Value.ToString().Trim() == "제어용")
                                strChannelPinMap = arrayCard06_NormalModePinMap[intChannelCount - 1].Trim();
                            else if (dgvMeasurement.Rows[i].Cells["DRPIType"].Value.ToString().Trim() == "정지용" && intChannelCount == 18)
                                strChannelPinMap = arrayCard06_NormalModePinMap[6].Trim();
                            else if (dgvMeasurement.Rows[i].Cells["DRPIType"].Value.ToString().Trim() == "정지용" && intChannelCount == 19)
                                strChannelPinMap = arrayCard06_NormalModePinMap[7].Trim();
                            else if (dgvMeasurement.Rows[i].Cells["DRPIType"].Value.ToString().Trim() == "정지용" && intChannelCount == 20)
                                strChannelPinMap = arrayCard06_NormalModePinMap[8].Trim();
                            else if (dgvMeasurement.Rows[i].Cells["DRPIType"].Value.ToString().Trim() == "정지용" && intChannelCount == 21)
                                strChannelPinMap = arrayCard06_NormalModePinMap[9].Trim();
                            else
                                strChannelPinMap = arrayCard06_NormalModePinMap[intChannelCount - 1].Trim();
                        }
                        else if (intDAMCount == 7)
                        {
                            if (dgvMeasurement.Rows[i].Cells["DRPIType"].Value.ToString().Trim() == "제어용")
                                strChannelPinMap = arrayCard07_NormalModePinMap[intChannelCount - 1].Trim();
                            else
                                strChannelPinMap = arrayCard07_NormalModePinMap[intChannelCount - 1].Trim();
                        }
                        else if (intDAMCount == 8)
                        {
                            if (dgvMeasurement.Rows[i].Cells["DRPIType"].Value.ToString().Trim() == "제어용")
                                strChannelPinMap = arrayCard08_NormalModePinMap[intChannelCount - 1].Trim();
                            else
                                strChannelPinMap = arrayCard08_NormalModePinMap[intChannelCount - 1].Trim();
                        }
                        else if (intDAMCount == 9)
                        {
                            if (dgvMeasurement.Rows[i].Cells["DRPIType"].Value.ToString().Trim() == "제어용")
                                strChannelPinMap = arrayCard09_NormalModePinMap[intChannelCount - 1].Trim();
                            else
                                strChannelPinMap = arrayCard09_NormalModePinMap[intChannelCount - 1].Trim();
                        }
                        else if (intDAMCount == 10)
                        {
                            if (dgvMeasurement.Rows[i].Cells["DRPIType"].Value.ToString().Trim() == "제어용")
                                strChannelPinMap = arrayCard10_NormalModePinMap[intChannelCount - 1].Trim();
                            else
                                strChannelPinMap = arrayCard10_NormalModePinMap[intChannelCount - 1].Trim();
                        }
                        #endregion
                    }
                    else
                    {
                        dgvMeasurement.Rows[i].Cells["DAMName"].Value = Gini.GetValue("DRPI", $"DRPIWheatstoneMode_SelectDAMRelay{intDAMCount}").Trim();

                        #region Wheatstone Mode PinMap 설정
                        if (intDAMCount == 1)
                        {
                            if (dgvMeasurement.Rows[i].Cells["DRPIType"].Value.ToString().Trim() == "제어용")
                                strChannelPinMap = arrayCard01_WheatstoneModePinMap[intChannelCount - 1].Trim();
                            else
                                strChannelPinMap = arrayCard01_WheatstoneModePinMap[intChannelCount - 1].Trim();
                        }
                        else if (intDAMCount == 2)
                        {
                            if (dgvMeasurement.Rows[i].Cells["DRPIType"].Value.ToString().Trim() == "제어용")
                                strChannelPinMap = arrayCard02_WheatstoneModePinMap[intChannelCount - 1].Trim();
                            else
                                strChannelPinMap = arrayCard02_WheatstoneModePinMap[intChannelCount - 1].Trim();
                        }
                        else if (intDAMCount == 3)
                        {
                            if (dgvMeasurement.Rows[i].Cells["DRPIType"].Value.ToString().Trim() == "제어용")
                                strChannelPinMap = arrayCard03_WheatstoneModePinMap[intChannelCount - 1].Trim();
                            else
                                strChannelPinMap = arrayCard03_WheatstoneModePinMap[intChannelCount - 1].Trim();
                        }
                        else if (intDAMCount == 4)
                        {
                            if (dgvMeasurement.Rows[i].Cells["DRPIType"].Value.ToString().Trim() == "제어용")
                                strChannelPinMap = arrayCard04_WheatstoneModePinMap[intChannelCount - 1].Trim();
                            else
                                strChannelPinMap = arrayCard04_WheatstoneModePinMap[intChannelCount - 1].Trim();
                        }
                        else if (intDAMCount == 5)
                        {
                            if (dgvMeasurement.Rows[i].Cells["DRPIType"].Value.ToString().Trim() == "제어용")
                                strChannelPinMap = arrayCard05_WheatstoneModePinMap[intChannelCount - 1].Trim();
                            else
                                strChannelPinMap = arrayCard05_WheatstoneModePinMap[intChannelCount - 1].Trim();
                        }
                        else if (intDAMCount == 6)
                        {
                            if (dgvMeasurement.Rows[i].Cells["DRPIType"].Value.ToString().Trim() == "제어용")
                                strChannelPinMap = arrayCard06_WheatstoneModePinMap[intChannelCount - 1].Trim();
                            else
                                strChannelPinMap = arrayCard06_WheatstoneModePinMap[intChannelCount - 1].Trim();
                        }
                        else if (intDAMCount == 7)
                        {
                            if (dgvMeasurement.Rows[i].Cells["DRPIType"].Value.ToString().Trim() == "제어용")
                                strChannelPinMap = arrayCard07_WheatstoneModePinMap[intChannelCount - 1].Trim();
                            else
                                strChannelPinMap = arrayCard07_WheatstoneModePinMap[intChannelCount - 1].Trim();
                        }
                        else if (intDAMCount == 8)
                        {
                            if (dgvMeasurement.Rows[i].Cells["DRPIType"].Value.ToString().Trim() == "제어용")
                                strChannelPinMap = arrayCard08_WheatstoneModePinMap[intChannelCount - 1].Trim();
                            else
                                strChannelPinMap = arrayCard08_WheatstoneModePinMap[intChannelCount - 1].Trim();
                        }
                        else if (intDAMCount == 9)
                        {
                            if (dgvMeasurement.Rows[i].Cells["DRPIType"].Value.ToString().Trim() == "제어용")
                                strChannelPinMap = arrayCard09_WheatstoneModePinMap[intChannelCount - 1].Trim();
                            else
                                strChannelPinMap = arrayCard09_WheatstoneModePinMap[intChannelCount - 1].Trim();
                        }
                        else if (intDAMCount == 10)
                        {
                            if (dgvMeasurement.Rows[i].Cells["DRPIType"].Value.ToString().Trim() == "제어용")
                                strChannelPinMap = arrayCard10_WheatstoneModePinMap[intChannelCount - 1].Trim();
                            else
                                strChannelPinMap = arrayCard10_WheatstoneModePinMap[intChannelCount - 1].Trim();
                        }
                        #endregion
                    }
                }

                dgvMeasurement.Rows[i].Cells["Channel"].Value = string.Format("Ch{0}", intChannelCount.ToString().Trim());
                dgvMeasurement.Rows[i].Cells["DAQPinMap"].Value = strChannelPinMap;
            }
        }

        /// <summary>
        /// 일반/휘스톤 모드 선택된 OI Card 측정 Thread 시작 
        /// </summary>
        public void threadStartTesterMeasurment()
        {
            bool boolResult = false;

            strDAQDeviceName = Gini.GetValue("Device", "DAQDeviceName").Trim();
            m_MeasureProcess.DigitalDAQ_CloseChannel();

            try
            {
                string strMeasurementMode = rboNormalMode.Checked ? "일반모드" : "휘스톤모드";

                decimal dcmMeasurementValue = 0M, dcmRefValue = 0M;
                string strControlName = "", strCoilName = "", strNormalModeDAQPinMap = "", strWheatstoneModeDAQPinMap = "", strDAQPinMap = ""
                    , strDAMRelayName = "", strChannelName = "", strRdcResult = "", strRacResult = "", strLResult = "", strCResult = "", strQResult = ""; ;
                int intCoilNumber = 0, intSleep = 300;

                decimal dcmDecisionRange_Rdc = 0M, dcmDecisionRange_Rac = 0M, dcmDecisionRange_L = 0M, dcmDecisionRange_C = 0M
                    , dcmDecisionRange_Q = 0M, dcmEffectiveStandardRange = 0M;

                dcmDecisionRange_Rdc = Convert.ToDecimal(Gini.GetValue("ReferenceValue", "RdcDecisionRange_ReferenceValue"));
                dcmDecisionRange_Rac = Convert.ToDecimal(Gini.GetValue("ReferenceValue", "RacDecisionRange_ReferenceValue"));
                dcmDecisionRange_L = Convert.ToDecimal(Gini.GetValue("ReferenceValue", "LDecisionRange_ReferenceValue"));
                dcmDecisionRange_C = Convert.ToDecimal(Gini.GetValue("ReferenceValue", "CDecisionRange_ReferenceValue"));
                dcmDecisionRange_Q = Convert.ToDecimal(Gini.GetValue("ReferenceValue", "QDecisionRange_ReferenceValue"));
                dcmEffectiveStandardRange = Convert.ToDecimal(Gini.GetValue("ReferenceValue", "EffectiveStandardRangeOfVariation"));

                decimal dcmDC_Resistance = 0M, dcmAC_Resistance = 0M, dcmL_Resistance = 0M, dcmC_Resistance = 0M
                    , dcmAC_MeasurementValue = 0M, dcmL_MeasurementValue = 0M;

                // LCR-Meter Range 설정
                m_MeasureProcess.functionLCRMeterRangeSetting(Function.FunctionLCRMeterinfo.intLCRMeer_RangeValue);

                // LCR-Meter 주파수 설정
                m_MeasureProcess.functionLCRMeterFrequencySetting(dcmFrequency);

                // LCR-Meter VoltageLevel 설정
                m_MeasureProcess.functionLCRMeterVoltageLevelSetting(dcmVoltageLevel);

                // LCR-Meter Mode 설정
                m_MeasureProcess.functionLCRMeterModeSetting(2);

                // 그리드의 행별 측정 진행
                for (int i = 0; i < dgvMeasurement.RowCount; i++)
                {
                    // DataGridView Select 컬럼의 체크박스 상태에 따라 측정 진행
                    if (!(bool)dgvMeasurement.Rows[i].Cells["Select"].Value) continue;

                    strControlName = dgvMeasurement.Rows[i].Cells["ControlName"].Value.ToString().Trim();
                    strCoilName = dgvMeasurement.Rows[i].Cells["CoilName"].Value.ToString().Trim();
                    intCoilNumber = dgvMeasurement.Rows[i].Cells["CoilNumber"].Value == null || dgvMeasurement.Rows[i].Cells["CoilNumber"].Value.ToString().Trim() == ""
                        ? 0 : Convert.ToInt32(dgvMeasurement.Rows[i].Cells["CoilNumber"].Value.ToString().Trim());
                    strDAMRelayName = dgvMeasurement.Rows[i].Cells["DAMName"].Value == null ? "" : dgvMeasurement.Rows[i].Cells["DAMName"].Value.ToString().Trim();
                    strChannelName = dgvMeasurement.Rows[i].Cells["Channel"].Value == null ? "" : dgvMeasurement.Rows[i].Cells["Channel"].Value.ToString().Trim();
                    dcmAC_MeasurementValue = 0M;
                    dcmL_MeasurementValue = 0M;
                    dgvMeasurement.Rows[i].Cells["Result"].Value = "";

                    if (strControlName.Trim() == "" || strCoilName.Trim() == "") continue;

                    // Channel 별로 DAQ Pin 번호 가져오기
                    strDAQPinMap = dgvMeasurement.Rows[i].Cells["DAQPinMap"].Value == null ? "" : dgvMeasurement.Rows[i].Cells["DAQPinMap"].Value.ToString().Trim();

                    if (strMeasurementMode.Trim() == "일반모드")
                    {
                        strWheatstoneModeDAQPinMap = strDAQPinMap;
                        strNormalModeDAQPinMap = strDAQPinMap;
                    }
                    else
                    {
                        strWheatstoneModeDAQPinMap = strDAQPinMap;
                        string[] arrayDAQPinMap = strDAQPinMap.Split(',');

                        strNormalModeDAQPinMap = arrayDAQPinMap[0] + "," + arrayDAQPinMap[1] + "," + arrayDAQPinMap[2] + "," + arrayDAQPinMap[3];
                    }

                    // 영점보정 값 가져오기
                    DataTable dt = new DataTable();
                    dt = m_db.GetSetOffsetDataGridViewDataInfo(strPlantName, strMeasurementMode, strDAMRelayName, strChannelName);

                    if (dt != null && dt.Rows.Count > 0)
                    {
                        dcmDC_Resistance = dt.Rows[0]["DC_Resistance"] == null || dt.Rows[0]["DC_Resistance"].ToString().Trim() == "" ? 0.000M : Convert.ToDecimal(dt.Rows[0]["DC_Resistance"].ToString().Trim());
                        dcmAC_Resistance = dt.Rows[0]["AC_Resistance"] == null || dt.Rows[0]["AC_Resistance"].ToString().Trim() == "" ? 0.000M : Convert.ToDecimal(dt.Rows[0]["AC_Resistance"].ToString().Trim());
                        dcmL_Resistance = dt.Rows[0]["Inductance"] == null || dt.Rows[0]["Inductance"].ToString().Trim() == "" ? 0.000M : Convert.ToDecimal(dt.Rows[0]["Inductance"].ToString().Trim());
                        dcmC_Resistance = dt.Rows[0]["Capacitance"] == null || dt.Rows[0]["Capacitance"].ToString().Trim() == "" ? 0.000M : Convert.ToDecimal(dt.Rows[0]["Capacitance"].ToString().Trim());
                    }
                    else
                    {
                        dcmDC_Resistance = 0.000M;
                        dcmAC_Resistance = 0.000M;
                        dcmL_Resistance = 0.000M;
                        dcmC_Resistance = 0.000000M;
                    }

                    // DAQ 핀 On
                    boolResult = m_MeasureProcess.functionDAQPinMapOnOff(strWheatstoneModeDAQPinMap, true);
                    Thread.Sleep(50);
                    
                    // 측정이 정지되었을 경우
                    if (boolMeasurementStop || !boolMeasurementStart) break;

                    // DC 측정
                    if (chkRdc.Checked)
                    {
                        #region DC 측정

                        // 컬럼 선택
                        dgvMeasurement.Rows[i].Cells["DC_ResistanceValue"].Selected = true;

                        // DC 데이터 측정값 가져오기
                        boolResult = m_MeasureProcess.functionRCSRdcMeasurement(ref dcmMeasurementValue, rboWheatstoneMode.Checked, intSleep);
                        intSleep = 100;

                        // 영점 보정 
                        dcmMeasurementValue = dcmMeasurementValue - dcmDC_Resistance;

                        // 기준값과 비교 및 편차 산출
                        if (rboCoilA.Checked)
                        {
                            if (strCoilName.Trim() == "A1")
                                dcmRefValue = dA_ControlRodRdc[0] == null ? 0.000M : dA_ControlRodRdc[0];
                            else if (strCoilName.Trim() == "A2")
                                dcmRefValue = dA_ControlRodRdc[1] == null ? 0.000M : dA_ControlRodRdc[1];
                            else
                                dcmRefValue = dA_ControlRodRdc[2] == null ? 0.000M : dA_ControlRodRdc[2];
                        }
                        else
                        {
                            if (strCoilName.Trim() == "B1")
                                dcmRefValue = dB_ControlRodRdc[0] == null ? 0.000M : dB_ControlRodRdc[0];
                            else if (strCoilName.Trim() == "B2")
                                dcmRefValue = dB_ControlRodRdc[1] == null ? 0.000M : dB_ControlRodRdc[1];
                            else
                                dcmRefValue = dB_ControlRodRdc[2] == null ? 0.000M : dB_ControlRodRdc[2];
                        }

                        if (rboWheatstoneMode.Checked)
                        {
                            // 선간 보정 
                            decimal dcmWheatstoneMode_Compensation = Convert.ToDecimal(Gini.GetValue("Wheatstone", "WheatstoneMode_Compensation"));
                            dcmMeasurementValue = dcmMeasurementValue + (dcmWheatstoneMode_Compensation * System.Math.Truncate(dcmMeasurementValue));
                        }

                        // 온도와 온도 증감 값 적용
                        if (dcmTemperature_ReferenceValue > 0)
                        {
                            if (dcmTemperature_ReferenceValue < dcmTemperature_ChangeValue)
                                dcmMeasurementValue = dcmMeasurementValue + dcmTemperature_Measurement;
                            else if (dcmTemperature_ReferenceValue > dcmTemperature_ChangeValue)
                                dcmMeasurementValue = dcmMeasurementValue - dcmTemperature_Measurement;
                        }

                        // 그리드에 측정 값 삽입
                        dgvMeasurement.Rows[i].Cells["DC_ResistanceValue"].Value = dcmMeasurementValue.ToString("F3");
                        dgvMeasurement.Rows[i].Cells["DC_Deviation"].Value = (dcmMeasurementValue - dcmRefValue).ToString("F3");

                        // 결과 체크
                        strRdcResult = m_MeasureProcess.GetMesurmentValueDecision(dcmMeasurementValue, dcmRefValue, dcmDecisionRange_Rdc, dcmEffectiveStandardRange, "DC");

                        if ((dcmMeasurementValue - dcmRefValue) == 0)
                            dgvMeasurement.Rows[i].Cells["DC_ResistanceValue"].Style.ForeColor = System.Drawing.Color.Black;
                        else
                        {
                            if (strRdcResult.Trim() == "부적합")
                            {
                                dgvMeasurement.Rows[i].Cells["DC_ResistanceValue"].Style.ForeColor = System.Drawing.Color.Red;
                            }
                            else if (strRdcResult.Trim() == "의심")
                            {
                                dgvMeasurement.Rows[i].Cells["DC_ResistanceValue"].Style.ForeColor = System.Drawing.Color.Blue;
                            }
                            else if (strRdcResult.Trim() == "적합")
                            {
                                dgvMeasurement.Rows[i].Cells["DC_ResistanceValue"].Style.ForeColor = System.Drawing.Color.Black;
                            }
                        }

                        // 컬럼 선택 취소
                        dgvMeasurement.Rows[i].Cells["DC_ResistanceValue"].Selected = false;
                        #endregion
                    }

                    // DAQ 핀 Off
                    boolResult = m_MeasureProcess.functionDAQPinMapOnOff(strWheatstoneModeDAQPinMap, false);
                    Thread.Sleep(50);

                    // 측정이 정지되었을 경우
                    if (boolMeasurementStop || !boolMeasurementStart) break;

                    // DAQ 핀 On
                    boolResult = m_MeasureProcess.functionDAQPinMapOnOff(strNormalModeDAQPinMap, true);
                    Thread.Sleep(50);

                    // AC 측정
                    if (chkRac.Checked)
                    {
                        #region AC 측정
                        // 컬럼 선택
                        dgvMeasurement.Rows[i].Cells["AC_ResistanceValue"].Selected = true;

                        // AC 데이터 측정값 가져오기
                        boolResult = m_MeasureProcess.functionRCSRacMeasurement(ref dcmMeasurementValue);
                        dcmAC_MeasurementValue = dcmMeasurementValue;
                        
                        // 영점 보정 
                        dcmMeasurementValue = dcmMeasurementValue - dcmAC_Resistance;
                        //dcmAC_MeasurementValue = dcmMeasurementValue;

                        // 기준값과 비교 및 편차 산출
                        if (rboCoilA.Checked)
                        {
                            if (strCoilName.Trim() == "A1")
                                dcmRefValue = dA_ControlRodRac[0] == null ? 0.000M : dA_ControlRodRac[0];
                            else if (strCoilName.Trim() == "A2")
                                dcmRefValue = dA_ControlRodRac[1] == null ? 0.000M : dA_ControlRodRac[1];
                            else
                                dcmRefValue = dA_ControlRodRac[2] == null ? 0.000M : dA_ControlRodRac[2];
                        }
                        else
                        {
                            if (strCoilName.Trim() == "B1")
                                dcmRefValue = dB_ControlRodRac[0] == null ? 0.000M : dB_ControlRodRac[0];
                            else if (strCoilName.Trim() == "B2")
                                dcmRefValue = dB_ControlRodRac[1] == null ? 0.000M : dB_ControlRodRac[1];
                            else
                                dcmRefValue = dB_ControlRodRac[2] == null ? 0.000M : dB_ControlRodRac[2];
                        }

                        // 온도와 온도 증감 값 적용
                        if (dcmTemperature_ReferenceValue > 0)
                        {
                            if (dcmTemperature_ReferenceValue < dcmTemperature_ChangeValue)
                                dcmMeasurementValue = dcmMeasurementValue + dcmTemperature_Measurement;
                            else if (dcmTemperature_ReferenceValue > dcmTemperature_ChangeValue)
                                dcmMeasurementValue = dcmMeasurementValue - dcmTemperature_Measurement;
                        }

                        // 그리드에 측정 값 삽입
                        dgvMeasurement.Rows[i].Cells["AC_ResistanceValue"].Value = dcmMeasurementValue.ToString("F3");
                        dgvMeasurement.Rows[i].Cells["AC_Deviation"].Value = (dcmMeasurementValue - dcmRefValue).ToString("F3");

                        // 결과 체크
                        strRacResult = m_MeasureProcess.GetMesurmentValueDecision(dcmMeasurementValue, dcmRefValue, dcmDecisionRange_Rac, dcmEffectiveStandardRange, "AC");

                        if ((dcmMeasurementValue - dcmRefValue) == 0)
                            dgvMeasurement.Rows[i].Cells["AC_ResistanceValue"].Style.ForeColor = System.Drawing.Color.Black;
                        else
                        {
                            if (strRacResult.Trim() == "부적합")
                            {
                                dgvMeasurement.Rows[i].Cells["AC_ResistanceValue"].Style.ForeColor = System.Drawing.Color.Red;
                            }
                            else if (strRacResult.Trim() == "의심")
                            {
                                dgvMeasurement.Rows[i].Cells["AC_ResistanceValue"].Style.ForeColor = System.Drawing.Color.Blue;
                            }
                            else if (strRacResult.Trim() == "적합")
                            {
                                dgvMeasurement.Rows[i].Cells["AC_ResistanceValue"].Style.ForeColor = System.Drawing.Color.Black;
                            }
                        }

                        // 컬럼 선택 취소
                        dgvMeasurement.Rows[i].Cells["AC_ResistanceValue"].Selected = false;
                        #endregion
                    }

                    // 측정이 정지되었을 경우
                    if (boolMeasurementStop || !boolMeasurementStart) break;

                    // L 측정
                    if (chkL.Checked)
                    {
                        #region L 측정
                        // 컬럼 선택
                        dgvMeasurement.Rows[i].Cells["L_InductanceValue"].Selected = true;

                        // L 데이터 측정값 가져오기
                        boolResult = m_MeasureProcess.functionRCSLMeasurement(ref dcmMeasurementValue);

                        // 영점 보정 
                        dcmMeasurementValue = dcmMeasurementValue - dcmL_Resistance;

                        // 기준값과 비교 및 편차 산출
                        if (rboCoilA.Checked)
                        {
                            if (strCoilName.Trim() == "A1")
                                dcmRefValue = dA_ControlRodL[0] == null ? 0.000M : dA_ControlRodL[0];
                            else if (strCoilName.Trim() == "A2")
                                dcmRefValue = dA_ControlRodL[1] == null ? 0.000M : dA_ControlRodL[1];
                            else
                                dcmRefValue = dA_ControlRodL[2] == null ? 0.000M : dA_ControlRodL[2];
                        }
                        else
                        {
                            if (strCoilName.Trim() == "B1")
                                dcmRefValue = dB_ControlRodL[0] == null ? 0.000M : dB_ControlRodL[0];
                            else if (strCoilName.Trim() == "B2")
                                dcmRefValue = dB_ControlRodL[1] == null ? 0.000M : dB_ControlRodL[1];
                            else
                                dcmRefValue = dB_ControlRodL[2] == null ? 0.000M : dB_ControlRodL[2];
                        }

                        // 그리드에 측정 값 삽입
                        dgvMeasurement.Rows[i].Cells["L_InductanceValue"].Value = dcmMeasurementValue.ToString("F3");
                        dgvMeasurement.Rows[i].Cells["L_Deviation"].Value = (dcmMeasurementValue - dcmRefValue).ToString("F3");
                        dcmL_MeasurementValue = dcmMeasurementValue;

                        // 결과 체크
                        strLResult = m_MeasureProcess.GetMesurmentValueDecision(dcmMeasurementValue, dcmRefValue, dcmDecisionRange_L, dcmEffectiveStandardRange, "L");

                        if ((dcmMeasurementValue - dcmRefValue) == 0)
                            dgvMeasurement.Rows[i].Cells["L_InductanceValue"].Style.ForeColor = System.Drawing.Color.Black;
                        else
                        {
                            if (strLResult.Trim() == "부적합")
                            {
                                dgvMeasurement.Rows[i].Cells["L_InductanceValue"].Style.ForeColor = System.Drawing.Color.Red;
                            }
                            else if (strLResult.Trim() == "의심")
                            {
                                dgvMeasurement.Rows[i].Cells["L_InductanceValue"].Style.ForeColor = System.Drawing.Color.Blue;
                            }
                            else if (strLResult.Trim() == "적합")
                            {
                                dgvMeasurement.Rows[i].Cells["L_InductanceValue"].Style.ForeColor = System.Drawing.Color.Black;
                            }
                        }

                        // 컬럼 선택 취소
                        dgvMeasurement.Rows[i].Cells["L_InductanceValue"].Selected = false;
                        #endregion
                    }

                    // 측정이 정지되었을 경우
                    if (boolMeasurementStop || !boolMeasurementStart) break;

                    // C 측정
                    if (chkC.Checked)
                    {
                        #region C 측정
                        // 컬럼 선택
                        dgvMeasurement.Rows[i].Cells["C_CapacitanceValue"].Selected = true;

                        // C 데이터 측정값 가져오기
                        boolResult = m_MeasureProcess.functionRCSCMeasurement(ref dcmMeasurementValue);

                        // 영점 보정 
                        dcmMeasurementValue = dcmMeasurementValue - dcmC_Resistance;

                        // 기준값과 비교 및 편차 산출
                        if (rboCoilA.Checked)
                        {
                            if (strCoilName.Trim() == "A1")
                                dcmRefValue = dA_ControlRodC[0] == null ? 0.000000M : dA_ControlRodC[0];
                            else if (strCoilName.Trim() == "A2")
                                dcmRefValue = dA_ControlRodC[1] == null ? 0.000000M : dA_ControlRodC[1];
                            else
                                dcmRefValue = dA_ControlRodC[2] == null ? 0.000000M : dA_ControlRodC[2];
                        }
                        else
                        {
                            if (strCoilName.Trim() == "B1")
                                dcmRefValue = dB_ControlRodC[0] == null ? 0.000000M : dB_ControlRodC[0];
                            else if (strCoilName.Trim() == "B2")
                                dcmRefValue = dB_ControlRodC[1] == null ? 0.000000M : dB_ControlRodC[1];
                            else
                                dcmRefValue = dB_ControlRodC[2] == null ? 0.000000M : dB_ControlRodC[2];
                        }

                        // 그리드에 측정 값 삽입
                        dgvMeasurement.Rows[i].Cells["C_CapacitanceValue"].Value = dcmMeasurementValue.ToString("F6");
                        dgvMeasurement.Rows[i].Cells["C_Deviation"].Value = (dcmMeasurementValue - dcmRefValue).ToString("F6");

                        // 결과 체크
                        strCResult = m_MeasureProcess.GetMesurmentValueDecision(dcmMeasurementValue, dcmRefValue, dcmDecisionRange_C, dcmEffectiveStandardRange, "C");

                        if ((dcmMeasurementValue - dcmRefValue) == 0)
                            dgvMeasurement.Rows[i].Cells["C_CapacitanceValue"].Style.ForeColor = System.Drawing.Color.Black;
                        else
                        {
                            if (strCResult.Trim() == "부적합")
                            {
                                dgvMeasurement.Rows[i].Cells["C_CapacitanceValue"].Style.ForeColor = System.Drawing.Color.Red;
                            }
                            else if (strCResult.Trim() == "의심")
                            {
                                dgvMeasurement.Rows[i].Cells["C_CapacitanceValue"].Style.ForeColor = System.Drawing.Color.Blue;
                            }
                            else if (strCResult.Trim() == "적합")
                            {
                                dgvMeasurement.Rows[i].Cells["C_CapacitanceValue"].Style.ForeColor = System.Drawing.Color.Black;
                            }
                        }

                        // 컬럼 선택 취소
                        dgvMeasurement.Rows[i].Cells["C_CapacitanceValue"].Selected = false;
                        #endregion
                    }

                    // 측정이 정지되었을 경우
                    if (boolMeasurementStop || !boolMeasurementStart) break;

                    // Q 측정
                    if (chkQ.Checked)
                    {
                        #region Q 측정
                        // 컬럼 선택
                        dgvMeasurement.Rows[i].Cells["Q_FactorValue"].Selected = true;

                        // Q 데이터 측정값 가져오기
                        //boolResult = m_MeasureProcess.functionRCSQMeasurement(ref dcmMeasurementValue);
                        if (!chkL.Checked || dcmL_MeasurementValue == 0M || dcmAC_MeasurementValue == 0M)
                            dcmMeasurementValue = 0.000M;
                        else
                            dcmMeasurementValue = (2M * 3.141592M * dcmFrequency * (dcmL_MeasurementValue / 1000)) / dcmAC_MeasurementValue;
                            
                        // 기준값과 비교 및 편차 산출
                        if (rboCoilA.Checked)
                        {
                            if (strCoilName.Trim() == "A1")
                                dcmRefValue = dA_ControlRodQ[0] == null ? 0.000M : dA_ControlRodQ[0];
                            else if (strCoilName.Trim() == "A2")
                                dcmRefValue = dA_ControlRodQ[1] == null ? 0.000M : dA_ControlRodQ[1];
                            else
                                dcmRefValue = dA_ControlRodQ[2] == null ? 0.000M : dA_ControlRodQ[2];
                        }
                        else
                        {
                            if (strCoilName.Trim() == "B1")
                                dcmRefValue = dB_ControlRodQ[0] == null ? 0.000M : dB_ControlRodQ[0];
                            else if (strCoilName.Trim() == "B2")
                                dcmRefValue = dB_ControlRodQ[1] == null ? 0.000M : dB_ControlRodQ[1];
                            else
                                dcmRefValue = dB_ControlRodQ[2] == null ? 0.000M : dB_ControlRodQ[2];
                        }

                        // 그리드에 측정 값 삽입
                        dgvMeasurement.Rows[i].Cells["Q_FactorValue"].Value = dcmMeasurementValue.ToString("F3");
                        dgvMeasurement.Rows[i].Cells["Q_Deviation"].Value = (dcmMeasurementValue - dcmRefValue).ToString("F3");

                        // 결과 체크
                        strQResult = m_MeasureProcess.GetMesurmentValueDecision(dcmMeasurementValue, dcmRefValue, dcmDecisionRange_Q, dcmEffectiveStandardRange, "Q");

                        if ((dcmMeasurementValue - dcmRefValue) == 0)
                            dgvMeasurement.Rows[i].Cells["Q_FactorValue"].Style.ForeColor = System.Drawing.Color.Black;
                        else
                        {
                            if (strQResult.Trim() == "부적합")
                            {
                                dgvMeasurement.Rows[i].Cells["Q_FactorValue"].Style.ForeColor = System.Drawing.Color.Red;
                            }
                            else if (strQResult.Trim() == "의심")
                            {
                                dgvMeasurement.Rows[i].Cells["Q_FactorValue"].Style.ForeColor = System.Drawing.Color.Blue;
                            }
                            else if (strQResult.Trim() == "적합")
                            {
                                dgvMeasurement.Rows[i].Cells["Q_FactorValue"].Style.ForeColor = System.Drawing.Color.Black;
                            }
                        }

                        // 컬럼 선택 취소
                        dgvMeasurement.Rows[i].Cells["Q_FactorValue"].Selected = false;
                        #endregion
                    }

                    if (dgvMeasurement.Rows[i].Cells["DC_ResistanceValue"].Style.ForeColor == System.Drawing.Color.Red
                        || dgvMeasurement.Rows[i].Cells["AC_ResistanceValue"].Style.ForeColor == System.Drawing.Color.Red
                        || dgvMeasurement.Rows[i].Cells["L_InductanceValue"].Style.ForeColor == System.Drawing.Color.Red
                        || dgvMeasurement.Rows[i].Cells["C_CapacitanceValue"].Style.ForeColor == System.Drawing.Color.Red
                        || dgvMeasurement.Rows[i].Cells["Q_FactorValue"].Style.ForeColor == System.Drawing.Color.Red)
                    {
                        dgvMeasurement.Rows[i].Cells["Result"].Value = "부적합";
                        dgvMeasurement.Rows[i].Cells["Result"].Style.ForeColor = System.Drawing.Color.Red;
                    }
                    else if (dgvMeasurement.Rows[i].Cells["DC_ResistanceValue"].Style.ForeColor == System.Drawing.Color.Blue
                        || dgvMeasurement.Rows[i].Cells["AC_ResistanceValue"].Style.ForeColor == System.Drawing.Color.Blue
                        || dgvMeasurement.Rows[i].Cells["L_InductanceValue"].Style.ForeColor == System.Drawing.Color.Blue
                        || dgvMeasurement.Rows[i].Cells["C_CapacitanceValue"].Style.ForeColor == System.Drawing.Color.Blue
                        || dgvMeasurement.Rows[i].Cells["Q_FactorValue"].Style.ForeColor == System.Drawing.Color.Blue)
                    {
                        dgvMeasurement.Rows[i].Cells["Result"].Value = "의심";
                        dgvMeasurement.Rows[i].Cells["Result"].Style.ForeColor = System.Drawing.Color.Blue;
                    }
                    else
                    {
                        dgvMeasurement.Rows[i].Cells["Result"].Value = "적합";
                        dgvMeasurement.Rows[i].Cells["Result"].Style.ForeColor = System.Drawing.Color.Black;
                    }

                    // DAQ 핀 Off
                    boolResult = m_MeasureProcess.functionDAQPinMapOnOff(strNormalModeDAQPinMap, false);
                    Thread.Sleep(50);
                }
            }
            catch (Exception ex)
            {
                frmMB.TopMost = true;
                frmMB.lblMessage.Text = "측정 오류";
                frmMB.ShowDialog();
                boolResult = false;
                System.Diagnostics.Debug.Print(ex.Message);
            }
            finally
            {
                // DAQ 핀 초기화
                m_MeasureProcess.DigitalDAQ_CloseChannel();

                btnMeasurementStart.Cursor = System.Windows.Forms.Cursors.Hand;
                btnMeasurementStop.Cursor = System.Windows.Forms.Cursors.Default;
                btnSave.Cursor = System.Windows.Forms.Cursors.Hand;

                if (boolMeasurementStop)
                {
                    frmMB.lblMessage.Text = "측정을 중단하였습니다.";
                    frmMB.TopMost = true;
                    frmMB.ShowDialog();
                }
                else if (!boolResult)
                {
                    frmMB.lblMessage.Text = "측정이 실패하였습니다.";
                    frmMB.TopMost = true;
                    frmMB.ShowDialog();
                }
                else
                {
                    btnMeasurementStop.Enabled = false;
                    btnClose.Enabled = false;

                    frmWorkMessageBox frmMessage = new frmWorkMessageBox();
                    frmMessage.lblMessage.Text = "측정이 완료되었습니다.\r\n저장하시겠습니까?";
                    frmMessage.ShowDialog();

                    if (frmMessage.boolOk)
                    {
                        btnSave_Click(null, null);
                    }

                    btnMeasurementStop.Enabled = true;
                    btnClose.Enabled = true;
                }

                boolMeasurementStart = false;

                //측정 시작 중단에 따라 버튼 활성화/비활성화 설정
                SetMeasurementStartButtonEnabled(true);

                // 카드 선택 버튼 활성화 설정
                btnSelectCard.Enabled = true;

                // 측정 시 Control 활성화/비활성화 설정 
                SetControlEnabled(true);

                if (threadMeasurementStart != null)
                    threadMeasurementStart.Abort();
            }
        }

       

        /// <summary>
        /// 코일 그룹 Checked Changed Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void rboCoil_CheckedChanged(object sender, EventArgs e)
        {
            if (boolMeasurementStart || !boolFormLoad) return;

            // 그리드 행 초기화
            dgvMeasurement.Rows.Clear();

            SetStopAndControlButtonAllClear(gbStop, gbControl, gbCardFrame);
            SetStopAndControlButtonAllInitialize();
        }

        /// <summary>
        /// 기준치 팝업 호출
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnReferenceValue_Click(object sender, EventArgs e)
        {
            if (boolMeasurementStart) return;

            frmDRPIReferenceValue frm = new frmDRPIReferenceValue(this);
            frm.ShowDialog();
        }

        /// <summary>
        /// 카드선택 Button Click Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSelectCard_Click(object sender, EventArgs e)
        {
            frmSelectDRPICard frm = new frmSelectDRPICard(this);
            //frm.TopMost = true;
            frm.ShowDialog();
        }

        /// <summary>
        /// DataGridView Cell Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgvMeasurement_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //this.Cursor = Cursors.WaitCursor;

            //try
            //{
            //    // 선택 컬럼의 체크박스 처리
            //    if (dgvMeasurement.Columns[e.ColumnIndex].Name.Trim() == "Select")
            //    {
            //        bool isChecked = (bool)dgvMeasurement[e.ColumnIndex, e.RowIndex].EditedFormattedValue;
            //        dgvMeasurement.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = !isChecked;
            //        dgvMeasurement.EndEdit();
            //    }
            //}
            //catch (Exception ex)
            //{
            //    System.Diagnostics.Debug.Print(ex.Message);
            //}

            //this.Cursor = Cursors.Default;
            if (((DataGridView)sender).RowCount < 1) return;

            isCheck = false;

            if (e.RowIndex >= 0 && e.ColumnIndex == 0)
            {
                ((DataGridView)sender).Rows[e.RowIndex].Cells["Select"].Value = !(bool)((DataGridView)sender).Rows[e.RowIndex].Cells["Select"].Value;
                allCheck.Checked = true;

                for (int i = 0; i < ((DataGridView)sender).Rows.Count; i++)
                {
                    if (!(bool)((DataGridView)sender).Rows[i].Cells["Select"].Value)
                    {
                        allCheck.Checked = false;
                        break;
                    }
                }
            }

            isCheck = true;
        }
    }
}
