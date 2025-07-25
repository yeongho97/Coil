using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Ports;
using System.IO;
using System.Text.RegularExpressions;

// TCP/IP (Lan) 통신용
using System.Net;
using System.Net.Sockets;
using System.Threading;

using NationalInstruments.DAQmx;
using Coil_Diagnostor.Function;

namespace Coil_Diagnostor
{
    public partial class frmConfiguration : Form
    {
        protected Function.FunctionMeasureProcess m_MeasureProcess = new Function.FunctionMeasureProcess();
        protected Function.FunctionDataControl m_db = new Function.FunctionDataControl();
        protected frmMessageBox frmMB = new frmMessageBox();
        protected string strPlantName = "";
        protected bool boolMessageOK = false;
        protected bool boolFormLoad = false;
        protected bool boollInitialize = false;

        public frmConfiguration()
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
                case Keys.F1: // 초기화 버튼
                    btnInitialize1.PerformClick();
                    break;
                case Keys.F2: // 설정 버튼
                    if (tabControl1.SelectedIndex == 0)
                        btnSave1.PerformClick();
                    else if (tabControl1.SelectedIndex == 1)
                        btnSave2.PerformClick();
                    else if (tabControl1.SelectedIndex == 2)
                        btnSave3.PerformClick();
                    else if (tabControl1.SelectedIndex == 3)
                        btnSave4.PerformClick();
                    else if (tabControl1.SelectedIndex == 4)
                        btnSave5.PerformClick();
                    break;
                case Keys.F3: // 통신연결 버튼
                    btnCommunication.PerformClick();
                    break;
                case Keys.F12: // 닫기 버튼
                    if (tabControl1.SelectedIndex == 0)
                        btnClose1.PerformClick();
                    else if (tabControl1.SelectedIndex == 1)
                        btnClose2.PerformClick();
                    else if (tabControl1.SelectedIndex == 2)
                        btnClose3.PerformClick();
                    else if (tabControl1.SelectedIndex == 3)
                        btnClose4.PerformClick();
                    else if (tabControl1.SelectedIndex == 4)
                        btnClose5.PerformClick();
                    break;
            }

            return base.ProcessCmdKey(ref msg, keyData);
        }

        /// <summary>
        /// Form Load
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmConfiguration_Load(object sender, EventArgs e)
        {
            boolFormLoad = true;

            strPlantName = Gini.GetValue("Device", "PlantName");

            // ComboBox 데이터 설정
            SetComboBoxDataBinding();

            // Control - 통신 장비 설정값 초기화
            SetControlInitialize1();

            // Control - 기준값 범위 초기화
            SetControlInitialize4();

            // Control - 취득 데이터 수 초기화
            SetControlInitialize5();

            // 환경설정 - 통신 장비 설정값 가져오기
            GetTabControl1ValueChange();

            // 환경설정 - 기준값 범위 가져오기
            GetTabControl4ValueChange();

            // 환경설정 - 취득 데이터 수 가져오기
            GetTabControl5ValueChange();

            boolFormLoad = false;
        }

        /// <summary>
        /// ComboBox 데이터 설정
        /// </summary>
        private void SetComboBoxDataBinding()
        {
            // DAQ Device
            cboDAQDevice.Items.Clear();

            foreach (string s in DaqSystem.Local.Devices)
            {
                cboDAQDevice.Items.Add(s);
            }

            cboDAQDevice.Items.Add("없음");

            cboDAQDevice.SelectedIndex = cboDAQDevice.Items.Count - 1;

            // 발전소 호기
            string[] strHogi = Gini.GetValue("Combo", "HogiData").Split(',');

            // 초기 값 추가
            cboHogi_RCS.Items.Add("초기값");
            cboHogi_DRPI.Items.Add("초기값");

            for (int i = 0; i < strHogi.Length; i++)
            {
                cboHogi_RCS.Items.Add(strHogi[i].Trim());
                cboHogi_DRPI.Items.Add(strHogi[i].Trim());
            }

            cboHogi_RCS.SelectedIndex = 0;
            cboHogi_DRPI.SelectedIndex = 0;
        }

        /// <summary>
        /// Control - 통신 장비 설정값 초기화
        /// </summary>
        private void SetControlInitialize1()
        {
            boollInitialize = true;

            try
            {
                // 통신 장비 Data 가져오기
                // LCR-Meter
                teLCRMeter_Addr.Text = "";
                teLCRMeter_ID.Text = "";

                // DMM
                bool bUseTDRMeter = Convert.ToBoolean(Gini.GetValue("Device", "USE_TDRMeter"));
                pnlTDRMeterSetting.Visible = bUseTDRMeter;
                teTDRMeter_IPAddress.Text = "";
                teTDRMeter_IPPort.Text = "";

                // DAQ Device Name Data 가져오기
                if (cboDAQDevice != null && cboDAQDevice.Items.Count > 0)
                    cboDAQDevice.SelectedItem = "없음";
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
            finally
            {
                boollInitialize = false;
            }
        }

        /// <summary>
        /// Control - 기준값(RCS) 초기화
        /// </summary>
        private void SetControlInitialize2()
        {
            boollInitialize = true;

            try
            {
                // RCS 기준값 Data 가져오기
                teRdcStop_RCSReferenceValue.Text = "0.000";
                teRdcUp_RCSReferenceValue.Text = "0.000";
                teRdcMove_RCSReferenceValue.Text = "0.000";
                teRacStop_RCSReferenceValue.Text = "0.000";
                teRacUp_RCSReferenceValue.Text = "0.000";
                teRacMove_RCSReferenceValue.Text = "0.000";
                teLStop_RCSReferenceValue.Text = "0.000";
                teLUp_RCSReferenceValue.Text = "0.000";
                teLMove_RCSReferenceValue.Text = "0.000";
                teCStop_RCSReferenceValue.Text = "0.000000";
                teCUp_RCSReferenceValue.Text = "0.000000";
                teCMove_RCSReferenceValue.Text = "0.000000";
                teQStop_RCSReferenceValue.Text = "0.000";
                teQUp_RCSReferenceValue.Text = "0.000";
                teQMove_RCSReferenceValue.Text = "0.000";
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
            finally
            {
                boollInitialize = false;
            }
        }

        /// <summary>
        /// Control - 기준값(DRPI) 초기화
        /// </summary>
        private void SetControlInitialize3()
        {
            boollInitialize = true;

            try
            {
                if (boolFormLoad)
                {
                    // 구분 Label 설정
                    string[] arrayLabelType = Gini.GetValue("Combo", "DRPIReferenceValueLabelType").Split(',');

                    for (int i = 0; i < arrayLabelType.Length; i++)
                    {
                        if (arrayLabelType[i].Trim() == "제어용")
                            lblControl.Text = arrayLabelType[i].Trim();
                        else if (arrayLabelType[i].Trim() == "정지용")
                            lblStop.Text = arrayLabelType[i].Trim();
                    }

                    // 제어용 - 코일 Label 설정
                    arrayLabelType = Gini.GetValue("Combo", "DRPIReferenceValueLabelControlName").Split(',');

                    for (int i = 0; i < arrayLabelType.Length; i++)
                    {
                        if (arrayLabelType[i].Trim() == "A1")
                            lblControlA1.Text = arrayLabelType[i].Trim();
                        else if (arrayLabelType[i].Trim() == "A2")
                            lblControlA2.Text = arrayLabelType[i].Trim();
                        else if (arrayLabelType[i].Trim() == "B1")
                            lblControlB1.Text = arrayLabelType[i].Trim();
                        else if (arrayLabelType[i].Trim() == "B2")
                            lblControlB2.Text = arrayLabelType[i].Trim();
                        else if (arrayLabelType[i].Trim() == "A3 ~ 21")
                            lblControlA.Text = arrayLabelType[i].Trim();
                        else if (arrayLabelType[i].Trim() == "B3 ~ 21")
                            lblControlB.Text = arrayLabelType[i].Trim();
                    }

                    // 정지용 - 코일 Label 설정
                    arrayLabelType = Gini.GetValue("Combo", "DRPIReferenceValueLabelStopName").Split(',');

                    for (int i = 0; i < arrayLabelType.Length; i++)
                    {
                        if (arrayLabelType[i].Trim() == "A1")
                            lblStopA1.Text = arrayLabelType[i].Trim();
                        else if (arrayLabelType[i].Trim() == "A2")
                            lblStopA2.Text = arrayLabelType[i].Trim();
                        else if (arrayLabelType[i].Trim() == "B1")
                            lblStopB1.Text = arrayLabelType[i].Trim();
                        else if (arrayLabelType[i].Trim() == "B2")
                            lblStopB2.Text = arrayLabelType[i].Trim();
                        else if (arrayLabelType[i].Trim() == "A3~6,18~21")
                            lblStopA.Text = arrayLabelType[i].Trim();
                        else if (arrayLabelType[i].Trim() == "B3~6,18~21")
                            lblStopB.Text = arrayLabelType[i].Trim();
                    }
                }

                // DRPI 기준값 Data 가져오기
                // 제어용
                teControlA1Rdc_DRPIReferenceValue.Text = "0.000";
                teControlA1Rac_DRPIReferenceValue.Text = "0.000";
                teControlA1L_DRPIReferenceValue.Text = "0.000";
                teControlA1C_DRPIReferenceValue.Text = "0.000000";
                teControlA1Q_DRPIReferenceValue.Text = "0.000";
                teControlA2Rdc_DRPIReferenceValue.Text = "0.000";
                teControlA2Rac_DRPIReferenceValue.Text = "0.000";
                teControlA2L_DRPIReferenceValue.Text = "0.000";
                teControlA2C_DRPIReferenceValue.Text = "0.000000";
                teControlA2Q_DRPIReferenceValue.Text = "0.000";
                teControlB1Rdc_DRPIReferenceValue.Text = "0.000";
                teControlB1Rac_DRPIReferenceValue.Text = "0.000";
                teControlB1L_DRPIReferenceValue.Text = "0.000";
                teControlB1C_DRPIReferenceValue.Text = "0.000000";
                teControlB1Q_DRPIReferenceValue.Text = "0.000";
                teControlB2Rdc_DRPIReferenceValue.Text = "0.000";
                teControlB2Rac_DRPIReferenceValue.Text = "0.000";
                teControlB2L_DRPIReferenceValue.Text = "0.000";
                teControlB2C_DRPIReferenceValue.Text = "0.000000";
                teControlB2Q_DRPIReferenceValue.Text = "0.000";
                teControlARdc_DRPIReferenceValue.Text = "0.000";
                teControlARac_DRPIReferenceValue.Text = "0.000";
                teControlAL_DRPIReferenceValue.Text = "0.000";
                teControlAC_DRPIReferenceValue.Text = "0.000000";
                teControlAQ_DRPIReferenceValue.Text = "0.000";
                teControlBRdc_DRPIReferenceValue.Text = "0.000";
                teControlBRac_DRPIReferenceValue.Text = "0.000";
                teControlBL_DRPIReferenceValue.Text = "0.000";
                teControlBC_DRPIReferenceValue.Text = "0.000000";
                teControlBQ_DRPIReferenceValue.Text = "0.000";

                // 정지용
                teStopA1Rdc_DRPIReferenceValue.Text = "0.000";
                teStopA1Rac_DRPIReferenceValue.Text = "0.000";
                teStopA1L_DRPIReferenceValue.Text = "0.000";
                teStopA1C_DRPIReferenceValue.Text = "0.000000";
                teStopA1Q_DRPIReferenceValue.Text = "0.000";
                teStopA2Rdc_DRPIReferenceValue.Text = "0.000";
                teStopA2Rac_DRPIReferenceValue.Text = "0.000";
                teStopA2L_DRPIReferenceValue.Text = "0.000";
                teStopA2C_DRPIReferenceValue.Text = "0.000000";
                teStopA2Q_DRPIReferenceValue.Text = "0.000";
                teStopB1Rdc_DRPIReferenceValue.Text = "0.000";
                teStopB1Rac_DRPIReferenceValue.Text = "0.000";
                teStopB1L_DRPIReferenceValue.Text = "0.000";
                teStopB1C_DRPIReferenceValue.Text = "0.000000";
                teStopB1Q_DRPIReferenceValue.Text = "0.000";
                teStopB2Rdc_DRPIReferenceValue.Text = "0.000";
                teStopB2Rac_DRPIReferenceValue.Text = "0.000";
                teStopB2L_DRPIReferenceValue.Text = "0.000";
                teStopB2C_DRPIReferenceValue.Text = "0.000000";
                teStopB2Q_DRPIReferenceValue.Text = "0.000";
                teStopARdc_DRPIReferenceValue.Text = "0.000";
                teStopARac_DRPIReferenceValue.Text = "0.000";
                teStopAL_DRPIReferenceValue.Text = "0.000";
                teStopAC_DRPIReferenceValue.Text = "0.000000";
                teStopAQ_DRPIReferenceValue.Text = "0.000";
                teStopBRdc_DRPIReferenceValue.Text = "0.000";
                teStopBRac_DRPIReferenceValue.Text = "0.000";
                teStopBL_DRPIReferenceValue.Text = "0.000";
                teStopBC_DRPIReferenceValue.Text = "0.000000";
                teStopBQ_DRPIReferenceValue.Text = "0.000";
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
            finally
            {
                boollInitialize = false;
            }
        }

        /// <summary>
        /// Control - 기준값 범위 초기화
        /// </summary>
        private void SetControlInitialize4()
        {
            boollInitialize = true;

            try
            {
                // 기준값 범위 Data 가져오기
                teRdcDecisionRange_ReferenceValue.Text = "0";
                teRacDecisionRange_ReferenceValue.Text = "0";
                teRdcDecisionRange_ReferenceValue.Text = "0";
                teRdcDecisionRange_ReferenceValue.Text = "0";
                teRdcDecisionRange_ReferenceValue.Text = "0";
                teEffectiveStandardRangeOfVariation.Text = "0";
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
            finally
            {
                boollInitialize = false;
            }
        }

        /// <summary>
        /// Control - 취득 데이터 수 초기화
        /// </summary>
        private void SetControlInitialize5()
        {
            boollInitialize = true;

            try
            {
                // 취득 데이터 수 Data 가져오기
                teWheatstoneDataNumber.Text = "0";
                teTemperature_ReferenceValue.Text = "0.00";
                teTemperatureUpDown_ReferenceValue.Text = "0.00";
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
            finally
            {
                boollInitialize = false;
            }
        }

        /// <summary>
        /// 환경설정 - 통신 장비 설정값 가져오기
        /// </summary>
        private void GetTabControl1ValueChange()
        {
            if (boollInitialize) return;

            try
            {
                // 통신 장비 Data 가져오기
                // LCR-Meter
                teLCRMeter_Addr.Text = Gini.GetValue("Device", "LCRMeter_Addr").Trim();
                teLCRMeter_ID.Text = Gini.GetValue("Device", "LCRMeter_ID").Trim();

                // DMM
                teTDRMeter_IPAddress.Text = Gini.GetValue("Device", "TDRMeter_IPAddress").Trim();
                teTDRMeter_IPPort.Text = Gini.GetValue("Device", "TDRMeter_IPPort").Trim();

                // DAQ Device Name Data 가져오기
                if (cboDAQDevice != null && cboDAQDevice.Items.Count > 0)
                    cboDAQDevice.SelectedItem = Gini.GetValue("Device", "DAQDeviceName").Trim();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
        }

        /// <summary>
        /// 환경설정 - 기준 값(RCS) 가져오기
        /// </summary>
        private void GetTabControl2ValueChange()
        {
            if (boollInitialize) return;

            DataTable dt;
            string strSelectHogi = "";
            string strSelectOhDegree = "";

            try
            {
                // RCS 기준값 Data 가져오기
                strSelectHogi = cboHogi_RCS.SelectedItem == null ? "초기값" : cboHogi_RCS.SelectedItem.ToString().Trim();
                strSelectOhDegree = cboOhDegree_RCS.SelectedItem == null ? "제 0 차" : cboOhDegree_RCS.SelectedItem.ToString().Trim();

                dt = new DataTable();
                dt = m_db.GetRCSReferenceValueDataInfo(strPlantName, strSelectHogi, strSelectOhDegree);

                if (dt != null && dt.Rows.Count > 0)
                {
                    teRdcStop_RCSReferenceValue.Text = dt.Rows[0]["RdcStop_RCSReferenceValue"].ToString().Trim();
                    teRdcUp_RCSReferenceValue.Text = dt.Rows[0]["RdcUp_RCSReferenceValue"].ToString().Trim();
                    teRdcMove_RCSReferenceValue.Text = dt.Rows[0]["RdcMove_RCSReferenceValue"].ToString().Trim();
                    teRacStop_RCSReferenceValue.Text = dt.Rows[0]["RacStop_RCSReferenceValue"].ToString().Trim();
                    teRacUp_RCSReferenceValue.Text = dt.Rows[0]["RacUp_RCSReferenceValue"].ToString().Trim();
                    teRacMove_RCSReferenceValue.Text = dt.Rows[0]["RacMove_RCSReferenceValue"].ToString().Trim();
                    teLStop_RCSReferenceValue.Text = dt.Rows[0]["LStop_RCSReferenceValue"].ToString().Trim();
                    teLUp_RCSReferenceValue.Text = dt.Rows[0]["LUp_RCSReferenceValue"].ToString().Trim();
                    teLMove_RCSReferenceValue.Text = dt.Rows[0]["LMove_RCSReferenceValue"].ToString().Trim();
                    teCStop_RCSReferenceValue.Text = dt.Rows[0]["CStop_RCSReferenceValue"].ToString().Trim();
                    teCUp_RCSReferenceValue.Text = dt.Rows[0]["CUp_RCSReferenceValue"].ToString().Trim();
                    teCMove_RCSReferenceValue.Text = dt.Rows[0]["CMove_RCSReferenceValue"].ToString().Trim();
                    teQStop_RCSReferenceValue.Text = dt.Rows[0]["QStop_RCSReferenceValue"].ToString().Trim();
                    teQUp_RCSReferenceValue.Text = dt.Rows[0]["QUp_RCSReferenceValue"].ToString().Trim();
                    teQMove_RCSReferenceValue.Text = dt.Rows[0]["QMove_RCSReferenceValue"].ToString().Trim();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
        }

        /// <summary>
        /// 환경설정 - 기준 값(DRPI) 가져오기
        /// </summary>
        private void GetTabControl3ValueChange()
        {
            if (boollInitialize) return;

            DataTable dt;
            string strSelectHogi = "";
            string strSelectOhDegree = "";

            try
            {
                // DRPI 기준값 Data 가져오기
                strSelectHogi = cboHogi_DRPI.SelectedItem == null ? "초기값" : cboHogi_DRPI.SelectedItem.ToString().Trim();
                strSelectOhDegree = cboOhDegree_DRPI.SelectedItem == null ? "제 0 차" : cboOhDegree_DRPI.SelectedItem.ToString().Trim();

                dt = new DataTable();
                dt = m_db.GetDRPIReferenceValueDataInfo(strPlantName, strSelectHogi, strSelectOhDegree);

                if (dt != null && dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        if (dt.Rows[i]["CoilCabinetType"] == null || dt.Rows[i]["CoilCabinetType"].ToString().Trim() == "") continue;
                        if (dt.Rows[i]["CoilCabinetName"] == null || dt.Rows[i]["CoilCabinetName"].ToString().Trim() == "") continue;

                        if (dt.Rows[i]["CoilCabinetType"].ToString().Trim() == "제어용" && dt.Rows[i]["CoilCabinetName"].ToString().Trim() == "A1")
                        {
                            // 제어용 A1
                            teControlA1Rdc_DRPIReferenceValue.Text = dt.Rows[i]["Rdc_DRPIReferenceValue"] == null ? "0.000" : dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim();
                            teControlA1Rac_DRPIReferenceValue.Text = dt.Rows[i]["Rac_DRPIReferenceValue"] == null ? "0.000" : dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim();
                            teControlA1L_DRPIReferenceValue.Text = dt.Rows[i]["L_DRPIReferenceValue"] == null ? "0.000" : dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim();
                            teControlA1C_DRPIReferenceValue.Text = dt.Rows[i]["C_DRPIReferenceValue"] == null ? "0.000000" : dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim();
                            teControlA1Q_DRPIReferenceValue.Text = dt.Rows[i]["Q_DRPIReferenceValue"] == null ? "0.000" : dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim();

                        }
                        else if (dt.Rows[i]["CoilCabinetType"].ToString().Trim() == "제어용" && dt.Rows[i]["CoilCabinetName"].ToString().Trim() == "A2")
                        {
                            // 제어용 A2
                            teControlA2Rdc_DRPIReferenceValue.Text = dt.Rows[i]["Rdc_DRPIReferenceValue"] == null ? "0.000" : dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim();
                            teControlA2Rac_DRPIReferenceValue.Text = dt.Rows[i]["Rac_DRPIReferenceValue"] == null ? "0.000" : dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim();
                            teControlA2L_DRPIReferenceValue.Text = dt.Rows[i]["L_DRPIReferenceValue"] == null ? "0.000" : dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim();
                            teControlA2C_DRPIReferenceValue.Text = dt.Rows[i]["C_DRPIReferenceValue"] == null ? "0.000000" : dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim();
                            teControlA2Q_DRPIReferenceValue.Text = dt.Rows[i]["Q_DRPIReferenceValue"] == null ? "0.000" : dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim();
                        }
                        else if (dt.Rows[i]["CoilCabinetType"].ToString().Trim() == "제어용" && dt.Rows[i]["CoilCabinetName"].ToString().Trim() == "B1")
                        {
                            // 제어용 B1
                            teControlB1Rdc_DRPIReferenceValue.Text = dt.Rows[i]["Rdc_DRPIReferenceValue"] == null ? "0.000" : dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim();
                            teControlB1Rac_DRPIReferenceValue.Text = dt.Rows[i]["Rac_DRPIReferenceValue"] == null ? "0.000" : dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim();
                            teControlB1L_DRPIReferenceValue.Text = dt.Rows[i]["L_DRPIReferenceValue"] == null ? "0.000" : dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim();
                            teControlB1C_DRPIReferenceValue.Text = dt.Rows[i]["C_DRPIReferenceValue"] == null ? "0.000000" : dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim();
                            teControlB1Q_DRPIReferenceValue.Text = dt.Rows[i]["Q_DRPIReferenceValue"] == null ? "0.000" : dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim();
                        }
                        else if (dt.Rows[i]["CoilCabinetType"].ToString().Trim() == "제어용" && dt.Rows[i]["CoilCabinetName"].ToString().Trim() == "B2")
                        {
                            // 제어용 B2
                            teControlB2Rdc_DRPIReferenceValue.Text = dt.Rows[i]["Rdc_DRPIReferenceValue"] == null ? "0.000" : dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim();
                            teControlB2Rac_DRPIReferenceValue.Text = dt.Rows[i]["Rac_DRPIReferenceValue"] == null ? "0.000" : dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim();
                            teControlB2L_DRPIReferenceValue.Text = dt.Rows[i]["L_DRPIReferenceValue"] == null ? "0.000" : dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim();
                            teControlB2C_DRPIReferenceValue.Text = dt.Rows[i]["C_DRPIReferenceValue"] == null ? "0.000000" : dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim();
                            teControlB2Q_DRPIReferenceValue.Text = dt.Rows[i]["Q_DRPIReferenceValue"] == null ? "0.000" : dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim();
                        }
                        else if (dt.Rows[i]["CoilCabinetType"].ToString().Trim() == "제어용" && dt.Rows[i]["CoilCabinetName"].ToString().Trim() == "A3 ~ 21")
                        {
                            // 제어용 A3 ~ 21
                            teControlARdc_DRPIReferenceValue.Text = dt.Rows[i]["Rdc_DRPIReferenceValue"] == null ? "0.000" : dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim();
                            teControlARac_DRPIReferenceValue.Text = dt.Rows[i]["Rac_DRPIReferenceValue"] == null ? "0.000" : dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim();
                            teControlAL_DRPIReferenceValue.Text = dt.Rows[i]["L_DRPIReferenceValue"] == null ? "0.000" : dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim();
                            teControlAC_DRPIReferenceValue.Text = dt.Rows[i]["C_DRPIReferenceValue"] == null ? "0.000000" : dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim();
                            teControlAQ_DRPIReferenceValue.Text = dt.Rows[i]["Q_DRPIReferenceValue"] == null ? "0.000" : dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim();
                        }
                        else if (dt.Rows[i]["CoilCabinetType"].ToString().Trim() == "제어용" && dt.Rows[i]["CoilCabinetName"].ToString().Trim() == "B3 ~ 21")
                        {
                            // 제어용 B3 ~ 21
                            teControlBRdc_DRPIReferenceValue.Text = dt.Rows[i]["Rdc_DRPIReferenceValue"] == null ? "0.000" : dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim();
                            teControlBRac_DRPIReferenceValue.Text = dt.Rows[i]["Rac_DRPIReferenceValue"] == null ? "0.000" : dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim();
                            teControlBL_DRPIReferenceValue.Text = dt.Rows[i]["L_DRPIReferenceValue"] == null ? "0.000" : dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim();
                            teControlBC_DRPIReferenceValue.Text = dt.Rows[i]["C_DRPIReferenceValue"] == null ? "0.000000" : dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim();
                            teControlBQ_DRPIReferenceValue.Text = dt.Rows[i]["Q_DRPIReferenceValue"] == null ? "0.000" : dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim();
                        }
                        else if (dt.Rows[i]["CoilCabinetType"].ToString().Trim() == "정지용" && dt.Rows[i]["CoilCabinetName"].ToString().Trim() == "A1")
                        {
                            // 정지용 A1
                            teStopA1Rdc_DRPIReferenceValue.Text = dt.Rows[i]["Rdc_DRPIReferenceValue"] == null ? "0.000" : dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim();
                            teStopA1Rac_DRPIReferenceValue.Text = dt.Rows[i]["Rac_DRPIReferenceValue"] == null ? "0.000" : dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim();
                            teStopA1L_DRPIReferenceValue.Text = dt.Rows[i]["L_DRPIReferenceValue"] == null ? "0.000" : dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim();
                            teStopA1C_DRPIReferenceValue.Text = dt.Rows[i]["C_DRPIReferenceValue"] == null ? "0.000000" : dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim();
                            teStopA1Q_DRPIReferenceValue.Text = dt.Rows[i]["Q_DRPIReferenceValue"] == null ? "0.000" : dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim();
                        }
                        else if (dt.Rows[i]["CoilCabinetType"].ToString().Trim() == "정지용" && dt.Rows[i]["CoilCabinetName"].ToString().Trim() == "A2")
                        {
                            // 정지용 A2
                            teStopA2Rdc_DRPIReferenceValue.Text = dt.Rows[i]["Rdc_DRPIReferenceValue"] == null ? "0.000" : dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim();
                            teStopA2Rac_DRPIReferenceValue.Text = dt.Rows[i]["Rac_DRPIReferenceValue"] == null ? "0.000" : dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim();
                            teStopA2L_DRPIReferenceValue.Text = dt.Rows[i]["L_DRPIReferenceValue"] == null ? "0.000" : dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim();
                            teStopA2C_DRPIReferenceValue.Text = dt.Rows[i]["C_DRPIReferenceValue"] == null ? "0.000000" : dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim();
                            teStopA2Q_DRPIReferenceValue.Text = dt.Rows[i]["Q_DRPIReferenceValue"] == null ? "0.000" : dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim();
                        }
                        else if (dt.Rows[i]["CoilCabinetType"].ToString().Trim() == "정지용" && dt.Rows[i]["CoilCabinetName"].ToString().Trim() == "B1")
                        {
                            // 정지용 B1
                            teStopB1Rdc_DRPIReferenceValue.Text = dt.Rows[i]["Rdc_DRPIReferenceValue"] == null ? "0.000" : dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim();
                            teStopB1Rac_DRPIReferenceValue.Text = dt.Rows[i]["Rac_DRPIReferenceValue"] == null ? "0.000" : dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim();
                            teStopB1L_DRPIReferenceValue.Text = dt.Rows[i]["L_DRPIReferenceValue"] == null ? "0.000" : dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim();
                            teStopB1C_DRPIReferenceValue.Text = dt.Rows[i]["C_DRPIReferenceValue"] == null ? "0.000000" : dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim();
                            teStopB1Q_DRPIReferenceValue.Text = dt.Rows[i]["Q_DRPIReferenceValue"] == null ? "0.000" : dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim();
                        }
                        else if (dt.Rows[i]["CoilCabinetType"].ToString().Trim() == "정지용" && dt.Rows[i]["CoilCabinetName"].ToString().Trim() == "B2")
                        {
                            // 정지용 B2
                            teStopB2Rdc_DRPIReferenceValue.Text = dt.Rows[i]["Rdc_DRPIReferenceValue"] == null ? "0.000" : dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim();
                            teStopB2Rac_DRPIReferenceValue.Text = dt.Rows[i]["Rac_DRPIReferenceValue"] == null ? "0.000" : dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim();
                            teStopB2L_DRPIReferenceValue.Text = dt.Rows[i]["L_DRPIReferenceValue"] == null ? "0.000" : dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim();
                            teStopB2C_DRPIReferenceValue.Text = dt.Rows[i]["C_DRPIReferenceValue"] == null ? "0.000000" : dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim();
                            teStopB2Q_DRPIReferenceValue.Text = dt.Rows[i]["Q_DRPIReferenceValue"] == null ? "0.000" : dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim();
                        }
                        else if (dt.Rows[i]["CoilCabinetType"].ToString().Trim() == "정지용" && dt.Rows[i]["CoilCabinetName"].ToString().Trim() == "A3~6,18~21")
                        {
                            // 정지용 A3~6,18~21
                            teStopARdc_DRPIReferenceValue.Text = dt.Rows[i]["Rdc_DRPIReferenceValue"] == null ? "0.000" : dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim();
                            teStopARac_DRPIReferenceValue.Text = dt.Rows[i]["Rac_DRPIReferenceValue"] == null ? "0.000" : dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim();
                            teStopAL_DRPIReferenceValue.Text = dt.Rows[i]["L_DRPIReferenceValue"] == null ? "0.000" : dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim();
                            teStopAC_DRPIReferenceValue.Text = dt.Rows[i]["C_DRPIReferenceValue"] == null ? "0.000000" : dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim();
                            teStopAQ_DRPIReferenceValue.Text = dt.Rows[i]["Q_DRPIReferenceValue"] == null ? "0.000" : dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim();
                        }
                        else if (dt.Rows[i]["CoilCabinetType"].ToString().Trim() == "정지용" && dt.Rows[i]["CoilCabinetName"].ToString().Trim() == "B3~6,18~21")
                        {
                            // 정지용 B3~6,18~21
                            teStopBRdc_DRPIReferenceValue.Text = dt.Rows[i]["Rdc_DRPIReferenceValue"] == null ? "0.000" : dt.Rows[i]["Rdc_DRPIReferenceValue"].ToString().Trim();
                            teStopBRac_DRPIReferenceValue.Text = dt.Rows[i]["Rac_DRPIReferenceValue"] == null ? "0.000" : dt.Rows[i]["Rac_DRPIReferenceValue"].ToString().Trim();
                            teStopBL_DRPIReferenceValue.Text = dt.Rows[i]["L_DRPIReferenceValue"] == null ? "0.000" : dt.Rows[i]["L_DRPIReferenceValue"].ToString().Trim();
                            teStopBC_DRPIReferenceValue.Text = dt.Rows[i]["C_DRPIReferenceValue"] == null ? "0.000000" : dt.Rows[i]["C_DRPIReferenceValue"].ToString().Trim();
                            teStopBQ_DRPIReferenceValue.Text = dt.Rows[i]["Q_DRPIReferenceValue"] == null ? "0.000" : dt.Rows[i]["Q_DRPIReferenceValue"].ToString().Trim();
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
        /// 환경설정 - 기준값 범위 가져오기
        /// </summary>
        private void GetTabControl4ValueChange()
        {
            if (boollInitialize) return;

            try
            {
                // 기준값 범위 Data 가져오기
                teRdcDecisionRange_ReferenceValue.Text = Gini.GetValue("ReferenceValue", "RdcDecisionRange_ReferenceValue").Trim();
                teRacDecisionRange_ReferenceValue.Text = Gini.GetValue("ReferenceValue", "RacDecisionRange_ReferenceValue").Trim();
                teLDecisionRange_ReferenceValue.Text = Gini.GetValue("ReferenceValue", "LDecisionRange_ReferenceValue").Trim();
                teCDecisionRange_ReferenceValue.Text = Gini.GetValue("ReferenceValue", "CDecisionRange_ReferenceValue").Trim();
                teQDecisionRange_ReferenceValue.Text = Gini.GetValue("ReferenceValue", "QDecisionRange_ReferenceValue").Trim();
                teEffectiveStandardRangeOfVariation.Text = Gini.GetValue("ReferenceValue", "EffectiveStandardRangeOfVariation").Trim();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
        }

        /// <summary>
        /// 환경설정 - 취득 데이터 수 가져오기
        /// </summary>
        private void GetTabControl5ValueChange()
        {
            if (boollInitialize) return;

            try
            {
                // 취득 데이터 수 Data 가져오기
                teWheatstoneDataNumber.Text = Gini.GetValue("Wheatstone", "WheatstoneDataNumber").Trim();
                teTemperature_ReferenceValue.Text = Convert.ToDecimal(Gini.GetValue("ReferenceValue", "Temperature_ReferenceValue")).ToString("F2").Trim();
                teTemperatureUpDown_ReferenceValue.Text = Convert.ToDecimal(Gini.GetValue("ReferenceValue", "TemperatureUpDown_ReferenceValue")).ToString("F5").Trim();
                nmWheatstoneMode_Calculate1.Value = Convert.ToDecimal(Gini.GetValue("Wheatstone", "WheatstoneMode_Calculate1"));
                nmWheatstoneMode_Calculate2.Value = Convert.ToDecimal(Gini.GetValue("Wheatstone", "WheatstoneMode_Calculate2"));
                nmWheatstoneMode_Calculate3.Value = Convert.ToDecimal(Gini.GetValue("Wheatstone", "WheatstoneMode_Calculate3"));
                nmWheatstoneMod_OffSet.Value = Convert.ToDecimal(Gini.GetValue("Wheatstone", "WheatstoneMod_OffSet"));
                nmWheatstoneMode_Compensation.Value = Convert.ToDecimal(Gini.GetValue("Wheatstone", "WheatstoneMode_Compensation"));
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
        }

        /// <summary>
        /// 초기화 버튼 이벤트
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnInitialize_Click(object sender, EventArgs e)
        {
            ((Button)sender).ForeColor = System.Drawing.Color.Blue;

            // 환경설정 - 통신 장비 설정값 가져오기
            GetTabControl1ValueChange();

            ((Button)sender).ForeColor = System.Drawing.Color.White;
        }

        /// <summary>
        /// 저장 버튼 이벤트
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSave_Click(object sender, EventArgs e)
        {
            string[] arrayData;
            string strHogi = "", strOhDegree = "", strCoilCabinetType = "", strCoilCabinetName = "";
            bool boolSave = false;
            string strMessage = "";
            bool boolMessage = false;

            try
            {
                ((Button)sender).ForeColor = System.Drawing.Color.Blue;
                string val;
                // 통신 장비 설정값 저장
                switch (((Button)sender).Name.Trim())
                {
                    case "btnSave1":
                        #region
                        strMessage = "통신 장비 설정 데이터를";
                        boolSave = true;

                        val = teLCRMeter_Addr.Text.Trim();
                        boolSave &= Gini.SetValue("Device", "LCRMeter_Addr", val);
                        val = teLCRMeter_ID.Text.Trim();
                        boolSave &= Gini.SetValue("Device", "LCRMeter_ID", val);
                        val = teTDRMeter_IPAddress.Text.Trim();
                        boolSave &= Gini.SetValue("Device", "TDRMeter_IPAddress", val);
                        val = teTDRMeter_IPPort.Text.Trim();
                        boolSave &= Gini.SetValue("Device", "TDRMeter_IPPort", val);
                        val = cboDAQDevice.SelectedItem == null ? "없음" : cboDAQDevice.SelectedItem.ToString().Trim();
                        boolSave &= Gini.SetValue("Device", "DAQDeviceName", val);

                        #endregion
                        break;
                    case "btnSave2":
                        #region
                        strMessage = "기준값(RCS) 데이터를";
                        strHogi = cboHogi_RCS.SelectedItem == null ? "초기값" : cboHogi_RCS.SelectedItem.ToString().Trim();
                        strOhDegree = cboOhDegree_RCS.SelectedItem == null ? "제 0 차" : cboOhDegree_RCS.SelectedItem.ToString().Trim();

                        arrayData = new string[15];

                        arrayData[0] = teRdcStop_RCSReferenceValue.Text.Trim();
                        arrayData[1] = teRdcUp_RCSReferenceValue.Text.Trim();
                        arrayData[2] = teRdcMove_RCSReferenceValue.Text.Trim();
                        arrayData[3] = teRacStop_RCSReferenceValue.Text.Trim();
                        arrayData[4] = teRacUp_RCSReferenceValue.Text.Trim();
                        arrayData[5] = teRacMove_RCSReferenceValue.Text.Trim();
                        arrayData[6] = teLStop_RCSReferenceValue.Text.Trim();
                        arrayData[7] = teLUp_RCSReferenceValue.Text.Trim();
                        arrayData[8] = teLMove_RCSReferenceValue.Text.Trim();
                        arrayData[9] = teCStop_RCSReferenceValue.Text.Trim();
                        arrayData[10] = teCUp_RCSReferenceValue.Text.Trim();
                        arrayData[11] = teCMove_RCSReferenceValue.Text.Trim();
                        arrayData[12] = teQStop_RCSReferenceValue.Text.Trim();
                        arrayData[13] = teQUp_RCSReferenceValue.Text.Trim();
                        arrayData[14] = teQMove_RCSReferenceValue.Text.Trim();

                        if ((m_db.GetRCSReferenceValueDataCount(strPlantName.Trim(), strHogi.Trim(), strOhDegree.Trim())) > 0)
                        {
                            boolSave = m_db.UpDateRCSReferenceValueDataInfo(strPlantName.Trim(), strHogi.Trim(), strOhDegree.Trim(), arrayData);
                        }
                        else
                        {
                            boolSave = m_db.InsertRCSReferenceValueDataInfo(strPlantName.Trim(), strHogi.Trim(), strOhDegree.Trim(), arrayData);
                        }
                        #endregion
                        break;
                    case "btnSave3":
                        #region
                        strHogi = cboHogi_DRPI.SelectedItem == null ? "초기값" : cboHogi_DRPI.SelectedItem.ToString().Trim();
                        strOhDegree = cboOhDegree_DRPI.SelectedItem == null ? "제 0 차" : cboOhDegree_DRPI.SelectedItem.ToString().Trim();

                        // 제어용 A1
                        strMessage = "기준값(DRPI)의 제어용 A1 데이터를";
                        arrayData = new string[5];

                        strCoilCabinetType = lblControl.Text.Trim();
                        strCoilCabinetName = lblControlA1.Text.Trim();

                        arrayData[0] = teControlA1Rdc_DRPIReferenceValue.Text.Trim();
                        arrayData[1] = teControlA1Rac_DRPIReferenceValue.Text.Trim();
                        arrayData[2] = teControlA1L_DRPIReferenceValue.Text.Trim();
                        arrayData[3] = teControlA1C_DRPIReferenceValue.Text.Trim();
                        arrayData[4] = teControlA1Q_DRPIReferenceValue.Text.Trim();

                        if ((m_db.GetDRPIReferenceValueDataCount(strPlantName.Trim(), strHogi.Trim(), strOhDegree.Trim(), strCoilCabinetType, strCoilCabinetName)) > 0)
                        {
                            boolSave = m_db.UpDateDRPIReferenceValueDataInfo(strPlantName.Trim(), strHogi.Trim(), strOhDegree.Trim(), strCoilCabinetType, strCoilCabinetName, arrayData);
                        }
                        else
                        {
                            boolSave = m_db.InsertDRPIReferenceValueDataInfo(strPlantName.Trim(), strHogi.Trim(), strOhDegree.Trim(), strCoilCabinetType, strCoilCabinetName, arrayData);
                        }

                        if (!boolSave)
                        {
                            boolMessage = true;

                            frmMB.lblMessage.Text = strMessage.Trim() + " 저장 중 실패";
                            frmMB.TopMost = true;
                            frmMB.ShowDialog();
                        }
                        else
                        {
                            // 제어용 A2
                            strMessage = "기준값(DRPI)의 제어용 A2 데이터를";
                            arrayData = new string[5];

                            strCoilCabinetType = lblControl.Text.Trim();
                            strCoilCabinetName = lblControlA2.Text.Trim();

                            arrayData[0] = teControlA2Rdc_DRPIReferenceValue.Text.Trim();
                            arrayData[1] = teControlA2Rac_DRPIReferenceValue.Text.Trim();
                            arrayData[2] = teControlA2L_DRPIReferenceValue.Text.Trim();
                            arrayData[3] = teControlA2C_DRPIReferenceValue.Text.Trim();
                            arrayData[4] = teControlA2Q_DRPIReferenceValue.Text.Trim();

                            if ((m_db.GetDRPIReferenceValueDataCount(strPlantName.Trim(), strHogi.Trim(), strOhDegree.Trim(), strCoilCabinetType, strCoilCabinetName)) > 0)
                            {
                                boolSave = m_db.UpDateDRPIReferenceValueDataInfo(strPlantName.Trim(), strHogi.Trim(), strOhDegree.Trim(), strCoilCabinetType, strCoilCabinetName, arrayData);
                            }
                            else
                            {
                                boolSave = m_db.InsertDRPIReferenceValueDataInfo(strPlantName.Trim(), strHogi.Trim(), strOhDegree.Trim(), strCoilCabinetType, strCoilCabinetName, arrayData);
                            }
                        }

                        if (!boolSave)
                        {
                            boolMessage = true;

                            frmMB.lblMessage.Text = strMessage.Trim() + " 저장 중 실패";
                            frmMB.TopMost = true;
                            frmMB.ShowDialog();
                        }
                        else
                        {
                            // 제어용 B1
                            strMessage = "기준값(DRPI)의 제어용 B1 데이터를";
                            arrayData = new string[5];

                            strCoilCabinetType = lblControl.Text.Trim();
                            strCoilCabinetName = lblControlB1.Text.Trim();

                            arrayData[0] = teControlB1Rdc_DRPIReferenceValue.Text.Trim();
                            arrayData[1] = teControlB1Rac_DRPIReferenceValue.Text.Trim();
                            arrayData[2] = teControlB1L_DRPIReferenceValue.Text.Trim();
                            arrayData[3] = teControlB1C_DRPIReferenceValue.Text.Trim();
                            arrayData[4] = teControlB1Q_DRPIReferenceValue.Text.Trim();

                            if ((m_db.GetDRPIReferenceValueDataCount(strPlantName.Trim(), strHogi.Trim(), strOhDegree.Trim(), strCoilCabinetType, strCoilCabinetName)) > 0)
                            {
                                boolSave = m_db.UpDateDRPIReferenceValueDataInfo(strPlantName.Trim(), strHogi.Trim(), strOhDegree.Trim(), strCoilCabinetType, strCoilCabinetName, arrayData);
                            }
                            else
                            {
                                boolSave = m_db.InsertDRPIReferenceValueDataInfo(strPlantName.Trim(), strHogi.Trim(), strOhDegree.Trim(), strCoilCabinetType, strCoilCabinetName, arrayData);
                            }
                        }

                        if (!boolSave)
                        {
                            boolMessage = true;

                            frmMB.lblMessage.Text = strMessage.Trim() + " 저장 중 실패";
                            frmMB.TopMost = true;
                            frmMB.ShowDialog();
                        }
                        else
                        {
                            // 제어용 B2
                            strMessage = "기준값(DRPI)의 제어용 B2 데이터를";
                            arrayData = new string[5];

                            strCoilCabinetType = lblControl.Text.Trim();
                            strCoilCabinetName = lblControlB2.Text.Trim();

                            arrayData[0] = teControlB2Rdc_DRPIReferenceValue.Text.Trim();
                            arrayData[1] = teControlB2Rac_DRPIReferenceValue.Text.Trim();
                            arrayData[2] = teControlB2L_DRPIReferenceValue.Text.Trim();
                            arrayData[3] = teControlB2C_DRPIReferenceValue.Text.Trim();
                            arrayData[4] = teControlB2Q_DRPIReferenceValue.Text.Trim();

                            if ((m_db.GetDRPIReferenceValueDataCount(strPlantName.Trim(), strHogi.Trim(), strOhDegree.Trim(), strCoilCabinetType, strCoilCabinetName)) > 0)
                            {
                                boolSave = m_db.UpDateDRPIReferenceValueDataInfo(strPlantName.Trim(), strHogi.Trim(), strOhDegree.Trim(), strCoilCabinetType, strCoilCabinetName, arrayData);
                            }
                            else
                            {
                                boolSave = m_db.InsertDRPIReferenceValueDataInfo(strPlantName.Trim(), strHogi.Trim(), strOhDegree.Trim(), strCoilCabinetType, strCoilCabinetName, arrayData);
                            }
                        }

                        if (!boolSave)
                        {
                            boolMessage = true;

                            frmMB.lblMessage.Text = strMessage.Trim() + " 저장 중 실패";
                            frmMB.TopMost = true;
                            frmMB.ShowDialog();
                        }
                        else
                        {
                            // 제어용 A3 ~ 21
                            strMessage = "기준값(DRPI)의 제어용 A3 ~ 21 데이터를";
                            arrayData = new string[5];

                            strCoilCabinetType = lblControl.Text.Trim();
                            strCoilCabinetName = lblControlA.Text.Trim();

                            arrayData[0] = teControlARdc_DRPIReferenceValue.Text.Trim();
                            arrayData[1] = teControlARac_DRPIReferenceValue.Text.Trim();
                            arrayData[2] = teControlAL_DRPIReferenceValue.Text.Trim();
                            arrayData[3] = teControlAC_DRPIReferenceValue.Text.Trim();
                            arrayData[4] = teControlAQ_DRPIReferenceValue.Text.Trim();

                            if ((m_db.GetDRPIReferenceValueDataCount(strPlantName.Trim(), strHogi.Trim(), strOhDegree.Trim(), strCoilCabinetType, strCoilCabinetName)) > 0)
                            {
                                boolSave = m_db.UpDateDRPIReferenceValueDataInfo(strPlantName.Trim(), strHogi.Trim(), strOhDegree.Trim(), strCoilCabinetType, strCoilCabinetName, arrayData);
                            }
                            else
                            {
                                boolSave = m_db.InsertDRPIReferenceValueDataInfo(strPlantName.Trim(), strHogi.Trim(), strOhDegree.Trim(), strCoilCabinetType, strCoilCabinetName, arrayData);
                            }
                        }

                        if (!boolSave)
                        {
                            boolMessage = true;

                            frmMB.lblMessage.Text = strMessage.Trim() + " 저장 중 실패";
                            frmMB.TopMost = true;
                            frmMB.ShowDialog();
                        }
                        else
                        {
                            // 제어용 B3 ~ 21
                            strMessage = "기준값(DRPI)의 제어용 B3 ~ 21 데이터를";
                            arrayData = new string[5];

                            strCoilCabinetType = lblControl.Text.Trim();
                            strCoilCabinetName = lblControlB.Text.Trim();

                            arrayData[0] = teControlBRdc_DRPIReferenceValue.Text.Trim();
                            arrayData[1] = teControlBRac_DRPIReferenceValue.Text.Trim();
                            arrayData[2] = teControlBL_DRPIReferenceValue.Text.Trim();
                            arrayData[3] = teControlBC_DRPIReferenceValue.Text.Trim();
                            arrayData[4] = teControlBQ_DRPIReferenceValue.Text.Trim();

                            if ((m_db.GetDRPIReferenceValueDataCount(strPlantName.Trim(), strHogi.Trim(), strOhDegree.Trim(), strCoilCabinetType, strCoilCabinetName)) > 0)
                            {
                                boolSave = m_db.UpDateDRPIReferenceValueDataInfo(strPlantName.Trim(), strHogi.Trim(), strOhDegree.Trim(), strCoilCabinetType, strCoilCabinetName, arrayData);
                            }
                            else
                            {
                                boolSave = m_db.InsertDRPIReferenceValueDataInfo(strPlantName.Trim(), strHogi.Trim(), strOhDegree.Trim(), strCoilCabinetType, strCoilCabinetName, arrayData);
                            }
                        }

                        if (!boolSave)
                        {
                            boolMessage = true;

                            frmMB.lblMessage.Text = strMessage.Trim() + " 저장 중 실패";
                            frmMB.TopMost = true;
                            frmMB.ShowDialog();
                        }
                        else
                        {
                            // 정지용 A1
                            strMessage = "기준값(DRPI)의 정지용 A1 데이터를";
                            arrayData = new string[5];

                            strCoilCabinetType = lblStop.Text.Trim();
                            strCoilCabinetName = lblStopA1.Text.Trim();

                            arrayData[0] = teStopA1Rdc_DRPIReferenceValue.Text.Trim();
                            arrayData[1] = teStopA1Rac_DRPIReferenceValue.Text.Trim();
                            arrayData[2] = teStopA1L_DRPIReferenceValue.Text.Trim();
                            arrayData[3] = teStopA1C_DRPIReferenceValue.Text.Trim();
                            arrayData[4] = teStopA1Q_DRPIReferenceValue.Text.Trim();

                            if ((m_db.GetDRPIReferenceValueDataCount(strPlantName.Trim(), strHogi.Trim(), strOhDegree.Trim(), strCoilCabinetType, strCoilCabinetName)) > 0)
                            {
                                boolSave = m_db.UpDateDRPIReferenceValueDataInfo(strPlantName.Trim(), strHogi.Trim(), strOhDegree.Trim(), strCoilCabinetType, strCoilCabinetName, arrayData);
                            }
                            else
                            {
                                boolSave = m_db.InsertDRPIReferenceValueDataInfo(strPlantName.Trim(), strHogi.Trim(), strOhDegree.Trim(), strCoilCabinetType, strCoilCabinetName, arrayData);
                            }
                        }

                        if (!boolSave)
                        {
                            boolMessage = true;

                            frmMB.lblMessage.Text = strMessage.Trim() + " 저장 중 실패";
                            frmMB.TopMost = true;
                            frmMB.ShowDialog();
                        }
                        else
                        {
                            // 정지용 A2
                            strMessage = "기준값(DRPI)의 정지용 A2 데이터를";
                            arrayData = new string[5];

                            strCoilCabinetType = lblStop.Text.Trim();
                            strCoilCabinetName = lblStopA2.Text.Trim();

                            arrayData[0] = teStopA2Rdc_DRPIReferenceValue.Text.Trim();
                            arrayData[1] = teStopA2Rac_DRPIReferenceValue.Text.Trim();
                            arrayData[2] = teStopA2L_DRPIReferenceValue.Text.Trim();
                            arrayData[3] = teStopA2C_DRPIReferenceValue.Text.Trim();
                            arrayData[4] = teStopA2Q_DRPIReferenceValue.Text.Trim();

                            if ((m_db.GetDRPIReferenceValueDataCount(strPlantName.Trim(), strHogi.Trim(), strOhDegree.Trim(), strCoilCabinetType, strCoilCabinetName)) > 0)
                            {
                                boolSave = m_db.UpDateDRPIReferenceValueDataInfo(strPlantName.Trim(), strHogi.Trim(), strOhDegree.Trim(), strCoilCabinetType, strCoilCabinetName, arrayData);
                            }
                            else
                            {
                                boolSave = m_db.InsertDRPIReferenceValueDataInfo(strPlantName.Trim(), strHogi.Trim(), strOhDegree.Trim(), strCoilCabinetType, strCoilCabinetName, arrayData);
                            }
                        }

                        if (!boolSave)
                        {
                            boolMessage = true;

                            frmMB.lblMessage.Text = strMessage.Trim() + " 저장 중 실패";
                            frmMB.TopMost = true;
                            frmMB.ShowDialog();
                        }
                        else
                        {
                            // 정지용 B1
                            strMessage = "기준값(DRPI)의 정지용 B1 데이터를";
                            arrayData = new string[5];

                            strCoilCabinetType = lblStop.Text.Trim();
                            strCoilCabinetName = lblStopB1.Text.Trim();

                            arrayData[0] = teStopB1Rdc_DRPIReferenceValue.Text.Trim();
                            arrayData[1] = teStopB1Rac_DRPIReferenceValue.Text.Trim();
                            arrayData[2] = teStopB1L_DRPIReferenceValue.Text.Trim();
                            arrayData[3] = teStopB1C_DRPIReferenceValue.Text.Trim();
                            arrayData[4] = teStopB1Q_DRPIReferenceValue.Text.Trim();

                            if ((m_db.GetDRPIReferenceValueDataCount(strPlantName.Trim(), strHogi.Trim(), strOhDegree.Trim(), strCoilCabinetType, strCoilCabinetName)) > 0)
                            {
                                boolSave = m_db.UpDateDRPIReferenceValueDataInfo(strPlantName.Trim(), strHogi.Trim(), strOhDegree.Trim(), strCoilCabinetType, strCoilCabinetName, arrayData);
                            }
                            else
                            {
                                boolSave = m_db.InsertDRPIReferenceValueDataInfo(strPlantName.Trim(), strHogi.Trim(), strOhDegree.Trim(), strCoilCabinetType, strCoilCabinetName, arrayData);
                            }
                        }

                        if (!boolSave)
                        {
                            boolMessage = true;

                            frmMB.lblMessage.Text = strMessage.Trim() + " 저장 중 실패";
                            frmMB.TopMost = true;
                            frmMB.ShowDialog();
                        }
                        else
                        {
                            // 정지용 B2
                            strMessage = "기준값(DRPI)의 정지용 B2 데이터를";
                            arrayData = new string[5];

                            strCoilCabinetType = lblStop.Text.Trim();
                            strCoilCabinetName = lblStopB2.Text.Trim();

                            arrayData[0] = teStopB2Rdc_DRPIReferenceValue.Text.Trim();
                            arrayData[1] = teStopB2Rac_DRPIReferenceValue.Text.Trim();
                            arrayData[2] = teStopB2L_DRPIReferenceValue.Text.Trim();
                            arrayData[3] = teStopB2C_DRPIReferenceValue.Text.Trim();
                            arrayData[4] = teStopB2Q_DRPIReferenceValue.Text.Trim();

                            if ((m_db.GetDRPIReferenceValueDataCount(strPlantName.Trim(), strHogi.Trim(), strOhDegree.Trim(), strCoilCabinetType, strCoilCabinetName)) > 0)
                            {
                                boolSave = m_db.UpDateDRPIReferenceValueDataInfo(strPlantName.Trim(), strHogi.Trim(), strOhDegree.Trim(), strCoilCabinetType, strCoilCabinetName, arrayData);
                            }
                            else
                            {
                                boolSave = m_db.InsertDRPIReferenceValueDataInfo(strPlantName.Trim(), strHogi.Trim(), strOhDegree.Trim(), strCoilCabinetType, strCoilCabinetName, arrayData);
                            }
                        }

                        if (!boolSave)
                        {
                            boolMessage = true;

                            frmMB.lblMessage.Text = strMessage.Trim() + " 저장 중 실패";
                            frmMB.TopMost = true;
                            frmMB.ShowDialog();
                        }
                        else
                        {
                            // 정지용 A3 ~ 21
                            strMessage = "기준값(DRPI)의 정지용 A3 ~ 21 데이터를";
                            arrayData = new string[5];

                            strCoilCabinetType = lblStop.Text.Trim();
                            strCoilCabinetName = lblStopA.Text.Trim();

                            arrayData[0] = teStopARdc_DRPIReferenceValue.Text.Trim();
                            arrayData[1] = teStopARac_DRPIReferenceValue.Text.Trim();
                            arrayData[2] = teStopAL_DRPIReferenceValue.Text.Trim();
                            arrayData[3] = teStopAC_DRPIReferenceValue.Text.Trim();
                            arrayData[4] = teStopAQ_DRPIReferenceValue.Text.Trim();

                            if ((m_db.GetDRPIReferenceValueDataCount(strPlantName.Trim(), strHogi.Trim(), strOhDegree.Trim(), strCoilCabinetType, strCoilCabinetName)) > 0)
                            {
                                boolSave = m_db.UpDateDRPIReferenceValueDataInfo(strPlantName.Trim(), strHogi.Trim(), strOhDegree.Trim(), strCoilCabinetType, strCoilCabinetName, arrayData);
                            }
                            else
                            {
                                boolSave = m_db.InsertDRPIReferenceValueDataInfo(strPlantName.Trim(), strHogi.Trim(), strOhDegree.Trim(), strCoilCabinetType, strCoilCabinetName, arrayData);
                            }
                        }

                        if (!boolSave)
                        {
                            boolMessage = true;

                            frmMB.lblMessage.Text = strMessage.Trim() + " 저장 중 실패";
                            frmMB.TopMost = true;
                            frmMB.ShowDialog();
                        }
                        else
                        {
                            // 정지용 B3 ~ 21
                            strMessage = "기준값(DRPI)의 정지용 B3 ~ 21 데이터를";
                            arrayData = new string[5];

                            strCoilCabinetType = lblStop.Text.Trim();
                            strCoilCabinetName = lblStopB.Text.Trim();

                            arrayData[0] = teStopBRdc_DRPIReferenceValue.Text.Trim();
                            arrayData[1] = teStopBRac_DRPIReferenceValue.Text.Trim();
                            arrayData[2] = teStopBL_DRPIReferenceValue.Text.Trim();
                            arrayData[3] = teStopBC_DRPIReferenceValue.Text.Trim();
                            arrayData[4] = teStopBQ_DRPIReferenceValue.Text.Trim();

                            if ((m_db.GetDRPIReferenceValueDataCount(strPlantName.Trim(), strHogi.Trim(), strOhDegree.Trim(), strCoilCabinetType, strCoilCabinetName)) > 0)
                            {
                                boolSave = m_db.UpDateDRPIReferenceValueDataInfo(strPlantName.Trim(), strHogi.Trim(), strOhDegree.Trim(), strCoilCabinetType, strCoilCabinetName, arrayData);
                            }
                            else
                            {
                                boolSave = m_db.InsertDRPIReferenceValueDataInfo(strPlantName.Trim(), strHogi.Trim(), strOhDegree.Trim(), strCoilCabinetType, strCoilCabinetName, arrayData);
                            }
                        }

                        if (boolSave) strMessage = "기준값(DRPI) 데이터를";
                        #endregion
                        break;
                    case "btnSave4":
                        #region
                        strMessage = "기준값 범위 데이터를";
                        boolSave = true;

                        val = teRdcDecisionRange_ReferenceValue.Text.Trim() == "" ? "0" : teRdcDecisionRange_ReferenceValue.Text.Trim();
                        boolSave &= Gini.SetValue("ReferenceValue", "teRdcDecisionRange_ReferenceValue", val);
                        val = teRacDecisionRange_ReferenceValue.Text.Trim() == "" ? "0" : teRacDecisionRange_ReferenceValue.Text.Trim();
                        boolSave &= Gini.SetValue("ReferenceValue", "teRacDecisionRange_ReferenceValue", val);
                        val = teLDecisionRange_ReferenceValue.Text.Trim() == "" ? "0" :  teLDecisionRange_ReferenceValue.Text.Trim();
                        boolSave &= Gini.SetValue("ReferenceValue", "teLDecisionRange_ReferenceValue", val);
                        val = teCDecisionRange_ReferenceValue.Text.Trim() == "" ? "0" : teCDecisionRange_ReferenceValue.Text.Trim();
                        boolSave &= Gini.SetValue("ReferenceValue", "teCDecisionRange_ReferenceValue", val);
                        val = teQDecisionRange_ReferenceValue.Text.Trim() == "" ? "0" : teQDecisionRange_ReferenceValue.Text.Trim();
                        boolSave &= Gini.SetValue("ReferenceValue", "teQDecisionRange_ReferenceValue", val);
                        val = teEffectiveStandardRangeOfVariation.Text.Trim() == "" ? "0" : teEffectiveStandardRangeOfVariation.Text.Trim();
                        boolSave &= Gini.SetValue("ReferenceValue", "teEffectiveStandardRangeOfVariation", val);

                        #endregion
                        break;
                    case "btnSave5":
                        #region
                        strMessage = "취득 데이터 수 데이터를";
                        boolSave = true;

                        val = teWheatstoneDataNumber.Text.Trim() == "" ? "0" : teWheatstoneDataNumber.Text.Trim();
                        boolSave &= Gini.SetValue("Wheatstone", "WheatstoneDataNumber", val);
                        val = teTemperature_ReferenceValue.Text.Trim() == "" ? "0.00" : teTemperature_ReferenceValue.Text.Trim();
                        boolSave &= Gini.SetValue("ReferenceValue", "Temperature_ReferenceValue", val);
                        val = teTemperatureUpDown_ReferenceValue.Text.Trim() == "" ? "0.00000" : teTemperatureUpDown_ReferenceValue.Text.Trim();
                        boolSave &= Gini.SetValue("ReferenceValue", "TemperatureUpDown_ReferenceValue", val);
                        val = nmWheatstoneMode_Calculate1.Value.ToString();
                        boolSave &= Gini.SetValue("Wheatstone", "WheatstoneMode_Calculate1", val);
                        val = nmWheatstoneMode_Calculate2.Value.ToString();
                        boolSave &= Gini.SetValue("Wheatstone", "WheatstoneMode_Calculate2", val);
                        val = nmWheatstoneMode_Calculate3.Value.ToString();
                        boolSave &= Gini.SetValue("Wheatstone", "WheatstoneMode_Calculate3", val);
                        val = nmWheatstoneMod_OffSet.Value.ToString();
                        boolSave &= Gini.SetValue("Wheatstone", "WheatstoneMod_OffSet", val);
                        val = nmWheatstoneMode_Compensation.Value.ToString();
                        boolSave &= Gini.SetValue("Wheatstone", "WheatstoneMode_Compensation", val);

                        #endregion
                        break;
                }

                if (boolSave)
                {
                    frmMB.lblMessage.Text = strMessage.Trim() + " 저장 완료";
                    frmMB.TopMost = true;
                    frmMB.ShowDialog();
                }
                else if (!boolMessage)
                {
                    frmMB.lblMessage.Text = strMessage.Trim() + " 저장 중 실패";
                    frmMB.TopMost = true;
                    frmMB.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                frmMB.lblMessage.Text = strMessage.Trim() + " 저장 중 오류";
                frmMB.TopMost = true;
                frmMB.ShowDialog();
                System.Diagnostics.Debug.Print(ex.Message);
            }
            finally
            {
                ((Button)sender).ForeColor = System.Drawing.Color.White;
            }
        }

        /// <summary>
        /// 닫기 버튼 이벤트
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();  
        }

        /// <summary>
        /// TextBox Leave Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TextBox_Leave(object sender, EventArgs e)
        {
            if (((TextBox)sender).Text.ToString().Trim() == "")
            {
                ((TextBox)sender).Text = "0.000";
            }

            /*
            // 기준값 RCS일 경우
            if (tabControl1.SelectedIndex == 1)
            {
                if (((TextBox)sender).Name.Trim() == "teRacStop_RCSReferenceValue" || ((TextBox)sender).Name.Trim() == "teLStop_RCSReferenceValue")
                {
                    // 정지 Q 기준값 구하기
                    decimal dcmRac_ReferenceValue = 0M, dcmL_RefernceValue = 0M;
                    dcmRac_ReferenceValue = teRacStop_RCSReferenceValue.Text == "" ? 0 : Convert.ToDecimal(teRacStop_RCSReferenceValue.Text.Trim());
                    dcmL_RefernceValue = teLStop_RCSReferenceValue.Text == "" ? 0 : Convert.ToDecimal(teLStop_RCSReferenceValue.Text.Trim());

                    if (dcmRac_ReferenceValue != 0 && dcmL_RefernceValue != 0)
                        teQStop_RCSReferenceValue.Text = SetQReferenceValueCalculation(dcmRac_ReferenceValue, dcmL_RefernceValue);
                    else
                        teQStop_RCSReferenceValue.Text = "0.000";
                }
                else if (((TextBox)sender).Name.Trim() == "teRacUp_RCSReferenceValue" || ((TextBox)sender).Name.Trim() == "teLUp_RCSReferenceValue")
                {
                    // 올림 Q 기준값 구하기
                    decimal dcmRac_ReferenceValue = 0M, dcmL_RefernceValue = 0M;
                    dcmRac_ReferenceValue = teRacUp_RCSReferenceValue.Text == "" ? 0 : Convert.ToDecimal(teRacUp_RCSReferenceValue.Text.Trim());
                    dcmL_RefernceValue = teLUp_RCSReferenceValue.Text == "" ? 0 : Convert.ToDecimal(teLUp_RCSReferenceValue.Text.Trim());

                    if (dcmRac_ReferenceValue != 0 && dcmL_RefernceValue != 0)
                        teQUp_RCSReferenceValue.Text = SetQReferenceValueCalculation(dcmRac_ReferenceValue, dcmL_RefernceValue);
                    else
                        teQUp_RCSReferenceValue.Text = "0.000";
                }
                else if (((TextBox)sender).Name.Trim() == "teRacMove_RCSReferenceValue" || ((TextBox)sender).Name.Trim() == "teLMove_RCSReferenceValue")
                {
                    // 이동 Q 기준값 구하기
                    decimal dcmRac_ReferenceValue = 0M, dcmL_RefernceValue = 0M;
                    dcmRac_ReferenceValue = teRacUp_RCSReferenceValue.Text == "" ? 0 : Convert.ToDecimal(teRacUp_RCSReferenceValue.Text.Trim());
                    dcmL_RefernceValue = teLMove_RCSReferenceValue.Text == "" ? 0 : Convert.ToDecimal(teLMove_RCSReferenceValue.Text.Trim());

                    if (dcmRac_ReferenceValue != 0 && dcmL_RefernceValue != 0)
                        teQMove_RCSReferenceValue.Text = SetQReferenceValueCalculation(dcmRac_ReferenceValue, dcmL_RefernceValue);
                    else
                        teQMove_RCSReferenceValue.Text = "0.000";
                }
            }
            else if (tabControl1.SelectedIndex == 2)
            {
                if (((TextBox)sender).Name.Trim() == "teControlA1Rac_DRPIReferenceValue" || ((TextBox)sender).Name.Trim() == "teControlA1L_DRPIReferenceValue")
                {
                    // 정지 Q 기준값 구하기
                    decimal dcmRac_ReferenceValue = 0M, dcmL_RefernceValue = 0M;
                    dcmRac_ReferenceValue = teControlA1Rac_DRPIReferenceValue.Text == "" ? 0 : Convert.ToDecimal(teControlA1Rac_DRPIReferenceValue.Text.Trim());
                    dcmL_RefernceValue = teControlA1L_DRPIReferenceValue.Text == "" ? 0 : Convert.ToDecimal(teControlA1L_DRPIReferenceValue.Text.Trim());

                    if (dcmRac_ReferenceValue != 0 && dcmL_RefernceValue != 0)
                        teControlA1Q_DRPIReferenceValue.Text = SetQReferenceValueCalculation(dcmRac_ReferenceValue, dcmL_RefernceValue);
                    else
                        teControlA1Q_DRPIReferenceValue.Text = "0.000";
                }
                else if (((TextBox)sender).Name.Trim() == "teControlA2Rac_DRPIReferenceValue" || ((TextBox)sender).Name.Trim() == "teControlA2L_DRPIReferenceValue")
                {
                    // 정지 Q 기준값 구하기
                    decimal dcmRac_ReferenceValue = 0M, dcmL_RefernceValue = 0M;
                    dcmRac_ReferenceValue = teControlA2Rac_DRPIReferenceValue.Text == "" ? 0 : Convert.ToDecimal(teControlA2Rac_DRPIReferenceValue.Text.Trim());
                    dcmL_RefernceValue = teControlA2L_DRPIReferenceValue.Text == "" ? 0 : Convert.ToDecimal(teControlA2L_DRPIReferenceValue.Text.Trim());

                    if (dcmRac_ReferenceValue != 0 && dcmL_RefernceValue != 0)
                        teControlA2Q_DRPIReferenceValue.Text = SetQReferenceValueCalculation(dcmRac_ReferenceValue, dcmL_RefernceValue);
                    else
                        teControlA2Q_DRPIReferenceValue.Text = "0.000";
                }
                else if (((TextBox)sender).Name.Trim() == "teControlARac_DRPIReferenceValue" || ((TextBox)sender).Name.Trim() == "teControlAL_DRPIReferenceValue")
                {
                    // 정지 Q 기준값 구하기
                    decimal dcmRac_ReferenceValue = 0M, dcmL_RefernceValue = 0M;
                    dcmRac_ReferenceValue = teControlARac_DRPIReferenceValue.Text == "" ? 0 : Convert.ToDecimal(teControlARac_DRPIReferenceValue.Text.Trim());
                    dcmL_RefernceValue = teControlAL_DRPIReferenceValue.Text == "" ? 0 : Convert.ToDecimal(teControlAL_DRPIReferenceValue.Text.Trim());

                    if (dcmRac_ReferenceValue != 0 && dcmL_RefernceValue != 0)
                        teControlAQ_DRPIReferenceValue.Text = SetQReferenceValueCalculation(dcmRac_ReferenceValue, dcmL_RefernceValue);
                    else
                        teControlAQ_DRPIReferenceValue.Text = "0.000";
                }
                else if (((TextBox)sender).Name.Trim() == "teControlB1Rac_DRPIReferenceValue" || ((TextBox)sender).Name.Trim() == "teControlB1L_DRPIReferenceValue")
                {
                    // 정지 Q 기준값 구하기
                    decimal dcmRac_ReferenceValue = 0M, dcmL_RefernceValue = 0M;
                    dcmRac_ReferenceValue = teControlB1Rac_DRPIReferenceValue.Text == "" ? 0 : Convert.ToDecimal(teControlB1Rac_DRPIReferenceValue.Text.Trim());
                    dcmL_RefernceValue = teControlB1L_DRPIReferenceValue.Text == "" ? 0 : Convert.ToDecimal(teControlB1L_DRPIReferenceValue.Text.Trim());

                    if (dcmRac_ReferenceValue != 0 && dcmL_RefernceValue != 0)
                        teControlB1Q_DRPIReferenceValue.Text = SetQReferenceValueCalculation(dcmRac_ReferenceValue, dcmL_RefernceValue);
                    else
                        teControlB1Q_DRPIReferenceValue.Text = "0.000";
                }
                else if (((TextBox)sender).Name.Trim() == "teControlB2Rac_DRPIReferenceValue" || ((TextBox)sender).Name.Trim() == "teControlB2L_DRPIReferenceValue")
                {
                    // 정지 Q 기준값 구하기
                    decimal dcmRac_ReferenceValue = 0M, dcmL_RefernceValue = 0M;
                    dcmRac_ReferenceValue = teControlB2Rac_DRPIReferenceValue.Text == "" ? 0 : Convert.ToDecimal(teControlB2Rac_DRPIReferenceValue.Text.Trim());
                    dcmL_RefernceValue = teControlB2L_DRPIReferenceValue.Text == "" ? 0 : Convert.ToDecimal(teControlB2L_DRPIReferenceValue.Text.Trim());

                    if (dcmRac_ReferenceValue != 0 && dcmL_RefernceValue != 0)
                        teControlB2Q_DRPIReferenceValue.Text = SetQReferenceValueCalculation(dcmRac_ReferenceValue, dcmL_RefernceValue);
                    else
                        teControlB2Q_DRPIReferenceValue.Text = "0.000";
                }
                else if (((TextBox)sender).Name.Trim() == "teControlBRac_DRPIReferenceValue" || ((TextBox)sender).Name.Trim() == "teControlBL_DRPIReferenceValue")
                {
                    // 정지 Q 기준값 구하기
                    decimal dcmRac_ReferenceValue = 0M, dcmL_RefernceValue = 0M;
                    dcmRac_ReferenceValue = teControlBRac_DRPIReferenceValue.Text == "" ? 0 : Convert.ToDecimal(teControlBRac_DRPIReferenceValue.Text.Trim());
                    dcmL_RefernceValue = teControlBL_DRPIReferenceValue.Text == "" ? 0 : Convert.ToDecimal(teControlBL_DRPIReferenceValue.Text.Trim());

                    if (dcmRac_ReferenceValue != 0 && dcmL_RefernceValue != 0)
                        teControlBQ_DRPIReferenceValue.Text = SetQReferenceValueCalculation(dcmRac_ReferenceValue, dcmL_RefernceValue);
                    else
                        teControlBQ_DRPIReferenceValue.Text = "0.000";
                }
                else if (((TextBox)sender).Name.Trim() == "teStopA1Rac_DRPIReferenceValue" || ((TextBox)sender).Name.Trim() == "teStopA1L_DRPIReferenceValue")
                {
                    // 정지 Q 기준값 구하기
                    decimal dcmRac_ReferenceValue = 0M, dcmL_RefernceValue = 0M;
                    dcmRac_ReferenceValue = teStopA1Rac_DRPIReferenceValue.Text == "" ? 0 : Convert.ToDecimal(teStopA1Rac_DRPIReferenceValue.Text.Trim());
                    dcmL_RefernceValue = teStopA1L_DRPIReferenceValue.Text == "" ? 0 : Convert.ToDecimal(teStopA1L_DRPIReferenceValue.Text.Trim());

                    if (dcmRac_ReferenceValue != 0 && dcmL_RefernceValue != 0)
                        teStopA1Q_DRPIReferenceValue.Text = SetQReferenceValueCalculation(dcmRac_ReferenceValue, dcmL_RefernceValue);
                    else
                        teStopA1Q_DRPIReferenceValue.Text = "0.000";
                }
                else if (((TextBox)sender).Name.Trim() == "teStopA2Rac_DRPIReferenceValue" || ((TextBox)sender).Name.Trim() == "teStopA2L_DRPIReferenceValue")
                {
                    // 정지 Q 기준값 구하기
                    decimal dcmRac_ReferenceValue = 0M, dcmL_RefernceValue = 0M;
                    dcmRac_ReferenceValue = teStopA2Rac_DRPIReferenceValue.Text == "" ? 0 : Convert.ToDecimal(teStopA2Rac_DRPIReferenceValue.Text.Trim());
                    dcmL_RefernceValue = teStopA2L_DRPIReferenceValue.Text == "" ? 0 : Convert.ToDecimal(teStopA2L_DRPIReferenceValue.Text.Trim());

                    if (dcmRac_ReferenceValue != 0 && dcmL_RefernceValue != 0)
                        teStopA2Q_DRPIReferenceValue.Text = SetQReferenceValueCalculation(dcmRac_ReferenceValue, dcmL_RefernceValue);
                    else
                        teStopA2Q_DRPIReferenceValue.Text = "0.000";
                }
                else if (((TextBox)sender).Name.Trim() == "teStopARac_DRPIReferenceValue" || ((TextBox)sender).Name.Trim() == "teStopAL_DRPIReferenceValue")
                {
                    // 정지 Q 기준값 구하기
                    decimal dcmRac_ReferenceValue = 0M, dcmL_RefernceValue = 0M;
                    dcmRac_ReferenceValue = teStopARac_DRPIReferenceValue.Text == "" ? 0 : Convert.ToDecimal(teStopARac_DRPIReferenceValue.Text.Trim());
                    dcmL_RefernceValue = teStopAL_DRPIReferenceValue.Text == "" ? 0 : Convert.ToDecimal(teStopAL_DRPIReferenceValue.Text.Trim());

                    if (dcmRac_ReferenceValue != 0 && dcmL_RefernceValue != 0)
                        teStopAQ_DRPIReferenceValue.Text = SetQReferenceValueCalculation(dcmRac_ReferenceValue, dcmL_RefernceValue);
                    else
                        teStopAQ_DRPIReferenceValue.Text = "0.000";
                }
                else if (((TextBox)sender).Name.Trim() == "teStopB1Rac_DRPIReferenceValue" || ((TextBox)sender).Name.Trim() == "teStopB1L_DRPIReferenceValue")
                {
                    // 정지 Q 기준값 구하기
                    decimal dcmRac_ReferenceValue = 0M, dcmL_RefernceValue = 0M;
                    dcmRac_ReferenceValue = teStopB1Rac_DRPIReferenceValue.Text == "" ? 0 : Convert.ToDecimal(teStopB1Rac_DRPIReferenceValue.Text.Trim());
                    dcmL_RefernceValue = teStopB1L_DRPIReferenceValue.Text == "" ? 0 : Convert.ToDecimal(teStopB1L_DRPIReferenceValue.Text.Trim());

                    if (dcmRac_ReferenceValue != 0 && dcmL_RefernceValue != 0)
                        teStopB1Q_DRPIReferenceValue.Text = SetQReferenceValueCalculation(dcmRac_ReferenceValue, dcmL_RefernceValue);
                    else 
                        teStopB1Q_DRPIReferenceValue.Text = "0.000";
                }
                else if (((TextBox)sender).Name.Trim() == "teStopB2Rac_DRPIReferenceValue" || ((TextBox)sender).Name.Trim() == "teStopB2L_DRPIReferenceValue")
                {
                    // 정지 Q 기준값 구하기
                    decimal dcmRac_ReferenceValue = 0M, dcmL_RefernceValue = 0M;
                    dcmRac_ReferenceValue = teStopB2Rac_DRPIReferenceValue.Text == "" ? 0 : Convert.ToDecimal(teStopB2Rac_DRPIReferenceValue.Text.Trim());
                    dcmL_RefernceValue = teStopB2L_DRPIReferenceValue.Text == "" ? 0 : Convert.ToDecimal(teStopB2L_DRPIReferenceValue.Text.Trim());

                    if (dcmRac_ReferenceValue != 0 && dcmL_RefernceValue != 0)
                        teStopB2Q_DRPIReferenceValue.Text = SetQReferenceValueCalculation(dcmRac_ReferenceValue, dcmL_RefernceValue);
                    else
                        teStopB2Q_DRPIReferenceValue.Text = "0.000";
                }
                else if (((TextBox)sender).Name.Trim() == "teStopBRac_DRPIReferenceValue" || ((TextBox)sender).Name.Trim() == "teStopBL_DRPIReferenceValue")
                {
                    // 정지 Q 기준값 구하기
                    decimal dcmRac_ReferenceValue = 0M, dcmL_RefernceValue = 0M;
                    dcmRac_ReferenceValue = teStopBRac_DRPIReferenceValue.Text == "" ? 0 : Convert.ToDecimal(teStopBRac_DRPIReferenceValue.Text.Trim());
                    dcmL_RefernceValue = teStopBL_DRPIReferenceValue.Text == "" ? 0 : Convert.ToDecimal(teStopBL_DRPIReferenceValue.Text.Trim());

                    if (dcmRac_ReferenceValue != 0 && dcmL_RefernceValue != 0)
                        teStopBQ_DRPIReferenceValue.Text = SetQReferenceValueCalculation(dcmRac_ReferenceValue, dcmL_RefernceValue);
                    else 
                        teStopBQ_DRPIReferenceValue.Text = "0.000";
                }
            }
             * */
        }

        /// <summary>
        /// Q 기준값 구하기
        /// </summary>
        /// <param name="_dRdc_ReferenceValue"></param>
        /// <param name="_dL_RefernceValue"></param>
        /// <returns></returns>
        private string SetQReferenceValueCalculation(decimal _dcmRac_ReferenceValue, decimal _dcmL_RefernceValue)
        {
            string strResult = "";
            decimal dcmQ_ReferenceValue = 0M;

            if (rboMeasurementType1.Checked)
            {
                dcmQ_ReferenceValue = (2M * 3.141592M * 120M * (_dcmL_RefernceValue / 1000)) / _dcmRac_ReferenceValue;
            }
            else
            {
                dcmQ_ReferenceValue = (2M * 3.141592M * 120M * (_dcmL_RefernceValue / 1000)) / _dcmRac_ReferenceValue;
            }

            strResult = (Math.Truncate(dcmQ_ReferenceValue * 1000) / 1000).ToString().Trim();

            return strResult;
        }

        /// <summary>
        /// TextBox Leave Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TextBox_C_Leave(object sender, EventArgs e)
        {
            if (((TextBox)sender).Text.ToString().Trim() == "")
            {
                ((TextBox)sender).Text = "0.000000";
            }
        }

        /// <summary>
        /// RCS or DRPI 호기 ComboBox 변경 이벤트
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cboHogi_SelectedIndexChanged(object sender, EventArgs e)
        {
            string strSelectHogi = "초기값";
            
            // 호기별 OH 차수 설정
            DataTable dt = new DataTable();

            if (((ComboBox)sender).Name.Trim() == "cboHogi_RCS")
            {
                strSelectHogi = cboHogi_RCS.SelectedItem == null ? "초기값" : cboHogi_RCS.SelectedItem.ToString().Trim();

                dt = m_db.GetRCSReferenceValueOhDegreeInfo(strPlantName, strSelectHogi);

                cboOhDegree_RCS.Items.Clear();

                // 초기 값 추가
                if (strSelectHogi.Trim() == "초기값")
                {
                    cboOhDegree_RCS.Items.Add("제 0 차");
                }
                else
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        if (dt.Rows[i]["Oh_Degree"] != null)
                        {
                            cboOhDegree_RCS.Items.Add(dt.Rows[i]["Oh_Degree"].ToString().Trim());
                        }
                    }
                }

                cboOhDegree_RCS.SelectedIndex = cboOhDegree_RCS.Items.Count - 1;

                if (strSelectHogi.Trim() == "초기값")
                    cboOhDegree_RCS.Enabled = false;
                else
                    cboOhDegree_RCS.Enabled = true;

                if (cboOhDegree_RCS.Items.Count < 1)
                {
                    // Control - 기준 값(RCS) 초기화
                    SetControlInitialize2();

                    // 환경설정 - 기준 값(RCS) 가져오기
                    GetTabControl2ValueChange();
                }
            }
            else if (((ComboBox)sender).Name.Trim() == "cboHogi_DRPI")
            {
                strSelectHogi = cboHogi_DRPI.SelectedItem == null ? "초기값" : cboHogi_DRPI.SelectedItem.ToString().Trim();

                dt = m_db.GetDRPIReferenceValueOhDegreeInfo(strPlantName, strSelectHogi);

                cboOhDegree_DRPI.Items.Clear();

                // 초기 값 추가
                if (strSelectHogi.Trim() == "초기값")
                {
                    cboOhDegree_DRPI.Items.Add("제 0 차");
                }
                else
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        if (dt.Rows[i]["Oh_Degree"] != null)
                        {
                            cboOhDegree_DRPI.Items.Add(dt.Rows[i]["Oh_Degree"].ToString().Trim());
                        }
                    }
                }

                cboOhDegree_DRPI.SelectedIndex = cboOhDegree_DRPI.Items.Count - 1;

                if (strSelectHogi.Trim() == "초기값")
                    cboOhDegree_DRPI.Enabled = false;
                else
                    cboOhDegree_DRPI.Enabled = true;

                if (cboOhDegree_DRPI.Items.Count < 1)
                {
                    // Control - 기준 값(DRPI) 초기화
                    SetControlInitialize3();

                    // 환경설정 - 기준 값(DRPI) 가져오기
                    GetTabControl3ValueChange();
                }
            }
        }

        /// <summary>
        /// RCS or DRPI OH 차수 ComboBox 변경 이벤트
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cboOhDegree_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (((ComboBox)sender).Name.Trim() == "cboOhDegree_RCS")
            {
                // Control - 기준 값(RCS) 초기화
                SetControlInitialize2();

                // 환경설정 - 기준 값(RCS) 가져오기
                GetTabControl2ValueChange();
            }
            else if (((ComboBox)sender).Name.Trim() == "cboOhDegree_DRPI")
            {
                // Control - 기준 값(DRPI) 초기화
                SetControlInitialize3();

                // 환경설정 - 기준 값(DRPI) 가져오기
                GetTabControl3ValueChange();
            }
        }

        /// <summary>
        /// 통신 확인
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCommunication_Click(object sender, EventArgs e)
        {
            bool boolDAQ = false;
            bool boolLCRMeter = false;

            try
            {
                // DAQ and LCR-Meter 
                string strDAQDeviceName = cboDAQDevice.SelectedItem == null ? "" : cboDAQDevice.SelectedItem.ToString().Trim();
                m_MeasureProcess.DaqAndLCRMeterCommunication(strDAQDeviceName.Trim(), teLCRMeter_Addr.Text.Trim(), ref boolDAQ, ref boolLCRMeter);

                //ledDAQ.Value = boolDAQ;
                //ledLCRMeter.Value = boolLCRMeter;
                ledDAQ.On = boolDAQ;
                ledLCRMeter.On = boolLCRMeter;

                // TDR-Meter (이더넷 연결 체크)
                if (pnlTDRMeterSetting.Visible == true)
                {
                    System.Net.NetworkInformation.Ping ping = new System.Net.NetworkInformation.Ping();
                    System.Net.NetworkInformation.PingOptions options = new System.Net.NetworkInformation.PingOptions();
                    options.DontFragment = true;
                    string data = "aaaaaaaaaaaaaaaaa";
                    byte[] buffer = ASCIIEncoding.ASCII.GetBytes(data);
                    int timeout = 120;
                    System.Net.NetworkInformation.PingReply reply = ping.Send(IPAddress.Parse(teTDRMeter_IPAddress.Text.Trim()), timeout, buffer, options);

                    if (reply.Status == System.Net.NetworkInformation.IPStatus.Success)
                    {
                        // 네트워크 사용 가능할 때~~
                        //ledTDRMeter.Value = true;
                        ledTDRMeter.On = true;
                    }
                    else
                    {
                        //ledTDRMeter.Value = false;
                        ledTDRMeter.On = false;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
            finally
            {
                if (m_MeasureProcess.m_LCRMeter.m_bConnected) m_MeasureProcess.m_LCRMeter.Close();
            }
        }

        /// <summary>
        /// NumericUpDown Leave Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void NumericUpDown_Leave(object sender, EventArgs e)
        {
            if (((NumericUpDown)sender).Text.ToString().Trim() == "")
            {
                ((NumericUpDown)sender).Value = 0.0M;
                ((NumericUpDown)sender).Text = "0.0";
            }
        }

        /// <summary>
        /// 판단범위 설정(Rdc, Rac, L, C, Q 일괄 설정)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void teRdcDecisionRange_ReferenceValue_TextChanged(object sender, EventArgs e)
        {
            if (teRdcDecisionRange_ReferenceValue.Text.Trim() == "" 
                || teRdcDecisionRange_ReferenceValue.Text.Trim().Length != Regex.Replace(teRdcDecisionRange_ReferenceValue.Text.Trim(), @"\D", "").Length)
            {
                MessageBox.Show("판단범위에 공백 또는 문자를 입력할 수 없습니다.");
                teRdcDecisionRange_ReferenceValue.Text = "10";
            }
        }

        private void teRdcDecisionRange_ReferenceValue_Leave(object sender, EventArgs e)
        {
            teRacDecisionRange_ReferenceValue.Text = teRdcDecisionRange_ReferenceValue.Text.Trim();
            teLDecisionRange_ReferenceValue.Text = teRdcDecisionRange_ReferenceValue.Text.Trim();
            teCDecisionRange_ReferenceValue.Text = teRdcDecisionRange_ReferenceValue.Text.Trim();
            teQDecisionRange_ReferenceValue.Text = teRdcDecisionRange_ReferenceValue.Text.Trim();
        }
    }
}
