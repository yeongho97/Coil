using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using Coil_Diagnostor.Function;

namespace Coil_Diagnostor
{
    public partial class frmSelfTest : Form
    {
        frmMain frmMain;
        protected Function.FunctionDataControl m_db = new Function.FunctionDataControl();
        protected frmMessageBox frmMB = new frmMessageBox();
        protected string strPlantName = "";
        protected bool boolMessageOK = false;
        protected bool boolFormLoad = false;
        protected bool boolMeasurement = false;

        public Thread _threadMeasurementStartState;
        public Thread _thread;

        public frmSelfTest()
        {
            CheckForIllegalCrossThreadCalls = false;
            InitializeComponent();
        }

        public frmSelfTest(frmMain _frm)
        {
            CheckForIllegalCrossThreadCalls = false;
            InitializeComponent();

            frmMain = _frm;
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
                    btnInitialize.PerformClick();
                    break;
                case Keys.F3: // 점검 버튼
                    btnCheck.PerformClick();
                    break;
                case Keys.F4: // 정지 버튼
                    btnStop.PerformClick();
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
        private void frmSelfTest_Load(object sender, EventArgs e)
        {
            boolFormLoad = true;

            strPlantName = Gini.GetValue("Device", "PlantName").Trim();

            // 초기화
            SetControlInitialize();

            // LCR - Meter 통신 연결
            if (frmMain.m_MeasureProcess.m_LCRMeter.Initialize(Gini.GetValue("Device", "LCRMeter_Addr").Trim()))
            {
                // Button Enabled 설정
                SetButtonEnabled(true);
            }
            else
            {
                btnCheck.Enabled = false;
                btnStop.Enabled = false;
            }

            boolFormLoad = false;
        }

        private void frmSelfTest_FormClosing(object sender, FormClosingEventArgs e)
        {
            frmMain.m_MeasureProcess.m_LCRMeter.Close();
        }

        /// <summary>
        /// Button Enabled 설정
        /// </summary>
        /// <param name="_boolEnabled"></param>
        private void SetButtonEnabled(bool _boolStop)
        {
            if (_boolStop)
            {
                btnInitialize.Enabled = true;       // 초기화
                btnCheck.Enabled = true;            // 점검
                btnStop.Enabled = false;            // 정지
                btnClose.Enabled = true;            // 닫기
            }
            else
            {
                btnInitialize.Enabled = false;      // 초기화
                btnCheck.Enabled = false;           // 점검
                btnStop.Enabled = true;             // 정지
                btnClose.Enabled = false;           // 닫기
            }
        }

        /// <summary>
        /// 초기화
        /// </summary>
        private void SetControlInitialize()
        {
            try
            {
                teNormalResistance.Text = "";
                teWheatstoneResistance.Text = "";

                foreach (Control c in groupBox3.Controls)
                {
                    if (c.GetType().Name.Trim() == "GroupBox")
                    {
                        switch (c.Text.Trim())
                        {
                            case "DAM 1":
                                foreach (Control c_btn in gbDAM1.Controls)
                                {
                                    if (c_btn.GetType().Name.Trim() == "Button")
                                    {
                                        c_btn.BackColor = System.Drawing.Color.Gainsboro;
                                    }
                                }
                                break;
                            case "DAM 2":
                                foreach (Control c_btn in gbDAM2.Controls)
                                {
                                    if (c_btn.GetType().Name.Trim() == "Button")
                                    {
                                        c_btn.BackColor = System.Drawing.Color.Gainsboro;
                                    }
                                }
                                break;
                            case "DAM 3":
                                foreach (Control c_btn in gbDAM3.Controls)
                                {
                                    if (c_btn.GetType().Name.Trim() == "Button")
                                    {
                                        c_btn.BackColor = System.Drawing.Color.Gainsboro;
                                    }
                                }
                                break;
                            case "DAM 4":
                                foreach (Control c_btn in gbDAM4.Controls)
                                {
                                    if (c_btn.GetType().Name.Trim() == "Button")
                                    {
                                        c_btn.BackColor = System.Drawing.Color.Gainsboro;
                                    }
                                }
                                break;
                            case "DAM 5":
                                foreach (Control c_btn in gbDAM5.Controls)
                                {
                                    if (c_btn.GetType().Name.Trim() == "Button")
                                    {
                                        c_btn.BackColor = System.Drawing.Color.Gainsboro;
                                    }
                                }
                                break;
                            case "DAM 6":
                                foreach (Control c_btn in gbDAM6.Controls)
                                {
                                    if (c_btn.GetType().Name.Trim() == "Button")
                                    {
                                        c_btn.BackColor = System.Drawing.Color.Gainsboro;
                                    }
                                }
                                break;
                            case "DAM 7":
                                foreach (Control c_btn in gbDAM7.Controls)
                                {
                                    if (c_btn.GetType().Name.Trim() == "Button")
                                    {
                                        c_btn.BackColor = System.Drawing.Color.Gainsboro;
                                    }
                                }
                                break;
                            case "DAM 8":
                                foreach (Control c_btn in gbDAM8.Controls)
                                {
                                    if (c_btn.GetType().Name.Trim() == "Button")
                                    {
                                        c_btn.BackColor = System.Drawing.Color.Gainsboro;
                                    }
                                }
                                break;
                            case "DAM 9":
                                foreach (Control c_btn in gbDAM9.Controls)
                                {
                                    if (c_btn.GetType().Name.Trim() == "Button")
                                    {
                                        c_btn.BackColor = System.Drawing.Color.Gainsboro;
                                    }
                                }
                                break;
                            case "DAM 10":
                                foreach (Control c_btn in gbDAM10.Controls)
                                {
                                    if (c_btn.GetType().Name.Trim() == "Button")
                                    {
                                        c_btn.BackColor = System.Drawing.Color.Gainsboro;
                                    }
                                }
                                break;
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
        /// 닫기 Button Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnClose_Click(object sender, EventArgs e)
        {
            if (!boolMeasurement)
                this.Close();
            else
            {
                frmMB.lblMessage.Text = "자체 점검 중이므로 점검 종료 또는 정지 후 닫기하십시오.";
                frmMB.TopMost = true;
                frmMB.ShowDialog();
            }
        }

        /// <summary>
        /// 정지 Button Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnStop_Click(object sender, EventArgs e)
        {
            try
            {
                boolMeasurement = false;
                Thread.Sleep(1000);

                if (_thread != null)
                    _thread.Abort();

                // Button Enabled 설정
                SetButtonEnabled(true);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
        }

        /// <summary>
        /// 점검 Button Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCheck_Click(object sender, EventArgs e)
        {
            try
            {
                boolMeasurement = true;

                // Button Enabled 설정
                SetButtonEnabled(false);

                // 초기화                
                SetControlInitialize();

                // 측정 시작
                _thread = new Thread(new ThreadStart(threadSelfTestMeasurementStart));
                _thread.Start();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
                boolMeasurement = false;
            }
        }

        /// <summary>
        /// Self Test Measurement Start Thread
        /// </summary>
        public void threadSelfTestMeasurementStart()
        {
            bool boolResult = false;
            string strDAQPinMap = "", strDAQDIOPinMap = "", strDAQAIPinMap = "";
            string strDAQDeviceName = Gini.GetValue("Device", "DAQDeviceName").Trim();
            int intSleep = 30;

            try
            {
                strDAQDIOPinMap = string.Format("{0}/port2/line0", strDAQDeviceName);
                strDAQAIPinMap = string.Format("{0}/ai1", strDAQDeviceName);

                #region Self Test DAM 1

                if (chkDAM1.Checked)
                {
                    // DAM 1 - Channel 1
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line1,{0}/port0/line23,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM1_Ch1, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 1 - Channel 2
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line2,{0}/port0/line23,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM1_Ch2, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 1 - Channel 3
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line3,{0}/port0/line23,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM1_Ch3, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 1 - Channel 4
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line4,{0}/port0/line23,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM1_Ch4, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 1 - Channel 5
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line5,{0}/port0/line23,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM1_Ch5, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 1 - Channel 6
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line6,{0}/port0/line23,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM1_Ch6, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 1 - Channel 7
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line7,{0}/port0/line23,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM1_Ch7, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 1 - Channel 8
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line8,{0}/port0/line23,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM1_Ch8, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 1 - Channel 9
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line9,{0}/port0/line23,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM1_Ch9, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 1 - Channel 10
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line10,{0}/port0/line23,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM1_Ch10, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 1 - Channel 11
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line11,{0}/port0/line23,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM1_Ch11, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 1 - Channel 12
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line12,{0}/port0/line23,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM1_Ch12, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 1 - Channel 13
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line13,{0}/port0/line23,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM1_Ch13, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 1 - Channel 14
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line14,{0}/port0/line23,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM1_Ch14, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 1 - Channel 15
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line15,{0}/port0/line23,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM1_Ch15, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 1 - Channel 16
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line16,{0}/port0/line23,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM1_Ch16, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 1 - Channel 17
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line17,{0}/port0/line23,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM1_Ch17, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 1 - Channel 18
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line18,{0}/port0/line23,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM1_Ch18, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 1 - Channel 19
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line19,{0}/port0/line23,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM1_Ch19, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 1 - Channel 20
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line20,{0}/port0/line23,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM1_Ch20, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 1 - Channel 21
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line21,{0}/port0/line23,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM1_Ch21, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }
                }

                #endregion

                #region Self Test DAM 2

                if (chkDAM2.Checked)
                {
                    // DAM 2 - Channel 1
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line1,{0}/port0/line24,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM2_Ch1, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 2 - Channel 2
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line2,{0}/port0/line24,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM2_Ch2, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 2 - Channel 3
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line3,{0}/port0/line24,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM2_Ch3, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 2 - Channel 4
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line4,{0}/port0/line24,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM2_Ch4, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 2 - Channel 5
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line5,{0}/port0/line24,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM2_Ch5, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 2 - Channel 6
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line6,{0}/port0/line24,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM2_Ch6, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 2 - Channel 7
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line7,{0}/port0/line24,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM2_Ch7, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 2 - Channel 8
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line8,{0}/port0/line24,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM2_Ch8, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 2 - Channel 9
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line9,{0}/port0/line24,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM2_Ch9, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 2 - Channel 10
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line10,{0}/port0/line24,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM2_Ch10, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 2 - Channel 11
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line11,{0}/port0/line24,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM2_Ch11, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 2 - Channel 12
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line12,{0}/port0/line24,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM2_Ch12, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 2 - Channel 13
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line13,{0}/port0/line24,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM2_Ch13, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 2 - Channel 14
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line14,{0}/port0/line24,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM2_Ch14, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 2 - Channel 15
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line15,{0}/port0/line24,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM2_Ch15, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 2 - Channel 16
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line16,{0}/port0/line24,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM2_Ch16, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 2 - Channel 17
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line17,{0}/port0/line24,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM2_Ch17, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 2 - Channel 18
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line18,{0}/port0/line24,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM2_Ch18, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 2 - Channel 19
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line19,{0}/port0/line24,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM2_Ch19, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 2 - Channel 20
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line20,{0}/port0/line24,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM2_Ch20, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 2 - Channel 21
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line21,{0}/port0/line24,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM2_Ch21, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }
                }

                #endregion

                #region Self Test DAM 3

                if (chkDAM3.Checked)
                {
                    // DAM 3 - Channel 1
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line1,{0}/port0/line25,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM3_Ch1, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 3 - Channel 2
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line2,{0}/port0/line25,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM3_Ch2, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 3 - Channel 3
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line3,{0}/port0/line25,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM3_Ch3, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 3 - Channel 4
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line4,{0}/port0/line25,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM3_Ch4, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 3 - Channel 5
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line5,{0}/port0/line25,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM3_Ch5, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 3 - Channel 6
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line6,{0}/port0/line25,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM3_Ch6, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 3 - Channel 7
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line7,{0}/port0/line25,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM3_Ch7, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 3 - Channel 8
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line8,{0}/port0/line25,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM3_Ch8, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 3 - Channel 9
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line9,{0}/port0/line25,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM3_Ch9, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 3 - Channel 10
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line10,{0}/port0/line25,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM3_Ch10, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 3 - Channel 11
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line11,{0}/port0/line25,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM3_Ch11, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 3 - Channel 12
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line12,{0}/port0/line25,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM3_Ch12, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 3 - Channel 13
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line13,{0}/port0/line25,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM3_Ch13, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 3 - Channel 14
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line14,{0}/port0/line25,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM3_Ch14, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 3 - Channel 15
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line15,{0}/port0/line25,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM3_Ch15, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 3 - Channel 16
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line16,{0}/port0/line25,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM3_Ch16, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 3 - Channel 17
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line17,{0}/port0/line25,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM3_Ch17, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 3 - Channel 18
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line18,{0}/port0/line25,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM3_Ch18, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 3 - Channel 19
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line19,{0}/port0/line25,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM3_Ch19, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 3 - Channel 20
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line20,{0}/port0/line25,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM3_Ch20, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 3 - Channel 21
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line21,{0}/port0/line25,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM3_Ch21, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }
                }

                #endregion

                #region Self Test DAM 4

                if (chkDAM4.Checked)
                {
                    // DAM 4 - Channel 1
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line1,{0}/port0/line26,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM4_Ch1, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 4 - Channel 2
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line2,{0}/port0/line26,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM4_Ch2, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 4 - Channel 3
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line3,{0}/port0/line26,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM4_Ch3, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 4 - Channel 4
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line4,{0}/port0/line26,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM4_Ch4, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 4 - Channel 5
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line5,{0}/port0/line26,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM4_Ch5, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 4 - Channel 6
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line6,{0}/port0/line26,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM4_Ch6, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 4 - Channel 7
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line7,{0}/port0/line26,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM4_Ch7, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 4 - Channel 8
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line8,{0}/port0/line26,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM4_Ch8, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 4 - Channel 9
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line9,{0}/port0/line26,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM4_Ch9, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 4 - Channel 10
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line10,{0}/port0/line26,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM4_Ch10, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 4 - Channel 11
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line11,{0}/port0/line26,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM4_Ch11, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 4 - Channel 12
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line12,{0}/port0/line26,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM4_Ch12, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 4 - Channel 13
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line13,{0}/port0/line26,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM4_Ch13, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 4 - Channel 14
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line14,{0}/port0/line26,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM4_Ch14, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 4 - Channel 15
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line15,{0}/port0/line26,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM4_Ch15, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 4 - Channel 16
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line16,{0}/port0/line26,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM4_Ch16, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 4 - Channel 17
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line17,{0}/port0/line26,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM4_Ch17, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 4 - Channel 18
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line18,{0}/port0/line26,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM4_Ch18, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 4 - Channel 19
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line19,{0}/port0/line26,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM4_Ch19, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 4 - Channel 20
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line20,{0}/port0/line26,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM4_Ch20, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 4 - Channel 21
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line21,{0}/port0/line26,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM4_Ch21, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }
                }

                #endregion

                #region Self Test DAM 5

                if (chkDAM5.Checked)
                {
                    // DAM 5 - Channel 1
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line1,{0}/port0/line27,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM5_Ch1, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 5 - Channel 2
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line2,{0}/port0/line27,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM5_Ch2, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 5 - Channel 3
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line3,{0}/port0/line27,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM5_Ch3, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 5 - Channel 4
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line4,{0}/port0/line27,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM5_Ch4, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 5 - Channel 5
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line5,{0}/port0/line27,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM5_Ch5, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 5 - Channel 6
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line6,{0}/port0/line27,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM5_Ch6, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 5 - Channel 7
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line7,{0}/port0/line27,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM5_Ch7, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 5 - Channel 8
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line8,{0}/port0/line27,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM5_Ch8, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 5 - Channel 9
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line9,{0}/port0/line27,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM5_Ch9, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 5 - Channel 10
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line10,{0}/port0/line27,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM5_Ch10, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 5 - Channel 11
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line11,{0}/port0/line27,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM5_Ch11, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 5 - Channel 12
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line12,{0}/port0/line27,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM5_Ch12, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 5 - Channel 13
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line13,{0}/port0/line27,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM5_Ch13, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 5 - Channel 14
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line14,{0}/port0/line27,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM5_Ch14, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 5 - Channel 15
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line15,{0}/port0/line27,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM5_Ch15, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 5 - Channel 16
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line16,{0}/port0/line27,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM5_Ch16, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 5 - Channel 17
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line17,{0}/port0/line27,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM5_Ch17, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 5 - Channel 18
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line18,{0}/port0/line27,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM5_Ch18, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 5 - Channel 19
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line19,{0}/port0/line27,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM5_Ch19, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 5 - Channel 20
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line20,{0}/port0/line27,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM5_Ch20, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 5 - Channel 21
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line21,{0}/port0/line27,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM5_Ch21, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }
                }

                #endregion

                #region Self Test DAM 6

                if (chkDAM6.Checked)
                {
                    // DAM 6 - Channel 1
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line1,{0}/port0/line28,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM6_Ch1, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 6 - Channel 2
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line2,{0}/port0/line28,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM6_Ch2, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 6 - Channel 3
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line3,{0}/port0/line28,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM6_Ch3, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 6 - Channel 4
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line4,{0}/port0/line28,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM6_Ch4, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 6 - Channel 5
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line5,{0}/port0/line28,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM6_Ch5, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 6 - Channel 6
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line6,{0}/port0/line28,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM6_Ch6, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 6 - Channel 7
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line7,{0}/port0/line28,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM6_Ch7, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 6 - Channel 8
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line8,{0}/port0/line28,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM6_Ch8, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 6 - Channel 9
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line9,{0}/port0/line28,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM6_Ch9, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 6 - Channel 10
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line10,{0}/port0/line28,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM6_Ch10, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 6 - Channel 11
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line11,{0}/port0/line28,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM6_Ch11, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 6 - Channel 12
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line12,{0}/port0/line28,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM6_Ch12, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 6 - Channel 13
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line13,{0}/port0/line28,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM6_Ch13, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 6 - Channel 14
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line14,{0}/port0/line28,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM6_Ch14, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 6 - Channel 15
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line15,{0}/port0/line28,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM6_Ch15, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 6 - Channel 16
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line16,{0}/port0/line28,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM6_Ch16, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 6 - Channel 17
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line17,{0}/port0/line28,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM6_Ch17, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 6 - Channel 18
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line18,{0}/port0/line28,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM6_Ch18, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 6 - Channel 19
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line19,{0}/port0/line28,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM6_Ch19, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 6 - Channel 20
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line20,{0}/port0/line28,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM6_Ch20, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 6 - Channel 21
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line21,{0}/port0/line28,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM6_Ch21, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }
                }

                #endregion

                #region Self Test DAM 7

                if (chkDAM7.Checked)
                {
                    // DAM 7 - Channel 1
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line1,{0}/port0/line29,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM7_Ch1, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 7 - Channel 2
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line2,{0}/port0/line29,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM7_Ch2, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 7 - Channel 3
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line3,{0}/port0/line29,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM7_Ch3, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 7 - Channel 4
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line4,{0}/port0/line29,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM7_Ch4, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 7 - Channel 5
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line5,{0}/port0/line29,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM7_Ch5, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 7 - Channel 6
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line6,{0}/port0/line29,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM7_Ch6, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 7 - Channel 7
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line7,{0}/port0/line29,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM7_Ch7, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 7 - Channel 8
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line8,{0}/port0/line29,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM7_Ch8, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 7 - Channel 9
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line9,{0}/port0/line29,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM7_Ch9, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 7 - Channel 10
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line10,{0}/port0/line29,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM7_Ch10, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 7 - Channel 11
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line11,{0}/port0/line29,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM7_Ch11, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 7 - Channel 12
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line12,{0}/port0/line29,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM7_Ch12, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 7 - Channel 13
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line13,{0}/port0/line29,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM7_Ch13, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 7 - Channel 14
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line14,{0}/port0/line29,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM7_Ch14, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 7 - Channel 15
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line15,{0}/port0/line29,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM7_Ch15, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 7 - Channel 16
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line16,{0}/port0/line29,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM7_Ch16, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 7 - Channel 17
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line17,{0}/port0/line29,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM7_Ch17, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 7 - Channel 18
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line18,{0}/port0/line29,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM7_Ch18, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 7 - Channel 19
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line19,{0}/port0/line29,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM7_Ch19, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 7 - Channel 20
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line20,{0}/port0/line29,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM7_Ch20, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 7 - Channel 21
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line21,{0}/port0/line29,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM7_Ch21, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }
                }

                #endregion

                #region Self Test DAM 8

                if (chkDAM8.Checked)
                {
                    // DAM 8 - Channel 1
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line1,{0}/port0/line30,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM8_Ch1, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 8 - Channel 2
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line2,{0}/port0/line30,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM8_Ch2, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 8 - Channel 3
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line3,{0}/port0/line30,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM8_Ch3, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 8 - Channel 4
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line4,{0}/port0/line30,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM8_Ch4, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 8 - Channel 5
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line5,{0}/port0/line30,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM8_Ch5, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 8 - Channel 6
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line6,{0}/port0/line30,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM8_Ch6, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 8 - Channel 7
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line7,{0}/port0/line30,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM8_Ch7, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 8 - Channel 8
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line8,{0}/port0/line30,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM8_Ch8, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 8 - Channel 9
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line9,{0}/port0/line30,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM8_Ch9, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 8 - Channel 10
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line10,{0}/port0/line30,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM8_Ch10, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 8 - Channel 11
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line11,{0}/port0/line30,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM8_Ch11, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 8 - Channel 12
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line12,{0}/port0/line30,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM8_Ch12, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 8 - Channel 13
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line13,{0}/port0/line30,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM8_Ch13, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 8 - Channel 14
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line14,{0}/port0/line30,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM8_Ch14, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 8 - Channel 15
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line15,{0}/port0/line30,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM8_Ch15, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 8 - Channel 16
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line16,{0}/port0/line30,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM8_Ch16, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 8 - Channel 17
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line17,{0}/port0/line30,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM8_Ch17, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 8 - Channel 18
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line18,{0}/port0/line30,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM8_Ch18, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 8 - Channel 19
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line19,{0}/port0/line30,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM8_Ch19, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 8 - Channel 20
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line20,{0}/port0/line30,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM8_Ch20, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 8 - Channel 21
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line21,{0}/port0/line30,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM8_Ch21, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }
                }

                #endregion

                #region Self Test DAM 9

                if (chkDAM9.Checked)
                {
                    // DAM 9 - Channel 1
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line1,{0}/port0/line31,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM9_Ch1, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 9 - Channel 2
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line2,{0}/port0/line31,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM9_Ch2, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 9 - Channel 3
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line3,{0}/port0/line31,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM9_Ch3, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 9 - Channel 4
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line4,{0}/port0/line31,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM9_Ch4, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 9 - Channel 5
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line5,{0}/port0/line31,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM9_Ch5, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 9 - Channel 6
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line6,{0}/port0/line31,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM9_Ch6, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 9 - Channel 7
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line7,{0}/port0/line31,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM9_Ch7, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 9 - Channel 8
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line8,{0}/port0/line31,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM9_Ch8, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 9 - Channel 9
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line9,{0}/port0/line31,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM9_Ch9, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 9 - Channel 10
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line10,{0}/port0/line31,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM9_Ch10, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 9 - Channel 11
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line11,{0}/port0/line31,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM9_Ch11, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 9 - Channel 12
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line12,{0}/port0/line31,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM9_Ch12, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 9 - Channel 13
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line13,{0}/port0/line31,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM9_Ch13, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 9 - Channel 14
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line14,{0}/port0/line31,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM9_Ch14, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 9 - Channel 15
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line15,{0}/port0/line31,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM9_Ch15, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 9 - Channel 16
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line16,{0}/port0/line31,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM9_Ch16, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 9 - Channel 17
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line17,{0}/port0/line31,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM9_Ch17, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 9 - Channel 18
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line18,{0}/port0/line31,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM9_Ch18, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 9 - Channel 19
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line19,{0}/port0/line31,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM9_Ch19, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 9 - Channel 20
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line20,{0}/port0/line31,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM9_Ch20, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 9 - Channel 21
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line21,{0}/port0/line31,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM9_Ch21, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }
                }

                #endregion

                #region Self Test DAM 10

                if (chkDAM10.Checked)
                {
                    // DAM 10 - Channel 1
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line1,{0}/port2/line1,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM10_Ch1, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 10 - Channel 2
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line2,{0}/port2/line1,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM10_Ch2, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 10 - Channel 3
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line3,{0}/port2/line1,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM10_Ch3, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 10 - Channel 4
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line4,{0}/port2/line1,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM10_Ch4, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 10 - Channel 5
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line5,{0}/port2/line1,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM10_Ch5, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 10 - Channel 6
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line6,{0}/port2/line1,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM10_Ch6, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 10 - Channel 7
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line7,{0}/port2/line1,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM10_Ch7, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 10 - Channel 8
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line8,{0}/port2/line1,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM10_Ch8, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 10 - Channel 9
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line9,{0}/port2/line1,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM10_Ch9, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 10 - Channel 10
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line10,{0}/port2/line1,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM10_Ch10, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 10 - Channel 11
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line11,{0}/port2/line1,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM10_Ch11, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 10 - Channel 12
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line12,{0}/port2/line1,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM10_Ch12, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 10 - Channel 13
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line13,{0}/port2/line1,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM10_Ch13, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 1 - Channel 14
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line14,{0}/port2/line1,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM10_Ch14, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 10 - Channel 15
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line15,{0}/port2/line1,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM10_Ch15, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 10 - Channel 16
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line16,{0}/port2/line1,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM10_Ch16, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 10 - Channel 17
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line17,{0}/port2/line1,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM10_Ch17, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 10 - Channel 18
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line18,{0}/port2/line1,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM10_Ch18, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 10 - Channel 19
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line19,{0}/port2/line1,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM10_Ch19, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 10 - Channel 20
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line20,{0}/port2/line1,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM10_Ch20, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }

                    // DAM 10 - Channel 21
                    if (boolMeasurement)
                    {
                        strDAQPinMap = string.Format("{0}/port0/line0,{0}/port0/line21,{0}/port2/line1,{0}/port2/line0", strDAQDeviceName);
                        boolResult = DaqChannelOnOffSetting(btnDAM10_Ch21, strDAQPinMap, strDAQDIOPinMap, strDAQAIPinMap);
                        Thread.Sleep(intSleep);
                    }
                }

                #endregion

                #region 일반 모드 전압값 취득

                if (chkNormalResistance.Checked)
                {
                    strDAQPinMap = string.Format(Gini.GetValue("SelfTest", "SelfTestNormalMode_DAQPinName").Trim(), strDAQDeviceName);
                    boolResult = DaqChannelOnOffLCRMeterValue(teNormalResistance, strDAQPinMap, false);
                }
                #endregion

                #region 휘스톤 모드 전압값 취득
                if (chkWheatstoneResistance.Checked)
                {
                    strDAQPinMap = string.Format(Gini.GetValue("SelfTest", "SelfTestWheatstoneMode_DAQPinName").Trim(), strDAQDeviceName);
                    boolResult = DaqChannelOnOffLCRMeterValue(teWheatstoneResistance, strDAQPinMap, true);
                }
                #endregion
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
                boolMeasurement = false;
            }
            finally
            {
                frmMain.m_MeasureProcess.DigitalDAQ_CloseChannel();

                if (boolResult && boolMeasurement)
                {
                    frmMB.lblMessage.Text = "Self Test를 완료합니다.";
                    frmMB.TopMost = true;
                    frmMB.ShowDialog();
                }
                else if (boolResult && !boolMeasurement)
                {
                    frmMB.lblMessage.Text = "Self Test을 취소하였습니다.";
                    frmMB.TopMost = true;
                    frmMB.ShowDialog();
                }
                else
                {
                    frmMB.lblMessage.Text = "Self Test를 실패하였습니다.";
                    frmMB.TopMost = true;
                    frmMB.ShowDialog();
                }

                boolMeasurement = false;

                // Button Enabled 설정
                SetButtonEnabled(true);

                if (_thread != null)
                    _thread.Abort();
            }
        }

        /// <summary>
        /// DAQ Pin On/Off와 DIO Channel 읽기어 상태 체크
        /// </summary>
        /// <param name="_btn"></param>
        /// <param name="strDAQPinMap"></param>
        /// <param name="_strDAQDIOPinMap"></param>
        /// <returns></returns>
        private bool DaqChannelOnOffSetting(Button _btn, string strDAQPinMap, string _strDAQDIOPinMap, string _strDAQAIPinMap)
        {
            bool boolResult = false;
            bool boolDIOChannelState = false;
            double dMeasurementValue = 0d;
                        
            boolResult = frmMain.m_MeasureProcess.m_DigitalDAQ.DigitalDAQ_WriteChannel(strDAQPinMap, true);

            if (boolResult)
            {
                //boolDIOChannelState = frmMain.m_MeasureProcess.m_DigitalDAQ.DigitalDAQ_ReaderChannel(_strDAQDIOPinMap);                
                for (int i = 0; i < 5; i++)
                {
                    boolDIOChannelState = frmMain.m_MeasureProcess.m_AIReader.startAITaskSelfTestReader(_strDAQAIPinMap, -10, 10, ref dMeasurementValue, 1000, 300);

                    if (boolDIOChannelState) break;
                }

                if (boolDIOChannelState)
                    _btn.BackColor = System.Drawing.Color.Lime;
                else
                    _btn.BackColor = System.Drawing.Color.Red;
            }

            boolResult = frmMain.m_MeasureProcess.m_DigitalDAQ.DigitalDAQ_WriteChannel(strDAQPinMap, false);

            return boolResult;
        }

        /// <summary>
        /// DAQ Pin On/Off와 LCR-Meter 값 읽어오기 (일반/휘스톤 모드)
        /// </summary>
        /// <param name="_btn"></param>
        /// <param name="strDAQPinMap"></param>
        /// <param name="_strDAQDIOPinMap"></param>
        /// <returns></returns>
        private bool DaqChannelOnOffLCRMeterValue(TextBox _te, string strDAQPinMap, bool boolWheatstoneMode)
        {
            bool boolResult = false;

            boolResult = frmMain.m_MeasureProcess.m_DigitalDAQ.DigitalDAQ_WriteChannel(strDAQPinMap, true);

            if (boolResult)
            {
                // 주파수 전송
                frmMain.m_MeasureProcess.m_LCRMeter.OptFreqSelect(100);

                // 전압레벨 전송
                frmMain.m_MeasureProcess.m_LCRMeter.OptLevelSelect(1);

                // DC_R 전환 전송
                frmMain.m_MeasureProcess.m_LCRMeter.OptSelect(Function.FunctionLCRMeterinfo.strCmd_LCRMeter_RDCLsSelect.Trim());

                // LCR Meter DC_R 값 읽어 오기
                decimal dcmValue = 0M;
                decimal dcmWheatstoneValue = 0M;

                // 휘스톤 모드일 경우
                if (boolWheatstoneMode)
                {
                    double dMeasurementValue = 0d, dRate = 1000d;
                    int intSamplesPerChannel = 1000;

                    Device.ClassAnalogDAQProcess m_AIReader = new Device.ClassAnalogDAQProcess();
                    Function.FunctionMeasureProcess m_Measure = new Function.FunctionMeasureProcess();

                    m_AIReader.startAITaskMultiReader(string.Format(Gini.GetValue("Wheatstone", "WheatstoneAIChannel_DAQPinName").Trim(), Gini.GetValue("Device", "DAQDeviceName").Trim())
                        , -10.00d, 10.00d, ref dMeasurementValue, dRate, intSamplesPerChannel);

                    dcmValue = Convert.ToDecimal(dMeasurementValue);
                    dcmWheatstoneValue = m_Measure.SetWheatstoneModeMeasurementValueCalculate(dcmValue);
                    _te.Text = dcmWheatstoneValue.ToString("F3").Trim();
                }
                else
                {
                    // 일반 모드 LCR Meter DC_R 값 읽어 오기
                    dcmValue = frmMain.m_MeasureProcess.m_LCRMeter.LCRReadFunRepeat(3, 0, false);
                    _te.Text = dcmValue.ToString("F3").Trim();
                }
            }

            boolResult = frmMain.m_MeasureProcess.m_DigitalDAQ.DigitalDAQ_WriteChannel(strDAQPinMap, false);

            return boolResult;
        }

        /// <summary>
        /// 초기화 Button Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnInitialize_Click(object sender, EventArgs e)
        {
            SetControlInitialize();
        }

        /// <summary>
        /// DAM ALL Checked Chaned Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void chkDAMAll_CheckedChanged(object sender, EventArgs e)
        {
            chkDAM1.Checked = chkDAMAll.Checked;
            chkDAM2.Checked = chkDAMAll.Checked;
            chkDAM3.Checked = chkDAMAll.Checked;
            chkDAM4.Checked = chkDAMAll.Checked;
            chkDAM5.Checked = chkDAMAll.Checked;
            chkDAM6.Checked = chkDAMAll.Checked;
            chkDAM7.Checked = chkDAMAll.Checked;
            chkDAM8.Checked = chkDAMAll.Checked;
            chkDAM9.Checked = chkDAMAll.Checked;
            chkDAM10.Checked = chkDAMAll.Checked;
        }
    }
}
