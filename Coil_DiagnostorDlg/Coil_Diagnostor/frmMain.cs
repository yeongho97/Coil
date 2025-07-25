using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using Coil_Diagnostor.Function;

namespace Coil_Diagnostor
{
	public partial class frmMain : Form
    {
        public Function.FunctionMeasureProcess m_MeasureProcess = new Function.FunctionMeasureProcess();
        public frmMessageBox frmMB = new frmMessageBox();

        public MultimediaTimer mmTimer = new MultimediaTimer();
        public bool boolPinOnOff = false;
        public string strDAQPinMap = "";

		public frmMain()
        {
            CheckForIllegalCrossThreadCalls = false;
			InitializeComponent();

            mmTimer.Tick += new System.EventHandler(mmTimer_Tick);
		}

        /// <summary>
        /// 폼 단축키 지정
        /// </summary>
        protected override bool ProcessCmdKey(ref System.Windows.Forms.Message msg, Keys keyData)
        {
            Keys key = keyData & ~(Keys.Shift | Keys.Control);

            //switch (key)
            //{
                //case Keys.F1: // 조회 버튼
                //    btnSearch.PerformClick();
                //    break;
                //case Keys.F2: // 기준값 표시/숨김 버튼
                //    btnReferenceValue.PerformClick();
                //    break;
                //case Keys.F3: // 평균값 표시/숨김 버튼
                //    btnAverageValue.PerformClick();
                //    break;
                //case Keys.F4: // Scaling 설정 버튼
                //    btnSearch.PerformClick();
                //    break;
                //case Keys.Control & Keys.F12: // 닫기 버튼
                //    btnExit.PerformClick();
                //    break;
            //}

            return base.ProcessCmdKey(ref msg, keyData);
        }

        #region ▣ Form 이벤트

        /// <summary>
        /// Form Load 될 때마다
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmMain_Load(object sender, EventArgs e)
        {
            // Daq 설정
            if (!m_MeasureProcess.AccessDigitalDAQChecking(Gini.GetValue("Device", "DAQDeviceName").Trim()))
            {
                frmMB.lblMessage.Text = "DAQ 연결을 확인하십시오.";
                frmMB.TopMost = true;
                frmMB.ShowDialog();
            }
            else
            {
                strDAQPinMap = string.Format(Gini.GetValue("Device", "OPROnOff_DAQPinMap").Trim(), Gini.GetValue("Device", "DAQDeviceName").Trim());
                mmTimer.Period = 1000;
                mmTimer.Resolution = 1;
                mmTimer.Start();
            }

            bool bUseTDRMeter = Convert.ToBoolean(Gini.GetValue("Device", "USE_TDRMeter"));
            if (!bUseTDRMeter)
            {
                tblpnlMenu.ColumnStyles[3].Width = 0;
                tblpnlMenu.ColumnStyles[4].Width = 0;
            }
        }

        /// <summary>
        /// Form Closing Evnet
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            m_MeasureProcess.DigitalDAQ_ChannelOnOff(strDAQPinMap, false);
        }

        #endregion


        #region ▣ 버튼 이벤트

		private void buttonForeColorSetting(System.Drawing.Color _color, Button _button, bool _boolButtonOnOff)
		{
			btnRCSDiagnosis.Enabled = _boolButtonOnOff;
			btnDRPIDiagnosis.Enabled = _boolButtonOnOff;
			btnRCSHistory.Enabled = _boolButtonOnOff;
			btnDRPIHistory.Enabled = _boolButtonOnOff;
			btnRCSReport.Enabled = _boolButtonOnOff;
			btnDRPIReport.Enabled = _boolButtonOnOff;
			btnSelfTest.Enabled = _boolButtonOnOff;
			btnTDR_RCSDiagnosis.Enabled = _boolButtonOnOff;
            btnTDR_RCSHistory.Enabled = _boolButtonOnOff;
            btnTDR_DRPIDiagnosis.Enabled = _boolButtonOnOff;
            btnTDR_DRPIHistory.Enabled = _boolButtonOnOff;
			btnSetOffset.Enabled = _boolButtonOnOff;
			btnConfiguration.Enabled = _boolButtonOnOff;
            btnAccessDB.Enabled = _boolButtonOnOff;
			
			if (!_boolButtonOnOff)
			{
				btnRCSDiagnosis.ForeColor = System.Drawing.Color.Silver;
				btnDRPIDiagnosis.ForeColor = System.Drawing.Color.Silver;
				btnRCSHistory.ForeColor = System.Drawing.Color.Silver;
				btnDRPIHistory.ForeColor = System.Drawing.Color.Silver;
				btnRCSReport.ForeColor = System.Drawing.Color.Silver;
				btnDRPIReport.ForeColor = System.Drawing.Color.Silver;
				btnSelfTest.ForeColor = System.Drawing.Color.Silver;
				btnTDR_RCSDiagnosis.ForeColor = System.Drawing.Color.Silver;
                btnTDR_RCSHistory.ForeColor = System.Drawing.Color.Silver;
                btnTDR_DRPIDiagnosis.ForeColor = System.Drawing.Color.Silver;
                btnTDR_DRPIHistory.ForeColor = System.Drawing.Color.Silver;
				btnSetOffset.ForeColor = System.Drawing.Color.Silver;
                btnConfiguration.ForeColor = System.Drawing.Color.Silver;
                btnAccessDB.ForeColor = System.Drawing.Color.Silver;

				switch (_button.Name.Trim())
				{
                    case "btnRCSDiagnosis":
						btnRCSDiagnosis.ForeColor = _color;
						btnRCSDiagnosis.Enabled = true;
						break;
                    case "btnDRPIDiagnosis":
						btnDRPIDiagnosis.ForeColor = _color;
						btnDRPIDiagnosis.Enabled = true;
						break;
                    case "btnRCSHistory":
						btnRCSHistory.ForeColor = _color;
						btnRCSHistory.Enabled = true;
						break;
                    case "btnDRPIHistory":
						btnDRPIHistory.ForeColor = _color;
						btnDRPIHistory.Enabled = true;
						break;
                    case "btnRCSReport":
						btnRCSReport.ForeColor = _color;
						btnRCSReport.Enabled = true;
						break;
                    case "btnDRPIReport":
						btnDRPIReport.ForeColor = _color;
						btnDRPIReport.Enabled = true;
						break;
                    case "btnSelfTest":
						btnSelfTest.ForeColor = _color;
						btnSelfTest.Enabled = true;
						break;
                    case "btnTDR_RCSDiagnosis":
						btnTDR_RCSDiagnosis.ForeColor = _color;
						btnTDR_RCSDiagnosis.Enabled = true;
						break;
                    case "btnTDR_RCSHistory":
						btnTDR_RCSHistory.ForeColor = _color;
						btnTDR_RCSHistory.Enabled = true;
                        break;
                    case "btnTDR_DRPIDiagnosis":
                        btnTDR_DRPIDiagnosis.ForeColor = _color;
                        btnTDR_DRPIDiagnosis.Enabled = true;
                        break;
                    case "btnTDR_DRPIHistory":
                        btnTDR_DRPIHistory.ForeColor = _color;
                        btnTDR_DRPIHistory.Enabled = true;
                        break;
                    case "btnSetOffset":
						btnSetOffset.ForeColor = _color;
						btnSetOffset.Enabled = true;
						break;
                    case "btnConfiguration":
						btnConfiguration.ForeColor = _color;
						btnConfiguration.Enabled = true;
                        break;
                    case "btnAccessDB":
                        btnAccessDB.ForeColor = _color;
                        btnAccessDB.Enabled = true;
                        break;
				}
			}
			else
			{
				btnRCSDiagnosis.ForeColor = System.Drawing.Color.White;
				btnDRPIDiagnosis.ForeColor = System.Drawing.Color.White;
				btnRCSHistory.ForeColor = System.Drawing.Color.White;
				btnDRPIHistory.ForeColor = System.Drawing.Color.White;
				btnRCSReport.ForeColor = System.Drawing.Color.White;
				btnDRPIReport.ForeColor = System.Drawing.Color.White;
				btnSelfTest.ForeColor = System.Drawing.Color.White;
				btnTDR_RCSDiagnosis.ForeColor = System.Drawing.Color.White;
                btnTDR_RCSHistory.ForeColor = System.Drawing.Color.White;
                btnTDR_DRPIDiagnosis.ForeColor = System.Drawing.Color.White;
                btnTDR_DRPIHistory.ForeColor = System.Drawing.Color.White;
				btnSetOffset.ForeColor = System.Drawing.Color.White;
                btnConfiguration.ForeColor = System.Drawing.Color.White;
                btnAccessDB.ForeColor = System.Drawing.Color.White;
			}
		}

        /// <summary>
        /// 프로그램 종료
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnExit_Click(object sender, EventArgs e)
        {
            mmTimer.Stop();
            
            this.Close();
        }

        /// <summary>
        /// RCS 진단 화면 호출
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnRCSDiagnosis_Click(object sender, EventArgs e)
		{
			buttonForeColorSetting(System.Drawing.Color.Yellow, ((Button)sender), false);

            frmRCSDiagnosis frm = new frmRCSDiagnosis();
            frm.ShowDialog();

			buttonForeColorSetting(System.Drawing.Color.White, ((Button)sender), true);
        }

        /// <summary>
        /// RCS 이력 화면 호출
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnRCSHistory_Click(object sender, EventArgs e)
        {
            buttonForeColorSetting(System.Drawing.Color.Yellow, ((Button)sender), false);

            frmRCSHistory frm = new frmRCSHistory();
            frm.ShowDialog();

            buttonForeColorSetting(System.Drawing.Color.White, ((Button)sender), true);
        }

        /// <summary>
        /// RCS 보고서 화면 호출
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnRCSReport_Click(object sender, EventArgs e)
        {
            buttonForeColorSetting(System.Drawing.Color.Yellow, ((Button)sender), false);

            frmRCSReport frm = new frmRCSReport();
            frm.ShowDialog();

            buttonForeColorSetting(System.Drawing.Color.White, ((Button)sender), true);
        }

        /// <summary>
        /// DRPI 진단 화면 호출
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnDRPIDiagnosis_Click(object sender, EventArgs e)
        {
            buttonForeColorSetting(System.Drawing.Color.Yellow, ((Button)sender), false);

            frmDRPIDiagnosis frm = new frmDRPIDiagnosis();
            frm.ShowDialog();

            buttonForeColorSetting(System.Drawing.Color.White, ((Button)sender), true);
        }

        /// <summary>
        /// DRPI 이력 화면 호출
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnDRPIHistory_Click(object sender, EventArgs e)
        {
            buttonForeColorSetting(System.Drawing.Color.Yellow, ((Button)sender), false);

            frmDRPIHistory frm = new frmDRPIHistory();
            frm.ShowDialog();

            buttonForeColorSetting(System.Drawing.Color.White, ((Button)sender), true);
        }

        /// <summary>
        /// DRPI 보고서 화면 호출
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnDRPIReport_Click(object sender, EventArgs e)
        {
            buttonForeColorSetting(System.Drawing.Color.Yellow, ((Button)sender), false);

            frmDRPIReport frm = new frmDRPIReport();
            frm.ShowDialog();

            buttonForeColorSetting(System.Drawing.Color.White, ((Button)sender), true);
        }

        /// <summary>
        /// TDR-RCS 진단 화면 호출
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnTDR_RCSDiagnosis_Click(object sender, EventArgs e)
        {
            buttonForeColorSetting(System.Drawing.Color.Yellow, ((Button)sender), false);

            //frmTDR_RCSDiagnosis frm = new frmTDR_RCSDiagnosis();
            //frm.ShowDialog();

            buttonForeColorSetting(System.Drawing.Color.White, ((Button)sender), true);
        }

        /// <summary>
        /// TDR-RCS 이력 화면 호출
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnTDR_RCSHistory_Click(object sender, EventArgs e)
        {
            buttonForeColorSetting(System.Drawing.Color.Yellow, ((Button)sender), false);

            //frmTDR_RCSHistory frm = new frmTDR_RCSHistory();
            //frm.ShowDialog();

            buttonForeColorSetting(System.Drawing.Color.White, ((Button)sender), true);
        }

        /// <summary>
        /// TDR-DRPI 진단 화면 호출
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnTDR_DRPIDiagnosis_Click(object sender, EventArgs e)
        {
            buttonForeColorSetting(System.Drawing.Color.Yellow, ((Button)sender), false);

            //frmTDR_DRPIDiagnosis frm = new frmTDR_DRPIDiagnosis();
            //frm.ShowDialog();

            buttonForeColorSetting(System.Drawing.Color.White, ((Button)sender), true);
        }

        /// <summary>
        /// TDR-DRPI 이력 화면 호출
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnTDR_DRPIHistory_Click(object sender, EventArgs e)
        {
            buttonForeColorSetting(System.Drawing.Color.Yellow, ((Button)sender), false);

            //frmTDR_DRPIHistory frm = new frmTDR_DRPIHistory();
            //frm.ShowDialog();

            buttonForeColorSetting(System.Drawing.Color.White, ((Button)sender), true);
        }

        /// <summary>
        /// 자체 진단 화면 호출
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSelfTest_Click(object sender, EventArgs e)
        {
            buttonForeColorSetting(System.Drawing.Color.Yellow, ((Button)sender), false);

            frmSelfTest frm = new frmSelfTest(this);
            frm.ShowDialog();

            buttonForeColorSetting(System.Drawing.Color.White, ((Button)sender), true);
        }

        /// <summary>
        /// 영점 보정 화면 호출
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSetOffset_Click(object sender, EventArgs e)
        {
            buttonForeColorSetting(System.Drawing.Color.Yellow, ((Button)sender), false);

            frmSetOffset frm = new frmSetOffset();
            frm.ShowDialog();

            buttonForeColorSetting(System.Drawing.Color.White, ((Button)sender), true);
        }

        /// <summary>
        /// 환경설정 화면 호출
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnConfiguration_Click(object sender, EventArgs e)
        {
            buttonForeColorSetting(System.Drawing.Color.Yellow, ((Button)sender), false);

            frmConfiguration frm = new frmConfiguration();
            frm.ShowDialog();

            buttonForeColorSetting(System.Drawing.Color.White, ((Button)sender), true);
        }

        /// <summary>
        /// 데이터 병합 화면 호출
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnAccessDB_Click(object sender, EventArgs e)
        {
            buttonForeColorSetting(System.Drawing.Color.Yellow, ((Button)sender), false);

            frmAccessDBReader frm = new frmAccessDBReader();
            frm.ShowDialog();

            buttonForeColorSetting(System.Drawing.Color.White, ((Button)sender), true);
        }

        #endregion
        
        /// <summary>
        /// OPR DAQ PinMap 자동 On/Off 처리
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void mmTimer_Tick(object sender, EventArgs e)
        {
            if (!boolPinOnOff)
            {
                boolPinOnOff = true;
                m_MeasureProcess.DigitalDAQ_ChannelOnOff(strDAQPinMap, boolPinOnOff);
            }
            else
            {
                boolPinOnOff = false;
                m_MeasureProcess.DigitalDAQ_ChannelOnOff(strDAQPinMap, boolPinOnOff);
            }
        }
    }
}
