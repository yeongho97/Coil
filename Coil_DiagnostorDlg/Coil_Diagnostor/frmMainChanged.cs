using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using System.Diagnostics;   // Process 사용을 위해 선언
using System.IO;
using Coil_Diagnostor.Function;

namespace Coil_Diagnostor
{
	public partial class frmMainChanged : Form
    {
        public frmMainChanged()
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
                case Keys.F1: // 조회 버튼
                    btnFilePathSetting.PerformClick();
                    break;
                case Keys.F2: // 기준값 표시/숨김 버튼
                    btnMain.PerformClick();
                    break;
                case Keys.F3: // 평균값 표시/숨김 버튼
                    btnCoilDiagnostor.PerformClick();
                    break;
                case Keys.Control & Keys.F12: // 닫기 버튼
                    btnExit.PerformClick();
                    break;
            }

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
        }

        /// <summary>
        /// Form Closing Evnet
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmMain_FormClosing(object sender, FormClosingEventArgs e)
        {
        }

        #endregion


        #region ▣ 버튼 이벤트

		private void buttonForeColorSetting(System.Drawing.Color _color, Button _button, bool _boolButtonOnOff)
		{
            btnFilePathSetting.Enabled = _boolButtonOnOff;
			btnCoilDiagnostor.Enabled = _boolButtonOnOff;
			btnMain.Enabled = _boolButtonOnOff;
			
			if (!_boolButtonOnOff)
            {
                btnFilePathSetting.ForeColor = System.Drawing.Color.Silver;
				btnCoilDiagnostor.ForeColor = System.Drawing.Color.Silver;
				btnMain.ForeColor = System.Drawing.Color.Silver;

				switch (_button.Name.Trim())
				{
                    case "btnRCSDiagnosis":
						btnCoilDiagnostor.ForeColor = _color;
						btnCoilDiagnostor.Enabled = true;
						break;
                    case "btnRCSHistory":
						btnMain.ForeColor = _color;
						btnMain.Enabled = true;
                        break;
                    case "btnFilePathSetting":
                        btnFilePathSetting.ForeColor = _color;
                        btnFilePathSetting.Enabled = true;
                        break;
				}
			}
			else
            {
                btnFilePathSetting.ForeColor = System.Drawing.Color.White;
				btnCoilDiagnostor.ForeColor = System.Drawing.Color.White;
				btnMain.ForeColor = System.Drawing.Color.White;
			}
		}

        /// <summary>
        /// 프로그램 종료
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// 기준 코일 성능시험기 실행 파일 경로 및 파일명 선택
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnFilePathSetting_Click(object sender, EventArgs e)
        {
            buttonForeColorSetting(System.Drawing.Color.Yellow, ((Button)sender), false);

            frmFilePathSetting frm = new frmFilePathSetting();
            frm.ShowDialog();

            buttonForeColorSetting(System.Drawing.Color.White, ((Button)sender), true);
        }

        /// <summary>
        /// 기존 코일 성능시험기 프로그램 호출
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCoilDiagnostor_Click(object sender, EventArgs e)
		{
			buttonForeColorSetting(System.Drawing.Color.Yellow, ((Button)sender), false);

            Process UserProcess = new Process();
            ProcessStartInfo startIndo = new ProcessStartInfo(Gini.GetValue("Device", "FilePathSetting").Trim());
            startIndo.UseShellExecute = false;
            startIndo.RedirectStandardOutput = true;
            UserProcess.StartInfo = startIndo;
            UserProcess.Start();
            StreamReader sReader = UserProcess.StandardOutput;
            string strMsg = sReader.ReadLine();
            sReader.Close();
            UserProcess.Close();

			buttonForeColorSetting(System.Drawing.Color.White, ((Button)sender), true);
        }

        /// <summary>
        /// 신규 코일 성능시험기 프로그램 호출
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnMain_Click(object sender, EventArgs e)
        {
            buttonForeColorSetting(System.Drawing.Color.Yellow, ((Button)sender), false);

            frmMain frm = new frmMain();
            frm.ShowDialog();

            buttonForeColorSetting(System.Drawing.Color.White, ((Button)sender), true);
        }

        #endregion
    }
}
