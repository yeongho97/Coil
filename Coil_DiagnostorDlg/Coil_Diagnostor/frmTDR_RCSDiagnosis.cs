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
using System.IO;

// TCP/IP (Lan) 통신용
using System.Net;
using System.Net.Sockets;
using System.Threading;
using Coil_Diagnostor.Function;
using Coil_Diagnostor.UserControl;

namespace Coil_Diagnostor
{
    public partial class frmTDR_RCSDiagnosis : System.Windows.Forms.Form
    {
        protected Function.FunctionDataControl m_db = new Function.FunctionDataControl();
        protected frmMessageBox frmMB = new frmMessageBox();
        protected string strPlantName = "";
        protected bool boolMessageOK = false;
        protected bool boolFormLoad = false;
        protected bool boollInitialize = false;
        public bool boolMeasurementStart = false;
        public bool boolMeasurementStop = false;
        public bool boolDataSaveStart = false;
        public bool boolNormalMode = false;

        private RCSRodCabinetPanel rcsPanel;

        public string strSelectHogi = "";
        public string strSelectOhDegree = "";
        public string strSelectPowerCabinet = "";
        public string strSelectControlRodName = "";
        public string strSelectCoilName = "";
        public string strMeasurement_StartDate = "";
        public string strMeasurement_EndDate = "";
        public string strDataSavePath = "";
        public string strImageFileName = "";
        public string strTextFileName = "";
        public string strImageDataPathFileName = "";
        public string strTextDataPathFileName = "";

        public string strURLImage = "";
        public string strURLText = "";

        public int intPowerCabinetButtonSelectIndex = 1;

        public IPAddress ipadress;

        public frmTDR_RCSDiagnosis()
        {
            CheckForIllegalCrossThreadCalls = false;
            InitializeComponent();

            // 저장 폴더 확인 및 생성
            string strDataSavePath = Application.StartupPath + @"\TDR_Data";
            System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(strDataSavePath);

            if (!di.Exists)
            {
                di.Create();
            }
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
        private void frmTDR_RCSDiagnosis_Load(object sender, EventArgs e)
        {
            boolFormLoad = true;

            strPlantName = Gini.GetValue("Device", "PlantName").Trim();
            rcsPanel = new RCSRodCabinetPanel(panel3, "33", 2);
            rcsPanel.Click += rcsPanel_Click;

            // ComboBox 데이터 설정
            SetComboBoxDataBinding();

            // 초기화
            SetControlInitialize();

            rcsPanel.TDRInitializePanel();

            // 포인트
            cboHogi.Focus();


            ipadress = IPAddress.Parse(Gini.GetValue("Device", "TDRMeter_IPAddress").Trim());

            strURLImage = string.Format("http://{0}/www/screenshot.bmp", Gini.GetValue("Device", "TDRMeter_IPAddress").Trim());
            strURLText = string.Format("http://{0}/www/trace.csv", Gini.GetValue("Device", "TDRMeter_IPAddress").Trim());

            if (MessageBox.Show("TDR-Meter 장비 Lan 선연결 및 Web Export를 On 하십시오.", "TDR-Meter 통신 확인", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == System.Windows.Forms.DialogResult.Yes)
            {
                // TDR-Meter (이더넷 연결 체크)
                System.Net.NetworkInformation.Ping ping = new System.Net.NetworkInformation.Ping();
                System.Net.NetworkInformation.PingOptions options = new System.Net.NetworkInformation.PingOptions();
                options.DontFragment = true;
                string data = "aaaaaaaaaaaaaaaaa";
                byte[] buffer = ASCIIEncoding.ASCII.GetBytes(data);
                int timeout = 120;
                System.Net.NetworkInformation.PingReply reply = ping.Send(IPAddress.Parse(Gini.GetValue("Device", "TDRMeter_IPAddress").Trim()), timeout, buffer, options);

                if (reply.Status == System.Net.NetworkInformation.IPStatus.Success)
                {
                    if (DownloadRemoteImageFile(strURLImage, "", true))
                    {
                        //ledTDRMeter.Value = true;
                        ledTDRMeter.On = true;
                        btnMeasurementStart.Enabled = true;
                    }
                    else
                    {
                        //ledTDRMeter.Value = false;
                        ledTDRMeter.On = false;
                        btnMeasurementStart.Enabled = false;
                    }
                }
                else
                {
                    //ledTDRMeter.Value = false;
                    ledTDRMeter.On = false;
                    btnMeasurementStart.Enabled = false;
                }

                reply = null;
            }
            else
            {
                string strMessage = "TDR-Meter Web Export On 설정 방법은 매뉴얼를 참조하십시오.";

                frmMB.lblMessage.Text = strMessage;
                frmMB.TopMost = true;
                frmMB.ShowDialog();
                this.Close();
            }

            boolFormLoad = false;
        }

        private void rcsPanel_Click(object sender, EventArgs e)
        {
            //
            RCSRodCabinetPanel p = sender as RCSRodCabinetPanel;
            pbTDRImage.Image = null;
            strSelectPowerCabinet = p.GetCabinetName();
            strSelectControlRodName = p.GetRodName();
        }

        /// <summary>
        /// ComboBox 데이터 설정
        /// </summary>
        private void SetComboBoxDataBinding()
        {
            // 발전소 호기
            string[] strHogi = Gini.GetValue("Combo", "HogiData").Split(',');

            cboHogi.Items.Clear();

            for (int i = 0; i < strHogi.Length; i++)
            {
                cboHogi.Items.Add(strHogi[i].Trim());
            }

            // 호기 선택
            if (Gini.GetValue("RCS", "SelectRCS_TDRHogi").Trim() == "")
            {
                // 최초 실행 시 기본 호기 선택
                cboHogi.SelectedIndex = 0;
            }
            else
            {
                // 직전 실행 시 기본 호기 선택
                cboHogi.SelectedItem = Gini.GetValue("RCS", "SelectRCS_TDRHogi").Trim();
            }
        }

        /// <summary>
        /// 초기화
        /// </summary>
        private void SetControlInitialize()
        {
            // 측정 항목 설정
            rboMeasurementItem_Stop.Checked = true;
            strSelectCoilName = "정지";

            if (pbTDRImage.Image != null) pbTDRImage.Image.Clone();
        }
        /// <summary>
        /// Form Closing Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmTDR_RCSDiagnosis_FormClosing(object sender, FormClosingEventArgs e)
        {

            Gini.SetValue("RCS", "SelectRCS_TDRHogi", cboHogi.SelectedItem.ToString().Trim());
            Gini.SetValue("RCS", "SelectRCS_TDROHDegree", teOhDegree2.Text.Trim());
        }


        /// <summary>
        /// 호기 Selected Index Changed Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cboHogi_SelectedIndexChanged(object sender, EventArgs e)
        {
            teOhDegree2.Text = Gini.GetValue("RCS", "SelectRCS_TDROHDegree").Trim();

            pbTDRImage.Image = null;
        }

        /// <summary>
        /// 정지/올림/이동 Radio Button Checked Changed Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void rboMeasurementItem_CheckedChanged(object sender, EventArgs e)
        {
            if (rboMeasurementItem_Stop.Checked)
                strSelectCoilName = "정지";
            else if (rboMeasurementItem_Up.Checked)
                strSelectCoilName = "올림";
            else if (rboMeasurementItem_Move.Checked)
                strSelectCoilName = "이동";
            
            pbTDRImage.Image = null;
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
        /// 정지 Button Evnet
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnMeasurementStop_Click(object sender, EventArgs e)
        {
            try
            {
                //bool boolResult = m_Client.CloseSocket();

                if (boolDataSaveStart)
                {
                    // 기준 데이터의 종료 일시 업데이터
                    boolDataSaveStart = m_db.UpDateTDRRCSDiagnosisHeaderDataGridViewDataInfo(strPlantName.Trim(), strSelectHogi.Trim(), strSelectOhDegree.Trim()
                    , strSelectPowerCabinet.Trim(), strSelectControlRodName.Trim(), strSelectCoilName.Trim(), strMeasurement_EndDate.Trim());
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
            finally
            {
                // 측정 시작 일시
                strMeasurement_StartDate = "";
                strMeasurement_EndDate = "";
                boolDataSaveStart = false;

                //ledTDRMeter.Value = false;
                ledTDRMeter.On = false;
            }
        }

        /// <summary>
        /// 측정 Button Evnet
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnMeasurementStart_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;

            try
            {
                pbTDRImage.Image = null;

                // 대상 호기 체크
                if (cboHogi == null || cboHogi.SelectedItem == null || cboHogi.SelectedItem.ToString().Trim() == "")
                {
                    frmMB.lblMessage.Text = "호기를 선택하십시오.";
                    frmMB.TopMost = true;
                    frmMB.ShowDialog();
                    return;
                }

                // 대상 차수 체크
                if (teOhDegree2.Text.Trim() == "")
                {
                    frmMB.lblMessage.Text = "차수를 입력하십시오.";
                    frmMB.TopMost = true;
                    frmMB.ShowDialog();
                    return;
                }

                // 전력합 선택 체크
                if (strSelectPowerCabinet.Trim() == "")
                {
                    frmMB.lblMessage.Text = "전력합을 선택하십시오.";
                    frmMB.TopMost = true;
                    frmMB.ShowDialog();
                    return;
                }

                // 제어봉 선택 체크
                if (strSelectControlRodName.Trim() == "")
                {
                    frmMB.lblMessage.Text = "제어봉을 선택하십시오.";
                    frmMB.TopMost = true;
                    frmMB.ShowDialog();
                    return;
                }

                // 코일명 선택 체크
                if (!rboMeasurementItem_Stop.Checked && !rboMeasurementItem_Up.Checked && !rboMeasurementItem_Move.Checked)
                {
                    frmMB.lblMessage.Text = "코일명을 선택하십시오.";
                    frmMB.TopMost = true;
                    frmMB.ShowDialog();
                    return;
                }

                // IP Address 설정 체크
                if (Gini.GetValue("Device", "TDRMeter_IPAddress").Trim() == "")
                {
                    frmMB.lblMessage.Text = "환경설정 화면에서 TDR IP Address를 설정하십시오.";
                    frmMB.TopMost = true;
                    frmMB.ShowDialog();
                    return;
                }

                // IP Port 설정 체크
                if (Gini.GetValue("Device", "TDRMeter_IPPort").Trim() == "")
                {
                    frmMB.lblMessage.Text = "환경설정 화면에서 TDR IP Port를 설정하십시오.";
                    frmMB.TopMost = true;
                    frmMB.ShowDialog();
                    return;
                }

                // Control 활성화/비활성화
                SetControlEnabled(false);

                //if (ledTDRMeter.Value)
                if (ledTDRMeter.On)
                {
                    strSelectHogi = cboHogi.SelectedItem == null ? "" : cboHogi.SelectedItem.ToString().Trim();
                    strSelectOhDegree = lblOhDegree1.Text.Trim() + " " + teOhDegree2.Text.Trim() + " " + lblOhDegree3.Text.Trim();

                    // 측정 시작 일시
                    DateTime dataTime = System.DateTime.Now;
                    strMeasurement_StartDate = dataTime.ToString("yyyy-MM-dd HH:mm:ss");
                    strMeasurement_EndDate = strMeasurement_StartDate;

                    // 경로 생성
                    strDataSavePath = string.Format(@"\TDR_Data\{0}\", dataTime.ToString("yyyy"));
                    string strDataSavePath1 = string.Format(@"{0}\TDR_Data\{1}\", Application.StartupPath.Trim(), dataTime.ToString("yyyy"));
                    DirectoryInfo di = new DirectoryInfo(strDataSavePath1.Trim());

                    if (!di.Exists)
                        di.Create();

                    // 받은 Image 파일 생성
                    strImageFileName = string.Format("TDR_RCD {0}_{1}_{2}_{3}_{4}_{5}.bmp", strSelectHogi, strSelectOhDegree
                        , strSelectPowerCabinet, strSelectControlRodName, strSelectCoilName, dataTime.ToString("yyyyMMdd_HHmmss"));

                    strTextFileName = string.Format("TDR_RCD {0}_{1}_{2}_{3}_{4}_{5}.csv", strSelectHogi, strSelectOhDegree
                        , strSelectPowerCabinet, strSelectControlRodName, strSelectCoilName, dataTime.ToString("yyyyMMdd_HHmmss"));

                    strImageDataPathFileName = strDataSavePath.Trim() + strImageFileName.Trim();
                    strTextDataPathFileName = strDataSavePath.Trim() + strTextFileName.Trim();

                    // TDR-Meter 화면 캡쳐 가져오기
                    if (DownloadRemoteImageFile(strURLImage, Application.StartupPath.Trim() + strImageDataPathFileName, false))
                    {
                        openFileDialog1.FileName = Application.StartupPath.Trim() + strImageDataPathFileName;
                        pbTDRImage.Image = new Bitmap(openFileDialog1.FileName);
                        pbTDRImage.SizeMode = PictureBoxSizeMode.Zoom;

                        // TDR-Meter 데이터 가져오기
                        if (fileDownload(strURLText, Application.StartupPath.Trim() + strTextDataPathFileName))
                        {
                            // 데이터 중복 체크
                            if ((m_db.GetTDRRCSDiagnosisHeaderDataGridViewDataCount(strPlantName.Trim(), strSelectHogi.Trim(), strSelectOhDegree.Trim()
                                , strSelectPowerCabinet.Trim(), strSelectControlRodName.Trim(), strSelectCoilName.Trim())) > 0)
                            {
                                // 기존 데이터 삭제
                                boolDataSaveStart = m_db.DeleteTDRRCSDiagnosisHeaderDataGridViewDataInfo(strPlantName.Trim(), strSelectHogi.Trim(), strSelectOhDegree.Trim()
                                , strSelectPowerCabinet.Trim(), strSelectControlRodName.Trim(), strSelectCoilName.Trim());
                            }

                            // 데이터 신규 저장
                            int intSeqNumber = Regex.Replace(strSelectCoilName.Trim(), @"\D", "") == "" ? 1 : Convert.ToInt32(Regex.Replace(strSelectCoilName.Trim(), @"\D", ""));
                            boolDataSaveStart = m_db.InsertTDRRCSDiagnosisHeaderDataGridViewDataInfo(strPlantName.Trim(), strSelectHogi.Trim(), strSelectOhDegree.Trim()
                                , strSelectPowerCabinet.Trim(), strSelectControlRodName.Trim(), strSelectCoilName.Trim(), strMeasurement_StartDate.Trim()
                                , strMeasurement_EndDate.Trim(), strImageDataPathFileName.Trim(), strTextDataPathFileName.Trim(), intSeqNumber);

                            frmMB.lblMessage.Text = "TDR-Meter 측정 완료";
                            frmMB.TopMost = true;
                            frmMB.ShowDialog();
                        }
                        else
                        {
                            frmMB.lblMessage.Text = "TDR-Meter 측정 실패";
                            frmMB.TopMost = true;
                            frmMB.ShowDialog();
                        }
                    }
                    else
                    {
                        frmMB.lblMessage.Text = "TDR-Meter 측정 실패";
                        frmMB.TopMost = true;
                        frmMB.ShowDialog();
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
            finally
            {
                // Control 활성화/비활성화
                SetControlEnabled(true);

                if (rboMeasurementItem_Stop.Checked)
                    rboMeasurementItem_Up.Checked = true;
                else if (rboMeasurementItem_Up.Checked)
                    rboMeasurementItem_Move.Checked = true;
                else
                    rboMeasurementItem_Stop.Checked = true;
            }

            System.Windows.Forms.Cursor.Current = Cursors.Default;
        }

        /// <summary>
        /// Control 활성화/비활성화
        /// </summary>
        /// <param name="boolEnabled"></param>
        private void SetControlEnabled(bool boolEnabled)
        {
            groupBox1.Enabled = boolEnabled;
            groupBox8.Enabled = boolEnabled;
            groupBox2.Enabled = boolEnabled;
            //groupBox10.Enabled = boolEnabled;
            //groupBox11.Enabled = boolEnabled;
            rcsPanel.SetControlEnabled(boolEnabled);
        }

        /// <summary>
        /// URL을 통해 이미지(Image) 다운로드(Download)해서 파일(File)로 저장(Save)
        /// 웹상의 이미지 URL을 알고 있을때.. 해당 URL로부터 이미지에 대한 데이터를 가져와 파일로 저장하는 C# 함수입니다. 
        /// 가끔 꼭 필요한 함수인데.. 필요할때 쉽게 찾아 볼 수 있도록 기록해 둡니다.
        /// </summary>
        /// <param name="uri"></param>
        /// <param name="fileName"></param>
        /// <returns></returns>
        private bool DownloadRemoteImageFile(string uri, string fileName, bool boolChecking)
        {
            bool boolResult = false;
            HttpWebRequest request = null;
            HttpWebResponse response = null;

            try
            {
                request = (HttpWebRequest)WebRequest.Create(uri);
                response = (HttpWebResponse)request.GetResponse();

                if (!boolChecking)
                {
                    bool bImage = response.ContentType.StartsWith("image", StringComparison.OrdinalIgnoreCase);

                    if ((response.StatusCode == HttpStatusCode.OK ||
                        response.StatusCode == HttpStatusCode.Moved ||
                        response.StatusCode == HttpStatusCode.Redirect) &&
                        bImage)
                    {
                        using (Stream inputStream = response.GetResponseStream())
                        using (Stream outputStream = File.OpenWrite(fileName))
                        {
                            byte[] buffer = new byte[8192];
                            int bytesRead;
                            do
                            {
                                bytesRead = inputStream.Read(buffer, 0, buffer.Length);
                                outputStream.Write(buffer, 0, bytesRead);
                            } while (bytesRead != 0);
                        }

                        boolResult = true;
                    }
                }
                else
                {
                    boolResult = true;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
            finally
            {
                if (response != null)
                {
                    response.Close();
                    response.Dispose();
                }

                if (request != null) request.Abort();
            }

            return boolResult;
        }

        /// <summary>
        /// 파일 다운로드
        /// </summary>
        /// <param name="url"></param>
        /// <param name="path"></param>
        /// <returns></returns>
        public bool fileDownload(string url, string path)
        {
            bool boolResult = false;
            try
            {
                WebClient webClient = new WebClient();

                webClient.DownloadFile(url, path);

                boolResult = true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            return boolResult;
        }

        /// <summary>
        /// 차수 Text Changed Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void teOhDegree2_TextChanged(object sender, EventArgs e)
        {
            pbTDRImage.Image = null;
        }
    }
}
