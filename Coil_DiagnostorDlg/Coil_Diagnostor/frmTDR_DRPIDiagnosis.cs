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
using Coil_Diagnostor.Function;


namespace Coil_Diagnostor
{
    public partial class frmTDR_DRPIDiagnosis : Form
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
        public bool boolControlRod = false;

        public string strSelectHogi = "";
        public string strSelectOhDegree = "";
        public string strSelectDRPIGroup = "";
        public string strSelectDRPIType = "";
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
        public IPAddress ipadress;

        private Panel DRPIrodPanel;
        private GroupBox gbStop;
        private GroupBox gbControl;

        public frmTDR_DRPIDiagnosis()
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
        private void frmTDR_DRPIDiagnosis_Load(object sender, EventArgs e)
        {
            boolFormLoad = true;

            strPlantName = Gini.GetValue("Device", "PlantName").Trim();

            string plantName = Gini.GetValue("Device", "PlantName");
            int rodCount = SetRodPanel(plantName);
            SetRodButtonDataBinding(DRPIrodPanel, rodCount);

            // ComboBox 데이터 설정
            SetComboBoxDataBinding();

            // 초기화
            SetControlInitialize();

            SetStopAndControlButtonInitialize(gbStop);
            SetStopAndControlButtonInitialize(gbControl);

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
                    returnVal = 33;
                    break;
                case "한빛 1발전소":
                    DRPIrodPanel = pnlRod52;
                    gbStop = gbStop52;
                    gbControl = gbControl52;
                    returnVal = 52;
                    break;
                default:
                    DRPIrodPanel = pnlRod52;
                    gbStop = gbStop52;
                    gbControl = gbControl52;
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
            if (Gini.GetValue("DRPI", "SelectDRPI_TDRHogi").Trim() == "")
            {
                // 최초 실행 시 기본 호기 선택
                cboHogi.SelectedIndex = 0;
            }
            else
            {
                // 직전 실행 시 기본 호기 선택
                cboHogi.SelectedItem = Gini.GetValue("DRPI", "SelectDRPI_TDRHogi").Trim();
            }
        }

        /// <summary>
        /// 초기화
        /// </summary>
        private void SetControlInitialize()
        {
            // 코일 그룹
            rboCoilA.Checked = true;
            strSelectDRPIGroup = rboCoilA.Text.Trim();
            strSelectDRPIType = "";
        }

        /// <summary>
        /// 정지용/제어용 Button 초기화
        /// </summary>
        private void SetStopAndControlButtonInitialize(GroupBox _gb)
        {
            foreach (Control c in _gb.Controls)
            {
                if (c.GetType().Name.Trim() == "Button")
                {
                    c.ForeColor = System.Drawing.Color.Black;
                    c.BackgroundImage = global::Coil_Diagnostor.Properties.Resources.버튼_4;
                    c.Enabled = true;
                }
            }
        }

        /// <summary>
        /// Form Closing Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmTDR_DRPIDiagnosis_FormClosing(object sender, FormClosingEventArgs e)
        {
            Gini.SetValue("DRPI", "SelectDRPI_TDRHogi", cboHogi.SelectedItem.ToString().Trim());
            Gini.SetValue("DRPI", "SelectDRPI_TDROHDegree", teOhDegree2.Text.Trim());
        }

        /// <summary>
        /// 정지용/제어용 Button Click Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnStopControl_Click(object sender, EventArgs e)
        {
            try
            {
                pbTDRImage.Image = null;

                // Button 선택 여부에 따라 글자 색변경 (선택 시 White, 선택 취소 시 Black)
                if (((Button)sender).ForeColor == System.Drawing.Color.Black)
                {
                    ((Button)sender).ForeColor = System.Drawing.Color.White;
                    ((Button)sender).BackgroundImage = global::Coil_Diagnostor.Properties.Resources.버튼_3;
                    strSelectControlRodName = ((Button)sender).Text.Trim();

                    // 전력함 표시 Button Color
                    SetControlButtonColor(gbStop, ((Button)sender));
                    SetControlButtonColor(gbControl, ((Button)sender));
                    
                    // 제어용/정지용에 따라 설정
                    if (rboCoilA.Checked)
                    {
                        strSelectDRPIGroup = rboCoilA.Text.Trim();
                        SetListBox(strSelectDRPIGroup);
                    }
                    else
                    {
                        strSelectDRPIGroup = rboCoilB.Text.Trim();
                        SetListBox(strSelectDRPIGroup);
                    }
                }
                else
                {
                    ((Button)sender).ForeColor = System.Drawing.Color.Black;
                    ((Button)sender).BackgroundImage = global::Coil_Diagnostor.Properties.Resources.버튼_4;

                    strSelectControlRodName = "";
                    strSelectDRPIType = "";
                    lboxCoilName.Items.Clear();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
        }

        /// <summary>
        /// 정지용/제어용 Button Color
        /// </summary>
        private void SetControlButtonColor(GroupBox _gb, Button _btn)
        {
            foreach (Control c in _gb.Controls)
            {
                if (c.GetType().Name.Trim() == "Button")
                {
                    if (c.Name.Trim() != _btn.Name.Trim())
                    {
                        c.ForeColor = System.Drawing.Color.Black;
                        c.BackgroundImage = global::Coil_Diagnostor.Properties.Resources.버튼_4;
                    }
                    else if (c.Name.Trim() == _btn.Name.Trim())
                    {
                        strSelectDRPIType = _gb.Text.Trim();
                    }
                }
            }
        }

        /// <summary>
        /// 호기 Selected Index Changed Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cboHogi_SelectedIndexChanged(object sender, EventArgs e)
        {
            pbTDRImage.Image = null;

            teOhDegree2.Text = Gini.GetValue("DRPI", "SelectDRPI_TDROHDegree").Trim();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void rboCoil_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                pbTDRImage.Image = null;

                // 제어용/정지용에 따라 설정
                if (rboCoilA.Checked)
                {
                    strSelectDRPIGroup = rboCoilA.Text.Trim();
                    SetListBox(strSelectDRPIGroup);
                }
                else
                {
                    strSelectDRPIGroup = rboCoilB.Text.Trim();
                    SetListBox(strSelectDRPIGroup);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
        }

        /// <summary>
        /// 제어용/정지용에 따라 설정
        /// </summary>
        /// <param name="_strCoilName"></param>
        private void SetListBox(string _strCoilName)
        {
            try
            {
                string[] strCoilName;

                if (strSelectDRPIType.Trim() == "정지용")
                    strCoilName = Gini.GetValue("DRPI", "Stop_Item").Split(',');
                else
                    strCoilName = Gini.GetValue("DRPI", "Control_Item").Split(',');

                lboxCoilName.Items.Clear();

                for (int i = 0; i < strCoilName.Length; i++)
                {
                    lboxCoilName.Items.Add(_strCoilName.Trim() + Regex.Replace(strCoilName[i].Trim(), @"\D", ""));
                }

                lboxCoilName.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
        }

        /// <summary>
        /// 코일Name 선택 ListBox Selected Index Changed Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void lboxCoilName_SelectedIndexChanged(object sender, EventArgs e)
        {
            pbTDRImage.Image = null;

            strSelectCoilName = lboxCoilName.SelectedItem.ToString().Trim();
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
        /// 측정 Button Evnet
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnMeasurementStart_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;

            int intMeasurementIndex = 0;

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

                // 코일그룹 선택 체크
                if (strSelectDRPIGroup.Trim() == "")
                {
                    frmMB.lblMessage.Text = "코일그룹을 선택하십시오.";
                    frmMB.TopMost = true;
                    frmMB.ShowDialog();
                    return;
                }

                // 제어봉 선택 체크
                if (strSelectControlRodName.Trim() == "" || strSelectDRPIType.Trim() == "")
                {
                    frmMB.lblMessage.Text = "제어봉을 선택하십시오.";
                    frmMB.TopMost = true;
                    frmMB.ShowDialog();
                    return;
                }

                // 코일명 선택 체크
                if (strSelectCoilName.Trim() == "")
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
                    intMeasurementIndex = lboxCoilName.SelectedIndex;

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
                    strImageFileName = string.Format("TDR_DRPI {0}_{1}_{2}_{3}_{4}_{5}_{6}.bmp", strSelectHogi, strSelectOhDegree
                        , strSelectDRPIGroup, strSelectDRPIType, strSelectControlRodName, strSelectCoilName, dataTime.ToString("yyyyMMdd_HHmmss"));

                    strTextFileName = string.Format("TDR_DRPI {0}_{1}_{2}_{3}_{4}_{5}_{6}.csv", strSelectHogi, strSelectOhDegree
                        , strSelectDRPIGroup, strSelectDRPIType, strSelectControlRodName, strSelectCoilName, dataTime.ToString("yyyyMMdd_HHmmss"));

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
                            if ((m_db.GetTDRDRPIDiagnosisHeaderDataGridViewDataCount(strPlantName.Trim(), strSelectHogi.Trim(), strSelectOhDegree.Trim()
                                , strSelectDRPIGroup.Trim(), strSelectDRPIType.Trim(), strSelectControlRodName.Trim(), strSelectCoilName.Trim())) > 0)
                            {
                                // 기존 데이터 삭제
                                boolDataSaveStart = m_db.DeleteTDRDRPIDiagnosisHeaderDataGridViewDataInfo(strPlantName.Trim(), strSelectHogi.Trim(), strSelectOhDegree.Trim()
                                , strSelectDRPIGroup.Trim(), strSelectDRPIType.Trim().Trim(), strSelectControlRodName.Trim(), strSelectCoilName.Trim());
                            }

                            // 데이터 신규 저장
                            int intSeqNumber = Regex.Replace(strSelectCoilName.Trim(), @"\D", "") == "" ? 1 : Convert.ToInt32(Regex.Replace(strSelectCoilName.Trim(), @"\D", ""));
                            boolDataSaveStart = m_db.InsertTDRDRPIDiagnosisHeaderDataGridViewDataInfo(strPlantName.Trim(), strSelectHogi.Trim(), strSelectOhDegree.Trim()
                                , strSelectDRPIGroup.Trim(), strSelectDRPIType.Trim().Trim(), strSelectControlRodName.Trim(), strSelectCoilName.Trim()
                                , strMeasurement_StartDate.Trim(), strMeasurement_EndDate.Trim(), strImageDataPathFileName.Trim(), strTextDataPathFileName.Trim(), intSeqNumber);

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
                else
                {
                    frmMB.lblMessage.Text = "이미";
                    frmMB.TopMost = true;
                    frmMB.ShowDialog();
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

                if (lboxCoilName.Items.Count > (intMeasurementIndex + 1))
                    lboxCoilName.SelectedIndex = intMeasurementIndex + 1;
            }

            System.Windows.Forms.Cursor.Current = Cursors.Default;
        }

        /// <summary>
        /// Control 활성화/비활성화
        /// </summary>
        /// <param name="boolEnabled"></param>
        private void SetControlEnabled(bool boolEnabled)
        {
            gbStop.Enabled = boolEnabled;
            gbControl.Enabled = boolEnabled;
            groupBox1.Enabled = boolEnabled;
            groupBox8.Enabled = boolEnabled;
            groupBox3.Enabled = boolEnabled;
            groupBox14.Enabled = boolEnabled;
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
