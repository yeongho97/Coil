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
    public partial class frmSetOffset : Form
    {
        protected Function.FunctionMeasureProcess m_MeasureProcess = new Function.FunctionMeasureProcess();
        protected Function.FunctionDataControl m_db = new Function.FunctionDataControl();
        protected frmMessageBox frmMB = new frmMessageBox();
        protected string strPlantName = "";
        protected bool boolMessageOK = false;
        protected bool boolFormLoad = false;
        public bool boolMeasurementStart = false;
        public bool boolMeasurementStop = false;
        protected CheckBox NormalAllCheck = new CheckBox();
        protected CheckBox WheatstoneAllCheck = new CheckBox();
        protected bool isNormalCheck = true;
        protected bool isWheatstoneCheck = true;
        public static Thread threadMeasurementStart;
        public static Thread threadMeasurementStop;
        public int intTabSelectIndex = 0;
        
        public decimal dcmFrequency = 120M;
        public decimal dcmVoltageLevel = 1M;
        public string strDAQDeviceName = "";

        public Thread _threadMeasurementStartState;
        public Thread _thread;

        public frmSetOffset()
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
                case Keys.F2: // 측정 버튼
                    btnMeasurement.PerformClick();
                    break;
                case Keys.F3: // 정지 버튼
                    btnStop.PerformClick();
                    break;
                case Keys.F4: // 저장 버튼
                    btnSave.PerformClick();
                    break;
                case Keys.F12: // 닫기 버튼
                    btnClose.PerformClick();
                    break;
            }

            return base.ProcessCmdKey(ref msg, keyData);
        }

        /// <summary>
        /// form Load
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmSetOffset_Load(object sender, EventArgs e)
        {
            boolFormLoad = true;

            strPlantName = Gini.GetValue("Device", "PlantName").Trim();

            // ComboBox 데이터 설정
            SetComboBoxDataBinding();

            // 일반모드 DataGridView Columns 설정
            SetNormalModeDataGridViewInitialize();
            cboDAMRelay_NormalMode.SelectedIndex = 0;

            // 휘스톤모드 DataGridView Columns 설정
            SetWheatstoneModeDataGridViewInitialize();
            cboDAMRelay_WheatstoneMode.SelectedIndex = 0;

            // LCR-Meter 접속 체크
            if (m_MeasureProcess.m_LCRMeter.Initialize(Gini.GetValue("Device", "LCRMeter_Addr").Trim()))
            {
                // Button Enabled 설정
                SetButtonEnabled(true);
            }
            else
            {
                btnMeasurement.Enabled = false;
                btnStop.Enabled = false;
                btnSave.Enabled = false;
            }
            
            boolFormLoad = false;
        }

        /// <summary>
        /// ComboBox 데이터 설정
        /// </summary>
        private void SetComboBoxDataBinding()
        {
            // 발전소 호기
            string[] strData = Gini.GetValue("Combo", "DAMRelayItem").Split(',');

            cboDAMRelay_NormalMode.Items.Clear();
            cboDAMRelay_WheatstoneMode.Items.Clear();

            for (int i = 0; i < strData.Length; i++)
            {
                cboDAMRelay_NormalMode.Items.Add(strData[i].Trim());
                cboDAMRelay_WheatstoneMode.Items.Add(strData[i].Trim());
            }
        }

        /// <summary>
        /// Button Enabled 설정
        /// </summary>
        /// <param name="_boolEnabled"></param>
        private void SetButtonEnabled(bool _boolStop)
        {
            if (_boolStop)
            {
                btnMeasurement.Enabled = true;      // 측정
                btnStop.Enabled = false;            // 정지
                btnSave.Enabled = true;             // 저장
                btnClose.Enabled = true;            // 닫기
            }
            else
            {
                btnMeasurement.Enabled = false;     // 측정
                btnStop.Enabled = true;             // 정지
                btnSave.Enabled = false;            // 저장
                btnClose.Enabled = false;           // 닫기
            }
        }

        /// <summary>
        /// 일반모드 DataGridView Columns 설정
        /// </summary>
        private void SetNormalModeDataGridViewInitialize()
        {
            try
            {
                //----- DataGridView Column 추가
                DataGridViewCheckBoxColumn Column1 = new DataGridViewCheckBoxColumn();
                DataGridViewTextBoxColumn Column2 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn Column3 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn Column4 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn Column5 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn Column6 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn Column7 = new DataGridViewTextBoxColumn();

                Column1.HeaderText = "";
                Column1.Name = "ChekBox";
                Column1.Width = 30;
                Column1.ReadOnly = true;
                Column1.Visible = true;
                Column1.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column1.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column2.HeaderText = "채널";
                Column2.Name = "ChannelName";
                Column2.Width = 90;
                Column2.ReadOnly = true;
                Column2.Visible = true;
                Column2.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column2.SortMode = DataGridViewColumnSortMode.NotSortable;
                Column2.DefaultCellStyle.BackColor = Color.LightGray;

                Column3.HeaderText = "채널핀번호";
                Column3.Name = "ChannelPinNo";
                Column3.Width = 120;
                Column3.ReadOnly = true;
                Column3.Visible = false;
                Column3.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column3.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column4.HeaderText = "DC 저항(Ω)";
                Column4.Name = "DC_Resistance";
                Column4.Width = 120;
                Column4.ReadOnly = false;
                Column4.Visible = true;
                Column4.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column4.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column5.HeaderText = "AC 저항(Ω)";
                Column5.Name = "AC_Resistance";
                Column5.Width = 120;
                Column5.ReadOnly = false;
                Column5.Visible = true;
                Column5.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column5.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column6.HeaderText = "인덕턴스(mH)";
                Column6.Name = "Inductance";
                Column6.Width = 120;
                Column6.ReadOnly = false;
                Column6.Visible = true;
                Column6.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column6.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column7.HeaderText = "캐패시턴스(uF)";
                Column7.Name = "Capacitance";
                Column7.Width = 100;
                Column7.ReadOnly = false;
                Column7.Visible = false;
                Column7.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column7.SortMode = DataGridViewColumnSortMode.NotSortable;

                dgvNormalMode.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] { Column1, Column2, Column3, Column4, Column5, Column6, Column7 });

                dgvNormalMode.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing;
                dgvNormalMode.ColumnHeadersHeight = 40;
                dgvNormalMode.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dgvNormalMode.RowHeadersWidth = 40;

                // DataGridView 기타 세팅
                dgvNormalMode.AllowUserToAddRows = false; // Row 추가 기능 미사용
                dgvNormalMode.AllowUserToDeleteRows = false; // Row 삭제 기능 미사용

                // CheckBox 세팅
                NormalAllCheck.Name = "allCheck";
                NormalAllCheck.CheckedChanged += new EventHandler(NormalCheckClick);
                NormalAllCheck.Size = new Size(13, 13);
                NormalAllCheck.Location = new Point((dgvNormalMode.Columns[0].Width / 2) - (NormalAllCheck.Width / 2), (dgvNormalMode.ColumnHeadersHeight / 2) - (NormalAllCheck.Height / 2));
                dgvNormalMode.Controls.Add(NormalAllCheck); // DataGridView에 CheckBox 추가( 헤더용 )
                NormalAllCheck.Checked = true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
        }

        /// <summary>
        /// 일반모드 DataGridView의 모든 CheckBoxCell Checked값 적용
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void NormalCheckClick(object sender, EventArgs e)
        {
            if (isNormalCheck)
            {
                if (NormalAllCheck.Checked)
                    for (int i = 0; i < dgvNormalMode.Rows.Count; i++)
                        dgvNormalMode.Rows[i].Cells[0].Value = true;
                else
                    for (int i = 0; i < dgvNormalMode.Rows.Count; i++)
                        dgvNormalMode.Rows[i].Cells[0].Value = false;

                dgvNormalMode.EndEdit(DataGridViewDataErrorContexts.Commit); // << 이거 안할경우 선택된 Cell이 CheckBoxCell일 경우 변화가 없는것처럼 보임
            }
        }

        /// <summary>
        /// 셀의 체크박스가 클릭될 때 셀안의 체크박스 Checked값 변경
        /// 전부 클릭이 되어있는지 확인해서 AIAllCheck의 Checked값 적용
        /// isAICheck는 AIAllCheck.Checked에 값을 적용할떄 AllCheckClick 함수가 실행되지 않기 위함
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgvNormalMode_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dgvNormalMode.RowCount < 1) return;

            isNormalCheck = false;

            if (e.RowIndex >= 0 && e.ColumnIndex == 0)
            {
                dgvNormalMode.Rows[e.RowIndex].Cells[0].Value = !(bool)dgvNormalMode.Rows[e.RowIndex].Cells[0].Value;
                NormalAllCheck.Checked = true;

                for (int i = 0; i < dgvNormalMode.Rows.Count; i++)
                {
                    if (!(bool)dgvNormalMode.Rows[i].Cells[0].Value)
                    {
                        NormalAllCheck.Checked = false;
                        break;
                    }
                }
            }

            isNormalCheck = true;
        }

        /// <summary>
        /// 휘스톤모드 DataGridView Columns 설정
        /// </summary>
        private void SetWheatstoneModeDataGridViewInitialize()
        {
            try
            {
                //----- DataGridView Column 추가
                DataGridViewCheckBoxColumn Column1 = new DataGridViewCheckBoxColumn();
                DataGridViewTextBoxColumn Column2 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn Column3 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn Column4 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn Column5 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn Column6 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn Column7 = new DataGridViewTextBoxColumn();

                Column1.HeaderText = "";
                Column1.Name = "ChekBox";
                Column1.Width = 30;
                Column1.ReadOnly = false;
                Column1.Visible = true;
                Column1.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column1.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column2.HeaderText = "채널";
                Column2.Name = "ChannelName";
                Column2.Width = 90;
                Column2.ReadOnly = true;
                Column2.Visible = true;
                Column2.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column2.SortMode = DataGridViewColumnSortMode.NotSortable;
                Column2.DefaultCellStyle.BackColor = Color.LightGray;

                Column3.HeaderText = "채널핀번호";
                Column3.Name = "ChannelPinNo";
                Column3.Width = 120;
                Column3.ReadOnly = false;
                Column3.Visible = false;
                Column3.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column3.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column4.HeaderText = "DC 저항(Ω)";
                Column4.Name = "DC_Resistance";
                Column4.Width = 120;
                Column4.ReadOnly = false;
                Column4.Visible = true;
                Column4.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column4.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column5.HeaderText = "AC 저항(Ω)";
                Column5.Name = "AC_Resistance";
                Column5.Width = 120;
                Column5.ReadOnly = false;
                Column5.Visible = true;
                Column5.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column5.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column6.HeaderText = "인덕턴스(mH)";
                Column6.Name = "Inductance";
                Column6.Width = 120;
                Column6.ReadOnly = false;
                Column6.Visible = true;
                Column6.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column6.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column7.HeaderText = "캐패시턴스(uF)";
                Column7.Name = "Capacitance";
                Column7.Width = 100;
                Column7.ReadOnly = false;
                Column7.Visible = false;
                Column7.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column7.SortMode = DataGridViewColumnSortMode.NotSortable;

                dgvWheatstoneMode.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] { Column1, Column2, Column3, Column4, Column5, Column6, Column7 });

                dgvWheatstoneMode.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing;
                dgvWheatstoneMode.ColumnHeadersHeight = 40;
                dgvWheatstoneMode.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dgvWheatstoneMode.RowHeadersWidth = 40;

                // DataGridView 기타 세팅
                dgvWheatstoneMode.AllowUserToAddRows = false; // Row 추가 기능 미사용
                dgvWheatstoneMode.AllowUserToDeleteRows = false; // Row 삭제 기능 미사용

                // CheckBox 세팅
                WheatstoneAllCheck.Name = "allCheck";
                WheatstoneAllCheck.CheckedChanged += new EventHandler(WheatstoneCheckClick);
                WheatstoneAllCheck.Size = new Size(13, 13);
                WheatstoneAllCheck.Location = new Point((dgvWheatstoneMode.Columns[0].Width / 2) - (WheatstoneAllCheck.Width / 2), (dgvWheatstoneMode.ColumnHeadersHeight / 2) - (WheatstoneAllCheck.Height / 2));
                dgvWheatstoneMode.Controls.Add(WheatstoneAllCheck); // DataGridView에 CheckBox 추가( 헤더용 )
                WheatstoneAllCheck.Checked = true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
        }

        /// <summary>
        /// 일반모드 DataGridView의 모든 CheckBoxCell Checked값 적용
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void WheatstoneCheckClick(object sender, EventArgs e)
        {
            if (isWheatstoneCheck)
            {
                if (WheatstoneAllCheck.Checked)
                    for (int i = 0; i < dgvWheatstoneMode.Rows.Count; i++)
                        dgvWheatstoneMode.Rows[i].Cells[0].Value = true;
                else
                    for (int i = 0; i < dgvWheatstoneMode.Rows.Count; i++)
                        dgvWheatstoneMode.Rows[i].Cells[0].Value = false;

                dgvWheatstoneMode.EndEdit(DataGridViewDataErrorContexts.Commit); // << 이거 안할경우 선택된 Cell이 CheckBoxCell일 경우 변화가 없는것처럼 보임
            }
        }

        /// <summary>
        /// 셀의 체크박스가 클릭될 때 셀안의 체크박스 Checked값 변경
        /// 전부 클릭이 되어있는지 확인해서 AIAllCheck의 Checked값 적용
        /// isAICheck는 AIAllCheck.Checked에 값을 적용할떄 AllCheckClick 함수가 실행되지 않기 위함
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgvWheatstoneMode_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dgvWheatstoneMode.RowCount < 1) return;

            isWheatstoneCheck = false;

            if (e.RowIndex >= 0 && e.ColumnIndex == 0)
            {
                dgvWheatstoneMode.Rows[e.RowIndex].Cells[0].Value = !(bool)dgvWheatstoneMode.Rows[e.RowIndex].Cells[0].Value;
                WheatstoneAllCheck.Checked = true;

                for (int i = 0; i < dgvWheatstoneMode.Rows.Count; i++)
                {
                    if (!(bool)dgvWheatstoneMode.Rows[i].Cells[0].Value)
                    {
                        WheatstoneAllCheck.Checked = false;
                        break;
                    }
                }
            }

            isWheatstoneCheck = true;
        }

        /// <summary>
        /// 휘스톤모드 DataGridView Data 가져오기
        /// </summary>
        private void GetWheatstonModeDataGridViewData()
        {
            try
            {

            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
        }

        /// <summary>
        /// 일반모드 DAM Relay ComboBox Selected Index Changed Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cboDAMRelay_NormalMode_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (boolMeasurementStart) return;

            // 일반모드 DataGridView Data 가져오기
            string strDAMRelayName = cboDAMRelay_NormalMode.SelectedItem == null ? "" : cboDAMRelay_NormalMode.SelectedItem.ToString().Trim();
            GetDataGridViewData("일반모드", strDAMRelayName);
        }

        /// <summary>
        /// 휘스톤모드 DAM Relay ComboBox Selected Index Changed Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cboDAMRelay_WheatstoneMode_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (boolMeasurementStart) return;

            // 휘스톤모드 DataGridView Data 가져오기
            string strDAMRelayName = cboDAMRelay_WheatstoneMode.SelectedItem == null ? "" : cboDAMRelay_WheatstoneMode.SelectedItem.ToString().Trim();
            GetDataGridViewData("휘스톤모드", strDAMRelayName);
        }

        /// <summary>
        /// 일반모드/휘스톤모드 DataGridView Data 가져오기
        /// </summary>
        private void GetDataGridViewData(string _strMeasurementMode, string strDAMRelayName)
        {
            try
            {
                DataTable dt = new DataTable();
                dt = m_db.GetSetOffsetDataGridViewDataInfo(strPlantName.Trim(), _strMeasurementMode.Trim(), strDAMRelayName);

                if (_strMeasurementMode.Trim() == "일반모드")
                    dgvNormalMode.Rows.Clear();
                else
                    dgvWheatstoneMode.Rows.Clear();

                if (dt != null && dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        if (_strMeasurementMode.Trim() == "일반모드")
                        {
                            int iRow = dgvNormalMode.Rows.Add();

                            dgvNormalMode.Rows[iRow].Cells["ChekBox"].Value = true;
                            dgvNormalMode.Rows[iRow].Cells["ChannelName"].Value = dt.Rows[i]["ChannelName"].ToString().Trim();
                            dgvNormalMode.Rows[iRow].Cells["DC_Resistance"].Value = dt.Rows[i]["DC_Resistance"].ToString().Trim();
                            dgvNormalMode.Rows[iRow].Cells["AC_Resistance"].Value = dt.Rows[i]["AC_Resistance"].ToString().Trim();
                            dgvNormalMode.Rows[iRow].Cells["Inductance"].Value = dt.Rows[i]["Inductance"].ToString().Trim();
                            dgvNormalMode.Rows[iRow].Cells["Capacitance"].Value = dt.Rows[i]["Capacitance"].ToString().Trim();
                        }
                        else
                        {
                            int iRow = dgvWheatstoneMode.Rows.Add();

                            dgvWheatstoneMode.Rows[iRow].Cells["ChekBox"].Value = true;
                            dgvWheatstoneMode.Rows[iRow].Cells["ChannelName"].Value = dt.Rows[i]["ChannelName"].ToString().Trim();
                            dgvWheatstoneMode.Rows[iRow].Cells["DC_Resistance"].Value = dt.Rows[i]["DC_Resistance"].ToString().Trim();
                            dgvWheatstoneMode.Rows[iRow].Cells["AC_Resistance"].Value = dt.Rows[i]["AC_Resistance"].ToString().Trim();
                            dgvWheatstoneMode.Rows[iRow].Cells["Inductance"].Value = dt.Rows[i]["Inductance"].ToString().Trim();
                            dgvWheatstoneMode.Rows[iRow].Cells["Capacitance"].Value = dt.Rows[i]["Capacitance"].ToString().Trim();
                        }
                    }
                }
                else
                {
                    if (_strMeasurementMode.Trim() == "일반모드")
                    {
                        frmMB.lblMessage.Text = "일반모드 데이터가 없습니다.";
                        frmMB.TopMost = true;
                        frmMB.ShowDialog();
                    }
                    else
                    {
                        frmMB.lblMessage.Text = "휘스톤모드 데이터가 없습니다.";
                        frmMB.TopMost = true;
                        frmMB.ShowDialog();
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

        private void btnSave_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;

            bool boolSave = false;

            try
            {
                btnSave.ForeColor = System.Drawing.Color.Blue;
                string strMessage = "";

                if (tabControl1.SelectedIndex == 0)
                {
                    strMessage = "일반모드";

                    if (dgvNormalMode.Rows.Count <= 0)
                    {
                        frmMB.lblMessage.Text = "일반모드 저장할 데이터가 없습니다.";
                        frmMB.TopMost = true;
                        frmMB.ShowDialog();
                        return;
                    }
                    else
                        boolSave = SaveNormalModeData();
                }
                else
                {
                    strMessage = "휘스톤모드";

                    if (dgvNormalMode.Rows.Count <= 0)
                    {
                        frmMB.lblMessage.Text = "휘스톤모드 저장할 데이터가 없습니다.";
                        frmMB.TopMost = true;
                        frmMB.ShowDialog();
                        return;
                    }
                    else
                        boolSave = SaveWheatstoneModeData();
                }

                if (boolSave)
                {
                    frmMB.lblMessage.Text = strMessage.Trim() + " 데이터를 저장 완료";
                    frmMB.TopMost = true;
                    frmMB.ShowDialog();
                }
                else
                {
                    frmMB.lblMessage.Text = strMessage.Trim() + " 데이터 저장 중 실패";
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
                btnSave.ForeColor = System.Drawing.Color.White;
            }

            System.Windows.Forms.Cursor.Current = Cursors.Default;
        }

        /// <summary>
        /// 일반모드 데이터 저장 처리
        /// </summary>
        /// <returns></returns>
        private bool SaveNormalModeData()
        {
            bool boolResult = false;

            try
            {
                string strMeasurementMode = "", strDAMRelayName = "", strChannelName = "", strChannelPinNo = "";
                decimal dcmDC_Resistance = 0.000M, dcmAC_Resistance = 0.000M, dcmInductance = 0.000M, dcmCapacitance = 0.000000M;

                strMeasurementMode = "일반모드";
                strDAMRelayName = cboDAMRelay_NormalMode.SelectedItem == null ? "" : cboDAMRelay_NormalMode.SelectedItem.ToString().Trim();

                for (int i = 0; i < dgvNormalMode.RowCount; i++)
                {
                    strChannelName = dgvNormalMode.Rows[i].Cells["ChannelName"].Value == null ? "" : dgvNormalMode.Rows[i].Cells["ChannelName"].Value.ToString().Trim();
                    strChannelPinNo = dgvNormalMode.Rows[i].Cells["ChannelPinNo"].Value == null ? "" : dgvNormalMode.Rows[i].Cells["ChannelPinNo"].Value.ToString().Trim();
                    dcmDC_Resistance = dgvNormalMode.Rows[i].Cells["DC_Resistance"].Value == null 
                        || dgvNormalMode.Rows[i].Cells["DC_Resistance"].Value.ToString().Trim() == "" 
                        ? 0.000M : Convert.ToDecimal(dgvNormalMode.Rows[i].Cells["DC_Resistance"].Value.ToString().Trim());
                    dcmAC_Resistance = dgvNormalMode.Rows[i].Cells["AC_Resistance"].Value == null
                        || dgvNormalMode.Rows[i].Cells["AC_Resistance"].Value.ToString().Trim() == "" 
                        ? 0.000M : Convert.ToDecimal(dgvNormalMode.Rows[i].Cells["AC_Resistance"].Value.ToString().Trim());
                    dcmInductance = dgvNormalMode.Rows[i].Cells["Inductance"].Value == null
                        || dgvNormalMode.Rows[i].Cells["Inductance"].Value.ToString().Trim() == "" 
                        ? 0.000M : Convert.ToDecimal(dgvNormalMode.Rows[i].Cells["Inductance"].Value.ToString().Trim());
                    dcmCapacitance = dgvNormalMode.Rows[i].Cells["Capacitance"].Value == null
                        || dgvNormalMode.Rows[i].Cells["Capacitance"].Value.ToString().Trim() == "" 
                        ? 0.000M : Convert.ToDecimal(dgvNormalMode.Rows[i].Cells["Capacitance"].Value.ToString().Trim());

                    if ((m_db.GetSetOffsetDataGridViewDataCount(strPlantName, strMeasurementMode, strDAMRelayName, strChannelName)) > 0)
                    {
                        boolResult = m_db.UpDateSetOffsetDataGridViewDataInfo(strPlantName, strMeasurementMode, strDAMRelayName, strChannelName
                            , dcmDC_Resistance, dcmAC_Resistance, dcmInductance, dcmCapacitance);
                    }
                    else
                    {
                        boolResult = m_db.InsertSetOffsetDataGridViewDataInfo(strPlantName, strMeasurementMode, strDAMRelayName, strChannelName
                            , dcmDC_Resistance, dcmAC_Resistance, dcmInductance, dcmCapacitance);
                    }                       
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            return boolResult;
        }

        /// <summary>
        /// 휘스톤모드 데이터 저장 처리
        /// </summary>
        /// <returns></returns>
        private bool SaveWheatstoneModeData()
        {
            bool boolResult = false;

            try
            {
                string strMeasurementMode = "", strDAMRelayName = "", strChannelName = "", strChannelPinNo = "";
                decimal dcmDC_Resistance = 0.000M, dcmAC_Resistance = 0.000M, dcmInductance = 0.000M, dcmCapacitance = 0.000000M;

                strMeasurementMode = "휘스톤모드";
                strDAMRelayName = cboDAMRelay_WheatstoneMode.SelectedItem == null ? "" : cboDAMRelay_WheatstoneMode.SelectedItem.ToString().Trim();

                for (int i = 0; i < dgvWheatstoneMode.RowCount; i++)
                {
                    strChannelName = dgvWheatstoneMode.Rows[i].Cells["ChannelName"].Value == null ? "" : dgvWheatstoneMode.Rows[i].Cells["ChannelName"].Value.ToString().Trim();
                    strChannelPinNo = dgvWheatstoneMode.Rows[i].Cells["ChannelPinNo"].Value == null ? "" : dgvWheatstoneMode.Rows[i].Cells["ChannelPinNo"].Value.ToString().Trim();
                    dcmDC_Resistance = dgvWheatstoneMode.Rows[i].Cells["DC_Resistance"].Value == null
                        || dgvWheatstoneMode.Rows[i].Cells["DC_Resistance"].Value.ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(dgvWheatstoneMode.Rows[i].Cells["DC_Resistance"].Value.ToString().Trim());
                    dcmAC_Resistance = dgvWheatstoneMode.Rows[i].Cells["AC_Resistance"].Value == null
                        || dgvWheatstoneMode.Rows[i].Cells["AC_Resistance"].Value.ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(dgvWheatstoneMode.Rows[i].Cells["AC_Resistance"].Value.ToString().Trim());
                    dcmInductance = dgvWheatstoneMode.Rows[i].Cells["Inductance"].Value == null
                        || dgvWheatstoneMode.Rows[i].Cells["Inductance"].Value.ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(dgvWheatstoneMode.Rows[i].Cells["Inductance"].Value.ToString().Trim());
                    dcmCapacitance = dgvWheatstoneMode.Rows[i].Cells["Capacitance"].Value == null
                        || dgvWheatstoneMode.Rows[i].Cells["Capacitance"].Value.ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(dgvWheatstoneMode.Rows[i].Cells["Capacitance"].Value.ToString().Trim());

                    if ((m_db.GetSetOffsetDataGridViewDataCount(strPlantName, strMeasurementMode, strDAMRelayName, strChannelName)) > 0)
                    {
                        boolResult = m_db.UpDateSetOffsetDataGridViewDataInfo(strPlantName, strMeasurementMode, strDAMRelayName, strChannelName
                            , dcmDC_Resistance, dcmAC_Resistance, dcmInductance, dcmCapacitance);
                    }
                    else
                    {
                        boolResult = m_db.InsertSetOffsetDataGridViewDataInfo(strPlantName, strMeasurementMode, strDAMRelayName, strChannelName
                            , dcmDC_Resistance, dcmAC_Resistance, dcmInductance, dcmCapacitance);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            return boolResult;
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
        /// 측정 Button Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnMeasurement_Click(object sender, EventArgs e)
        {
            if (boolMeasurementStart) return;

            try
            {
                boolMeasurementStart = true;
                btnMeasurement.ForeColor = System.Drawing.Color.Blue;

                // Button Enabled 설정
                SetButtonEnabled(false);

                // 측정 시작
                threadMeasurementStart = new Thread(new ThreadStart(threadSetOffsetMeasurementStart));
                threadMeasurementStart.Start();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
            finally
            {
                btnMeasurement.ForeColor = System.Drawing.Color.White;
            }
        }

        /// <summary>
        /// 일반모드/휘스톤모드 Measurement Start Thread
        /// </summary>
        public void threadSetOffsetMeasurementStart()
        {
            bool boolWheatstoneMode = tabControl1.SelectedIndex == 0 ? false : true;

            if (!boolWheatstoneMode)
                threadSetOffsetMeasurementValue(boolWheatstoneMode, dgvNormalMode);
            else
                threadSetOffsetMeasurementValue(boolWheatstoneMode, dgvWheatstoneMode);
        }

        public void threadSetOffsetMeasurementValue(bool _boolWheatstoneMode, DataGridView _dgv)
        {
            bool boolResult = false;

            strDAQDeviceName = Gini.GetValue("Device", "DAQDeviceName").Trim();
            m_MeasureProcess.DigitalDAQ_CloseChannel();

            try
            {
                decimal dcmMeasurementValue = 0M;
                string strChannelName = "", strNormalModeDAQPinMap = "", strWheatstoneModeDAQPinMap = "", strDAMRelayName = "";
                int intSleep = 500;

                if (!_boolWheatstoneMode)
                    strDAMRelayName = cboDAMRelay_NormalMode.SelectedItem == null ? "DAM 1" : cboDAMRelay_NormalMode.SelectedItem.ToString().Trim();
                else
                    strDAMRelayName = cboDAMRelay_WheatstoneMode.SelectedItem == null ? "DAM 1" : cboDAMRelay_WheatstoneMode.SelectedItem.ToString().Trim();

                // LCR-Meter Range 설정
                m_MeasureProcess.functionLCRMeterRangeSetting(Function.FunctionLCRMeterinfo.intLCRMeer_RangeValue);

                // LCR-Meter 주파수 설정
                m_MeasureProcess.functionLCRMeterFrequencySetting(dcmFrequency);

                // LCR-Meter VoltageLevel 설정
                m_MeasureProcess.functionLCRMeterVoltageLevelSetting(dcmVoltageLevel);

                // LCR-Meter Mode 설정
                m_MeasureProcess.functionLCRMeterModeSetting(2);

                // 그리드의 행별 측정 진행
                for (int i = 0; i < _dgv.RowCount; i++)
                {
                    strChannelName = _dgv.Rows[i].Cells["ChannelName"].Value.ToString().Trim();

                    if (strChannelName.Trim() == "") continue;
                    if ((bool)_dgv.Rows[i].Cells["ChekBox"].Value == false) continue;

                    // Channel 별로 DAQ Pin 번호 가져오기
                    strNormalModeDAQPinMap = GetNormalModeChannelPinName(strDAQDeviceName, strDAMRelayName, strChannelName);
                    strWheatstoneModeDAQPinMap = GetWheatstoneModeChannelPinName(strDAQDeviceName, strDAMRelayName, strChannelName);

                    // DAQ 핀 On (DC일 경우만 일반과 휘스톤은 DAQ 핀이 틀림)
                    if (!_boolWheatstoneMode)
                    {
                        _dgv.Rows[i].Cells["ChannelPinNo"].Value = strNormalModeDAQPinMap;
                        boolResult = m_MeasureProcess.functionDAQPinMapOnOff(strNormalModeDAQPinMap, true);
                        Thread.Sleep(50);
                    }
                    else
                    {
                        _dgv.Rows[i].Cells["ChannelPinNo"].Value = strWheatstoneModeDAQPinMap;
                        boolResult = m_MeasureProcess.functionDAQPinMapOnOff(strWheatstoneModeDAQPinMap, true);
                        Thread.Sleep(50);
                    }

                    // 측정이 정지되었을 경우
                    if (boolMeasurementStop || !boolMeasurementStart) break;

                    // DC 측정
                    #region DC 측정

                    // 컬럼 선택
                    _dgv.Rows[i].Cells["DC_Resistance"].Selected = true;

                    // DC 데이터 측정값 가져오기
                    boolResult = m_MeasureProcess.functionRCSRdcMeasurement(ref dcmMeasurementValue, _boolWheatstoneMode, intSleep);
                    intSleep = 100;

                    // 그리드에 측정 값 삽입
                    _dgv.Rows[i].Cells["DC_Resistance"].Value = dcmMeasurementValue.ToString("F3");

                    // 컬럼 선택 취소
                    _dgv.Rows[i].Cells["DC_Resistance"].Selected = false;

                    #endregion

                    // DAQ 핀 On
                    if (!_boolWheatstoneMode)
                        boolResult = m_MeasureProcess.functionDAQPinMapOnOff(strNormalModeDAQPinMap, false);
                    else
                        boolResult = m_MeasureProcess.functionDAQPinMapOnOff(strWheatstoneModeDAQPinMap, false);

                    Thread.Sleep(50);

                    // 측정이 정지되었을 경우
                    if (boolMeasurementStop || !boolMeasurementStart) break;

                    // DAQ 핀 On (DC를 제외한 나머지는 일반과 휘스톤 DAQ 핀이 같음)
                    boolResult = m_MeasureProcess.functionDAQPinMapOnOff(strNormalModeDAQPinMap, true);
                    Thread.Sleep(50);

                    // AC 측정
                    #region AC 측정
                    // 컬럼 선택
                    _dgv.Rows[i].Cells["AC_Resistance"].Selected = true;

                    // AC 데이터 측정값 가져오기
                    boolResult = m_MeasureProcess.functionRCSRacMeasurement(ref dcmMeasurementValue);

                    // 그리드에 측정 값 삽입
                    _dgv.Rows[i].Cells["AC_Resistance"].Value = dcmMeasurementValue.ToString("F3");

                    // 컬럼 선택 취소
                    _dgv.Rows[i].Cells["AC_Resistance"].Selected = false;
                    #endregion

                    // 측정이 정지되었을 경우
                    if (boolMeasurementStop || !boolMeasurementStart) break;

                    // L 측정
                    #region L 측정
                    // 컬럼 선택
                    _dgv.Rows[i].Cells["Inductance"].Selected = true;

                    // L 데이터 측정값 가져오기
                    boolResult = m_MeasureProcess.functionRCSLMeasurement(ref dcmMeasurementValue);

                    // 그리드에 측정 값 삽입
                    _dgv.Rows[i].Cells["Inductance"].Value = dcmMeasurementValue.ToString("F3");

                    // 컬럼 선택 취소
                    _dgv.Rows[i].Cells["Inductance"].Selected = false;
                    #endregion

                    // 측정이 정지되었을 경우
                    if (boolMeasurementStop || !boolMeasurementStart) break;

                    // C 측정
                    #region C 측정
                    //// 컬럼 선택
                    //_dgv.Rows[i].Cells["Capacitance"].Selected = true;

                    //// C 데이터 측정값 가져오기
                    //boolResult = m_MeasureProcess.functionRCSCMeasurement(ref dcmMeasurementValue);

                    //// 그리드에 측정 값 삽입
                    //_dgv.Rows[i].Cells["Capacitance"].Value = dcmMeasurementValue.ToString("F6");

                    //// 컬럼 선택 취소
                    //_dgv.Rows[i].Cells["Capacitance"].Selected = false;
                    #endregion

                    // 측정이 정지되었을 경우
                    //if (boolMeasurementStop || !boolMeasurementStart) break;

                    // Q 측정
                    #region Q 측정
                    //// 컬럼 선택
                    //_dgv.Rows[i].Cells["Q_FactorValue"].Selected = true;

                    //// Q 데이터 측정값 가져오기
                    //boolResult = m_MeasureProcess.functionRCSQMeasurement(ref dcmMeasurementValue);

                    //// 그리드에 측정 값 삽입
                    //_dgv.Rows[i].Cells["Q_FactorValue"].Value = dcmMeasurementValue.ToString("F3");

                    //// 컬럼 선택 취소
                    //_dgv.Rows[i].Cells["Q_FactorValue"].Selected = false;
                    #endregion

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

                btnMeasurement.Cursor = System.Windows.Forms.Cursors.Hand;
                btnStop.Cursor = System.Windows.Forms.Cursors.Default;
                btnSave.Cursor = System.Windows.Forms.Cursors.Hand;

                if (boolMeasurementStop)
                {
                    frmMB.lblMessage.Text = "측정을 중단하였습니다.";
                    frmMB.TopMost = true;
                    frmMB.ShowDialog();
                }
                else
                {
                    btnStop.Enabled = false;
                    btnClose.Enabled = false;

                    frmWorkMessageBox frmMessage = new frmWorkMessageBox();
                    frmMessage.lblMessage.Text = "측정이 완료되었습니다.\r\n저장하시겠습니까?";
                    frmMessage.ShowDialog();

                    if (frmMessage.boolOk)
                    {
                        btnSave_Click(null, null);
                    }

                    btnStop.Enabled = true;
                    btnClose.Enabled = true;
                }

                boolMeasurementStart = false;

                //측정 시작 중단에 따라 버튼 활성화/비활성화 설정
                SetButtonEnabled(true);

                if (threadMeasurementStart != null)
                    threadMeasurementStart.Abort();
            }
        }

        /// <summary>
        /// 일반모드 Channel 별로 DAQ Pin 번호 가져오기
        /// </summary>
        /// <param name="_strCoilName"></param>
        /// <returns></returns>
        private string GetNormalModeChannelPinName(string _strDAQDeviceName, string _strDAMRelayName, string _strChannelName)
        {
            string strResult = "";

            // DAM Name
            strResult = string.Format(Gini.GetValue("DAM_PinName", $"{_strDAMRelayName.Replace(" ", "")}_RelayCardNumberDAQPinName".Trim()), _strDAQDeviceName);
            // Channel 
            strResult = string.Format(Gini.GetValue("Channel_PinName", $"NormalMode{_strChannelName}_DAQPinName"), _strDAQDeviceName, strResult);
           
            return strResult;
        }

        /// <summary>
        /// 휘스톤모드 Channel 별로 DAQ Pin 번호 가져오기
        /// </summary>
        /// <param name="_strCoilName"></param>
        /// <returns></returns>
        private string GetWheatstoneModeChannelPinName(string _strDAQDeviceName, string _strDAMRelayName, string _strChannelName)
        {
            string strResult = "";

            // DAM Name
            strResult = string.Format(Gini.GetValue("DAM_PinName", $"{_strDAMRelayName.Replace(" ", "")}_RelayCardNumberDAQPinName".Trim()), _strDAQDeviceName);
            // Channel 
            strResult = string.Format(Gini.GetValue("Channel_PinName", $"WheatstoneMode{_strChannelName}_DAQPinName"), _strDAQDeviceName, strResult);

            return strResult;
        }

        /// <summary>
        /// Tab Control Selected Index Changed Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (btnSave.ForeColor == System.Drawing.Color.Blue || boolMeasurementStart)
            {
                if (intTabSelectIndex != tabControl1.SelectedIndex)
                {
                    tabControl1.SelectedIndex = intTabSelectIndex;
                    frmMB.lblMessage.Text = "저장 중이므로 일반 및 휘스톤 모드를 변경할 수 없습니다.";
                    frmMB.TopMost = true;
                    frmMB.ShowDialog();
                }
            }
            else
            {
                intTabSelectIndex = tabControl1.SelectedIndex;
            }
        }
    }
}
