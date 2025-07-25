using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using Coil_Diagnostor.Function;

// Excel Export 용
using Excel = Microsoft.Office.Interop.Excel;

namespace Coil_Diagnostor
{
    public partial class frmRCSReport : Form
    {
        protected Function.FunctionMeasureProcess m_MeasureProcess = new Function.FunctionMeasureProcess();
        protected Function.FunctionDataControl m_db = new Function.FunctionDataControl();
        protected Function.FunctionChart m_Chart = new Function.FunctionChart();
        protected frmMessageBox frmMB = new frmMessageBox();
        protected string strPlantName = "";
        protected bool boolFormLoad = false;
		protected string strMeasurementResult = "";
        protected string strMeasurementDate = "";
        protected string strReferenceHogi = "초기값";
        protected string strReferenceOHDegree = "제 0 차";

        protected decimal dcmStopAveRdc = 0.000M;
        protected decimal dcmUpAveRdc = 0.000M;
        protected decimal dcmMoveAveRdc = 0.000M;
        protected decimal dcmStopAveRac = 0.000M;
        protected decimal dcmUpAveRac = 0.000M;
        protected decimal dcmMoveAveRac = 0.000M;
        protected decimal dcmStopAveL = 0.000M;
        protected decimal dcmUpAveL = 0.000M;
        protected decimal dcmMoveAveL = 0.000M;
        protected decimal dcmStopAveC = 0.000000M;
        protected decimal dcmUpAveC = 0.000000M;
        protected decimal dcmMoveAveC = 0.000000M;
        protected decimal dcmStopAveQ = 0.000M;
        protected decimal dcmUpAveQ = 0.000M;
        protected decimal dcmMoveAveQ = 0.000M;

        protected decimal dcmStopRefRdc = 0.000M;
        protected decimal dcmUpRefRdc = 0.000M;
        protected decimal dcmMoveRefeRdc = 0.000M;
        protected decimal dcmStopRefRac = 0.000M;
        protected decimal dcmUpRefRac = 0.000M;
        protected decimal dcmMoveRefRac = 0.000M;
        protected decimal dcmStopRefL = 0.000M;
        protected decimal dcmUpRefL = 0.000M;
        protected decimal dcmMoveRefL = 0.000M;
        protected decimal dcmStopRefC = 0.000000M;
        protected decimal dcmUpRefC = 0.000000M;
        protected decimal dcmMoveRefC = 0.000000M;
        protected decimal dcmStopRefQ = 0.000M;
        protected decimal dcmUpRefQ = 0.000M;
        protected decimal dcmMoveRefQ = 0.000M;

        private Panel pnlPowerCabint = null;

        System.Windows.Forms.DataVisualization.Charting.Series sDC = new System.Windows.Forms.DataVisualization.Charting.Series();
        System.Windows.Forms.DataVisualization.Charting.Series sAC = new System.Windows.Forms.DataVisualization.Charting.Series();
        System.Windows.Forms.DataVisualization.Charting.Series sL = new System.Windows.Forms.DataVisualization.Charting.Series();
        System.Windows.Forms.DataVisualization.Charting.Series sC = new System.Windows.Forms.DataVisualization.Charting.Series();
        System.Windows.Forms.DataVisualization.Charting.Series sQFactor = new System.Windows.Forms.DataVisualization.Charting.Series();

        public frmRCSReport()
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
                case Keys.F5: // 설정 버튼
                    btnCheckBoxAllSelect52.PerformClick();
                    break;
                case Keys.F6: // 설정 버튼
                    btnCheckBoxAllClear52.PerformClick();
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
        private void frmRCSReport_Load(object sender, EventArgs e)
        {
            strPlantName = Gini.GetValue("Device", "PlantName").Trim();

            // ComboBox 데이터 설정
            SetComboBoxDataBinding();

            // 그리드 초기 설정
            SetDataGridViewInitialize();

            string plantName = Gini.GetValue("Device", "PlantName");
            SetRodPanel(plantName);
            SetCheckBoxDataBinding(pnlPowerCabint);

            chkMeasurementTarget_CheckedChanged(chkC, null);
        }

        // 20240401 한인석
        private void SetRodPanel(string plantName)
        { 
            switch (plantName)
            {
                case "고리 1발전소":
                    pnlPowerCabint = pnlRod33;
                    break;
                case "한빛 1발전소":
                    pnlPowerCabint = pnlRod52;
                    break;
                default:
                    pnlPowerCabint = pnlRod52;
                    break;
            }
            pnlPowerCabint.Visible = true;

            btnCheckBoxAllClear52.PerformClick();
            btnCheckBoxAllClear33.PerformClick();
        }
          
        // 20240329 한인석
        private void SetCheckBoxDataBinding(Panel p)
        {
            if (p == null) return;
            string[] powerCabinets = Gini.GetValue("RCS", "RCSMeasurementGroup_Item").Split(',');
            foreach(Control grv in p.Controls) 
            {
                //if(grv.GetType().Name.Trim() == "GroupBox")
                if(grv is GroupBox)
                {
                    string powerCabinetName = $"chk{grv.Name.Substring(2)}";
                    var cb = grv.Controls.Find(powerCabinetName, true).FirstOrDefault();
                    if(cb != null)
                    {
                        cb.Text = powerCabinets[powerCabinetName[powerCabinetName.Length - 1] - '0'];
                        string key = $"RCSPowerCabinetItem_{cb.Text}";
                        string[] rodNames = Gini.GetValue("RCS", key).Split(',');

                        foreach (Control rod in grv.Controls)
                        {
                            CheckBox r = rod as CheckBox;
                            if (r.Name != cb.Name)
                            {
                                int i = Convert.ToInt32(r.Name.Substring(cb.Name.Length + 1));
                                if(i >= rodNames.Length || rodNames[i] == "N/A")
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
            dt = m_db.GetRCSReferenceValueOhDegreeData(strPlantName.Trim());

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
            if (Gini.GetValue("RCS", "SelectRCS_Hogi").Trim() == "")
            {
                // 최초 실행 시 기본 호기 선택
                cboHogi.SelectedIndex = 0;
            }
            else
            {
                // 직전 실행 시 기본 호기 선택
                cboHogi.SelectedItem = Gini.GetValue("RCS", "SelectRCS_Hogi").Trim();
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
            string[] strTypeItem = Gini.GetValue("Combo", "RCSType_Item").Split(',');

            cboMeasurementType.Items.Clear();

            for (int i = 0; i < strTypeItem.Length; i++)
            {
                cboMeasurementType.Items.Add(strTypeItem[i].Trim());
            }

            cboMeasurementType.SelectedIndex = 0;


            // 코일명
            string[] strCoilNameItem = Gini.GetValue("Combo", "RCSCoilName_Item").Split(',');

            cboCoilName.Items.Clear();

            cboCoilName.Items.Add("전체");

            for (int i = 0; i < strCoilNameItem.Length; i++)
            {
                cboCoilName.Items.Add(strCoilNameItem[i].Trim());
            }

            cboCoilName.SelectedIndex = 0;


            // 비교대상s
            string[] strComparisonTargetItem = Gini.GetValue("Combo", "ComparisonTarget_Item").Split(',');

            cboComparisonTarget.Items.Clear();

            for (int i = 0; i < strComparisonTargetItem.Length; i++)
            {
                cboComparisonTarget.Items.Add(strComparisonTargetItem[i].Trim());
            }

            cboComparisonTarget.SelectedIndex = 0;
        }

        /// <summary>
        /// 그리드 초기 설정
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

                Column2.HeaderText = "코일명";
                Column2.Name = "CoilName";
                Column2.Width = 80;
                Column2.ReadOnly = true;
                Column2.Visible = true;
                Column2.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column2.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column3.HeaderText = "결과";
                Column3.Name = "Result";
                Column3.Width = 80;
                Column3.ReadOnly = true;
                Column3.Visible = true;
                Column3.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column3.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column4.HeaderText = "Index";
                Column4.Name = "CoilNumber";
                Column4.Width = 80;
                Column4.ReadOnly = true;
                Column4.Visible = false;
                Column4.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column4.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column5.HeaderText = "DC 저항(Ω)";
                Column5.Name = "DC_ResistanceValue";
                Column5.Width = 100;
                Column5.ReadOnly = true;
                Column5.Visible = chkRdc.Checked;
                Column5.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column5.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column6.HeaderText = "편차(Ω)";
                Column6.Name = "DC_Deviation";
                Column6.Width = 80;
                Column6.ReadOnly = true;
                Column6.Visible = chkRdc.Checked;
                Column6.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column6.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column7.HeaderText = "AC 저항(Ω)";
                Column7.Name = "AC_ResistanceValue";
                Column7.Width = 100;
                Column7.ReadOnly = true;
                Column7.Visible = chkRac.Checked;
                Column7.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column7.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column8.HeaderText = "편차(Ω)";
                Column8.Name = "AC_Deviation";
                Column8.Width = 80;
                Column8.ReadOnly = true;
                Column8.Visible = chkRac.Checked;
                Column8.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column8.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column9.HeaderText = "인덕턴스(mH)";
                Column9.Name = "L_InductanceValue";
                Column9.Width = 100;
                Column9.ReadOnly = true;
                Column9.Visible = chkL.Checked;
                Column9.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column9.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column10.HeaderText = "편차(mH)";
                Column10.Name = "L_Deviation";
                Column10.Width = 80;
                Column10.ReadOnly = true;
                Column10.Visible = chkL.Checked;
                Column10.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column10.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column11.HeaderText = "캐패시턴스(uF)";
                Column11.Name = "C_CapacitanceValue";
                Column11.Width = 100;
                Column11.ReadOnly = true;
                Column11.Visible = chkC.Checked;
                Column11.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column11.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column12.HeaderText = "편차(uF)";
                Column12.Name = "C_Deviation";
                Column12.Width = 80;
                Column12.ReadOnly = true;
                Column12.Visible = chkC.Checked;
                Column12.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column12.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column13.HeaderText = "Q Factor";
                Column13.Name = "Q_FactorValue";
                Column13.Width = 100;
                Column13.ReadOnly = true;
                Column13.Visible = chkQ.Checked;
                Column13.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column13.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column14.HeaderText = "편차";
                Column14.Name = "Q_Deviation";
                Column14.Width = 80;
                Column14.ReadOnly = true;
                Column14.Visible = chkQ.Checked;
                Column14.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column14.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column15.HeaderText = "발전소";
                Column15.Name = "PlantName";
                Column15.Width = 80;
                Column15.ReadOnly = true;
                Column15.Visible = false;
                Column15.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column15.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column16.HeaderText = "호기";
                Column16.Name = "Hogi";
                Column16.Width = 80;
                Column16.ReadOnly = true;
                Column16.Visible = false;
                Column16.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column16.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column17.HeaderText = "OH 차수";
                Column17.Name = "Oh_Degree";
                Column17.Width = 80;
                Column17.ReadOnly = true;
                Column17.Visible = false;
                Column17.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column17.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column18.HeaderText = "전력함";
                Column18.Name = "PowerCabinet";
                Column18.Width = 80;
                Column18.ReadOnly = true;
                Column18.Visible = false;
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

        private void chkPowerCabinet_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cb = sender as CheckBox;
            string grBoxName = $"gb{cb.Name.Substring(3)}";
            var gb = this.Controls.Find(grBoxName, true).FirstOrDefault(); 

            foreach (Control rod in gb.Controls)
            {
                if(rod is CheckBox && rod.Name != cb.Name && rod.Visible)
                {
                    ((CheckBox)rod).Checked = cb.Checked;
                }
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
                string strSelectHogi = cboHogi.SelectedItem == null ? "" : cboHogi.SelectedItem.ToString().Trim();

                // 호기별 OH 차수 설정
                DataTable dt = new DataTable();
                dt = m_db.GetRCSDiagnosisHeaderOhDegreeInfo(strPlantName, strSelectHogi);

                cboOhDegree.Items.Clear();

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (dt.Rows[i]["Oh_Degree"] != null)
                    {
                        cboOhDegree.Items.Add(dt.Rows[i]["Oh_Degree"].ToString().Trim());
                    }
                }

                if (Gini.GetValue("RCS", "SelectRCS_Hogi").Trim() == "")
                    cboOhDegree.SelectedIndex = dt.Rows.Count - 1;
                else
                    cboOhDegree.SelectedItem = "제 " + Gini.GetValue("RCS", "SelectRCS_OHDegree").Trim() + " 차";

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
        /// 측정타입 Selected Index Changed Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cboMeasurementType_SelectedIndexChanged(object sender, EventArgs e)
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
            if (gbPowerCabinet52_0.Enabled)
                chkPowerCabinet52_0.Checked = true;

            if (gbPowerCabinet52_1.Enabled)
                chkPowerCabinet52_1.Checked = true;

            if (gbPowerCabinet52_2.Enabled)
                chkPowerCabinet52_2.Checked = true;

            if (gbPowerCabinet52_3.Enabled)
                chkPowerCabinet52_3.Checked = true;

            if (gbPowerCabinet52_4.Enabled)
                chkPowerCabinet52_4.Checked = true;
        }

        /// <summary>
        /// 전체해제 Button Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCheckBoxAllClear_Click(object sender, EventArgs e)
        {
            if (gbPowerCabinet52_0.Enabled)
                chkPowerCabinet52_0.Checked = false;

            if (gbPowerCabinet52_1.Enabled)
                chkPowerCabinet52_1.Checked = false;

            if (gbPowerCabinet52_2.Enabled)
                chkPowerCabinet52_2.Checked = false;

            if (gbPowerCabinet52_3.Enabled)
                chkPowerCabinet52_3.Checked = false;

            if (gbPowerCabinet52_4.Enabled)
                chkPowerCabinet52_4.Checked = false;
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
            string strPath = Application.StartupPath + @"\Report";
            System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(strPath);

            if (!di.Exists)
            {
                di.Create();
            }

            System.Diagnostics.Process.Start(strPath);
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

                // 그리드와 그래프 초기화
                SetDataGridViewAndGraphInitialize();

                string strHogi = cboHogi.SelectedItem == null ? "" : cboHogi.SelectedItem.ToString().Trim();
                string strOhDegree = cboOhDegree.SelectedItem == null ? "" : cboOhDegree.SelectedItem.ToString().Trim();
                string strMeasurementType = cboMeasurementType.SelectedItem == null ? "" : cboMeasurementType.SelectedItem.ToString().Trim();
                string strFrequency = cboFrequency.SelectedItem == null ? "" : cboFrequency.SelectedItem.ToString().Trim();
                string strCoilName = cboCoilName.SelectedItem == null ? "" : cboCoilName.SelectedItem.ToString().Trim();
                string strComparisonTarget = cboComparisonTarget.SelectedItem == null ? "" : cboComparisonTarget.SelectedItem.ToString().Trim();
                string str1AC = "", str2AC = "", str1BD = "", str2BD = "", strSCD = "", strControlName = "";

                strControlName = GetCheckBoxCheckedTrueItem(pnlPowerCabint);

                if (strControlName.Trim() == "")
                {
                    frmMB.lblMessage.Text = "보고서 출력할 전력합을 선택하십시오.";
                    frmMB.TopMost = true;
                    frmMB.ShowDialog();
                    return;
                }

                DataTable dtDB = new DataTable();
                DataTable dtAverage = new DataTable();

                // 기준값 가져오기
                GetRCSReferenceValue();

                if (cboMeasurementType.SelectedItem != null && cboMeasurementType.SelectedItem.ToString().Trim() == "평균치")
                {
                    // 데이터 평균
                    dtDB = m_db.GetRCSReportDataMeasure(strPlantName.Trim(), strHogi, strOhDegree, strFrequency, strMeasurementType, strCoilName
                        , strControlName, strReferenceHogi, strReferenceOHDegree);

                    dtAverage = m_db.GetRCSReportDataAverage(strPlantName.Trim(), strHogi, strOhDegree, strFrequency, strMeasurementType, strCoilName
                        , strControlName, strReferenceHogi, strReferenceOHDegree);
                }
                else
                {
                    dtDB = m_db.GetRCSReportDataMeasure(strPlantName.Trim(), strHogi, strOhDegree, strFrequency, strMeasurementType, strCoilName
                        , strControlName, strReferenceHogi, strReferenceOHDegree);

                    dtAverage = dtDB;
                }

                if (dtDB == null || dtDB.Rows.Count <= 0)
                {
                    frmMB.lblMessage.Text = "데이터가 없습니다.";
                    frmMB.TopMost = true;
                    frmMB.ShowDialog();
                }
                else
                {
                    // 평균치 산출
                    SetDataAverage(dtAverage);

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

                    if (dtDB != null && dtDB.Rows.Count > 0)
                    {
                        string strXLabelValue = "", strXLabelCheck = "";

                        for (int i = 0; i < dtDB.Rows.Count; i++)
                        {
                            if (strXLabelCheck.Trim() != dtDB.Rows[i]["ControlName"].ToString().Trim())
                            {
                                strXLabelCheck = dtDB.Rows[i]["ControlName"].ToString().Trim();

                                if (strXLabelValue.Trim() == "")
                                    strXLabelValue = dtDB.Rows[i]["ControlName"].ToString().Trim();
                                else
                                    strXLabelValue = strXLabelValue + "," + dtDB.Rows[i]["ControlName"].ToString().Trim();
                            }
                        }

                        string[] arrayXLabelValue = strXLabelValue.Split(',');
                        int intXLabelValueCount = arrayXLabelValue.Length;

                        // 표 형식 
                        SetRCSReferenceValueTableMark(ref dtDB, strComparisonTarget);

                        // 측정값 및 편차 그래프 그리기
                        SetRCSReferenceValueMeasurement(dtAverage, arrayXLabelValue, intXLabelValueCount);

						if (dgvMeasurement.RowCount > 0)
						{
							// 엑셀 보고서 보내기
							if (ExcelReportExport(strHogi, strOhDegree, strMeasurementType, strCoilName, strComparisonTarget))
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
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
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
        /// 기준 값(RCS) 가져오기
        /// </summary>
        private void GetRCSReferenceValue()
        {
            DataTable dt;
            string strSelectHogi = "";
            string strSelectOhDegree = "";

            try
            {
                // RCS 기준값 Data 가져오기
                strSelectHogi = cboHogi.SelectedItem == null ? "초기값" : cboHogi.SelectedItem.ToString().Trim();
                strSelectOhDegree = cboOhDegree.SelectedItem == null ? "제 0 차" : cboOhDegree.SelectedItem.ToString().Trim();

                dt = new DataTable();

                // 해당 차수의 기준 값 유무 체크
                if ((m_db.GetRCSReferenceValueDataCount(strPlantName.Trim(), strSelectHogi.Trim(), strSelectOhDegree.Trim())) > 0)
                {
                    // 해당 차수의 기준 값을 가져온다.
                    strReferenceHogi = strSelectHogi.Trim();
                    strReferenceOHDegree = strSelectOhDegree.Trim();
                    dt = m_db.GetRCSReferenceValueDataInfo(strPlantName.Trim(), strReferenceHogi.Trim(), strReferenceOHDegree.Trim());
                }
                else
                {
                    // MAX OH 차수를 가져온다.
                    int intMaxOhDegree = m_db.GetRCSReferenceValueDataMaxOhDegree(strPlantName.Trim(), strSelectHogi.Trim());

                    if (intMaxOhDegree > 0)
                    {
                        // MAX 차수의 기준 값을 가져온다.
                        strReferenceHogi = strSelectHogi.Trim();
                        strReferenceOHDegree = "제 " + intMaxOhDegree.ToString().Trim() + " 차";
                        dt = m_db.GetRCSReferenceValueDataInfo(strPlantName.Trim(), strReferenceHogi.Trim(), strReferenceOHDegree.Trim());
                    }
                    else
                    {
                        // 초기값 0 차수의 기준 값을 가져온다.
                        strReferenceHogi = "초기값";
                        strReferenceOHDegree = "제 0 차";
                        dt = m_db.GetRCSReferenceValueDataInfo(strPlantName.Trim(), strReferenceHogi.Trim(), strReferenceOHDegree.Trim());
                    }
                }

                if (dt != null && dt.Rows.Count > 0)
                {
                    dcmStopRefRdc = dt.Rows[0]["RdcStop_RCSReferenceValue"] == null || dt.Rows[0]["RdcStop_RCSReferenceValue"].ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(dt.Rows[0]["RdcStop_RCSReferenceValue"].ToString().Trim());
                    dcmUpRefRdc = dt.Rows[0]["RdcUp_RCSReferenceValue"] == null || dt.Rows[0]["RdcUp_RCSReferenceValue"].ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(dt.Rows[0]["RdcUp_RCSReferenceValue"].ToString().Trim());
                    dcmMoveRefeRdc = dt.Rows[0]["RdcMove_RCSReferenceValue"] == null || dt.Rows[0]["RdcMove_RCSReferenceValue"].ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(dt.Rows[0]["RdcMove_RCSReferenceValue"].ToString().Trim());
                    dcmStopRefRac = dt.Rows[0]["RacStop_RCSReferenceValue"] == null || dt.Rows[0]["RacStop_RCSReferenceValue"].ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(dt.Rows[0]["RacStop_RCSReferenceValue"].ToString().Trim());
                    dcmUpRefRac = dt.Rows[0]["RacUp_RCSReferenceValue"] == null || dt.Rows[0]["RacUp_RCSReferenceValue"].ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(dt.Rows[0]["RacUp_RCSReferenceValue"].ToString().Trim());
                    dcmMoveRefRac = dt.Rows[0]["RacMove_RCSReferenceValue"] == null || dt.Rows[0]["RacMove_RCSReferenceValue"].ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(dt.Rows[0]["RacMove_RCSReferenceValue"].ToString().Trim());
                    dcmStopRefL = dt.Rows[0]["LStop_RCSReferenceValue"] == null || dt.Rows[0]["LStop_RCSReferenceValue"].ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(dt.Rows[0]["LStop_RCSReferenceValue"].ToString().Trim());
                    dcmUpRefL = dt.Rows[0]["LUp_RCSReferenceValue"] == null || dt.Rows[0]["LUp_RCSReferenceValue"].ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(dt.Rows[0]["LUp_RCSReferenceValue"].ToString().Trim());
                    dcmMoveRefL = dt.Rows[0]["LMove_RCSReferenceValue"] == null || dt.Rows[0]["LMove_RCSReferenceValue"].ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(dt.Rows[0]["LMove_RCSReferenceValue"].ToString().Trim());
                    dcmStopRefC = dt.Rows[0]["CStop_RCSReferenceValue"] == null || dt.Rows[0]["CStop_RCSReferenceValue"].ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(dt.Rows[0]["CStop_RCSReferenceValue"].ToString().Trim());
                    dcmUpRefC = dt.Rows[0]["CUp_RCSReferenceValue"] == null || dt.Rows[0]["CUp_RCSReferenceValue"].ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(dt.Rows[0]["CUp_RCSReferenceValue"].ToString().Trim());
                    dcmMoveRefC = dt.Rows[0]["CMove_RCSReferenceValue"] == null || dt.Rows[0]["CMove_RCSReferenceValue"].ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(dt.Rows[0]["CMove_RCSReferenceValue"].ToString().Trim());
                    dcmStopRefQ = dt.Rows[0]["QStop_RCSReferenceValue"] == null || dt.Rows[0]["QStop_RCSReferenceValue"].ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(dt.Rows[0]["QStop_RCSReferenceValue"].ToString().Trim());
                    dcmUpRefQ = dt.Rows[0]["QUp_RCSReferenceValue"] == null || dt.Rows[0]["QUp_RCSReferenceValue"].ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(dt.Rows[0]["QUp_RCSReferenceValue"].ToString().Trim());
                    dcmMoveRefQ = dt.Rows[0]["QMove_RCSReferenceValue"] == null || dt.Rows[0]["QMove_RCSReferenceValue"].ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(dt.Rows[0]["QMove_RCSReferenceValue"].ToString().Trim());
                }
                else
                {                    
                    dcmStopRefRdc = 0.000M;
                    dcmUpRefRdc = 0.000M;
                    dcmMoveRefeRdc = 0.000M;
                    dcmStopRefRac = 0.000M;
                    dcmUpRefRac = 0.000M;
                    dcmMoveRefRac = 0.000M;
                    dcmStopRefL = 0.000M;
                    dcmUpRefL = 0.000M;
                    dcmMoveRefL = 0.000M;
                    dcmStopRefC = 0.000000M;
                    dcmUpRefC = 0.000000M;
                    dcmMoveRefC = 0.000000M;
                    dcmStopRefQ = 0.000M;
                    dcmUpRefQ = 0.000M;
                    dcmMoveRefQ = 0.000M;
                }
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
                if (grv is GroupBox)
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

        /// <summary>
        /// 평균 구하기
        /// </summary>
        /// <param name="_dt"></param>
        /// <returns></returns>
        private void SetDataAverage(DataTable _dt)
        {
            decimal dcmDecisionRange_Rdc = 0M, dcmDecisionRange_Rac = 0M, dcmDecisionRange_L = 0M, dcmDecisionRange_C = 0M
                , dcmDecisionRange_Q = 0M, dcmEffectiveStandardRange = 0M;
            dcmDecisionRange_Rdc = Convert.ToDecimal(Gini.GetValue("ReferenceValue", "RdcDecisionRange_ReferenceValue"));
            dcmDecisionRange_Rac = Convert.ToDecimal(Gini.GetValue("ReferenceValue", "RacDecisionRange_ReferenceValue"));
            dcmDecisionRange_L = Convert.ToDecimal(Gini.GetValue("ReferenceValue", "LDecisionRange_ReferenceValue"));
            dcmDecisionRange_C = Convert.ToDecimal(Gini.GetValue("ReferenceValue", "CDecisionRange_ReferenceValue"));
            dcmDecisionRange_Q = Convert.ToDecimal(Gini.GetValue("ReferenceValue", "QDecisionRange_ReferenceValue"));
            dcmEffectiveStandardRange = Convert.ToDecimal(Gini.GetValue("ReferenceValue", "EffectiveStandardRangeOfVariation"));

            decimal dcmSumRdcStop = 0.000M, dcmSumRacStop = 0.000M, dcmSumLStop = 0.000M, dcmSumCStop = 0.000M, dcmSumQStop = 0.000M;
            decimal dcmSumRdcUp = 0.000M, dcmSumRacUp = 0.000M, dcmSumLUp = 0.000M, dcmSumCUp = 0.000M, dcmSumQUp = 0.000M;
            decimal dcmSumRdcMove = 0.000M, dcmSumRacMove = 0.000M, dcmSumLMove = 0.000M, dcmSumCMove = 0.000M, dcmSumQMove = 0.000M;
            int intSumStopCount = 0, intSumUpCount = 0, intSumMoveCount = 0;

            for (int i = 0; i < _dt.Rows.Count; i++)
            {
                // 평균 구하기위해 합계 산출
                if (_dt.Rows[i]["CoilName"].ToString().Trim() == "정지")
                {
                    dcmSumRdcStop += _dt.Rows[i]["DC_ResistanceValue"] == null || _dt.Rows[i]["DC_ResistanceValue"].ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(_dt.Rows[i]["DC_ResistanceValue"].ToString().Trim());
                    dcmSumRacStop += _dt.Rows[i]["AC_ResistanceValue"] == null || _dt.Rows[i]["AC_ResistanceValue"].ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(_dt.Rows[i]["AC_ResistanceValue"].ToString().Trim());
                    dcmSumLStop += _dt.Rows[i]["L_InductanceValue"] == null || _dt.Rows[i]["L_InductanceValue"].ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(_dt.Rows[i]["L_InductanceValue"].ToString().Trim());
                    dcmSumCStop += _dt.Rows[i]["C_CapacitanceValue"] == null || _dt.Rows[i]["C_CapacitanceValue"].ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(_dt.Rows[i]["C_CapacitanceValue"].ToString().Trim());
                    dcmSumQStop += _dt.Rows[i]["Q_FactorValue"] == null || _dt.Rows[i]["Q_FactorValue"].ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(_dt.Rows[i]["Q_FactorValue"].ToString().Trim());
                    intSumStopCount++;
                }
                else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "올림")
                {
                    dcmSumRdcUp += _dt.Rows[i]["DC_ResistanceValue"] == null || _dt.Rows[i]["DC_ResistanceValue"].ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(_dt.Rows[i]["DC_ResistanceValue"].ToString().Trim());
                    dcmSumRacUp += _dt.Rows[i]["AC_ResistanceValue"] == null || _dt.Rows[i]["AC_ResistanceValue"].ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(_dt.Rows[i]["AC_ResistanceValue"].ToString().Trim());
                    dcmSumLUp += _dt.Rows[i]["L_InductanceValue"] == null || _dt.Rows[i]["L_InductanceValue"].ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(_dt.Rows[i]["L_InductanceValue"].ToString().Trim());
                    dcmSumCUp += _dt.Rows[i]["C_CapacitanceValue"] == null || _dt.Rows[i]["C_CapacitanceValue"].ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(_dt.Rows[i]["C_CapacitanceValue"].ToString().Trim());
                    dcmSumQUp += _dt.Rows[i]["Q_FactorValue"] == null || _dt.Rows[i]["Q_FactorValue"].ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(_dt.Rows[i]["Q_FactorValue"].ToString().Trim());
                    intSumUpCount++;
                }
                else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "이동")
                {
                    dcmSumRdcMove += _dt.Rows[i]["DC_ResistanceValue"] == null || _dt.Rows[i]["DC_ResistanceValue"].ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(_dt.Rows[i]["DC_ResistanceValue"].ToString().Trim());
                    dcmSumRacMove += _dt.Rows[i]["AC_ResistanceValue"] == null || _dt.Rows[i]["AC_ResistanceValue"].ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(_dt.Rows[i]["AC_ResistanceValue"].ToString().Trim());
                    dcmSumLMove += _dt.Rows[i]["L_InductanceValue"] == null || _dt.Rows[i]["L_InductanceValue"].ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(_dt.Rows[i]["L_InductanceValue"].ToString().Trim());
                    dcmSumCMove += _dt.Rows[i]["C_CapacitanceValue"] == null || _dt.Rows[i]["C_CapacitanceValue"].ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(_dt.Rows[i]["C_CapacitanceValue"].ToString().Trim());
                    dcmSumQMove += _dt.Rows[i]["Q_FactorValue"] == null || _dt.Rows[i]["Q_FactorValue"].ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(_dt.Rows[i]["Q_FactorValue"].ToString().Trim());
                    intSumMoveCount++;
                }
            }

            decimal dcmAvgRdcStop = intSumStopCount > 0 ? dcmSumRdcStop / intSumStopCount : 0M;
            decimal dcmAvgRacStop = intSumStopCount > 0 ? dcmSumRacStop / intSumStopCount : 0M;
            decimal dcmAvgLStop = intSumStopCount > 0 ? dcmSumLStop / intSumStopCount : 0M;
            decimal dcmAvgCStop = intSumStopCount > 0 ? dcmSumCStop / intSumStopCount : 0M;
            decimal dcmAvgQStop = intSumStopCount > 0 ? dcmSumQStop / intSumStopCount : 0M;
            decimal dcmAvgRdcUp = intSumUpCount > 0 ? dcmSumRdcUp / intSumUpCount : 0M;
            decimal dcmAvgRacUp = intSumUpCount > 0 ? dcmSumRacUp / intSumUpCount : 0M;
            decimal dcmAvgLUp = intSumUpCount > 0 ? dcmSumLUp / intSumUpCount : 0M;
            decimal dcmAvgCUp = intSumUpCount > 0 ? dcmSumCUp / intSumUpCount : 0M;
            decimal dcmAvgQUp = intSumUpCount > 0 ? dcmSumQUp / intSumUpCount : 0M;
            decimal dcmAvgRdcMove = intSumMoveCount > 0 ? dcmSumRdcMove / intSumMoveCount : 0M;
            decimal dcmAvgRacMove = intSumMoveCount > 0 ? dcmSumRacMove / intSumMoveCount : 0M;
            decimal dcmAvgLMove = intSumMoveCount > 0 ? dcmSumLMove / intSumMoveCount : 0M;
            decimal dcmAvgCMove = intSumMoveCount > 0 ? dcmSumCMove / intSumMoveCount : 0M;
            decimal dcmAvgQMove = intSumMoveCount > 0 ? dcmSumQMove / intSumMoveCount : 0M;

            dcmSumRdcStop = 0.000M; dcmSumRacStop = 0.000M; dcmSumLStop = 0.000M; dcmSumCStop = 0.000M; dcmSumQStop = 0.000M;
            dcmSumRdcUp = 0.000M; dcmSumRacUp = 0.000M; dcmSumLUp = 0.000M; dcmSumCUp = 0.000M; dcmSumQUp = 0.000M;
            dcmSumRdcMove = 0.000M; dcmSumRacMove = 0.000M; dcmSumLMove = 0.000M; dcmSumCMove = 0.000M; dcmSumQMove = 0.000M;
            decimal dcmReferenceValue = Convert.ToDecimal(Gini.GetValue("ReferenceValue", "RdcDecisionRange_ReferenceValue"));
            decimal dcmMeasureValue = 0M, dcmChkValue = 0M;

            int intAvgRdcStopCount = 0, intAvgRdcUpCount = 0, intAvgRdcMoveCount = 0, intAvgRacStopCount = 0, intAvgRacUpCount = 0, intAvgRacMoveCount = 0
                , intAvgLStopCount = 0, intAvgLUpCount = 0, intAvgLMoveCount = 0, intAvgCStopCount = 0, intAvgCUpCount = 0, intAvgCMoveCount = 0
                , intAvgQStopCount = 0, intAvgQUpCount = 0, intAvgQMoveCount = 0;

            for (int i = 0; i < _dt.Rows.Count; i++)
            {
                // 평균 구하기위해 합계 산출
                if (_dt.Rows[i]["CoilName"].ToString().Trim() == "정지")
                {
                    dcmMeasureValue = _dt.Rows[i]["DC_ResistanceValue"] == null || _dt.Rows[i]["DC_ResistanceValue"].ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(_dt.Rows[i]["DC_ResistanceValue"].ToString().Trim());

                    dcmMeasureValue = dcmMeasureValue - dcmAvgRdcStop;

                    if (dcmMeasureValue < 0) dcmMeasureValue = -dcmMeasureValue;

                    dcmChkValue = dcmMeasureValue * (dcmReferenceValue / 100);

                    if (dcmMeasureValue > dcmChkValue)
                    {
                        dcmSumRdcStop += _dt.Rows[i]["DC_ResistanceValue"] == null || _dt.Rows[i]["DC_ResistanceValue"].ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(_dt.Rows[i]["DC_ResistanceValue"].ToString().Trim());
                        intAvgRdcStopCount++;
                    }

                    dcmMeasureValue = _dt.Rows[i]["AC_ResistanceValue"] == null || _dt.Rows[i]["AC_ResistanceValue"].ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(_dt.Rows[i]["AC_ResistanceValue"].ToString().Trim());

                    dcmMeasureValue = dcmMeasureValue - dcmAvgRacStop;

                    if (dcmMeasureValue < 0) dcmMeasureValue = -dcmMeasureValue;

                    dcmChkValue = dcmMeasureValue * (dcmReferenceValue / 100);

                    if (dcmMeasureValue > dcmChkValue)
                    {
                        dcmSumRacStop += _dt.Rows[i]["AC_ResistanceValue"] == null || _dt.Rows[i]["AC_ResistanceValue"].ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(_dt.Rows[i]["AC_ResistanceValue"].ToString().Trim());
                        intAvgRacStopCount++;
                    }

                    dcmMeasureValue = _dt.Rows[i]["L_InductanceValue"] == null || _dt.Rows[i]["L_InductanceValue"].ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(_dt.Rows[i]["L_InductanceValue"].ToString().Trim());

                    dcmMeasureValue = dcmMeasureValue - dcmAvgLStop;

                    if (dcmMeasureValue < 0) dcmMeasureValue = -dcmMeasureValue;

                    dcmChkValue = dcmMeasureValue * (dcmReferenceValue / 100);

                    if (dcmMeasureValue > dcmChkValue)
                    {
                        dcmSumLStop += _dt.Rows[i]["L_InductanceValue"] == null || _dt.Rows[i]["L_InductanceValue"].ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(_dt.Rows[i]["L_InductanceValue"].ToString().Trim());
                        intAvgLStopCount++;
                    }

                    dcmMeasureValue = _dt.Rows[i]["C_CapacitanceValue"] == null || _dt.Rows[i]["C_CapacitanceValue"].ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(_dt.Rows[i]["C_CapacitanceValue"].ToString().Trim());

                    dcmMeasureValue = dcmMeasureValue - dcmAvgCStop;

                    if (dcmMeasureValue < 0) dcmMeasureValue = -dcmMeasureValue;

                    dcmChkValue = dcmMeasureValue * (dcmReferenceValue / 100);

                    if (dcmMeasureValue > dcmChkValue)
                    {
                        dcmSumCStop += _dt.Rows[i]["C_CapacitanceValue"] == null || _dt.Rows[i]["C_CapacitanceValue"].ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(_dt.Rows[i]["C_CapacitanceValue"].ToString().Trim());
                        intAvgCStopCount++;
                    }

                    dcmMeasureValue = _dt.Rows[i]["Q_FactorValue"] == null || _dt.Rows[i]["Q_FactorValue"].ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(_dt.Rows[i]["Q_FactorValue"].ToString().Trim());

                    dcmMeasureValue = dcmMeasureValue - dcmAvgQStop;

                    if (dcmMeasureValue < 0) dcmMeasureValue = -dcmMeasureValue;

                    dcmChkValue = dcmMeasureValue * (dcmReferenceValue / 100);

                    if (dcmMeasureValue > dcmChkValue)
                    {
                        dcmSumQStop += _dt.Rows[i]["Q_FactorValue"] == null || _dt.Rows[i]["Q_FactorValue"].ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(_dt.Rows[i]["Q_FactorValue"].ToString().Trim());
                        intAvgQStopCount++;
                    }
                }
                else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "올림")
                {
                    dcmMeasureValue = _dt.Rows[i]["DC_ResistanceValue"] == null || _dt.Rows[i]["DC_ResistanceValue"].ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(_dt.Rows[i]["DC_ResistanceValue"].ToString().Trim());

                    dcmMeasureValue = dcmMeasureValue - dcmAvgRdcUp;

                    if (dcmMeasureValue < 0) dcmMeasureValue = -dcmMeasureValue;

                    dcmChkValue = dcmMeasureValue * (dcmReferenceValue / 100);

                    if (dcmMeasureValue > dcmChkValue)
                    {
                        dcmSumRdcUp += _dt.Rows[i]["DC_Deviation"] == null || _dt.Rows[i]["DC_Deviation"].ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(_dt.Rows[i]["DC_Deviation"].ToString().Trim());
                        intAvgRdcUpCount++;
                    }

                    dcmMeasureValue = _dt.Rows[i]["AC_ResistanceValue"] == null || _dt.Rows[i]["AC_ResistanceValue"].ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(_dt.Rows[i]["AC_ResistanceValue"].ToString().Trim());

                    dcmMeasureValue = dcmMeasureValue - dcmAvgRacUp;

                    if (dcmMeasureValue < 0) dcmMeasureValue = -dcmMeasureValue;

                    dcmChkValue = dcmMeasureValue * (dcmReferenceValue / 100);

                    if (dcmMeasureValue > dcmChkValue)
                    {
                        dcmSumRacUp += _dt.Rows[i]["AC_ResistanceValue"] == null || _dt.Rows[i]["AC_ResistanceValue"].ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(_dt.Rows[i]["AC_ResistanceValue"].ToString().Trim());
                        intAvgRacUpCount++;
                    }

                    dcmMeasureValue = _dt.Rows[i]["L_InductanceValue"] == null || _dt.Rows[i]["L_InductanceValue"].ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(_dt.Rows[i]["L_InductanceValue"].ToString().Trim());

                    dcmMeasureValue = dcmMeasureValue - dcmAvgLUp;

                    if (dcmMeasureValue < 0) dcmMeasureValue = -dcmMeasureValue;

                    dcmChkValue = dcmMeasureValue * (dcmReferenceValue / 100);

                    if (dcmMeasureValue > dcmChkValue)
                    {
                        dcmSumLUp += _dt.Rows[i]["L_InductanceValue"] == null || _dt.Rows[i]["L_InductanceValue"].ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(_dt.Rows[i]["L_InductanceValue"].ToString().Trim());
                        intAvgLUpCount++;
                    }

                    dcmMeasureValue = _dt.Rows[i]["C_CapacitanceValue"] == null || _dt.Rows[i]["C_CapacitanceValue"].ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(_dt.Rows[i]["C_CapacitanceValue"].ToString().Trim());

                    dcmMeasureValue = dcmMeasureValue - dcmAvgCUp;

                    if (dcmMeasureValue < 0) dcmMeasureValue = -dcmMeasureValue;

                    dcmChkValue = dcmMeasureValue * (dcmReferenceValue / 100);

                    if (dcmMeasureValue > dcmChkValue)
                    {
                        dcmSumCUp += _dt.Rows[i]["C_CapacitanceValue"] == null || _dt.Rows[i]["C_CapacitanceValue"].ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(_dt.Rows[i]["C_CapacitanceValue"].ToString().Trim());
                        intAvgCUpCount++;
                    }

                    dcmMeasureValue = _dt.Rows[i]["Q_FactorValue"] == null || _dt.Rows[i]["Q_FactorValue"].ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(_dt.Rows[i]["Q_FactorValue"].ToString().Trim());

                    dcmMeasureValue = dcmMeasureValue - dcmAvgQUp;

                    if (dcmMeasureValue < 0) dcmMeasureValue = -dcmMeasureValue;

                    dcmChkValue = dcmMeasureValue * (dcmReferenceValue / 100);

                    if (dcmMeasureValue > dcmChkValue)
                    {
                        dcmSumQUp += _dt.Rows[i]["Q_FactorValue"] == null || _dt.Rows[i]["Q_FactorValue"].ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(_dt.Rows[i]["Q_FactorValue"].ToString().Trim());
                        intAvgQUpCount++;
                    }
                }
                else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "이동")
                {
                    dcmMeasureValue = _dt.Rows[i]["DC_ResistanceValue"] == null || _dt.Rows[i]["DC_ResistanceValue"].ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(_dt.Rows[i]["DC_ResistanceValue"].ToString().Trim());

                    dcmMeasureValue = dcmMeasureValue - dcmAvgRdcMove;

                    if (dcmMeasureValue < 0) dcmMeasureValue = -dcmMeasureValue;

                    dcmChkValue = dcmMeasureValue * (dcmReferenceValue / 100);

                    if (dcmMeasureValue > dcmChkValue)
                    {
                        dcmSumRdcMove += _dt.Rows[i]["DC_Deviation"] == null || _dt.Rows[i]["DC_Deviation"].ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(_dt.Rows[i]["DC_Deviation"].ToString().Trim());
                        intAvgRdcMoveCount++;
                    }

                    dcmMeasureValue = _dt.Rows[i]["AC_ResistanceValue"] == null || _dt.Rows[i]["AC_ResistanceValue"].ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(_dt.Rows[i]["AC_ResistanceValue"].ToString().Trim());

                    dcmMeasureValue = dcmMeasureValue - dcmAvgRacMove;

                    if (dcmMeasureValue < 0) dcmMeasureValue = -dcmMeasureValue;

                    dcmChkValue = dcmMeasureValue * (dcmReferenceValue / 100);

                    if (dcmMeasureValue > dcmChkValue)
                    {
                        dcmSumRacMove += _dt.Rows[i]["AC_ResistanceValue"] == null || _dt.Rows[i]["AC_ResistanceValue"].ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(_dt.Rows[i]["AC_ResistanceValue"].ToString().Trim());
                        intAvgRacMoveCount++;
                    }

                    dcmMeasureValue = _dt.Rows[i]["L_InductanceValue"] == null || _dt.Rows[i]["L_InductanceValue"].ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(_dt.Rows[i]["L_InductanceValue"].ToString().Trim());

                    dcmMeasureValue = dcmMeasureValue - dcmAvgLMove;

                    if (dcmMeasureValue < 0) dcmMeasureValue = -dcmMeasureValue;

                    dcmChkValue = dcmMeasureValue * (dcmReferenceValue / 100);

                    if (dcmMeasureValue > dcmChkValue)
                    {
                        dcmSumLMove += _dt.Rows[i]["L_InductanceValue"] == null || _dt.Rows[i]["L_InductanceValue"].ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(_dt.Rows[i]["L_InductanceValue"].ToString().Trim());
                        intAvgLMoveCount++;
                    }

                    dcmMeasureValue = _dt.Rows[i]["C_CapacitanceValue"] == null || _dt.Rows[i]["C_CapacitanceValue"].ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(_dt.Rows[i]["C_CapacitanceValue"].ToString().Trim());

                    dcmMeasureValue = dcmMeasureValue - dcmAvgCMove;

                    if (dcmMeasureValue < 0) dcmMeasureValue = -dcmMeasureValue;

                    dcmChkValue = dcmMeasureValue * (dcmReferenceValue / 100);

                    if (dcmMeasureValue > dcmChkValue)
                    {
                        dcmSumCMove += _dt.Rows[i]["C_CapacitanceValue"] == null || _dt.Rows[i]["C_CapacitanceValue"].ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(_dt.Rows[i]["C_CapacitanceValue"].ToString().Trim());
                        intAvgCMoveCount++;
                    }

                    dcmMeasureValue = _dt.Rows[i]["Q_FactorValue"] == null || _dt.Rows[i]["Q_FactorValue"].ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(_dt.Rows[i]["Q_FactorValue"].ToString().Trim());

                    dcmMeasureValue = dcmMeasureValue - dcmAvgQMove;

                    if (dcmMeasureValue < 0) dcmMeasureValue = -dcmMeasureValue;

                    dcmChkValue = dcmMeasureValue * (dcmReferenceValue / 100);

                    if (dcmMeasureValue > dcmChkValue)
                    {
                        dcmSumQMove += _dt.Rows[i]["Q_FactorValue"] == null || _dt.Rows[i]["Q_FactorValue"].ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(_dt.Rows[i]["Q_FactorValue"].ToString().Trim());
                        intAvgQMoveCount++;
                    }
                }
            }

            // 정지 평균값 설정
            if (intAvgRdcStopCount > 0)
                dcmStopAveRdc = dcmSumRdcStop / intAvgRdcStopCount;
            else
                dcmStopAveRdc = 0.000M;

            if (intAvgRacStopCount > 0)
                dcmStopAveRac = dcmSumRacStop / intAvgRacStopCount;
            else
                dcmStopAveRac = 0.000M;

            if (intAvgLStopCount > 0)
                dcmStopAveL = dcmSumLStop / intAvgLStopCount;
            else
                dcmStopAveL = 0.000M;

            if (intAvgCStopCount > 0)
                dcmStopAveC = dcmSumCStop / intAvgCStopCount;
            else
                dcmStopAveC = 0.000M;

            if (intAvgQStopCount > 0)
                dcmStopAveQ = dcmSumQStop / intAvgQStopCount;
            else
                dcmStopAveQ = 0.000M;

            // 올림 평균값 설정
            if (intAvgRdcUpCount > 0)
                dcmUpAveRdc = dcmSumRdcUp / intAvgRdcUpCount;
            else
                dcmUpAveRdc = 0.000M;

            if (intAvgRacUpCount > 0)
                dcmUpAveRac = dcmSumRacUp / intAvgRacUpCount;
            else
                dcmUpAveRac = 0.000M;

            if (intAvgLUpCount > 0)
                dcmUpAveL = dcmSumLUp / intAvgLUpCount;
            else
                dcmUpAveL = 0.000M;

            if (intAvgCUpCount > 0)
                dcmUpAveC = dcmSumCUp / intAvgCUpCount;
            else
                dcmUpAveC = 0.000M;

            if (intAvgQUpCount > 0)
                dcmUpAveQ = dcmSumQUp / intAvgQUpCount;
            else
                dcmUpAveQ = 0.000M;

            // 이동 평균값 설정
            if (intAvgRdcMoveCount > 0)
                dcmMoveAveRdc = dcmSumRdcMove / intAvgRdcMoveCount;
            else
                dcmMoveAveRdc = 0.000M;

            if (intAvgRacMoveCount > 0)
                dcmMoveAveRac = dcmSumRacMove / intAvgRacMoveCount;
            else
                dcmMoveAveRac = 0.000M;

            if (intAvgLMoveCount > 0)
                dcmMoveAveL = dcmSumLMove / intAvgLMoveCount;
            else
                dcmMoveAveL = 0.000M;

            if (intAvgCMoveCount > 0)
                dcmMoveAveC = dcmSumCMove / intAvgCMoveCount;
            else
                dcmMoveAveC = 0.000M;

            if (intAvgQMoveCount > 0)
                dcmMoveAveQ = dcmSumQMove / intAvgQMoveCount;
            else
                dcmMoveAveQ = 0.000M;

            return;
        }

        /// <summary>
        /// 표 형식 / 평균 값(RCS) 계산
        /// </summary>
        /// <param name="_dt"></param>
        private void SetRCSReferenceValueTableMark(ref DataTable _dt, string _strComparisonTarget)
        {
            try
            {
                decimal dcmDecision_StopRdc = 0.000M, dcmDecision_StopRac = 0.000M, dcmDecision_StopL = 0.000M, dcmDecision_StopC = 0.000M, dcmDecision_StopQ = 0.000M
                    , dcmDecision_UpRdc = 0.000M, dcmDecision_UpRac = 0.000M, dcmDecision_UpL = 0.000M, dcmDecision_UpC = 0.000M, dcmDecision_UpQ = 0.000M
                    , dcmDecision_MoveRdc = 0.000M, dcmDecision_MoveRac = 0.000M, dcmDecision_MoveL = 0.000M, dcmDecision_MoveC = 0.000M, dcmDecision_MoveQ = 0.000M;

                // 비교대상
                if (_strComparisonTarget.Trim() == "평균값")
                {
                    // 평균값
                    dcmDecision_StopRdc = dcmStopAveRdc;
                    dcmDecision_StopRac = dcmStopAveRac;
                    dcmDecision_StopL = dcmStopAveL;
                    dcmDecision_StopC = dcmStopAveC;
                    dcmDecision_StopQ = dcmStopAveQ;

                    dcmDecision_UpRdc = dcmUpAveRdc;
                    dcmDecision_UpRac = dcmUpAveRac;
                    dcmDecision_UpL = dcmUpAveL;
                    dcmDecision_UpC = dcmUpAveC;
                    dcmDecision_UpQ = dcmUpAveQ;

                    dcmDecision_MoveRdc = dcmMoveAveRdc;
                    dcmDecision_MoveRac = dcmMoveAveRac;
                    dcmDecision_MoveL = dcmMoveAveL;
                    dcmDecision_MoveC = dcmMoveAveC;
                    dcmDecision_MoveQ = dcmMoveAveQ;
                }
                else
                {
                    // 기준값
                    dcmDecision_StopRdc = dcmStopRefRdc;
                    dcmDecision_StopRac = dcmStopRefRac;
                    dcmDecision_StopL = dcmStopRefL;
                    dcmDecision_StopC = dcmStopRefC;
                    dcmDecision_StopQ = dcmStopRefQ;

                    dcmDecision_UpRdc = dcmUpRefRdc;
                    dcmDecision_UpRac = dcmUpRefRac;
                    dcmDecision_UpL = dcmUpRefL;
                    dcmDecision_UpC = dcmUpRefC;
                    dcmDecision_UpQ = dcmUpRefQ;

                    dcmDecision_MoveRdc = dcmMoveRefeRdc;
                    dcmDecision_MoveRac = dcmMoveRefRac;
                    dcmDecision_MoveL = dcmMoveRefL;
                    dcmDecision_MoveC = dcmMoveRefC;
                    dcmDecision_MoveQ = dcmMoveRefQ;
                }

                decimal dcmDecisionRange_Rdc = 0M, dcmDecisionRange_Rac = 0M, dcmDecisionRange_L = 0M, dcmDecisionRange_C = 0M
                    , dcmDecisionRange_Q = 0M, dcmEffectiveStandardRange = 0M;
                dcmDecisionRange_Rdc = Convert.ToDecimal(Gini.GetValue("ReferenceValue", "RdcDecisionRange_ReferenceValue"));
                dcmDecisionRange_Rac = Convert.ToDecimal(Gini.GetValue("ReferenceValue", "RacDecisionRange_ReferenceValue"));
                dcmDecisionRange_L = Convert.ToDecimal(Gini.GetValue("ReferenceValue", "LDecisionRange_ReferenceValue"));
                dcmDecisionRange_C = Convert.ToDecimal(Gini.GetValue("ReferenceValue", "CDecisionRange_ReferenceValue"));
                dcmDecisionRange_Q = Convert.ToDecimal(Gini.GetValue("ReferenceValue", "QDecisionRange_ReferenceValue"));
                dcmEffectiveStandardRange = Convert.ToDecimal(Gini.GetValue("ReferenceValue", "EffectiveStandardRangeOfVariation"));

                decimal dcmSumRdcStop = 0.000M, dcmSumRacStop = 0.000M, dcmSumLStop = 0.000M, dcmSumCStop = 0.000M, dcmSumQcStop = 0.000M;
                decimal dcmSumRdcUp = 0.000M, dcmSumRacUp = 0.000M, dcmSumLUp = 0.000M, dcmSumCUp = 0.000M, dcmSumQcUp = 0.000M;
                decimal dcmSumRdcMove = 0.000M, dcmSumRacMove = 0.000M, dcmSumLMove = 0.000M, dcmSumCMove = 0.000M, dcmSumQcMove = 0.000M;
                int intSumStopCount = 0, intSumUpCount = 0, intSumMoveCount = 0, intPASS = 0, intFAIL = 0, intDoubt = 0;
                string strRdcResult = "", strRacResult = "", strLResult = "", strCResult = "", strQResult = "";
                decimal dcmDecision_Rdc = 0M, dcmDecision_Rac = 0M, dcmDecision_L = 0M, dcmDecision_C = 0M, dcmDecision_Q = 0M;

                for (int i = 0; i < _dt.Rows.Count; i++)
                {
                    int iRow = dgvMeasurement.Rows.Add();

                    dgvMeasurement.Rows[iRow].Cells["ControlName"].Value = _dt.Rows[i]["ControlName"].ToString().Trim();
                    dgvMeasurement.Rows[iRow].Cells["CoilName"].Value = _dt.Rows[i]["CoilName"].ToString().Trim();
                    dgvMeasurement.Rows[iRow].Cells["CoilNumber"].Value = _dt.Rows[i]["CoilNumber"].ToString().Trim();
                    dgvMeasurement.Rows[iRow].Cells["Result"].Value = "";

                    #region Rdc 체크
                    decimal dcmRdcValue = _dt.Rows[i]["DC_ResistanceValue"] == null || _dt.Rows[i]["DC_ResistanceValue"].ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(_dt.Rows[i]["DC_ResistanceValue"].ToString().Trim());
                    decimal dcmRdcDeviationValue = _dt.Rows[i]["DC_Deviation"] == null || _dt.Rows[i]["DC_Deviation"].ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(_dt.Rows[i]["DC_Deviation"].ToString().Trim());

                    if (_dt.Rows[i]["DC_ResistanceValue"] == null || _dt.Rows[i]["DC_ResistanceValue"].ToString().Trim() == "")
                    {
                        strRdcResult = "부적합";
                    }
                    else
                    {
                        if (_dt.Rows[i]["CoilName"].ToString().Trim() == "정지")
                        {
                            dcmDecision_Rdc = dcmDecision_StopRdc;

                            strRdcResult = m_MeasureProcess.GetMesurmentValueDecision(dcmRdcValue, dcmDecision_StopRdc, dcmDecisionRange_Rdc, dcmEffectiveStandardRange, "DC");
                        }
                        else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "올림")
                        {
                            dcmDecision_Rdc = dcmDecision_UpRdc;

                            strRdcResult = m_MeasureProcess.GetMesurmentValueDecision(dcmRdcValue, dcmDecision_UpRdc, dcmDecisionRange_Rdc, dcmEffectiveStandardRange, "DC");
                        }
                        else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "이동")
                        {
                            dcmDecision_Rdc = dcmDecision_MoveRdc;

                            strRdcResult = m_MeasureProcess.GetMesurmentValueDecision(dcmRdcValue, dcmDecision_MoveRdc, dcmDecisionRange_Rdc, dcmEffectiveStandardRange, "DC");
                        }
                        else
                            strRdcResult = "부적합";
                    }

                    dgvMeasurement.Rows[iRow].Cells["DC_ResistanceValue"].Value = dcmRdcValue.ToString("F3").Trim();
                    dgvMeasurement.Rows[iRow].Cells["DC_Deviation"].Value = dcmRdcDeviationValue.ToString("F3").Trim();

                    if (dcmRdcDeviationValue == 0)
                        dgvMeasurement.Rows[iRow].Cells["DC_ResistanceValue"].Style.ForeColor = System.Drawing.Color.Black;
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
                    decimal dcmRacDeviationValue = _dt.Rows[i]["AC_Deviation"] == null || _dt.Rows[i]["AC_Deviation"].ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(_dt.Rows[i]["AC_Deviation"].ToString().Trim());

                    if (_dt.Rows[i]["AC_ResistanceValue"] == null || _dt.Rows[i]["AC_ResistanceValue"].ToString().Trim() == "")
                    {
                        strRacResult = "부적합";
                    }
                    else
                    {
                        if (_dt.Rows[i]["CoilName"].ToString().Trim() == "정지")
                        {
                            dcmDecision_Rac = dcmDecision_StopRac;

                            strRacResult = m_MeasureProcess.GetMesurmentValueDecision(dcmRacValue, dcmDecision_StopRac, dcmDecisionRange_Rac, dcmEffectiveStandardRange, "AC");
                        }
                        else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "올림")
                        {
                            dcmDecision_Rac = dcmDecision_UpRac;

                            strRacResult = m_MeasureProcess.GetMesurmentValueDecision(dcmRacValue, dcmDecision_UpRac, dcmDecisionRange_Rac, dcmEffectiveStandardRange, "AC");
                        }
                        else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "이동")
                        {
                            dcmDecision_Rac = dcmDecision_MoveRac;

                            strRacResult = m_MeasureProcess.GetMesurmentValueDecision(dcmRacValue, dcmDecision_MoveRac, dcmDecisionRange_Rac, dcmEffectiveStandardRange, "AC");
                        }
                        else
                            strRacResult = "부적합";
                    }

                    dgvMeasurement.Rows[iRow].Cells["AC_ResistanceValue"].Value = dcmRacValue.ToString("F3").Trim();
                    dgvMeasurement.Rows[iRow].Cells["AC_Deviation"].Value = dcmRacDeviationValue.ToString("F3").Trim();

                    if (dcmRacDeviationValue == 0)
                        dgvMeasurement.Rows[iRow].Cells["AC_ResistanceValue"].Style.ForeColor = System.Drawing.Color.Black;
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
                    decimal dcmLDeviationValue = _dt.Rows[i]["L_Deviation"] == null || _dt.Rows[i]["L_Deviation"].ToString().Trim() == ""
                        ? 0.000M : Convert.ToDecimal(_dt.Rows[i]["L_Deviation"].ToString().Trim());

                    if (_dt.Rows[i]["L_InductanceValue"] == null || _dt.Rows[i]["L_InductanceValue"].ToString().Trim() == "")
                    {
                        strLResult = "부적합";
                    }
                    else
                    {
                        if (_dt.Rows[i]["CoilName"].ToString().Trim() == "정지")
                        {
                            dcmDecision_L = dcmDecision_StopL;

                            strLResult = m_MeasureProcess.GetMesurmentValueDecision(dcmLValue, dcmDecision_StopL, dcmDecisionRange_L, dcmEffectiveStandardRange, "L");
                        }
                        else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "올림")
                        {
                            dcmDecision_L = dcmDecision_UpL;

                            strLResult = m_MeasureProcess.GetMesurmentValueDecision(dcmLValue, dcmDecision_UpL, dcmDecisionRange_L, dcmEffectiveStandardRange, "L");
                        }
                        else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "이동")
                        {
                            dcmDecision_L = dcmDecision_MoveL;

                            strLResult = m_MeasureProcess.GetMesurmentValueDecision(dcmLValue, dcmDecision_MoveL, dcmDecisionRange_L, dcmEffectiveStandardRange, "L");
                        }
                        else
                            strLResult = "부적합";
                    }

                    dgvMeasurement.Rows[iRow].Cells["L_InductanceValue"].Value = dcmLValue.ToString("F3").Trim();
                    dgvMeasurement.Rows[iRow].Cells["L_Deviation"].Value = dcmLDeviationValue.ToString("F3").Trim();

                    if (dcmLDeviationValue == 0)
                        dgvMeasurement.Rows[iRow].Cells["L_InductanceValue"].Style.ForeColor = System.Drawing.Color.Black;
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
                    decimal dcmCDeviationValue = _dt.Rows[i]["C_Deviation"] == null || _dt.Rows[i]["C_Deviation"].ToString().Trim() == ""
                        ? 0.000000M : Convert.ToDecimal(_dt.Rows[i]["C_Deviation"].ToString().Trim());

                    if (_dt.Rows[i]["C_CapacitanceValue"] == null || _dt.Rows[i]["C_CapacitanceValue"].ToString().Trim() == "")
                    {
                        strCResult = "부적합";
                    }
                    else
                    {
                        if (_dt.Rows[i]["CoilName"].ToString().Trim() == "정지")
                        {
                            dcmDecision_C = dcmDecision_StopC;

                            strCResult = m_MeasureProcess.GetMesurmentValueDecision(dcmCValue, dcmDecision_StopC, dcmDecisionRange_C, dcmEffectiveStandardRange, "C");
                        }
                        else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "올림")
                        {
                            dcmDecision_C = dcmDecision_UpC;

                            strCResult = m_MeasureProcess.GetMesurmentValueDecision(dcmCValue, dcmDecision_UpC, dcmDecisionRange_C, dcmEffectiveStandardRange, "C");
                        }
                        else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "이동")
                        {
                            dcmDecision_C = dcmDecision_MoveC;

                            strCResult = m_MeasureProcess.GetMesurmentValueDecision(dcmCValue, dcmDecision_MoveC, dcmDecisionRange_C, dcmEffectiveStandardRange, "C");
                        }
                        else
                            strCResult = "부적합";
                    }

                    dgvMeasurement.Rows[iRow].Cells["C_CapacitanceValue"].Value = dcmCValue.ToString("F6").Trim();
                    dgvMeasurement.Rows[iRow].Cells["C_Deviation"].Value = dcmCDeviationValue.ToString("F6").Trim();

                    if (dcmCDeviationValue == 0)
                        dgvMeasurement.Rows[iRow].Cells["C_CapacitanceValue"].Style.ForeColor = System.Drawing.Color.Black;
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
                    decimal dcmQDeviationValue = _dt.Rows[i]["Q_Deviation"] == null || _dt.Rows[i]["Q_Deviation"].ToString().Trim() == ""
                        ? 0.0M : Convert.ToDecimal(_dt.Rows[i]["Q_Deviation"].ToString().Trim());

                    if (_dt.Rows[i]["Q_FactorValue"] == null || _dt.Rows[i]["Q_FactorValue"].ToString().Trim() == "")
                    {
                        strQResult = "부적합";
                    }
                    else
                    {
                        if (_dt.Rows[i]["CoilName"].ToString().Trim() == "정지")
                        {
                            dcmDecision_Q = dcmDecision_StopQ;

                            strQResult = m_MeasureProcess.GetMesurmentValueDecision(dcmQValue, dcmDecision_StopQ, dcmDecisionRange_Q, dcmEffectiveStandardRange, "Q");
                        }
                        else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "올림")
                        {
                            dcmDecision_Q = dcmDecision_UpQ;

                            strQResult = m_MeasureProcess.GetMesurmentValueDecision(dcmQValue, dcmDecision_UpQ, dcmDecisionRange_Q, dcmEffectiveStandardRange, "Q");
                        }
                        else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "이동")
                        {
                            dcmDecision_Q = dcmDecision_MoveQ;

                            strQResult = m_MeasureProcess.GetMesurmentValueDecision(dcmQValue, dcmDecision_MoveQ, dcmDecisionRange_Q, dcmEffectiveStandardRange, "Q");
                        }
                        else
                            strQResult = "부적합";
                    }

                    dgvMeasurement.Rows[iRow].Cells["Q_FactorValue"].Value = dcmQValue.ToString("F3").Trim();
                    dgvMeasurement.Rows[iRow].Cells["Q_Deviation"].Value = dcmQDeviationValue.ToString("F3").Trim();

                    if (dcmQDeviationValue == 0)
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

                    dgvMeasurement.Rows[iRow].Cells["PlantName"].Value = _dt.Rows[i]["PlantName"].ToString().Trim();
                    dgvMeasurement.Rows[iRow].Cells["Hogi"].Value = _dt.Rows[i]["Hogi"].ToString().Trim();
                    dgvMeasurement.Rows[iRow].Cells["Oh_Degree"].Value = _dt.Rows[i]["Oh_Degree"].ToString().Trim();
                    dgvMeasurement.Rows[iRow].Cells["PowerCabinet"].Value = _dt.Rows[i]["PowerCabinet"].ToString().Trim();

                    // 평균 구하기위해 합계 산출
                    if (_dt.Rows[i]["CoilName"].ToString().Trim() == "정지")
                    {
                        dcmSumRdcStop += dgvMeasurement.Rows[i].Cells["DC_ResistanceValue"].Value == null || dgvMeasurement.Rows[i].Cells["DC_ResistanceValue"].Value.ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(dgvMeasurement.Rows[i].Cells["DC_ResistanceValue"].Value.ToString().Trim());
                        dcmSumRacStop += dgvMeasurement.Rows[i].Cells["AC_ResistanceValue"].Value == null || dgvMeasurement.Rows[i].Cells["AC_ResistanceValue"].Value.ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(dgvMeasurement.Rows[i].Cells["AC_ResistanceValue"].Value.ToString().Trim());
                        dcmSumLStop += dgvMeasurement.Rows[i].Cells["L_InductanceValue"].Value == null || dgvMeasurement.Rows[i].Cells["L_InductanceValue"].Value.ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(dgvMeasurement.Rows[i].Cells["L_InductanceValue"].Value.ToString().Trim());
                        dcmSumCStop += dgvMeasurement.Rows[i].Cells["C_CapacitanceValue"].Value == null || dgvMeasurement.Rows[i].Cells["C_CapacitanceValue"].Value.ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(dgvMeasurement.Rows[i].Cells["C_CapacitanceValue"].Value.ToString().Trim());
                        dcmSumQcStop += dgvMeasurement.Rows[i].Cells["Q_FactorValue"].Value == null || dgvMeasurement.Rows[i].Cells["Q_FactorValue"].Value.ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(dgvMeasurement.Rows[i].Cells["Q_FactorValue"].Value.ToString().Trim());
                        intSumStopCount++;
                    }
                    else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "올림")
                    {
                        dcmSumRdcUp += dgvMeasurement.Rows[i].Cells["DC_ResistanceValue"].Value == null || dgvMeasurement.Rows[i].Cells["DC_ResistanceValue"].Value.ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(dgvMeasurement.Rows[i].Cells["DC_ResistanceValue"].Value.ToString().Trim());
                        dcmSumRacUp += dgvMeasurement.Rows[i].Cells["AC_ResistanceValue"].Value == null || dgvMeasurement.Rows[i].Cells["AC_ResistanceValue"].Value.ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(dgvMeasurement.Rows[i].Cells["AC_ResistanceValue"].Value.ToString().Trim());
                        dcmSumLUp += dgvMeasurement.Rows[i].Cells["L_InductanceValue"].Value == null || dgvMeasurement.Rows[i].Cells["L_InductanceValue"].Value.ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(dgvMeasurement.Rows[i].Cells["L_InductanceValue"].Value.ToString().Trim());
                        dcmSumCUp += dgvMeasurement.Rows[i].Cells["C_CapacitanceValue"].Value == null || dgvMeasurement.Rows[i].Cells["C_CapacitanceValue"].Value.ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(dgvMeasurement.Rows[i].Cells["C_CapacitanceValue"].Value.ToString().Trim());
                        dcmSumQcUp += dgvMeasurement.Rows[i].Cells["Q_FactorValue"].Value == null || dgvMeasurement.Rows[i].Cells["Q_FactorValue"].Value.ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(dgvMeasurement.Rows[i].Cells["Q_FactorValue"].Value.ToString().Trim());
                        intSumUpCount++;
                    }
                    else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "이동")
                    {
                        dcmSumRdcMove += dgvMeasurement.Rows[i].Cells["DC_ResistanceValue"].Value == null || dgvMeasurement.Rows[i].Cells["DC_ResistanceValue"].Value.ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(dgvMeasurement.Rows[i].Cells["DC_ResistanceValue"].Value.ToString().Trim());
                        dcmSumRacMove += dgvMeasurement.Rows[i].Cells["AC_ResistanceValue"].Value == null || dgvMeasurement.Rows[i].Cells["AC_ResistanceValue"].Value.ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(dgvMeasurement.Rows[i].Cells["AC_ResistanceValue"].Value.ToString().Trim());
                        dcmSumLMove += dgvMeasurement.Rows[i].Cells["L_InductanceValue"].Value == null || dgvMeasurement.Rows[i].Cells["L_InductanceValue"].Value.ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(dgvMeasurement.Rows[i].Cells["L_InductanceValue"].Value.ToString().Trim());
                        dcmSumCMove += dgvMeasurement.Rows[i].Cells["C_CapacitanceValue"].Value == null || dgvMeasurement.Rows[i].Cells["C_CapacitanceValue"].Value.ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(dgvMeasurement.Rows[i].Cells["C_CapacitanceValue"].Value.ToString().Trim());
                        dcmSumQcMove += dgvMeasurement.Rows[i].Cells["Q_FactorValue"].Value == null || dgvMeasurement.Rows[i].Cells["Q_FactorValue"].Value.ToString().Trim() == ""
                            ? 0.000M : Convert.ToDecimal(dgvMeasurement.Rows[i].Cells["Q_FactorValue"].Value.ToString().Trim());
                        intSumMoveCount++;
                    }
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
        public void SetRCSReferenceValueMeasurement(DataTable _dt, string[] _arrayXLabelValue, int _intXLabelValueCount)
        {
            try
            {
                double[] dStopRdc = new double[_intXLabelValueCount];
                double[] dUpRdc = new double[_intXLabelValueCount];
                double[] dMoveRdc = new double[_intXLabelValueCount];

                double[] dStopRac = new double[_intXLabelValueCount];
                double[] dUpRac = new double[_intXLabelValueCount];
                double[] dMoveRac = new double[_intXLabelValueCount];

                double[] dStopL = new double[_intXLabelValueCount];
                double[] dUpL = new double[_intXLabelValueCount];
                double[] dMoveL = new double[_intXLabelValueCount];

                double[] dStopC = new double[_intXLabelValueCount];
                double[] dUpC = new double[_intXLabelValueCount];
                double[] dMoveC = new double[_intXLabelValueCount];

                double[] dStopQ = new double[_intXLabelValueCount];
                double[] dUpQ = new double[_intXLabelValueCount];
                double[] dMoveQ = new double[_intXLabelValueCount];

                double[] dDevStopRdc = new double[_intXLabelValueCount];
                double[] dDevUpRdc = new double[_intXLabelValueCount];
                double[] dDevMoveRdc = new double[_intXLabelValueCount];
                double[] dDevStopRac = new double[_intXLabelValueCount];
                double[] dDevUpRac = new double[_intXLabelValueCount];
                double[] dDevMoveRac = new double[_intXLabelValueCount];

                double[] dDevStopL = new double[_intXLabelValueCount];
                double[] dDevUpL = new double[_intXLabelValueCount];
                double[] dDevMoveL = new double[_intXLabelValueCount];
                double[] dDevStopC = new double[_intXLabelValueCount];
                double[] dDevUpC = new double[_intXLabelValueCount];
                double[] dDevMoveC = new double[_intXLabelValueCount];
                double[] dDevStopQ = new double[_intXLabelValueCount];
                double[] dDevUpQ = new double[_intXLabelValueCount];
                double[] dDevMoveQ = new double[_intXLabelValueCount];

                double dRdcValue = 0d, dRacValue = 0d, dLValue = 0d, dCValue = 0d, dQValue = 0d
                    , dDevRdcValue = 0d, dDevRacValue = 0d, dDevLValue = 0d, dDevCValue = 0d, dDevQValue = 0d;
                int intStopIndex = 0, intUpIndex = 0, intMoveIndex = 0;

                for (int i = 0; i < _dt.Rows.Count; i++)
                {
                    if (_dt.Rows[i]["CoilName"].ToString().Trim() == "정지")
                    {
                        dRdcValue = _dt.Rows[i]["DC_ResistanceValue"] == null || _dt.Rows[i]["DC_ResistanceValue"].ToString().Trim() == ""
                            ? 0.000d : Convert.ToDouble(_dt.Rows[i]["DC_ResistanceValue"].ToString().Trim());
                        dStopRdc[intStopIndex] = dRdcValue;

                        dRacValue = _dt.Rows[i]["AC_ResistanceValue"] == null || _dt.Rows[i]["AC_ResistanceValue"].ToString().Trim() == ""
                            ? 0.000d : Convert.ToDouble(_dt.Rows[i]["AC_ResistanceValue"].ToString().Trim());
                        dStopRac[intStopIndex] = dRacValue;

                        dLValue = _dt.Rows[i]["L_InductanceValue"] == null || _dt.Rows[i]["L_InductanceValue"].ToString().Trim() == ""
                            ? 0.000d : Convert.ToDouble(_dt.Rows[i]["L_InductanceValue"].ToString().Trim());
                        dStopL[intStopIndex] = dLValue;

                        dCValue = _dt.Rows[i]["C_CapacitanceValue"] == null || _dt.Rows[i]["C_CapacitanceValue"].ToString().Trim() == ""
                            ? 0.000d : Convert.ToDouble(_dt.Rows[i]["C_CapacitanceValue"].ToString().Trim());
                        dStopC[intStopIndex] = dCValue;

                        dQValue = _dt.Rows[i]["Q_FactorValue"] == null || _dt.Rows[i]["Q_FactorValue"].ToString().Trim() == ""
                            ? 0.000d : Convert.ToDouble(_dt.Rows[i]["Q_FactorValue"].ToString().Trim());
                        dStopQ[intStopIndex] = dQValue;

                        dDevRdcValue = _dt.Rows[i]["DC_Deviation"] == null || _dt.Rows[i]["DC_Deviation"].ToString().Trim() == ""
                            ? 0.000d : Convert.ToDouble(_dt.Rows[i]["DC_Deviation"].ToString().Trim());
                        dDevStopRdc[intStopIndex] = dDevRdcValue;

                        dDevRacValue = _dt.Rows[i]["AC_Deviation"] == null || _dt.Rows[i]["AC_Deviation"].ToString().Trim() == ""
                            ? 0.000d : Convert.ToDouble(_dt.Rows[i]["AC_Deviation"].ToString().Trim());
                        dDevStopRac[intStopIndex] = dDevRacValue;

                        dDevLValue = _dt.Rows[i]["L_Deviation"] == null || _dt.Rows[i]["L_Deviation"].ToString().Trim() == ""
                            ? 0.000d : Convert.ToDouble(_dt.Rows[i]["L_Deviation"].ToString().Trim());
                        dDevStopL[intStopIndex] = dDevLValue;

                        dDevCValue = _dt.Rows[i]["C_Deviation"] == null || _dt.Rows[i]["C_Deviation"].ToString().Trim() == ""
                            ? 0.000d : Convert.ToDouble(_dt.Rows[i]["C_Deviation"].ToString().Trim());
                        dDevStopC[intStopIndex] = dDevCValue;

                        dDevQValue = _dt.Rows[i]["Q_Deviation"] == null || _dt.Rows[i]["Q_Deviation"].ToString().Trim() == ""
                            ? 0.000d : Convert.ToDouble(_dt.Rows[i]["Q_Deviation"].ToString().Trim());
                        dDevStopQ[intStopIndex] = dDevQValue;

                        intStopIndex++;
                    }
                    else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "올림")
                    {
                        dRdcValue = _dt.Rows[i]["DC_ResistanceValue"] == null || _dt.Rows[i]["DC_ResistanceValue"].ToString().Trim() == ""
                        ? 0.000d : Convert.ToDouble(_dt.Rows[i]["DC_ResistanceValue"].ToString().Trim());
                        dUpRdc[intUpIndex] = dRdcValue;

                        dRacValue = _dt.Rows[i]["AC_ResistanceValue"] == null || _dt.Rows[i]["AC_ResistanceValue"].ToString().Trim() == ""
                            ? 0.000d : Convert.ToDouble(_dt.Rows[i]["AC_ResistanceValue"].ToString().Trim());
                        dUpRac[intUpIndex] = dRacValue;

                        dLValue = _dt.Rows[i]["L_InductanceValue"] == null || _dt.Rows[i]["L_InductanceValue"].ToString().Trim() == ""
                        ? 0.000d : Convert.ToDouble(_dt.Rows[i]["L_InductanceValue"].ToString().Trim());
                        dUpL[intUpIndex] = dLValue;

                        dCValue = _dt.Rows[i]["C_CapacitanceValue"] == null || _dt.Rows[i]["C_CapacitanceValue"].ToString().Trim() == ""
                            ? 0.000d : Convert.ToDouble(_dt.Rows[i]["C_CapacitanceValue"].ToString().Trim());
                        dUpC[intUpIndex] = dCValue;

                        dQValue = _dt.Rows[i]["Q_FactorValue"] == null || _dt.Rows[i]["Q_FactorValue"].ToString().Trim() == ""
                            ? 0.000d : Convert.ToDouble(_dt.Rows[i]["Q_FactorValue"].ToString().Trim());
                        dUpQ[intUpIndex] = dQValue;

                        dDevRdcValue = _dt.Rows[i]["DC_Deviation"] == null || _dt.Rows[i]["DC_Deviation"].ToString().Trim() == ""
                        ? 0.000d : Convert.ToDouble(_dt.Rows[i]["DC_Deviation"].ToString().Trim());
                        dDevUpRdc[intUpIndex] = dDevRdcValue;

                        dDevRacValue = _dt.Rows[i]["AC_Deviation"] == null || _dt.Rows[i]["AC_Deviation"].ToString().Trim() == ""
                            ? 0.000d : Convert.ToDouble(_dt.Rows[i]["AC_Deviation"].ToString().Trim());
                        dDevUpRac[intUpIndex] = dDevRacValue;

                        dDevLValue = _dt.Rows[i]["L_Deviation"] == null || _dt.Rows[i]["L_Deviation"].ToString().Trim() == ""
                        ? 0.000d : Convert.ToDouble(_dt.Rows[i]["L_Deviation"].ToString().Trim());
                        dDevUpL[intUpIndex] = dDevLValue;

                        dDevCValue = _dt.Rows[i]["C_Deviation"] == null || _dt.Rows[i]["C_Deviation"].ToString().Trim() == ""
                            ? 0.000d : Convert.ToDouble(_dt.Rows[i]["C_Deviation"].ToString().Trim());
                        dDevUpC[intUpIndex] = dDevCValue;

                        dDevQValue = _dt.Rows[i]["Q_Deviation"] == null || _dt.Rows[i]["Q_Deviation"].ToString().Trim() == ""
                            ? 0.000d : Convert.ToDouble(_dt.Rows[i]["Q_Deviation"].ToString().Trim());
                        dDevUpQ[intUpIndex] = dDevQValue;

                        intUpIndex++;
                    }
                    else if (_dt.Rows[i]["CoilName"].ToString().Trim() == "이동")
                    {
                        dRdcValue = _dt.Rows[i]["DC_ResistanceValue"] == null || _dt.Rows[i]["DC_ResistanceValue"].ToString().Trim() == ""
                        ? 0.000d : Convert.ToDouble(_dt.Rows[i]["DC_ResistanceValue"].ToString().Trim());
                        dMoveRdc[intMoveIndex] = dRdcValue;

                        dRacValue = _dt.Rows[i]["AC_ResistanceValue"] == null || _dt.Rows[i]["AC_ResistanceValue"].ToString().Trim() == ""
                            ? 0.000d : Convert.ToDouble(_dt.Rows[i]["AC_ResistanceValue"].ToString().Trim());
                        dMoveRac[intMoveIndex] = dRacValue;

                        dLValue = _dt.Rows[i]["L_InductanceValue"] == null || _dt.Rows[i]["L_InductanceValue"].ToString().Trim() == ""
                        ? 0.000d : Convert.ToDouble(_dt.Rows[i]["L_InductanceValue"].ToString().Trim());
                        dMoveL[intMoveIndex] = dLValue;

                        dCValue = _dt.Rows[i]["C_CapacitanceValue"] == null || _dt.Rows[i]["C_CapacitanceValue"].ToString().Trim() == ""
                            ? 0.000d : Convert.ToDouble(_dt.Rows[i]["C_CapacitanceValue"].ToString().Trim());
                        dMoveC[intMoveIndex] = dCValue;

                        dQValue = _dt.Rows[i]["Q_FactorValue"] == null || _dt.Rows[i]["Q_FactorValue"].ToString().Trim() == ""
                            ? 0.000d : Convert.ToDouble(_dt.Rows[i]["Q_FactorValue"].ToString().Trim());
                        dMoveQ[intMoveIndex] = dQValue;

                        dDevRdcValue = _dt.Rows[i]["DC_Deviation"] == null || _dt.Rows[i]["DC_Deviation"].ToString().Trim() == ""
                        ? 0.000d : Convert.ToDouble(_dt.Rows[i]["DC_Deviation"].ToString().Trim());
                        dDevMoveRdc[intMoveIndex] = dDevRdcValue;

                        dDevRacValue = _dt.Rows[i]["AC_Deviation"] == null || _dt.Rows[i]["AC_Deviation"].ToString().Trim() == ""
                            ? 0.000d : Convert.ToDouble(_dt.Rows[i]["AC_Deviation"].ToString().Trim());
                        dDevMoveRac[intMoveIndex] = dDevRacValue;

                        dDevLValue = _dt.Rows[i]["L_Deviation"] == null || _dt.Rows[i]["L_Deviation"].ToString().Trim() == ""
                        ? 0.000d : Convert.ToDouble(_dt.Rows[i]["L_Deviation"].ToString().Trim());
                        dDevMoveL[intMoveIndex] = dDevLValue;

                        dDevCValue = _dt.Rows[i]["C_Deviation"] == null || _dt.Rows[i]["C_Deviation"].ToString().Trim() == ""
                            ? 0.000d : Convert.ToDouble(_dt.Rows[i]["C_Deviation"].ToString().Trim());
                        dDevMoveC[intMoveIndex] = dDevCValue;

                        dDevQValue = _dt.Rows[i]["Q_Deviation"] == null || _dt.Rows[i]["Q_Deviation"].ToString().Trim() == ""
                            ? 0.000d : Convert.ToDouble(_dt.Rows[i]["Q_Deviation"].ToString().Trim());
                        dDevMoveQ[intMoveIndex] = dDevQValue;

                        intMoveIndex++;
                    }
                }

                // DC 저항 Chart 그리기  
                SetChart_Depict(chartMeasurementValueDC, "DC 저항", dStopRdc, dUpRdc, dMoveRdc, _arrayXLabelValue, 0, "", "");

                // AC 저항 Chart 그리기  
                SetChart_Depict(chartMeasurementValueAC, "AC 저항", dStopRac, dUpRac, dMoveRac, _arrayXLabelValue, 0, "", "");

                // L 저항 Chart 그리기  
                SetChart_Depict(chartMeasurementValueL, "인덕턴스", dStopL, dUpL, dMoveL, _arrayXLabelValue, 0, "", "");

                // C 저항 Chart 그리기  
                SetChart_Depict(chartMeasurementValueC, "캐패시턴스", dStopC, dUpC, dMoveC, _arrayXLabelValue, 0, "", "");

                // Q 저항 Chart 그리기  
                SetChart_Depict(chartMeasurementValueQ, "Q-Factor", dStopQ, dUpQ, dMoveQ, _arrayXLabelValue, 0, "", "");

                // DC 편차 Chart 그리기  
                SetChart_Depict(chartAverageValueDC, "DC 저항", dDevStopRdc, dDevUpRdc, dDevMoveRdc, _arrayXLabelValue, 0, "", "");

                // AC 편차 Chart 그리기  
                SetChart_Depict(chartAverageValueAC, "AC 저항", dDevStopRac, dDevUpRac, dDevMoveRac, _arrayXLabelValue, 0, "", "");

                // L 저항 Chart 그리기  
                SetChart_Depict(chartAverageValueL, "인덕턴스 편차", dDevStopL, dDevUpL, dDevMoveL, _arrayXLabelValue, 0, "", "");

                // C 저항 Chart 그리기  
                SetChart_Depict(chartAverageValueC, "캐패시턴스 편차", dDevStopC, dDevUpC, dDevMoveC, _arrayXLabelValue, 0, "", "");

                // Q 저항 Chart 그리기  
                SetChart_Depict(chartAverageValueQ, "Q-Factor 편차", dDevStopQ, dDevUpQ, dDevMoveQ, _arrayXLabelValue, 0, "", "");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
        }

        private void SetChart_Depict(System.Windows.Forms.DataVisualization.Charting.Chart cht, string chartName
            , double[] dStopValue, double[] dUpValue, double[] dMoveValue, string[] arrayXLabelValue
            , int chartOption, string coilName, string chartType)
        {
            sDC = cht.Series.Add("정지");
            sDC = cht.Series.Add("올림");
            sDC = cht.Series.Add("이동");

            sDC.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
            sDC.XValueType = System.Windows.Forms.DataVisualization.Charting.ChartValueType.Int32;
            sDC.YValueType = System.Windows.Forms.DataVisualization.Charting.ChartValueType.Double;

            byte chartAreaNo;
            chartAreaNo = 0; //차트 영역 번호.
            m_Chart.InitialChart(ref cht, chartName, 1);

            m_Chart.MakeLegend(ref cht, "정지", "right", chartAreaNo);
            m_Chart.AddChartColumn(ref cht, chartAreaNo, chartAreaNo, "정지", ref arrayXLabelValue, ref dStopValue, true, System.Drawing.Color.Lime, 0, true);

            m_Chart.MakeLegend(ref cht, "올림", "right", chartAreaNo);
            m_Chart.AddChartColumn(ref cht, chartAreaNo, chartAreaNo, "올림", ref arrayXLabelValue, ref dUpValue, true, System.Drawing.Color.Blue, 0, true);

            m_Chart.MakeLegend(ref cht, "이동", "right", chartAreaNo);
            m_Chart.AddChartColumn(ref cht, chartAreaNo, chartAreaNo, "이동", ref arrayXLabelValue, ref dMoveValue, true, System.Drawing.Color.Red, 0, true);

            if (arrayXLabelValue.Length < 30)
                cht.ChartAreas[0].AxisX.Interval = 1; // 차수별일 경우는 X 축 1씩 증가 값 설정
            else
                cht.ChartAreas[0].AxisX.Interval = 10; // 제어봉일 경우는 X 축 10씩 증가 값 설정
        }

		private bool ExcelReportExport(string _strHogi, string _strOhDegree, string _strMeasurementType, string _strCoilName, string _strComparisonTarget)
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
			string strTitle = string.Format("RCS 이력 보고서");
			string strStartPath = Application.StartupPath;
            			
			try
			{
                string fileName = strStartPath + @"\코일보고서Form_Rv1.xlsx";

				DateTime  nowDataTime = System.DateTime.Now;

				if (_strCoilName.Trim() == "" && _strCoilName.Trim() == "전체")
					newfileName = strStartPath + @"\Report\" + "RCS 보고서_" + _strHogi + "_" + _strOhDegree + "_" + _strMeasurementType + "_(" + nowDataTime.ToString("yyyyMMddHHmmss") + ").xlsx";
				else if (_strMeasurementType.Trim() == "" && _strMeasurementType.Trim() == "전체")
					newfileName = strStartPath + @"\Report\" + "RCS 보고서_" + _strHogi + "_" + _strOhDegree + "_" + _strCoilName + "_(" + nowDataTime.ToString("yyyyMMddHHmmss") + ").xlsx";
				else
					newfileName = strStartPath + @"\Report\" + "RCS 보고서_" + _strHogi + "_" + _strOhDegree + "_" + _strMeasurementType + "_" + _strCoilName + "_(" + nowDataTime.ToString("yyyyMMddHHmmss") + ").xlsx";
				
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

				string strPowerCabinet = "", strTemperature_ReferenceValue = "", strFrequency = "", strVoltageLevel = "", strMeasurementDate = "";

                Queue<string> arrayStrPowerCabinet = new Queue<string>();
                if (chkPowerCabinet52_0.Checked) arrayStrPowerCabinet.Enqueue(chkPowerCabinet52_0.Text);
                if (chkPowerCabinet52_1.Checked) arrayStrPowerCabinet.Enqueue(chkPowerCabinet52_1.Text);
                if (chkPowerCabinet52_2.Checked) arrayStrPowerCabinet.Enqueue(chkPowerCabinet52_2.Text);
                if (chkPowerCabinet52_3.Checked) arrayStrPowerCabinet.Enqueue(chkPowerCabinet52_3.Text);
                if (chkPowerCabinet52_4.Checked) arrayStrPowerCabinet.Enqueue(chkPowerCabinet52_4.Text);
                if (chkPowerCabinet33_0.Checked) arrayStrPowerCabinet.Enqueue(chkPowerCabinet33_0.Text);
                if (chkPowerCabinet33_1.Checked) arrayStrPowerCabinet.Enqueue(chkPowerCabinet33_1.Text);
                if (chkPowerCabinet33_2.Checked) arrayStrPowerCabinet.Enqueue(chkPowerCabinet33_2.Text);
                int n = arrayStrPowerCabinet.Count;
                for( int i = 0; i < n; i++ )
                {
                    if (i == 0)
                    {
                        strPowerCabinet = $"'{arrayStrPowerCabinet.Dequeue()}'";
                    }
                    else
                    {
                        strPowerCabinet = $"{strPowerCabinet},'{arrayStrPowerCabinet.Dequeue()}'";
                    }
                }

                DataTable dt = new DataTable();
				dt = m_db.GetRCSDiagnosisHeaderDataGridViewDataInfo(strPlantName, _strHogi, _strOhDegree, strPowerCabinet);

				if (dt != null && dt.Rows.Count > 0)
				{
					strTemperature_ReferenceValue = dt.Rows[0]["Temperature_ReferenceValue"].ToString().Trim();
					strFrequency = dt.Rows[0]["Frequency"].ToString().Trim();
					strVoltageLevel = dt.Rows[0]["VoltageLevel"].ToString().Trim();
					strMeasurementDate = dt.Rows[0]["MeasurementDate"].ToString().Trim();
				}

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
				oRng.Value = _strMeasurementType.Trim();

				// 코일
				oRng = xlsxSheet1.get_Range("I3", "J3"); //해당 범위의 셀 획득
				oRng.MergeCells = true; //머지
				oRng.Value = _strCoilName.Trim();

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
                oRng.Value = "비고";

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

				int intLeft = 2, intTop = 1, intWidth = 915, intHeight = 296;

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
        private void  GetItemTitleName(ref string strItemRefDC, ref string strItemAveDC, ref string strItemRefAC, ref string strItemAveAC
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

        private void button2_Click(object sender, EventArgs e)
        {
            if (gbPowerCabinet33_0.Enabled)
                chkPowerCabinet33_0.Checked = true;

            if (gbPowerCabinet33_1.Enabled)
                chkPowerCabinet33_1.Checked = true;

            if (gbPowerCabinet33_2.Enabled)
                chkPowerCabinet33_2.Checked = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (gbPowerCabinet33_0.Enabled)
                chkPowerCabinet33_0.Checked = false;

            if (gbPowerCabinet33_1.Enabled)
                chkPowerCabinet33_1.Checked = false;

            if (gbPowerCabinet33_2.Enabled)
                chkPowerCabinet33_2.Checked = false;
        }
    }
}
