using Coil_Diagnostor.Function;
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

namespace Coil_Diagnostor
{
    public partial class frmRCSHistory : Form
    {
        protected Function.FunctionMeasureProcess m_MeasureProcess = new Function.FunctionMeasureProcess();
        protected Function.FunctionDataControl m_db = new Function.FunctionDataControl();
        protected Function.FunctionChart m_Chart = new Function.FunctionChart();
        protected frmMessageBox frmMB = new frmMessageBox();
        protected string strPlantName = "";
        protected bool boolFormLoad = false;
        protected string strReferenceHogi = "초기값";
        protected string strReferenceOHDegree = "제 0 차";

        System.Windows.Forms.DataVisualization.Charting.Series sSeries = new System.Windows.Forms.DataVisualization.Charting.Series();

        public frmRCSHistory()
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
        private void frmRCSHistory_Load(object sender, EventArgs e)
        {
            strPlantName = Gini.GetValue("Device", "PlantName").Trim();

            // ComboBox 데이터 설정
            SetComboBoxDataBinding();

            // 그리드 초기 설정
            SetDataGridViewInitialize();

            // 정지,올림,이동 label 설정
            SetReferenceValuetAndAverageValueLabelLocation(0);
            SetReferenceValuetAndAverageValueLabelLocation(1);

            // 정지,올림,이동 label 기준값 Visible 설정
            SetReferenceValueLabelVisible(false);

            // 정지,올림,이동 label 평균값 Visible 설정
            SetAverageValueLabelVisible(false);

            tabControl.SelectedIndex = 0;
            if (tabControl.SelectedIndex == 1 || tabControl.SelectedIndex == 2)
            {
                btnReferenceValue.Enabled = true;
                btnAverageValue.Enabled = true;
            }
            else
            {
                btnReferenceValue.Enabled = false;
                btnAverageValue.Enabled = false;
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

            // 측정타입
            string[] strTypeItem = Gini.GetValue("Combo", "RCSType_Item").Split(',');

            cboMeasurementType.Items.Clear();

            for (int i = 0; i < strTypeItem.Length; i++)
            {
                cboMeasurementType.Items.Add(strTypeItem[i].Trim());
            }

            cboMeasurementType.SelectedIndex = 1;


            // 코일명
            string[] strCoilNameItem = Gini.GetValue("Combo", "RCSCoilName_Item").Split(',');

            cboCoilName.Items.Clear();

            cboCoilName.Items.Add("전체");

            for (int i = 0; i < strCoilNameItem.Length; i++)
            {
                cboCoilName.Items.Add(strCoilNameItem[i].Trim());
            }

            cboCoilName.SelectedIndex = 0;


            // 비교대상 
            string[] strComparisonTargetItem = Gini.GetValue("Combo", "ComparisonTarget_Item").Split(',');

            cboComparisonTarget.Items.Clear();

            for (int i = 0; i < strComparisonTargetItem.Length; i++)
            {
                cboComparisonTarget.Items.Add(strComparisonTargetItem[i].Trim());
            }

            cboComparisonTarget.SelectedIndex = 0;

            // 그룹
            string[] strMeasurementGroup_Item = Gini.GetValue("RCS", "RCSMeasurementGroup_Item").Split(',');    // todo : check

            cboPowerCabinet.Items.Clear();

            cboPowerCabinet.Items.Add("전체");

            for (int i = 0; i < strMeasurementGroup_Item.Length; i++)
            {
                cboPowerCabinet.Items.Add(strMeasurementGroup_Item[i].Trim());
            }

            cboPowerCabinet.SelectedIndex = 0;
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
            }
            else
            {
                btnReferenceValue.Enabled = false;
                btnAverageValue.Enabled = false;
            }

            // Rdc, Rac 활성화 및 비활성화
            SetTabControlDCAC();
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
                string strRCSPowerCabinetItem = "";

                string[] strMeasurementGroup_Item = Gini.GetValue("RCS", "RCSMeasurementGroup_Item").Split(',');
                foreach (string cabinet in strMeasurementGroup_Item)
                {
                    if (strRCSPowerCabinetItem.Length == 0)
                    {
                        strRCSPowerCabinetItem = Gini.GetValue("RCS", $"RCSPowerCabinetItem_{cabinet}").Trim();
                    }
                    else
                    {
                        strRCSPowerCabinetItem = strRCSPowerCabinetItem + "," + Gini.GetValue("RCS", $"RCSPowerCabinetItem_{cabinet}").Trim();
                    }
                }

                string[] arrayRCSPowerCabinetItem = strRCSPowerCabinetItem.Split(',');

                if (cboHistorySearchType.SelectedItem.ToString().Trim() == "기간별")
                {
                    // 제어봉
                    cboControlRod.Items.Clear();

                    for (int i = 0; i < arrayRCSPowerCabinetItem.Length; i++)
                    {
                        if (arrayRCSPowerCabinetItem[i].Trim() != "N/A")
                            cboControlRod.Items.Add(arrayRCSPowerCabinetItem[i].Trim());
                    }

                    cboControlRod.SelectedIndex = 0;
                }
                else
                {
                    // 제어봉
                    cboControlRod.Items.Clear();

                    cboControlRod.Items.Add("전체");

                    for (int i = 0; i < arrayRCSPowerCabinetItem.Length; i++)
                    {
                        if (arrayRCSPowerCabinetItem[i].Trim() != "N/A")
                            cboControlRod.Items.Add(arrayRCSPowerCabinetItem[i].Trim());
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

            string strMeasurementType = cboMeasurementType.SelectedItem == null ? "" : cboMeasurementType.SelectedItem.ToString().Trim();

            if (strMeasurementType.Trim() == "측정치")
                btnScalingSetting.Enabled = true;
            else
                btnScalingSetting.Enabled = false;
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
        private void cboPowerCabinet_SelectedIndexChanged(object sender, EventArgs e)
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

                // 정지,올림,이동 label 설정
                SetReferenceValuetAndAverageValueLabelLocation(0);
                SetReferenceValuetAndAverageValueLabelLocation(1);
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

                // 정지,올림,이동 label 설정
                SetReferenceValuetAndAverageValueLabelLocation(0);
                SetReferenceValuetAndAverageValueLabelLocation(1);
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
            lblTab1StopLed_ReferenceValue.Visible = _boolVisible;
            lblTab1StopText_ReferenceValue.Visible = _boolVisible;
            lblTab1UpLed_ReferenceValue.Visible = _boolVisible;
            lblTab1UpText_ReferenceValue.Visible = _boolVisible;
            lblTab1MoveLed_ReferenceValue.Visible = _boolVisible;
            lblTab1MoveText_ReferenceValue.Visible = _boolVisible;

            lblTab2StopLed_ReferenceValue.Visible = _boolVisible;
            lblTab2topText_ReferenceValue.Visible = _boolVisible;
            lblTab2UpLed_ReferenceValue.Visible = _boolVisible;
            lblTab2UpText_ReferenceValue.Visible = _boolVisible;
            lblTab2MoveLed_ReferenceValue.Visible = _boolVisible;
            lblTab2MoveText_ReferenceValue.Visible = _boolVisible;

            if (chartMeasurementValueDC.Series.Count > 3)
                chartMeasurementValueDC.Series[3].Enabled = _boolVisible;

            if (chartMeasurementValueDC.Series.Count > 4)
                chartMeasurementValueDC.Series[4].Enabled = _boolVisible;

            if (chartMeasurementValueDC.Series.Count > 5)
                chartMeasurementValueDC.Series[5].Enabled = _boolVisible;

            if (chartMeasurementValueAC.Series.Count > 3)
                chartMeasurementValueAC.Series[3].Enabled = _boolVisible;

            if (chartMeasurementValueAC.Series.Count > 4)
                chartMeasurementValueAC.Series[4].Enabled = _boolVisible;

            if (chartMeasurementValueAC.Series.Count > 5)
                chartMeasurementValueAC.Series[5].Enabled = _boolVisible;

            if (chartMeasurementValueL.Series.Count > 3)
                chartMeasurementValueL.Series[3].Enabled = _boolVisible;

            if (chartMeasurementValueL.Series.Count > 4)
                chartMeasurementValueL.Series[4].Enabled = _boolVisible;

            if (chartMeasurementValueL.Series.Count > 5)
                chartMeasurementValueL.Series[5].Enabled = _boolVisible;

            if (chartMeasurementValueC.Series.Count > 3)
                chartMeasurementValueC.Series[3].Enabled = _boolVisible;

            if (chartMeasurementValueC.Series.Count > 4)
                chartMeasurementValueC.Series[4].Enabled = _boolVisible;

            if (chartMeasurementValueC.Series.Count > 5)
                chartMeasurementValueC.Series[5].Enabled = _boolVisible;

            if (chartMeasurementValueQ.Series.Count > 3)
                chartMeasurementValueQ.Series[3].Enabled = _boolVisible;

            if (chartMeasurementValueQ.Series.Count > 4)
                chartMeasurementValueQ.Series[4].Enabled = _boolVisible;

            if (chartMeasurementValueQ.Series.Count > 5)
                chartMeasurementValueQ.Series[5].Enabled = _boolVisible;
        }

        /// <summary>
        /// 정지,올림,이동 label 평균값 Visible 설정
        /// </summary>
        private void SetAverageValueLabelVisible(bool _boolVisible)
        {
            lblTab1StopLed_AverageValue.Visible = _boolVisible;
            lblTab1StopText_AverageValue.Visible = _boolVisible;
            lblTab1UpLed_AverageValue.Visible = _boolVisible;
            lblTab1UpText_AverageValue.Visible = _boolVisible;
            lblTab1MoveLed_AverageValue.Visible = _boolVisible;
            lblTab1MoveText_AverageValue.Visible = _boolVisible;

            lblTab2StopLed_AverageValue.Visible = _boolVisible;
            lblTab2StopText_AverageValue.Visible = _boolVisible;
            lblTab2UpLed_AverageValue.Visible = _boolVisible;
            lblTab2UpText_AverageValue.Visible = _boolVisible;
            lblTab2MoveLed_AverageValue.Visible = _boolVisible;
            lblTab2MoveText_AverageValue.Visible = _boolVisible;

            if (chartMeasurementValueDC.Series.Count > 6)
                chartMeasurementValueDC.Series[6].Enabled = _boolVisible;

            if (chartMeasurementValueDC.Series.Count > 7)
                chartMeasurementValueDC.Series[7].Enabled = _boolVisible;

            if (chartMeasurementValueDC.Series.Count > 8)
                chartMeasurementValueDC.Series[8].Enabled = _boolVisible;

            if (chartMeasurementValueAC.Series.Count > 6)
                chartMeasurementValueAC.Series[6].Enabled = _boolVisible;

            if (chartMeasurementValueAC.Series.Count > 7)
                chartMeasurementValueAC.Series[7].Enabled = _boolVisible;

            if (chartMeasurementValueAC.Series.Count > 8)
                chartMeasurementValueAC.Series[8].Enabled = _boolVisible;

            if (chartMeasurementValueL.Series.Count > 6)
                chartMeasurementValueL.Series[6].Enabled = _boolVisible;

            if (chartMeasurementValueL.Series.Count > 7)
                chartMeasurementValueL.Series[7].Enabled = _boolVisible;

            if (chartMeasurementValueL.Series.Count > 8)
                chartMeasurementValueL.Series[8].Enabled = _boolVisible;

            if (chartMeasurementValueC.Series.Count > 6)
                chartMeasurementValueC.Series[6].Enabled = _boolVisible;

            if (chartMeasurementValueC.Series.Count > 7)
                chartMeasurementValueC.Series[7].Enabled = _boolVisible;

            if (chartMeasurementValueC.Series.Count > 8)
                chartMeasurementValueC.Series[8].Enabled = _boolVisible;

            if (chartMeasurementValueQ.Series.Count > 6)
                chartMeasurementValueQ.Series[6].Enabled = _boolVisible;

            if (chartMeasurementValueQ.Series.Count > 7)
                chartMeasurementValueQ.Series[7].Enabled = _boolVisible;

            if (chartMeasurementValueQ.Series.Count > 8)
                chartMeasurementValueQ.Series[8].Enabled = _boolVisible;
        }

        /// <summary>
        /// 정지,올림,이동 label 설정
        /// </summary>
        private void SetReferenceValuetAndAverageValueLabelLocation(int _intSelect)
        {
            switch (_intSelect)
            {
                case 0:
                    if (btnReferenceValue.Text.Trim() == "기준값 표시 (F3)")
                    {
                        if (btnAverageValue.Text.Trim() == "평균값 숨김 (F4)")
                        {
                            lblTab1StopLed.Location = new System.Drawing.Point(10, 75);
                            lblTab1StopText.Location = new System.Drawing.Point(44, 75);
                            lblTab1UpLed.Location = new System.Drawing.Point(10, 108);
                            lblTab1UpText.Location = new System.Drawing.Point(44, 108);
                            lblTab1MoveLed.Location = new System.Drawing.Point(10, 141);
                            lblTab1MoveText.Location = new System.Drawing.Point(44, 141);

                            lblTab1StopLed_AverageValue.Location = new System.Drawing.Point(10, 174);
                            lblTab1StopText_AverageValue.Location = new System.Drawing.Point(44, 174);
                            lblTab1UpLed_AverageValue.Location = new System.Drawing.Point(10, 207);
                            lblTab1UpText_AverageValue.Location = new System.Drawing.Point(44, 207);
                            lblTab1MoveLed_AverageValue.Location = new System.Drawing.Point(10, 240);
                            lblTab1MoveText_AverageValue.Location = new System.Drawing.Point(44, 240);
                        }
                        else
                        {
                            lblTab1StopLed.Location = new System.Drawing.Point(10, 75);
                            lblTab1StopText.Location = new System.Drawing.Point(44, 75);
                            lblTab1UpLed.Location = new System.Drawing.Point(10, 225);
                            lblTab1UpText.Location = new System.Drawing.Point(44, 225);
                            lblTab1MoveLed.Location = new System.Drawing.Point(10, 375);
                            lblTab1MoveText.Location = new System.Drawing.Point(44, 375);
                        }
                    }
                    else
                    {
                        lblTab1StopLed.Location = new System.Drawing.Point(10, 75);
                        lblTab1StopText.Location = new System.Drawing.Point(44, 75);
                        lblTab1UpLed.Location = new System.Drawing.Point(10, 108);
                        lblTab1UpText.Location = new System.Drawing.Point(44, 108);
                        lblTab1MoveLed.Location = new System.Drawing.Point(10, 141);
                        lblTab1MoveText.Location = new System.Drawing.Point(44, 141);

                        lblTab1StopLed_ReferenceValue.Location = new System.Drawing.Point(10, 174);
                        lblTab1StopText_ReferenceValue.Location = new System.Drawing.Point(44, 174);
                        lblTab1UpLed_ReferenceValue.Location = new System.Drawing.Point(10, 207);
                        lblTab1UpText_ReferenceValue.Location = new System.Drawing.Point(44, 207);
                        lblTab1MoveLed_ReferenceValue.Location = new System.Drawing.Point(10, 240);
                        lblTab1MoveText_ReferenceValue.Location = new System.Drawing.Point(44, 240);

                        if (btnAverageValue.Text.Trim() == "평균값 숨김 (F4)")
                        {
                            lblTab1StopLed_AverageValue.Location = new System.Drawing.Point(10, 273);
                            lblTab1StopText_AverageValue.Location = new System.Drawing.Point(44, 273);
                            lblTab1UpLed_AverageValue.Location = new System.Drawing.Point(10, 306);
                            lblTab1UpText_AverageValue.Location = new System.Drawing.Point(44, 306);
                            lblTab1MoveLed_AverageValue.Location = new System.Drawing.Point(10, 339);
                            lblTab1MoveText_AverageValue.Location = new System.Drawing.Point(44, 339);
                        }
                    }
                    break;
                case 1:
                    if (btnReferenceValue.Text.Trim() == "기준값 표시 (F3)")
                    {
                        if (btnAverageValue.Text.Trim() == "평균값 숨김 (F4)")
                        {
                            lblTab2StopLed.Location = new System.Drawing.Point(3, 8);
                            lblTab2StopText.Location = new System.Drawing.Point(29, 8);
                            lblTab2UpLed.Location = new System.Drawing.Point(3, 32);
                            lblTab2UpText.Location = new System.Drawing.Point(29, 32);
                            lblTab2MoveLed.Location = new System.Drawing.Point(3, 56);
                            lblTab2MoveText.Location = new System.Drawing.Point(29, 56);

                            lblTab2StopLed_AverageValue.Location = new System.Drawing.Point(3, 80);
                            lblTab2StopText_AverageValue.Location = new System.Drawing.Point(29, 80);
                            lblTab2UpLed_AverageValue.Location = new System.Drawing.Point(3, 104);
                            lblTab2UpText_AverageValue.Location = new System.Drawing.Point(29, 104);
                            lblTab2MoveLed_AverageValue.Location = new System.Drawing.Point(3, 128);
                            lblTab2MoveText_AverageValue.Location = new System.Drawing.Point(29, 128);
                        }
                        else
                        {
                            lblTab2StopLed.Location = new System.Drawing.Point(3, 58);
                            lblTab2StopText.Location = new System.Drawing.Point(29, 58);
                            lblTab2UpLed.Location = new System.Drawing.Point(3, 108);
                            lblTab2UpText.Location = new System.Drawing.Point(29, 108);
                            lblTab2MoveLed.Location = new System.Drawing.Point(3, 158);
                            lblTab2MoveText.Location = new System.Drawing.Point(29, 158);
                        }
                    }
                    else
                    {
                        lblTab2StopLed.Location = new System.Drawing.Point(3, 8);
                        lblTab2StopText.Location = new System.Drawing.Point(29, 8);
                        lblTab2UpLed.Location = new System.Drawing.Point(3, 32);
                        lblTab2UpText.Location = new System.Drawing.Point(29, 32);
                        lblTab2MoveLed.Location = new System.Drawing.Point(3, 56);
                        lblTab2MoveText.Location = new System.Drawing.Point(29, 56);

                        lblTab2StopLed_ReferenceValue.Location = new System.Drawing.Point(3, 80);
                        lblTab2topText_ReferenceValue.Location = new System.Drawing.Point(29, 80);
                        lblTab2UpLed_ReferenceValue.Location = new System.Drawing.Point(3, 104);
                        lblTab2UpText_ReferenceValue.Location = new System.Drawing.Point(29, 104);
                        lblTab2MoveLed_ReferenceValue.Location = new System.Drawing.Point(3, 128);
                        lblTab2MoveText_ReferenceValue.Location = new System.Drawing.Point(29, 128);

                        if (btnAverageValue.Text.Trim() == "평균값 숨김 (F4)")
                        {
                            lblTab2StopLed_AverageValue.Location = new System.Drawing.Point(3, 152);
                            lblTab2StopText_AverageValue.Location = new System.Drawing.Point(29, 152);
                            lblTab2UpLed_AverageValue.Location = new System.Drawing.Point(3, 176);
                            lblTab2UpText_AverageValue.Location = new System.Drawing.Point(29, 176);
                            lblTab2MoveLed_AverageValue.Location = new System.Drawing.Point(3, 200);
                            lblTab2MoveText_AverageValue.Location = new System.Drawing.Point(29, 200);
                        }
                    }
                    break;
            }
        }

        /// <summary>
        /// 기준 값 Setting Button Event
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

                string[] arrayData = new string[15];

                arrayData[0] = teRdcStop_RCSAverageValue.Text.Trim();
                arrayData[1] = teRdcUp_RCSAverageValue.Text.Trim();
                arrayData[2] = teRdcMove_RCSAverageValue.Text.Trim();
                arrayData[3] = teRacStop_RCSAverageValue.Text.Trim();
                arrayData[4] = teRacUp_RCSAverageValue.Text.Trim();
                arrayData[5] = teRacMove_RCSAverageValue.Text.Trim();
                arrayData[6] = teLStop_RCSAverageValue.Text.Trim();
                arrayData[7] = teLUp_RCSAverageValue.Text.Trim();
                arrayData[8] = teLMove_RCSAverageValue.Text.Trim();
                arrayData[9] = teCStop_RCSAverageValue.Text.Trim();
                arrayData[10] = teCUp_RCSAverageValue.Text.Trim();
                arrayData[11] = teCMove_RCSAverageValue.Text.Trim();
                arrayData[12] = teQStop_RCSAverageValue.Text.Trim();
                arrayData[13] = teQUp_RCSAverageValue.Text.Trim();
                arrayData[14] = teQMove_RCSAverageValue.Text.Trim();

                if ((m_db.GetRCSReferenceValueDataCount(strPlantName.Trim(), strHogi.Trim(), strOhDegree.Trim())) > 0)
                {
                    boolSave = m_db.UpDateRCSReferenceValueDataInfo(strPlantName.Trim(), strHogi.Trim(), strOhDegree.Trim(), arrayData);
                }
                else
                {
                    boolSave = m_db.InsertRCSReferenceValueDataInfo(strPlantName.Trim(), strHogi.Trim(), strOhDegree.Trim(), arrayData);
                }

                if (boolSave)
                {
                    // 기준 값(RCS) 초기화
                    SetRCSReferenceValueInitialize();

                    // 기준값 가져오기
                    GetRCSReferenceValue();

                    // 표 형식 결과 재판단 및 업데이터
                    SetDataGridViewUpdate();

                    frmMB.lblMessage.Text = "기준 값 설정이 완료";
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
                decimal dStop_Rdc = 0M, dStop_Rac = 0M, dStop_L = 0M, dStop_C = 0M, dStop_Q = 0M, dUp_Rdc = 0M, dUp_Rac = 0M
                    , dUp_L = 0M, dUp_C = 0M, dUp_Q = 0M, dMove_Rdc = 0M, dMove_Rac = 0M, dMove_L = 0M, dMove_C = 0M, dMove_Q = 0M;

                if (cboComparisonTarget.SelectedItem.ToString().Trim() == "평균값")
                {
                    dStop_Rdc = teRdcStop_RCSAverageValue.Text == "" ? 0.000M : Convert.ToDecimal(teRdcStop_RCSAverageValue.Text.Trim());
                    dStop_Rac = teRacStop_RCSAverageValue.Text == "" ? 0.000M : Convert.ToDecimal(teRacStop_RCSAverageValue.Text.Trim());
                    dStop_L = teLStop_RCSAverageValue.Text == "" ? 0.000M : Convert.ToDecimal(teLStop_RCSAverageValue.Text.Trim());
                    dStop_C = teCStop_RCSAverageValue.Text == "" ? 0.000M : Convert.ToDecimal(teCStop_RCSAverageValue.Text.Trim());
                    dStop_Q = teQStop_RCSAverageValue.Text == "" ? 0.000M : Convert.ToDecimal(teQStop_RCSAverageValue.Text.Trim());
                    dUp_Rdc = teRdcUp_RCSAverageValue.Text == "" ? 0.000M : Convert.ToDecimal(teRdcUp_RCSAverageValue.Text.Trim());
                    dUp_Rac = teRacUp_RCSAverageValue.Text == "" ? 0.000M : Convert.ToDecimal(teRacUp_RCSAverageValue.Text.Trim());
                    dUp_L = teLUp_RCSAverageValue.Text == "" ? 0.000M : Convert.ToDecimal(teLUp_RCSAverageValue.Text.Trim());
                    dUp_C = teCUp_RCSAverageValue.Text == "" ? 0.000M : Convert.ToDecimal(teCUp_RCSAverageValue.Text.Trim());
                    dUp_Q = teQUp_RCSAverageValue.Text == "" ? 0.000M : Convert.ToDecimal(teQUp_RCSAverageValue.Text.Trim());
                    dMove_Rdc = teRdcMove_RCSAverageValue.Text == "" ? 0.000M : Convert.ToDecimal(teRdcMove_RCSAverageValue.Text.Trim());
                    dMove_Rac = teRacMove_RCSAverageValue.Text == "" ? 0.000M : Convert.ToDecimal(teRacMove_RCSAverageValue.Text.Trim());
                    dMove_L = teLMove_RCSAverageValue.Text == "" ? 0.000M : Convert.ToDecimal(teLMove_RCSAverageValue.Text.Trim());
                    dMove_C = teCMove_RCSAverageValue.Text == "" ? 0.000M : Convert.ToDecimal(teCMove_RCSAverageValue.Text.Trim());
                    dMove_Q = teQMove_RCSAverageValue.Text == "" ? 0.000M : Convert.ToDecimal(teQMove_RCSAverageValue.Text.Trim());

                }
                else
                {
                    dStop_Rdc = teRdcStop_RCSReferenceValue.Text == "" ? 0.000M : Convert.ToDecimal(teRdcStop_RCSReferenceValue.Text.Trim());
                    dStop_Rac = teRacStop_RCSReferenceValue.Text == "" ? 0.000M : Convert.ToDecimal(teRacStop_RCSReferenceValue.Text.Trim());
                    dStop_L = teLStop_RCSReferenceValue.Text == "" ? 0.000M : Convert.ToDecimal(teLStop_RCSReferenceValue.Text.Trim());
                    dStop_C = teCStop_RCSReferenceValue.Text == "" ? 0.000M : Convert.ToDecimal(teCStop_RCSReferenceValue.Text.Trim());
                    dStop_Q = teQStop_RCSReferenceValue.Text == "" ? 0.000M : Convert.ToDecimal(teQStop_RCSReferenceValue.Text.Trim());
                    dUp_Rdc = teRdcUp_RCSReferenceValue.Text == "" ? 0.000M : Convert.ToDecimal(teRdcUp_RCSReferenceValue.Text.Trim());
                    dUp_Rac = teRacUp_RCSReferenceValue.Text == "" ? 0.000M : Convert.ToDecimal(teRacUp_RCSReferenceValue.Text.Trim());
                    dUp_L = teLUp_RCSReferenceValue.Text == "" ? 0.000M : Convert.ToDecimal(teLUp_RCSReferenceValue.Text.Trim());
                    dUp_C = teCUp_RCSReferenceValue.Text == "" ? 0.000M : Convert.ToDecimal(teCUp_RCSReferenceValue.Text.Trim());
                    dUp_Q = teQUp_RCSReferenceValue.Text == "" ? 0.000M : Convert.ToDecimal(teQUp_RCSReferenceValue.Text.Trim());
                    dMove_Rdc = teRdcMove_RCSReferenceValue.Text == "" ? 0.000M : Convert.ToDecimal(teRdcMove_RCSReferenceValue.Text.Trim());
                    dMove_Rac = teRacMove_RCSReferenceValue.Text == "" ? 0.000M : Convert.ToDecimal(teRacMove_RCSReferenceValue.Text.Trim());
                    dMove_L = teLMove_RCSReferenceValue.Text == "" ? 0.000M : Convert.ToDecimal(teLMove_RCSReferenceValue.Text.Trim());
                    dMove_C = teCMove_RCSReferenceValue.Text == "" ? 0.000M : Convert.ToDecimal(teCMove_RCSReferenceValue.Text.Trim());
                    dMove_Q = teQMove_RCSReferenceValue.Text == "" ? 0.000M : Convert.ToDecimal(teQMove_RCSReferenceValue.Text.Trim());

                }

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
                    string strCoilName = dgvMeasurement.Rows[i].Cells["CoilName"].Value == null ? "" : dgvMeasurement.Rows[i].Cells["CoilName"].Value.ToString().Trim();

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

                    if (strCoilName.Trim() == "정지")
                    {
                        dgvMeasurement.Rows[i].Cells["DC_Deviation"].Value = (dStop_Rdc - dRdc).ToString("F3");
                        dgvMeasurement.Rows[i].Cells["AC_Deviation"].Value = (dStop_Rac - dRac).ToString("F3");
                        dgvMeasurement.Rows[i].Cells["L_Deviation"].Value = (dStop_L - dL).ToString("F6");
                        dgvMeasurement.Rows[i].Cells["C_Deviation"].Value = (dStop_C - dC).ToString("F3");

                        dcmRefRdcValue = dStop_Rdc;
                        dcmRefRacValue = dStop_Rac;
                        dcmRefLValue = dStop_L;
                        dcmRefCValue = dStop_C;
                        dcmRefQValue = dStop_Q;
                    }
                    else if (strCoilName.Trim() == "올림")
                    {
                        dgvMeasurement.Rows[i].Cells["DC_Deviation"].Value = (dUp_Rdc - dRdc).ToString("F3");
                        dgvMeasurement.Rows[i].Cells["AC_Deviation"].Value = (dUp_Rac - dRac).ToString("F3");
                        dgvMeasurement.Rows[i].Cells["L_Deviation"].Value = (dUp_L - dL).ToString("F3");
                        dgvMeasurement.Rows[i].Cells["C_Deviation"].Value = (dUp_C - dC).ToString("F6");
                        dgvMeasurement.Rows[i].Cells["Q_Deviation"].Value = (dUp_Q - dQ).ToString("F3");

                        dcmRefRdcValue = dUp_Rdc;
                        dcmRefRacValue = dUp_Rac;
                        dcmRefLValue = dUp_L;
                        dcmRefCValue = dUp_C;
                        dcmRefQValue = dUp_Q;
                    }
                    else if (strCoilName.Trim() == "이동")
                    {
                        dgvMeasurement.Rows[i].Cells["DC_Deviation"].Value = (dMove_Rdc - dRdc).ToString("F3");
                        dgvMeasurement.Rows[i].Cells["AC_Deviation"].Value = (dMove_Rac - dRac).ToString("F3");
                        dgvMeasurement.Rows[i].Cells["L_Deviation"].Value = (dMove_L - dL).ToString("F3");
                        dgvMeasurement.Rows[i].Cells["C_Deviation"].Value = (dMove_C - dC).ToString("F6");
                        dgvMeasurement.Rows[i].Cells["Q_Deviation"].Value = (dMove_Q - dQ).ToString("F3");

                        dcmRefRdcValue = dMove_Rdc;
                        dcmRefRacValue = dMove_Rac;
                        dcmRefLValue = dMove_L;
                        dcmRefCValue = dMove_C;
                        dcmRefQValue = dMove_Q;
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

                    string strHogi = dgvMeasurement.Rows[i].Cells["Hogi"].Value == null ? "" : dgvMeasurement.Rows[i].Cells["Hogi"].Value.ToString().Trim();
                    string strOhDegree = dgvMeasurement.Rows[i].Cells["Oh_Degree"].Value == null ? "" : dgvMeasurement.Rows[i].Cells["Oh_Degree"].Value.ToString().Trim();
                    string strPowerCabinet = dgvMeasurement.Rows[i].Cells["PowerCabinet"].Value == null ? "" : dgvMeasurement.Rows[i].Cells["PowerCabinet"].Value.ToString().Trim();
                    string strControlName = dgvMeasurement.Rows[i].Cells["ControlName"].Value == null ? "" : dgvMeasurement.Rows[i].Cells["ControlName"].Value.ToString().Trim();
                    int intCoilNumber = dgvMeasurement.Rows[i].Cells["CoilNumber"].Value == null || dgvMeasurement.Rows[i].Cells["CoilNumber"].Value.ToString().Trim() == ""
                        ? 0 : Convert.ToInt32(dgvMeasurement.Rows[i].Cells["CoilNumber"].Value.ToString().Trim());
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

                    if ((m_db.GetRCSDiagnosisDetailDataGridViewDataCount(strPlantName.Trim(), strHogi, strOhDegree, strPowerCabinet, strControlName, strCoilName, intCoilNumber)) > 0)
                    {
                        // 기존 데이터 Update
                        boolDetail = m_db.UpdateRCSDiagnosisDetailDataGridViewDataInfo(strPlantName.Trim(), strHogi, strOhDegree, strPowerCabinet
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
        /// 닫기 Button Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// 조회 Button Event
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

                // 기준 값(RCS) 초기화
                SetRCSReferenceValueInitialize();

                // 평균 값(RCS) 초기화
                SetRCSAverageValueInitialize();

                // 기준값 가져오기
                GetRCSReferenceValue();

                // Data 가져오기
                GetRCSMeasurementDataBinding();

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
        /// 기준 값(RCS) 초기화
        /// </summary>
        private void SetRCSReferenceValueInitialize()
        {
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

        /// <summary>
        /// 평균 값(RCS) 초기화
        /// </summary>
        private void SetRCSAverageValueInitialize()
        {
            teRdcStop_RCSAverageValue.Text = "0.000";
            teRdcUp_RCSAverageValue.Text = "0.000";
            teRdcMove_RCSAverageValue.Text = "0.000";
            teRacStop_RCSAverageValue.Text = "0.000";
            teRacUp_RCSAverageValue.Text = "0.000";
            teRacMove_RCSAverageValue.Text = "0.000";
            teLStop_RCSAverageValue.Text = "0.000";
            teLUp_RCSAverageValue.Text = "0.000";
            teLMove_RCSAverageValue.Text = "0.000";
            teCStop_RCSAverageValue.Text = "0.000000";
            teCUp_RCSAverageValue.Text = "0.000000";
            teCMove_RCSAverageValue.Text = "0.000000";
            teQStop_RCSAverageValue.Text = "0.000";
            teQUp_RCSAverageValue.Text = "0.000";
            teQMove_RCSAverageValue.Text = "0.000";
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
                        dt = m_db.GetRCSReferenceValueDataInfo(strPlantName.Trim(), strSelectHogi.Trim(), intMaxOhDegree.ToString().Trim());
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
        /// Data 가져오기
        /// </summary>
        private void GetRCSMeasurementDataBinding()
        {
            try
            {
                string strHogi = cboHogi.SelectedItem == null ? "" : cboHogi.SelectedItem.ToString().Trim();
                string strOhDegree = cboOhDegree.SelectedItem == null ? "" : cboOhDegree.SelectedItem.ToString().Trim();
                string strMeasurementType = cboMeasurementType.SelectedItem == null ? "" : cboMeasurementType.SelectedItem.ToString().Trim();
                string strFrequency = cboFrequency.SelectedItem == null ? "" : cboFrequency.SelectedItem.ToString().Trim();
                string strPowerCabinet = cboPowerCabinet.SelectedItem == null ? "" : cboPowerCabinet.SelectedItem.ToString().Trim();
                string strControlRod = cboControlRod.SelectedItem == null ? "" : cboControlRod.SelectedItem.ToString().Trim();
                string strCoilName = cboCoilName.SelectedItem == null ? "" : cboCoilName.SelectedItem.ToString().Trim();
                string strComparisonTarget = cboComparisonTarget.SelectedItem == null ? "" : cboComparisonTarget.SelectedItem.ToString().Trim();

                DataTable dtDB = new DataTable();
                DataTable dtAverage = new DataTable();

                // 측정 타입이 평균치일 경우
                if (cboMeasurementType.SelectedItem != null && cboMeasurementType.SelectedItem.ToString().Trim() == "평균치")
                {
                    dtDB = m_db.GetRCSDiagnosisDetailDataGridViewDataAverage(strPlantName.Trim(), strHogi, strOhDegree, strMeasurementType
                            , strFrequency, strPowerCabinet, strControlRod, strCoilName, strReferenceHogi, strReferenceOHDegree);

                    dtAverage = dtDB;
                }
                else
                {
                    dtDB = m_db.GetRCSDiagnosisDetailDataGridViewDataMeasure(strPlantName.Trim(), strHogi, strOhDegree, strMeasurementType
                            , strFrequency, strPowerCabinet, strControlRod, strCoilName, strReferenceHogi, strReferenceOHDegree);

                    dtAverage = m_db.GetRCSDiagnosisDetailDataGridViewDataAverage(strPlantName.Trim(), strHogi, strOhDegree, strMeasurementType
                            , strFrequency, strPowerCabinet, strControlRod, strCoilName, strReferenceHogi, strReferenceOHDegree);
                }

                if (dtDB == null || dtDB.Rows.Count <= 0)
                {
                    frmMB.lblMessage.Text = "데이터가 없습니다.";
                    frmMB.TopMost = true;
                    frmMB.ShowDialog();
                }
                else
                {
                    // 평균 계산
                    SetControlRodCoilGroupAverageCalculation(dtAverage);

                    // 표 형식
                    // 측정 타입이 평균치일 경우
                    if (cboMeasurementType.SelectedItem != null && cboMeasurementType.SelectedItem.ToString().Trim() == "평균치")
                    {
                        // 측정타입이 측정치일 경우
                        if (dtAverage != null && dtAverage.Rows.Count > 0)
                            SetRCSReferenceValueTableMark(ref dtAverage, strComparisonTarget);
                    }
                    else
                    {
                        // 측정타입이 측정치일 경우
                        if (dtDB != null && dtDB.Rows.Count > 0)
                            SetRCSReferenceValueTableMark(ref dtDB, strComparisonTarget);

                        SetDataAverage(ref dtAverage, dtDB);
                    }

                    if (dtAverage != null && dtAverage.Rows.Count > 0)
                    {
                        string strXLabelValue = "", strXLabelCheck = "";
                        int intXLabelValueCountStop = 0, intXLabelValueCountUp = 0, intXLabelValueCountMove = 0;

                        for (int i = 0; i < dtAverage.Rows.Count; i++)
                        {
                            if (strXLabelCheck.Trim() != dtAverage.Rows[i]["ControlName"].ToString().Trim())
                            {
                                strXLabelCheck = dtAverage.Rows[i]["ControlName"].ToString().Trim();

                                if (strXLabelValue.Trim() == "")
                                    strXLabelValue = dtAverage.Rows[i]["ControlName"].ToString().Trim();
                                else
                                    strXLabelValue = strXLabelValue + "," + dtAverage.Rows[i]["ControlName"].ToString().Trim();
                            }

                            if (dtAverage.Rows[i]["CoilName"].ToString().Trim() == "정지")
                                intXLabelValueCountStop++;

                            if (dtAverage.Rows[i]["CoilName"].ToString().Trim() == "올림")
                                intXLabelValueCountUp++;

                            if (dtAverage.Rows[i]["CoilName"].ToString().Trim() == "이동")
                                intXLabelValueCountMove++;
                        }

                        string[] arrayXLabelValue = strXLabelValue.Split(',');
                        int intXLabelValueCount = arrayXLabelValue.Length;

                        // 측정값 및 편차 그래프 그리기
                        SetRCSReferenceValueMeasurement(dtAverage, arrayXLabelValue, intXLabelValueCountStop, intXLabelValueCountUp, intXLabelValueCountMove);
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
            string strDB_PlantNeme = "", strDB_Hogi = "", strDB_Oh_Degree = "", strDB_PowerCabinet = "", strDB_ControlName = "", strDB_CoilName = "";
            int intCoilCount = 0;

            for (int i = 0; i < dtAverage.Rows.Count; i++)
            {
                if (strDB_PlantNeme.Trim() != dtAverage.Rows[i]["PlantName"].ToString().Trim()
                    || strDB_Hogi.Trim() != dtAverage.Rows[i]["Hogi"].ToString().Trim()
                    || strDB_Oh_Degree.Trim() != dtAverage.Rows[i]["Oh_Degree"].ToString().Trim()
                    || strDB_PowerCabinet.Trim() != dtAverage.Rows[i]["PowerCabinet"].ToString().Trim()
                    || strDB_ControlName.Trim() != dtAverage.Rows[i]["ControlName"].ToString().Trim()
                    || strDB_CoilName.Trim() != dtAverage.Rows[i]["CoilName"].ToString().Trim())
                {
                    strDB_PlantNeme = dtAverage.Rows[i]["PlantName"].ToString().Trim();
                    strDB_Hogi = dtAverage.Rows[i]["Hogi"].ToString().Trim();
                    strDB_Oh_Degree = dtAverage.Rows[i]["Oh_Degree"].ToString().Trim();
                    strDB_PowerCabinet = dtAverage.Rows[i]["PowerCabinet"].ToString().Trim();
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
                            && strDB_PowerCabinet.Trim() == dtDB.Rows[j]["PowerCabinet"].ToString().Trim()
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
        /// 평균 계산
        /// </summary>
        /// <param name="_dt"></param>
        private void SetControlRodCoilGroupAverageCalculation(DataTable _dt)
        {
            try
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
                            dcmSumRdcUp += _dt.Rows[i]["DC_ResistanceValue"] == null || _dt.Rows[i]["DC_ResistanceValue"].ToString().Trim() == ""
                                ? 0.000M : Convert.ToDecimal(_dt.Rows[i]["DC_ResistanceValue"].ToString().Trim());
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
                            dcmSumRdcMove += _dt.Rows[i]["DC_ResistanceValue"] == null || _dt.Rows[i]["DC_ResistanceValue"].ToString().Trim() == ""
                                ? 0.000M : Convert.ToDecimal(_dt.Rows[i]["DC_ResistanceValue"].ToString().Trim());
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
					teRdcStop_RCSAverageValue.Text = (dcmSumRdcStop / intAvgRdcStopCount).ToString("F3").Trim();
				else
					teRdcStop_RCSAverageValue.Text = "0.000";

				if (intAvgRacStopCount > 0)
					teRacStop_RCSAverageValue.Text = (dcmSumRacStop / intAvgRacStopCount).ToString("F3").Trim();
				else
					teRacStop_RCSAverageValue.Text = "0.000";

				if (intAvgLStopCount > 0)
					teLStop_RCSAverageValue.Text = (dcmSumLStop / intAvgLStopCount).ToString("F3").Trim();
				else
					teLStop_RCSAverageValue.Text = "0.000";

				if (intAvgCStopCount > 0)
					teCStop_RCSAverageValue.Text = (dcmSumCStop / intAvgCStopCount).ToString("F6").Trim();
				else
					teCStop_RCSAverageValue.Text = "0.000";

				if (intAvgQStopCount > 0)
					teQStop_RCSAverageValue.Text = (dcmSumQStop / intAvgQStopCount).ToString("F3").Trim();
				else
					teQStop_RCSAverageValue.Text = "0.000";

                // 올림 평균값 설정
				if (intAvgRdcUpCount > 0)
					teRdcUp_RCSAverageValue.Text = (dcmSumRdcUp / intAvgRdcUpCount).ToString("F3").Trim();
				else
					teRdcUp_RCSAverageValue.Text = "0.000";

				if (intAvgRacUpCount > 0)
					teRacUp_RCSAverageValue.Text = (dcmSumRacUp / intAvgRacUpCount).ToString("F3").Trim();
				else
					teRacUp_RCSAverageValue.Text = "0.000";

				if (intAvgLUpCount > 0)
					teLUp_RCSAverageValue.Text = (dcmSumLUp / intAvgLUpCount).ToString("F3").Trim();
				else
					teLUp_RCSAverageValue.Text = "0.000";

				if (intAvgCUpCount > 0)
					teCUp_RCSAverageValue.Text = (dcmSumCUp / intAvgCUpCount).ToString("F6").Trim();
				else
					teCUp_RCSAverageValue.Text = "0.000";

				if (intAvgQUpCount > 0)
					teQUp_RCSAverageValue.Text = (dcmSumQUp / intAvgQUpCount).ToString("F3").Trim();
				else
					teQUp_RCSAverageValue.Text = "0.000";

                // 이동 평균값 설정
				if (intAvgRdcMoveCount > 0)
					teRdcMove_RCSAverageValue.Text = (dcmSumRdcMove / intAvgRdcMoveCount).ToString("F3").Trim();
				else
                    teRdcMove_RCSAverageValue.Text = "0.000";

				if (intAvgRacMoveCount > 0)
					teRacMove_RCSAverageValue.Text = (dcmSumRacMove / intAvgRacMoveCount).ToString("F3").Trim();
				else
                    teRacMove_RCSAverageValue.Text = "0.000";

				if (intAvgLMoveCount > 0)
					teLMove_RCSAverageValue.Text = (dcmSumLMove / intAvgLMoveCount).ToString("F3").Trim();
				else
                    teLMove_RCSAverageValue.Text = "0.000";

				if (intAvgCMoveCount > 0)
					teCMove_RCSAverageValue.Text = (dcmSumCMove / intAvgCMoveCount).ToString("F6").Trim();
				else
                    teCMove_RCSAverageValue.Text = "0.000";

				if (intAvgQMoveCount > 0)
					teQMove_RCSAverageValue.Text = (dcmSumQMove / intAvgQMoveCount).ToString("F3").Trim();
				else
                    teQMove_RCSAverageValue.Text = "0.000";
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
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
                    dcmDecision_StopRdc = teRdcStop_RCSAverageValue.Text == "" ? 0.000M : Convert.ToDecimal(teRdcStop_RCSAverageValue.Text.Trim());
                    dcmDecision_StopRac = teRacStop_RCSAverageValue.Text == "" ? 0.000M : Convert.ToDecimal(teRacStop_RCSAverageValue.Text.Trim());
                    dcmDecision_StopL = teLStop_RCSAverageValue.Text == "" ? 0.000M : Convert.ToDecimal(teLStop_RCSAverageValue.Text.Trim());
                    dcmDecision_StopC = teCStop_RCSAverageValue.Text == "" ? 0.000M : Convert.ToDecimal(teCStop_RCSAverageValue.Text.Trim());
                    dcmDecision_StopQ = teQStop_RCSAverageValue.Text == "" ? 0.000M : Convert.ToDecimal(teQStop_RCSAverageValue.Text.Trim());

                    dcmDecision_UpRdc = teRdcUp_RCSAverageValue.Text == "" ? 0.000M : Convert.ToDecimal(teRdcUp_RCSAverageValue.Text.Trim());
                    dcmDecision_UpRac = teRacUp_RCSAverageValue.Text == "" ? 0.000M : Convert.ToDecimal(teRacUp_RCSAverageValue.Text.Trim());
                    dcmDecision_UpL = teLUp_RCSAverageValue.Text == "" ? 0.000M : Convert.ToDecimal(teLUp_RCSAverageValue.Text.Trim());
                    dcmDecision_UpC = teCUp_RCSAverageValue.Text == "" ? 0.000M : Convert.ToDecimal(teCUp_RCSAverageValue.Text.Trim());
                    dcmDecision_UpQ = teQUp_RCSAverageValue.Text == "" ? 0.000M : Convert.ToDecimal(teQUp_RCSAverageValue.Text.Trim());

                    dcmDecision_MoveRdc = teRdcMove_RCSAverageValue.Text == "" ? 0.000M : Convert.ToDecimal(teRdcMove_RCSAverageValue.Text.Trim());
                    dcmDecision_MoveRac = teRacMove_RCSAverageValue.Text == "" ? 0.000M : Convert.ToDecimal(teRacMove_RCSAverageValue.Text.Trim());
                    dcmDecision_MoveL = teLMove_RCSAverageValue.Text == "" ? 0.000M : Convert.ToDecimal(teLMove_RCSAverageValue.Text.Trim());
                    dcmDecision_MoveC = teCMove_RCSAverageValue.Text == "" ? 0.000M : Convert.ToDecimal(teCMove_RCSAverageValue.Text.Trim());
                    dcmDecision_MoveQ = teQMove_RCSAverageValue.Text == "" ? 0.000M : Convert.ToDecimal(teQMove_RCSAverageValue.Text.Trim());
                }
                else
                {
                    // 기준값
                    dcmDecision_StopRdc = teRdcStop_RCSReferenceValue.Text == "" ? 0.000M : Convert.ToDecimal(teRdcStop_RCSReferenceValue.Text.Trim());
                    dcmDecision_StopRac = teRacStop_RCSReferenceValue.Text == "" ? 0.000M : Convert.ToDecimal(teRacStop_RCSReferenceValue.Text.Trim());
                    dcmDecision_StopL = teLStop_RCSReferenceValue.Text == "" ? 0.000M : Convert.ToDecimal(teLStop_RCSReferenceValue.Text.Trim());
                    dcmDecision_StopC = teCStop_RCSReferenceValue.Text == "" ? 0.000M : Convert.ToDecimal(teCStop_RCSReferenceValue.Text.Trim());
                    dcmDecision_StopQ = teQStop_RCSReferenceValue.Text == "" ? 0.000M : Convert.ToDecimal(teQStop_RCSReferenceValue.Text.Trim());

                    dcmDecision_UpRdc = teRdcUp_RCSReferenceValue.Text == "" ? 0.000M : Convert.ToDecimal(teRdcUp_RCSReferenceValue.Text.Trim());
                    dcmDecision_UpRac = teRacUp_RCSReferenceValue.Text == "" ? 0.000M : Convert.ToDecimal(teRacUp_RCSReferenceValue.Text.Trim());
                    dcmDecision_UpL = teLUp_RCSReferenceValue.Text == "" ? 0.000M : Convert.ToDecimal(teLUp_RCSReferenceValue.Text.Trim());
                    dcmDecision_UpC = teCUp_RCSReferenceValue.Text == "" ? 0.000M : Convert.ToDecimal(teCUp_RCSReferenceValue.Text.Trim());
                    dcmDecision_UpQ = teQUp_RCSReferenceValue.Text == "" ? 0.000M : Convert.ToDecimal(teQUp_RCSReferenceValue.Text.Trim());

                    dcmDecision_MoveRdc = teRdcMove_RCSReferenceValue.Text == "" ? 0.000M : Convert.ToDecimal(teRdcMove_RCSReferenceValue.Text.Trim());
                    dcmDecision_MoveRac = teRacMove_RCSReferenceValue.Text == "" ? 0.000M : Convert.ToDecimal(teRacMove_RCSReferenceValue.Text.Trim());
                    dcmDecision_MoveL = teLMove_RCSReferenceValue.Text == "" ? 0.000M : Convert.ToDecimal(teLMove_RCSReferenceValue.Text.Trim());
                    dcmDecision_MoveC = teCMove_RCSReferenceValue.Text == "" ? 0.000M : Convert.ToDecimal(teCMove_RCSReferenceValue.Text.Trim());
                    dcmDecision_MoveQ = teQMove_RCSReferenceValue.Text == "" ? 0.000M : Convert.ToDecimal(teQMove_RCSReferenceValue.Text.Trim());
                }

                decimal dcmDecisionRange_Rdc = 0M, dcmDecisionRange_Rac = 0M, dcmDecisionRange_L = 0M, dcmDecisionRange_C = 0M
                    , dcmDecisionRange_Q = 0M, dcmEffectiveStandardRange = 0M;
                dcmDecisionRange_Rdc = Convert.ToDecimal(Gini.GetValue("ReferenceValue", "RdcDecisionRange_ReferenceValue"));
                dcmDecisionRange_Rac = Convert.ToDecimal(Gini.GetValue("ReferenceValue", "RacDecisionRange_ReferenceValue"));
                dcmDecisionRange_L = Convert.ToDecimal(Gini.GetValue("ReferenceValue", "LDecisionRange_ReferenceValue"));
                dcmDecisionRange_C = Convert.ToDecimal(Gini.GetValue("ReferenceValue", "CDecisionRange_ReferenceValue"));
                dcmDecisionRange_Q = Convert.ToDecimal(Gini.GetValue("ReferenceValue", "QDecisionRange_ReferenceValue"));
                dcmEffectiveStandardRange = Convert.ToDecimal(Gini.GetValue("ReferenceValue", "EffectiveStandardRangeOfVariation"));

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
                    dgvMeasurement.Rows[iRow].Cells["DC_Deviation"].Value = (dcmRdcValue - dcmDecision_Rdc).ToString("F3").Trim();
                    _dt.Rows[i]["DC_Deviation"] = (dcmRdcValue - dcmDecision_Rdc).ToString("F3").Trim();

                    if ((dcmRdcValue - dcmDecision_Q) == 0)
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
                    dgvMeasurement.Rows[iRow].Cells["AC_Deviation"].Value = (dcmRacValue - dcmDecision_Rac).ToString("F3").Trim();
                    _dt.Rows[i]["AC_Deviation"] = (dcmRacValue - dcmDecision_Rac).ToString("F3").Trim();

                    if ((dcmRacValue - dcmDecision_Q) == 0)
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
                    dgvMeasurement.Rows[iRow].Cells["L_Deviation"].Value = (dcmLValue - dcmDecision_L).ToString("F3").Trim();
                    _dt.Rows[i]["L_Deviation"] = (dcmLValue - dcmDecision_L).ToString("F3").Trim();

                    if ((dcmLValue - dcmDecision_Q) == 0)
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
                    dgvMeasurement.Rows[iRow].Cells["C_Deviation"].Value = (dcmCValue - dcmDecision_C).ToString("F6").Trim();
                    _dt.Rows[i]["C_Deviation"] = (dcmCValue - dcmDecision_C).ToString("F6").Trim();

                    if ((dcmCValue - dcmDecision_Q) == 0)
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
                    dgvMeasurement.Rows[iRow].Cells["Q_Deviation"].Value = (dcmQValue - dcmDecision_Q).ToString("F3").Trim();
                    _dt.Rows[i]["Q_Deviation"] = (dcmQValue - dcmDecision_Q).ToString("F3").Trim();

                    if ((dcmQValue - dcmDecision_Q) == 0)
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

                    dgvMeasurement.Rows[iRow].Cells["PlantName"].Value = _dt.Rows[i]["PlantName"].ToString().Trim();
                    dgvMeasurement.Rows[iRow].Cells["Hogi"].Value = _dt.Rows[i]["Hogi"].ToString().Trim();
                    dgvMeasurement.Rows[iRow].Cells["Oh_Degree"].Value = _dt.Rows[i]["Oh_Degree"].ToString().Trim();
                    dgvMeasurement.Rows[iRow].Cells["PowerCabinet"].Value = _dt.Rows[i]["PowerCabinet"].ToString().Trim();
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
        public void SetRCSReferenceValueMeasurement(DataTable _dt, string[] _arrayXLabelValue, int _intXLabelValueCountStop, int _intXLabelValueCountUp, int _intXLabelValueCountMove)
        {
            double[] dStopRdc = new double[_intXLabelValueCountStop];
            double[] dUpRdc = new double[_intXLabelValueCountUp];
            double[] dMoveRdc = new double[_intXLabelValueCountMove];

            double[] dRefStopRdc = new double[_intXLabelValueCountStop];
            double[] dRefUpRdc = new double[_intXLabelValueCountUp];
            double[] dRefMoveRdc = new double[_intXLabelValueCountMove];

            double[] dAveStopRdc = new double[_intXLabelValueCountStop];
            double[] dAveUpRdc = new double[_intXLabelValueCountUp];
            double[] dAveMoveRdc = new double[_intXLabelValueCountMove];

            double[] dStopRac = new double[_intXLabelValueCountStop];
            double[] dUpRac = new double[_intXLabelValueCountUp];
            double[] dMoveRac = new double[_intXLabelValueCountMove];

            double[] dRefStopRac = new double[_intXLabelValueCountStop];
            double[] dRefUpRac = new double[_intXLabelValueCountUp];
            double[] dRefMoveRac = new double[_intXLabelValueCountMove];

            double[] dAveStopRac = new double[_intXLabelValueCountStop];
            double[] dAveUpRac = new double[_intXLabelValueCountUp];
            double[] dAveMoveRac = new double[_intXLabelValueCountMove];
            double[] dStopL = new double[_intXLabelValueCountStop];
            double[] dUpL = new double[_intXLabelValueCountUp];
            double[] dMoveL = new double[_intXLabelValueCountMove];

            double[] dRefStopL = new double[_intXLabelValueCountStop];
            double[] dRefUpL = new double[_intXLabelValueCountUp];
            double[] dRefMoveL = new double[_intXLabelValueCountMove];

            double[] dAveStopL = new double[_intXLabelValueCountStop];
            double[] dAveUpL = new double[_intXLabelValueCountUp];
            double[] dAveMoveL = new double[_intXLabelValueCountMove];

            double[] dStopC = new double[_intXLabelValueCountStop];
            double[] dUpC = new double[_intXLabelValueCountUp];
            double[] dMoveC = new double[_intXLabelValueCountMove];

            double[] dRefStopC = new double[_intXLabelValueCountStop];
            double[] dRefUpC = new double[_intXLabelValueCountUp];
            double[] dRefMoveC = new double[_intXLabelValueCountMove];

            double[] dAveStopC = new double[_intXLabelValueCountStop];
            double[] dAveUpC = new double[_intXLabelValueCountUp];
            double[] dAveMoveC = new double[_intXLabelValueCountMove];

            double[] dStopQ = new double[_intXLabelValueCountStop];
            double[] dUpQ = new double[_intXLabelValueCountUp];
            double[] dMoveQ = new double[_intXLabelValueCountMove];

            double[] dRefStopQ = new double[_intXLabelValueCountStop];
            double[] dRefUpQ = new double[_intXLabelValueCountUp];
            double[] dRefMoveQ = new double[_intXLabelValueCountMove];

            double[] dAveStopQ = new double[_intXLabelValueCountStop];
            double[] dAveUpQ = new double[_intXLabelValueCountUp];
            double[] dAveMoveQ = new double[_intXLabelValueCountMove];

            double[] dDevStopRdc = new double[_intXLabelValueCountStop];
            double[] dDevUpRdc = new double[_intXLabelValueCountUp];
            double[] dDevMoveRdc = new double[_intXLabelValueCountMove];
            double[] dDevStopRac = new double[_intXLabelValueCountStop];
            double[] dDevUpRac = new double[_intXLabelValueCountUp];
            double[] dDevMoveRac = new double[_intXLabelValueCountMove];

            double[] dDevStopL = new double[_intXLabelValueCountStop];
            double[] dDevUpL = new double[_intXLabelValueCountUp];
            double[] dDevMoveL = new double[_intXLabelValueCountMove];
            double[] dDevStopC = new double[_intXLabelValueCountStop];
            double[] dDevUpC = new double[_intXLabelValueCountUp];
            double[] dDevMoveC = new double[_intXLabelValueCountMove];
            double[] dDevStopQ = new double[_intXLabelValueCountStop];
            double[] dDevUpQ = new double[_intXLabelValueCountUp];
            double[] dDevMoveQ = new double[_intXLabelValueCountMove];

            double dRdcValue = 0d, dRacValue = 0d, dLValue = 0d, dCValue = 0d, dQValue = 0d
                , dDevRdcValue = 0d, dDevRacValue = 0d, dDevLValue = 0d, dDevCValue = 0d, dDevQValue = 0d;
            int intStopIndex = 0, intUpIndex = 0, intMoveIndex = 0;

            try
            {
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

                        if (dRefStopRdc.Length > intStopIndex)
                        {
                            dRefStopRdc[intStopIndex] = teRdcStop_RCSReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teRdcStop_RCSReferenceValue.Text.Trim());
                            dAveStopRdc[intStopIndex] = teRdcStop_RCSAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teRdcStop_RCSAverageValue.Text.Trim());
                            dRefStopRac[intStopIndex] = teRacStop_RCSReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teRdcStop_RCSReferenceValue.Text.Trim());
                            dAveStopRac[intStopIndex] = teRacStop_RCSAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teRacStop_RCSAverageValue.Text.Trim());
                            dRefStopL[intStopIndex] = teLStop_RCSReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teLStop_RCSReferenceValue.Text.Trim());
                            dAveStopL[intStopIndex] = teLStop_RCSAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teLStop_RCSAverageValue.Text.Trim());
                            dRefStopC[intStopIndex] = teCStop_RCSReferenceValue.Text.Trim() == "" ? 0.000000d : Convert.ToDouble(teCStop_RCSReferenceValue.Text.Trim());
                            dAveStopC[intStopIndex] = teCStop_RCSAverageValue.Text.Trim() == "" ? 0.000000d : Convert.ToDouble(teCStop_RCSAverageValue.Text.Trim());
                            dRefStopQ[intStopIndex] = teQStop_RCSReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teQStop_RCSReferenceValue.Text.Trim());
                            dAveStopQ[intStopIndex] = teQStop_RCSAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teQStop_RCSAverageValue.Text.Trim());
                        }

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

                        if (dRefUpRdc.Length > intUpIndex)
                        {
                            dRefUpRdc[intUpIndex] = teRdcUp_RCSReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teRdcUp_RCSReferenceValue.Text.Trim());
                            dAveUpRdc[intUpIndex] = teRdcUp_RCSAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teRdcUp_RCSAverageValue.Text.Trim());
                            dRefUpRac[intUpIndex] = teRacUp_RCSReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teRdcUp_RCSReferenceValue.Text.Trim());
                            dAveUpRac[intUpIndex] = teRacUp_RCSAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teRacUp_RCSAverageValue.Text.Trim());
                            dRefUpL[intUpIndex] = teLUp_RCSReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teLUp_RCSReferenceValue.Text.Trim());
                            dAveUpL[intUpIndex] = teLUp_RCSAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teLUp_RCSAverageValue.Text.Trim());
                            dRefUpC[intUpIndex] = teCUp_RCSReferenceValue.Text.Trim() == "" ? 0.000000d : Convert.ToDouble(teCUp_RCSReferenceValue.Text.Trim());
                            dAveUpC[intUpIndex] = teCUp_RCSAverageValue.Text.Trim() == "" ? 0.000000d : Convert.ToDouble(teCUp_RCSAverageValue.Text.Trim());
                            dRefUpQ[intUpIndex] = teQUp_RCSReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teQUp_RCSReferenceValue.Text.Trim());
                            dAveUpQ[intUpIndex] = teQUp_RCSAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teQUp_RCSAverageValue.Text.Trim());
                        }

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

                        if (dRefMoveRdc.Length > intMoveIndex)
                        {
                            dRefMoveRdc[intMoveIndex] = teRdcMove_RCSReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teRdcMove_RCSReferenceValue.Text.Trim());
                            dAveMoveRdc[intMoveIndex] = teRdcMove_RCSAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teRdcMove_RCSAverageValue.Text.Trim());
                            dRefMoveRac[intMoveIndex] = teRacMove_RCSReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teRdcMove_RCSReferenceValue.Text.Trim());
                            dAveMoveRac[intMoveIndex] = teRacMove_RCSAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teRacMove_RCSAverageValue.Text.Trim());
                            dRefMoveL[intMoveIndex] = teLMove_RCSReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teLMove_RCSReferenceValue.Text.Trim());
                            dAveMoveL[intMoveIndex] = teLMove_RCSAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teLMove_RCSAverageValue.Text.Trim());
                            dRefMoveC[intMoveIndex] = teCMove_RCSReferenceValue.Text.Trim() == "" ? 0.000000d : Convert.ToDouble(teCMove_RCSReferenceValue.Text.Trim());
                            dAveMoveC[intMoveIndex] = teCMove_RCSAverageValue.Text.Trim() == "" ? 0.000000d : Convert.ToDouble(teCMove_RCSAverageValue.Text.Trim());
                            dRefMoveQ[intMoveIndex] = teQMove_RCSReferenceValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teQMove_RCSReferenceValue.Text.Trim());
                            dAveMoveQ[intMoveIndex] = teQMove_RCSAverageValue.Text.Trim() == "" ? 0.000d : Convert.ToDouble(teQMove_RCSAverageValue.Text.Trim());
                        }

                        intMoveIndex++;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            // DC 저항 Chart 그리기  
            SetChart_Depict(chartMeasurementValueDC, "DC 저항", dStopRdc, dUpRdc, dMoveRdc, dRefStopRdc, dRefUpRdc
                , dRefMoveRdc, dAveStopRdc, dAveUpRdc, dAveMoveRdc, _arrayXLabelValue, 0, "", "");

            // AC 저항 Chart 그리기  
            SetChart_Depict(chartMeasurementValueAC, "AC 저항", dStopRac, dUpRac, dMoveRac, dRefStopRac, dRefUpRac
                , dRefMoveRac, dAveStopRac, dAveUpRac, dAveMoveRac, _arrayXLabelValue, 0, "", "");

            // L 저항 Chart 그리기  
            SetChart_Depict(chartMeasurementValueL, "인덕턴스", dStopL, dUpL, dMoveL, dRefStopL, dRefUpL
                , dRefMoveL, dAveStopL, dAveUpL, dAveMoveL, _arrayXLabelValue, 0, "", "");

            // C 저항 Chart 그리기  
            SetChart_Depict(chartMeasurementValueC, "캐패시턴스", dStopC, dUpC, dMoveC, dRefStopC, dRefUpC
                , dRefMoveC, dAveStopC, dAveUpC, dAveMoveC, _arrayXLabelValue, 0, "", "");

            // Q 저항 Chart 그리기  
            SetChart_Depict(chartMeasurementValueQ, "Q-Factor", dStopQ, dUpQ, dMoveQ, dRefStopQ, dRefUpQ
                , dRefMoveQ, dAveStopQ, dAveUpQ, dAveMoveQ, _arrayXLabelValue, 0, "", "");

            // DC 편차 Chart 그리기  
            SetChart_Depict(chartAverageValueDC, "DC 편차", dDevStopRdc, dDevUpRdc, dDevMoveRdc, _arrayXLabelValue, 0, "", "");

            // AC 편차 Chart 그리기  
            SetChart_Depict(chartAverageValueAC, "AC 편차", dDevStopRac, dDevUpRac, dDevMoveRac, _arrayXLabelValue, 0, "", "");

            // L 저항 Chart 그리기  
            SetChart_Depict(chartAverageValueL, "인덕턴스 편차", dDevStopL, dDevUpL, dDevMoveL, _arrayXLabelValue, 0, "", "");

            // C 저항 Chart 그리기  
            SetChart_Depict(chartAverageValueC, "캐패시턴스 편차", dDevStopC, dDevUpC, dDevMoveC, _arrayXLabelValue, 0, "", "");

            // Q 저항 Chart 그리기  
            SetChart_Depict(chartAverageValueQ, "Q-Factor 편차", dDevStopQ, dDevUpQ, dDevMoveQ, _arrayXLabelValue, 0, "", "");
        }

        public void SetChart_Depict(System.Windows.Forms.DataVisualization.Charting.Chart cht, string chartName
            , double[] dStopValue, double[] dUpValue, double[] dMoveValue, double[] dRefStopValue, double[] dRefUpValue
            , double[] dRefMoveValue, double[] dAveStopValue, double[] dAveUpValue, double[] dAveMoveValue, string[] arrayXLabelValue
            , int chartOption, string coilName, string chartType)
        {
            sSeries = cht.Series.Add("정지");
            sSeries = cht.Series.Add("올림");
            sSeries = cht.Series.Add("이동");
            sSeries = cht.Series.Add("정지 기준값");
            sSeries = cht.Series.Add("올림 기준값");
            sSeries = cht.Series.Add("이동 기준값");
            sSeries = cht.Series.Add("정지 평균값");
            sSeries = cht.Series.Add("올림 평균값");
            sSeries = cht.Series.Add("이동 평균값");

            sSeries.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
            sSeries.XValueType = System.Windows.Forms.DataVisualization.Charting.ChartValueType.Int32;
            sSeries.YValueType = System.Windows.Forms.DataVisualization.Charting.ChartValueType.Double;

            byte chartAreaNo;
            chartAreaNo = 0; //차트 영역 번호.
            m_Chart.InitialChart(ref cht, chartName, 1);

            m_Chart.MakeLegend(ref cht, "정지", "right", chartAreaNo);
            m_Chart.AddChartColumn(ref cht, chartAreaNo, chartAreaNo, "정지", ref arrayXLabelValue, ref dStopValue, true, System.Drawing.Color.Lime, 0, true);

            m_Chart.MakeLegend(ref cht, "올림", "right", chartAreaNo);
            m_Chart.AddChartColumn(ref cht, chartAreaNo, chartAreaNo, "올림", ref arrayXLabelValue, ref dUpValue, true, System.Drawing.Color.Blue, 0, true);

            m_Chart.MakeLegend(ref cht, "이동", "right", chartAreaNo);
            m_Chart.AddChartColumn(ref cht, chartAreaNo, chartAreaNo, "이동", ref arrayXLabelValue, ref dMoveValue, true, System.Drawing.Color.Red, 0, true);

            m_Chart.MakeLegend(ref cht, "정지", "right", chartAreaNo);
            m_Chart.AddChartColumn(ref cht, chartAreaNo, chartAreaNo, "정지 기준값", ref arrayXLabelValue, ref dRefStopValue, true, System.Drawing.Color.Lime, 1, false);

            m_Chart.MakeLegend(ref cht, "올림", "right", chartAreaNo);
            m_Chart.AddChartColumn(ref cht, chartAreaNo, chartAreaNo, "올림 기준값", ref arrayXLabelValue, ref dRefUpValue, true, System.Drawing.Color.Blue, 1, false);

            m_Chart.MakeLegend(ref cht, "이동", "right", chartAreaNo);
            m_Chart.AddChartColumn(ref cht, chartAreaNo, chartAreaNo, "이동 기준값", ref arrayXLabelValue, ref dRefMoveValue, true, System.Drawing.Color.Red, 1, false);

            m_Chart.MakeLegend(ref cht, "정지", "right", chartAreaNo);
            m_Chart.AddChartColumn(ref cht, chartAreaNo, chartAreaNo, "정지 평균값", ref arrayXLabelValue, ref dAveStopValue, true, System.Drawing.Color.Lime, 2, false);

            m_Chart.MakeLegend(ref cht, "올림", "right", chartAreaNo);
            m_Chart.AddChartColumn(ref cht, chartAreaNo, chartAreaNo, "올림 평균값", ref arrayXLabelValue, ref dAveUpValue, true, System.Drawing.Color.Blue, 2, false);

            m_Chart.MakeLegend(ref cht, "이동", "right", chartAreaNo);
            m_Chart.AddChartColumn(ref cht, chartAreaNo, chartAreaNo, "이동 평균값", ref arrayXLabelValue, ref dAveMoveValue, true, System.Drawing.Color.Red, 2, false);

            if (arrayXLabelValue.Length < 30)
                cht.ChartAreas[0].AxisX.Interval = 1; // 차수별일 경우는 X 축 1씩 증가 값 설정
            else
                cht.ChartAreas[0].AxisX.Interval = 10; // 제어봉일 경우는 X 축 10씩 증가 값 설정
        }

        public void SetChart_Depict(System.Windows.Forms.DataVisualization.Charting.Chart cht, string chartName
            , double[] dStopValue, double[] dUpValue, double[] dMoveValue, string[] arrayXLabelValue
            , int chartOption, string coilName, string chartType)
        {
            sSeries = cht.Series.Add("정지");
            sSeries = cht.Series.Add("올림");
            sSeries = cht.Series.Add("이동");

            sSeries.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
            sSeries.XValueType = System.Windows.Forms.DataVisualization.Charting.ChartValueType.Int32;
            sSeries.YValueType = System.Windows.Forms.DataVisualization.Charting.ChartValueType.Double;

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

                tlpAveDCAC.SetRowSpan(panel8, 2);

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

                tlpAveDCAC.SetRowSpan(panel8, 1);

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

                tlpAveDCAC.SetRowSpan(panel8, 1);

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

                tlpRefLCQ.SetRowSpan(panel9, 3);

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

                tlpRefLCQ.SetRowSpan(panel10, 3);

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

                tlpRefLCQ.SetRowSpan(panel9, 2);

                tlpRefLCQ.Controls.Add(this.chartMeasurementValueC, 0, 1);
                tlpRefLCQ.SetRowSpan(chartMeasurementValueC, 1);
                tlpRefLCQ.Controls.Add(this.chartMeasurementValueL, 0, 0);
                tlpRefLCQ.SetRowSpan(chartMeasurementValueL, 1);

                tlpAveLCQ.RowCount = 2;
                tlpAveLCQ.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
                tlpAveLCQ.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));

                tlpRefLCQ.SetRowSpan(panel10, 2);

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

                tlpRefLCQ.SetRowSpan(panel9, 2);

                tlpRefLCQ.Controls.Add(this.chartMeasurementValueQ, 0, 1);
                tlpRefLCQ.SetRowSpan(chartMeasurementValueQ, 1);
                tlpRefLCQ.Controls.Add(this.chartMeasurementValueL, 0, 0);
                tlpRefLCQ.SetRowSpan(chartMeasurementValueL, 1);

                tlpAveLCQ.RowCount = 2;
                tlpAveLCQ.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
                tlpAveLCQ.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));

                tlpRefLCQ.SetRowSpan(panel10, 2);

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

                tlpRefLCQ.SetRowSpan(panel9, 2);

                tlpRefLCQ.Controls.Add(this.chartMeasurementValueQ, 0, 1);
                tlpRefLCQ.SetRowSpan(chartMeasurementValueQ, 1);
                tlpRefLCQ.Controls.Add(this.chartMeasurementValueC, 0, 0);
                tlpRefLCQ.SetRowSpan(chartMeasurementValueC, 1);

                tlpAveLCQ.RowCount = 2;
                tlpAveLCQ.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
                tlpAveLCQ.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));

                tlpRefLCQ.SetRowSpan(panel10, 2);

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

                tlpRefLCQ.SetRowSpan(panel9, 1);

                tlpRefLCQ.Controls.Add(this.chartMeasurementValueL, 0, 0);
                tlpRefLCQ.SetRowSpan(chartMeasurementValueL, 1);

                tlpAveLCQ.RowCount = 1;
                tlpAveLCQ.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));

                tlpRefLCQ.SetRowSpan(panel10, 1);

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

                tlpRefLCQ.SetRowSpan(panel9, 1);

                tlpRefLCQ.Controls.Add(this.chartMeasurementValueC, 0, 0);
                tlpRefLCQ.SetRowSpan(chartMeasurementValueC, 1);

                tlpAveLCQ.RowCount = 1;
                tlpAveLCQ.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));

                tlpRefLCQ.SetRowSpan(panel10, 1);

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

                tlpRefLCQ.SetRowSpan(panel9, 1);

                tlpRefLCQ.Controls.Add(this.chartMeasurementValueQ, 0, 0);
                tlpRefLCQ.SetRowSpan(chartMeasurementValueQ, 1);

                tlpAveLCQ.RowCount = 1;
                tlpAveLCQ.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));

                tlpRefLCQ.SetRowSpan(panel10, 1);

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
