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
using Coil_Diagnostor.Function;

namespace Coil_Diagnostor
{
    public partial class frmAccessDBReader : Form
    {
        protected Function.FunctionMeasureProcess m_MeasureProcess = new Function.FunctionMeasureProcess();
        protected Function.FunctionDataControl m_db = new Function.FunctionDataControl();
        protected Function.FunctionMSAccessDataControl m_Accessdb = new Function.FunctionMSAccessDataControl();
        protected frmMessageBox frmMB = new frmMessageBox();
        protected string strPlantName = "";
        protected bool boolFormLoad = false;

        protected CheckBox allHeaderCheck = new CheckBox();
        protected bool isHeaderCheck = true;

        protected CheckBox allDetailCheck = new CheckBox();
        protected bool isDetailCheck = true;

        public frmAccessDBReader()
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
                case Keys.F1: // Access DB File 폴더바로가기 버튼
                    btnAccessDBFileNameOpen.PerformClick();
                    break;
                case Keys.F2: // 가져오기 버튼
                    btnSearch.PerformClick();
                    break;
                case Keys.F3: // 병합 버튼
                    btnSave.PerformClick();
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
        private void frmAccessDBReader_Load(object sender, EventArgs e)
        {
            boolFormLoad = true;

            strPlantName = Gini.GetValue("Device", "PlantName");

            // ComboBox 데이터 설정
            SetComboBoxDataBinding();

            // Header DataGridView 초기화
            SetHeaderDataGridViewInitialize();

            // 기준값 Data 그리드 초기 설정
            SetReferenceValueDataGridViewInitialize();

            // 초기화
            SetControlInitialize();

            boolFormLoad = false;
        }

        /// <summary>
        /// ComboBox 데이터 설정
        /// </summary>
        private void SetComboBoxDataBinding()
        {

        }

        /// <summary>
        /// 초기화
        /// </summary>
        private void SetControlInitialize()
        {
            // 호기 기본 선택
            nmHogi.Value = 1;

            // 측정 유형
            rboDRPI.Checked = false;
            rboRCS.Checked = true;

            // 폴더 초기화
            teAccessDBFileName.Text = Gini.GetValue("DB", "AccessDBFileName");

            // OH 차수
            nmOh_Degree.Value = Convert.ToDecimal(Gini.GetValue("DB", "AccessDBMaxOhDegree", "1"));

        }

        /// <summary>
        /// 측정 Data 그리드 초기 설정
        /// </summary>
        private void SetHeaderDataGridViewInitialize()
        {
            try
            {
                //----- DataGridView Column 추가
                DataGridViewCheckBoxColumn chkcolumn = new DataGridViewCheckBoxColumn();
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
                DataGridViewTextBoxColumn Column20 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn Column21 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn Column22 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn Column23 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn Column24 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn Column25 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn Column26 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn Column27 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn Column28 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn Column29 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn Column30 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn Column31 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn Column32 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn Column33 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn Column34 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn Column35 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn Column36 = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn Column37 = new DataGridViewTextBoxColumn();

                chkcolumn.HeaderText = "";
                chkcolumn.Name = "Select";
                chkcolumn.Width = 35;
                chkcolumn.ReadOnly = false;
                chkcolumn.Visible = true;
                chkcolumn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                chkcolumn.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column1.HeaderText = "데이터 상태";
                Column1.Name = "UpLoadText";
                Column1.Width = 300;
                Column1.ReadOnly = true;
                Column1.Visible = true;
                Column1.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                Column1.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column2.HeaderText = "발전소";
                Column2.Name = "PlantName";
                Column2.Width = 80;
                Column2.ReadOnly = true;
                Column2.Visible = false;
                Column2.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column2.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column3.HeaderText = "호기";
                Column3.Name = "Hogi";
                Column3.Width = 60;
                Column3.ReadOnly = true;
                Column3.Visible = true;
                Column3.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column3.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column4.HeaderText = "OH 차수";
                Column4.Name = "Oh_Degree";
                Column4.Width = 60;
                Column4.ReadOnly = true;
                Column4.Visible = true;
                Column4.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column4.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column5.HeaderText = "DRPI 그룹";
                Column5.Name = "DRPIGroup";
                Column5.Width = 70;
                Column5.ReadOnly = true;
                Column5.Visible = true;
                Column5.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column5.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column6.HeaderText = "DRPI 타입";
                Column6.Name = "DRPIType";
                Column6.Width = 70;
                Column6.ReadOnly = true;
                Column6.Visible = true;
                Column6.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column6.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column7.HeaderText = "전력함";
                Column7.Name = "PowerCabinet";
                Column7.Width = 70;
                Column7.ReadOnly = true;
                Column7.Visible = true;
                Column7.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column7.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column8.HeaderText = "제어봉ID";
                Column8.Name = "ControlSeqNumber";
                Column8.Width = 70;
                Column8.ReadOnly = true;
                Column8.Visible = false;
                Column8.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column8.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column9.HeaderText = "제어봉";
                Column9.Name = "ControlName";
                Column9.Width = 70;
                Column9.ReadOnly = true;
                Column9.Visible = true;
                Column9.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column9.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column15.HeaderText = "코일ID";
                Column15.Name = "SeqNumber";
                Column15.Width = 70;
                Column15.ReadOnly = true;
                Column15.Visible = false;
                Column15.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column15.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column16.HeaderText = "코일명";
                Column16.Name = "CoilName";
                Column16.Width = 70;
                Column16.ReadOnly = true;
                Column16.Visible = true;
                Column16.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column16.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column17.HeaderText = "코일일련번호";
                Column17.Name = "CoilNumber";
                Column17.Width = 100;
                Column17.ReadOnly = true;
                Column17.Visible = true;
                Column17.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column17.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column14.HeaderText = "측정횟수";
                Column14.Name = "MeasurementCount";
                Column14.Width = 60;
                Column14.ReadOnly = true;
                Column14.Visible = true;
                Column14.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column14.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column34.HeaderText = "Header 측정일시";
                Column34.Name = "HeaderDate";
                Column34.Width = 150;
                Column34.ReadOnly = true;
                Column34.Visible = true;
                Column34.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column34.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column35.HeaderText = "Detail 측정일시";
                Column35.Name = "DetailDate";
                Column35.Width = 150;
                Column35.ReadOnly = true;
                Column35.Visible = true;
                Column35.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column35.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column18.HeaderText = "DC 측정값";
                Column18.Name = "DC_ResistanceValue";
                Column18.Width = 70;
                Column18.ReadOnly = true;
                Column18.Visible = true;
                Column18.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column18.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column19.HeaderText = "편차";
                Column19.Name = "DC_Deviation";
                Column19.Width = 70;
                Column19.ReadOnly = true;
                Column19.Visible = false;
                Column19.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column19.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column20.HeaderText = "DC 상태";
                Column20.Name = "Item_Rdc";
                Column20.Width = 60;
                Column20.ReadOnly = true;
                Column20.Visible = false;
                Column20.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column20.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column21.HeaderText = "AC 측정값";
                Column21.Name = "AC_ResistanceValue";
                Column21.Width = 70;
                Column21.ReadOnly = true;
                Column21.Visible = true;
                Column21.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column21.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column22.HeaderText = "편차";
                Column22.Name = "AC_Deviation";
                Column22.Width = 70;
                Column22.ReadOnly = true;
                Column22.Visible = false;
                Column22.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column22.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column23.HeaderText = "AC 상태";
                Column23.Name = "Item_Rac";
                Column23.Width = 60;
                Column23.ReadOnly = true;
                Column23.Visible = false;
                Column23.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column23.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column24.HeaderText = "인덕턴스";
                Column24.Name = "L_InductanceValue";
                Column24.Width = 70;
                Column24.ReadOnly = true;
                Column24.Visible = true;
                Column24.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column24.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column25.HeaderText = "편차";
                Column25.Name = "L_Deviation";
                Column25.Width = 70;
                Column25.ReadOnly = true;
                Column25.Visible = false;
                Column25.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column25.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column26.HeaderText = "L 상태";
                Column26.Name = "Item_L";
                Column26.Width = 50;
                Column26.ReadOnly = true;
                Column26.Visible = false;
                Column26.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column26.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column27.HeaderText = "캐패시턴스";
                Column27.Name = "C_CapacitanceValue";
                Column27.Width = 80;
                Column27.ReadOnly = true;
                Column27.Visible = true;
                Column27.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column27.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column28.HeaderText = "편차";
                Column28.Name = "C_Deviation";
                Column28.Width = 70;
                Column28.ReadOnly = true;
                Column28.Visible = false;
                Column28.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column28.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column29.HeaderText = "C 상태";
                Column29.Name = "Item_C";
                Column29.Width = 50;
                Column29.ReadOnly = true;
                Column29.Visible = false;
                Column29.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column29.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column30.HeaderText = "Q Factor";
                Column30.Name = "Q_FactorValue";
                Column30.Width = 70;
                Column30.ReadOnly = true;
                Column30.Visible = true;
                Column30.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column30.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column31.HeaderText = "편차";
                Column31.Name = "Q_Deviation";
                Column31.Width = 70;
                Column31.ReadOnly = true;
                Column31.Visible = false;
                Column31.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column31.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column32.HeaderText = "Q 상태";
                Column32.Name = "Item_Q";
                Column32.Width = 50;
                Column32.ReadOnly = true;
                Column32.Visible = false;
                Column32.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column32.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column33.HeaderText = "측정모드";
                Column33.Name = "MeasurementMode";
                Column33.Width = 70;
                Column33.ReadOnly = true;
                Column33.Visible = true;
                Column33.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column33.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column36.HeaderText = "측정결과";
                Column36.Name = "MeasurementResult";
                Column36.Width = 80;
                Column36.ReadOnly = true;
                Column36.Visible = false;
                Column36.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column36.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column10.HeaderText = "온도";
                Column10.Name = "Temperature_ReferenceValue";
                Column10.Width = 50;
                Column10.ReadOnly = true;
                Column10.Visible = true;
                Column10.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column10.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column11.HeaderText = "온도증감";
                Column11.Name = "TemperatureUpDown_ReferenceValue";
                Column11.Width = 70;
                Column11.ReadOnly = true;
                Column11.Visible = true;
                Column11.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column11.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column12.HeaderText = "주파수";
                Column12.Name = "Frequency";
                Column12.Width = 70;
                Column12.ReadOnly = true;
                Column12.Visible = true;
                Column12.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column12.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column13.HeaderText = "전압레벨";
                Column13.Name = "VoltageLevel";
                Column13.Width = 70;
                Column13.ReadOnly = true;
                Column13.Visible = true;
                Column13.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column13.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column37.HeaderText = "TAKE_ID";
                Column37.Name = "TAKE_ID";
                Column37.Width = 80;
                Column37.ReadOnly = true;
                Column37.Visible = true;
                Column37.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column37.SortMode = DataGridViewColumnSortMode.NotSortable;

                dgvAccessData.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] { chkcolumn, Column1, Column2, Column3
                    , Column4, Column5, Column6, Column7, Column8, Column9, Column15, Column16, Column17, Column14, Column34, Column35
                    , Column18, Column19, Column20, Column21, Column22, Column23, Column24, Column25, Column26, Column27, Column28
                    , Column29, Column30, Column31, Column32, Column33, Column36, Column10, Column11, Column12, Column13, Column37});

                dgvAccessData.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing;
                dgvAccessData.ColumnHeadersHeight = 40;
                dgvAccessData.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dgvAccessData.RowHeadersWidth = 40;

                // DataGridView 기타 세팅
                dgvAccessData.AllowUserToAddRows = false; // Row 추가 기능 미사용
                dgvAccessData.AllowUserToDeleteRows = false; // Row 삭제 기능 미사용

                // CheckBox 세팅
                allHeaderCheck.Name = "allHeaderCheck";
                allHeaderCheck.CheckedChanged += new EventHandler(AllHeaderCheckClick);
                allHeaderCheck.Size = new Size(13, 13);
                allHeaderCheck.Location = new Point(((dgvAccessData.Columns[0].Width / 2) - (allHeaderCheck.Width / 2)), (dgvAccessData.ColumnHeadersHeight / 2) - (allHeaderCheck.Height / 2));
                dgvAccessData.Controls.Add(allHeaderCheck); // DataGridView에 CheckBox 추가( 헤더용 )
                allHeaderCheck.Checked = false;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
        }

        /// <summary>
        /// 기준값 Data 그리드 초기 설정
        /// </summary>
        private void SetReferenceValueDataGridViewInitialize()
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

                Column1.HeaderText = "데이터 상태";
                Column1.Name = "UpLoadText";
                Column1.Width = 300;
                Column1.ReadOnly = true;
                Column1.Visible = true;
                Column1.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                Column1.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column2.HeaderText = "발전소";
                Column2.Name = "PlantName";
                Column2.Width = 80;
                Column2.ReadOnly = true;
                Column2.Visible = false;
                Column2.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column2.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column3.HeaderText = "호기";
                Column3.Name = "Hogi";
                Column3.Width = 60;
                Column3.ReadOnly = true;
                Column3.Visible = true;
                Column3.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column3.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column4.HeaderText = "OH 차수";
                Column4.Name = "Oh_Degree";
                Column4.Width = 60;
                Column4.ReadOnly = true;
                Column4.Visible = true;
                Column4.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column4.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column12.HeaderText = "DRPI 그룹";
                Column12.Name = "DRPIGroup";
                Column12.Width = 70;
                Column12.ReadOnly = true;
                Column12.Visible = true;
                Column12.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column12.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column13.HeaderText = "DRPI 타입";
                Column13.Name = "DRPIType";
                Column13.Width = 70;
                Column13.ReadOnly = true;
                Column13.Visible = true;
                Column13.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column13.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column5.HeaderText = "코일ID";
                Column5.Name = "COIL_ID";
                Column5.Width = 70;
                Column5.ReadOnly = true;
                Column5.Visible = true;
                Column5.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column5.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column6.HeaderText = "코일명";
                Column6.Name = "CoilName";
                Column6.Width = 70;
                Column6.ReadOnly = true;
                Column6.Visible = true;
                Column6.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column6.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column7.HeaderText = "DC 기준값";
                Column7.Name = "AVG_DC";
                Column7.Width = 80;
                Column7.ReadOnly = true;
                Column7.Visible = true;
                Column7.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column7.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column8.HeaderText = "AC 기준값";
                Column8.Name = "AVG_AC";
                Column8.Width = 80;
                Column8.ReadOnly = true;
                Column8.Visible = true;
                Column8.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column8.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column9.HeaderText = "인덕턴스";
                Column9.Name = "AVG_L";
                Column9.Width = 80;
                Column9.ReadOnly = true;
                Column9.Visible = true;
                Column9.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column9.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column10.HeaderText = "캐패시턴스";
                Column10.Name = "AVG_C";
                Column10.Width = 80;
                Column10.ReadOnly = true;
                Column10.Visible = true;
                Column10.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column10.SortMode = DataGridViewColumnSortMode.NotSortable;

                Column11.HeaderText = "Q Factor";
                Column11.Name = "AVG_Q";
                Column11.Width = 80;
                Column11.ReadOnly = true;
                Column11.Visible = true;
                Column11.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Column11.SortMode = DataGridViewColumnSortMode.NotSortable;

                dgvReferenceValue.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] { Column1, Column2, Column3
                    , Column4, Column12, Column13, Column5, Column6, Column7, Column8,Column9, Column10, Column11 });

                dgvReferenceValue.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing;
                dgvReferenceValue.ColumnHeadersHeight = 40;
                dgvReferenceValue.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dgvReferenceValue.RowHeadersWidth = 40;

                // DataGridView 기타 세팅
                dgvReferenceValue.AllowUserToAddRows = false; // Row 추가 기능 미사용
                dgvReferenceValue.AllowUserToDeleteRows = false; // Row 삭제 기능 미사용
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
        public void AllHeaderCheckClick(object sender, EventArgs e)
        {
            if (isHeaderCheck)
            {
                if (allHeaderCheck.Checked)
                    for (int i = 0; i < dgvAccessData.Rows.Count; i++)
                        dgvAccessData.Rows[i].Cells[0].Value = true;
                else
                    for (int i = 0; i < dgvAccessData.Rows.Count; i++)
                        dgvAccessData.Rows[i].Cells[0].Value = false;

                dgvAccessData.EndEdit(DataGridViewDataErrorContexts.Commit); // << 이거 안할경우 선택된 Cell이 CheckBoxCell일 경우 변화가 없는것처럼 보임
            }
        }

        /// <summary>
        /// DataGridView Cell Click Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgvAccessData_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (((DataGridView)sender).RowCount < 1) return;

            isHeaderCheck = false;

            if (e.RowIndex >= 0 && e.ColumnIndex == 0)
            {
                ((DataGridView)sender).Rows[e.RowIndex].Cells["Select"].Value = !(bool)((DataGridView)sender).Rows[e.RowIndex].Cells["Select"].Value;
                allHeaderCheck.Checked = true;

                for (int i = 0; i < ((DataGridView)sender).Rows.Count; i++)
                {
                    if (!(bool)((DataGridView)sender).Rows[i].Cells["Select"].Value)
                    {
                        allHeaderCheck.Checked = false;
                        break;
                    }
                }
            }

            isHeaderCheck = true;
        }

        /// <summary>
        /// DataGridView ColumnDividerWidthChanged Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgvAccessData_ColumnDividerWidthChanged(object sender, DataGridViewColumnEventArgs e)
        {
            CheckBoxFixed((DataGridView)sender, allHeaderCheck);
        }

        /// <summary>
        /// DataGridView RowHeadersWidthChanged Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgvAccessData_RowHeadersWidthChanged(object sender, EventArgs e)
        {
            CheckBoxFixed((DataGridView)sender, allHeaderCheck);
        }

        /// <summary>
        /// DataGridView Scroll Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgvAccessData_Scroll(object sender, ScrollEventArgs e)
        {
            CheckBoxFixed((DataGridView)sender, allHeaderCheck);
        }

        /// <summary>
        /// 콤보박스 위치 변경함수
        /// </summary>
        /// <param name="dgv">데이터그리드뷰</param>
        /// <param name="c">콤보박스</param>
        public void CheckBoxFixed(DataGridView _dgv, Control c)
        {
            // 표시된 첫번째 열이 0일경우 >> 스크롤 해서 "선택" 열이 넘어가면 1로 됨
            if (_dgv.FirstDisplayedScrollingColumnIndex == 0)
            {
                // 현재 스크롤 되어 보이지 않는 열의 가로 길이 >> "선택"열의 Width가 60이고 스크롤되어 "선택"열이 반만 보일경우 해당값은 30임
                // 따라서 열의 가려진 크기가 "선택"행의 Width - CheckBox.Width - 7보다 크다면 숨김처리( 해당 처리를 안하면 CheckBox는 자동 숨기기가 안되서 공중에 붕뜬 느낌으로 남아있음 )
                // 따라서 열의 가려진 크기가 "선택"행의 Width - CheckBox.Width - 7은 >> 60 - 13 - 7 따라서 체크박스가 숨겨진 길이에 닿는 순간임
                if (_dgv.FirstDisplayedScrollingColumnHiddenWidth > (_dgv.Columns[0].Width - c.Width - 7))
                    c.Visible = false;
                // 위 내용이 아니면 다시 보이게 Visible = true 처리
                // Point.X 계산은 RowHeadersWidth + "선택"행 Width - CheckBox.WIdth - 7 - 숨겨진행의 길이
                // Point.Y 계산은 ( 열 높이 / 2 ) - ( CheckBox.Heiht / 2 ) >> css 에서 가운데 마추는 형식과 비슷하다고 보면 됨
                else
                {
                    c.Location = new Point((_dgv.Columns[0].Width / 2) - (c.Width / 2), (_dgv.ColumnHeadersHeight / 2) - (c.Height / 2));
                    c.Visible = true;
                }
            }
            // 표시된 첫번째 열이 0보다 크면 "선택"행은 숨겨진 상태므로 숨김 처리
            else if (_dgv.FirstDisplayedScrollingColumnIndex > 0)
                c.Visible = false;
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
                ((NumericUpDown)sender).Value = 1M;
                ((NumericUpDown)sender).Text = "1";
            }
        }

        /// <summary>
        /// RCS 선택
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void rboRCS_CheckedChanged(object sender, EventArgs e)
        {
            if (rboRCS.Checked)
            {
                dgvAccessData.Columns["DRPIGroup"].Visible = false;
                dgvAccessData.Columns["DRPIType"].Visible = false;

                dgvAccessData.Columns["PowerCabinet"].Visible = true;
            }

            dgvAccessData.Rows.Clear();
        }

        /// <summary>
        /// DRPI 선택
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void rboDRPI_CheckedChanged(object sender, EventArgs e)
        {
            if (rboDRPI.Checked)
            {
                dgvAccessData.Columns["DRPIGroup"].Visible = true;
                dgvAccessData.Columns["DRPIType"].Visible = true;

                dgvAccessData.Columns["PowerCabinet"].Visible = false;
            }

            dgvAccessData.Rows.Clear();
        }

        /// <summary>
        /// 호기 변경 Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cboHogi_SelectedIndexChanged(object sender, EventArgs e)
        {
            dgvAccessData.Rows.Clear();
        }

        /// <summary>
        /// Oh Degree 변경 Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void nmOh_Degree_ValueChanged(object sender, EventArgs e)
        {
            dgvAccessData.Rows.Clear();
        }

        /// <summary>
        /// 폴더 오픈
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnAccessDBFileNameOpen_Click(object sender, EventArgs e)
        {
            try
            {
                dgvAccessData.Rows.Clear();

                /*
                // 파일 오픈
                OpenFileDialog ofd = new OpenFileDialog();
                ofd.Filter = "Access File(*.mdb)|*.mdb|AllFiles(*.*)|*.*";
                ofd.Title = "Access DB 파일을 선택해 주십시오";

                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    teAccessDBFileName.Text = ofd.FileName.Trim();
                }
                * */

                FolderBrowserDialog dialog = new FolderBrowserDialog();
                dialog.SelectedPath = teAccessDBFileName.Text.Trim();

                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    teAccessDBFileName.Text = dialog.SelectedPath.Trim();
                }

                if(!Gini.SetValue("DB", "AccessDBFileName", teAccessDBFileName.Text.Trim()))
                {
                    frmMB.lblMessage.Text = "INI 저장 실패";
                    frmMB.TopMost = true;
                    frmMB.ShowDialog();
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
            if (!Gini.SetValue("DB", "AccessDBMaxOhDegree", nmOh_Degree.Value.ToString().Trim()))
            {
                frmMB.lblMessage.Text = "INI 저장 실패";
                frmMB.TopMost = true;
                frmMB.ShowDialog();
            }

            this.Close();
        }

        /// <summary>
        /// 병합 Button Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSave_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;

            bool boolHeader = false;
            bool boolDetail = false;
            progressBar1.Visible = true;
            int intDataCount = 0;

            try
            {
                if (dgvAccessData.RowCount > 0)
                {
                    int intMeasurementCount = 0, intItem_Rdc = 0, intItem_Rac = 0, intItem_L = 0, intItem_C = 0, intItem_Q = 0
                        , intControlSeqNumber = 0, intCoilNumber = 0, intSeqNumber = 0, intTAKE_ID = 0;
                    string strHogi = "", strOh_Degree = "", strDRPIGroup = "", strDRPIType = "", strPowerCabinet = "", strFrequency = "", strVoltageLevel = ""
                        , strMeasurementMode = "", strMeasurementResult = "", strHeaderDate = "", strDetailDate = "", strControlName = "", strCoilName = ""
                        , strHeaderMessage = "";
                    decimal dcmTemperature_ReferenceValue = 0M, dcmTemperatureUpDown_ReferenceValue = 0M
                        , dcmDC_VAL = 0M, dcmAC_VAL = 0M, dcmL_VAL = 0M, dcmC_VAL = 0M, dcmQ_VAL = 0M;
                    
                    progressBar1.Value = 0;

                    for (int i = 0; i < dgvAccessData.RowCount; i++)
                    {
                        dgvAccessData.Rows[i].Selected = true;

                        if ((bool)dgvAccessData.Rows[i].Cells["Select"].Value)
                        {
                            string strSelectDRPIGroup = dgvAccessData.Rows[i].Cells["DRPIGroup"].Value == null ? "" : dgvAccessData.Rows[i].Cells["DRPIGroup"].Value.ToString().Trim();
                            string strSelectDRPIType = dgvAccessData.Rows[i].Cells["DRPIType"].Value == null ? "" : dgvAccessData.Rows[i].Cells["DRPIType"].Value.ToString().Trim();
                            string strSelectPowerCabinet = dgvAccessData.Rows[i].Cells["PowerCabinet"].Value == null ? "" : dgvAccessData.Rows[i].Cells["PowerCabinet"].Value.ToString().Trim();
                            string strSelectControlName = dgvAccessData.Rows[i].Cells["ControlName"].Value == null ? "" : dgvAccessData.Rows[i].Cells["ControlName"].Value.ToString().Trim();

                            if (dgvAccessData.Rows[i].Cells["Hogi"].Value.ToString().Trim() != strHogi.Trim()
                                || dgvAccessData.Rows[i].Cells["Oh_Degree"].Value.ToString().Trim() != strOh_Degree.Trim()
                                || strSelectDRPIGroup.Trim() != strDRPIGroup.Trim()
                                || strSelectDRPIType.Trim() != strDRPIType.Trim()
                                || strSelectPowerCabinet.Trim() != strPowerCabinet.Trim()
                                || strSelectControlName.Trim() != strControlName.Trim())
                            {
                                strHogi = dgvAccessData.Rows[i].Cells["Hogi"].Value == null ? "" : dgvAccessData.Rows[i].Cells["Hogi"].Value.ToString().Trim();
                                strOh_Degree = dgvAccessData.Rows[i].Cells["Oh_Degree"].Value == null ? "" : dgvAccessData.Rows[i].Cells["Oh_Degree"].Value.ToString().Trim();
                                strDRPIGroup = dgvAccessData.Rows[i].Cells["DRPIGroup"].Value == null ? "" : dgvAccessData.Rows[i].Cells["DRPIGroup"].Value.ToString().Trim();
                                strDRPIType = dgvAccessData.Rows[i].Cells["DRPIType"].Value == null ? "" : dgvAccessData.Rows[i].Cells["DRPIType"].Value.ToString().Trim();
                                strPowerCabinet = dgvAccessData.Rows[i].Cells["PowerCabinet"].Value == null ? "" : dgvAccessData.Rows[i].Cells["PowerCabinet"].Value.ToString().Trim();
                                strControlName = dgvAccessData.Rows[i].Cells["ControlName"].Value == null ? "" : dgvAccessData.Rows[i].Cells["ControlName"].Value.ToString().Trim();

                                dcmTemperature_ReferenceValue = dgvAccessData.Rows[i].Cells["Temperature_ReferenceValue"].Value == null
                                    || dgvAccessData.Rows[i].Cells["Temperature_ReferenceValue"].Value.ToString().Trim() == ""
                                    ? 0M : Convert.ToDecimal(dgvAccessData.Rows[i].Cells["Temperature_ReferenceValue"].Value.ToString().Trim());
                                dcmTemperatureUpDown_ReferenceValue = dgvAccessData.Rows[i].Cells["TemperatureUpDown_ReferenceValue"].Value == null
                                    || dgvAccessData.Rows[i].Cells["TemperatureUpDown_ReferenceValue"].Value.ToString().Trim() == ""
                                    ? 0M : Convert.ToDecimal(dgvAccessData.Rows[i].Cells["TemperatureUpDown_ReferenceValue"].Value.ToString().Trim());

                                strFrequency = dgvAccessData.Rows[i].Cells["Frequency"].Value == null ? "" : dgvAccessData.Rows[i].Cells["Frequency"].Value.ToString().Trim();
                                strVoltageLevel = dgvAccessData.Rows[i].Cells["VoltageLevel"].Value == null ? "" : dgvAccessData.Rows[i].Cells["VoltageLevel"].Value.ToString().Trim();
                                strMeasurementMode = dgvAccessData.Rows[i].Cells["MeasurementMode"].Value == null ? "" : dgvAccessData.Rows[i].Cells["MeasurementMode"].Value.ToString().Trim();
                                strMeasurementResult = dgvAccessData.Rows[i].Cells["MeasurementResult"].Value == null ? "" : dgvAccessData.Rows[i].Cells["MeasurementResult"].Value.ToString().Trim();
                                strHeaderDate = dgvAccessData.Rows[i].Cells["HeaderDate"].Value == null ? "" : dgvAccessData.Rows[i].Cells["HeaderDate"].Value.ToString().Trim();
                                strDetailDate = dgvAccessData.Rows[i].Cells["DetailDate"].Value == null ? "" : dgvAccessData.Rows[i].Cells["DetailDate"].Value.ToString().Trim();

                                intMeasurementCount = dgvAccessData.Rows[i].Cells["MeasurementCount"].Value == null
                                    || dgvAccessData.Rows[i].Cells["MeasurementCount"].Value.ToString().Trim() == ""
                                    ? 0 : Convert.ToInt32(dgvAccessData.Rows[i].Cells["MeasurementCount"].Value.ToString().Trim());
                                intItem_Rdc = dgvAccessData.Rows[i].Cells["Item_Rdc"].Value == null
                                    || dgvAccessData.Rows[i].Cells["Item_Rdc"].Value.ToString().Trim() == ""
                                    ? 0 : Convert.ToInt32(dgvAccessData.Rows[i].Cells["Item_Rdc"].Value.ToString().Trim());
                                intItem_Rac = dgvAccessData.Rows[i].Cells["Item_Rac"].Value == null
                                    || dgvAccessData.Rows[i].Cells["Item_Rac"].Value.ToString().Trim() == ""
                                    ? 0 : Convert.ToInt32(dgvAccessData.Rows[i].Cells["Item_Rac"].Value.ToString().Trim());
                                intItem_L = dgvAccessData.Rows[i].Cells["Item_L"].Value == null
                                    || dgvAccessData.Rows[i].Cells["Item_L"].Value.ToString().Trim() == ""
                                    ? 0 : Convert.ToInt32(dgvAccessData.Rows[i].Cells["Item_L"].Value.ToString().Trim());
                                intItem_C = dgvAccessData.Rows[i].Cells["Item_C"].Value == null
                                    || dgvAccessData.Rows[i].Cells["Item_C"].Value.ToString().Trim() == ""
                                    ? 0 : Convert.ToInt32(dgvAccessData.Rows[i].Cells["Item_C"].Value.ToString().Trim());
                                intItem_Q = dgvAccessData.Rows[i].Cells["Item_Q"].Value == null
                                    || dgvAccessData.Rows[i].Cells["Item_Q"].Value.ToString().Trim() == ""
                                    ? 0 : Convert.ToInt32(dgvAccessData.Rows[i].Cells["Item_Q"].Value.ToString().Trim());
                                intTAKE_ID = dgvAccessData.Rows[i].Cells["TAKE_ID"].Value == null
                                    || dgvAccessData.Rows[i].Cells["TAKE_ID"].Value.ToString().Trim() == ""
                                    ? 0 : Convert.ToInt32(dgvAccessData.Rows[i].Cells["TAKE_ID"].Value.ToString().Trim());

                                if (rboRCS.Checked)
                                {
                                    // Header 기존 데이터 체크
                                    if ((m_db.GetRCSDiagnosisHeaderDataGridViewDataCount(strPlantName.Trim(), strHogi.Trim(), strOh_Degree.Trim(), strPowerCabinet.Trim(), strControlName.Trim()) > 0))
                                    {
                                        // 기존 데이터 삭제
                                        boolHeader = m_db.DeleteRCSDiagnosisHeaderDataGridViewDataInfo(strPlantName.Trim(), strHogi, strOh_Degree, strPowerCabinet, strControlName.Trim());
                                    }

                                    // 데이터 저상
                                    boolHeader = m_db.InsertRCSDiagnosisHeaderDataGridViewDataInfo(strPlantName.Trim(), strHogi.Trim(), strOh_Degree.Trim(), strPowerCabinet.Trim()
                                        , strControlName.Trim(), dcmTemperature_ReferenceValue, dcmTemperatureUpDown_ReferenceValue, strFrequency.Trim(), strVoltageLevel.Trim()
                                        , intMeasurementCount, intItem_Rdc, intItem_Rac, intItem_L, intItem_C, intItem_Q, strMeasurementMode.Trim(), strHeaderDate.Trim()
                                        , strMeasurementResult.Trim(), intTAKE_ID);
                                }
                                else
                                {
                                    // Header 기존 데이터 체크
                                    if ((m_db.GetDRPIDiagnosisHeaderDataGridViewDataCount(strPlantName.Trim(), strHogi.Trim(), strOh_Degree.Trim(), strDRPIGroup.Trim(), strDRPIType.Trim(), strControlName.Trim()) > 0))
                                    {
                                        // 기존 데이터 삭제
                                        boolHeader = m_db.DeleteDRPIDiagnosisHeaderDataGridViewDataInfo(strPlantName.Trim(), strHogi.Trim(), strOh_Degree.Trim(), strDRPIGroup.Trim(), strDRPIType.Trim(), strControlName.Trim());
                                    }

                                    // 데이터 저상
                                    boolHeader = m_db.InsertDRPIDiagnosisHeaderDataGridViewDataInfo(strPlantName.Trim(), strHogi.Trim(), strOh_Degree.Trim(), strDRPIGroup.Trim(), strDRPIType.Trim()
                                        , strControlName.Trim(), dcmTemperature_ReferenceValue, dcmTemperatureUpDown_ReferenceValue, strFrequency.Trim(), strVoltageLevel.Trim()
                                        , intMeasurementCount, intItem_Rdc, intItem_Rac, intItem_L, intItem_C, intItem_Q, strMeasurementMode.Trim(), strHeaderDate.Trim()
                                        , strMeasurementResult.Trim(), intTAKE_ID);
                                }

                                if (boolHeader)
                                    strHeaderMessage = "Header 저장 완료";
                                else
                                    strHeaderMessage = "Header 저장 실패";
                                
                                dgvAccessData.Rows[i].Cells["UpLoadText"].Value = strHeaderMessage;
                            }

                            strCoilName = dgvAccessData.Rows[i].Cells["CoilName"].Value == null ? "" : dgvAccessData.Rows[i].Cells["CoilName"].Value.ToString().Trim();

                            intControlSeqNumber = dgvAccessData.Rows[i].Cells["ControlSeqNumber"].Value == null
                                || dgvAccessData.Rows[i].Cells["ControlSeqNumber"].Value.ToString().Trim() == ""
                                ? 0 : Convert.ToInt32(dgvAccessData.Rows[i].Cells["ControlSeqNumber"].Value.ToString().Trim());
                            intCoilNumber = dgvAccessData.Rows[i].Cells["CoilNumber"].Value == null
                                || dgvAccessData.Rows[i].Cells["CoilNumber"].Value.ToString().Trim() == ""
                                ? 0 : Convert.ToInt32(dgvAccessData.Rows[i].Cells["CoilNumber"].Value.ToString().Trim());
                            intSeqNumber = dgvAccessData.Rows[i].Cells["SeqNumber"].Value == null
                                || dgvAccessData.Rows[i].Cells["SeqNumber"].Value.ToString().Trim() == ""
                                ? 0 : Convert.ToInt32(dgvAccessData.Rows[i].Cells["SeqNumber"].Value.ToString().Trim());

                            dcmDC_VAL = dgvAccessData.Rows[i].Cells["DC_ResistanceValue"].Value == null
                                || dgvAccessData.Rows[i].Cells["DC_ResistanceValue"].Value.ToString().Trim() == ""
                                ? 0.000M : Convert.ToDecimal(dgvAccessData.Rows[i].Cells["DC_ResistanceValue"].Value.ToString().Trim());
                            dcmAC_VAL = dgvAccessData.Rows[i].Cells["AC_ResistanceValue"].Value == null
                                || dgvAccessData.Rows[i].Cells["AC_ResistanceValue"].Value.ToString().Trim() == ""
                                ? 0.000M : Convert.ToDecimal(dgvAccessData.Rows[i].Cells["AC_ResistanceValue"].Value.ToString().Trim());
                            dcmL_VAL = dgvAccessData.Rows[i].Cells["L_InductanceValue"].Value == null
                                || dgvAccessData.Rows[i].Cells["L_InductanceValue"].Value.ToString().Trim() == ""
                                ? 0.000M : Convert.ToDecimal(dgvAccessData.Rows[i].Cells["L_InductanceValue"].Value.ToString().Trim());
                            dcmC_VAL = dgvAccessData.Rows[i].Cells["C_CapacitanceValue"].Value == null
                                || dgvAccessData.Rows[i].Cells["C_CapacitanceValue"].Value.ToString().Trim() == ""
                                ? 0.000000M : Convert.ToDecimal(dgvAccessData.Rows[i].Cells["C_CapacitanceValue"].Value.ToString().Trim());
                            dcmQ_VAL = dgvAccessData.Rows[i].Cells["Q_FactorValue"].Value == null
                                || dgvAccessData.Rows[i].Cells["Q_FactorValue"].Value.ToString().Trim() == ""
                                ? 0.0M : Convert.ToDecimal(dgvAccessData.Rows[i].Cells["Q_FactorValue"].Value.ToString().Trim());

                            if (rboRCS.Checked)
                            {
                                // Detail 기준 데이터 체크 
                                if ((m_db.GetRCSDiagnosisDetailDataGridViewDataCount(strPlantName.Trim(), strHogi.Trim(), strOh_Degree.Trim(), strPowerCabinet.Trim(), strControlName.Trim(), strCoilName.Trim(), intCoilNumber)) > 0)
                                {
                                    // 기존 데이터 삭제
                                    boolDetail = m_db.DeleteRCSDiagnosisDetailDataGridViewDataInfo(strPlantName.Trim(), strHogi.Trim(), strOh_Degree.Trim(), strPowerCabinet.Trim()
                                        , strControlName.Trim(), strCoilName.Trim(), intCoilNumber);
                                }

                                // 데이터 저장
                                boolDetail = m_db.InsertRCSDiagnosisDetailDataGridViewDataInfo(strPlantName.Trim(), strHogi.Trim(), strOh_Degree.Trim(), strPowerCabinet.Trim()
                                    , strControlName.Trim(), strCoilName.Trim(), intCoilNumber, dcmDC_VAL, 0.000M, dcmAC_VAL, 0.000M, dcmL_VAL, 0.000M, dcmC_VAL, 0.000000M
                                    , dcmQ_VAL, 0.0M, strMeasurementResult.Trim(), strDetailDate.Trim(), intSeqNumber, intControlSeqNumber);
                            }
                            else
                            {
                                // Detail 기준 데이터 체크 
                                if ((m_db.GetDRPIDiagnosisDetailDataGridViewDataCount(strPlantName.Trim(), strHogi.Trim(), strOh_Degree.Trim(), strDRPIGroup.Trim(), strDRPIType.Trim(), strControlName.Trim(), strCoilName.Trim(), intCoilNumber)) > 0)
                                {
                                    // 기존 데이터 삭제
                                    boolDetail = m_db.DeleteDRPIDiagnosisDetailDataGridViewDataInfo(strPlantName.Trim(), strHogi.Trim(), strOh_Degree.Trim(), strDRPIGroup.Trim(), strDRPIType.Trim()
                                        , strControlName.Trim(), strCoilName.Trim(), intCoilNumber);
                                }

                                // 데이터 저장
                                boolDetail = m_db.InsertDRPIDiagnosisDetailDataGridViewDataInfo(strPlantName.Trim(), strHogi.Trim(), strOh_Degree.Trim(), strDRPIGroup.Trim(), strDRPIType.Trim()
                                    , strControlName.Trim(), strCoilName.Trim(), intCoilNumber, strMeasurementMode.Trim(), dcmDC_VAL, 0.000M, dcmAC_VAL, 0.000M, dcmL_VAL, 0.000M, dcmC_VAL, 0.000000M
                                    , dcmQ_VAL, 0.0M, strMeasurementResult.Trim(), strDetailDate.Trim(), intSeqNumber, intControlSeqNumber);
                            }

                            intDataCount++;

                            if (boolDetail)
                                dgvAccessData.Rows[i].Cells["UpLoadText"].Value = string.Format("{0}, Detail 저장 완료", strHeaderMessage.ToString().Trim());
                            else
                                dgvAccessData.Rows[i].Cells["UpLoadText"].Value = string.Format("{0}, Detail 저장 실패", strHeaderMessage.ToString().Trim());
                        }

                        progressBar1.Value = (i * 100) / dgvAccessData.RowCount;
                        dgvAccessData.Rows[i].Selected = false;
                    }
                }
                
                // 기준 값 가져오기
                if (chkReferenceValue.Checked)
                    boolDetail = SaveReferenceValueData();

                progressBar1.Value = 100;

                frmMB.lblMessage.Text = "저장 완료";
                frmMB.TopMost = true;
                frmMB.ShowDialog();
            }
            catch (Exception ex)
            {
                frmMB.lblMessage.Text = "저장 실패";
                frmMB.TopMost = true;
                frmMB.ShowDialog();
                System.Diagnostics.Debug.Print(ex.Message);
            }
            finally
            {
                progressBar1.Visible = false;
            }

            System.Windows.Forms.Cursor.Current = Cursors.Default;
        }

        private bool SaveReferenceValueData()
        {
            bool boolResult = false;
            string strSelectHogi = nmHogi.Value == null ? "1 호기" : (Convert.ToInt32(nmHogi.Value).ToString().Trim() + " 호기");
            string strOh_Degree = nmOh_Degree.Value == null ? "제 0 차" : ("제 " + nmOh_Degree.Value.ToString().Trim() + " 차");

            if (rboRCS.Checked)
            {
                #region ▣ RCS 기준값 저장

                double dRef_RdcStop = 0d, dRef_RdcUP = 0d, dRef_RdcMove = 0d;
                double dRef_RacStop = 0d, dRef_RacUP = 0d, dRef_RacMove = 0d;
                double dRef_LStop = 0d, dRef_LUP = 0d, dRef_LMove = 0d;
                double dRef_CStop = 0d, dRef_CUP = 0d, dRef_CMove = 0d;
                double dRef_QStop = 0d, dRef_QUP = 0d, dRef_QMove = 0d;

                for (int i = 0; i < dgvReferenceValue.RowCount; i++)
                {
                    if (dgvReferenceValue.Rows[i].Cells["CoilName"].Value.ToString().Trim() == "정지")
                    {
                        dRef_RdcStop = dgvReferenceValue.Rows[i].Cells["AVG_DC"].Value == null
                            || dgvReferenceValue.Rows[i].Cells["AVG_DC"].Value.ToString().Trim() == ""
                            ? 0d : Convert.ToDouble(dgvReferenceValue.Rows[i].Cells["AVG_DC"].Value.ToString().Trim());

                        dRef_RacStop = dgvReferenceValue.Rows[i].Cells["AVG_AC"].Value == null
                            || dgvReferenceValue.Rows[i].Cells["AVG_AC"].Value.ToString().Trim() == ""
                            ? 0d : Convert.ToDouble(dgvReferenceValue.Rows[i].Cells["AVG_AC"].Value.ToString().Trim());

                        dRef_LStop = dgvReferenceValue.Rows[i].Cells["AVG_L"].Value == null
                            || dgvReferenceValue.Rows[i].Cells["AVG_L"].Value.ToString().Trim() == ""
                            ? 0d : Convert.ToDouble(dgvReferenceValue.Rows[i].Cells["AVG_L"].Value.ToString().Trim());

                        dRef_CStop = dgvReferenceValue.Rows[i].Cells["AVG_C"].Value == null
                            || dgvReferenceValue.Rows[i].Cells["AVG_C"].Value.ToString().Trim() == ""
                            ? 0d : Convert.ToDouble(dgvReferenceValue.Rows[i].Cells["AVG_C"].Value.ToString().Trim());

                        dRef_QStop = dgvReferenceValue.Rows[i].Cells["AVG_Q"].Value == null
                            || dgvReferenceValue.Rows[i].Cells["AVG_Q"].Value.ToString().Trim() == ""
                            ? 0d : Convert.ToDouble(dgvReferenceValue.Rows[i].Cells["AVG_Q"].Value.ToString().Trim());
                    }
                    else if (dgvReferenceValue.Rows[i].Cells["CoilName"].Value.ToString().Trim() == "올림")
                    {
                        dRef_RdcUP = dgvReferenceValue.Rows[i].Cells["AVG_DC"].Value == null
                            || dgvReferenceValue.Rows[i].Cells["AVG_DC"].Value.ToString().Trim() == ""
                            ? 0d : Convert.ToDouble(dgvReferenceValue.Rows[i].Cells["AVG_DC"].Value.ToString().Trim());

                        dRef_RacUP = dgvReferenceValue.Rows[i].Cells["AVG_AC"].Value == null
                            || dgvReferenceValue.Rows[i].Cells["AVG_AC"].Value.ToString().Trim() == ""
                            ? 0d : Convert.ToDouble(dgvReferenceValue.Rows[i].Cells["AVG_AC"].Value.ToString().Trim());

                        dRef_LUP = dgvReferenceValue.Rows[i].Cells["AVG_L"].Value == null
                            || dgvReferenceValue.Rows[i].Cells["AVG_L"].Value.ToString().Trim() == ""
                            ? 0d : Convert.ToDouble(dgvReferenceValue.Rows[i].Cells["AVG_L"].Value.ToString().Trim());

                        dRef_CUP = dgvReferenceValue.Rows[i].Cells["AVG_C"].Value == null
                            || dgvReferenceValue.Rows[i].Cells["AVG_C"].Value.ToString().Trim() == ""
                            ? 0d : Convert.ToDouble(dgvReferenceValue.Rows[i].Cells["AVG_C"].Value.ToString().Trim());

                        dRef_QUP = dgvReferenceValue.Rows[i].Cells["AVG_Q"].Value == null
                            || dgvReferenceValue.Rows[i].Cells["AVG_Q"].Value.ToString().Trim() == ""
                            ? 0d : Convert.ToDouble(dgvReferenceValue.Rows[i].Cells["AVG_Q"].Value.ToString().Trim());
                    }
                    else if (dgvReferenceValue.Rows[i].Cells["CoilName"].Value.ToString().Trim() == "이동")
                    {
                        dRef_RdcMove = dgvReferenceValue.Rows[i].Cells["AVG_DC"].Value == null
                            || dgvReferenceValue.Rows[i].Cells["AVG_DC"].Value.ToString().Trim() == ""
                            ? 0d : Convert.ToDouble(dgvReferenceValue.Rows[i].Cells["AVG_DC"].Value.ToString().Trim());

                        dRef_RacMove = dgvReferenceValue.Rows[i].Cells["AVG_AC"].Value == null
                            || dgvReferenceValue.Rows[i].Cells["AVG_AC"].Value.ToString().Trim() == ""
                            ? 0d : Convert.ToDouble(dgvReferenceValue.Rows[i].Cells["AVG_AC"].Value.ToString().Trim());

                        dRef_LMove = dgvReferenceValue.Rows[i].Cells["AVG_L"].Value == null
                            || dgvReferenceValue.Rows[i].Cells["AVG_L"].Value.ToString().Trim() == ""
                            ? 0d : Convert.ToDouble(dgvReferenceValue.Rows[i].Cells["AVG_L"].Value.ToString().Trim());

                        dRef_CMove = dgvReferenceValue.Rows[i].Cells["AVG_C"].Value == null
                            || dgvReferenceValue.Rows[i].Cells["AVG_C"].Value.ToString().Trim() == ""
                            ? 0d : Convert.ToDouble(dgvReferenceValue.Rows[i].Cells["AVG_C"].Value.ToString().Trim());

                        dRef_QMove = dgvReferenceValue.Rows[i].Cells["AVG_Q"].Value == null
                            || dgvReferenceValue.Rows[i].Cells["AVG_Q"].Value.ToString().Trim() == ""
                            ? 0d : Convert.ToDouble(dgvReferenceValue.Rows[i].Cells["AVG_Q"].Value.ToString().Trim());
                    }
                }

                if (dgvReferenceValue.RowCount > 0)
                {
                    string[] arrayData = new string[15];

                    arrayData[0] = dRef_RdcStop.ToString().Trim();
                    arrayData[1] = dRef_RdcUP.ToString().Trim();
                    arrayData[2] = dRef_RdcMove.ToString().Trim();
                    arrayData[3] = dRef_RacStop.ToString().Trim();
                    arrayData[4] = dRef_RacUP.ToString().Trim();
                    arrayData[5] = dRef_RacMove.ToString().Trim();
                    arrayData[6] = dRef_LStop.ToString().Trim();
                    arrayData[7] = dRef_LUP.ToString().Trim();
                    arrayData[8] = dRef_LMove.ToString().Trim();
                    arrayData[9] = dRef_CStop.ToString().Trim();
                    arrayData[10] = dRef_CUP.ToString().Trim();
                    arrayData[11] = dRef_CMove.ToString().Trim();
                    arrayData[12] = dRef_QStop.ToString().Trim();
                    arrayData[13] = dRef_QUP.ToString().Trim();
                    arrayData[14] = dRef_QMove.ToString().Trim();

                    if (m_db.GetRCSReferenceValueDataCount(strPlantName.Trim(), strSelectHogi.Trim(), strOh_Degree.Trim()) > 0)
                        boolResult = m_db.UpDateRCSReferenceValueDataInfo(strPlantName.Trim(), strSelectHogi.Trim(), strOh_Degree.Trim(), arrayData);
                    else
                        boolResult = m_db.InsertRCSReferenceValueDataInfo(strPlantName.Trim(), strSelectHogi.Trim(), strOh_Degree.Trim(), arrayData);
                }

                #endregion
            }
            else
            {
                #region ▣ DRPI 기준값 저장

                string[] arrayData = new string[5];
                string strCoilCabinetType = "", strCoilCabinetName = "", strCheckCoilCabinetType = "", strCheckCoilCabinetName = ""
                    , strSaveCoilCabinetName = "", strDRPIData = "";
                double dSumRef_Rdc = 0d, dSumRef_Rac = 0d, dSumRef_L = 0d, dSumRef_C = 0d, dSumRef_Q = 0d;
                double dDataCount = 0;

                for (int i = 0; i < dgvReferenceValue.RowCount; i++)
                {
                    strCoilCabinetName = dgvReferenceValue.Rows[i].Cells["DRPIGroup"].Value == null
                            ? "" : dgvReferenceValue.Rows[i].Cells["DRPIGroup"].Value.ToString().Trim();
                    strCoilCabinetType = dgvReferenceValue.Rows[i].Cells["DRPIType"].Value == null
                            ? "" : dgvReferenceValue.Rows[i].Cells["DRPIType"].Value.ToString().Trim();

                    string strCoilName = strCoilCabinetName.Trim() + (dgvReferenceValue.Rows[i].Cells["CoilName"].Value == null
                            ? "" : Regex.Replace(dgvReferenceValue.Rows[i].Cells["CoilName"].Value.ToString().Trim(), @"\D", ""));

                    if (strCoilName.Trim() == "A1" || strCoilName.Trim() == "B1" || strCoilName.Trim() == "A2" || strCoilName.Trim() == "B2")
                    {
                        if (strCoilCabinetName.Trim() != strCheckCoilCabinetName.Trim() || strCoilCabinetType.Trim() != strCheckCoilCabinetType.Trim())
                        {
                            strCheckCoilCabinetName = strCoilCabinetName.Trim();
                            strCheckCoilCabinetType = strCoilCabinetType.Trim();

                            if (strDRPIData.Trim() == "")
                                strDRPIData = strCoilCabinetName.Trim() + "," + strCoilCabinetType.Trim();
                            else
                                strDRPIData += "/" + strCoilCabinetName.Trim() + "," + strCoilCabinetType.Trim();
                        }

                        arrayData[0] = dgvReferenceValue.Rows[i].Cells["AVG_DC"].Value == null
                            || dgvReferenceValue.Rows[i].Cells["AVG_DC"].Value.ToString().Trim() == ""
                            ? "0.000" : dgvReferenceValue.Rows[i].Cells["AVG_DC"].Value.ToString().Trim();

                        arrayData[1] = dgvReferenceValue.Rows[i].Cells["AVG_AC"].Value == null
                            || dgvReferenceValue.Rows[i].Cells["AVG_AC"].Value.ToString().Trim() == ""
                            ? "0.000" : dgvReferenceValue.Rows[i].Cells["AVG_AC"].Value.ToString().Trim();

                        arrayData[2] = dgvReferenceValue.Rows[i].Cells["AVG_L"].Value == null
                            || dgvReferenceValue.Rows[i].Cells["AVG_L"].Value.ToString().Trim() == ""
                            ? "0.000" : dgvReferenceValue.Rows[i].Cells["AVG_L"].Value.ToString().Trim();

                        arrayData[3] = dgvReferenceValue.Rows[i].Cells["AVG_C"].Value == null
                            || dgvReferenceValue.Rows[i].Cells["AVG_C"].Value.ToString().Trim() == ""
                            ? "0.000000" : dgvReferenceValue.Rows[i].Cells["AVG_C"].Value.ToString().Trim();

                        arrayData[4] = dgvReferenceValue.Rows[i].Cells["AVG_Q"].Value == null
                            || dgvReferenceValue.Rows[i].Cells["AVG_Q"].Value.ToString().Trim() == ""
                            ? "0.000" : dgvReferenceValue.Rows[i].Cells["AVG_Q"].Value.ToString().Trim();

                        if (m_db.GetDRPIReferenceValueDataCount(strPlantName.Trim(), strSelectHogi.Trim(), strOh_Degree.Trim(), strCoilCabinetType, strCoilName) > 0)
                            boolResult = m_db.UpDateDRPIReferenceValueDataInfo(strPlantName.Trim(), strSelectHogi.Trim(), strOh_Degree.Trim(), strCoilCabinetType, strCoilName, arrayData);
                        else
                            boolResult = m_db.InsertDRPIReferenceValueDataInfo(strPlantName.Trim(), strSelectHogi.Trim(), strOh_Degree.Trim(), strCoilCabinetType, strCoilName, arrayData);
                    }
                }

                string[] arrayDRPIData = strDRPIData.Split('/');
                dDataCount = 0;
                strCheckCoilCabinetType = "";
                strCheckCoilCabinetName = "";
                strSaveCoilCabinetName = "";

                for (int j = 0; j < arrayDRPIData.Length; j++)
                {
                    string[] arrayDRPIDataInfo = arrayDRPIData[j].Split(',');
                    for (int i = 0; i < dgvReferenceValue.RowCount; i++)
                    {
                        strCoilCabinetName = dgvReferenceValue.Rows[i].Cells["DRPIGroup"].Value == null
                                ? "" : dgvReferenceValue.Rows[i].Cells["DRPIGroup"].Value.ToString().Trim();
                        strCoilCabinetType = dgvReferenceValue.Rows[i].Cells["DRPIType"].Value == null
                                ? "" : dgvReferenceValue.Rows[i].Cells["DRPIType"].Value.ToString().Trim();

                        if (arrayDRPIDataInfo[0].Trim() != strCoilCabinetName.Trim() || arrayDRPIDataInfo[1].Trim() != strCoilCabinetType.Trim()) continue;

                        string strCoilName = strCoilCabinetName.Trim() + (dgvReferenceValue.Rows[i].Cells["CoilName"].Value == null
                                ? "" : Regex.Replace(dgvReferenceValue.Rows[i].Cells["CoilName"].Value.ToString().Trim(), @"\D", ""));

                        if (strCoilName.Trim() == "A1" || strCoilName.Trim() == "B1" || strCoilName.Trim() == "A2" || strCoilName.Trim() == "B2") continue;

                        if (strCoilName.Trim() != "A1" && strCoilName.Trim() != "B1" && strCoilName.Trim() != "A2" && strCoilName.Trim() != "B2")
                        {
                            if (strCoilCabinetName.Trim() != strCheckCoilCabinetName.Trim() || strCoilCabinetType.Trim() != strCheckCoilCabinetType.Trim())
                            {
                                if (dDataCount > 0)
                                {
                                    strSaveCoilCabinetName = "";
                                    strSaveCoilCabinetName = Regex.Replace(strCheckCoilCabinetName.Trim(), @"\d", "") + (strCheckCoilCabinetType.Trim() == "정지용" ? "3~6,18~21" : "3 ~ 21");

                                    arrayData[0] = (dSumRef_Rdc / dDataCount).ToString("F3").Trim();
                                    arrayData[1] = (dSumRef_Rac / dDataCount).ToString("F3").Trim();
                                    arrayData[2] = (dSumRef_L / dDataCount).ToString("F3").Trim();
                                    arrayData[3] = (dSumRef_C / dDataCount).ToString("F6").Trim();
                                    arrayData[4] = (dSumRef_Q / dDataCount).ToString("F3").Trim();

                                    if (m_db.GetDRPIReferenceValueDataCount(strPlantName.Trim(), strSelectHogi.Trim(), strOh_Degree.Trim(), strCheckCoilCabinetType, strSaveCoilCabinetName) > 0)
                                        boolResult = m_db.UpDateDRPIReferenceValueDataInfo(strPlantName.Trim(), strSelectHogi.Trim(), strOh_Degree.Trim(), strCheckCoilCabinetType, strSaveCoilCabinetName, arrayData);
                                    else
                                        boolResult = m_db.InsertDRPIReferenceValueDataInfo(strPlantName.Trim(), strSelectHogi.Trim(), strOh_Degree.Trim(), strCheckCoilCabinetType, strSaveCoilCabinetName, arrayData);

                                    dDataCount = 0;
                                    dSumRef_Rdc = 0d;
                                    dSumRef_Rac = 0d;
                                    dSumRef_L = 0d;
                                    dSumRef_C = 0d;
                                    dSumRef_Q = 0d;
                                }
                            }

                            strCheckCoilCabinetName = strCoilCabinetName.Trim();
                            strCheckCoilCabinetType = strCoilCabinetType.Trim();

                            dSumRef_Rdc += dgvReferenceValue.Rows[i].Cells["AVG_DC"].Value == null
                                || dgvReferenceValue.Rows[i].Cells["AVG_DC"].Value.ToString().Trim() == ""
                                ? 0d : Convert.ToDouble(dgvReferenceValue.Rows[i].Cells["AVG_DC"].Value.ToString().Trim());

                            dSumRef_Rac += dgvReferenceValue.Rows[i].Cells["AVG_AC"].Value == null
                                || dgvReferenceValue.Rows[i].Cells["AVG_AC"].Value.ToString().Trim() == ""
                                ? 0d : Convert.ToDouble(dgvReferenceValue.Rows[i].Cells["AVG_AC"].Value.ToString().Trim());

                            dSumRef_L += dgvReferenceValue.Rows[i].Cells["AVG_L"].Value == null
                                || dgvReferenceValue.Rows[i].Cells["AVG_L"].Value.ToString().Trim() == ""
                                ? 0d : Convert.ToDouble(dgvReferenceValue.Rows[i].Cells["AVG_L"].Value.ToString().Trim());

                            dSumRef_C += dgvReferenceValue.Rows[i].Cells["AVG_C"].Value == null
                                || dgvReferenceValue.Rows[i].Cells["AVG_C"].Value.ToString().Trim() == ""
                                ? 0d : Convert.ToDouble(dgvReferenceValue.Rows[i].Cells["AVG_C"].Value.ToString().Trim());

                            dSumRef_Q += dgvReferenceValue.Rows[i].Cells["AVG_Q"].Value == null
                                || dgvReferenceValue.Rows[i].Cells["AVG_Q"].Value.ToString().Trim() == ""
                                ? 0d : Convert.ToDouble(dgvReferenceValue.Rows[i].Cells["AVG_Q"].Value.ToString().Trim());

                            dDataCount++;
                        }
                    }
                }

                if (dDataCount > 0)
                {
                    strSaveCoilCabinetName = "";
                    strSaveCoilCabinetName = Regex.Replace(strCheckCoilCabinetName.Trim(), @"\d", "") + (strCheckCoilCabinetType.Trim() == "정지용" ? "3~6,18~21" : "3 ~ 21");

                    arrayData[0] = (dSumRef_Rdc / dDataCount).ToString("F3").Trim();
                    arrayData[1] = (dSumRef_Rac / dDataCount).ToString("F3").Trim();
                    arrayData[2] = (dSumRef_L / dDataCount).ToString("F3").Trim();
                    arrayData[3] = (dSumRef_C / dDataCount).ToString("F6").Trim();
                    arrayData[4] = (dSumRef_Q / dDataCount).ToString("F3").Trim();

                    if (m_db.GetDRPIReferenceValueDataCount(strPlantName.Trim(), strSelectHogi.Trim(), strOh_Degree.Trim(), strCheckCoilCabinetType, strSaveCoilCabinetName) > 0)
                        boolResult = m_db.UpDateDRPIReferenceValueDataInfo(strPlantName.Trim(), strSelectHogi.Trim(), strOh_Degree.Trim(), strCheckCoilCabinetType, strSaveCoilCabinetName, arrayData);
                    else
                        boolResult = m_db.InsertDRPIReferenceValueDataInfo(strPlantName.Trim(), strSelectHogi.Trim(), strOh_Degree.Trim(), strCheckCoilCabinetType, strSaveCoilCabinetName, arrayData);

                    dDataCount = 0;
                    dSumRef_Rdc = 0d;
                    dSumRef_Rac = 0d;
                    dSumRef_L = 0d;
                    dSumRef_C = 0d;
                    dSumRef_Q = 0d;
                }
                
                #endregion
            }

            return boolResult;
        }

        /// <summary>
        /// 가져오기 Button Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSearch_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            progressBar1.Visible = true;

            try
            {
                dgvAccessData.Rows.Clear();
                dgvReferenceValue.Rows.Clear();

                if (teAccessDBFileName.Text.Trim() == "")
                {
                    frmMB.lblMessage.Text = "Access DB를 선택하십시오.";
                    frmMB.TopMost = true;
                    frmMB.ShowDialog();

                    return;
                }

                // Header Data 가져오기
                int intDataCount = 0;
                bool boolResult = false;

                boolResult = GetSelectFileData(ref intDataCount);

                if (chkReferenceValue.Checked)
                    GetSelectReferenceValueFileData();

                frmMB.lblMessage.Text = "Access DB에서 데이터 가져오기 완료";
                frmMB.TopMost = true;
                frmMB.ShowDialog();
            }
            catch (Exception ex)
            {
                frmMB.lblMessage.Text = "AAccess DB에서 데이터 가져오기 실패";
                frmMB.TopMost = true;
                frmMB.ShowDialog();
                System.Diagnostics.Debug.Print(ex.Message);
            }

            System.Windows.Forms.Cursor.Current = Cursors.Default;
        }

        /// <summary>
        /// 선택 File Data 가져오기
        /// </summary>
        private bool GetSelectFileData(ref int _intDataCount)
        {
            bool boolResult = false;

            try
            {
                string strFielder = teAccessDBFileName.Text.Trim();
                string strCommonDBFile = strFielder.Trim() + @"\YG1_CMN.mdb";
                string strRCSDBFile = @"\YG1_RCS.mdb";
                string strDRPIFile = @"\YG1_DRPI.mdb";
                string strAccessDBFileName = "";

                if (rboDRPI.Checked)
                    strAccessDBFileName = strFielder.Trim() + strDRPIFile.Trim();
                else
                    strAccessDBFileName = strFielder.Trim() + strRCSDBFile.Trim();

                // DB 파일 체크
                System.IO.FileInfo fi1 = new System.IO.FileInfo(strAccessDBFileName);
                System.IO.FileInfo fi2 = new System.IO.FileInfo(strCommonDBFile);

                if (!fi1.Exists || !fi2.Exists)
                {
                    frmMB.lblMessage.Text = "기존 데이터(.mdb) 파일이 없습니다.";
                    frmMB.TopMost = true;
                    frmMB.ShowDialog();
                    return false;
                }

                string strQuery = "";
                string strTableName_Take = "[" + strAccessDBFileName.Trim() + "].[TB_TAKE] AS TB_TAKE";
                string strTableName_TakeVal = "[" + strAccessDBFileName.Trim() + "].[TB_TAKE_VAL] AS TB_TAKE_VAL";
                string strTableName_MasterATM = "[" + strCommonDBFile.Trim() + "].[TB_MASTER_ATM] AS TB_MASTER_ATM";
                string strTableName_CodeCTRLGroup = "[" + strAccessDBFileName.Trim() + "].[TB_CODE_CTRL_GROUP] AS TB_CODE_CTRL_GROUP";
                string strTableName_CodeControl = "[" + strAccessDBFileName.Trim() + "].[TB_CODE_CONTROL] AS TB_CODE_CONTROL";
                string strTableName_CoilCode = "[" + strAccessDBFileName.Trim() + "].[TB_CODE_COIL] AS TB_CODE_COIL";
                string strTableName_DRPICoilCode = "[" + strAccessDBFileName.Trim() + "].[TB_DRPI_CODE_COIL] AS TB_DRPI_CODE_COIL";
                int intOh_Degree = nmOh_Degree.Value == null ? 0 : Convert.ToInt32(nmOh_Degree.Value);
                string strSelectHogi = nmHogi.Value == null ? "1 호기" : Convert.ToInt32(nmHogi.Value).ToString().Trim() + " 호기";

                strFielder = string.Format("TB_TAKE.*, TB_TAKE_VAL.*", strTableName_Take, strTableName_TakeVal, strTableName_MasterATM);
                strFielder = string.Format("{0}, (SELECT TB_MASTER_ATM.ATM_NAME FROM {2} WHERE TB_MASTER_ATM.ATM_ID = TB_TAKE.ATM_ID) AS ATM_NAME"
                    , strFielder, strTableName_Take, strTableName_MasterATM);
                strFielder = string.Format("{0}, (SELECT TB_CODE_CONTROL.CTRL_CODE FROM {1} WHERE TB_CODE_CONTROL.CTRL_ID = TB_TAKE.CTRL_ID) AS CTRL_CODE"
                    , strFielder, strTableName_CodeControl);

                // 코일명 
                if (rboDRPI.Checked)
                {
                    strFielder = string.Format("{0}, '' AS GRP_NAME", strFielder);
                    strFielder = string.Format("{0}, (SELECT TB_DRPI_CODE_COIL.COIL_CODE FROM {1} WHERE TB_DRPI_CODE_COIL.COIL_ID = TB_TAKE_VAL.COIL_ID) AS COIL_CODE"
                        , strFielder, strTableName_DRPICoilCode);
                }
                else
                {
                    strFielder = string.Format("{0}, (SELECT TB_CODE_CTRL_GROUP.GRP_NAME FROM {1} LEFT OUTER JOIN {2} ON TB_CODE_CTRL_GROUP.CTRL_GRP = TB_CODE_CONTROL.CTRL_GRP WHERE TB_CODE_CONTROL.CTRL_ID = TB_TAKE.CTRL_ID) AS GRP_NAME"
                        , strFielder, strTableName_CodeControl, strTableName_CodeCTRLGroup);
                    strFielder = string.Format("{0}, (SELECT TB_CODE_COIL.COIL_CODE FROM {1} WHERE TB_CODE_COIL.COIL_ID = TB_TAKE_VAL.COIL_ID) AS COIL_CODE"
                        , strFielder, strTableName_CoilCode);
                }

                strQuery = string.Format("SELECT {0} ", strFielder);
                strQuery = string.Format("{0} FROM {1}", strQuery, strTableName_Take);
                strQuery = string.Format("{0} LEFT OUTER JOIN {1}", strQuery, strTableName_TakeVal);
                strQuery = string.Format("{0} ON TB_TAKE_VAL.TAKE_ID = TB_TAKE.TAKE_ID", strQuery);

                if (rboDRPI.Checked)
                    strQuery = string.Format("{0} WHERE TB_TAKE.OH_ID = {2} AND TB_TAKE.DRPI_TYPE <> 0"
                        , strQuery, strTableName_Take, intOh_Degree);
                else
                    strQuery = string.Format("{0} WHERE TB_TAKE.OH_ID = {2} AND TB_TAKE.DRPI_TYPE = 0"
                        , strQuery, strTableName_Take, intOh_Degree);

                strQuery = string.Format("{0} ORDER BY TB_TAKE.CTRL_ID, TB_TAKE.TAKE_ID, TB_TAKE_VAL.COIL_ID, TB_TAKE_VAL.TAKE_COUNT, TB_TAKE.TAKE_TIME, TB_TAKE_VAL.TAKE_TIME ;"
                    , strQuery, strTableName_Take, strTableName_TakeVal);

                DataTable dt = new DataTable();

                dt = m_Accessdb.SelectOleDb(strAccessDBFileName.Trim(), strQuery);

                if (dt == null || dt.Rows.Count <= 0)
                {
                    _intDataCount = 0;
                }
                else
                {
                    _intDataCount = dt.Rows.Count;
                    progressBar1.Value = 0;
                    
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        if (string.Format("{0} 호기", Regex.Replace(dt.Rows[i]["ATM_NAME"].ToString().Trim(), @"\D", "")) != strSelectHogi.Trim()
                            || dt.Rows[i]["TB_TAKE_VAL.TAKE_ID"] == null || dt.Rows[i]["TB_TAKE_VAL.TAKE_ID"].ToString().Trim() == "")
                        {
                            progressBar1.Value = ((i * 100) / _intDataCount) / 2;
                            continue;
                        }

                        int intRowIndex = dgvAccessData.Rows.Add();

                        dgvAccessData.Rows[intRowIndex].Cells["Select"].Value = false;
                        dgvAccessData.Rows[intRowIndex].Cells["UpLoadText"].Value = "";

                        dgvAccessData.Rows[intRowIndex].Cells["PlantName"].Value = strPlantName.Trim();

                        // 호기 
                        dgvAccessData.Rows[intRowIndex].Cells["Hogi"].Value = dt.Rows[i]["ATM_NAME"] == null ? "" : string.Format("{0} 호기", Regex.Replace(dt.Rows[i]["ATM_NAME"].ToString().Trim(), @"\D", ""));

                        dgvAccessData.Rows[intRowIndex].Cells["Oh_Degree"].Value = dt.Rows[i]["OH_ID"] == null ? "" : string.Format("제 {0} 차", dt.Rows[i]["OH_ID"].ToString().Trim());

                        if (rboDRPI.Checked)
                        {
                            if (dt.Rows[i]["DRPI_TYPE"].ToString().Trim() == "1")
                                dgvAccessData.Rows[intRowIndex].Cells["DRPIGroup"].Value = "A";
                            else
                                dgvAccessData.Rows[intRowIndex].Cells["DRPIGroup"].Value = "B";

                            if (dt.Rows[i]["COIL_TYPE"].ToString().Trim() == "1")
                                dgvAccessData.Rows[intRowIndex].Cells["DRPIType"].Value = "제어용";
                            else
                                dgvAccessData.Rows[intRowIndex].Cells["DRPIType"].Value = "정지용";
                        }
                        else
                        {
                            dgvAccessData.Rows[intRowIndex].Cells["DRPIGroup"].Value = "";
                            dgvAccessData.Rows[intRowIndex].Cells["DRPIType"].Value = "";
                            dgvAccessData.Rows[intRowIndex].Cells["PowerCabinet"].Value = dt.Rows[i]["GRP_NAME"] == null ? "" : dt.Rows[i]["GRP_NAME"].ToString().Trim().Replace("전력함", "");
                        }

                        dgvAccessData.Rows[intRowIndex].Cells["ControlSeqNumber"].Value = dt.Rows[i]["CTRL_ID"] == null ? "" : dt.Rows[i]["CTRL_ID"].ToString().Trim();
                        dgvAccessData.Rows[intRowIndex].Cells["ControlName"].Value = dt.Rows[i]["CTRL_CODE"] == null ? "" : dt.Rows[i]["CTRL_CODE"].ToString().Trim();
                        dgvAccessData.Rows[intRowIndex].Cells["Temperature_ReferenceValue"].Value = dt.Rows[i]["TEMPER"] == null ? "" : dt.Rows[i]["TEMPER"].ToString().Trim();
                        dgvAccessData.Rows[intRowIndex].Cells["TemperatureUpDown_ReferenceValue"].Value = Gini.GetValue("ReferenceValue", "TemperatureUpDown_ReferenceValue").Trim();
                        dgvAccessData.Rows[intRowIndex].Cells["Frequency"].Value = dt.Rows[i]["BAND"] == null ? "120 Hz" : dt.Rows[i]["BAND"].ToString().Trim() + " Hz";
                        dgvAccessData.Rows[intRowIndex].Cells["VoltageLevel"].Value = dt.Rows[i]["VOLT"] == null ? "1000mV" : dt.Rows[i]["VOLT"].ToString().Trim() + "mV";
                        dgvAccessData.Rows[intRowIndex].Cells["MeasurementCount"].Value = "1";
                        dgvAccessData.Rows[intRowIndex].Cells["SeqNumber"].Value = dt.Rows[i]["COIL_ID"] == null ? "" : dt.Rows[i]["COIL_ID"].ToString().Trim();
                        
                        if (rboRCS.Checked)
                        {
                            dgvAccessData.Rows[intRowIndex].Cells["CoilName"].Value = dt.Rows[i]["COIL_CODE"] == null ? "" : dt.Rows[i]["COIL_CODE"].ToString().Trim().Replace("권선", "");
                        }
                        else
                        {
                            if (dt.Rows[i]["COIL_CODE"].ToString().Trim() == "A7" && dt.Rows[i]["COIL_TYPE"].ToString().Trim() == "2")
                            {
                                if (dt.Rows[i]["DRPI_TYPE"].ToString().Trim() == "1")
                                    dgvAccessData.Rows[intRowIndex].Cells["CoilName"].Value = "A18";
                                else
                                    dgvAccessData.Rows[intRowIndex].Cells["CoilName"].Value = "B18";
                            }
                            else if (dt.Rows[i]["COIL_CODE"].ToString().Trim() == "A8" && dt.Rows[i]["COIL_TYPE"].ToString().Trim() == "2")
                            {
                                if (dt.Rows[i]["DRPI_TYPE"].ToString().Trim() == "1")
                                    dgvAccessData.Rows[intRowIndex].Cells["CoilName"].Value = "A19";
                                else
                                    dgvAccessData.Rows[intRowIndex].Cells["CoilName"].Value = "B19";
                            }
                            else if (dt.Rows[i]["COIL_CODE"].ToString().Trim() == "A9" && dt.Rows[i]["COIL_TYPE"].ToString().Trim() == "2")
                            {
                                if (dt.Rows[i]["DRPI_TYPE"].ToString().Trim() == "1")
                                    dgvAccessData.Rows[intRowIndex].Cells["CoilName"].Value = "A20";
                                else
                                    dgvAccessData.Rows[intRowIndex].Cells["CoilName"].Value = "B20";
                            }
                            else if (dt.Rows[i]["COIL_CODE"].ToString().Trim() == "A10" && dt.Rows[i]["COIL_TYPE"].ToString().Trim() == "2")
                            {
                                if (dt.Rows[i]["DRPI_TYPE"].ToString().Trim() == "1")
                                    dgvAccessData.Rows[intRowIndex].Cells["CoilName"].Value = "A21";
                                else
                                    dgvAccessData.Rows[intRowIndex].Cells["CoilName"].Value = "B21";
                            }
                            else
                            {
                                if (dt.Rows[i]["DRPI_TYPE"].ToString().Trim() == "1")
                                    dgvAccessData.Rows[intRowIndex].Cells["CoilName"].Value = dt.Rows[i]["COIL_CODE"] == null ? "" : ("A" + Regex.Replace(dt.Rows[i]["COIL_CODE"].ToString().Trim(), @"\D", ""));
                                else
                                    dgvAccessData.Rows[intRowIndex].Cells["CoilName"].Value = dt.Rows[i]["COIL_CODE"] == null ? "" : ("B" + Regex.Replace(dt.Rows[i]["COIL_CODE"].ToString().Trim(), @"\D", ""));
                            }
                        }

                        dgvAccessData.Rows[intRowIndex].Cells["CoilNumber"].Value = dt.Rows[i]["TAKE_COUNT"] == null ? "" : dt.Rows[i]["TAKE_COUNT"].ToString().Trim();

                        string strDC_VAL = Regex.Replace(dt.Rows[i]["DC_VAL"].ToString(), @"\d", "");
                        string strAC_VAL = Regex.Replace(dt.Rows[i]["AC_VAL"].ToString(), @"\d", "");
                        string strIND_VAL = Regex.Replace(dt.Rows[i]["IND_VAL"].ToString(), @"\d", "");
                        string strC_VAL = Regex.Replace(dt.Rows[i]["DC_VAL"].ToString(), @"\d", "");
                        string strQ_VAL = Regex.Replace(dt.Rows[i]["Q_VAL"].ToString(), @"\d", ""); 

                        double dDC_VAL = dt.Rows[i]["DC_VAL"] == null || dt.Rows[i]["DC_VAL"].ToString().Trim() == ""
                            ? 0d : (strDC_VAL.Contains("E") == true ? 0d : Convert.ToDouble(dt.Rows[i]["DC_VAL"].ToString().Trim()));
                        double dAC_VAL = dt.Rows[i]["AC_VAL"] == null || dt.Rows[i]["AC_VAL"].ToString().Trim() == ""
                            ? 0d : (strAC_VAL.Contains("E") == true ? 0d : Convert.ToDouble(dt.Rows[i]["AC_VAL"].ToString().Trim()));
                        double dIND_VAL = dt.Rows[i]["IND_VAL"] == null || dt.Rows[i]["IND_VAL"].ToString().Trim() == ""
                            ? 0d : (strIND_VAL.Contains("E") == true ? 0d : Convert.ToDouble(dt.Rows[i]["IND_VAL"].ToString().Trim()));
                        double dC_VAL = dt.Rows[i]["C_VAL"] == null || dt.Rows[i]["C_VAL"].ToString().Trim() == ""
                            ? 0d : (strC_VAL.Contains("E") == true ? 0d : Convert.ToDouble(dt.Rows[i]["C_VAL"].ToString().Trim()));
                        double dQ_VAL = dt.Rows[i]["Q_VAL"] == null || dt.Rows[i]["Q_VAL"].ToString().Trim() == ""
                            ? 0d : (strQ_VAL.Contains("E") == true ? 0d : Convert.ToDouble(dt.Rows[i]["Q_VAL"].ToString().Trim()));

                        dgvAccessData.Rows[intRowIndex].Cells["DC_ResistanceValue"].Value = dDC_VAL.ToString("F3").Trim();
                        dgvAccessData.Rows[intRowIndex].Cells["DC_Deviation"].Value = "0.000";
                        dgvAccessData.Rows[intRowIndex].Cells["Item_Rdc"].Value = dDC_VAL == 0 ? "1" : "0";
                        dgvAccessData.Rows[intRowIndex].Cells["AC_ResistanceValue"].Value = dAC_VAL.ToString("F3").Trim();
                        dgvAccessData.Rows[intRowIndex].Cells["AC_Deviation"].Value = "0.000";
                        dgvAccessData.Rows[intRowIndex].Cells["Item_Rac"].Value = dAC_VAL == 0 ? "1" : "0";
                        dgvAccessData.Rows[intRowIndex].Cells["L_InductanceValue"].Value = dIND_VAL.ToString("F3").Trim();
                        dgvAccessData.Rows[intRowIndex].Cells["L_Deviation"].Value = "0.000";
                        dgvAccessData.Rows[intRowIndex].Cells["Item_L"].Value = dIND_VAL == 0 ? "1" : "0";
                        dgvAccessData.Rows[intRowIndex].Cells["C_CapacitanceValue"].Value = dC_VAL.ToString("F6").Trim();
                        dgvAccessData.Rows[intRowIndex].Cells["C_Deviation"].Value = "0.000000";
                        dgvAccessData.Rows[intRowIndex].Cells["Item_C"].Value = dC_VAL == 0 ? "1" : "0";
                        dgvAccessData.Rows[intRowIndex].Cells["Q_FactorValue"].Value = dQ_VAL.ToString("F3").Trim();
                        dgvAccessData.Rows[intRowIndex].Cells["Q_Deviation"].Value = "0.0000";
                        dgvAccessData.Rows[intRowIndex].Cells["Item_Q"].Value = dQ_VAL == 0 ? "1" : "0";
                        dgvAccessData.Rows[intRowIndex].Cells["MeasurementMode"].Value = "일반모드";
                        dgvAccessData.Rows[intRowIndex].Cells["HeaderDate"].Value = dt.Rows[i]["TB_TAKE.TAKE_TIME"] == null ? "" : dt.Rows[i]["TB_TAKE.TAKE_TIME"].ToString().Trim();
                        dgvAccessData.Rows[intRowIndex].Cells["DetailDate"].Value = dt.Rows[i]["TB_TAKE_VAL.TAKE_TIME"] == null ? "" : dt.Rows[i]["TB_TAKE_VAL.TAKE_TIME"].ToString().Trim();
                        dgvAccessData.Rows[intRowIndex].Cells["MeasurementResult"].Value = "";
                        dgvAccessData.Rows[intRowIndex].Cells["TAKE_ID"].Value = dt.Rows[i]["TB_TAKE_VAL.TAKE_ID"] == null ? "" : dt.Rows[i]["TB_TAKE_VAL.TAKE_ID"].ToString().Trim();
                        
                        progressBar1.Value = ((i * 100) / _intDataCount) / 2;
                    }
                }

                progressBar1.Value = 50;

                int intMeasurementCount = 0, iniHeaderDataCount = 0, iniDetailDataCount = 0, intCoilNumber = 0;
                string strHogi = "", strOh_Degree = "", strDRPIGroup = "", strDRPIType = "", strControlName = "", strPowerCabinet = "", strCoilName = ""
                    , strDataCheckMessage = "", strTake_ID = "", strTake_ID_Check = "";

                for (int i = 0; i < dgvAccessData.RowCount; i++)
                {
                    if (dgvAccessData.Rows[i].Cells["Hogi"].Value.ToString().Trim() != strSelectHogi.Trim())
                    {
                        progressBar1.Value = 50 + (((i * 100) / _intDataCount) / 2);
                        continue;
                    }

                    if (rboDRPI.Checked)
                    {
                        if (dgvAccessData.Rows[i].Cells["TAKE_ID"].Value.ToString().Trim() != strTake_ID.Trim()
                            || dgvAccessData.Rows[i].Cells["CoilName"].Value.ToString().Trim() != strCoilName.Trim())
                        {
                            strDataCheckMessage = "";

                            strTake_ID = dgvAccessData.Rows[i].Cells["TAKE_ID"].Value.ToString().Trim();
                            strCoilName = dgvAccessData.Rows[i].Cells["CoilName"].Value.ToString().Trim();

                            intMeasurementCount = 0;
                            for (int j = 0; j < dgvAccessData.RowCount; j++)
                            {
                                if (dgvAccessData.Rows[j].Cells["TAKE_ID"].Value.ToString().Trim() == strTake_ID.Trim()
                                    && dgvAccessData.Rows[j].Cells["CoilName"].Value.ToString().Trim() == strCoilName.Trim())
                                {
                                    intCoilNumber = dgvAccessData.Rows[j].Cells["CoilNumber"].Value == null || dgvAccessData.Rows[j].Cells["CoilNumber"].Value.ToString().Trim() == ""
                                        ? 0 : Convert.ToInt32(dgvAccessData.Rows[j].Cells["CoilNumber"].Value.ToString().Trim());

                                    if (intMeasurementCount < intCoilNumber)
                                        intMeasurementCount = intCoilNumber;
                                }
                            }

                            strHogi = dgvAccessData.Rows[i].Cells["Hogi"].Value == null ? "" : dgvAccessData.Rows[i].Cells["Hogi"].Value.ToString().Trim();
                            strOh_Degree = dgvAccessData.Rows[i].Cells["Oh_Degree"].Value == null ? "" : dgvAccessData.Rows[i].Cells["Oh_Degree"].Value.ToString().Trim();
                            strDRPIGroup = dgvAccessData.Rows[i].Cells["DRPIGroup"].Value == null ? "" : dgvAccessData.Rows[i].Cells["DRPIGroup"].Value.ToString().Trim();
                            strDRPIType = dgvAccessData.Rows[i].Cells["DRPIType"].Value == null ? "" : dgvAccessData.Rows[i].Cells["DRPIType"].Value.ToString().Trim();
                            strControlName = dgvAccessData.Rows[i].Cells["ControlName"].Value == null ? "" : dgvAccessData.Rows[i].Cells["ControlName"].Value.ToString().Trim();
                            strCoilName = dgvAccessData.Rows[i].Cells["CoilName"].Value == null ? "" : dgvAccessData.Rows[i].Cells["CoilName"].Value.ToString().Trim();
                            intCoilNumber = dgvAccessData.Rows[i].Cells["CoilNumber"].Value == null || dgvAccessData.Rows[i].Cells["CoilNumber"].Value.ToString().Trim() == ""
                                ? 0 : Convert.ToInt32(dgvAccessData.Rows[i].Cells["CoilNumber"].Value.ToString().Trim());

                            // Header 기존 데이터 체크
                            iniHeaderDataCount = m_db.GetDRPIDiagnosisHeaderDataGridViewDataCount(strPlantName.Trim(), strHogi.Trim(), strOh_Degree.Trim(), strDRPIGroup.Trim(), strDRPIType.Trim(), strControlName.Trim());

                            if (iniHeaderDataCount > 0)
                                strDataCheckMessage = string.Format("Header Table에 이미 데이터가 있습니다.");

                            // Detail 기준 데이터 체크 
                            iniDetailDataCount = m_db.GetDRPIDiagnosisDetailDataGridViewDataCount(strPlantName.Trim(), strHogi.Trim(), strOh_Degree.Trim(), strDRPIGroup.Trim(), strDRPIType.Trim(), strControlName.Trim(), strCoilName.Trim(), intCoilNumber);

                            if (iniDetailDataCount > 0 && iniHeaderDataCount > 0)
                                strDataCheckMessage = string.Format("Header와 Detail Table에 이미 데이터가 있습니다.");
                        }
                    }
                    else
                    {
                        strDataCheckMessage = "";

                        if (dgvAccessData.Rows[i].Cells["TAKE_ID"].Value.ToString().Trim() != strTake_ID.Trim()
                            || dgvAccessData.Rows[i].Cells["CoilName"].Value.ToString().Trim() != strCoilName.Trim())
                        {
                            strTake_ID = dgvAccessData.Rows[i].Cells["TAKE_ID"].Value.ToString().Trim();
                            strCoilName = dgvAccessData.Rows[i].Cells["CoilName"].Value.ToString().Trim();

                            intMeasurementCount = 0;
                            for (int j = 0; j < dgvAccessData.RowCount; j++)
                            {
                                if (dgvAccessData.Rows[j].Cells["TAKE_ID"].Value.ToString().Trim() == strTake_ID.Trim()
                                    && dgvAccessData.Rows[j].Cells["CoilName"].Value.ToString().Trim() == strCoilName.Trim())
                                {
                                    intCoilNumber = dgvAccessData.Rows[j].Cells["CoilNumber"].Value == null || dgvAccessData.Rows[j].Cells["CoilNumber"].Value.ToString().Trim() == ""
                                        ? 0 : Convert.ToInt32(dgvAccessData.Rows[j].Cells["CoilNumber"].Value.ToString().Trim());

                                    if (intMeasurementCount < intCoilNumber)
                                        intMeasurementCount = intCoilNumber;
                                }
                            }
                        }

                        strHogi = dgvAccessData.Rows[i].Cells["Hogi"].Value == null ? "" : dgvAccessData.Rows[i].Cells["Hogi"].Value.ToString().Trim();
                        strOh_Degree = dgvAccessData.Rows[i].Cells["Oh_Degree"].Value == null ? "" : dgvAccessData.Rows[i].Cells["Oh_Degree"].Value.ToString().Trim();
                        strPowerCabinet = dgvAccessData.Rows[i].Cells["PowerCabinet"].Value == null ? "" : dgvAccessData.Rows[i].Cells["PowerCabinet"].Value.ToString().Trim();
                        strControlName = dgvAccessData.Rows[i].Cells["ControlName"].Value == null ? "" : dgvAccessData.Rows[i].Cells["ControlName"].Value.ToString().Trim();
                        strCoilName = dgvAccessData.Rows[i].Cells["CoilName"].Value == null ? "" : dgvAccessData.Rows[i].Cells["CoilName"].Value.ToString().Trim();
                        intCoilNumber = dgvAccessData.Rows[i].Cells["CoilNumber"].Value == null || dgvAccessData.Rows[i].Cells["CoilNumber"].Value.ToString().Trim() == ""
                            ? 0 : Convert.ToInt32(dgvAccessData.Rows[i].Cells["CoilNumber"].Value.ToString().Trim());

                        // Header 기존 데이터 체크
                        iniHeaderDataCount = m_db.GetRCSDiagnosisHeaderDataGridViewDataCount(strPlantName.Trim(), strHogi.Trim(), strOh_Degree.Trim(), strPowerCabinet.Trim(), strControlName.Trim());

                        if (iniHeaderDataCount > 0)
                            strDataCheckMessage = string.Format("Header Table에 이미 데이터가 있습니다.");

                        // Detail 기준 데이터 체크 
                        iniDetailDataCount = m_db.GetRCSDiagnosisDetailDataGridViewDataCount(strPlantName.Trim(), strHogi.Trim(), strOh_Degree.Trim(), strPowerCabinet.Trim(), strControlName.Trim(), strCoilName.Trim(), intCoilNumber);

                        if (iniDetailDataCount > 0 && iniHeaderDataCount > 0)
                            strDataCheckMessage = string.Format("Header와 Detail Table에 이미 데이터가 있습니다.");
                    };

                    int iniCheckDataCount = 0;
                    strTake_ID_Check = "";

                    for (int ii = 0; ii < dgvAccessData.RowCount; ii++)
                    {
                        if (dgvAccessData.Rows[ii].Cells["TAKE_ID"].Value.ToString().Trim() == strTake_ID.Trim()) continue;

                        if (rboRCS.Checked)
                        {
                            if (dgvAccessData.Rows[ii].Cells["TAKE_ID"].Value.ToString().Trim() != strTake_ID_Check.Trim()
                                && dgvAccessData.Rows[ii].Cells["Hogi"].Value.ToString().Trim() == strHogi.Trim()
                                && dgvAccessData.Rows[ii].Cells["PowerCabinet"].Value.ToString().Trim() == strPowerCabinet.Trim()
                                && dgvAccessData.Rows[ii].Cells["ControlName"].Value.ToString().Trim() == strControlName.Trim()
                                && dgvAccessData.Rows[ii].Cells["CoilName"].Value.ToString().Trim() == strCoilName.Trim()
                                && dgvAccessData.Rows[ii].Cells["CoilNumber"].Value.ToString().Trim() == intCoilNumber.ToString().Trim())
                            {
                                strTake_ID_Check = dgvAccessData.Rows[ii].Cells["TAKE_ID"].Value.ToString().Trim();
                                iniCheckDataCount++;
                            }
                        }
                        else
                        {
                            if (dgvAccessData.Rows[ii].Cells["TAKE_ID"].Value.ToString().Trim() != strTake_ID_Check.Trim()
                                && dgvAccessData.Rows[ii].Cells["Hogi"].Value.ToString().Trim() == strHogi.Trim()
                                && dgvAccessData.Rows[ii].Cells["Oh_Degree"].Value.ToString().Trim() == strOh_Degree.Trim()
                                && dgvAccessData.Rows[ii].Cells["DRPIGroup"].Value.ToString().Trim() == strDRPIGroup.Trim()
                                && dgvAccessData.Rows[ii].Cells["DRPITYPE"].Value.ToString().Trim() == strDRPIType.Trim()
                                && dgvAccessData.Rows[ii].Cells["ControlName"].Value.ToString().Trim() == strControlName.Trim()
                                && dgvAccessData.Rows[ii].Cells["CoilName"].Value.ToString().Trim() == strCoilName.Trim()
                                && dgvAccessData.Rows[ii].Cells["CoilNumber"].Value.ToString().Trim() == intCoilNumber.ToString().Trim())
                            {
                                strTake_ID_Check = dgvAccessData.Rows[ii].Cells["TAKE_ID"].Value.ToString().Trim();
                                iniCheckDataCount++;
                            }
                        }
                    }

                    string strMessage = "";

                    if (iniCheckDataCount > 0 && strDataCheckMessage.Trim() != "")
                        strMessage = string.Format("중복 및 {0}", strDataCheckMessage.Trim());
                    else if (iniCheckDataCount > 0)
                        strMessage = string.Format("중복 데이터가 있습니다.");
                    else
                        strMessage = strDataCheckMessage.Trim();


                    dgvAccessData.Rows[i].Cells["UpLoadText"].Value = strMessage.Trim() == "" ? "신규" : strMessage.Trim();
                    dgvAccessData.Rows[i].Cells["MeasurementCount"].Value = intMeasurementCount.ToString().Trim();

                    progressBar1.Value = 50 + (((i * 100) / _intDataCount) / 2);
                }

                progressBar1.Value = 100;
                boolResult = true;
            }
            catch (Exception ex)
            {
                boolResult = false;
                System.Diagnostics.Debug.Print(ex.Message);
            }
            finally
            {
                progressBar1.Visible = false;
            }

            return boolResult;
        }

        /// <summary>
        /// 기준값 선택 File Data 가져오기
        /// </summary>
        private void GetSelectReferenceValueFileData()
        {
            try
            {
                string strFielder = teAccessDBFileName.Text.Trim();
                string strCommonDBFile = strFielder.Trim() + @"\YG1_CMN.mdb";
                string strRCSDBFile = @"\YG1_RCS.mdb";
                string strDRPIFile = @"\YG1_DRPI.mdb";
                string strAccessDBFileName = "";

                if (rboDRPI.Checked)
                    strAccessDBFileName = strFielder.Trim() + strDRPIFile.Trim();
                else
                    strAccessDBFileName = strFielder.Trim() + strRCSDBFile.Trim();

                // DB 파일 체크
                System.IO.FileInfo fi1 = new System.IO.FileInfo(strAccessDBFileName);
                System.IO.FileInfo fi2 = new System.IO.FileInfo(strCommonDBFile);

                if (!fi1.Exists || !fi2.Exists)
                {
                    frmMB.lblMessage.Text = "기존 데이터(.mdb) 파일이 없습니다.";
                    frmMB.TopMost = true;
                    frmMB.ShowDialog();
                    return;
                }

                string strQuery = "";
                string strTableName_OPTVAG = "[" + strAccessDBFileName.Trim() + "].[TB_OPT_AVG] AS TB_OPT_AVG";
                string strTableName_MasterATM = "[" + strCommonDBFile.Trim() + "].[TB_MASTER_ATM] AS TB_MASTER_ATM";
                string strTableName_CoilCode = "[" + strAccessDBFileName.Trim() + "].[TB_CODE_COIL] AS TB_CODE_COIL";
                string strTableName_DRPICoilCode = "[" + strAccessDBFileName.Trim() + "].[TB_DRPI_CODE_COIL] AS TB_DRPI_CODE_COIL";
                int intOh_Degree = nmOh_Degree.Value == null ? 0 : Convert.ToInt32(nmOh_Degree.Value);
                string strSelectHogi = nmHogi.Value == null ? "1 호기" : Convert.ToInt32(nmHogi.Value).ToString().Trim() + " 호기";

                strFielder = string.Format("TB_OPT_AVG.*", strTableName_OPTVAG);
                strFielder = string.Format("{0}, (SELECT TB_MASTER_ATM.ATM_NAME FROM {2} WHERE TB_MASTER_ATM.ATM_ID = TB_OPT_AVG.ATM_ID) AS ATM_NAME"
                    , strFielder, strTableName_OPTVAG, strTableName_MasterATM);

                // 코일명 
                if (rboDRPI.Checked)
                    strFielder = string.Format("{0}, (SELECT TB_DRPI_CODE_COIL.COIL_CODE FROM {1} WHERE TB_DRPI_CODE_COIL.COIL_ID = TB_OPT_AVG.COIL_ID) AS COIL_CODE"
                        , strFielder, strTableName_DRPICoilCode);
                else
                    strFielder = string.Format("{0}, (SELECT TB_CODE_COIL.COIL_CODE FROM {1} WHERE TB_CODE_COIL.COIL_ID = TB_OPT_AVG.COIL_ID) AS COIL_CODE"
                        , strFielder, strTableName_CoilCode);

                strQuery = string.Format("SELECT {0} ", strFielder);
                strQuery = string.Format("{0} FROM {1}", strQuery, strTableName_OPTVAG);

                if (rboDRPI.Checked)
                {
                    strQuery = string.Format("{0} WHERE TB_OPT_AVG.OH_ID = {2} AND TB_OPT_AVG.DRPI_TYPE IN (1, 2)"
                       , strQuery, strTableName_OPTVAG, intOh_Degree);

                    strQuery = string.Format("{0} ORDER BY TB_OPT_AVG.ATM_ID, TB_OPT_AVG.OH_ID, TB_OPT_AVG.DRPI_TYPE, TB_OPT_AVG.COIL_TYPE, TB_OPT_AVG.COIL_ID ;"
                        , strQuery, strTableName_OPTVAG);
                }
                else
                {
                    strQuery = string.Format("{0} WHERE TB_OPT_AVG.OH_ID = {2} AND TB_OPT_AVG.DRPI_TYPE = 0"
                        , strQuery, strTableName_OPTVAG, intOh_Degree);

                    strQuery = string.Format("{0} ORDER BY TB_OPT_AVG.ATM_ID, TB_OPT_AVG.OH_ID, TB_OPT_AVG.DRPY_TYPE, TB_OPT_AVG.COIL_TYPE, TB_OPT_AVG.COIL_ID ;"
                        , strQuery, strTableName_OPTVAG);
                }

                DataTable dt = new DataTable();

                dt = m_Accessdb.SelectOleDb(strAccessDBFileName.Trim(), strQuery);

                if (dt != null && dt.Rows.Count > 0)
                {
                    string strMessage = "";
                    decimal dcmQ_Factor = 0M, dcmAC_Value = 0M, dcmL_Value = 0M;

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        if (string.Format("{0} 호기", Regex.Replace(dt.Rows[i]["ATM_NAME"].ToString().Trim(), @"\D", "")) != strSelectHogi.Trim())
                        {
                            continue;
                        }

                        int intRowIndex = dgvReferenceValue.Rows.Add();

                        if (rboRCS.Checked)
                        {
                            if (m_db.GetRCSReferenceValueDataCount(strPlantName.Trim(), strSelectHogi.Trim(), "제 " + intOh_Degree.ToString().Trim() + " 차수") > 0)
                                strMessage = "해당 차수의 기준값이 존재합니다.";
                            else
                                strMessage = "신규";
                        }
                        else
                        {
                            if (m_db.GetDRPIReferenceValueDataCount(strPlantName.Trim(), strSelectHogi.Trim(), "제 " + intOh_Degree.ToString().Trim() + " 차수") > 0)
                                strMessage = "해당 차수의 기준값이 존재합니다.";
                            else
                                strMessage = "신규";
                        }

                        dgvReferenceValue.Rows[intRowIndex].Cells["UpLoadText"].Value = strMessage.Trim();
                        dgvReferenceValue.Rows[intRowIndex].Cells["PlantName"].Value = strPlantName.Trim();
                        dgvReferenceValue.Rows[intRowIndex].Cells["Hogi"].Value = dt.Rows[i]["ATM_NAME"] == null ? "" : string.Format("{0} 호기", Regex.Replace(dt.Rows[i]["ATM_NAME"].ToString().Trim(), @"\D", ""));
                        dgvReferenceValue.Rows[intRowIndex].Cells["Oh_Degree"].Value = dt.Rows[i]["OH_ID"] == null ? "" : string.Format("제 {0} 차", dt.Rows[i]["OH_ID"].ToString().Trim());

                        if (rboDRPI.Checked)
                        {
                            if (dt.Rows[i]["DRPI_TYPE"].ToString().Trim() == "1")
                                dgvReferenceValue.Rows[intRowIndex].Cells["DRPIGroup"].Value = "A";
                            else
                                dgvReferenceValue.Rows[intRowIndex].Cells["DRPIGroup"].Value = "B";

                            if (dt.Rows[i]["COIL_TYPE"].ToString().Trim() == "1")
                                dgvReferenceValue.Rows[intRowIndex].Cells["DRPIType"].Value = "제어용";
                            else
                                dgvReferenceValue.Rows[intRowIndex].Cells["DRPIType"].Value = "정지용";
                        }
                        else
                        {
                            dgvReferenceValue.Rows[intRowIndex].Cells["DRPIGroup"].Value = "";
                            dgvReferenceValue.Rows[intRowIndex].Cells["DRPIType"].Value = "";
                        }

                        dgvReferenceValue.Rows[intRowIndex].Cells["COIL_ID"].Value = dt.Rows[i]["COIL_ID"] == null ? "" : dt.Rows[i]["COIL_ID"].ToString().Trim();
                        dgvReferenceValue.Rows[intRowIndex].Cells["CoilName"].Value = dt.Rows[i]["COIL_CODE"] == null ? "" : dt.Rows[i]["COIL_CODE"].ToString().Trim().Replace("권선", ""); ;
                        dgvReferenceValue.Rows[intRowIndex].Cells["AVG_DC"].Value = dt.Rows[i]["AVG_DC"] == null ? "" : dt.Rows[i]["AVG_DC"].ToString().Trim();
                        dgvReferenceValue.Rows[intRowIndex].Cells["AVG_AC"].Value = dt.Rows[i]["AVG_AC"] == null ? "" : dt.Rows[i]["AVG_AC"].ToString().Trim();
                        dgvReferenceValue.Rows[intRowIndex].Cells["AVG_L"].Value = dt.Rows[i]["AVG_L"] == null ? "" : dt.Rows[i]["AVG_L"].ToString().Trim();
                        dgvReferenceValue.Rows[intRowIndex].Cells["AVG_C"].Value = dt.Rows[i]["AVG_C"] == null ? "" : dt.Rows[i]["AVG_C"].ToString().Trim();

                        if (dt.Rows[i]["AVG_AC"] != null && dt.Rows[i]["AVG_AC"].ToString().Trim() != "")
                            dcmAC_Value = Convert.ToDecimal(dt.Rows[i]["AVG_AC"].ToString().Trim());
                        else
                            dcmAC_Value = 0M;

                        if (dt.Rows[i]["AVG_L"] != null && dt.Rows[i]["AVG_L"].ToString().Trim() != "")
                            dcmL_Value = Convert.ToDecimal(dt.Rows[i]["AVG_L"].ToString().Trim());
                        else
                            dcmL_Value = 0M;

                        if (dcmAC_Value > 0)
                            dcmQ_Factor = (2M * 3.141592M * 120M * (dcmL_Value / 1000M)) / dcmAC_Value;
                        else
                            dcmQ_Factor = 0M;

                        dgvReferenceValue.Rows[intRowIndex].Cells["AVG_Q"].Value = dt.Rows[i]["AVG_Q"] == null ? "" : dcmQ_Factor.ToString("F3").Trim();
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
        }
    }
}
