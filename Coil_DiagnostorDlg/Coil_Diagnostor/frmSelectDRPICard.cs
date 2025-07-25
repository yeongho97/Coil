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

namespace Coil_Diagnostor
{
    public partial class frmSelectDRPICard : Form
    {
        protected string strPlantName = "";
        protected frmDRPIDiagnosis frmDRPI;
        protected bool boolFormLoad = false;

        protected ComboBox[] arrayDAM = new ComboBox[10];
        protected string[] arrayOldDAM = new string[10];

        public frmSelectDRPICard()
        {
            CheckForIllegalCrossThreadCalls = false;
            InitializeComponent();
        }

        public frmSelectDRPICard(frmDRPIDiagnosis _frm)
        {
            CheckForIllegalCrossThreadCalls = false;
            InitializeComponent();

            frmDRPI = _frm;
        }

        /// <summary>
        /// 폼 단축키 지정
        /// </summary>
        protected override bool ProcessCmdKey(ref System.Windows.Forms.Message msg, Keys keyData)
        {
            Keys key = keyData & ~(Keys.Shift | Keys.Control);

            switch (key)
            {
                case Keys.F3: // 초기화 버튼
                    btnInitialize.PerformClick();
                    break;
                case Keys.F4: // 확인 버튼
                    btnConfirm.PerformClick();
                    break;
                case Keys.F12: // 닫기 버튼
                    btnCancel.PerformClick();
                    break;
            }

            return base.ProcessCmdKey(ref msg, keyData);
        }

        /// <summary>
        /// Form Load
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmSelectDRPICard_Load(object sender, EventArgs e)
        {
            boolFormLoad = true;

            strPlantName = Gini.GetValue("Device", "PlantName").Trim();

            // Array ComboBox 설정
            SetArrayComboBox();

            // ComboBox 데이터 설정
            SetComboBoxDataBinding();

            // Check Box 설정
            SetCheckBoxInitialize(true);

            boolFormLoad = false;
        }

        /// <summary>
        /// Array ComboBox 설정
        /// </summary>
        private void SetArrayComboBox()
        {
            arrayDAM[0] = cboDAMRelay1;
            arrayDAM[1] = cboDAMRelay2;
            arrayDAM[2] = cboDAMRelay3;
            arrayDAM[3] = cboDAMRelay4;
            arrayDAM[4] = cboDAMRelay5;
            arrayDAM[5] = cboDAMRelay6;
            arrayDAM[6] = cboDAMRelay7;
            arrayDAM[7] = cboDAMRelay8;
            arrayDAM[8] = cboDAMRelay9;
            arrayDAM[9] = cboDAMRelay10;
        }

        /// <summary>
        /// ComboBox 데이터 설정
        /// </summary>
        private void SetComboBoxDataBinding()
        {
            // 발전소 호기
            string[] strData = Gini.GetValue("Combo", "DAMRelayItem").Split(',');

            cboDAMRelay1.Items.Clear();
            cboDAMRelay2.Items.Clear();
            cboDAMRelay3.Items.Clear();
            cboDAMRelay4.Items.Clear();
            cboDAMRelay5.Items.Clear();
            cboDAMRelay6.Items.Clear();
            cboDAMRelay7.Items.Clear();
            cboDAMRelay8.Items.Clear();
            cboDAMRelay9.Items.Clear();
            cboDAMRelay10.Items.Clear();

            if (strData.Length > 10)
            {
                for (int i = 0; i < 10; i++)
                {
                    cboDAMRelay1.Items.Add(string.Format("DAM {0}", (i + 1).ToString().Trim()));
                    cboDAMRelay2.Items.Add(string.Format("DAM {0}", (i + 1).ToString().Trim()));
                    cboDAMRelay3.Items.Add(string.Format("DAM {0}", (i + 1).ToString().Trim()));
                    cboDAMRelay4.Items.Add(string.Format("DAM {0}", (i + 1).ToString().Trim()));
                    cboDAMRelay5.Items.Add(string.Format("DAM {0}", (i + 1).ToString().Trim()));
                    cboDAMRelay6.Items.Add(string.Format("DAM {0}", (i + 1).ToString().Trim()));
                    cboDAMRelay7.Items.Add(string.Format("DAM {0}", (i + 1).ToString().Trim()));
                    cboDAMRelay8.Items.Add(string.Format("DAM {0}", (i + 1).ToString().Trim()));
                    cboDAMRelay9.Items.Add(string.Format("DAM {0}", (i + 1).ToString().Trim()));
                    cboDAMRelay10.Items.Add(string.Format("DAM {0}", (i + 1).ToString().Trim()));
                }
            }
            else
            {
                for (int i = 0; i < strData.Length; i++)
                {
                    cboDAMRelay1.Items.Add(strData[i].Trim());
                    cboDAMRelay2.Items.Add(strData[i].Trim());
                    cboDAMRelay3.Items.Add(strData[i].Trim());
                    cboDAMRelay4.Items.Add(strData[i].Trim());
                    cboDAMRelay5.Items.Add(strData[i].Trim());
                    cboDAMRelay6.Items.Add(strData[i].Trim());
                    cboDAMRelay7.Items.Add(strData[i].Trim());
                    cboDAMRelay8.Items.Add(strData[i].Trim());
                    cboDAMRelay9.Items.Add(strData[i].Trim());
                    cboDAMRelay10.Items.Add(strData[i].Trim());
                }
            }

            if (frmDRPI.boolNormalMode)
            {
                cboDAMRelay1.SelectedItem = Gini.GetValue("DRPI", "DRPINormalMode_SelectDAMRelay1").Trim();
                cboDAMRelay2.SelectedItem = Gini.GetValue("DRPI", "DRPINormalMode_SelectDAMRelay2").Trim();
                cboDAMRelay3.SelectedItem = Gini.GetValue("DRPI", "DRPINormalMode_SelectDAMRelay3").Trim();
                cboDAMRelay4.SelectedItem = Gini.GetValue("DRPI", "DRPINormalMode_SelectDAMRelay4").Trim();
                cboDAMRelay5.SelectedItem = Gini.GetValue("DRPI", "DRPINormalMode_SelectDAMRelay5").Trim();
                cboDAMRelay6.SelectedItem = Gini.GetValue("DRPI", "DRPINormalMode_SelectDAMRelay6").Trim();
                cboDAMRelay7.SelectedItem = Gini.GetValue("DRPI", "DRPINormalMode_SelectDAMRelay7").Trim();
                cboDAMRelay8.SelectedItem = Gini.GetValue("DRPI", "DRPINormalMode_SelectDAMRelay8").Trim();
                cboDAMRelay9.SelectedItem = Gini.GetValue("DRPI", "DRPINormalMode_SelectDAMRelay9").Trim();
                cboDAMRelay10.SelectedItem = Gini.GetValue("DRPI", "DRPINormalMode_SelectDAMRelay10").Trim();
            }
            else
            {
                cboDAMRelay1.SelectedItem = Gini.GetValue("DRPI", "DRPIWheatstoneMode_SelectDAMRelay1").Trim();
                cboDAMRelay2.SelectedItem = Gini.GetValue("DRPI", "DRPIWheatstoneMode_SelectDAMRelay2").Trim();
                cboDAMRelay3.SelectedItem = Gini.GetValue("DRPI", "DRPIWheatstoneMode_SelectDAMRelay3").Trim();
                cboDAMRelay4.SelectedItem = Gini.GetValue("DRPI", "DRPIWheatstoneMode_SelectDAMRelay4").Trim();
                cboDAMRelay5.SelectedItem = Gini.GetValue("DRPI", "DRPIWheatstoneMode_SelectDAMRelay5").Trim();
                cboDAMRelay6.SelectedItem = Gini.GetValue("DRPI", "DRPIWheatstoneMode_SelectDAMRelay6").Trim();
                cboDAMRelay7.SelectedItem = Gini.GetValue("DRPI", "DRPIWheatstoneMode_SelectDAMRelay7").Trim();
                cboDAMRelay8.SelectedItem = Gini.GetValue("DRPI", "DRPIWheatstoneMode_SelectDAMRelay8").Trim();
                cboDAMRelay9.SelectedItem = Gini.GetValue("DRPI", "DRPIWheatstoneMode_SelectDAMRelay9").Trim();
                cboDAMRelay10.SelectedItem = Gini.GetValue("DRPI", "DRPIWheatstoneMode_SelectDAMRelay10").Trim();
            }

            arrayOldDAM[0] = cboDAMRelay1.SelectedItem.ToString().Trim();
            arrayOldDAM[1] = cboDAMRelay2.SelectedItem.ToString().Trim();
            arrayOldDAM[2] = cboDAMRelay3.SelectedItem.ToString().Trim();
            arrayOldDAM[3] = cboDAMRelay4.SelectedItem.ToString().Trim();
            arrayOldDAM[4] = cboDAMRelay5.SelectedItem.ToString().Trim();
            arrayOldDAM[5] = cboDAMRelay6.SelectedItem.ToString().Trim();
            arrayOldDAM[6] = cboDAMRelay7.SelectedItem.ToString().Trim();
            arrayOldDAM[7] = cboDAMRelay8.SelectedItem.ToString().Trim();
            arrayOldDAM[8] = cboDAMRelay9.SelectedItem.ToString().Trim();
            arrayOldDAM[9] = cboDAMRelay10.SelectedItem.ToString().Trim();
        }

        /// <summary>
        /// Check Box 설정
        /// </summary>
        private void SetCheckBoxInitialize(bool _boolCheck)
        {
            chkDAM1_ChannelAll.Checked = _boolCheck;
            chkDAM2_ChannelAll.Checked = _boolCheck;
            chkDAM3_ChannelAll.Checked = _boolCheck;
            chkDAM4_ChannelAll.Checked = _boolCheck;
            chkDAM5_ChannelAll.Checked = _boolCheck;
            chkDAM6_ChannelAll.Checked = _boolCheck;
            chkDAM7_ChannelAll.Checked = _boolCheck;
            chkDAM8_ChannelAll.Checked = _boolCheck;
            chkDAM9_ChannelAll.Checked = _boolCheck;
            chkDAM10_ChannelAll.Checked = _boolCheck;
        }

        /// <summary>
        /// 전체 선택 Check Box Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void chkChannelAll_CheckedChanged(object sender, EventArgs e)
        {
            if (((CheckBox)sender).Name.Trim() == "chkDAM1_ChannelAll")
            {
                // 첫번째
                chkDAM1_Channel1.Checked = chkDAM1_ChannelAll.Checked;
                chkDAM1_Channel2.Checked = chkDAM1_ChannelAll.Checked;
                chkDAM1_Channel3.Checked = chkDAM1_ChannelAll.Checked;
                chkDAM1_Channel4.Checked = chkDAM1_ChannelAll.Checked;
                chkDAM1_Channel5.Checked = chkDAM1_ChannelAll.Checked;
                chkDAM1_Channel6.Checked = chkDAM1_ChannelAll.Checked;
                chkDAM1_Channel7.Checked = chkDAM1_ChannelAll.Checked;
                chkDAM1_Channel8.Checked = chkDAM1_ChannelAll.Checked;
                chkDAM1_Channel9.Checked = chkDAM1_ChannelAll.Checked;
                chkDAM1_Channel10.Checked = chkDAM1_ChannelAll.Checked;
                chkDAM1_Channel11.Checked = chkDAM1_ChannelAll.Checked;
                chkDAM1_Channel12.Checked = chkDAM1_ChannelAll.Checked;
                chkDAM1_Channel13.Checked = chkDAM1_ChannelAll.Checked;
                chkDAM1_Channel14.Checked = chkDAM1_ChannelAll.Checked;
                chkDAM1_Channel15.Checked = chkDAM1_ChannelAll.Checked;
                chkDAM1_Channel16.Checked = chkDAM1_ChannelAll.Checked;
                chkDAM1_Channel17.Checked = chkDAM1_ChannelAll.Checked;
                chkDAM1_Channel18.Checked = chkDAM1_ChannelAll.Checked;
                chkDAM1_Channel19.Checked = chkDAM1_ChannelAll.Checked;
                chkDAM1_Channel20.Checked = chkDAM1_ChannelAll.Checked;
                chkDAM1_Channel21.Checked = chkDAM1_ChannelAll.Checked;
            }
            else if (((CheckBox)sender).Name.Trim() == "chkDAM2_ChannelAll")
            {
                // 두번째
                chkDAM2_Channel1.Checked = chkDAM2_ChannelAll.Checked;
                chkDAM2_Channel2.Checked = chkDAM2_ChannelAll.Checked;
                chkDAM2_Channel3.Checked = chkDAM2_ChannelAll.Checked;
                chkDAM2_Channel4.Checked = chkDAM2_ChannelAll.Checked;
                chkDAM2_Channel5.Checked = chkDAM2_ChannelAll.Checked;
                chkDAM2_Channel6.Checked = chkDAM2_ChannelAll.Checked;
                chkDAM2_Channel7.Checked = chkDAM2_ChannelAll.Checked;
                chkDAM2_Channel8.Checked = chkDAM2_ChannelAll.Checked;
                chkDAM2_Channel9.Checked = chkDAM2_ChannelAll.Checked;
                chkDAM2_Channel10.Checked = chkDAM2_ChannelAll.Checked;
                chkDAM2_Channel11.Checked = chkDAM2_ChannelAll.Checked;
                chkDAM2_Channel12.Checked = chkDAM2_ChannelAll.Checked;
                chkDAM2_Channel13.Checked = chkDAM2_ChannelAll.Checked;
                chkDAM2_Channel14.Checked = chkDAM2_ChannelAll.Checked;
                chkDAM2_Channel15.Checked = chkDAM2_ChannelAll.Checked;
                chkDAM2_Channel16.Checked = chkDAM2_ChannelAll.Checked;
                chkDAM2_Channel17.Checked = chkDAM2_ChannelAll.Checked;
                chkDAM2_Channel18.Checked = chkDAM2_ChannelAll.Checked;
                chkDAM2_Channel19.Checked = chkDAM2_ChannelAll.Checked;
                chkDAM2_Channel20.Checked = chkDAM2_ChannelAll.Checked;
                chkDAM2_Channel21.Checked = chkDAM2_ChannelAll.Checked;
            }
            else if (((CheckBox)sender).Name.Trim() == "chkDAM3_ChannelAll")
            {
                // 세번째
                chkDAM3_Channel1.Checked = chkDAM3_ChannelAll.Checked;
                chkDAM3_Channel2.Checked = chkDAM3_ChannelAll.Checked;
                chkDAM3_Channel3.Checked = chkDAM3_ChannelAll.Checked;
                chkDAM3_Channel4.Checked = chkDAM3_ChannelAll.Checked;
                chkDAM3_Channel5.Checked = chkDAM3_ChannelAll.Checked;
                chkDAM3_Channel6.Checked = chkDAM3_ChannelAll.Checked;
                chkDAM3_Channel7.Checked = chkDAM3_ChannelAll.Checked;
                chkDAM3_Channel8.Checked = chkDAM3_ChannelAll.Checked;
                chkDAM3_Channel9.Checked = chkDAM3_ChannelAll.Checked;
                chkDAM3_Channel10.Checked = chkDAM3_ChannelAll.Checked;
                chkDAM3_Channel11.Checked = chkDAM3_ChannelAll.Checked;
                chkDAM3_Channel12.Checked = chkDAM3_ChannelAll.Checked;
                chkDAM3_Channel13.Checked = chkDAM3_ChannelAll.Checked;
                chkDAM3_Channel14.Checked = chkDAM3_ChannelAll.Checked;
                chkDAM3_Channel15.Checked = chkDAM3_ChannelAll.Checked;
                chkDAM3_Channel16.Checked = chkDAM3_ChannelAll.Checked;
                chkDAM3_Channel17.Checked = chkDAM3_ChannelAll.Checked;
                chkDAM3_Channel18.Checked = chkDAM3_ChannelAll.Checked;
                chkDAM3_Channel19.Checked = chkDAM3_ChannelAll.Checked;
                chkDAM3_Channel20.Checked = chkDAM3_ChannelAll.Checked;
                chkDAM3_Channel21.Checked = chkDAM3_ChannelAll.Checked;
            }
            else if (((CheckBox)sender).Name.Trim() == "chkDAM4_ChannelAll")
            {
                // 네번째
                chkDAM4_Channel1.Checked = chkDAM4_ChannelAll.Checked;
                chkDAM4_Channel2.Checked = chkDAM4_ChannelAll.Checked;
                chkDAM4_Channel3.Checked = chkDAM4_ChannelAll.Checked;
                chkDAM4_Channel4.Checked = chkDAM4_ChannelAll.Checked;
                chkDAM4_Channel5.Checked = chkDAM4_ChannelAll.Checked;
                chkDAM4_Channel6.Checked = chkDAM4_ChannelAll.Checked;
                chkDAM4_Channel7.Checked = chkDAM4_ChannelAll.Checked;
                chkDAM4_Channel8.Checked = chkDAM4_ChannelAll.Checked;
                chkDAM4_Channel9.Checked = chkDAM4_ChannelAll.Checked;
                chkDAM4_Channel10.Checked = chkDAM4_ChannelAll.Checked;
                chkDAM4_Channel11.Checked = chkDAM4_ChannelAll.Checked;
                chkDAM4_Channel12.Checked = chkDAM4_ChannelAll.Checked;
                chkDAM4_Channel13.Checked = chkDAM4_ChannelAll.Checked;
                chkDAM4_Channel14.Checked = chkDAM4_ChannelAll.Checked;
                chkDAM4_Channel15.Checked = chkDAM4_ChannelAll.Checked;
                chkDAM4_Channel16.Checked = chkDAM4_ChannelAll.Checked;
                chkDAM4_Channel17.Checked = chkDAM4_ChannelAll.Checked;
                chkDAM4_Channel18.Checked = chkDAM4_ChannelAll.Checked;
                chkDAM4_Channel19.Checked = chkDAM4_ChannelAll.Checked;
                chkDAM4_Channel20.Checked = chkDAM4_ChannelAll.Checked;
                chkDAM4_Channel21.Checked = chkDAM4_ChannelAll.Checked;
            }
            else if (((CheckBox)sender).Name.Trim() == "chkDAM5_ChannelAll")
            {
                // 다섯번째
                chkDAM5_Channel1.Checked = chkDAM5_ChannelAll.Checked;
                chkDAM5_Channel2.Checked = chkDAM5_ChannelAll.Checked;
                chkDAM5_Channel3.Checked = chkDAM5_ChannelAll.Checked;
                chkDAM5_Channel4.Checked = chkDAM5_ChannelAll.Checked;
                chkDAM5_Channel5.Checked = chkDAM5_ChannelAll.Checked;
                chkDAM5_Channel6.Checked = chkDAM5_ChannelAll.Checked;
                chkDAM5_Channel7.Checked = chkDAM5_ChannelAll.Checked;
                chkDAM5_Channel8.Checked = chkDAM5_ChannelAll.Checked;
                chkDAM5_Channel9.Checked = chkDAM5_ChannelAll.Checked;
                chkDAM5_Channel10.Checked = chkDAM5_ChannelAll.Checked;
                chkDAM5_Channel11.Checked = chkDAM5_ChannelAll.Checked;
                chkDAM5_Channel12.Checked = chkDAM5_ChannelAll.Checked;
                chkDAM5_Channel13.Checked = chkDAM5_ChannelAll.Checked;
                chkDAM5_Channel14.Checked = chkDAM5_ChannelAll.Checked;
                chkDAM5_Channel15.Checked = chkDAM5_ChannelAll.Checked;
                chkDAM5_Channel16.Checked = chkDAM5_ChannelAll.Checked;
                chkDAM5_Channel17.Checked = chkDAM5_ChannelAll.Checked;
                chkDAM5_Channel18.Checked = chkDAM5_ChannelAll.Checked;
                chkDAM5_Channel19.Checked = chkDAM5_ChannelAll.Checked;
                chkDAM5_Channel20.Checked = chkDAM5_ChannelAll.Checked;
                chkDAM5_Channel21.Checked = chkDAM5_ChannelAll.Checked;
            }
            else if (((CheckBox)sender).Name.Trim() == "chkDAM6_ChannelAll")
            {
                // 여섯번째
                chkDAM6_Channel1.Checked = chkDAM6_ChannelAll.Checked;
                chkDAM6_Channel2.Checked = chkDAM6_ChannelAll.Checked;
                chkDAM6_Channel3.Checked = chkDAM6_ChannelAll.Checked;
                chkDAM6_Channel4.Checked = chkDAM6_ChannelAll.Checked;
                chkDAM6_Channel5.Checked = chkDAM6_ChannelAll.Checked;
                chkDAM6_Channel6.Checked = chkDAM6_ChannelAll.Checked;
                chkDAM6_Channel7.Checked = chkDAM6_ChannelAll.Checked;
                chkDAM6_Channel8.Checked = chkDAM6_ChannelAll.Checked;
                chkDAM6_Channel9.Checked = chkDAM6_ChannelAll.Checked;
                chkDAM6_Channel10.Checked = chkDAM6_ChannelAll.Checked;
                chkDAM6_Channel11.Checked = chkDAM6_ChannelAll.Checked;
                chkDAM6_Channel12.Checked = chkDAM6_ChannelAll.Checked;
                chkDAM6_Channel13.Checked = chkDAM6_ChannelAll.Checked;
                chkDAM6_Channel14.Checked = chkDAM6_ChannelAll.Checked;
                chkDAM6_Channel15.Checked = chkDAM6_ChannelAll.Checked;
                chkDAM6_Channel16.Checked = chkDAM6_ChannelAll.Checked;
                chkDAM6_Channel17.Checked = chkDAM6_ChannelAll.Checked;
                chkDAM6_Channel18.Checked = chkDAM6_ChannelAll.Checked;
                chkDAM6_Channel19.Checked = chkDAM6_ChannelAll.Checked;
                chkDAM6_Channel20.Checked = chkDAM6_ChannelAll.Checked;
                chkDAM6_Channel21.Checked = chkDAM6_ChannelAll.Checked;
            }
            else if (((CheckBox)sender).Name.Trim() == "chkDAM7_ChannelAll")
            {
                // 여섯번째
                chkDAM7_Channel1.Checked = chkDAM7_ChannelAll.Checked;
                chkDAM7_Channel2.Checked = chkDAM7_ChannelAll.Checked;
                chkDAM7_Channel3.Checked = chkDAM7_ChannelAll.Checked;
                chkDAM7_Channel4.Checked = chkDAM7_ChannelAll.Checked;
                chkDAM7_Channel5.Checked = chkDAM7_ChannelAll.Checked;
                chkDAM7_Channel6.Checked = chkDAM7_ChannelAll.Checked;
                chkDAM7_Channel7.Checked = chkDAM7_ChannelAll.Checked;
                chkDAM7_Channel8.Checked = chkDAM7_ChannelAll.Checked;
                chkDAM7_Channel9.Checked = chkDAM7_ChannelAll.Checked;
                chkDAM7_Channel10.Checked = chkDAM7_ChannelAll.Checked;
                chkDAM7_Channel11.Checked = chkDAM7_ChannelAll.Checked;
                chkDAM7_Channel12.Checked = chkDAM7_ChannelAll.Checked;
                chkDAM7_Channel13.Checked = chkDAM7_ChannelAll.Checked;
                chkDAM7_Channel14.Checked = chkDAM7_ChannelAll.Checked;
                chkDAM7_Channel15.Checked = chkDAM7_ChannelAll.Checked;
                chkDAM7_Channel16.Checked = chkDAM7_ChannelAll.Checked;
                chkDAM7_Channel17.Checked = chkDAM7_ChannelAll.Checked;
                chkDAM7_Channel18.Checked = chkDAM7_ChannelAll.Checked;
                chkDAM7_Channel19.Checked = chkDAM7_ChannelAll.Checked;
                chkDAM7_Channel20.Checked = chkDAM7_ChannelAll.Checked;
                chkDAM7_Channel21.Checked = chkDAM7_ChannelAll.Checked;
            }
            else if (((CheckBox)sender).Name.Trim() == "chkDAM8_ChannelAll")
            {
                // 여섯번째
                chkDAM8_Channel1.Checked = chkDAM8_ChannelAll.Checked;
                chkDAM8_Channel2.Checked = chkDAM8_ChannelAll.Checked;
                chkDAM8_Channel3.Checked = chkDAM8_ChannelAll.Checked;
                chkDAM8_Channel4.Checked = chkDAM8_ChannelAll.Checked;
                chkDAM8_Channel5.Checked = chkDAM8_ChannelAll.Checked;
                chkDAM8_Channel6.Checked = chkDAM8_ChannelAll.Checked;
                chkDAM8_Channel7.Checked = chkDAM8_ChannelAll.Checked;
                chkDAM8_Channel8.Checked = chkDAM8_ChannelAll.Checked;
                chkDAM8_Channel9.Checked = chkDAM8_ChannelAll.Checked;
                chkDAM8_Channel10.Checked = chkDAM8_ChannelAll.Checked;
                chkDAM8_Channel11.Checked = chkDAM8_ChannelAll.Checked;
                chkDAM8_Channel12.Checked = chkDAM8_ChannelAll.Checked;
                chkDAM8_Channel13.Checked = chkDAM8_ChannelAll.Checked;
                chkDAM8_Channel14.Checked = chkDAM8_ChannelAll.Checked;
                chkDAM8_Channel15.Checked = chkDAM8_ChannelAll.Checked;
                chkDAM8_Channel16.Checked = chkDAM8_ChannelAll.Checked;
                chkDAM8_Channel17.Checked = chkDAM8_ChannelAll.Checked;
                chkDAM8_Channel18.Checked = chkDAM8_ChannelAll.Checked;
                chkDAM8_Channel19.Checked = chkDAM8_ChannelAll.Checked;
                chkDAM8_Channel20.Checked = chkDAM8_ChannelAll.Checked;
                chkDAM8_Channel21.Checked = chkDAM8_ChannelAll.Checked;
            }
            else if (((CheckBox)sender).Name.Trim() == "chkDAM9_ChannelAll")
            {
                // 여섯번째
                chkDAM9_Channel1.Checked = chkDAM9_ChannelAll.Checked;
                chkDAM9_Channel2.Checked = chkDAM9_ChannelAll.Checked;
                chkDAM9_Channel3.Checked = chkDAM9_ChannelAll.Checked;
                chkDAM9_Channel4.Checked = chkDAM9_ChannelAll.Checked;
                chkDAM9_Channel5.Checked = chkDAM9_ChannelAll.Checked;
                chkDAM9_Channel6.Checked = chkDAM9_ChannelAll.Checked;
                chkDAM9_Channel7.Checked = chkDAM9_ChannelAll.Checked;
                chkDAM9_Channel8.Checked = chkDAM9_ChannelAll.Checked;
                chkDAM9_Channel9.Checked = chkDAM9_ChannelAll.Checked;
                chkDAM9_Channel10.Checked = chkDAM9_ChannelAll.Checked;
                chkDAM9_Channel11.Checked = chkDAM9_ChannelAll.Checked;
                chkDAM9_Channel12.Checked = chkDAM9_ChannelAll.Checked;
                chkDAM9_Channel13.Checked = chkDAM9_ChannelAll.Checked;
                chkDAM9_Channel14.Checked = chkDAM9_ChannelAll.Checked;
                chkDAM9_Channel15.Checked = chkDAM9_ChannelAll.Checked;
                chkDAM9_Channel16.Checked = chkDAM9_ChannelAll.Checked;
                chkDAM9_Channel17.Checked = chkDAM9_ChannelAll.Checked;
                chkDAM9_Channel18.Checked = chkDAM9_ChannelAll.Checked;
                chkDAM9_Channel19.Checked = chkDAM9_ChannelAll.Checked;
                chkDAM9_Channel20.Checked = chkDAM9_ChannelAll.Checked;
                chkDAM9_Channel21.Checked = chkDAM9_ChannelAll.Checked;
            }
            else if (((CheckBox)sender).Name.Trim() == "chkDAM10_ChannelAll")
            {
                // 여섯번째
                chkDAM10_Channel1.Checked = chkDAM10_ChannelAll.Checked;
                chkDAM10_Channel2.Checked = chkDAM10_ChannelAll.Checked;
                chkDAM10_Channel3.Checked = chkDAM10_ChannelAll.Checked;
                chkDAM10_Channel4.Checked = chkDAM10_ChannelAll.Checked;
                chkDAM10_Channel5.Checked = chkDAM10_ChannelAll.Checked;
                chkDAM10_Channel6.Checked = chkDAM10_ChannelAll.Checked;
                chkDAM10_Channel7.Checked = chkDAM10_ChannelAll.Checked;
                chkDAM10_Channel8.Checked = chkDAM10_ChannelAll.Checked;
                chkDAM10_Channel9.Checked = chkDAM10_ChannelAll.Checked;
                chkDAM10_Channel10.Checked = chkDAM10_ChannelAll.Checked;
                chkDAM10_Channel11.Checked = chkDAM10_ChannelAll.Checked;
                chkDAM10_Channel12.Checked = chkDAM10_ChannelAll.Checked;
                chkDAM10_Channel12.Checked = chkDAM10_ChannelAll.Checked;
                chkDAM10_Channel13.Checked = chkDAM10_ChannelAll.Checked;
                chkDAM10_Channel14.Checked = chkDAM10_ChannelAll.Checked;
                chkDAM10_Channel15.Checked = chkDAM10_ChannelAll.Checked;
                chkDAM10_Channel16.Checked = chkDAM10_ChannelAll.Checked;
                chkDAM10_Channel17.Checked = chkDAM10_ChannelAll.Checked;
                chkDAM10_Channel18.Checked = chkDAM10_ChannelAll.Checked;
                chkDAM10_Channel19.Checked = chkDAM10_ChannelAll.Checked;
                chkDAM10_Channel20.Checked = chkDAM10_ChannelAll.Checked;
                chkDAM10_Channel21.Checked = chkDAM10_ChannelAll.Checked;
            }
        }

        /// <summary>
        /// 초기화 Button Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnInitialize_Click(object sender, EventArgs e)
        {
            cboDAMRelay1.SelectedIndex = 0;
            cboDAMRelay2.SelectedIndex = 1;
            cboDAMRelay3.SelectedIndex = 2;
            cboDAMRelay4.SelectedIndex = 3;
            cboDAMRelay5.SelectedIndex = 4;
            cboDAMRelay6.SelectedIndex = 5;
            cboDAMRelay7.SelectedIndex = 6;
            cboDAMRelay8.SelectedIndex = 7;
            cboDAMRelay9.SelectedIndex = 8;
            cboDAMRelay10.SelectedIndex = 9;

            arrayOldDAM[0] = cboDAMRelay1.SelectedItem.ToString().Trim();
            arrayOldDAM[1] = cboDAMRelay2.SelectedItem.ToString().Trim();
            arrayOldDAM[2] = cboDAMRelay3.SelectedItem.ToString().Trim();
            arrayOldDAM[3] = cboDAMRelay4.SelectedItem.ToString().Trim();
            arrayOldDAM[4] = cboDAMRelay5.SelectedItem.ToString().Trim();
            arrayOldDAM[5] = cboDAMRelay6.SelectedItem.ToString().Trim();
            arrayOldDAM[6] = cboDAMRelay7.SelectedItem.ToString().Trim();
            arrayOldDAM[7] = cboDAMRelay8.SelectedItem.ToString().Trim();
            arrayOldDAM[8] = cboDAMRelay9.SelectedItem.ToString().Trim();
            arrayOldDAM[9] = cboDAMRelay10.SelectedItem.ToString().Trim();
        }

        /// <summary>
        /// 취소 Button Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// 확인 Button Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnConfirm_Click(object sender, EventArgs e)
        {
            try
            {
                // RCS 측정 화면에 선택 Card 보내기
                for (int i = 0; i < arrayDAM.Length; i++)
                {
                    if (frmDRPI.boolNormalMode)
                    {
                        // 일반 모드
                        SetSelectRelayCardNormalModePinMap(arrayDAM[i].SelectedItem.ToString().Trim(), i);
                    }
                    else
                    {
                        // 휘스톤 모드
                        SetSelectRelayCardWheatstoneModePinMap(arrayDAM[i].SelectedItem.ToString().Trim(), i);
                    }
                }

                this.Close();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
        }

        /// <summary>
        /// 일반 모드
        /// </summary>
        /// <param name="_strRelayCard"></param>
        /// <param name="intArrayCardPinMapIndex"></param>
        private void SetSelectRelayCardNormalModePinMap(string _strRelayCard, int intArrayCardPinMapIndex)
        {
            string strRelayCardPinMap = "";
            //_strSelectDAMRelay
            strRelayCardPinMap = string.Format(Gini.GetValue("DAM_PinName", $"{_strRelayCard.Replace(" ", "")}_RelayCardNumberDAQPinName".Trim()), Gini.GetValue("Device", "DAQDeviceName").Trim());

            if (intArrayCardPinMapIndex == 0)
            {
                for (int i = 0; i < frmDRPI.arrayCard01_NormalModePinMap.Length; i++)
                {
                    frmDRPI.arrayCard01_NormalModePinMap[i] = string.Format(Gini.GetValue("Channel_PinName", $"NormalModeCh{i + 1}_DAQPinName").Trim(), Gini.GetValue("Device", "DAQDeviceName"), strRelayCardPinMap);
                }
                Gini.SetValue("DRPI", $"DRPINormalMode_SelectDAMRelay1", _strRelayCard.Trim());
            }
            else if (intArrayCardPinMapIndex == 1)
            {
                for (int i = 0; i < frmDRPI.arrayCard02_NormalModePinMap.Length; i++)
                {
                    frmDRPI.arrayCard02_NormalModePinMap[i] = string.Format(Gini.GetValue("Channel_PinName", $"NormalModeCh{i + 1}_DAQPinName").Trim(), Gini.GetValue("Device", "DAQDeviceName"), strRelayCardPinMap);
                }
                Gini.SetValue("DRPI", $"DRPINormalMode_SelectDAMRelay2", _strRelayCard.Trim());
            }
            else if (intArrayCardPinMapIndex == 2)
            {
                for (int i = 0; i < frmDRPI.arrayCard03_NormalModePinMap.Length; i++)
                {
                    frmDRPI.arrayCard03_NormalModePinMap[i] = string.Format(Gini.GetValue("Channel_PinName", $"NormalModeCh{i + 1}_DAQPinName").Trim(), Gini.GetValue("Device", "DAQDeviceName"), strRelayCardPinMap);
                }
                Gini.SetValue("DRPI", $"DRPINormalMode_SelectDAMRelay3", _strRelayCard.Trim());
            }
            else if (intArrayCardPinMapIndex == 3)
            {
                for (int i = 0; i < frmDRPI.arrayCard04_NormalModePinMap.Length; i++)
                {
                    frmDRPI.arrayCard04_NormalModePinMap[i] = string.Format(Gini.GetValue("Channel_PinName", $"NormalModeCh{i + 1}_DAQPinName").Trim(), Gini.GetValue("Device", "DAQDeviceName"), strRelayCardPinMap);
                }
                Gini.SetValue("DRPI", $"DRPINormalMode_SelectDAMRelay4", _strRelayCard.Trim());
            }
            else if (intArrayCardPinMapIndex == 4)
            {
                for (int i = 0; i < frmDRPI.arrayCard05_NormalModePinMap.Length; i++)
                {
                    frmDRPI.arrayCard05_NormalModePinMap[i] = string.Format(Gini.GetValue("Channel_PinName", $"NormalModeCh{i + 1}_DAQPinName").Trim(), Gini.GetValue("Device", "DAQDeviceName"), strRelayCardPinMap);
                }
                Gini.SetValue("DRPI", $"DRPINormalMode_SelectDAMRelay5", _strRelayCard.Trim());
            }
            else if (intArrayCardPinMapIndex == 5)
            {
                for (int i = 0; i < frmDRPI.arrayCard06_NormalModePinMap.Length; i++)
                {
                    frmDRPI.arrayCard06_NormalModePinMap[i] = string.Format(Gini.GetValue("Channel_PinName", $"NormalModeCh{i + 1}_DAQPinName").Trim(), Gini.GetValue("Device", "DAQDeviceName"), strRelayCardPinMap);
                }
                Gini.SetValue("DRPI", $"DRPINormalMode_SelectDAMRelay6", _strRelayCard.Trim());
            }
            else if (intArrayCardPinMapIndex == 6)
            {
                for (int i = 0; i < frmDRPI.arrayCard07_NormalModePinMap.Length; i++)
                {
                    frmDRPI.arrayCard07_NormalModePinMap[i] = string.Format(Gini.GetValue("Channel_PinName", $"NormalModeCh{i + 1}_DAQPinName").Trim(), Gini.GetValue("Device", "DAQDeviceName"), strRelayCardPinMap);
                }
                Gini.SetValue("DRPI", $"DRPINormalMode_SelectDAMRelay7", _strRelayCard.Trim());
            }
            else if (intArrayCardPinMapIndex == 7)
            {
                for (int i = 0; i < frmDRPI.arrayCard08_NormalModePinMap.Length; i++)
                {
                    frmDRPI.arrayCard08_NormalModePinMap[i] = string.Format(Gini.GetValue("Channel_PinName", $"NormalModeCh{i + 1}_DAQPinName").Trim(), Gini.GetValue("Device", "DAQDeviceName"), strRelayCardPinMap);
                }
                Gini.SetValue("DRPI", $"DRPINormalMode_SelectDAMRelay8", _strRelayCard.Trim());
            }
            else if (intArrayCardPinMapIndex == 8)
            {
                for (int i = 0; i < frmDRPI.arrayCard09_NormalModePinMap.Length; i++)
                {
                    frmDRPI.arrayCard09_NormalModePinMap[i] = string.Format(Gini.GetValue("Channel_PinName", $"NormalModeCh{i + 1}_DAQPinName").Trim(), Gini.GetValue("Device", "DAQDeviceName"), strRelayCardPinMap);
                }
                Gini.SetValue("DRPI", $"DRPINormalMode_SelectDAMRelay9", _strRelayCard.Trim());
            }
            else if (intArrayCardPinMapIndex == 9)
            {
                for (int i = 0; i < frmDRPI.arrayCard10_NormalModePinMap.Length; i++)
                {
                    frmDRPI.arrayCard10_NormalModePinMap[i] = string.Format(Gini.GetValue("Channel_PinName", $"NormalModeCh{i + 1}_DAQPinName").Trim(), Gini.GetValue("Device", "DAQDeviceName"), strRelayCardPinMap);
                }
                Gini.SetValue("DRPI", $"DRPINormalMode_SelectDAMRelay10", _strRelayCard.Trim());
            }
        }

        /// <summary>
        /// 휘스톤 모드
        /// </summary>
        /// <param name="_strRelayCard"></param>
        /// <param name="intArrayCardPinMapIndex"></param>
        private void SetSelectRelayCardWheatstoneModePinMap(string _strRelayCard, int intArrayCardPinMapIndex)
        {
            string strRelayCardPinMap = "";
            //_strSelectDAMRelay
            strRelayCardPinMap = string.Format(Gini.GetValue("DAM_PinName", $"{_strRelayCard.Replace(" ", "")}_RelayCardNumberDAQPinName".Trim()), Gini.GetValue("Device", "DAQDeviceName").Trim());

            if (intArrayCardPinMapIndex == 0)
            { 
                for (int i = 0; i < frmDRPI.arrayCard01_WheatstoneModePinMap.Length; i++)
                {
                    frmDRPI.arrayCard01_WheatstoneModePinMap[i] = string.Format(Gini.GetValue("Channel_PinName", $"WheatstoneModeCh{i + 1}_DAQPinName").Trim(), Gini.GetValue("Device", "DAQDeviceName"), strRelayCardPinMap);
                }
                Gini.SetValue("DRPI", $"DRPIWheatstoneMode_SelectDAMRelay1", _strRelayCard.Trim());
            }
            else if (intArrayCardPinMapIndex == 1)
            {
                for (int i = 0; i < frmDRPI.arrayCard02_WheatstoneModePinMap.Length; i++)
                {
                    frmDRPI.arrayCard02_WheatstoneModePinMap[i] = string.Format(Gini.GetValue("Channel_PinName", $"WheatstoneModeCh{i + 1}_DAQPinName").Trim(), Gini.GetValue("Device", "DAQDeviceName"), strRelayCardPinMap);
                }
                Gini.SetValue("DRPI", $"DRPIWheatstoneMode_SelectDAMRelay2", _strRelayCard.Trim());
            }
            else if (intArrayCardPinMapIndex == 2)
            {
                for (int i = 0; i < frmDRPI.arrayCard03_WheatstoneModePinMap.Length; i++)
                {
                    frmDRPI.arrayCard03_WheatstoneModePinMap[i] = string.Format(Gini.GetValue("Channel_PinName", $"WheatstoneModeCh{i + 1}_DAQPinName").Trim(), Gini.GetValue("Device", "DAQDeviceName"), strRelayCardPinMap);
                }
                Gini.SetValue("DRPI", $"DRPIWheatstoneMode_SelectDAMRelay3", _strRelayCard.Trim());
            }
            else if (intArrayCardPinMapIndex == 3)
            {
                for (int i = 0; i < frmDRPI.arrayCard04_WheatstoneModePinMap.Length; i++)
                {
                    frmDRPI.arrayCard04_WheatstoneModePinMap[i] = string.Format(Gini.GetValue("Channel_PinName", $"WheatstoneModeCh{i + 1}_DAQPinName").Trim(), Gini.GetValue("Device", "DAQDeviceName"), strRelayCardPinMap);
                }
                Gini.SetValue("DRPI", $"DRPIWheatstoneMode_SelectDAMRelay4", _strRelayCard.Trim());
            }
            else if (intArrayCardPinMapIndex == 4)
            {
                for (int i = 0; i < frmDRPI.arrayCard05_WheatstoneModePinMap.Length; i++)
                {
                    frmDRPI.arrayCard05_WheatstoneModePinMap[i] = string.Format(Gini.GetValue("Channel_PinName", $"WheatstoneModeCh{i + 1}_DAQPinName").Trim(), Gini.GetValue("Device", "DAQDeviceName"), strRelayCardPinMap);
                }
                Gini.SetValue("DRPI", $"DRPIWheatstoneMode_SelectDAMRelay5", _strRelayCard.Trim());
            }
            else if (intArrayCardPinMapIndex == 5)
            {
                for (int i = 0; i < frmDRPI.arrayCard06_WheatstoneModePinMap.Length; i++)
                {
                    frmDRPI.arrayCard06_WheatstoneModePinMap[i] = string.Format(Gini.GetValue("Channel_PinName", $"WheatstoneModeCh{i + 1}_DAQPinName").Trim(), Gini.GetValue("Device", "DAQDeviceName"), strRelayCardPinMap);
                }
                Gini.SetValue("DRPI", $"DRPIWheatstoneMode_SelectDAMRelay6", _strRelayCard.Trim());
            }
            else if (intArrayCardPinMapIndex == 6)
            {
                for (int i = 0; i < frmDRPI.arrayCard07_WheatstoneModePinMap.Length; i++)
                {
                    frmDRPI.arrayCard07_WheatstoneModePinMap[i] = string.Format(Gini.GetValue("Channel_PinName", $"WheatstoneModeCh{i + 1}_DAQPinName").Trim(), Gini.GetValue("Device", "DAQDeviceName"), strRelayCardPinMap);
                }
                Gini.SetValue("DRPI", $"DRPIWheatstoneMode_SelectDAMRelay7", _strRelayCard.Trim());
            }
            else if (intArrayCardPinMapIndex == 7)
            {
                for (int i = 0; i < frmDRPI.arrayCard08_WheatstoneModePinMap.Length; i++)
                {
                    frmDRPI.arrayCard08_WheatstoneModePinMap[i] = string.Format(Gini.GetValue("Channel_PinName", $"WheatstoneModeCh{i + 1}_DAQPinName").Trim(), Gini.GetValue("Device", "DAQDeviceName"), strRelayCardPinMap);
                }
                Gini.SetValue("DRPI", $"DRPIWheatstoneMode_SelectDAMRelay8", _strRelayCard.Trim());
            }
            else if (intArrayCardPinMapIndex == 8)
            {
                for (int i = 0; i < frmDRPI.arrayCard09_WheatstoneModePinMap.Length; i++)
                {
                    frmDRPI.arrayCard09_WheatstoneModePinMap[i] = string.Format(Gini.GetValue("Channel_PinName", $"WheatstoneModeCh{i + 1}_DAQPinName").Trim(), Gini.GetValue("Device", "DAQDeviceName"), strRelayCardPinMap);
                }
                Gini.SetValue("DRPI", $"DRPIWheatstoneMode_SelectDAMRelay9", _strRelayCard.Trim());
            }
            else if (intArrayCardPinMapIndex == 9)
            {
                for (int i = 0; i < frmDRPI.arrayCard10_WheatstoneModePinMap.Length; i++)
                {
                    frmDRPI.arrayCard10_WheatstoneModePinMap[i] = string.Format(Gini.GetValue("Channel_PinName", $"WheatstoneModeCh{i + 1}_DAQPinName").Trim(), Gini.GetValue("Device", "DAQDeviceName"), strRelayCardPinMap);
                }
                Gini.SetValue("DRPI", $"DRPIWheatstoneMode_SelectDAMRelay10", _strRelayCard.Trim());
            }
        }

        /// <summary>
        /// 카드 ComboBox Selected Index Changed Evnet
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cboDAMRelay_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (boolFormLoad) return;

            try
            {
                string strSelectDAM = ((ComboBox)sender).SelectedItem.ToString().Trim();

                int intCBOIndex = 0, intNextIndex = 0;
                for (int i = 0; i < arrayOldDAM.Length; i++)
                {
                    if (arrayOldDAM[i].Trim() == strSelectDAM.Trim())
                        intNextIndex = i;

                    if (((ComboBox)sender).Name.Trim() == arrayDAM[i].Name.Trim())
                        intCBOIndex = i;
                }

                arrayDAM[intNextIndex].SelectedItem = arrayOldDAM[intCBOIndex].Trim();
                arrayOldDAM[intCBOIndex] = strSelectDAM.Trim();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
        }
    }
}
