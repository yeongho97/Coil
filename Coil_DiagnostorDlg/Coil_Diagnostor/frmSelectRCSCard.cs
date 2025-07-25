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
    public partial class frmSelectRCSCard : Form
    {
        protected string strPlantName = "";
        protected frmRCSDiagnosis frmRCS;
        protected bool boolFormLoad = false;
        protected string strSelectDAM = "";

        protected ComboBox[] arrayDAM = new ComboBox[6];
        protected string[] arrayOldDAM = new string[6];

        public frmSelectRCSCard()
        {
            CheckForIllegalCrossThreadCalls = false;
            InitializeComponent();
        }

        public frmSelectRCSCard(frmRCSDiagnosis _frm)
        {
            CheckForIllegalCrossThreadCalls = false;
            InitializeComponent();

            frmRCS = _frm;
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
        private void frmSelectRCSCard_Load(object sender, EventArgs e)
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

            for (int i = 0; i < 10; i++)
            {
                cboDAMRelay1.Items.Add(strData[i].Trim());
                cboDAMRelay2.Items.Add(strData[i].Trim());
                cboDAMRelay3.Items.Add(strData[i].Trim());
                cboDAMRelay4.Items.Add(strData[i].Trim());
                cboDAMRelay5.Items.Add(strData[i].Trim());
                cboDAMRelay6.Items.Add(strData[i].Trim());
            }

            if (frmRCS.boolNormalMode)
            {
                cboDAMRelay1.SelectedItem = Gini.GetValue("RCS", "RCSNormalMode_SelectDAMRelay1").Trim();
                cboDAMRelay2.SelectedItem = Gini.GetValue("RCS", "RCSNormalMode_SelectDAMRelay2").Trim();
                cboDAMRelay3.SelectedItem = Gini.GetValue("RCS", "RCSNormalMode_SelectDAMRelay3").Trim();
                cboDAMRelay4.SelectedItem = Gini.GetValue("RCS", "RCSNormalMode_SelectDAMRelay4").Trim();
                cboDAMRelay5.SelectedItem = Gini.GetValue("RCS", "RCSNormalMode_SelectDAMRelay5").Trim();
                cboDAMRelay6.SelectedItem = Gini.GetValue("RCS", "RCSNormalMode_SelectDAMRelay6").Trim();
            }
            else
            {
                cboDAMRelay1.SelectedItem = Gini.GetValue("RCS", "RCSWheatstoneMode_SelectDAMRelay1").Trim();
                cboDAMRelay2.SelectedItem = Gini.GetValue("RCS", "RCSWheatstoneMode_SelectDAMRelay2").Trim();
                cboDAMRelay3.SelectedItem = Gini.GetValue("RCS", "RCSWheatstoneMode_SelectDAMRelay3").Trim();
                cboDAMRelay4.SelectedItem = Gini.GetValue("RCS", "RCSWheatstoneMode_SelectDAMRelay4").Trim();
                cboDAMRelay5.SelectedItem = Gini.GetValue("RCS", "RCSWheatstoneMode_SelectDAMRelay5").Trim();
                cboDAMRelay6.SelectedItem = Gini.GetValue("RCS", "RCSWheatstoneMode_SelectDAMRelay6").Trim();
            }

            arrayOldDAM[0] = cboDAMRelay1.SelectedItem.ToString().Trim();
            arrayOldDAM[1] = cboDAMRelay2.SelectedItem.ToString().Trim();
            arrayOldDAM[2] = cboDAMRelay3.SelectedItem.ToString().Trim();
            arrayOldDAM[3] = cboDAMRelay4.SelectedItem.ToString().Trim();
            arrayOldDAM[4] = cboDAMRelay5.SelectedItem.ToString().Trim();
            arrayOldDAM[5] = cboDAMRelay6.SelectedItem.ToString().Trim();
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

            arrayOldDAM[0] = cboDAMRelay1.SelectedItem.ToString().Trim();
            arrayOldDAM[1] = cboDAMRelay2.SelectedItem.ToString().Trim();
            arrayOldDAM[2] = cboDAMRelay3.SelectedItem.ToString().Trim();
            arrayOldDAM[3] = cboDAMRelay4.SelectedItem.ToString().Trim();
            arrayOldDAM[4] = cboDAMRelay5.SelectedItem.ToString().Trim();
            arrayOldDAM[5] = cboDAMRelay6.SelectedItem.ToString().Trim();
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
                    if (frmRCS.boolNormalMode)
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

            strRelayCardPinMap = string.Format(Gini.GetValue("DAM_PinName", $"{_strRelayCard.Replace(" ", "")}_RelayCardNumberDAQPinName".Trim()), Gini.GetValue("Device", "DAQDeviceName").Trim());

            if (intArrayCardPinMapIndex == 0)
            {
                for (int i = 0; i < frmRCS.arrayCard01_NormalModePinMap.Length; i++)
                {
                    frmRCS.arrayCard01_NormalModePinMap[i] = string.Format(Gini.GetValue("Channel_PinName", $"NormalModeCh{i + 1}_DAQPinName").Trim(), Gini.GetValue("Device", "DAQDeviceName"), strRelayCardPinMap);
                }
                Gini.SetValue("RCS", $"RCSNormalMode_SelectDAMRelay1", _strRelayCard.Trim());

            }
            else if (intArrayCardPinMapIndex == 1)
            {
                for (int i = 0; i < frmRCS.arrayCard02_NormalModePinMap.Length; i++)
                {
                    frmRCS.arrayCard02_NormalModePinMap[i] = string.Format(Gini.GetValue("Channel_PinName", $"NormalModeCh{i + 1}_DAQPinName").Trim(), Gini.GetValue("Device", "DAQDeviceName"), strRelayCardPinMap);
                }
                Gini.SetValue("RCS", $"RCSNormalMode_SelectDAMRelay2", _strRelayCard.Trim());
            }
            else if (intArrayCardPinMapIndex == 2)
            {
                for (int i = 0; i < frmRCS.arrayCard03_NormalModePinMap.Length; i++)
                {
                    frmRCS.arrayCard03_NormalModePinMap[i] = string.Format(Gini.GetValue("Channel_PinName", $"NormalModeCh{i + 1}_DAQPinName").Trim(), Gini.GetValue("Device", "DAQDeviceName"), strRelayCardPinMap);
                }
                Gini.SetValue("RCS", $"RCSNormalMode_SelectDAMRelay3", _strRelayCard.Trim());
            }
            else if (intArrayCardPinMapIndex == 3)
            {
                for (int i = 0; i < frmRCS.arrayCard04_NormalModePinMap.Length; i++)
                {
                    frmRCS.arrayCard04_NormalModePinMap[i] = string.Format(Gini.GetValue("Channel_PinName", $"NormalModeCh{i + 1}_DAQPinName").Trim(), Gini.GetValue("Device", "DAQDeviceName"), strRelayCardPinMap);
                }
                Gini.SetValue("RCS", $"RCSNormalMode_SelectDAMRelay4", _strRelayCard.Trim());
            }
            else if (intArrayCardPinMapIndex == 4)
            {
                for (int i = 0; i < frmRCS.arrayCard05_NormalModePinMap.Length; i++)
                {
                    frmRCS.arrayCard05_NormalModePinMap[i] = string.Format(Gini.GetValue("Channel_PinName", $"NormalModeCh{i + 1}_DAQPinName").Trim(), Gini.GetValue("Device", "DAQDeviceName"), strRelayCardPinMap);
                }
                Gini.SetValue("RCS", $"RCSNormalMode_SelectDAMRelay5", _strRelayCard.Trim());
            }
            else if (intArrayCardPinMapIndex == 5)
            {
                for (int i = 0; i < frmRCS.arrayCard06_NormalModePinMap.Length; i++)
                {
                    frmRCS.arrayCard06_NormalModePinMap[i] = string.Format(Gini.GetValue("Channel_PinName", $"NormalModeCh{i + 1}_DAQPinName").Trim(), Gini.GetValue("Device", "DAQDeviceName"), strRelayCardPinMap);
                }
                Gini.SetValue("RCS", $"RCSNormalMode_SelectDAMRelay6", _strRelayCard.Trim());
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

            strRelayCardPinMap = string.Format(Gini.GetValue("DAM_PinName", $"{_strRelayCard.Replace(" ", "")}_RelayCardNumberDAQPinName".Trim()), Gini.GetValue("Device", "DAQDeviceName").Trim());

            if (intArrayCardPinMapIndex == 0)
            {
                for (int i = 0; i < frmRCS.arrayCard01_WheatstoneModePinMap.Length; i++)
                {
                    frmRCS.arrayCard01_WheatstoneModePinMap[i] = string.Format(Gini.GetValue("Channel_PinName", $"WheatstoneModeCh{i + 1}_DAQPinName").Trim(), Gini.GetValue("Device", "DAQDeviceName"), strRelayCardPinMap);
                }
                Gini.SetValue("RCS", $"RCSWheatstoneMode_SelectDAMRelay1", _strRelayCard.Trim());
            }
            else if (intArrayCardPinMapIndex == 1)
            {
                for (int i = 0; i < frmRCS.arrayCard02_WheatstoneModePinMap.Length; i++)
                {
                    frmRCS.arrayCard02_WheatstoneModePinMap[i] = string.Format(Gini.GetValue("Channel_PinName", $"WheatstoneModeCh{i + 1}_DAQPinName").Trim(), Gini.GetValue("Device", "DAQDeviceName"), strRelayCardPinMap);
                }
                Gini.SetValue("RCS", $"RCSWheatstoneMode_SelectDAMRelay2", _strRelayCard.Trim());
            }
            else if (intArrayCardPinMapIndex == 2)
            {
                for (int i = 0; i < frmRCS.arrayCard03_WheatstoneModePinMap.Length; i++)
                {
                    frmRCS.arrayCard03_WheatstoneModePinMap[i] = string.Format(Gini.GetValue("Channel_PinName", $"WheatstoneModeCh{i + 1}_DAQPinName").Trim(), Gini.GetValue("Device", "DAQDeviceName"), strRelayCardPinMap);
                }
                Gini.SetValue("RCS", $"RCSWheatstoneMode_SelectDAMRelay3", _strRelayCard.Trim());
            }
            else if (intArrayCardPinMapIndex == 3)
            {
                for (int i = 0; i < frmRCS.arrayCard04_WheatstoneModePinMap.Length; i++)
                {
                    frmRCS.arrayCard04_WheatstoneModePinMap[i] = string.Format(Gini.GetValue("Channel_PinName", $"WheatstoneModeCh{i + 1}_DAQPinName").Trim(), Gini.GetValue("Device", "DAQDeviceName"), strRelayCardPinMap);
                }
                Gini.SetValue("RCS", $"RCSWheatstoneMode_SelectDAMRelay4", _strRelayCard.Trim());
            }
            else if (intArrayCardPinMapIndex == 4)
            {
                for (int i = 0; i < frmRCS.arrayCard05_WheatstoneModePinMap.Length; i++)
                {
                    frmRCS.arrayCard05_WheatstoneModePinMap[i] = string.Format(Gini.GetValue("Channel_PinName", $"WheatstoneModeCh{i + 1}_DAQPinName").Trim(), Gini.GetValue("Device", "DAQDeviceName"), strRelayCardPinMap);
                }
                Gini.SetValue("RCS", $"RCSWheatstoneMode_SelectDAMRelay5", _strRelayCard.Trim());
            }
            else if (intArrayCardPinMapIndex == 5)
            {
                for (int i = 0; i < frmRCS.arrayCard06_WheatstoneModePinMap.Length; i++)
                {
                    frmRCS.arrayCard06_WheatstoneModePinMap[i] = string.Format(Gini.GetValue("Channel_PinName", $"WheatstoneModeCh{i + 1}_DAQPinName").Trim(), Gini.GetValue("Device", "DAQDeviceName"), strRelayCardPinMap);
                }
                Gini.SetValue("RCS", $"RCSWheatstoneMode_SelectDAMRelay6", _strRelayCard.Trim());
            }
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
                bool boolNextIndex = false;

                int intCBOIndex = 0, intNextIndex = 0;
                for (int i = 0; i < arrayOldDAM.Length; i++)
                {
                    if (arrayOldDAM[i].Trim() == strSelectDAM.Trim())
                    {
                        intNextIndex = i;
                        boolNextIndex = true;
                    }

                    if (((ComboBox)sender).Name.Trim() == arrayDAM[i].Name.Trim())
                        intCBOIndex = i;
                }

                if (boolNextIndex)
                    arrayDAM[intNextIndex].SelectedItem = arrayOldDAM[intCBOIndex].Trim();
                
                arrayOldDAM[intCBOIndex] = strSelectDAM.Trim(); 
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
            finally
            {
            }
        }
    }
}
