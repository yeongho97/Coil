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
    public partial class frmFilePathSetting : Form
    {
        public frmFilePathSetting()
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
                case Keys.F1: // 환경 설정 버튼
                    btnFilePathSetting.PerformClick();
                    break;
                case Keys.Control & Keys.F12: // 닫기 버튼
                    btnClose.PerformClick();
                    break;
            }

            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void frmFilePathSetting_Load(object sender, EventArgs e)
        {
            teFilePathSetting.Text = Gini.GetValue("Device", "FilePathSetting").Trim();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnFilePathSetting_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog ofd = new OpenFileDialog();
                ofd.Filter = "Access File(*.exe)|*.exe|AllFiles(*.*)|*.*";
                ofd.Title = "Access DB 파일을 선택해 주십시오";

                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    teFilePathSetting.Text = ofd.FileName.Trim();
                }

                Gini.SetValue("Device", "FilePathSetting", teFilePathSetting.Text.Trim());
      
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
        }
    }
}
