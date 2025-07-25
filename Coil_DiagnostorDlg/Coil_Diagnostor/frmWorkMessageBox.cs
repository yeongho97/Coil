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
    public partial class frmWorkMessageBox : Form
    {
        public bool boolOk = false;

        public frmWorkMessageBox()
        {
            CheckForIllegalCrossThreadCalls = false;
            InitializeComponent();
        }

        private void btnConfirm_Click(object sender, EventArgs e)
        {
            boolOk = true;
            this.Close();
        }

        private void frmWorkMessageBox_Load(object sender, EventArgs e)
        {
            boolOk = false;
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            boolOk = false;
            this.Close();
        }
    }
}
