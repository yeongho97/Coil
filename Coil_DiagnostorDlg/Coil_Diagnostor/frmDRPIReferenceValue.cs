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
    public partial class frmDRPIReferenceValue : Form
    {
        frmDRPIDiagnosis frmDRPI;

        public frmDRPIReferenceValue()
        {
            InitializeComponent();
        }

        public frmDRPIReferenceValue(frmDRPIDiagnosis _frm)
        {
            InitializeComponent();

            frmDRPI = _frm;
        }

        /// <summary>
        /// Form Load
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmDRPIReferenceValue_Load(object sender, EventArgs e)
        {
            teControlA1Rdc_DRPIReferenceValue.Text = frmDRPI.dA_ControlRodRdc[0] == null ? "0.000" : frmDRPI.dA_ControlRodRdc[0].ToString("F3");
            teControlA1Rac_DRPIReferenceValue.Text = frmDRPI.dA_ControlRodRac[0] == null ? "0.000" : frmDRPI.dA_ControlRodRac[0].ToString("F3");
            teControlA1L_DRPIReferenceValue.Text = frmDRPI.dA_ControlRodL[0] == null ? "0.000" : frmDRPI.dA_ControlRodL[0].ToString("F3");
            teControlA1C_DRPIReferenceValue.Text = frmDRPI.dA_ControlRodC[0] == null ? "0.000000" : frmDRPI.dA_ControlRodC[0].ToString("F6");
            teControlA1Q_DRPIReferenceValue.Text = frmDRPI.dA_ControlRodQ[0] == null ? "0.000" : frmDRPI.dA_ControlRodQ[0].ToString("F3");

            teControlA2Rdc_DRPIReferenceValue.Text = frmDRPI.dA_ControlRodRdc[1] == null ? "0.000" : frmDRPI.dA_ControlRodRdc[1].ToString("F3");
            teControlA2Rac_DRPIReferenceValue.Text = frmDRPI.dA_ControlRodRac[1] == null ? "0.000" : frmDRPI.dA_ControlRodRac[1].ToString("F3");
            teControlA2L_DRPIReferenceValue.Text = frmDRPI.dA_ControlRodL[1] == null ? "0.000" : frmDRPI.dA_ControlRodL[1].ToString("F3");
            teControlA2C_DRPIReferenceValue.Text = frmDRPI.dA_ControlRodC[1] == null ? "0.000000" : frmDRPI.dA_ControlRodC[1].ToString("F6");
            teControlA2Q_DRPIReferenceValue.Text = frmDRPI.dA_ControlRodQ[1] == null ? "0.000" : frmDRPI.dA_ControlRodQ[1].ToString("F3");

            teControlARdc_DRPIReferenceValue.Text = frmDRPI.dA_ControlRodRdc[2] == null ? "0.000" : frmDRPI.dA_ControlRodRdc[2].ToString("F3");
            teControlARac_DRPIReferenceValue.Text = frmDRPI.dA_ControlRodRac[2] == null ? "0.000" : frmDRPI.dA_ControlRodRac[2].ToString("F3");
            teControlAL_DRPIReferenceValue.Text = frmDRPI.dA_ControlRodL[2] == null ? "0.000" : frmDRPI.dA_ControlRodL[2].ToString("F3");
            teControlAC_DRPIReferenceValue.Text = frmDRPI.dA_ControlRodC[2] == null ? "0.000000" : frmDRPI.dA_ControlRodC[2].ToString("F6");
            teControlAQ_DRPIReferenceValue.Text = frmDRPI.dA_ControlRodQ[2] == null ? "0.000" : frmDRPI.dA_ControlRodQ[2].ToString("F3");

            teStopA1Rdc_DRPIReferenceValue.Text = frmDRPI.dA_StopRdc[0] == null ? "0.000" : frmDRPI.dA_StopRdc[0].ToString("F3");
            teStopA1Rac_DRPIReferenceValue.Text = frmDRPI.dA_StopRac[0] == null ? "0.000" : frmDRPI.dA_StopRac[0].ToString("F3");
            teStopA1L_DRPIReferenceValue.Text = frmDRPI.dA_StopL[0] == null ? "0.000" : frmDRPI.dA_StopL[0].ToString("F3");
            teStopA1C_DRPIReferenceValue.Text = frmDRPI.dA_StopC[0] == null ? "0.000000" : frmDRPI.dA_StopC[0].ToString("F6");
            teStopA1Q_DRPIReferenceValue.Text = frmDRPI.dA_StopQ[0] == null ? "0.000" : frmDRPI.dA_StopQ[0].ToString("F3");

            teStopA2Rdc_DRPIReferenceValue.Text = frmDRPI.dA_StopRdc[1] == null ? "0.000" : frmDRPI.dA_StopRdc[1].ToString("F3");
            teStopA2Rac_DRPIReferenceValue.Text = frmDRPI.dA_StopRac[1] == null ? "0.000" : frmDRPI.dA_StopRac[1].ToString("F3");
            teStopA2L_DRPIReferenceValue.Text = frmDRPI.dA_StopL[1] == null ? "0.000" : frmDRPI.dA_StopL[1].ToString("F3");
            teStopA2C_DRPIReferenceValue.Text = frmDRPI.dA_StopC[1] == null ? "0.000000" : frmDRPI.dA_StopC[1].ToString("F6");
            teStopA2Q_DRPIReferenceValue.Text = frmDRPI.dA_StopQ[1] == null ? "0.000" : frmDRPI.dA_StopQ[1].ToString("F3");

            teStopARdc_DRPIReferenceValue.Text = frmDRPI.dA_StopRdc[2] == null ? "0.000" : frmDRPI.dA_StopRdc[2].ToString("F3");
            teStopARac_DRPIReferenceValue.Text = frmDRPI.dA_StopRac[2] == null ? "0.000" : frmDRPI.dA_StopRac[2].ToString("F3");
            teStopAL_DRPIReferenceValue.Text = frmDRPI.dA_StopL[2] == null ? "0.000" : frmDRPI.dA_StopL[2].ToString("F3");
            teStopAC_DRPIReferenceValue.Text = frmDRPI.dA_StopC[2] == null ? "0.000000" : frmDRPI.dA_StopC[2].ToString("F6");
            teStopAQ_DRPIReferenceValue.Text = frmDRPI.dA_StopQ[2] == null ? "0.000" : frmDRPI.dA_StopQ[2].ToString("F3");

            teControlB1Rdc_DRPIReferenceValue.Text = frmDRPI.dB_ControlRodRdc[0] == null ? "0.000" : frmDRPI.dB_ControlRodRdc[0].ToString("F3");
            teControlB1Rac_DRPIReferenceValue.Text = frmDRPI.dB_ControlRodRac[0] == null ? "0.000" : frmDRPI.dB_ControlRodRac[0].ToString("F3");
            teControlB1L_DRPIReferenceValue.Text = frmDRPI.dB_ControlRodL[0] == null ? "0.000" : frmDRPI.dB_ControlRodL[0].ToString("F3");
            teControlB1C_DRPIReferenceValue.Text = frmDRPI.dB_ControlRodC[0] == null ? "0.000000" : frmDRPI.dB_ControlRodC[0].ToString("F6");
            teControlB1Q_DRPIReferenceValue.Text = frmDRPI.dB_ControlRodQ[0] == null ? "0.000" : frmDRPI.dB_ControlRodQ[0].ToString("F3");

            teControlB2Rdc_DRPIReferenceValue.Text = frmDRPI.dB_ControlRodRdc[1] == null ? "0.000" : frmDRPI.dB_ControlRodRdc[1].ToString("F3");
            teControlB2Rac_DRPIReferenceValue.Text = frmDRPI.dB_ControlRodRac[1] == null ? "0.000" : frmDRPI.dB_ControlRodRac[1].ToString("F3");
            teControlB2L_DRPIReferenceValue.Text = frmDRPI.dB_ControlRodL[1] == null ? "0.000" : frmDRPI.dB_ControlRodL[1].ToString("F3");
            teControlB2C_DRPIReferenceValue.Text = frmDRPI.dB_ControlRodC[1] == null ? "0.000000" : frmDRPI.dB_ControlRodC[1].ToString("F6");
            teControlB2Q_DRPIReferenceValue.Text = frmDRPI.dB_ControlRodQ[1] == null ? "0.000" : frmDRPI.dB_ControlRodQ[1].ToString("F3");

            teControlBRdc_DRPIReferenceValue.Text = frmDRPI.dB_ControlRodRdc[2] == null ? "0.000" : frmDRPI.dB_ControlRodRdc[2].ToString("F3");
            teControlBRac_DRPIReferenceValue.Text = frmDRPI.dB_ControlRodRac[2] == null ? "0.000" : frmDRPI.dB_ControlRodRac[2].ToString("F3");
            teControlBL_DRPIReferenceValue.Text = frmDRPI.dB_ControlRodL[2] == null ? "0.000" : frmDRPI.dB_ControlRodL[2].ToString("F3");
            teControlBC_DRPIReferenceValue.Text = frmDRPI.dB_ControlRodC[2] == null ? "0.000000" : frmDRPI.dB_ControlRodC[2].ToString("F6");
            teControlBQ_DRPIReferenceValue.Text = frmDRPI.dB_ControlRodQ[2] == null ? "0.000" : frmDRPI.dB_ControlRodQ[2].ToString("F3");

            teStopB1Rdc_DRPIReferenceValue.Text = frmDRPI.dB_StopRdc[0] == null ? "0.000" : frmDRPI.dB_StopRdc[0].ToString("F3");
            teStopB1Rac_DRPIReferenceValue.Text = frmDRPI.dB_StopRac[0] == null ? "0.000" : frmDRPI.dB_StopRac[0].ToString("F3");
            teStopB1L_DRPIReferenceValue.Text = frmDRPI.dB_StopL[0] == null ? "0.000" : frmDRPI.dB_StopL[0].ToString("F3");
            teStopB1C_DRPIReferenceValue.Text = frmDRPI.dB_StopC[0] == null ? "0.000000" : frmDRPI.dB_StopC[0].ToString("F6");
            teStopB1Q_DRPIReferenceValue.Text = frmDRPI.dB_StopQ[0] == null ? "0.000" : frmDRPI.dB_StopQ[0].ToString("F3");

            teStopB2Rdc_DRPIReferenceValue.Text = frmDRPI.dB_StopRdc[1] == null ? "0.000" : frmDRPI.dB_StopRdc[1].ToString("F3");
            teStopB2Rac_DRPIReferenceValue.Text = frmDRPI.dB_StopRac[1] == null ? "0.000" : frmDRPI.dB_StopRac[1].ToString("F3");
            teStopB2L_DRPIReferenceValue.Text = frmDRPI.dB_StopL[1] == null ? "0.000" : frmDRPI.dB_StopL[1].ToString("F3");
            teStopB2C_DRPIReferenceValue.Text = frmDRPI.dB_StopC[1] == null ? "0.000000" : frmDRPI.dB_StopC[1].ToString("F6");
            teStopB2Q_DRPIReferenceValue.Text = frmDRPI.dB_StopQ[1] == null ? "0.000" : frmDRPI.dB_StopQ[1].ToString("F3");

            teStopBRdc_DRPIReferenceValue.Text = frmDRPI.dB_StopRdc[2] == null ? "0.000" : frmDRPI.dB_StopRdc[2].ToString("F3");
            teStopBRac_DRPIReferenceValue.Text = frmDRPI.dB_StopRac[2] == null ? "0.000" : frmDRPI.dB_StopRac[2].ToString("F3");
            teStopBL_DRPIReferenceValue.Text = frmDRPI.dB_StopL[2] == null ? "0.000" : frmDRPI.dB_StopL[2].ToString("F3");
            teStopBC_DRPIReferenceValue.Text = frmDRPI.dB_StopC[2] == null ? "0.000000" : frmDRPI.dB_StopC[2].ToString("F6");
            teStopBQ_DRPIReferenceValue.Text = frmDRPI.dB_StopQ[2] == null ? "0.000" : frmDRPI.dB_StopQ[2].ToString("F3");
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
    }
}
