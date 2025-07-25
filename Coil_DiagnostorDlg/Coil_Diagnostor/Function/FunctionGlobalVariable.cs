/*using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Coil_Diagnostor.Function
{
    public class FunctionGlobalVariable
    {
        public string ucPlantName = Properties.Settings.Default.PlantName.Trim();
        public string ucLCRMeter_Addr = Properties.Settings.Default.LCRMeter_Addr.Trim();
        public string ucLCRMeter_ID = Properties.Settings.Default.LCRMeter_ID.Trim();
        public string ucDMM1_Addr = Properties.Settings.Default.TDRMeter_IPAddress.Trim();
        public string ucDMM1_ID = Properties.Settings.Default.TDRMeter_IPPort.Trim();
        public string ucDAQDeviceName = Properties.Settings.Default.DAQDeviceName.Trim();
        public decimal ucRdcDecisionRange_ReferenceValue = Properties.Settings.Default.RdcDecisionRange_ReferenceValue;
        public decimal ucRacDecisionRange_ReferenceValue = Properties.Settings.Default.RacDecisionRange_ReferenceValue;
        public decimal ucLDecisionRange_ReferenceValue = Properties.Settings.Default.LDecisionRange_ReferenceValue;
        public decimal ucCDecisionRange_ReferenceValue = Properties.Settings.Default.CDecisionRange_ReferenceValue;
        public decimal ucQDecisionRange_ReferenceValue = Properties.Settings.Default.QDecisionRange_ReferenceValue;
        public decimal ucEffectiveStandardRangeOfVariation = Properties.Settings.Default.EffectiveStandardRangeOfVariation;
        public int ucWheatstoneDataNumber = Properties.Settings.Default.WheatstoneDataNumber;
        public decimal ucTemperature_ReferenceValue = Properties.Settings.Default.Temperature_ReferenceValue;
        public decimal ucTemperatureUpDown_ReferenceValue = Properties.Settings.Default.TemperatureUpDown_ReferenceValue;

        /// <summary>
        /// 환경설정 값 초기화
        /// </summary>
        public void initialize()
        {
            ucLCRMeter_Addr = "";
            ucLCRMeter_ID = "";
            ucDMM1_Addr = "";
            ucDMM1_ID = "";
            ucDAQDeviceName = "";
            ucRdcDecisionRange_ReferenceValue = 0;
            ucRacDecisionRange_ReferenceValue = 0;
            ucLDecisionRange_ReferenceValue = 0;
            ucCDecisionRange_ReferenceValue = 0;
            ucQDecisionRange_ReferenceValue = 0;
            ucEffectiveStandardRangeOfVariation = 0;
            ucWheatstoneDataNumber = 0;
            ucTemperature_ReferenceValue = 0.00M;
            ucTemperatureUpDown_ReferenceValue = 0.00M;
        }

        /// <summary>
        /// 환경설정 값 변경
        /// </summary>
        public void SetValue()
        {
            Properties.Settings.Default.LCRMeter_Addr = ucLCRMeter_Addr.Trim();
            Properties.Settings.Default.LCRMeter_ID = ucLCRMeter_ID.Trim();
            Properties.Settings.Default.TDRMeter_IPAddress = ucDMM1_Addr.Trim();
            Properties.Settings.Default.TDRMeter_IPPort = ucDMM1_ID.Trim();
            Properties.Settings.Default.DAQDeviceName = ucDAQDeviceName.Trim();
            Properties.Settings.Default.RdcDecisionRange_ReferenceValue = ucRdcDecisionRange_ReferenceValue;
            Properties.Settings.Default.RacDecisionRange_ReferenceValue = ucRacDecisionRange_ReferenceValue;
            Properties.Settings.Default.LDecisionRange_ReferenceValue = ucLDecisionRange_ReferenceValue;
            Properties.Settings.Default.CDecisionRange_ReferenceValue = ucCDecisionRange_ReferenceValue;
            Properties.Settings.Default.QDecisionRange_ReferenceValue = ucQDecisionRange_ReferenceValue;
            Properties.Settings.Default.EffectiveStandardRangeOfVariation = ucEffectiveStandardRangeOfVariation;
            Properties.Settings.Default.WheatstoneDataNumber = ucWheatstoneDataNumber;
            Properties.Settings.Default.Temperature_ReferenceValue = ucTemperature_ReferenceValue;
            Properties.Settings.Default.TemperatureUpDown_ReferenceValue = ucTemperatureUpDown_ReferenceValue;

            Properties.Settings.Default.Save();
        }

        /// <summary>
        /// 환경설정 값 가져오기
        /// </summary>
        public void GetValue()
        {
            ucPlantName = Properties.Settings.Default.PlantName.Trim();
            ucLCRMeter_Addr = Properties.Settings.Default.LCRMeter_Addr.Trim();
            ucLCRMeter_ID = Properties.Settings.Default.LCRMeter_ID.Trim();
            ucDMM1_Addr = Properties.Settings.Default.TDRMeter_IPAddress.Trim();
            ucDMM1_ID = Properties.Settings.Default.TDRMeter_IPPort.Trim();
            ucDAQDeviceName = Properties.Settings.Default.DAQDeviceName.Trim();
            ucRdcDecisionRange_ReferenceValue = Properties.Settings.Default.RdcDecisionRange_ReferenceValue;
            ucRacDecisionRange_ReferenceValue = Properties.Settings.Default.RacDecisionRange_ReferenceValue;
            ucLDecisionRange_ReferenceValue = Properties.Settings.Default.LDecisionRange_ReferenceValue;
            ucCDecisionRange_ReferenceValue = Properties.Settings.Default.CDecisionRange_ReferenceValue;
            ucQDecisionRange_ReferenceValue = Properties.Settings.Default.QDecisionRange_ReferenceValue;
            ucEffectiveStandardRangeOfVariation = Properties.Settings.Default.EffectiveStandardRangeOfVariation;
            ucWheatstoneDataNumber = Properties.Settings.Default.WheatstoneDataNumber;
            ucTemperature_ReferenceValue = Properties.Settings.Default.Temperature_ReferenceValue;
            ucTemperatureUpDown_ReferenceValue = Properties.Settings.Default.TemperatureUpDown_ReferenceValue;
        }
    }
}
*/