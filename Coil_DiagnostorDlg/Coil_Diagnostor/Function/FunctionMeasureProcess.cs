using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO.Ports;
using System.Threading;
using System.Data;
using System.Windows.Forms;
using System.IO;

// DAQ
using NationalInstruments;
using NationalInstruments.DAQmx;

namespace Coil_Diagnostor.Function
{
    #region ▣ Four Card 측정 로직

    public class FunctionMeasureProcess
    {
        //frmRCSDiagnosis frmRCS;
        //frmDRPIDiagnosis frmDRPI;
        
        public Device.ClassDigitalDAQProcess m_DigitalDAQ = new Device.ClassDigitalDAQProcess();
        // 20240429 한인석
        // LCR meter : usb -> GPIB 
        // ClassLCRMeterProcess class not exist function -> delete
        // double -> decimal 
        //public Device.ClassLCRMeterProcess m_LCRMeter = new Device.ClassLCRMeterProcess();
        public Device.ClassLCRMeterGPIBProcess m_LCRMeter = new Device.ClassLCRMeterGPIBProcess();
        public Device.ClassAnalogDAQProcess m_AIReader = new Device.ClassAnalogDAQProcess();

        public FunctionMeasureProcess()
        {
            
        }

        #region ▣ Device 장비 통신 접속 확인

        public bool DaqAndLCRMeterCommunication(string strDeviceName, string strLCRMeter, ref bool boolDAQ, ref bool boolLCRMeter)
        {
            bool boolResult = false;

            try
            {
                boolDAQ = m_DigitalDAQ.DigitalDAQ_Setting(strDeviceName);

                boolLCRMeter = m_LCRMeter.Initialize(strLCRMeter);

                if (boolDAQ && boolLCRMeter)
                    boolResult = true;
                else
                    boolResult = false;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
                boolResult = false;
            }

            return boolResult;
        }

        /// <summary>
        /// DigitalDAQ 장비 통신 접속 확인
        /// </summary>
        /// <returns></returns>
        public bool AccessDigitalDAQChecking(string strDeviceName)
        {
            bool boolResult = false;
            try
            {
                boolResult = m_DigitalDAQ.DigitalDAQ_Setting(strDeviceName);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
                boolResult = false;
            }

            return boolResult;
        }

        /// <summary>
        /// Digital DAQ On/Off
        /// </summary>
        /// <param name="_strDAQPinMap"></param>
        /// <param name="boolPinOnOff"></param>
        public bool DigitalDAQ_ChannelOnOff(string _strDAQPinMap, bool boolPinOnOff)
        {
            bool boolResult = false;

            boolResult = m_DigitalDAQ.MainDigitalDAQ_WriteChannel(_strDAQPinMap, boolPinOnOff);

            return boolResult;
        }

        /// <summary>
        /// Digital DAQ Pin Reader
        /// </summary>
        /// <param name="_strDAQPinMap"></param>
        /// <param name="boolPinOnOff"></param>
        public bool DigitalDAQ_ChannelReader(string _strDAQPinMap)
        {
            bool boolResult = false;

            boolResult = m_DigitalDAQ.DigitalDAQ_ReaderChannel(_strDAQPinMap);

            return boolResult;
        }

        /// <summary>
        /// Digital DAQ All Close 
        /// </summary>
        public void DigitalDAQ_CloseChannel()
        {
            string strDAQPinMap = "{0}/port0/line0,{0}/port0/line1,{0}/port0/line2,{0}/port0/line3,{0}/port0/line4,{0}/port0/line5,{0}/port0/line6,{0}/port0/line7";
            strDAQPinMap = strDAQPinMap + ",{0}/port0/line8,{0}/port0/line9,{0}/port0/line10,{0}/port0/line11,{0}/port0/line12,{0}/port0/line13,{0}/port0/line14,{0}/port0/line15";
            strDAQPinMap = strDAQPinMap + ",{0}/port0/line16,{0}/port0/line17,{0}/port0/line18,{0}/port0/line19,{0}/port0/line20,{0}/port0/line21,{0}/port0/line23";
            strDAQPinMap = strDAQPinMap + ",{0}/port0/line24,{0}/port0/line25,{0}/port0/line26,{0}/port0/line27,{0}/port0/line28,{0}/port0/line29,{0}/port0/line30,{0}/port0/line31";
            strDAQPinMap = strDAQPinMap + ",{0}/port1/line0,{0}/port1/line1,{0}/port1/line2,{0}/port1/line3,{0}/port1/line4,{0}/port1/line5,{0}/port1/line6,{0}/port1/line7";
            strDAQPinMap = strDAQPinMap + ",{0}/port2/line0,{0}/port2/line1,{0}/port2/line2,{0}/port2/line3,{0}/port2/line4,{0}/port2/line5,{0}/port2/line6,{0}/port2/line7";
            //strDAQPinMap = string.Format(strDAQPinMap, Properties.Settings.Default.DAQDeviceName.Trim());
            strDAQPinMap = string.Format(strDAQPinMap, Gini.GetValue("Device", "DAQDeviceName").Trim());

            m_DigitalDAQ.DigitalDAQ_WriteChannel(strDAQPinMap, false);
        }

        #endregion


        #region ▣ 측정 로직

        /// <summary>
        /// LCR-Meter 주파수 설정
        /// </summary>
        /// <param name="dcmFrequency"></param>
        /// <returns></returns>
        public bool functionLCRMeterFrequencySetting(decimal dcmFrequency)
        {
            bool boolResult = false;

            try
            {
                // 주파수 설정
                boolResult = m_LCRMeter.OptFreqSelect(dcmFrequency);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            return boolResult;
        }

        /// <summary>
        /// LCR-Meter VoltageLevel 설정
        /// </summary>
        /// <param name="dcmVoltageLevel"></param>
        /// <returns></returns>
        public bool functionLCRMeterVoltageLevelSetting(decimal dcmVoltageLevel)
        {
            bool boolResult = false;

            try
            {
                // Voltage Level 설정
                boolResult = m_LCRMeter.OptLevelSelect(dcmVoltageLevel);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            return boolResult;
        }

        /// <summary>
        /// LCR-Meter Mode 설정
        /// </summary>
        /// <param name="dcmVoltageLevel"></param>
        /// <returns></returns>
        public bool functionLCRMeterModeSetting(int intMode)
        {
            bool boolResult = false;

            try
            {
                // Voltage Level 설정
                boolResult = m_LCRMeter.OptModeSelect(intMode);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            return boolResult;
        }

        /// <summary>
        /// LCR-Meter Range 설정
        /// </summary>
        /// <param name="dcmRange"></param>
        /// <returns></returns>
        public bool functionLCRMeterRangeSetting(decimal dcmRange)
        {
            bool boolResult = false;

            try
            {
                // Range 설정
                boolResult = m_LCRMeter.OptRangeSelect(dcmRange);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            return boolResult;
        }

        /// <summary>
        /// RCS 일반/휘스톤 모드 Rdc(DC) 저항 측정
        /// </summary>
        /// <param name="_frm"></param>
        /// <param name="dcmMeasurementValue"></param>
        /// <param name="dcmFrequency"></param>
        /// <param name="dcmVoltageLevel"></param>
        /// <returns></returns>
        public bool functionDAQPinMapOnOff(string strDAQPinMap, bool boolDAQPinMapOnOFF)
        {
            bool boolResult = false;

            try
            {
                boolResult = m_DigitalDAQ.DigitalDAQ_WriteChannel(strDAQPinMap, boolDAQPinMapOnOFF);
                Thread.Sleep(100);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
            finally
            {
                if (!boolDAQPinMapOnOFF)
                    DigitalDAQ_CloseChannel();
            }

            return boolResult;
        }

        /// <summary>
        /// RCS 일반/휘스톤 모드 Rdc(DC) 저항 측정
        /// </summary>
        /// <param name="_frm"></param>
        /// <param name="dcmMeasurementValue"></param>
        /// <param name="dcmFrequency"></param>
        /// <param name="dcmVoltageLevel"></param>
        /// <returns></returns>
        public bool functionRCSRdcMeasurement(ref decimal dcmMeasurementValue, bool boolWheatstoneMode, int intSleep)
        {
            bool boolResult = false;
            decimal dcmValue = 0M;

            try
            {
                // DC_R 전환 전송
                m_LCRMeter.OptSelect(Function.FunctionLCRMeterinfo.strCmd_LCRMeter_RDCLsSelect.Trim());
                Thread.Sleep(intSleep);
                
                // 휘스톤 모드일 경우
                if (boolWheatstoneMode)
                {
                    double dMeasurementValue = 0d, dRate = 1000d;
                    int intSamplesPerChannel = 1000;
                    
                    //m_AIReader.startAITaskMultiReader(string.Format(Properties.Settings.Default.WheatstoneAIChannel_DAQPinName.Trim(), Properties.Settings.Default.DAQDeviceName.Trim())
                    //    , -10.00d, 10.00d, ref dMeasurementValue, dRate, intSamplesPerChannel);
                    m_AIReader.startAITaskMultiReader(string.Format(Gini.GetValue("Wheatstone", "WheatstoneAIChannel_DAQPinName").Trim(), Gini.GetValue("Device", "DAQDeviceName").Trim())
                        , -10.00d, 10.00d, ref dMeasurementValue, dRate, intSamplesPerChannel);

                    dcmValue = Convert.ToDecimal(dMeasurementValue);
                    dcmMeasurementValue = SetWheatstoneModeMeasurementValueCalculate(dcmValue);
                }
                else
                {
                    // LCR Meter DC_R 값 읽어 오기
                    dcmValue = m_LCRMeter.LCRReadFunRepeat(1, 100, false);

                    if (dcmValue == 0M)
                    {
                        // LCR Meter DC_R 값 읽어 오기
                        dcmValue = m_LCRMeter.LCRReadFunRepeat(1, 100, false);

                        if (dcmValue == 0M)
                        {
                            // LCR Meter DC_R 값 읽어 오기
                            dcmValue = m_LCRMeter.LCRReadFunRepeat(1, 100, false);
                        }
                    }

                    dcmMeasurementValue = dcmValue;
                }

                boolResult = true;
            }
            catch (Exception ex)
            {
                boolResult = false;
                System.Diagnostics.Debug.Print(ex.Message);
            }

            return boolResult;
        }

        /// <summary>
        /// 휘스톤 모드일 경우 측정 값을 계산
        /// </summary>
        /// <param name="dcmValue"></param>
        /// <returns></returns>
        public decimal SetWheatstoneModeMeasurementValueCalculate(decimal dcmValue)
        {
            decimal dcmResult = 0M;

            try
            {
                //if (Properties.Settings.Default.MeasurementType == 1)
                //{
                //    // 휘스톤 모드 (선간 보정 제외)
                //    dcmResult = (Properties.Settings.Default.WheatstoneMode_Calculate1 - (Properties.Settings.Default.WheatstoneMode_Calculate2) * dcmValue)
                //        / (Properties.Settings.Default.WheatstoneMode_Calculate3 + dcmValue);
                //}
                //else
                //{
                // 휘스톤 모드 (선간 보정 적용)
                string strOffset = Gini.GetValue("Wheatstone", "WheatstoneMod_OffSet");
                decimal dOffSet = string.IsNullOrEmpty(strOffset) ? 0 : Convert.ToDecimal(strOffset);
                dcmValue = dcmValue - dOffSet;
                //dcmResult = (Properties.Settings.Default.WheatstoneMode_Calculate1 - (Properties.Settings.Default.WheatstoneMode_Calculate2) * dcmValue)
                dcmResult = (Convert.ToDecimal(Gini.GetValue("Wheatstone", "WheatstoneMode_Calculate1")) - Convert.ToDecimal(Gini.GetValue("Wheatstone", "WheatstoneMode_Calculate2")) * dcmValue)
                        / (Convert.ToDecimal(Gini.GetValue("Wheatstone", "WheatstoneMode_Calculate3")) + dcmValue);
                //}
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            return dcmResult;
        }

        /// <summary>
        /// RCS 일반/휘스톤 모드 Rac(AC) 저항 측정
        /// </summary>
        /// <param name="_frm"></param>
        /// <param name="dcmMeasurementValue"></param>
        /// <param name="dcmFrequency"></param>
        /// <param name="dcmVoltageLevel"></param>
        /// <returns></returns>
        public bool functionRCSRacMeasurement(ref decimal dcmMeasurementValue)
        {
            bool boolResult = false;

            try
            {
                // AC_R 전환 전송
                m_LCRMeter.OptSelect(Function.FunctionLCRMeterinfo.strCmd_LCRMeter_RACLsSelect.Trim());
                Thread.Sleep(100);

                // LCR Meter AC_R 값 읽어 오기
                dcmMeasurementValue = m_LCRMeter.LCRReadFunRepeat(3, 100, false);

                if (dcmMeasurementValue == 0M)
                {
                    // LCR Meter AC_R 값 읽어 오기
                    dcmMeasurementValue = m_LCRMeter.LCRReadFunRepeat(3, 100, false);
                }

                boolResult = true;
            }
            catch (Exception ex)
            {
                boolResult = false;
                System.Diagnostics.Debug.Print(ex.Message);
            }

            return boolResult;
        }

        /// <summary>
        /// RCS 일반/휘스톤 모드 L 저항 측정
        /// </summary>
        /// <param name="_frm"></param>
        /// <param name="dcmMeasurementValue"></param>
        /// <param name="dcmFrequency"></param>
        /// <param name="dcmVoltageLevel"></param>
        /// <returns></returns>
        public bool functionRCSLMeasurement(ref decimal dcmMeasurementValue)
        {
            bool boolResult = false;

            try
            {
                // L 전환 전송
                m_LCRMeter.OptSelect(Function.FunctionLCRMeterinfo.strCmd_LCRMeter_LsSelect.Trim());
                Thread.Sleep(100);

                // LCR Meter L_R 값 읽어 오기
                dcmMeasurementValue = m_LCRMeter.LCRReadFunRepeat(1, 100, true);

                // LCR Meter DC_R 값 읽어 오기
                dcmMeasurementValue = m_LCRMeter.LCRReadFunRepeat(1, 100, true);

                if (dcmMeasurementValue == 0M)
                {
                    // LCR Meter DC_R 값 읽어 오기
                    dcmMeasurementValue = m_LCRMeter.LCRReadFunRepeat(1, 100, true);
                }

                boolResult = true;
            }
            catch (Exception ex)
            {
                boolResult = false;
                System.Diagnostics.Debug.Print(ex.Message);
            }

            return boolResult;
        }

        /// <summary>
        /// RCS 일반/휘스톤 모드 C 저항 측정
        /// </summary>
        /// <param name="_frm"></param>
        /// <param name="dcmMeasurementValue"></param>
        /// <param name="dcmFrequency"></param>
        /// <param name="dcmVoltageLevel"></param>
        /// <returns></returns>
        public bool functionRCSCMeasurement(ref decimal dcmMeasurementValue)
        {
            bool boolResult = false;

            try
            {
                // C 전환 전송
                m_LCRMeter.OptSelect(Function.FunctionLCRMeterinfo.strCmd_LCRMeter_CpSelect.Trim());
                Thread.Sleep(100);

                // LCR Meter DC_R 값 읽어 오기
                dcmMeasurementValue = m_LCRMeter.LCRReadFunRepeat(1, 0, false);

                // LCR Meter DC_R 값 읽어 오기
                dcmMeasurementValue = m_LCRMeter.LCRReadFunRepeat(1, 100, false);

                if (dcmMeasurementValue == 0M)
                {
                    // LCR Meter DC_R 값 읽어 오기
                    dcmMeasurementValue = m_LCRMeter.LCRReadFunRepeat(1, 100, false);
                }

                boolResult = true;
            }
            catch (Exception ex)
            {
                boolResult = false;
                System.Diagnostics.Debug.Print(ex.Message);
            }

            return boolResult;
        }

        /// <summary>
        /// RCS 일반/휘스톤 모드 Q 저항 측정
        /// </summary>
        /// <param name="_frm"></param>
        /// <param name="dcmMeasurementValue"></param>
        /// <param name="dcmFrequency"></param>
        /// <param name="dcmVoltageLevel"></param>
        /// <returns></returns>
        public bool functionRCSQMeasurement(ref decimal dcmMeasurementValue)
        {
            bool boolResult = false;

            try
            {
                // Q 전환 전송
                m_LCRMeter.OptSelect(Function.FunctionLCRMeterinfo.strCmd_LCRMeter_QSelect.Trim());
                Thread.Sleep(100);

                // LCR Meter DC_R 값 읽어 오기
                dcmMeasurementValue = m_LCRMeter.LCRReadFunRepeat(1, 0, false);

                boolResult = true;
            }
            catch (Exception ex)
            {
                boolResult = false;
                System.Diagnostics.Debug.Print(ex.Message);
            }

            return boolResult;
        }

        #endregion

        #region ▣ 화면 공통 적용 함수

        /// <summary>
        /// 기준값 기준으로 측정값 결과를 판단 로직
        /// </summary>
        /// <param name="_dcmMeasurementValue"></param>
        /// <param name="_dcmReferenceValue"></param>
        /// <param name="_dcmDecisionRange"></param>
        /// <param name="_dcmEffectiveStandardRange"></param>
        /// <returns></returns>
        public string GetMesurmentValueDecision(decimal _dcmMeasurementValue, decimal _dcmReferenceValue, decimal _dcmDecisionRange, decimal _dcmEffectiveStandardRange
            , string _strItem)
        {
            string strResult = "";

            decimal dcmDecisionRangeMaxValue = _dcmReferenceValue + (_dcmReferenceValue * (_dcmDecisionRange / 100));
            decimal dcmEffectiveStandardRangeMaxValue = _dcmReferenceValue + (_dcmReferenceValue * (_dcmEffectiveStandardRange / 100));

            if (_strItem.Trim() == "C")
            {
                dcmDecisionRangeMaxValue = Math.Truncate(dcmDecisionRangeMaxValue * 1000000) / 1000000;
                dcmEffectiveStandardRangeMaxValue = Math.Truncate(dcmEffectiveStandardRangeMaxValue * 1000000) / 1000000;
            }
            else
            {
                dcmDecisionRangeMaxValue = Math.Truncate(dcmDecisionRangeMaxValue * 1000) / 1000;
                dcmEffectiveStandardRangeMaxValue = Math.Truncate(dcmEffectiveStandardRangeMaxValue * 1000) / 1000;
            }

            decimal dcmDecisionRangeMinValue = _dcmReferenceValue - (_dcmReferenceValue * (_dcmDecisionRange / 100));
            decimal dcmEffectiveStandardRangeMinValue = _dcmReferenceValue - (_dcmReferenceValue * (_dcmEffectiveStandardRange / 100));

            if (_strItem.Trim() == "C")
            {
                dcmDecisionRangeMinValue = Math.Truncate(dcmDecisionRangeMinValue * 1000000) / 1000000;
                dcmEffectiveStandardRangeMinValue = Math.Truncate(dcmEffectiveStandardRangeMinValue * 1000000) / 1000000;
            }
            else
            {
                dcmDecisionRangeMinValue = Math.Truncate(dcmDecisionRangeMinValue * 1000) / 1000;
                dcmEffectiveStandardRangeMinValue = Math.Truncate(dcmEffectiveStandardRangeMinValue * 1000) / 1000;
            }

            if (_dcmMeasurementValue < 0)
                _dcmMeasurementValue = -_dcmMeasurementValue;

            if (dcmDecisionRangeMaxValue < 0)
                dcmDecisionRangeMaxValue = -dcmDecisionRangeMaxValue;

            if (dcmEffectiveStandardRangeMaxValue < 0)
                dcmEffectiveStandardRangeMaxValue = -dcmEffectiveStandardRangeMaxValue;

            if (dcmDecisionRangeMinValue < 0)
                dcmDecisionRangeMinValue = -dcmDecisionRangeMinValue;

            if (dcmEffectiveStandardRangeMinValue < 0)
                dcmEffectiveStandardRangeMinValue = -dcmEffectiveStandardRangeMinValue;
            
            if (_dcmMeasurementValue != 0 && dcmDecisionRangeMaxValue != 0)
            {
                if (_dcmMeasurementValue >= dcmDecisionRangeMaxValue)
                    strResult = "부적합";
                else if (_dcmMeasurementValue <= dcmDecisionRangeMinValue)
                    strResult = "부적합";
                else if (_dcmMeasurementValue < dcmDecisionRangeMaxValue && _dcmMeasurementValue >= dcmEffectiveStandardRangeMaxValue)
                    strResult = "의심";
                else if (_dcmMeasurementValue > dcmDecisionRangeMinValue && _dcmMeasurementValue <= dcmEffectiveStandardRangeMinValue)
                    strResult = "의심h";
                else
                    strResult = "적합";
            }
            else
            {
                if (_dcmMeasurementValue < dcmDecisionRangeMaxValue && _dcmMeasurementValue >= dcmEffectiveStandardRangeMaxValue)
                    strResult = "의심";
                else if (_dcmMeasurementValue > dcmDecisionRangeMinValue && _dcmMeasurementValue <= dcmEffectiveStandardRangeMinValue)
                    strResult = "의심";
                else if (_dcmMeasurementValue == 0 && dcmDecisionRangeMinValue > 0)
                    strResult = "부적합";
                else 
                    strResult = "적합";
            }

            return strResult;
        }

        #endregion
    }
    #endregion
}
