using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO.Ports;
using System.Data;
using System.IO;
using NationalInstruments.VisaNS;
using Coil_Diagnostor.Function;
using NationalInstruments.NI4882;

namespace Coil_Diagnostor.Device
{
    public class ClassLCRMeterProcess
    {
        private ResourceManager rmLCR = ResourceManager.GetLocalManager();
        private MessageBasedSession ecLCR;

        protected string m_strID;                       //장비의 ID 스트링
        protected string m_strAddr;
        public bool m_bConnected = false;
        protected bool m_bStatus = false;
        protected string[] strResNames;

        public ClassLCRMeterProcess()
        {
            m_strID = Gini.GetValue("Device", "LCRMeter_ID");
        }

        public bool Initialize(string address)
        {
            if (m_bConnected == true)
                return false;

            try
            {
                m_strAddr = address;

                strResNames = rmLCR.FindResources(m_strAddr);
                ecLCR = new MessageBasedSession(strResNames[0]);

                // Set to the Remote State
                new UsbSession(ecLCR.ResourceName).ControlRen(RenMode.Assert);

                // Clear the device
                ecLCR.Clear();
                // Clear the error status
                ecLCR.Write("*CLS");
                ecLCR.Write("*RST");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
                m_bConnected = false;
                return false;
            }

            if (GetID() == m_strID.Substring(0, 8))
                m_bConnected = true;
            else
                m_bConnected = false;

            return true;
        }

        public void Close()
        {
            // Clear the Remote state
            if (m_bConnected == true)
            {
                new UsbSession(ecLCR.ResourceName).ControlRen(RenMode.Deassert);

                ecLCR.Terminate();
                ecLCR.Dispose();
                m_bConnected = false;
            }
        }

        /// <summary>
        /// 장비의 ID를 가져옵니다.
        /// </summary>
        /// <returns></returns>
 
        public string GetID()
        {
            string strRet = "";

            try
            {
                strRet = ecLCR.Query("*IDN?");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            return strRet.Substring(0, 8);
        }

        public bool IsOpen
        {
            get { return m_bConnected; }
        }

        public string DeviceAddress
        {
            get { return m_strAddr; }
            set { m_strAddr = value; }
        }

        public bool IsSelfTestOK()
        {
            if (GetID() == m_strID)
                m_bConnected = true;
            else
                m_bConnected = false;

            return m_bConnected;
        }
               
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public bool OptSelect(string _strCommand)          
        {
            string strData;

            if (m_bConnected == false)
                return false;

            strData = String.Format(_strCommand.Trim());
            
            try
            {
                ecLCR.Write(strData);

                System.Threading.Thread.Sleep(500);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
                return false;
            }

            return true;
        }

        /// <summary>
        /// LCR 데이터 반복 읽기
        /// </summary>
        /// <param name="nCnt"></param>
        /// <returns></returns>
        public decimal LCRReadFunRepeat(int nCnt, int nSleep, bool _boolCapacitance)          
        {
            double dSum = 0;
            decimal dValue = 0M;
            string strRet = "";
            int tempFirstInd = 0;
            string strFirstVal = "";

            try
            {
                System.Threading.Thread.Sleep(nSleep);

                for (int i = 0; i < nCnt; i++)
                {
                    ecLCR.Write(Function.FunctionLCRMeterinfo.strCmd_LCRMeter_FETCHOutp.Trim());
                    strRet = ecLCR.ReadString();
                    tempFirstInd = strRet.IndexOf(",");

                    string[] arrayRet = strRet.Split(',');

                    // data example : +2.03556E-5,+7.98219E-3,+0
                    if (_boolCapacitance)
                    {
                        strFirstVal = arrayRet[0].Substring(0, tempFirstInd);

                        dSum = dSum + Convert.ToDouble(strFirstVal.Trim()) * 1000;
                    }
                    else
                    {
                        strFirstVal = arrayRet[1].Substring(0, tempFirstInd);

                        dSum = dSum + Convert.ToDouble(strFirstVal.Trim());
                    }
                }

                dValue = Convert.ToDecimal(dSum) / Convert.ToDecimal(nCnt);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            return dValue;
        }

        /// <summary>
        /// 주파수 설정
        /// </summary>
        /// <param name="nIdx"></param>
        /// <returns></returns>
        public bool OptFreqSelect(decimal nIdx)   
        {
            string strData;

            if (m_bConnected == false)
                return false;

            strData = String.Format(Function.FunctionLCRMeterinfo.strCmd_LCRMeter_FREQuency.Trim(), nIdx);

            try
            {
                ecLCR.Write(strData);
                
                System.Threading.Thread.Sleep(10);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
                return false;
            }

            return true;
        }

        /// <summary>
        /// 전압레벨 설정
        /// </summary>
        /// <param name="nIdx"></param>
        /// <returns></returns>
        public bool OptLevelSelect(decimal nIdx)   
        {
            string strData;

            if (m_bConnected == false)
                return false;

            strData = String.Format(Function.FunctionLCRMeterinfo.strCmd_LCRMeter_VOLTage.Trim(), nIdx);

            try
            {
                ecLCR.Write(strData);

                System.Threading.Thread.Sleep(10);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
                return false;
            }

            return true;            
        }

        /// <summary>
        /// Mode 설정
        /// </summary>
        /// <param name="nIdx"></param>
        /// <returns></returns>
        public bool OptModeSelect(int nIdx)
        {
            string strData;

            if (m_bConnected == false)
                return false;

            strData = String.Format(Function.FunctionLCRMeterinfo.strCmd_LCRMeter_APERMED.Trim(), nIdx);

            try
            {
                ecLCR.Write(strData);

                System.Threading.Thread.Sleep(10);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
                return false;
            }

            return true;
        }

        /// <summary>
        /// Range 설정
        /// </summary>
        /// <param name="nIdx"></param>
        /// <returns></returns>
        public bool OptRangeSelect(decimal nIdx)
        {
            string strData;

            if (m_bConnected == false)
                return false;

            strData = String.Format(Function.FunctionLCRMeterinfo.strCmd_LCRMeter_Range.Trim(), nIdx);

            try
            {
                ecLCR.Write(strData);

                System.Threading.Thread.Sleep(10);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
                return false;
            }

            return true;
        }
	}

    public class ClassLCRMeterGPIBProcess
    {
        public string m_strID;                       //장비의 ID 스트링
        public string m_strAddr;
        public bool m_bConnected = false;
        public bool m_bStatus = false;

        private NationalInstruments.NI4882.Device LCRdevice;
        int BoardID = 0;
        byte PriAddress = 0;

        public ClassLCRMeterGPIBProcess()
        {
            //m_strID = frmMain.m_Global.ucLCRMeter_ID.Trim();
            m_strID = Gini.GetValue("Device", "LCRMeter_ID").Trim();
        }

        public ClassLCRMeterGPIBProcess(string _strID)
        {
            m_strID = _strID.Trim();
        }

        public bool Initialize(string address)
        {
            try
            {
                m_strAddr = address;

                if (address.Length > 0)
                {
                    int tempBoard = address.IndexOf(":");

                    string strBoard = address.Substring(0, tempBoard);
                    string strBoard2 = strBoard.Replace("GPIB", "");

                    if (strBoard2.Length > 0)
                        BoardID = System.Convert.ToInt16(strBoard2);

                    int tempPriAdd = address.LastIndexOf(":");
                    string strPriAdd;

                    if (tempPriAdd - 1 - (tempBoard + 2) > 0)
                    {
                        strPriAdd = address.Substring(tempBoard + 2, tempPriAdd - 1 - (tempBoard + 2));
                        PriAddress = System.Convert.ToByte(strPriAdd);
                    }
                }

                LCRdevice = new NationalInstruments.NI4882.Device(BoardID, PriAddress, 0);
            }
            catch (IOException ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
                m_bConnected = false;
                return false;
            }

            if (GetID() == m_strID.Substring(0, 8))
                m_bConnected = true;
            else
                m_bConnected = false;

            return m_bConnected;
        }

        public void Close()
        {
            if (LCRdevice != null)
            {
                LCRdevice.Dispose();
            }
        }

        // 장비의 ID를 가져옵니다.
        public string GetID()
        {
            string strRet = "";

            try
            {
                LCRdevice.Write("*IDN?");
                System.Threading.Thread.Sleep(500);
                strRet = LCRdevice.ReadString();

                if (strRet.Trim() != "" && strRet.Length > 8)
                    strRet = strRet.Substring(0, 8);
            }
            catch (IOException ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
            // 20240531 한인석 : GPIB Exception 처리 추가
            catch (InvalidOperationException ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            return strRet;
        }

        public bool IsOpen
        {
            get { return m_bConnected; }
        }

        public string DeviceAddress
        {
            get { return m_strAddr; }
            set { m_strAddr = value; }
        }

        public bool IsSelfTestOK()
        {
            string strGetID = GetID();
            if (strGetID.Length > 0 && strGetID.Substring(0, 7) == m_strID.Substring(0, 7))
                m_bConnected = true;
            else
                m_bConnected = false;

            return m_bConnected;
        }

        /// <summary>
        /// AUTO Range 설정
        /// </summary>
        /// <returns></returns>
        //public bool OptAUTORangeSelect()
        //{
        //    string strData;

        //    if (m_bConnected == false)
        //        return false;

        //    strData = String.Format(FunctionLCRMeterinfo.strCmd_LCRMeter_AUTORange.Trim());

        //    try
        //    {
        //        LCRdevice.Write(strData);

        //        System.Threading.Thread.Sleep(10);
        //    }
        //    catch (IOException ex)
        //    {
        //        System.Diagnostics.Debug.Print(ex.Message);
        //        return false;
        //    }

        //    return true;
        //}

        /// <summary>
        /// AUTO DCI ISO AUTO 설정
        /// </summary>
        /// <param name="_strOnOff"></param>
        /// <returns></returns>
        //public bool OptAUTODCIISOSelect(string _strOnOff)
        //{
        //    string strData;

        //    if (m_bConnected == false)
        //        return false;

        //    strData = String.Format(FunctionLCRMeterinfo.strCmd_LCRMeter_AUTODCIISO.Trim(), _strOnOff.Trim());

        //    try
        //    {
        //        LCRdevice.Write(strData);

        //        System.Threading.Thread.Sleep(10);
        //    }
        //    catch (IOException ex)
        //    {
        //        System.Diagnostics.Debug.Print(ex.Message);
        //        return false;
        //    }

        //    return true;
        //}

        /// <summary>
        /// AUTO DCI RNG AUTO 설정
        /// </summary>
        /// <param name="_strOnOff"></param>
        /// <returns></returns>
        //public bool OptAUTODCIRNGSelect(string _strOnOff)
        //{
        //    string strData;

        //    if (m_bConnected == false)
        //        return false;

        //    strData = String.Format(FunctionLCRMeterinfo.strCmd_LCRMeter_AUTODCIRNG.Trim(), _strOnOff.Trim());

        //    try
        //    {
        //        LCRdevice.Write(strData);

        //        System.Threading.Thread.Sleep(10);
        //    }
        //    catch (IOException ex)
        //    {
        //        System.Diagnostics.Debug.Print(ex.Message);
        //        return false;
        //    }

        //    return true;
        //}

        /// <summary>
        /// LCR 데이터 반복 읽기
        /// </summary>
        /// <param name="nCnt"></param>
        /// <returns></returns>
        //public double LCRReadFunRepeat(int nCnt, int nSleep, bool _boolCapacitance)
        public decimal LCRReadFunRepeat(int nCnt, int nSleep, bool _boolCapacitance)
        {
            double dSum = 0;
            //double dValue = 0d;
            decimal dValue = 0m;
            string strRet = "";
            int tempFirstInd = 0;
            string strFirstVal = "";

            try
            {
                System.Threading.Thread.Sleep(nSleep);

                for (int i = 0; i < nCnt; i++)
                {
                    LCRdevice.Write(FunctionLCRMeterinfo.strCmd_LCRMeter_FETCHOutp.Trim());
                    strRet = LCRdevice.ReadString();
                    tempFirstInd = strRet.IndexOf(",");

                    string[] arrayRet = strRet.Split(',');

                    if (_boolCapacitance)
                    {
                        strFirstVal = arrayRet[0].Substring(0, tempFirstInd);

                        dSum = dSum + Convert.ToDouble(strFirstVal.Trim()) * 1000;
                    }
                    else
                    {
                        strFirstVal = arrayRet[1].Substring(0, tempFirstInd);

                        dSum = dSum + Convert.ToDouble(strFirstVal.Trim());
                    }

                    System.Diagnostics.Debug.WriteLine(DateTime.Now.ToString("ss.fff"));
                }

                dValue = Convert.ToDecimal(dSum) / Convert.ToDecimal(nCnt);
            }
            catch (IOException ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
            // 20240531 한인석 : GPIB Exception 처리 추가
            catch (InvalidOperationException ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            return dValue;
        }

        /// <summary>
        /// LCR 데이터 반복 읽기
        /// </summary>
        /// <param name="nCnt"></param>
        /// <returns></returns>
        public double LCRReadFunRepeatOne(int nCnt, int nSleep, bool _boolCapacitance)
        {
            double dSum = 0;
            double dValue = 0d;
            string strRet = "";
            int tempFirstInd = 0;
            string strFirstVal = "";

            try
            {
                System.Threading.Thread.Sleep(nSleep);
                double[] dDataValue = new double[nCnt];
                int nCount = 0;

                for (int i = 0; i < nCnt; i++)
                {
                    LCRdevice.Write(FunctionLCRMeterinfo.strCmd_LCRMeter_FETCHOutp.Trim());
                    strRet = LCRdevice.ReadString();
                    tempFirstInd = strRet.IndexOf(",");

                    string[] arrayRet = strRet.Split(',');

                    if (_boolCapacitance)
                    {
                        strFirstVal = arrayRet[0].Substring(0, tempFirstInd);

                        dDataValue[i] = Convert.ToDouble(strFirstVal.Trim()) * 1000;
                    }
                    else
                    {
                        strFirstVal = arrayRet[1].Substring(0, tempFirstInd);

                        dDataValue[i] = Convert.ToDouble(strFirstVal.Trim());
                    }

                    dSum = dSum + dDataValue[i];
                    nCount++;
                }

                dValue = dSum / nCount;
            }
            catch (IOException ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
            // 20240531 한인석 : GPIB Exception 처리 추가
            catch (InvalidOperationException ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            return dValue;
        }

        /// <summary>
        /// Range 설정
        /// </summary>
        /// <param name="nIdx"></param>
        /// <returns></returns>
        public bool OptRangeSelect(decimal nIdx)
        {
            string strData;

            if (m_bConnected == false)
                return false;

            strData = String.Format(FunctionLCRMeterinfo.strCmd_LCRMeter_Range.Trim(), nIdx);

            try
            {
                LCRdevice.Write(strData);

                System.Threading.Thread.Sleep(10);
            }
            catch (IOException ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
                return false;
            }
            // 20240531 한인석 : GPIB Exception 처리 추가
            catch (InvalidOperationException ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
            return true;
        }

        public bool OptZSelect()           //Z
        {
            string strData;

            if (m_bConnected == false)
                return false;

            strData = String.Format(":FUNCtion:IMPedance ZTD");

            try
            {
                LCRdevice.Write(strData);

                System.Threading.Thread.Sleep(500);
            }
            catch (IOException ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
                return false;
            }
            // 20240531 한인석 : GPIB Exception 처리 추가
            catch (GpibException ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
            return true;
        }

        public double LCRReadFunRepeat(int nCnt)          //LCR 데이터 반복 읽기
        {
            double dSum = 0, dResultValue = 0;
            string strRet = "";
            int tempFirstInd = 0, nSumCount = 0;
            string strFirstVal = "";

            try
            {
                for (int i = 0; i < nCnt; i++)
                {
                    LCRdevice.Write("FETCH?");
                    System.Threading.Thread.Sleep(50);
                    strRet = LCRdevice.ReadString();

                    tempFirstInd = strRet.IndexOf(",");

                    strFirstVal = strRet.Substring(0, tempFirstInd);

                    if (strFirstVal.Trim() != "" && Convert.ToDouble(strFirstVal.Trim()) < 90000000000.00)
                    {
                        dSum = dSum + Convert.ToDouble(strFirstVal.Trim());
                        nSumCount++;
                    }
                }

                if (nSumCount > 0)
                    dResultValue = dSum / nSumCount;
                else
                    dResultValue = 0;
            }
            catch (IOException ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);

            }
            // 20240531 한인석 : GPIB Exception 처리 추가
            catch (InvalidOperationException ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
            return dResultValue;
        }

        /// <summary>
        /// Mode 설정
        /// </summary>
        /// <param name="nIdx"></param>
        /// <returns></returns>
        public bool OptModeSelect(int nIdx)
        {
            string strData;

            if (m_bConnected == false)
                return false;

            strData = String.Format(FunctionLCRMeterinfo.strCmd_LCRMeter_APERMED.Trim(), nIdx);

            try
            {
                LCRdevice.Write(strData);

                System.Threading.Thread.Sleep(10);
            }
            catch (IOException ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
                return false;
            }
            // 20240531 한인석 : GPIB Exception 처리 추가
            catch (InvalidOperationException ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
            return true;
        }

        /// <summary>
        /// Mode Short 설정
        /// </summary>
        /// <param name="nIdx"></param>
        /// <returns></returns>
        //public bool OptModeShortSelect(int nIdx)
        //{
        //    string strData;

        //    if (m_bConnected == false)
        //        return false;

        //    strData = String.Format(FunctionLCRMeterinfo.strCmd_LCRMeter_APERSHORT.Trim(), nIdx);

        //    try
        //    {
        //        LCRdevice.Write(strData);

        //        System.Threading.Thread.Sleep(10);
        //    }
        //    catch (IOException ex)
        //    {
        //        System.Diagnostics.Debug.Print(ex.Message);
        //        return false;
        //    }

        //    return true;
        //}

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public bool OptSelect(string _strCommand)
        {
            string strData;

            if (m_bConnected == false)
                return false;

            strData = String.Format(_strCommand.Trim());

            try
            {
                LCRdevice.Write(strData);

                System.Threading.Thread.Sleep(500);
            }
            catch (IOException ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
                return false;
            }
            // 20240531 한인석 : GPIB Exception 처리 추가
            catch (GpibException ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
            return true;
        }

        //public bool OptFreqSelect(double dIdx)   //주파수 설정
        public bool OptFreqSelect(decimal dIdx)   //주파수 설정
        {
            string strData;

            if (m_bConnected == false)
                return false;

            strData = String.Format(":FREQuency {0}", dIdx);

            try
            {
                LCRdevice.Write(strData);

                System.Threading.Thread.Sleep(10);
            }
            catch (IOException ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
                return false;
            }
            // 20240531 한인석 : GPIB Exception 처리 추가
            catch (InvalidOperationException ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
            return true;
        }

        //public bool OptLevelSelect(double nIdx)   //전압레벨 설정
        public bool OptLevelSelect(decimal nIdx)   //전압레벨 설정
        {
            string strData;

            if (m_bConnected == false)
                return false;

            strData = String.Format(":VOLTage {0}", nIdx);

            try
            {
                LCRdevice.Write(strData);

                System.Threading.Thread.Sleep(10);
            }
            catch (IOException ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
                return false;
            }
            // 20240531 한인석 : GPIB Exception 처리 추가
            catch (InvalidOperationException ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
            return true;
        }
    }
}
