using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO.Ports;

namespace Coil_Diagnostor.Function
{
    public class FunctionDeviceInfo
    {

    }

    /// <summary>
    /// LCR Meter Info or Command 
    /// </summary>
    public class FunctionLCRMeterinfo
    {
        // Power Supply Command 설정
        public static string strCmd_LCRMeter_IDN = "*IDN?";                            // Reads out the product information (manufacturer, model number, serial number, and firmware version number) of the E4980A/AL. (Query Only)
        public static string strCmd_LCRMeter_FETCHOutp = "FETCH?";                     // Returns the Volteage or Current Value in Volts
        public static string strCmd_LCRMeter_Clean = "*CLS";                           // Clear status
        public static string strCmd_LCRMeter_Reset = "*RST";                           // Resets the instrument settings. The preset state is different from that when resetting is performed using the :SYSTem:PRESet. (No Query)
        public static string strCmd_LCRMeter_RDCLsSelect = ":FUNC:IMP:TYPE LSRD";      // DC 저항
        public static string strCmd_LCRMeter_RACLsSelect = ":FUNC:IMP:TYPE LSRS";      // AC 저항
        public static string strCmd_LCRMeter_LsSelect = ":FUNC:IMP:TYPE LSD";          // Inductance (L)
        public static string strCmd_LCRMeter_CpSelect = ":FUNC:IMP:TYPE CPD";           // 캐패시턴스(uP) C - Cp-D
        public static string strCmd_LCRMeter_CsSelect = ":FUNC:IMP:TYPE CSD";           // 캐패시턴스(uP) C - Cs-D
        public static string strCmd_LCRMeter_QSelect = ":FUNC:IMP:TYPE LSQ";           // Q-Factor
        public static string strCmd_LCRMeter_FUNCtion = ":FUNCtion:IMPedance ZTD";     // Selects the measurement function.
        public static string strCmd_LCRMeter_FREQuency = ":FREQuency {0}";             // 주파수 설정
        public static string strCmd_LCRMeter_VOLTage = ":VOLTage {0}";                 // 전압레벨 설정
        public static string strCmd_LCRMeter_Range = ":FUNC:IMP:RANGe {0}";             // Range 설정
        public static int intLCRMeer_RangeValue = 100;                                  // 설정 Range Value
        public static string strCmd_LCRMeter_APERMED = ":APER MED,{0}";                 // MODE 설정
    }
}
