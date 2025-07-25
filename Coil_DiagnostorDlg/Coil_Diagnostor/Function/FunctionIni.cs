using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

namespace Coil_Diagnostor.Function
{
    // 20240327 한인석
    // ini file read/write
    internal class IniFile   
    {
        string Path;
        string defaultSection = Assembly.GetExecutingAssembly().GetName().Name;
        //-------------------------------------------------------------------------------
        [DllImport("kernel32", CharSet = CharSet.Unicode)]
        static extern long WritePrivateProfileString(string section, string key, string value, string FilePath);

        [DllImport("kernel32", CharSet = CharSet.Unicode)]
        static extern int GetPrivateProfileString(string section, string key, string Default, StringBuilder RetVal, int Size, string FilePath);
        //-------------------------------------------------------------------------------
        public IniFile(string IniPath)
        {
            Path = new FileInfo(IniPath).FullName;
        }
        //-------------------------------------------------------------------------------
        public string Read(string key)
        {
            var RetVal = new StringBuilder(255);
            GetPrivateProfileString(defaultSection, key, "", RetVal, 255, Path);
            return RetVal.ToString();
        }
        public string Read(string section, string key)
        {
            var RetVal = new StringBuilder(255);
            GetPrivateProfileString(section, key, "", RetVal, 255, Path);
            return RetVal.ToString();
        }
        public string Read(string section, string key, string defaultVal)
        {
            var RetVal = new StringBuilder(255);
            GetPrivateProfileString(section, key, defaultVal, RetVal, 255, Path);
            return RetVal.ToString();
        }
        //-------------------------------------------------------------------------------
        public void Write(string key, string value)
        {
            WritePrivateProfileString(defaultSection, key, value, Path);
        }
        public void Write(string section, string key, string value)
        {
            WritePrivateProfileString(section, key, value, Path);
        }
        //-------------------------------------------------------------------------------
        public void Deletekey(string key)
        {
            Write(defaultSection, key, null);
        }
        public void Deletekey(string section, string key)
        {
            Write(section, key, null);
        }
        //-------------------------------------------------------------------------------
        public void Deletesection(string section)
        {
            Write(section, null, null);
        }
        //-------------------------------------------------------------------------------
        public bool keyExists(string key)
        {
            return Read(defaultSection, key).Length > 0;
        }
        public bool keyExists(string key, string section)
        {
            return Read(section, key).Length > 0;
        }
        //-------------------------------------------------------------------------------
    }
    // 20240328 한인석
    // global variables control
    // static class for use across all forms
    public static class Gini
    {
        private static IniFile m_iniFile = new IniFile("setting.ini");
        private static Dictionary<string, Dictionary<string, string>> m_data = new Dictionary<string, Dictionary<string, string>>();
        //-------------------------------------------------------------------------------
        public static string GetValue(string section, string key, string defaultVal = null)
        {
            string returnVal;

            if (!m_data.ContainsKey(section))
            {
                m_data[section] = new Dictionary<string, string>();
            }

            if (m_data[section].ContainsKey(key))
            {
                returnVal = m_data[section][key];
            }
            else
            {
                LoadFromFile(section, key, defaultVal);
                returnVal = m_data[section][key];
            }
            return returnVal;
        }
        public static bool SetValue(string section, string key, string value)
        { 
            if(SaveToFile(section, key, value))
            {
                m_data[section][key] = value;
                return true;
            }
            else
            {
                return false;
            }
        }
        //-------------------------------------------------------------------------------
        private static void LoadFromFile(string section, string key, string defaultVal = null)
        {
            m_data[section][key] = m_iniFile.Read(section, key, defaultVal);
        }
        private static bool SaveToFile(string section, string key, string value)
        {
            try
            {
                m_iniFile.Write(section, key, value);
            }
            catch
            {
                return false;
            }
            return true;
        }
        //-------------------------------------------------------------------------------
    }
}
