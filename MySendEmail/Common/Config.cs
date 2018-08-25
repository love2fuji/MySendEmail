﻿using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;

namespace MySendEmail.Common
{
    public class Config
    {
        private static Configuration m_Config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
        public static log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public static void SetValue(string key, string value)
        {
            if (!string.IsNullOrEmpty(key))
            {
                m_Config.AppSettings.Settings.Remove(key);
                if (!string.IsNullOrEmpty(value))
                {
                    m_Config.AppSettings.Settings.Add(key, value);
                }
                m_Config.Save();
            }
        }


        public static string GetValue(string key)
        {
            string result = "";
            if (m_Config.AppSettings.Settings[key] != null)
            {
                result = m_Config.AppSettings.Settings[key].Value;
            }
            return result;
        }
    }
}
