using System;
using System.IO;

namespace PDF_Handler
{
    /// <summary>
    /// 寫日誌類
    /// </summary>
    public static class LogUtil
    {
        /// <summary>
        /// 配置默認路徑
        /// </summary>
        private static string defaultPath = System.Configuration.ConfigurationManager.AppSettings["logPath"];

        #region Exception異常日誌

        /// <summary>
        /// 寫異常日誌，存放到默認路徑
        /// </summary>
        /// <param name="ex">異常類</param>
        public static void WriteError(Exception ex)
        {
            WriteError(ex, defaultPath);
        }

        /// <summary>
        /// 寫異常日誌，存放到指定路徑
        /// </summary>
        /// <param name="ex">異常類</param>
        /// <param name="path">日誌存放路徑</param>
        public static void WriteError(Exception ex, string path)
        {
            string errMsg = CreateErrorMeg(ex);
            WriteLog(errMsg, path, LogType.Error);
        }

        #endregion

        #region 普通日誌

        /// <summary>
        /// 寫普通日誌，存放到默認路徑，使用默認日誌類型
        /// </summary>
        /// <param name="msg">日誌內容</param>
        public static void WriteLog(string msg)
        {
            WriteLog(msg, LogType.Info);
        }

        /// <summary>
        /// 寫普通日誌，存放到默認路徑，使用指定日誌類型
        /// </summary>
        /// <param name="msg">日誌內容</param>
        /// <param name="logType">日誌類型</param>
        public static void WriteLog(string msg, LogType logType)
        {
            WriteLog(msg, defaultPath, logType);
        }

        /// <summary>
        /// 寫普通日誌，存放到指定路徑，使用默認日誌類型
        /// </summary>
        /// <param name="msg">日誌內容</param>
        /// <param name="path">日誌存放路徑</param>
        public static void WriteLog(string msg, string path)
        {
            WriteLog(msg, path, LogType.Info);
        }

        /// <summary>
        /// 寫普通日誌，存放到指定路徑，使用指定日誌類型
        /// </summary>
        /// <param name="msg">日誌內容</param>
        /// <param name="path">日誌存放路徑</param>
        /// <param name="logType">日誌類型</param>
        public static void WriteLog(string msg, string path, LogType logType)
        {
            string fileName = path.Trim('\\') + "\\" + CreateFileName(logType);
            string logContext = FormatMsg(msg, logType);
            WriteFile(logContext, fileName);
        }

        #endregion

        #region 其他輔助方法

        /// <summary>
        /// 寫日誌到文件
        /// </summary>
        /// <param name="logContext">日誌內容</param>
        /// <param name="fullName">文件名</param>
        private static void WriteFile(string logContext, string fullName)
        {
            FileStream fs = null;
            StreamWriter sw = null;

            int splitIndex = fullName.LastIndexOf('\\');
            if (splitIndex == -1) return;
            string path = fullName.Substring(0, splitIndex);

            if (!Directory.Exists(path)) Directory.CreateDirectory(path);

            try
            {
                if (!File.Exists(fullName)) fs = new FileStream(fullName, FileMode.CreateNew);
                else fs = new FileStream(fullName, FileMode.Append);

                sw = new StreamWriter(fs);
                sw.WriteLine(logContext);
            }
            finally
            {
                if (sw != null)
                {
                    sw.Close();
                    sw.Dispose();
                }
                if (fs != null)
                {
                    fs.Close();
                    fs.Dispose();
                }
            }
        }

        /// <summary>
        /// 格式化日誌，日誌是默認類型
        /// </summary>
        /// <param name="msg">日誌內容</param>
        /// <returns>格式化後的日誌</returns>
        private static string FormatMsg(string msg)
        {
            return FormatMsg(msg, LogType.Info);
        }

        /// <summary>
        /// 格式化日誌
        /// </summary>
        /// <param name="msg">日誌內容</param>
        /// <param name="logType">日誌類型</param>
        /// <returns>格式化後的日誌</returns>
        private static string FormatMsg(string msg, LogType logType)
        {
            string result;
            string header = string.Format("[{0}][{1} {2}] ", logType.ToString(), DateTime.Now.ToShortDateString(), DateTime.Now.ToShortTimeString());
            result = header + msg;
            return result;
        }

        /// <summary>
        /// 從異常類中獲取日誌內容
        /// </summary>
        /// <param name="ex">異常類</param>
        /// <returns>日誌內容</returns>
        private static string CreateErrorMeg(Exception ex)
        {
            string result = string.Empty;
            result += ex.Message + "\r\n";
            result += ex.StackTrace + "\r\n";
            return result;
        }

        /// <summary>
        /// 生成日誌文件名
        /// </summary>
        /// <param name="logType">日誌類型</param>
        /// <returns>日誌文件名</returns>
        private static string CreateFileName(LogType logType)
        {
            string result = DateTime.Now.ToString("yyyy-MM-dd");
            if (logType != LogType.Info)
                result = logType.ToString() + result + ".log";
            return result;
        }

        #endregion
    }

    /// <summary>
    /// 日誌類型
    /// </summary>
    public enum LogType
    {
        Error,
        Info,
        Option
    }
}
