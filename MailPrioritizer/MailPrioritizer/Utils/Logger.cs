using System;
using System.Globalization;
using System.IO;

namespace MailPrioritizer.Utils
{
    public enum LogLevel
    {
        Debug = 0,
        Info = 1,
        Warn = 2,
        Error = 3
    }

    /// <summary>파일 기반 로거. 스레드 안전. 7일 초과 로그 자동 삭제.</summary>
    public static class Logger
    {
        private static readonly object _lock = new object();
        private static readonly string LogDir =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                         "MailPrioritizer", "Logs");
        private const string FilePrefix = "MailPrioritizer-";
        private const string FileSuffix = ".log";
        private const int RetentionDays = 7;

        private static bool _initialized;

        public static LogLevel MinLevel { get; set; } = LogLevel.Info;

        /// <summary>로그 디렉토리 생성 + 오래된 로그 정리. ThisAddIn_Startup에서 1회 호출.</summary>
        public static void Initialize()
        {
            try
            {
                Directory.CreateDirectory(LogDir);
                _initialized = true;
                CleanupOldLogs();
                Info("=== MailPrioritizer started ===");
            }
            catch (Exception)
            {
                // 로거 초기화 실패 시에도 Add-in 정상 동작 보장
            }
        }

        public static void Debug(string message)
        {
            Write(LogLevel.Debug, message, null);
        }

        public static void Info(string message)
        {
            Write(LogLevel.Info, message, null);
        }

        public static void Warn(string message)
        {
            Write(LogLevel.Warn, message, null);
        }

        public static void Error(string message, Exception ex = null)
        {
            Write(LogLevel.Error, message, ex);
        }

        private static void Write(LogLevel level, string message, Exception ex)
        {
            if (!_initialized) return;
            if (level < MinLevel) return;

            string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
            string levelStr = level.ToString().ToUpperInvariant();
            string line = string.Format("[{0}] [{1}] {2}", timestamp, levelStr, message);

            if (ex != null)
            {
                line += Environment.NewLine
                     + string.Format("  Exception: {0}: {1}", ex.GetType().FullName, ex.Message)
                     + Environment.NewLine
                     + "  StackTrace: " + ex.StackTrace;

                if (ex.InnerException != null)
                {
                    line += Environment.NewLine
                         + string.Format("  InnerException: {0}: {1}",
                               ex.InnerException.GetType().FullName,
                               ex.InnerException.Message);
                }
            }

            lock (_lock)
            {
                try
                {
                    File.AppendAllText(GetLogFilePath(), line + Environment.NewLine);
                }
                catch (Exception)
                {
                    // 로그 기록 실패가 Add-in을 중단시키면 안 됨
                }
            }
        }

        private static string GetLogFilePath()
        {
            return Path.Combine(LogDir,
                FilePrefix + DateTime.Now.ToString("yyyy-MM-dd") + FileSuffix);
        }

        private static void CleanupOldLogs()
        {
            try
            {
                string[] files = Directory.GetFiles(LogDir, FilePrefix + "*" + FileSuffix);
                DateTime cutoff = DateTime.Now.Date.AddDays(-RetentionDays);

                foreach (string file in files)
                {
                    string name = Path.GetFileNameWithoutExtension(file);
                    string dateStr = name.Substring(FilePrefix.Length);
                    DateTime fileDate;
                    if (DateTime.TryParseExact(dateStr, "yyyy-MM-dd",
                            CultureInfo.InvariantCulture,
                            DateTimeStyles.None, out fileDate))
                    {
                        if (fileDate < cutoff)
                        {
                            File.Delete(file);
                        }
                    }
                }
            }
            catch (Exception)
            {
                // 정리 실패는 치명적이지 않음
            }
        }
    }
}
