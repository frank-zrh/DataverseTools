using NLog;
using System;
using System.Runtime.CompilerServices;

namespace ExcelToDataverseTool
{
    public static class LLog
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        public static void LogException(Exception ex, [CallerFilePath] string filePath = "", [CallerLineNumber] int lineNumber = 0)
        {
            string location = $"@{System.IO.Path.GetFileName(filePath)}:{lineNumber}";
            logger.Error(ex, $"Exception occurred {location}");
        }

        public static void LogError(string message, [CallerFilePath] string filePath = "", [CallerLineNumber] int lineNumber = 0)
        {
            string location = $"@{System.IO.Path.GetFileName(filePath)}:{lineNumber}";
            logger.Error($"Error: {message} {location}");
        }

        public static void LogWarning(string message, [CallerFilePath] string filePath = "", [CallerLineNumber] int lineNumber = 0)
        {
            string location = $"@{System.IO.Path.GetFileName(filePath)}:{lineNumber}";
            logger.Warn($"Warning: {message} {location}");
        }

        public static void LogInfo(string message, [CallerFilePath] string filePath = "", [CallerLineNumber] int lineNumber = 0)
        {
            string location = $"@{System.IO.Path.GetFileName(filePath)}:{lineNumber}";
            logger.Info($"Info: {message} {location}");
        }

        public static void LogTrace(string message, [CallerFilePath] string filePath = "", [CallerLineNumber] int lineNumber = 0)
        {
            string location = $"@{System.IO.Path.GetFileName(filePath)}:{lineNumber}";
            logger.Trace($"Trace: {message} {location}");
        }
    }
}