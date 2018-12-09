using Formatter.Classes;
using Formatter.Interfaces;
using Logger.Interfaces;
using System;
using System.Collections.Generic;
using System.IO;

namespace Logger.Classes
{
    public class FileLogger : ILogger
    {
        #region Members

        protected string _logFilePath = string.Empty;
        protected IMessageFormatter _formatter = null;

        #endregion Members

        #region Constructor

        public FileLogger() : this(Path.Combine(Environment.CurrentDirectory, "LOCCalculator.txt")) { }

        public FileLogger(string logFilePath) : this(logFilePath, new DefaultFormatter()) { }
        
        public FileLogger(string logFilePath, IMessageFormatter formatter)
        {
            if (File.Exists(logFilePath))
            {
                File.Delete(logFilePath);
            }
            File.Create(logFilePath).Close();
            _logFilePath = logFilePath;
            _formatter = formatter;
        }

        #endregion Constructor

        #region Methods

        public void LogMessage(string message)
        {
            File.AppendAllText(_logFilePath, _formatter.FormatMessage(message));
        }

        public void LogSection(string message)
        {
            File.AppendAllText(_logFilePath, _formatter.FormatSection(message));
        }

        public void LogError(string error)
        {
            File.AppendAllText(_logFilePath, _formatter.FormatError(error));
        }

        public void LogExecutionSummary(string sectionHeader, Dictionary<string, string> executionDetails)
        {
            File.AppendAllText(_logFilePath, _formatter.FormatExecutionSummary(sectionHeader,executionDetails));
        }

        #endregion Methods
    }
}
