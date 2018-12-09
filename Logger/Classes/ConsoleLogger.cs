using Formatter.Classes;
using Formatter.Interfaces;
using Logger.Interfaces;
using System;
using System.Collections.Generic;

namespace Logger.Classes
{
    public class ConsoleLogger : ILogger
    {
        #region Members

        protected IMessageFormatter _formatter = null;

        #endregion Members

        #region Constructor

        public ConsoleLogger() : this(new DefaultFormatter()) { }

        public ConsoleLogger(IMessageFormatter formatter)
        {
            _formatter = formatter;
        }

        #endregion Constructor

        #region Methods

        public void LogMessage(string message)
        {
            Console.WriteLine(_formatter.FormatMessage(message));
        }

        public void LogSection(string message)
        {
            Console.WriteLine(_formatter.FormatSection(message));
        }

        public void LogError(string error)
        {
            Console.WriteLine(_formatter.FormatError(error));
        }

        public void LogExecutionSummary(string sectionHeader, Dictionary<string, string> executionDetails)
        {
            Console.WriteLine(_formatter.FormatExecutionSummary(sectionHeader, executionDetails));
        }

        #endregion Methods
    }
}
