using Formatter.Interfaces;
using System;
using System.Collections.Generic;
using System.Text;

namespace Formatter.Classes
{
    /// <summary>
    /// Default formatter for messages
    /// </summary>
    public class DefaultFormatter : IMessageFormatter
    {
        #region Methods

        #region Implemented Methods

        /// <summary>
        /// Formats error message
        /// </summary>
        /// <param name="message">Input message</param>
        /// <returns>Formatted message</returns>
        public string FormatError(string message)
        {
            StringBuilder errorMessage = new StringBuilder();
            errorMessage.AppendLine("####################");
            errorMessage.AppendLine("## Error occurred - " + message);
            errorMessage.AppendLine("####################");

            return errorMessage.ToString();
        }

        /// <summary>
        /// Formats execution summary
        /// </summary>
        /// <param name="sectionHeader">execution summary header</param>
        /// <param name="executionDetails">execution details</param>
        /// <returns>Formatted message</returns>
        public string FormatExecutionSummary(string sectionHeader, Dictionary<string, string> executionDetails)
        {
            StringBuilder sectionMessage = new StringBuilder();
            sectionMessage.AppendLine($"---------- {sectionHeader} ----------");

            foreach (var item in executionDetails)
            {
                sectionMessage.AppendLine($"{item.Key,-30} - {item.Value}");
            }

            string splitter = string.Empty;
            for (int i = 0; i < sectionHeader.Length + 22; i++)
            {
                splitter += "-";
            }
            sectionMessage.AppendLine(splitter);

            return sectionMessage.ToString();
        }

        /// <summary>
        /// Formats standard message
        /// </summary>
        /// <param name="message">input message</param>
        /// <returns>Formatted message</returns>
        public string FormatMessage(string message)
        {
            return $"[{DateTime.Now.ToString()}] - {message}{Environment.NewLine}";
        }

        /// <summary>
        /// Formats a section beginning/ ending
        /// </summary>
        /// <param name="message">section name</param>
        /// <returns>Formatted message</returns>
        public string FormatSection(string message)
        {
            StringBuilder sectionMessage = new StringBuilder();
            string splitter = string.Empty;
            for (int i = 0; i < message.Length + 6; i++)
            {
                splitter += "*";
            }
            sectionMessage.AppendLine(splitter);
            sectionMessage.AppendLine($"** {message} **");
            sectionMessage.AppendLine(splitter);

            return sectionMessage.ToString();
        }

        #endregion Implemented Methods

        #endregion Methods
    }
}
