using System.Collections.Generic;

namespace Logger.Interfaces
{
    public interface ILogger
    {
        void LogSection(string message);
        void LogMessage(string message);
        void LogError(string error);
        void LogExecutionSummary(string sectionHeader, Dictionary<string, string> detailedMessage);
    }
}
