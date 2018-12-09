using System.Collections.Generic;

namespace Formatter.Interfaces
{
    public interface IMessageFormatter
    {
        string FormatSection(string message);
        string FormatMessage(string message);
        string FormatError(string message);
        string FormatExecutionSummary(string sectionHeader, Dictionary<string, string> executionDetails);
    }
}
