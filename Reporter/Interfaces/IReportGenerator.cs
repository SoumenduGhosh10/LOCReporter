using System.Collections.Generic;

namespace Reporter.Interfaces
{
    public interface IReportGenerator
    {
        string GenerateReport(Dictionary<string, string> clocResults);
    }
}
