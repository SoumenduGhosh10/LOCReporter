using System.Collections.Generic;

namespace LOCCalculator.Interfaces
{
    public interface ILOCCalculator
    {
        Dictionary<string, string> CalculateLOC(Dictionary<string, string> inputFolders);
    }
}
