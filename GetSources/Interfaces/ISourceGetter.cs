using System.Collections.Generic;

namespace SourcesGetter.Interfaces
{
    public interface ISourceGetter
    {
        Dictionary<string, string> GetLatestSource();
    }
}
