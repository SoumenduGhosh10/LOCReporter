using LOCCalculator.Classes;
using LOCCalculator.Interfaces;
using Logger.Classes;
using Logger.Interfaces;
using Reader.Classes;
using Reader.Interfaces;
using Reporter.Classes;
using Reporter.Interfaces;
using SourcesGetter.Classes;
using SourcesGetter.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sample
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                ILogger logger = new FileLogger();
                AReader reader = new XMLReader();

                ISourceGetter sources = new GetSourcesFromGitHub(logger, reader);
                ILOCCalculator locCalculator = new CLOCCalculator(logger, reader);
                IReportGenerator reporter = new ExcelReportGenerator(logger);

                Dictionary<string, string> inputFolderForLOC = sources.GetLatestSource();
                Console.WriteLine("Fetched all sources from GitHub");
                Dictionary<string, string> resultFiles = locCalculator.CalculateLOC(inputFolderForLOC);
                Console.WriteLine("Calculated LOC for all products");
                string resultFile = reporter.GenerateReport(resultFiles);
                Console.WriteLine($"LOC report generated at - {resultFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception occurred - {ex.Message}");
            }
        }
    }
}
