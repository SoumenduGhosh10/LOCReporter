using LOCCalculator.Interfaces;
using Logger.Interfaces;
using Reader.Interfaces;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace LOCCalculator.Classes
{
    public class CLOCCalculator : ILOCCalculator
    {
        #region Constants

        protected const string clocExeName = "cloc-1.80.exe";

        protected const string ExcludeDirectoryOption = "--exclude-dir";
        protected const string ExcludeExtensionOption = "--exclude-ext";
        protected const string ExcludeFileOption = "--exclude-list-file";
        protected const string ExcludeFileByRegExOption = "--not-match-f";
        protected const string CustomLanguageDefinitionOption = "--force-lang-def";
        protected const string SkipUniquness = "--skip-uniqueness";
        protected const string EnableFileAndLanguageBasedReport = "--by-file-by-lang";
        protected const string OutputInXMLOption = "--xml";
        protected const string OutputFilePath = "--out";

        protected const string clocxmlReportMainNode = "results";
        protected const string clocxmlReportheaderNode = "header";
        protected const string clocxmlReportTotalFilesNode = "n_files";
        protected const string clocxmlReportTotalLinesNode = "n_lines";
        protected const string clocxmlReportFilesPerSecNode = "files_per_second";
        protected const string clocxmlReportLinesPerSerNode = "lines_per_second";

        #endregion

        #region Members

        protected ILogger _logger;
        protected AReader _reader;

        #endregion Members

        #region Constructor

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="logger">Logger type</param>
        /// <param name="reader">Reader type</param>
        public CLOCCalculator(ILogger logger, AReader reader)
        {
            _logger = logger;
            _reader = reader;
        }

        #endregion Constructor

        #region Methods

        #region Implemented Methods

        /// <summary>
        /// Calculate LOC for products
        /// </summary>
        /// <param name="inputFolders">input folder for calculating LOC</param>
        /// <returns>List of xml report paths created by cloc.exe</returns>
        public Dictionary<string, string> CalculateLOC(Dictionary<string, string> inputFolders)
        {
            try
            {
                // Read details from file
                _logger.LogSection("Reading cloc parameter details from config file");
                Dictionary<string, string> locDetails = _reader.ReadLOCCalculatorDetails();
                _logger.LogSection("Read cloc parameter details from config file successfully");

                // Validate the details
                _logger.LogSection("Validating inputs");
                ValidateInputs(locDetails, inputFolders);
                _logger.LogSection("Validated inputs successfully");

                // Read inputs
                _logger.LogSection("Reading inputs");
                CLOCParameters clocparameterDetails = ReadInputs(locDetails);
                _logger.LogSection("Read inputs successfully");

                // Perform loc calculation
                _logger.LogSection("Calculating LOC");
                Dictionary<string, string> resultFiles = CalculateLOC(inputFolders, clocparameterDetails);
                _logger.LogSection("Calculated LOC successfully");

                return resultFiles;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.Message);
                throw;
            }    
        }

        #endregion Implemented Methods

        #region Validate Inputs

        /// <summary>
        /// Validate inputs required for LOC calculation
        /// </summary>
        /// <param name="locDetails">cloc parameter details</param>
        /// <param name="inputFolders">input folders for LOC</param>
        protected void ValidateInputs(Dictionary<string, string> locDetails, Dictionary<string, string> inputFolders)
        {
            string message = string.Empty;

            // Validate input folders
            foreach (var item in inputFolders)
            {
                if (!Directory.Exists(item.Value))
                {
                    message = $"No folder exist at location - {item.Value}";
                    _logger.LogError(message);
                    throw new DirectoryNotFoundException(message);
                }
                _logger.LogMessage($"Validated input folder '{item.Value}' for product - {item.Key}");
            }

            // Validate loc calculator details
            string clocPath = locDetails[_reader.locCalculatorAttribute];

            if (IsValueInvalid(clocPath))
            {
                clocPath = Path.Combine(Environment.CurrentDirectory, clocExeName);
            }

            if (!File.Exists(clocPath))
            {
                message = $"cloc.exe doesn't exist at path {clocPath}";
                _logger.LogError(message);
                throw new FileNotFoundException(message);
            }
            _logger.LogMessage($"Validated that cloc.exe exists at {clocPath}");

            // Validate Exclude Files path
            string excludeFilesPath = locDetails[_reader.excludeFilesList];
            if (!IsValueInvalid(excludeFilesPath))
            {
                if (!File.Exists(excludeFilesPath))
                {
                    message = $"No file exists at path {excludeFilesPath}";
                    _logger.LogError(message);
                    throw new FileNotFoundException(message);
                }
                else
                {
                    _logger.LogMessage($"Validated that exclude file exists at {excludeFilesPath}");
                }
            }
        }

        #endregion Validate Inputs

        #region Read Inputs

        /// <summary>
        /// Converts input to clocParamters class
        /// </summary>
        /// <param name="locDetails"></param>
        /// <returns></returns>
        protected CLOCParameters ReadInputs(Dictionary<string, string> locDetails)
        {
            string clocPath = locDetails[_reader.locCalculatorAttribute];
            if (IsValueInvalid(clocPath))
            {
                clocPath = Path.Combine(Environment.CurrentDirectory, clocExeName);
            }

            CLOCParameters locParams = new CLOCParameters
            {
                LOCCalculatorPath = clocPath
            };
            _logger.LogMessage("Created cloc parameters object");

            string attributeValue = string.Empty;
            if (locDetails.ContainsKey(_reader.excludeDirectoryAttribute))
            {
                attributeValue = locDetails[_reader.excludeDirectoryAttribute];
                if (!IsValueInvalid(attributeValue))
                {
                    locParams.ExcludeDirectory = attributeValue;
                    _logger.LogMessage("Exclude Directory parameter was mentioned and is added to the parameters");
                }
            }

            if (locDetails.ContainsKey(_reader.excludeExtensionAttribute))
            {
                attributeValue = locDetails[_reader.excludeExtensionAttribute];
                if (!IsValueInvalid(attributeValue))
                {
                    locParams.ExcludeExtension = attributeValue;
                    _logger.LogMessage("Exclude Extension parameter was mentioned and is added to the parameters");
                }
            }

            if (locDetails.ContainsKey(_reader.excludeByRegExAttribute))
            {
                attributeValue = locDetails[_reader.excludeByRegExAttribute];
                if (!IsValueInvalid(attributeValue))
                {
                    locParams.ExcludeFileByRegEx = attributeValue;
                    _logger.LogMessage("Exclude Extension parameter was mentioned and is added to the parameters");
                }
            }

            if (locDetails.ContainsKey(_reader.excludeFilesList))
            {
                attributeValue = locDetails[_reader.excludeFilesList];
                if (!IsValueInvalid(attributeValue))
                {
                    locParams.ExcludeFiles = attributeValue;
                    _logger.LogMessage("Exclude Files List parameter was mentioned and is added to the parameters");
                }
            }

            return locParams;
        }

        #endregion Read Inputs

        #region Calculate LOC

        /// <summary>
        /// Calculate LOC for products
        /// </summary>
        /// <param name="inputFolders">input folders for LOC</param>
        /// <param name="clocparameterDetails">cloc parameter details</param>
        /// <returns></returns>
        protected Dictionary<string, string> CalculateLOC(Dictionary<string, string> inputFolders, CLOCParameters clocparameterDetails)
        {
            Dictionary<string, string> resultPaths = new Dictionary<string, string>();

            // Prepare cloc parameters for command line call
            _logger.LogMessage("Preparing common cloc.exe arguments");
            string clocArgument = PrepareCLOCArguments(clocparameterDetails);
            _logger.LogMessage("Prepared common cloc.exe arguments");

            string clocArgumentForCurrentFolder = string.Empty;
            string resultFile = string.Empty;
            DateTime start, end;
            TimeSpan span;
            foreach (var folderEntry in inputFolders)
            {
                // Add input folder and output file path to arguments
                resultFile = GetResultFilePath(folderEntry.Key, clocparameterDetails.LOCCalculatorPath);
                _logger.LogMessage($"Generated result file path for {folderEntry.Key}");

                _logger.LogMessage($"Preparing cloc.exe arguments specific for {folderEntry.Key}");
                clocArgumentForCurrentFolder = AppendInputFolderAndOutputFilePathToArgument(clocArgument, folderEntry.Value, resultFile);
                _logger.LogMessage($"Prepared cloc.exe arguments specific for {folderEntry.Key}");

                // Calculate loc
                start = DateTime.Now;
                _logger.LogMessage($"Calculating LOC for {folderEntry.Key}");
                CalculateLOC(clocparameterDetails.LOCCalculatorPath, clocArgumentForCurrentFolder);
                _logger.LogMessage($"LOC calculation completed for {folderEntry.Key}");
                end = DateTime.Now;
                span = end - start;

                // Log cloc execution details
                LogCLOCExecutionDetails(resultFile, span);

                // Add result file path to collection
                resultPaths.Add(folderEntry.Key, resultFile);
            }

            return resultPaths;
        }

        protected string PrepareCLOCArguments(CLOCParameters clocparameterDetails)
        {
            StringBuilder arguments = new StringBuilder();

            // Append directories to be excluded if any
            if (!IsValueInvalid(clocparameterDetails.ExcludeDirectory))
            {
                arguments.Append($" {ExcludeDirectoryOption}={clocparameterDetails.ExcludeDirectory}");
                _logger.LogMessage("Appended directories to be excluded");
            }

            // Append extensions to be excluded if any
            if (!IsValueInvalid(clocparameterDetails.ExcludeExtension))
            {
                arguments.Append($" {ExcludeExtensionOption}={clocparameterDetails.ExcludeExtension}");
                _logger.LogMessage("Appended extensions to be excluded");
            }

            // Append Exclude files by regular expression option
            if (!IsValueInvalid(clocparameterDetails.ExcludeFileByRegEx))
            {
                arguments.Append($" {ExcludeFileByRegExOption}={clocparameterDetails.ExcludeFileByRegEx}");
                _logger.LogMessage("Appended regular expression check to exclude file");
            }

            // Append files to be excluded if any
            if (!IsValueInvalid(clocparameterDetails.ExcludeFiles))
            {
                arguments.Append($" {ExcludeFileOption}=\"{clocparameterDetails.ExcludeFiles}\"");
                _logger.LogMessage("Appended files to exclude");
            }

            // To do custom language definition

            // Append Skip uniqueness option
            arguments.Append($" {SkipUniquness}");
            _logger.LogMessage("Appended skip uniqueness option");

            // Enable both language based and file based report
            arguments.Append($" {EnableFileAndLanguageBasedReport}");
            _logger.LogMessage("Appended enable file and language based report option");

            // Append the output file path
            // Append --xml as we want output in xml format 
            // which will help in Excel report generation
            arguments.Append($" {OutputInXMLOption}");
            _logger.LogMessage("Appended option to generate output in xml format");

            _logger.LogMessage("Appended all cloc.exe arguments");

            return arguments.ToString();
        }

        protected string GetResultFilePath(string ProductName, string clocPath)
        {
            FileInfo clocFileInfo = new FileInfo(clocPath);

            string parentDir = clocFileInfo.DirectoryName;

            return Path.Combine(parentDir, $"{ProductName}LOCReport.xml");
        }

        protected string AppendInputFolderAndOutputFilePathToArgument(string clocArgument, string inputFolder, string resultFile)
        {
            StringBuilder arguments = new StringBuilder();

            // Append input folder
            arguments.Append($" {inputFolder}");

            arguments.Append(clocArgument);

            // Append the output file path
            arguments.Append($" {OutputFilePath}=\"{resultFile}\"");

            return arguments.ToString();
        }

        protected void CalculateLOC(string lOCCalculatorPath, string clocArgumentForCurrentFolder)
        {
            // Initialize process with exe location and arguments
            Process lobjCLOCProcess = new Process();
            lobjCLOCProcess.StartInfo.FileName = lOCCalculatorPath;

            lobjCLOCProcess.StartInfo.Arguments = clocArgumentForCurrentFolder;
            lobjCLOCProcess.StartInfo.UseShellExecute = false;
            lobjCLOCProcess.StartInfo.RedirectStandardError = true;

            // Start process
            _logger.LogMessage("Triggering cloc.exe");
            lobjCLOCProcess.Start();
            lobjCLOCProcess.WaitForExit();

            if (lobjCLOCProcess.ExitCode != 0)
            {
                string errorMessage = $"cloc.exe completed unsuccessfully. Reason - {lobjCLOCProcess.StandardError.ReadToEnd()}";
                _logger.LogError(errorMessage);
                throw new Exception(errorMessage);
            }
            _logger.LogMessage("cloc.exe executed successfully");
        }

        protected void LogCLOCExecutionDetails(string resultFile, TimeSpan span)
        {
            Dictionary<string, string> details = new Dictionary<string, string>();
            var document = XDocument.Load(resultFile);

            IEnumerable<XElement> headerNode = document.Element(clocxmlReportMainNode).Elements(clocxmlReportheaderNode);

            details.Add("Total no of files read", headerNode.Elements(clocxmlReportTotalFilesNode).FirstOrDefault().Value);
            details.Add("Total no of lines read", headerNode.Elements(clocxmlReportTotalLinesNode).FirstOrDefault().Value);
            details.Add("No of files read per second", headerNode.Elements(clocxmlReportFilesPerSecNode).FirstOrDefault().Value);
            details.Add("No of lines read per second", headerNode.Elements(clocxmlReportLinesPerSerNode).FirstOrDefault().Value);
            details.Add("cloc.exe xml report path", resultFile);
            details.Add("Time taken", span.ToString());

            headerNode = null;

            _logger.LogExecutionSummary("cloc.exe execution details", details);
        }

        #endregion Calculate LOC

        protected bool IsValueInvalid(string value)
        {
            return string.IsNullOrEmpty(value) || string.IsNullOrWhiteSpace(value);
        }

        #endregion Methods

        #region Inner Class

        protected class CLOCParameters
        {
            public string LOCCalculatorPath { get; set; }
            public string ExcludeDirectory { get; set; }
            public string ExcludeExtension { get; set; }
            public string ExcludeFileByRegEx { get; set; }
            public string ExcludeFiles { get; set; }
        }

        #endregion Inner Class
    }
}
