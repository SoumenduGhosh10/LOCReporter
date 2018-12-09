using LibGit2Sharp;
using Logger.Interfaces;
using Reader.Interfaces;
using SourcesGetter.Interfaces;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace SourcesGetter.Classes
{
    /// <summary>
    /// Getting sources from open source GitHub repositories
    /// </summary>
    public class GetSourcesFromGitHub : ISourceGetter
    {
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
        public GetSourcesFromGitHub(ILogger logger, AReader reader)
        {
            _logger = logger;
            _reader = reader;
        }

        #endregion Constructor

        #region Methods

        #region Implemented Methods

        /// <summary>
        /// Get sources from GitHub server
        /// </summary>
        /// <returns>
        /// A dictionary containing Product name as key and local sources path as value
        /// </returns>
        public Dictionary<string, string> GetLatestSource()
        {
            try
            {
                // Read details from file
                _logger.LogSection("Reading GitHub details from config file");
                string sourceServerDetails = _reader.ReadSourceServerDetails();
                _logger.LogSection("Read GitHub details from config file successfully");

                // Validate the details
                _logger.LogSection("Validating GitHub inputs");
                ValidateInputs(sourceServerDetails);
                _logger.LogSection("Validated GitHub inputs successfully");

                // Read inputs
                _logger.LogSection("Reading GitHub details");
                List<GitHubDetails> serverDetails = ReadInputs(sourceServerDetails);
                _logger.LogSection("Read GitHub details successfully");

                // Perform get
                _logger.LogSection("Fetching latest sources from GitHub");
                Dictionary<string, string> localSourceDir = GetSource(serverDetails);
                _logger.LogSection("Fetched latest sources from GitHub");

                return localSourceDir;
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
        /// Validate GitHub details read from config files
        /// </summary>
        /// <param name="sourceServerDetails">GitHub details read from config file</param>
        protected void ValidateInputs(string sourceServerDetails)
        {
            string errorMessage = string.Empty;

            if (IsValueInvalid(sourceServerDetails))
            {
                errorMessage = "Source server details cannot be empty/ null/ white-spaces";
                _logger.LogError(errorMessage);
                throw new Exception(errorMessage);
            }
            _logger.LogMessage("Source server details is not empty/ null/ white-spaces");

            string[] connectionDetails = sourceServerDetails.Split(new string[] { ";" }, StringSplitOptions.RemoveEmptyEntries);

            if (connectionDetails.Length == 0)
            {
                errorMessage = "Source server details are invalid";
                _logger.LogError(errorMessage);
                throw new Exception(errorMessage);
            }
            _logger.LogMessage("Source server details split successfully");

            int counter = 0;
            foreach (var connection in connectionDetails)
            {
                if (connection.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries).Length < 3)
                {
                    errorMessage = $"GitHub connection details are not configured properly for connection - '{connection}'. A valid connection should have at least 3 entries";
                    _logger.LogError(errorMessage);
                    throw new Exception(errorMessage);
                }
                counter += 1;
                _logger.LogMessage($"{counter} source server details split successfully");
            }
        }

        /// <summary>
        /// Check whether a string is null/ empty/ whitespace(s)
        /// </summary>
        /// <param name="value">input value</param>
        /// <returns>True if value is null/ empty/ whitespace(s)</returns>
        protected bool IsValueInvalid(string value)
        {
            return string.IsNullOrEmpty(value) || string.IsNullOrWhiteSpace(value);
        }

        #endregion Validate Inputs

        #region Read Inputs

        /// <summary>
        /// Read input details sent form config file
        /// and create List<GitHubDetails> objects from that
        /// </summary>
        /// <param name="sourceServerDetails">GitHub details read from config file</param>
        /// <returns>List<GitHubDetails></returns>
        protected List<GitHubDetails> ReadInputs(string sourceServerDetails)
        {
            return
                sourceServerDetails
                .Split(new string[] { ";" }, StringSplitOptions.RemoveEmptyEntries)
                .Select(GitHubDetails.ParseFromString)
                .ToList();
        }

        #endregion Read Inputs

        #region Get Sources

        /// <summary>
        /// Get latest sources from GitHub
        /// </summary>
        /// <param name="serverDetails">GitHub server details</param>
        /// <returns>
        /// A dictionary containing Product name as key and local sources path as value
        /// </returns>
        protected Dictionary<string, string> GetSource(List<GitHubDetails> serverDetails)
        {
            Dictionary<string, string> localSourcePaths = new Dictionary<string, string>();

            foreach (GitHubDetails sourceDetail in serverDetails)
            {
                DateTime start = DateTime.Now;
                // Delete the repository
                DeleteDirectory(sourceDetail.LocalPath);
                _logger.LogMessage($"Cleaned up source directory - {sourceDetail.LocalPath}");

                // Recreate the empty directory
                Directory.CreateDirectory(sourceDetail.LocalPath);
                _logger.LogMessage($"Created empty directory to fetch sources at - {sourceDetail.LocalPath}");

                // Get latest sources
                CloneRepositoryForProduct(sourceDetail);
                _logger.LogMessage("Fetched latest source at - " + sourceDetail.LocalPath);

                // Delete folders not needed
                DeleteUnnecessaryDirectories(sourceDetail);
                DateTime end = DateTime.Now;

                TimeSpan span = end - start;
                _logger.LogMessage($"Fetching sources for project {sourceDetail.ProjectName} took {span.ToString()}");

                // Log LibGit2Sharp execution details
                LogLibGit2SharpExecutionDetails(sourceDetail.ProjectName, sourceDetail.LocalPath, span);

                localSourcePaths.Add(sourceDetail.ProjectName, sourceDetail.LocalPath);
            }
            return localSourcePaths;
        }

        /// <summary>
        /// Clone repository for GitHub details
        /// </summary>
        /// <param name="sourceDetail">GitHub details</param>
        protected void CloneRepositoryForProduct(GitHubDetails sourceDetail)
        {
            Repository.Clone(sourceDetail.GitHubURL, sourceDetail.LocalPath);
        }

        /// <summary>
        /// Delete unnecessary folders
        /// </summary>
        /// <param name="sourceDetail">GitHub details</param>
        protected void DeleteUnnecessaryDirectories(GitHubDetails sourceDetail)
        {
            if (null != sourceDetail.ProjectsToInclude && sourceDetail.ProjectsToInclude.Length > 0)
            {
                DirectoryInfo sourceDirectory = new DirectoryInfo(sourceDetail.LocalPath);
                foreach (var directory in sourceDirectory.GetDirectories())
                {
                    if (!sourceDetail.ProjectsToInclude.Contains(directory.Name))
                    {
                        DeleteDirectory(directory.FullName);
                    }
                }
                sourceDirectory = null;
                _logger.LogMessage("Cleaned up not required sources from - " + sourceDetail.LocalPath);
            }
        }

        /// <summary>
        /// Deletes a directory
        /// </summary>
        /// <param name="directoryPath">directory path to delete</param>
        protected void DeleteDirectory(string directoryPath)
        {
            if (Directory.Exists(directoryPath))
            {
                foreach (var subDirectory in Directory.GetDirectories(directoryPath))
                {
                    DeleteDirectory(subDirectory);
                }
                foreach (var file in Directory.GetFiles(directoryPath))
                {
                    var currentFileInfo = new FileInfo(file);
                    currentFileInfo.Attributes = FileAttributes.Normal;
                    currentFileInfo.Delete();
                }
                Directory.Delete(directoryPath);
            }
        }

        /// <summary>
        /// Log execution details 
        /// </summary>
        /// <param name="projectName">Project name</param>
        /// <param name="localPath">Local path</param>
        /// <param name="span">time span</param>
        protected void LogLibGit2SharpExecutionDetails(string projectName, string localPath, TimeSpan span)
        {
            Dictionary<string, string> details = new Dictionary<string, string>();

            details.Add("Project name", projectName);
            details.Add("Local path", localPath);
            details.Add("Time taken", span.ToString());

            _logger.LogExecutionSummary("LibGit2Sharp.dll execution details", details);
        }

        #endregion Get Sources

        #endregion Methods

        #region Inner Class

        /// <summary>
        /// GitHub details
        /// </summary>
        protected class GitHubDetails
        {
            public string ProjectName { get; set; }
            public string GitHubURL { get; set; }
            public string LocalPath { get; set; }
            public string[] ProjectsToInclude { get; set; }

            /// <summary>
            /// Read GitHub details from string
            /// </summary>
            /// <param name="arg">GitHub details as string</param>
            /// <returns>GitHubDetails object</returns>
            internal static GitHubDetails ParseFromString(string arg)
            {
                var connectionDetails = arg.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);

                GitHubDetails sourceDetails = new GitHubDetails
                {
                    ProjectName = connectionDetails[0],
                    GitHubURL = connectionDetails[1],
                    LocalPath = connectionDetails[2]  
                };

                if (connectionDetails.Length > 3)
                {
                    sourceDetails.ProjectsToInclude = connectionDetails[5].Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                }

                return sourceDetails;
            }
        }

        #endregion Inner Class
    }
}
