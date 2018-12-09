using System.Collections.Generic;

namespace Reader.Interfaces
{
    public abstract class AReader
    {
        public readonly string SourceServerGroupName = "SourceServers";
        public readonly string SourceServerNodeName = "Server";
        public readonly string LocNodeName = "LOC";
        public readonly string locCalculatorAttribute = "LOCCalculatorPath";
        public readonly string excludeDirectoryAttribute = "ExcludedDirectory";
        public readonly string excludeExtensionAttribute = "ExcludedExtension";
        public readonly string excludeByRegExAttribute = "ExcludeFileByRegularExpression";
        public readonly string excludeFilesList = "ExcludedFilesList";

        public abstract string ReadSourceServerDetails();
        public abstract Dictionary<string, string> ReadLOCCalculatorDetails();
    }
}
