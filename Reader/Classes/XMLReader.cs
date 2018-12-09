using Reader.Interfaces;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace Reader.Classes
{
    /// <summary>
    /// Reads details from an xml file
    /// </summary>
    public class XMLReader : AReader
    {
        #region Constants

        protected const string MainRegion = "config";

        #endregion Constants

        #region Members

        protected string _inputFilePath = string.Empty;

        #endregion Members

        #region Constructor

        public XMLReader() : this(Path.Combine(Environment.CurrentDirectory, "LOCConfig.xml")) { }

        public XMLReader(string inputFilePath)
        {
            // Validate Inputs
            ValidateInputs(inputFilePath);

            _inputFilePath = inputFilePath;
        }

        #endregion Constructor

        #region Methods

        #region Implemented Methods

        public override string ReadSourceServerDetails()
        {
            var document = XDocument.Load(_inputFilePath);

            string sourceServerdetails = string.Empty;
            foreach (XElement item in document.Element(MainRegion).Element(SourceServerGroupName).Elements(SourceServerNodeName))
            {
                sourceServerdetails += item.Value + ";";
            }
            document = null;

            return sourceServerdetails;
        }

        public override Dictionary<string, string> ReadLOCCalculatorDetails()
        {
            var document = XDocument.Load(_inputFilePath);

            IEnumerable<XElement> locNode = document.Element(MainRegion).Elements(LocNodeName);

            Dictionary<string, string> locDeatils = new Dictionary<string, string>();

            locDeatils.Add(locCalculatorAttribute, locNode.Select(x => x.Attribute(locCalculatorAttribute)?.Value).FirstOrDefault());
            locDeatils.Add(excludeDirectoryAttribute, locNode.Select(x => x.Attribute(excludeDirectoryAttribute)?.Value).FirstOrDefault());
            locDeatils.Add(excludeExtensionAttribute, locNode.Select(x => x.Attribute(excludeExtensionAttribute)?.Value).FirstOrDefault());
            locDeatils.Add(excludeByRegExAttribute, locNode.Select(x => x.Attribute(excludeByRegExAttribute)?.Value).FirstOrDefault());
            locDeatils.Add(excludeFilesList, locNode.Select(x => x.Attribute(excludeFilesList)?.Value).FirstOrDefault());

            return locDeatils;
        }

        #endregion Implemented Methods

        #region Validate Inputs

        protected void ValidateInputs(string filePath)
        {
            if (string.IsNullOrEmpty(filePath) || string.IsNullOrWhiteSpace(filePath))
            {
                throw new Exception("Input file path cannot be empty/ null/ white-spaces");
            }

            if (!File.Exists(filePath))
            {
                throw new Exception("Input file does not exist at path - " + filePath);
            }

            if (filePath.Substring(filePath.LastIndexOf(".") + 1).CompareTo("xml") != 0)
            {
                throw new Exception("Input file type is not xml and hence cannot be read with an xml reader");
            }
        }

        #endregion Validate Inputs

        #endregion Methods
    }
}
