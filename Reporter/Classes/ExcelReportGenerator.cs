using Logger.Interfaces;
using Reporter.Interfaces;
using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Xml.Linq;
using System.Runtime.InteropServices;

namespace Reporter.Classes
{
    public class ExcelReportGenerator : IReportGenerator
    {
        #region Constants

        protected const string ExcelReportName = "LOCReport.xlsx";

        protected const string PivotTableReportName = "ProductSummary";
        protected const string ProductDetailsReportName = "ProductDetails";
        protected const string PivotTableName = "Summary";
        protected const string PivotTableRange = "B2";
        protected const string PiovotTableValuesFieldName = "LOC_Count";
        protected const string PiovotTableRowFieldName = "Products";
        protected const string PiovotTableColumnFieldName = "Languages";

        protected const int ExcelHeaderRow = 2;
        protected const string ProductColumn = "B";
        protected const string LanguageColumn = "C";
        protected const string FilesCountColumn = "D";
        protected const string LOCColumn = "E";
        protected const string FileNameColumn = "B";
        protected const string FileLOCColumn = "C";
        protected const string ExtensionGroupColumn = "D";
        
        protected const string ProductHeader = "Product";
        protected const string LanguageHeader = "Language";
        protected const string FileNameHeader = "File";
        protected const string FilesCountHeder = "File_Count";
        protected const string LOCHeader = "LOC";
        protected const string ExtensionGroupHeader = "Extension Group";

        protected const string ResultNode = "results";
        protected const string LanguagesNode = "languages";
        protected const string LanguageNode = "language";
        protected const string FilesNode = "files";
        protected const string FileNode = "file";
        protected const string NameAttribute = "name";
        protected const string FilesCountAttribute = "files_count";
        protected const string CodeAttrribute = "code";
        protected const string FileNameAttribute = "name";
        protected const string FileLOCAttribute = "code";
        protected const string ExtGroupAttribute = "language";

        protected const string PivotTableStyle = "TableStyleMedium20";
        protected const string LanguageReportTableStyle = "TableStyleMedium21";
        protected const string FileReportTableStyle = "TableStyleLight21";

        #endregion Constants

        #region Members

        protected ILogger _logger;
        protected string _stratingCellProductDetails = string.Empty;
        protected string _endingCellProductDetails = string.Empty;

        #endregion Members

        #region Constructors

        public ExcelReportGenerator(ILogger logger)
        {
            _logger = logger;
        }

        #endregion Constructors

        #region Methods

        #region Implemented Methods

        public string GenerateReport(Dictionary<string, string> clocResults)
        {
            try
            {
                // Validate input
                _logger.LogSection("Validating inputs for excel report generation");
                ValidateInput(clocResults);
                _logger.LogSection("Validated inputs for excel report generation successfully");

                // Create excel file
                _logger.LogSection("Reading inputs for excel report generation");
                string excelReportPath = CreateExcelReportFile();
                _logger.LogSection("Read inputs for excel report generation successfully");

                // Generate language based report
                _logger.LogSection("Generating language based report");
                GenerateLangaugeBasedReport(excelReportPath, clocResults);
                _logger.LogSection("Generated language based report successfully");

                // Generate file based report
                _logger.LogSection("Generating file based report");
                GenerateFileBasedReport(excelReportPath, clocResults);
                _logger.LogSection("Generated file based report successfully");

                // Generate pivot table
                _logger.LogSection("Generating pivot table");
                GeneratePivotTable(excelReportPath, clocResults);
                _logger.LogSection("Generated pivot successfully");

                return excelReportPath;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.Message);
                throw;
            }
            
        }

        #endregion Implemented Methods

        #region Validate Inputs

        protected void ValidateInput(Dictionary<string, string> clocResults)
        {
            foreach (var clocResultEntry in clocResults)
            {
                if (!File.Exists(clocResultEntry.Value))
                {
                    string message = $"For product {clocResultEntry.Key} no cloc result file exists at - {clocResultEntry.Value}";
                    _logger.LogError(message);
                    throw new FileNotFoundException(message);
                }
                _logger.LogMessage($"Validated that for product {clocResultEntry.Key} cloc result files exists at - {clocResultEntry.Value}");
            }
        }

        #endregion Validate Inputs

        #region Create Excel Report File

        protected string CreateExcelReportFile()
        {
            string reportFilePath = Path.Combine(Environment.CurrentDirectory, ExcelReportName);
            if (File.Exists(reportFilePath))
            {
                File.Delete(reportFilePath);
            }
            Application excelApp = null;
            Workbook workBook = null;
            object misValue = System.Reflection.Missing.Value;
            try
            {
                excelApp = new Application();

                workBook = excelApp.Workbooks.Add(misValue);
                workBook.SaveAs(reportFilePath, XlFileFormat.xlWorkbookDefault, misValue, misValue, misValue, misValue, XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                workBook.Close(true, misValue, misValue);
                excelApp.Quit();
                ReleaseObject(workBook);
                ReleaseObject(excelApp);
            }

            return reportFilePath;
        }

        #endregion Create Excel Report File

        #region Language Based Report

        protected void GenerateLangaugeBasedReport(string excelReportPath, Dictionary<string, string> clocResults)
        {
            Application excelApp = new Application();
            excelApp.DisplayAlerts = false;
            excelApp.Interactive = false;

            Workbook workbook = null;
            Worksheet productSummary = null;
            Worksheet productDetails = null;

            try
            {
                // Open excel for writing
                workbook = excelApp.Workbooks.Open(excelReportPath);
                _logger.LogMessage("Opened excel file for writing report");

                // Pick the second sheet
                if (excelApp.Application.Sheets.Count >= 1)
                {
                    for (int i = 2; i <= excelApp.Application.Sheets.Count; i++)
                    {
                        excelApp.Worksheets[i].Delete();
                    }
                    productSummary = (Worksheet)workbook.Worksheets[1];
                    productDetails = (Worksheet)workbook.Worksheets.Add(After: productSummary);
                }
                else
                {
                    productSummary = (Worksheet)workbook.Worksheets.Add();
                    productDetails = (Worksheet)workbook.Worksheets.Add(After: productSummary);
                }
                _logger.LogMessage("Picked up the second worksheet");

                // Rename sheet to Language based report
                productDetails.Name = ProductDetailsReportName;

                // Write column headers
                productDetails.Cells[ExcelHeaderRow, ProductColumn] = ProductHeader;
                productDetails.Cells[ExcelHeaderRow, LanguageColumn] = LanguageHeader;
                productDetails.Cells[ExcelHeaderRow, FilesCountColumn] = FilesCountHeder;
                productDetails.Cells[ExcelHeaderRow, LOCColumn] = LOCHeader;
                _logger.LogMessage("Header columns are written");

                // Apply header style
                productDetails.Range[productDetails.Cells[ExcelHeaderRow, ProductColumn], productDetails.Cells[ExcelHeaderRow, LOCColumn]].Font.Bold = true;
                _logger.LogMessage("Added styles to header columns");

                // Initialize counters for writing in excel
                _stratingCellProductDetails = $"{ProductColumn}{ExcelHeaderRow}";

                int rowCounter = ExcelHeaderRow + 1;
                int startRowForProduct = -1;
                int endRowForProduct = -1;

                XDocument report = null;
                IEnumerable<XElement> languagesDetailsNode = null;

                string language = string.Empty;
                string filesCount = string.Empty;
                string locCount = string.Empty;
                startRowForProduct = rowCounter;

                foreach (var entry in clocResults)
                {
                    report = XDocument.Load(entry.Value);
                    languagesDetailsNode = report.Element(ResultNode).Element(LanguagesNode).Elements(LanguageNode);

                    // Initializing variables for current product
                    language = string.Empty;
                    filesCount = string.Empty;
                    locCount = string.Empty;
                    startRowForProduct = rowCounter;

                    // Reading all language based details from cloc generated xml report
                    foreach (var node in languagesDetailsNode)
                    {
                        // reading all language based details
                        language = node.Attribute(NameAttribute).Value;
                        filesCount = node.Attribute(FilesCountAttribute).Value;
                        locCount = node.Attribute(CodeAttrribute).Value;

                        // Writing all details to excel report
                        productDetails.Cells[rowCounter, ProductColumn] = entry.Key;
                        productDetails.Cells[rowCounter, LanguageColumn] = language;
                        productDetails.Cells[rowCounter, FilesCountColumn] = filesCount;
                        productDetails.Cells[rowCounter, LOCColumn] = locCount;

                        rowCounter += 1;
                    }
                    _logger.LogMessage($"Written all language based details for {entry.Key} to excel report");

                    // Group language based details by product
                    endRowForProduct = rowCounter - 1;
                    (productDetails.Rows[$"{startRowForProduct}:{endRowForProduct - 1}"]).Group();
                    _logger.LogMessage($"Grouped all language based details for {entry.Key} to excel report");
                }

                _endingCellProductDetails = $"{LOCColumn}{endRowForProduct}";

                // Define table style
                Range SourceRange = (Range)productDetails.Range[productDetails.Cells[ExcelHeaderRow, ProductColumn], productDetails.Cells[endRowForProduct, LOCColumn]];
                SourceRange.Worksheet.ListObjects.Add(
                    XlListObjectSourceType.xlSrcRange,
                    SourceRange,
                    Type.Missing,
                    XlYesNoGuess.xlYes,
                    Type.Missing).Name = ProductDetailsReportName;
                SourceRange.Select();
                SourceRange.Worksheet.ListObjects[ProductDetailsReportName].TableStyle = LanguageReportTableStyle;
                _logger.LogMessage("Table style is added");

                // Auto fit all columns
                productDetails.Columns.EntireColumn.AutoFit();
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                object misValue = System.Reflection.Missing.Value;
                workbook.Save();
                workbook.Close();

                productSummary = null;
                productDetails = null;
                excelApp.Quit();

                ReleaseObject(workbook);
                ReleaseObject(excelApp);
            }
        }

        #endregion Language Based Report

        #region File Based Report

        private void GenerateFileBasedReport(string excelReportPath, Dictionary<string, string> clocResults)
        {
            Application excelApp = new Application();
            excelApp.DisplayAlerts = false;
            excelApp.Interactive = false;

            Workbook workbook = null;
            Worksheet lastWorkSheet = null;
            Worksheet currentWorkSheet = null;

            try
            {
                // Open excel for writing
                workbook = excelApp.Workbooks.Open(excelReportPath);
                _logger.LogMessage("Opened excel file for writing report");

                // If more than sheets are present delete all the unnecessary sheets                
                if (excelApp.Application.Sheets.Count > 2)
                {
                    for (int i = 3; i <= excelApp.Application.Sheets.Count; i++)
                    {
                        excelApp.Worksheets[i].Delete();
                    }
                }

                // Initialize counters for writing in excel
                int rowCounter;

                XDocument report = null;
                IEnumerable<XElement> filesDetailsNode = null;

                string fileName = string.Empty;
                string locCount = string.Empty;
                string extGroup = string.Empty;

                foreach (var entry in clocResults)
                {
                    // Add sheet after product details
                    lastWorkSheet = excelApp.Worksheets[excelApp.Application.Sheets.Count];
                    currentWorkSheet = (Worksheet)workbook.Worksheets.Add(After: lastWorkSheet);
                    _logger.LogMessage($"Added sheet for product - {entry.Key}");

                    // Rename sheet to the product name
                    currentWorkSheet.Name = entry.Key;

                    // Write column headers
                    currentWorkSheet.Cells[ExcelHeaderRow, FileNameColumn] = FileNameHeader;
                    currentWorkSheet.Cells[ExcelHeaderRow, FileLOCColumn] = LOCHeader;
                    currentWorkSheet.Cells[ExcelHeaderRow, ExtensionGroupColumn] = ExtensionGroupHeader;
                    _logger.LogMessage("Header columns are written");

                    // Apply header style
                    currentWorkSheet.Range[currentWorkSheet.Cells[ExcelHeaderRow, FileNameColumn], currentWorkSheet.Cells[ExcelHeaderRow, ExtensionGroupColumn]].Font.Bold = true;
                    _logger.LogMessage("Added styles to header columns");

                    report = XDocument.Load(entry.Value);
                    filesDetailsNode = report.Element(ResultNode).Element(FilesNode).Elements(FileNode);

                    // Initializing variables for current product
                    fileName = string.Empty;
                    locCount = string.Empty;
                    extGroup = string.Empty;
                    rowCounter = ExcelHeaderRow + 1;

                    // Reading all language based details from cloc generated xml report
                    foreach (var node in filesDetailsNode)
                    {
                        // reading all language based details
                        fileName = node.Attribute(FileNameAttribute).Value;
                        locCount = node.Attribute(FileLOCAttribute).Value;
                        extGroup = node.Attribute(ExtGroupAttribute).Value;

                        // Writing all details to excel report
                        currentWorkSheet.Cells[rowCounter, FileNameColumn] = fileName;
                        currentWorkSheet.Cells[rowCounter, FileLOCColumn] = locCount;
                        currentWorkSheet.Cells[rowCounter, ExtensionGroupColumn] = extGroup;

                        rowCounter += 1;
                    }
                    _logger.LogMessage($"Written all fileS based details for {entry.Key} to excel report");

                    // Define table style
                    Range SourceRange = (Range)currentWorkSheet.Range[currentWorkSheet.Cells[ExcelHeaderRow, FileNameColumn], currentWorkSheet.Cells[(rowCounter - 1), ExtensionGroupColumn]];
                    SourceRange.Worksheet.ListObjects.Add(
                        XlListObjectSourceType.xlSrcRange,
                        SourceRange,
                        Type.Missing,
                        XlYesNoGuess.xlYes,
                        Type.Missing).Name = entry.Key;
                    SourceRange.Select();
                    SourceRange.Worksheet.ListObjects[entry.Key].TableStyle = FileReportTableStyle;
                    _logger.LogMessage("Table style is added");

                    // Auto fit all columns
                    currentWorkSheet.Columns.EntireColumn.AutoFit();
                }
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                object misValue = System.Reflection.Missing.Value;
                workbook.Save();
                workbook.Close();

                lastWorkSheet = null;
                currentWorkSheet = null;

                excelApp.Quit();

                ReleaseObject(workbook);
                ReleaseObject(excelApp);
            }
        }

        #endregion File Based Report

        #region Pivot Table

        private void GeneratePivotTable(string excelReportPath, Dictionary<string, string> clocResults)
        {
            Application excelApp = new Application();
            excelApp.DisplayAlerts = false;
            excelApp.Interactive = false;

            Workbook workbook = null;
            Worksheet pivotTableWorkSheet = null;
            Worksheet productDetailsWorkSheet = null;

            try
            {
                // Open excel for writing
                workbook = excelApp.Workbooks.Open(excelReportPath);
                _logger.LogMessage("Opened excel file for writing report");

                // Pick the second sheet
                if (excelApp.Application.Sheets.Count < 1)
                {
                    pivotTableWorkSheet = (Worksheet)workbook.Worksheets.Add();
                }
                else
                {
                    pivotTableWorkSheet = (Worksheet)excelApp.Worksheets[1];
                }
                _logger.LogMessage("Picked up the first worksheet");

                // Rename sheet to product summary
                pivotTableWorkSheet.Name = PivotTableReportName;

                // Open 'Product Details' sheet
                productDetailsWorkSheet = workbook.Worksheets[2];

                Range productDetailsDataRange = null;
                if (!(IsValueValid(_stratingCellProductDetails) || IsValueValid(_endingCellProductDetails)))
                {
                    string message = "The value for starting cell and/or ending cell are/is null/empty/whitespace(s). To compute pivot table, valid value for starting cell and ending cell in product details report are required. Hence pivot table generation cannot continue";
                    _logger.LogError(message);
                    throw new Exception(message);
                }
                else
                {
                    productDetailsDataRange = ((Worksheet)productDetailsWorkSheet).Range[$"{_stratingCellProductDetails}:{_endingCellProductDetails}"];
                    _logger.LogMessage("Data range calculation in product details table completed successfully");
                }

                // Set the range for Pivot table
                Range tableRange = ((Worksheet)pivotTableWorkSheet).Range[PivotTableRange];
                _logger.LogMessage("Calculated pivot table range successfully from product details report");

                // Set Pivot cache
                PivotCache cache = workbook.PivotCaches().Add(XlPivotTableSourceType.xlDatabase, productDetailsDataRange);
                _logger.LogMessage("Created a pivot cache");

                // Create a pivot table
                PivotTable productSummaryTable = pivotTableWorkSheet.PivotTables().Add(PivotCache: cache, TableDestination: tableRange, TableName: PivotTableName);
                _logger.LogMessage("Created a pivot table");

                // Creating Pivot fields (pivot field = column name in 'Product Details' sheet)

                // Selecting the fields for values
                PivotField valueFieldLOC = productSummaryTable.PivotFields(LOCHeader);
                valueFieldLOC.Orientation = XlPivotFieldOrientation.xlDataField;
                valueFieldLOC.Name = PiovotTableValuesFieldName; // This field name cannot be same with column field and it can't be empty
                _logger.LogMessage($"Added {LOCHeader} as value field to the pivot table");

                // Selecting the column field
                PivotField columnFieldLanguages = productSummaryTable.PivotFields(LanguageHeader);
                columnFieldLanguages.Orientation = XlPivotFieldOrientation.xlColumnField;
                columnFieldLanguages.Name = PiovotTableColumnFieldName;
                _logger.LogMessage($"Added {LanguageHeader} as column field to the pivot table");

                // Selecting the row field
                PivotField rowFieldProduct = productSummaryTable.PivotFields(ProductHeader);
                rowFieldProduct.Orientation = XlPivotFieldOrientation.xlRowField;
                rowFieldProduct.Name = PiovotTableRowFieldName;
                _logger.LogMessage($"Added {ProductHeader} as row field to the pivot table");

                // Define a table-style for the Pivot table
                productSummaryTable.TableStyle2 = PivotTableStyle;
                _logger.LogMessage($"Added style to the pivot table");

                // Auto fit all columns
                pivotTableWorkSheet.Columns.EntireColumn.AutoFit();
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                workbook.Save();
                workbook.Close();

                pivotTableWorkSheet = null;
                productDetailsWorkSheet = null;
                excelApp.Quit();

                ReleaseObject(workbook);
                ReleaseObject(excelApp);
            }
        }

        #endregion Pivot Table

        protected void ReleaseObject(object pobjObjectToRelease)
        {
            try
            {
                Marshal.ReleaseComObject(pobjObjectToRelease);
                pobjObjectToRelease = null;
            }
            catch (Exception) { }
        }

        protected bool IsValueValid(string value)
        {
            return !(string.IsNullOrEmpty(value) || string.IsNullOrWhiteSpace(value));
        }

        #endregion Methods
    }
}
