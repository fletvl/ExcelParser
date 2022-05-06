using System;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;
using System.IO;
using System.Collections.Generic;
using ExcelParser.Interface;

namespace ExcelParser.Parsers
{
    public abstract class ExcelParserBase<TExcelViewModel> : IExcelParserBase<TExcelViewModel>
    {
        public WorkbookPart WorkbookPart { get; set; }
        public WorksheetPart WorksheetPart { get; set; }
        public bool IsValidate { get; set; } = true;
        public SharedStringTable StringTable { get; set; }

        public ExcelParserBase(string filePath)
        {
            ValidateExcelFile(filePath);
        }

        private void ValidateExcelFile(string filePath)
        {
            if (!File.Exists(filePath))
            {
                IsValidate = false;
                return;
            }

            using (Stream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fileStream, false))
                {
                    if (doc.WorkbookPart == null)
                    {
                        IsValidate = false;
                        return;
                    }

                    WorkbookPart = doc.WorkbookPart;

                    // check first visible sheet
                    Sheet theSheet = WorkbookPart.Workbook.Descendants<Sheet>().First();
                    if (theSheet == null)
                    {
                        IsValidate = false;
                        return;
                    }

                    var stringTable = WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                    if (stringTable == null || stringTable?.SharedStringTable == null)
                    {
                        IsValidate = false;
                        return;
                    }

                    StringTable = stringTable.SharedStringTable;

                    var wsPart = (WorksheetPart)WorkbookPart.GetPartById(theSheet.Id);  

                    if (wsPart == null)
                    {
                        IsValidate = false;
                        return;
                    }

                    WorksheetPart = wsPart;

                    var rows = WorksheetPart.Worksheet.Descendants<Row>().ToList();
                    if (rows == null)
                    {
                        IsValidate = false;
                    }
                }
            }
        }

        public abstract List<TExcelViewModel> GetExcelModel();

        public abstract TExcelViewModel ParseExcelRowModel(uint rowIndex);

        public virtual bool CheckColumnsName(uint rowIndex)
        {
            return true;
        }

        protected string GetCellValue(string cellAddress)
        {
            try
            {
                Cell cell = WorksheetPart.Worksheet.Descendants<Cell>().Where(c => c.CellReference == cellAddress).FirstOrDefault();
                if (cell == null)
                    return "";

                string cellValue = cell.InnerText;
                if (string.IsNullOrEmpty(cellValue))
                    return cellAddress.StartsWith("C") ? "0" : "";

                //if cell's DataType is null but InnerText is exist then it's OLE DateTime or Number
                if (cell.DataType == null)
                {
                    if (cell.StyleIndex == null)
                        return cell.InnerText.Trim();

                    int styleIndex = (int)cell.StyleIndex.Value;
                    CellFormat cellFormat = (CellFormat)WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.ElementAt(styleIndex);
                    uint formatId = cellFormat.NumberFormatId.Value;

                    // NumberFormatId 14 is ShortDate and 165 is LongDate
                    if (formatId == 14 || formatId == 165)
                    {
                        if (double.TryParse(cellValue, out double oaDate))
                        {
                            cellValue = DateTime.FromOADate(oaDate).ToShortDateString();
                        }
                    }
                    else
                    {
                        cellValue = Math.Round(decimal.Parse(cellValue, System.Globalization.CultureInfo.InvariantCulture), 2, MidpointRounding.AwayFromZero).ToString().Replace(',', '.');
                    }
                }
                else if (cell.DataType.Value == CellValues.SharedString)
                {
                    cellValue = StringTable.ElementAt(int.Parse(cellValue)).InnerText.Trim();
                }

                return cellValue;
            }
            catch (Exception ex)
            {
                return $"Failed to get cell value: {ex.Message}";
            }
        }
    }
}
