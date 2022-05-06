using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.Linq;

namespace ExcelParser.Parsers
{
    public class SimpleExcelParser : ExcelParserBase<string>
    {
        public SimpleExcelParser(string filePath) : base(filePath) { }

        public override List<string> GetExcelModel()
        {
            List<string> excelModel = new List<string>();

            var rows = WorksheetPart.Worksheet.Descendants<Row>().ToList();

            foreach (var row in rows)
            {
                var cells = row.ChildElements.Cast<Cell>().ToList();
                if (cells.All(x => x.CellValue == null))
                    continue;

                var rowModel = ParseExcelRowModel(row.RowIndex);
                excelModel.Add(rowModel);
            }

            return excelModel;
        }

        public override string ParseExcelRowModel(uint rowIndex)
        {
           return base.GetCellValue($"A{rowIndex}");
        }
    }
}
