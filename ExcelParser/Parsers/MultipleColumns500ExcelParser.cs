using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.Linq;

namespace ExcelParser.Parsers
{
    public class Excel500RowsTestModel
    {
        public string ColumnA { get; set; }
        public string ColumnB { get; set; }
        public string ColumnC { get; set; }
        public string ColumnD { get; set; }
        public string ColumnE { get; set; }
    }

    public class MultipleColumns500ExcelParser : ExcelParserBase<Excel500RowsTestModel>
    {
        public MultipleColumns500ExcelParser(string filePath) : base(filePath) { }

        public override bool CheckColumnsName(uint rowIndex)
        {
            if (base.GetCellValue($"A{rowIndex}") == nameof(Excel500RowsTestModel.ColumnA) &&
                base.GetCellValue($"B{rowIndex}") == nameof(Excel500RowsTestModel.ColumnB) &&
                base.GetCellValue($"C{rowIndex}") == nameof(Excel500RowsTestModel.ColumnC) &&
                base.GetCellValue($"D{rowIndex}") == nameof(Excel500RowsTestModel.ColumnD) &&
                base.GetCellValue($"E{rowIndex}") == nameof(Excel500RowsTestModel.ColumnE))
            {
                return true;
            }
            else
                return false;
        }

        public override List<Excel500RowsTestModel> GetExcelModel()
        {
            List<Excel500RowsTestModel> excelModel = new List<Excel500RowsTestModel>();

            var rows = WorksheetPart.Worksheet.Descendants<Row>().ToList();

            foreach (var row in rows.Skip(1))
            {
                var cells = row.ChildElements.Cast<Cell>().ToList();
                if (cells.All(x => x.CellValue == null))
                    continue;

                var rowModel = ParseExcelRowModel(row.RowIndex);
                excelModel.Add(rowModel);
            }

            return excelModel;
        }

        public override Excel500RowsTestModel ParseExcelRowModel(uint rowIndex)
        {
            return new Excel500RowsTestModel()
            {
                ColumnA = base.GetCellValue($"A{rowIndex}"),
                ColumnB = base.GetCellValue($"B{rowIndex}"),
                ColumnC = base.GetCellValue($"C{rowIndex}"),
                ColumnD = base.GetCellValue($"D{rowIndex}"),
                ColumnE = base.GetCellValue($"E{rowIndex}")
            };
        }
    }
}
