using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.Linq;

namespace ExcelParser.Parsers
{
    public class ExcelTestModel
    {
        public string ColumnA { get; set; }
        public string ColumnB { get; set; }
        public string ColumnC { get; set; }
    }

    public class MultipleColumnsExcelParser : ExcelParserBase<ExcelTestModel>
    {
        public MultipleColumnsExcelParser(string filePath) : base(filePath) { }

        public override bool CheckColumnsName(uint rowIndex)
        {
            if (base.GetCellValue($"A{rowIndex}") == nameof(ExcelTestModel.ColumnA) &&
                base.GetCellValue($"B{rowIndex}") == nameof(ExcelTestModel.ColumnB) &&
                base.GetCellValue($"C{rowIndex}") == nameof(ExcelTestModel.ColumnC))
            {
                return true;
            }
            else
                return false;
        }

        public override List<ExcelTestModel> GetExcelModel()
        {
            List<ExcelTestModel> excelModel = new List<ExcelTestModel>();

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

        public override ExcelTestModel ParseExcelRowModel(uint rowIndex)
        {
            return new ExcelTestModel()
            {
                ColumnA = base.GetCellValue($"A{rowIndex}"),
                ColumnB = base.GetCellValue($"B{rowIndex}"),
                ColumnC = base.GetCellValue($"C{rowIndex}"),
            };
        }
    }
}
