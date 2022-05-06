using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;

namespace ExcelParser.Interface
{
    public interface IExcelParserBase<TExcelViewModel>
    {
        WorkbookPart WorkbookPart { get; set; }
        WorksheetPart WorksheetPart { get; set; }
        bool IsValidate { get; set; }
        SharedStringTable StringTable { get; set; }

        List<TExcelViewModel> GetExcelModel();
        TExcelViewModel ParseExcelRowModel(uint rowIndex);
        bool CheckColumnsName(uint rowIndex);
    }
}
