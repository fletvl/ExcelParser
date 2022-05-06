using ExcelParser.Parsers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;

namespace Test
{
    [TestClass]
    public class ExcelParserTest
    {
        [TestMethod]
        public void ParseSimpleExcelTest()
        {
            string excelPath = "TestExcel.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"ExcelTestFiles\", excelPath);

            SimpleExcelParser parser = new SimpleExcelParser(path);

            var model = parser.GetExcelModel();
        }

        [TestMethod]
        public void ParseMultipleColumnsExcelTest()
        {
            string excelPath = "TestExcelColumn.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"ExcelTestFiles\", excelPath);

            MultipleColumnsExcelParser parser = new MultipleColumnsExcelParser(path);

            Assert.IsTrue(parser.CheckColumnsName(1));

            var model = parser.GetExcelModel();
        }

        [TestMethod]
        public void ParseMultipleColumns500RowsExcelTest()
        {
            string excelPath = "TestExcelColumns500.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"ExcelTestFiles\", excelPath);

            MultipleColumns500ExcelParser parser = new MultipleColumns500ExcelParser(path);

            Assert.IsTrue(parser.CheckColumnsName(1));

            var model = parser.GetExcelModel();
        }
    }
}
