using Microsoft.VisualStudio.TestTools.UnitTesting;
using Dexiom.EPPlusExporter;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Dexiom.EPPlusExporterTests.Helpers;
using OfficeOpenXml.Table;

namespace Dexiom.EPPlusExporter.Tests
{
    [TestClass()]
    public class TableExporterTests
    {
        [TestMethod()]
        public void TableCreationTest()
        {
            const string tableName = "MyTable";
            var data = new[]
            {
                new { TextValue = "SomeText", DateValue = DateTime.Now, DoubleValue = 10.2, IntValue = 5}
            };

            var exporter = EnumerableExporter.Create(data);
            exporter.WorksheetName = tableName;
            exporter.TableStyle = TableStyles.Dark10;

            var excelPackage = exporter.CreateExcelPackage();

            var sheetToCheck = excelPackage.Workbook.Worksheets.Last();
            Assert.IsTrue(sheetToCheck.Tables[tableName].TableStyle == exporter.TableStyle);
        }
    }
}