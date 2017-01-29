using Microsoft.VisualStudio.TestTools.UnitTesting;
using Dexiom.EPPlusExporter;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Dexiom.EPPlusExporterTests.Extensions;
using Dexiom.EPPlusExporterTests.Helpers;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace Dexiom.EPPlusExporter.Tests
{

    [TestClass()]
    public class EnumerableExporterTests
    {
        [TestMethod()]
        public void CreateExcelPackageTest()
        {
            var data = new[]
            {
                new { TextValue = "SomeText", DateValue = DateTime.Now, DoubleValue = 10.2, IntValue = 5}
            };

            var excelPackage = EnumerableExporter.Create(data).CreateExcelPackage();
            Assert.IsTrue(excelPackage.Workbook.Worksheets.Count == 1);
        }
        
        [TestMethod()]
        public void AddWorksheetToExistingPackageTest()
        {
            var data = new[]
            {
                new { TextValue = "SomeText", DateValue = DateTime.Now, DoubleValue = 10.2, IntValue = 5}
            };

            var excelPackage = TestHelper.FakeAnExistingDocument();
            EnumerableExporter.Create(data).AddWorksheetToExistingPackage(excelPackage);
            Assert.IsTrue(excelPackage.Workbook.Worksheets.Count == 2);
            //TestHelper.OpenDocumentIfRequired(excelPackage);
        }

        [TestMethod()]
        public void DefaultNumberFormatTest()
        {
            var data = new[]
            {
                new { TextValue = "SomeText", DateValue = DateTime.Now, DoubleValue = 10.2, IntValue = 5}
            };

            var excelWorksheet = EnumerableExporter.Create(data)
                .DefaultNumberFormat(typeof(DateTime), "DATE: yyyy-MM-dd")
                .DefaultNumberFormat(typeof(double), "0.00 $")
                .DefaultNumberFormat(typeof(int), "00")
                .CreateExcelPackage()
                .Workbook.Worksheets.First();
            
            Assert.IsTrue(excelWorksheet.Cells[2, 2].Text == DateTime.Today.ToString("DATE: yyyy-MM-dd")); //DateValue
            Assert.IsTrue(excelWorksheet.Cells[2, 3].Text == "10.20 $"); //DoubleValue
            Assert.IsTrue(excelWorksheet.Cells[2, 4].Text == "05"); //IntValue
        }

        [TestMethod()]
        public void NumberFormatForTest()
        {
            var data = new[]
            {
                new { TextValue = "SomeText", DateValue = DateTime.Now, DoubleValue = 10.2, IntValue = 5}
            };

            var excelWorksheet = EnumerableExporter.Create(data)
                .NumberFormatFor(n => n.DateValue, "DATE: yyyy-MM-dd")
                .NumberFormatFor(n => n.DoubleValue, "0.00 $")
                .NumberFormatFor(n => n.IntValue, "00")
                .CreateExcelPackage()
                .Workbook.Worksheets.First();

            Assert.IsTrue(excelWorksheet.Cells[2, 2].Text == DateTime.Today.ToString("DATE: yyyy-MM-dd")); //DateValue
            Assert.IsTrue(excelWorksheet.Cells[2, 3].Text == "10.20 $"); //DoubleValue
            Assert.IsTrue(excelWorksheet.Cells[2, 4].Text == "05"); //IntValue
        }

        [TestMethod()]
        public void IgnoreTest()
        {
            var data = new[]
            {
                new { TextValue = "SomeText", DateValue = DateTime.Now, DoubleValue = 10.2, IntValue = 5}
            };

            var excelWorksheet = EnumerableExporter.Create(data)
                .Ignore(n => n.TextValue)
                .CreateExcelPackage()
                .Workbook.Worksheets.First();

            Assert.IsTrue(excelWorksheet.Cells[1, 1].Text == "Date Value");
        }


        [TestMethod()]
        public void TextFormatForTest()
        {
            var data = new[]
            {
                new { TextValue = "SomeText", DateValue = DateTime.Now, DoubleValue = 10.2, IntValue = 5}
            };

            const string textFormat = "Prefix: {0}";
            const string dateFormat = "{0:yyyy-MM-dd HH:mm}";
            var exporter = EnumerableExporter.Create(data)
                .TextFormatFor(n => n.TextValue, textFormat)
                .TextFormatFor(n => n.DateValue, dateFormat);

            var excelPackage = exporter.CreateExcelPackage();
            var excelWorksheet = excelPackage.Workbook.Worksheets.First();

            Assert.IsTrue(excelWorksheet.Cells[2, 1].Text == string.Format(textFormat, data.First().TextValue)); //TextValue
            Assert.IsTrue(excelWorksheet.Cells[2, 2].Text == string.Format(dateFormat, data.First().DateValue)); //DateValue
            Assert.IsTrue(excelWorksheet.Cells[2, 2].Value.ToString() == string.Format(dateFormat, data.First().DateValue)); //DateValue
        }
    }
}