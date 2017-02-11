using Microsoft.VisualStudio.TestTools.UnitTesting;
using Dexiom.EPPlusExporter;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Dexiom.EPPlusExporterTests.Helpers;
using OfficeOpenXml.Style;

namespace Dexiom.EPPlusExporter.Tests
{
    [TestClass()]
    public class ObjectExporterTests
    {
        [TestMethod()]
        public void CreateExcelPackageTest()
        {
            var data = new { TextValue = "SomeText", DateValue = DateTime.Now, DoubleValue = 10.2, IntValue = 5 };

            var excelPackage = ObjectExporter.Create(data).CreateExcelPackage();
            Assert.IsTrue(excelPackage.Workbook.Worksheets.Count == 1);
        }

        [TestMethod()]
        public void AddWorksheetToExistingPackageTest()
        {
            var data = new { TextValue = "SomeText", DateValue = DateTime.Now, DoubleValue = 10.2, IntValue = 5 };

            var excelPackage = TestHelper.FakeAnExistingDocument();
            ObjectExporter.Create(data).AddWorksheetToExistingPackage(excelPackage);
            Assert.IsTrue(excelPackage.Workbook.Worksheets.Count == 2);
            //TestHelper.OpenDocumentIfRequired(excelPackage);
        }

        [TestMethod()]
        public void ExportNullTest()
        {
            IList<Tuple<string, int, bool>> data = null;

            // ReSharper disable once ExpressionIsAlwaysNull
            Assert.IsNull(ObjectExporter.Create(data).CreateExcelPackage());
            // ReSharper disable once ExpressionIsAlwaysNull
            Assert.IsNull(ObjectExporter.Create(data).AddWorksheetToExistingPackage(TestHelper.FakeAnExistingDocument()));
        }

        [TestMethod()]
        public void WorksheetConfigurationTest()
        {
            const string newWorksheetName = "NewSheet";
            var data = new[]
            {
                new { TextValue = "SomeText", DateValue = DateTime.Now, DoubleValue = 10.2, IntValue = 5}
            };

            var excelPackage = TestHelper.FakeAnExistingDocument();
            var eporter = ObjectExporter.Create(data);

            //set properties
            eporter.WorksheetName = newWorksheetName;
            eporter.AddWorksheetToExistingPackage(excelPackage);

            //check properties
            var sheetToCheck = excelPackage.Workbook.Worksheets.Last();
            Assert.IsTrue(sheetToCheck.Name == newWorksheetName);
        }

        [TestMethod()]
        public void DefaultNumberFormatTest()
        {
            var data = new { TextValue = "SomeText", DateValue = DateTime.Now, DoubleValue = 10.2, IntValue = 5 };

            var excelWorksheet = ObjectExporter.Create(data)
                .DefaultNumberFormat(typeof(DateTime), "DATE: yyyy-MM-dd")
                .DefaultNumberFormat(typeof(double), "0.00 $")
                .DefaultNumberFormat(typeof(int), "00")
                .CreateExcelPackage()
                .Workbook.Worksheets.First();

            Assert.IsTrue(excelWorksheet.Cells[3, 2].Text == DateTime.Today.ToString("DATE: yyyy-MM-dd")); //DateValue
            Assert.IsTrue(excelWorksheet.Cells[4, 2].Text == "10.20 $"); //DoubleValue
            Assert.IsTrue(excelWorksheet.Cells[5, 2].Text == "05"); //IntValue
        }

        [TestMethod()]
        public void NumberFormatForTest()
        {
            var data = new { TextValue = "SomeText", DateValue = DateTime.Now, DoubleValue = 10.2, IntValue = 5 };

            var excelWorksheet = ObjectExporter.Create(data)
                .NumberFormatFor(n => n.DateValue, "DATE: yyyy-MM-dd")
                .NumberFormatFor(n => n.DoubleValue, "0.00 $")
                .NumberFormatFor(n => n.IntValue, "00")
                .CreateExcelPackage()
                .Workbook.Worksheets.First();

            Assert.IsTrue(excelWorksheet.Cells[3, 2].Text == DateTime.Today.ToString("DATE: yyyy-MM-dd")); //DateValue
            Assert.IsTrue(excelWorksheet.Cells[4, 2].Text == "10.20 $"); //DoubleValue
            Assert.IsTrue(excelWorksheet.Cells[5, 2].Text == "05"); //IntValue
        }

        [TestMethod()]
        public void IgnoreTest()
        {
            var data = new { TextValue = "SomeText", DateValue = DateTime.Now, DoubleValue = 10.2, IntValue = 5 };

            var excelWorksheet = ObjectExporter.Create(data)
                .Ignore(n => n.TextValue)
                .CreateExcelPackage()
                .Workbook.Worksheets.First();

            Assert.IsTrue(excelWorksheet.Cells[2, 1].Text == "Date Value");
        }


        [TestMethod()]
        public void TextFormatForTest()
        {
            var data = new { TextValue = "SomeText", DateValue = DateTime.Now, DoubleValue = 10.2, IntValue = 5 };

            const string textFormat = "Prefix: {0}";
            const string dateFormat = "{0:yyyy-MM-dd HH:mm}";
            var exporter = ObjectExporter.Create(data)
                .TextFormatFor(n => n.TextValue, textFormat)
                .TextFormatFor(n => n.DateValue, dateFormat);

            var excelPackage = exporter.CreateExcelPackage();
            var excelWorksheet = excelPackage.Workbook.Worksheets.First();

            Assert.IsTrue(excelWorksheet.Cells[2, 2].Text == string.Format(textFormat, data.TextValue)); //TextValue
            Assert.IsTrue(excelWorksheet.Cells[3, 2].Text == string.Format(dateFormat, data.DateValue)); //DateValue
            Assert.IsTrue(excelWorksheet.Cells[3, 2].Value.ToString() == string.Format(dateFormat, data.DateValue)); //DateValue
        }

        [TestMethod()]
        public void StyleForTest()
        {
            var data = new { TextValue = "SomeText", DateValue = DateTime.Now, DoubleValue = 10.2, IntValue = 5 };
            
            const string dateFormat = "yyyy-MM-dd HH:mm";
            var exporter = ObjectExporter.Create(data)
                .StyleFor(n => n.DateValue, n => n.Numberformat.Format = dateFormat);

            var excelPackage = exporter.CreateExcelPackage();
            var excelWorksheet = excelPackage.Workbook.Worksheets.First();
            
            Assert.IsTrue(excelWorksheet.Cells[3, 2].Text == data.DateValue.ToString(dateFormat)); //DateValue
        }

        [TestMethod()]
        public void ConditionalStyleForTest()
        {
            var data = new { TextValue = "SomeText", DateValue = DateTime.Now, DoubleValue = 10.2, IntValue = 5 };

            var exporter = ObjectExporter.Create(data)
                .ConditionalStyleFor(n => n.DateValue, (entry, style) =>
                {
                    if (entry.DoubleValue < 1)
                    {
                        style.Border.Bottom.Style = ExcelBorderStyle.Dashed;
                    }
                })
                .ConditionalStyleFor(n => n.DoubleValue, (entry, style) =>
                {
                    if (entry.DoubleValue > 1)
                    {
                        style.Border.Bottom.Style = ExcelBorderStyle.Thick;
                    }
                })
                .ConditionalStyleFor(n => n.IntValue, (entry, style) =>
                {
                    if (entry.DoubleValue > 1)
                    {
                        style.Border.Bottom.Style = ExcelBorderStyle.Dashed;
                    }
                });

            var excelPackage = exporter.CreateExcelPackage();
            //TestHelper.OpenDocument(excelPackage);

            var excelWorksheet = excelPackage.Workbook.Worksheets.First();
            Assert.IsTrue(excelWorksheet.Cells[3, 1].Style.Border.Bottom.Style == ExcelBorderStyle.None);
            Assert.IsTrue(excelWorksheet.Cells[4, 1].Style.Border.Bottom.Style == ExcelBorderStyle.Thick);
            Assert.IsTrue(excelWorksheet.Cells[5, 2].Style.Border.Bottom.Style == ExcelBorderStyle.Dashed);
        }
    }
}