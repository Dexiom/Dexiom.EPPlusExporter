using Microsoft.VisualStudio.TestTools.UnitTesting;
using Dexiom.EPPlusExporter;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Dexiom.EPPlusExporterTests.Extensions;
using Dexiom.EPPlusExporterTests.Helpers;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;

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
        public void AppendToExcelPackageTest()
        {
            var data = new[]
            {
                new { TextValue = "SomeText", DateValue = DateTime.Now, DoubleValue = 10.2, IntValue = 5}
            };

            var excelPackage = TestHelper.FakeAnExistingDocument();
            EnumerableExporter.Create(data)
                .CustomizeTable(range =>
                {
                    var newRange = range.Worksheet.Cells[range.End.Row, range.Start.Column, range.End.Row, range.End.Column];
                    newRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    newRange.Style.Fill.BackgroundColor.SetColor(Color.HotPink);
                })
                .AppendToExcelPackage(excelPackage);

            //TestHelper.OpenDocument(excelPackage);

            Assert.IsTrue(excelPackage.Workbook.Worksheets.Count == 2);
        }

        [TestMethod()]
        public void ExportEmptyEnumerableTest()
        {
            var data = Enumerable.Empty<Tuple<string, int, bool>>();

            var excelPackage = EnumerableExporter.Create(data).CreateExcelPackage();
            //TestHelper.OpenDocument(excelPackage);

            Assert.IsNotNull(excelPackage);
        }

        [TestMethod()]
        public void ExportNullTest()
        {
            IList<Tuple<string, int, bool>> data = null;

            // ReSharper disable once ExpressionIsAlwaysNull
            Assert.IsNull(EnumerableExporter.Create(data).CreateExcelPackage());
            // ReSharper disable once ExpressionIsAlwaysNull
            Assert.IsNull(ObjectExporter.Create(data).AppendToExcelPackage(TestHelper.FakeAnExistingDocument()));
        }
        
        [TestMethod()]
        public void ConfigureTest()
        {
            var data = new[]
            {
                new { TextValue = "SomeText", DateValue = DateTime.Now, DoubleValue = 10.2, IntValue = 5}
            };

            var dynamicProperties = new[]
            {
                DynamicProperty.Create(data, "DynamicColumn1", "Display Name 1", typeof(DateTime?), n => DateTime.Now.AddDays(n.IntValue - 4)),
                DynamicProperty.Create(data, "DynamicColumn2", "Display Name 2", typeof(double), n => n.DoubleValue - 0.2)
            };

            var excelPackage = EnumerableExporter.Create(data, dynamicProperties)
                .Configure(n => n.IntValue, configuration =>
                {
                    configuration.Header.Text = "";
                })
                .Configure(n => n.DateValue, configuration =>
                {
                    configuration.Header.Text = " ";
                    configuration.Header.SetStyle = style =>
                    {
                        style.Border.Bottom.Style = ExcelBorderStyle.Thick;
                    };
                    configuration.Content.NumberFormat = "dd-MM-yyyy";
                    configuration.Content.SetStyle = style =>
                    {
                        style.Border.Left.Style = ExcelBorderStyle.Dashed;
                        style.Border.Right.Style = ExcelBorderStyle.Dashed;
                    };
                })
                .Configure(new []{ "DynamicColumn1", "IntValue" }, n =>
                    {
                        n.Header.SetStyle = style =>
                        {
                            style.Font.Bold = true;
                            style.Font.Color.SetColor(Color.Black);
                        };
                    })
                .CustomizeTable(range =>
                {
                    var newRange = range.Worksheet.Cells[range.End.Row, range.Start.Column, range.End.Row, range.End.Column];
                    newRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    newRange.Style.Fill.BackgroundColor.SetColor(Color.HotPink);
                })
                .CreateExcelPackage();

            TestHelper.OpenDocument(excelPackage);


            var excelWorksheet = excelPackage.Workbook.Worksheets.First();
            
            //header
            Assert.IsTrue(excelWorksheet.Cells[1, 2].Style.Border.Bottom.Style == ExcelBorderStyle.Thick);
            Assert.IsTrue(excelWorksheet.Cells[1, 2].Text == " ");
            Assert.IsTrue(excelWorksheet.Cells[1, 4].Text == "Int Value");
            Assert.IsTrue(excelWorksheet.Cells[1, 1].Style.Fill.BackgroundColor.Rgb != "FFFF69B4");

            //data
            Assert.IsTrue(excelWorksheet.Cells[2, 2].Text == DateTime.Now.ToString("dd-MM-yyyy"));
            Assert.IsTrue(excelWorksheet.Cells[2, 2].Style.Border.Left.Style == ExcelBorderStyle.Dashed);
            Assert.IsTrue(excelWorksheet.Cells[2, 2].Style.Border.Right.Style == ExcelBorderStyle.Dashed);
            Assert.IsTrue(excelWorksheet.Cells[2, 1].Style.Fill.BackgroundColor.Rgb == "FFFF69B4");
        }

        #region Fluent Interface Tests

        [TestMethod()]
        public void WorksheetConfigurationTest()
        {
            const string newWorksheetName = "1 - NewSheet";
            const string newWorksheetExpectedTableName = "_1_-_NewSheet";
            var data = new[]
            {
                new { TextValue = "SomeText", DateValue = DateTime.Now, DoubleValue = 10.2, IntValue = 5}
            };

            var excelPackage = TestHelper.FakeAnExistingDocument();
            var eporter = EnumerableExporter.Create(data);

            //set properties
            eporter.WorksheetName = newWorksheetName;
            var sheetToCheck = eporter.AppendToExcelPackage(excelPackage);

            //TestHelper.OpenDocument(excelPackage);

            //check properties
            Assert.IsTrue(sheetToCheck.Name == newWorksheetName);
            Assert.IsNotNull(sheetToCheck.Tables[newWorksheetExpectedTableName]);
            
        }

        [TestMethod()]
        public void DefaultNumberFormatTest()
        {
            var data = new[]
            {
                new { TextValue = "SomeText", DateValue = DateTime.Now, DoubleValue = 10.2, IntValue = 5}
            };

            var dynamicProperties = new[]
            {
                DynamicProperty.Create(data, "DynamicColumn1", "Display Name 1", typeof(DateTime?), n => DateTime.Now.AddDays(n.IntValue - 4)),
                DynamicProperty.Create(data, "DynamicColumn2", "Display Name 2", typeof(double), n => n.DoubleValue - 0.2)
            };


            var excelPackage = EnumerableExporter.Create(data, dynamicProperties)
                .DefaultNumberFormat(typeof(DateTime), "| yyyy-MM-dd")
                .DefaultNumberFormat(typeof(DateTime?), "|| yyyy-MM-dd")
                .DefaultNumberFormat(typeof(double), "0.00 $")
                .DefaultNumberFormat(typeof(int), "00")
                .CreateExcelPackage();
            var excelWorksheet = excelPackage.Workbook.Worksheets.First();

            //TestHelper.OpenDocument(excelPackage);

            string numberDecimalSeparator = NumberFormatInfo.CurrentInfo.NumberDecimalSeparator;

            Assert.IsTrue(excelWorksheet.Cells[2, 2].Text == DateTime.Today.ToString("| yyyy-MM-dd")); //DateValue
            Assert.IsTrue(excelWorksheet.Cells[2, 3].Text == $"10{numberDecimalSeparator}20 $"); //DoubleValue
            Assert.IsTrue(excelWorksheet.Cells[2, 4].Text == "05"); //IntValue
            Assert.IsTrue(excelWorksheet.Cells[2, 5].Text == DateTime.Today.AddDays(1).ToString("|| yyyy-MM-dd")); //DynamicColumn1
            Assert.IsTrue(excelWorksheet.Cells[2, 6].Text == $"10{numberDecimalSeparator}00 $"); //DynamicColumn2

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

            string numberDecimalSeparator = NumberFormatInfo.CurrentInfo.NumberDecimalSeparator;

            Assert.IsTrue(excelWorksheet.Cells[2, 2].Text == DateTime.Today.ToString("DATE: yyyy-MM-dd")); //DateValue
            Assert.IsTrue(excelWorksheet.Cells[2, 3].Text == $"10{numberDecimalSeparator}20 $"); //DoubleValue
            Assert.IsTrue(excelWorksheet.Cells[2, 4].Text == "05"); //IntValue
        }

        [TestMethod()]
        public void DisplayTest()
        {
            var data = new[]
            {
                new { TextValue = "SomeText", DateValue = DateTime.Now, DoubleValue = 10.2, IntValue = 5}
            };

            var excelWorksheet = EnumerableExporter.Create(data)
                .Ignore(n => n.DateValue)
                .Display(n => new
                {
                    n.TextValue,
                    n.DoubleValue
                })
                .CreateExcelPackage()
                .Workbook.Worksheets.First();

            Assert.IsTrue(excelWorksheet.Cells[1, 1].Text == "Text Value");
            Assert.IsTrue(excelWorksheet.Cells[1, 2].Text == "Double Value");
            Assert.IsTrue(excelWorksheet.Cells[1, 3].Text == string.Empty);
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

        [TestMethod()]
        public void StyleForTest()
        {
            var data = new[]
            {
                new { TextValue = "SomeText", DateValue = DateTime.Now, DoubleValue = 10.2, IntValue = 5}
            };
            
            const string dateFormat = "yyyy-MM-dd HH:mm";
            var exporter = EnumerableExporter.Create(data)
                .StyleFor(n => n.DateValue, n => n.Numberformat.Format = dateFormat);

            var excelPackage = exporter.CreateExcelPackage();
            var excelWorksheet = excelPackage.Workbook.Worksheets.First();
            Assert.IsTrue(excelWorksheet.Cells[2, 2].Text == data.First().DateValue.ToString(dateFormat)); //DateValue
        }

        [TestMethod()]
        public void HeaderStyleForTest()
        {
            var data = new[]
            {
                new { TextValue = "SomeText", DateValue = DateTime.Now, DoubleValue = 10.2, IntValue = 5}
            };
    
            var exporter = EnumerableExporter.Create(data)
                .HeaderStyleFor(n => new { n.DateValue, n.DoubleValue, n.IntValue }, 
                    style => style.Border.Bottom.Style = ExcelBorderStyle.Thick);

            var excelPackage = exporter.CreateExcelPackage();
            //TestHelper.OpenDocument(excelPackage);

            var excelWorksheet = excelPackage.Workbook.Worksheets.First();
            Assert.IsTrue(excelWorksheet.Cells[1, 2].Style.Border.Bottom.Style == ExcelBorderStyle.Thick);
        }
        
        [TestMethod()]
        public void ConditionalStyleForTest()
        {
            var data = new[]
            {
                new { TextValue = "SomeText0", DateValue = DateTime.Now, DoubleValue = 0, IntValue = 5},
                new { TextValue = "SomeText1", DateValue = DateTime.Now, DoubleValue = 1, IntValue = 5},
                new { TextValue = "SomeText2", DateValue = DateTime.Now, DoubleValue = 2, IntValue = 5},
                new { TextValue = "SomeText3", DateValue = DateTime.Now, DoubleValue = 3, IntValue = 5}
            };

            var exporter = EnumerableExporter.Create(data)
                .ConditionalStyleFor(n => n.DoubleValue, (entry, style) =>
                {
                    if (entry.DoubleValue > 1)
                    {
                        style.Border.Bottom.Style = ExcelBorderStyle.Thick;
                    }
                });

            var excelPackage = exporter.CreateExcelPackage();
            //TestHelper.OpenDocument(excelPackage);

            var excelWorksheet = excelPackage.Workbook.Worksheets.First();
            Assert.IsTrue(excelWorksheet.Cells[3, 3].Style.Border.Bottom.Style == ExcelBorderStyle.None);
            Assert.IsTrue(excelWorksheet.Cells[4, 3].Style.Border.Bottom.Style == ExcelBorderStyle.Thick);
        }

        [TestMethod()]
        public void FormulaForTest()
        {
            var data = new[]
            {
                new { TextValue = "SomeText0", DateValue = DateTime.Now, DoubleValue = 0, IntValue = 5},
                new { TextValue = "SomeText1", DateValue = DateTime.Now, DoubleValue = 1, IntValue = 5},
                new { TextValue = "SomeText2", DateValue = DateTime.Now, DoubleValue = 2, IntValue = 5},
                new { TextValue = "SomeText3", DateValue = DateTime.Now, DoubleValue = 3, IntValue = 5}
            };

            var exporter = EnumerableExporter.Create(data)
                .FormulaFor(n => n.TextValue, (row, value) => $"=\"Text=\" & \"{value}-\" & \"{row.DoubleValue:0.00}\"");

            var excelPackage = exporter.CreateExcelPackage();
            //TestHelper.OpenDocument(excelPackage);

            var excelWorksheet = excelPackage.Workbook.Worksheets.First();
            Assert.IsTrue(excelWorksheet.Cells[3, 1].Formula == "=\"Text=\" & \"SomeText1-\" & \"1.00\"");
            Assert.IsTrue(excelWorksheet.Cells[4, 1].Formula == "=\"Text=\" & \"SomeText2-\" & \"2.00\"");
        }

        #endregion
    }
}