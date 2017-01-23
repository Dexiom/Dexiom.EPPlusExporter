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
using Dexiom.EPPlusExporterTests.Model;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace Dexiom.EPPlusExporter.Tests
{

    [TestClass()]
    public class EnumerableExporterTests
    {
        [TestMethod()]
        public void EnumerableExporterTest()
        {
            TestHelper.CreateEmployeeExporter();
        }

        [TestMethod()]
        public void CreateExcelPackageTest()
        {
            var exporter = TestHelper.CreateEmployeeExporter()
                .Ignore(n => n.Phone)
                .DisplayFormatFor(n => n.UserName, "** {0} **")
                .DisplayFormatFor(n => n.DateOfBirth, "{0:u}");

            var excelPackage = exporter.CreateExcelPackage();

            TestHelper.OpenDocumentIfRequired(excelPackage);
        }

        [TestMethod()]
        public void AddWorksheetToExistingPackage()
        {
            var exporter = TestHelper.CreateEmployeeExporter();

            var excelPackage = TestHelper.FakeAnExistingDocument();

            Console.WriteLine("exporter.AppendToExistingPackage");
            exporter.AddWorksheetToExistingPackage(excelPackage);
            TestHelper.OpenDocumentIfRequired(excelPackage);

            Assert.IsTrue(excelPackage.Workbook.Worksheets.Count == 2);
        }

        [TestMethod()]
        public void DisplayFormatForTest()
        {
            const string token = "*CUSTOM_FORMAT*";
            var exporter = TestHelper.CreateEmployeeExporter()
                .DisplayFormatFor(n => n.UserName, token + " {0}")
                .DisplayFormatFor(n => n.DateOfBirth, "{0:yyyy-MM-dd}");

            var excelPackage = exporter.CreateExcelPackage();
            var excelWorksheet = excelPackage.Workbook.Worksheets.First();

            //check text format
            {
                var myCell = excelWorksheet.Cells[2, 1].FlagInfo();
                myCell.Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                Console.WriteLine(myCell.Text);
                Assert.IsTrue(myCell.Text.StartsWith(token));
            }

            //check date format
            {
                var myCell = excelWorksheet.Cells[2, 6].FlagInfo();
                myCell.Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                var myDate = DateTime.ParseExact(myCell.Text, "yyyy-MM-dd", null);
                Console.WriteLine($"[{myCell.Text}] was converted to [{myDate}]");
            }

            TestHelper.OpenDocumentIfRequired(excelPackage);
        }

        [TestMethod()]
        public void IgnoreTest()
        {
            var exporter = TestHelper.CreateEmployeeExporter()
                .Ignore(n => n.UserName);

            var excelPackage = exporter.CreateExcelPackage();

            //check if the UserName column was removed
            var excelWorksheet = excelPackage.Workbook.Worksheets.First();
            var myCell = excelWorksheet.Cells[1, 1].FlagInfo();
            Assert.IsTrue(myCell.Text == "FirstName");

            TestHelper.OpenDocumentIfRequired(excelPackage);
        }

        [TestMethod()]
        public void UseAnonymousEnumerable()
        {
            var employees = TestHelper.GetEmployees().Select(n => new
            {
                Login = n.UserName,
                Mail = n.Email
            });

            var exporter = EnumerableExporter.Create(employees);

            var excelPackage = exporter.CreateExcelPackage();

            //check if the UserName column was removed
            var excelWorksheet = excelPackage.Workbook.Worksheets.First();
            var myCell = excelWorksheet.Cells[1, 2].FlagInfo();
            Assert.IsTrue(myCell.Text == "Mail");

            TestHelper.OpenDocumentIfRequired(excelPackage);
        }
    }
}