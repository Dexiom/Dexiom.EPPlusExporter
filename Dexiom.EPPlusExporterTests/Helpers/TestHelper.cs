using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Bogus;
using Dexiom.EPPlusExporter;
using Dexiom.EPPlusExporterTests.Model;
using OfficeOpenXml;

namespace Dexiom.EPPlusExporterTests.Helpers
{
    public static class TestHelper
    {
        public static void OpenDocumentIfRequired(ExcelPackage excelPackage, bool alwaysOpen = false)
        {
#if DEBUG
            alwaysOpen = true;
#endif
            if (alwaysOpen)
            {
                Console.WriteLine("Opening document");

                Directory.CreateDirectory("temp");
                var fileInfo = new FileInfo($"temp\\Test_{Guid.NewGuid():N}.xlsx");
                excelPackage.SaveAs(fileInfo);
                Process.Start(fileInfo.FullName);
            }
        }

        public static EnumerableExporter<Employee> CreateEmployeeExporter()
        {
            Console.WriteLine("CreateEmployeeExporter");

            return new EnumerableExporter<Employee>(GetEmployees());
        }

        public static IEnumerable<Employee> GetEmployees()
        {
            Console.WriteLine("GetEmployees");

            var faker = new Faker<Employee>().CustomInstantiator(n => new Employee(new Person()));
            return faker.Generate(1000);
        }

        public static ExcelPackage FakeAnExistingDocument()
        {
            Console.WriteLine("FakeAnExistingDocument");

            var retVal = new ExcelPackage();
            var worksheet = retVal.Workbook.Worksheets.Add("IAmANormalSheet");
            worksheet.Cells[1, 1].Value = "I am a normal sheet";
            worksheet.Cells[2, 1].Value = "with completly irrelevant data in it!";

            return retVal;
        }
    }
}
