using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Bogus;
using Dexiom.EPPlusExporter;
using OfficeOpenXml;

namespace EPPlusExporterDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            //ExportSimpleObject();
            ExportEnumerable();
        }

        private static void ExportSimpleObject()
        {
            Console.WriteLine("Create fake data...");
            var fakePerson = new Employee();

            Console.WriteLine("Exporting Simple Object...");
            var exporter = ObjectExporter.Create(fakePerson);
            
            var excelPackage = exporter.CreateExcelPackage();
            SaveAndOpenDocument(excelPackage);
        }

        private static void ExportEnumerable()
        {
            Console.WriteLine("Create fake data...");
            var faker = new Faker<Employee>().CustomInstantiator(n => new Employee());
            var data = faker.Generate(1000);

            Console.WriteLine("Exporting Enumerable...");
            var exporter = EnumerableExporter.Create(data)
                .Ignore(n => n.Phone)
                .TextFormatFor(n => n.UserName, "==> {0}")
                .StyleFor(n => n.DateOfBirth, style => style.Numberformat.Format = "YYYY-MMM-DD HH:mm:ss");
            
            var excelPackage = exporter.CreateExcelPackage();
            SaveAndOpenDocument(excelPackage);
        }

        private static void SaveAndOpenDocument(ExcelPackage excelPackage)
        {
            Console.WriteLine("Opening document");
            
            Directory.CreateDirectory("temp");
            var fileInfo = new FileInfo($"temp\\Test_{Guid.NewGuid():N}.xlsx");
            excelPackage.SaveAs(fileInfo);
            Process.Start(fileInfo.FullName);
        }
    }
}
