using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Bogus;
using Dexiom.EPPlusExporter;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace EPPlusExporterDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            ExportSimpleObject();
            //ExportEnumerable();
        }

        private static void ExportSimpleObject()
        {
            Console.WriteLine("Create fake data...");
            var fakePerson = new Employee();

            Console.WriteLine("Exporting Simple Object...");
            var exporter = ObjectExporter.Create(fakePerson)
                .Ignore(n => n.Phone)
                .DefaultNumberFormat(typeof(DateTime), "yyyy-MM-dd")
                .DefaultNumberFormat(typeof(double), "#,##0.00 $")
                .NumberFormatFor(n => n.ShoeSize, "0.0");
            
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
                .DefaultNumberFormat(typeof(DateTime), "yyyy-MM-dd")
                .NumberFormatFor(n => n.DateHired, "dd-MM-yyyy")
                .NumberFormatFor(n => n.DateHired, "dd-MM-yyyy")
                .NumberFormatFor(n => n.ShoeSize, "0.0")
                .NumberFormatFor(n => n.ChangeInPocket, "0.00 $")
                .NumberFormatFor(n => n.CarValue, "#,##0.00 $")
                .Ignore(n => n.Email)
                .TextFormatFor(n => n.Phone, "Cell: {0}")
                .StyleFor(n => n.DateContractEnd, style =>
                {
                    style.Fill.Gradient.Color1.SetColor(Color.Yellow);
                    style.Fill.Gradient.Color2.SetColor(Color.Green);
                });


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
