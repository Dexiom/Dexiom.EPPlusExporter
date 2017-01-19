using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Bogus;
using Dexiom.EPPlusExporter;

namespace EPPlusExporterDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Create fake data...");
            var faker = new Faker<Employee>().CustomInstantiator(n => new Employee(new Person()));
            var data = faker.Generate(1000);

            Console.WriteLine("Export to Excel...");
            var fileInfo = new FileInfo("EPPlusExporterDemo.xlsx");
            var exporter = new EnumerableExporter<Employee>(data)
                .Ignore(n => n.Phone)
                //.Ignore(n => n.DateOfBirth)
                .DisplayFormatFor(n => n.UserName, "Toto: {0}")
                .DisplayFormatFor(n => n.DateOfBirth, "{0:u}");

            var excelPackage = exporter.CreateExcelPackage();
            excelPackage.SaveAs(fileInfo);

            Console.WriteLine("Open Document!");
            Process.Start(fileInfo.FullName);
        }
    }
}
