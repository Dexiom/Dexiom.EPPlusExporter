using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Dexiom.EPPlusExporter;
using OfficeOpenXml.Table;

namespace EPPlusExporterDemo
{
    public static class EnumerableDemo
    {
        public static void Sample1()
        {
            var data = new[]
            {
                new { TextValue = "Text #1", DateValue = DateTime.Now, DoubleValue = 10.1, IntValue = 1},
                new { TextValue = "Text #2", DateValue = DateTime.Now, DoubleValue = 10.2, IntValue = 2},
                new { TextValue = "Text #3", DateValue = DateTime.Now, DoubleValue = 10.3, IntValue = 3},
                new { TextValue = "Text #4", DateValue = DateTime.Now, DoubleValue = 10.4, IntValue = 4}
            };

            var excelPackage = EnumerableExporter.Create(data).CreateExcelPackage();

            Program.SaveAndOpenDocument(excelPackage);
        }

        public static void Sample2()
        {
            var data = new[]
            {
                new { TextValue = "Text #1", DateValue = DateTime.Now, DoubleValue = 10.1, IntValue = 1},
                new { TextValue = "Text #2", DateValue = DateTime.Now, DoubleValue = 10.2, IntValue = 2},
                new { TextValue = "Text #3", DateValue = DateTime.Now, DoubleValue = 10.3, IntValue = 3},
                new { TextValue = "Text #4", DateValue = DateTime.Now, DoubleValue = 10.4, IntValue = 4}
            };

            var exporter = EnumerableExporter.Create(data)
                .TextFormatFor(n => n.TextValue, "Prefix: {0}")
                .NumberFormatFor(n => n.DateValue, "yyyy-MM-dd")
                .DefaultNumberFormat(typeof(double), "0.00 $")
                .Ignore(n => n.IntValue);

            exporter.WorksheetName = "MyData";
            exporter.TableStyle = TableStyles.Medium2;

            var excelPackage = exporter.CreateExcelPackage();

            Program.SaveAndOpenDocument(excelPackage);
        }

        public static void Sample3()
        {
            var data = new[]
            {
                new { TextValue = "Text #1", DateValue = DateTime.Now, DoubleValue = 10.1, IntValue = 1},
                new { TextValue = "Text #2", DateValue = DateTime.Now, DoubleValue = 10.2, IntValue = 2},
                new { TextValue = "Text #3", DateValue = DateTime.Now, DoubleValue = 10.3, IntValue = 3},
                new { TextValue = "Text #4", DateValue = DateTime.Now, DoubleValue = 10.4, IntValue = 4}
            };

            var exporter = EnumerableExporter.Create(data)
                .NumberFormatFor(n => n.DateValue, "yyyy-mmm-dd")
                .NumberFormatFor(n => new
                {
                    n.DoubleValue,
                    n.IntValue
                }, "0.00");
            
            var excelPackage = exporter.CreateExcelPackage();

            Program.SaveAndOpenDocument(excelPackage);
        }
    }
}
