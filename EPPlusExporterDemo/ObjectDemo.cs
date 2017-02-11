using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Dexiom.EPPlusExporter;
using OfficeOpenXml.Table;

namespace EPPlusExporterDemo
{
    public static class ObjectDemo
    {
        public static void Sample1()
        {
            var data = new {TextValue = "Text #1", DateValue = DateTime.Now, DoubleValue = 10.1, IntValue = 1};

            var excelPackage = ObjectExporter.Create(data).CreateExcelPackage();

            Program.SaveAndOpenDocument(excelPackage);
        }

        public static void Sample2()
        {
            var data = new {TextValue = "Text #1", DateValue = DateTime.Now, DoubleValue = 10.1, IntValue = 1};

            var exporter = ObjectExporter.Create(data)
                .TextFormatFor(n => n.TextValue, "Prefix: {0}")
                .NumberFormatFor(n => n.DateValue, "dd-MM-yyyy")
                .Configure(n => n.DateValue, configuration => configuration.Header.Text = "MyDate")
                .DefaultNumberFormat(typeof(double), "0.00 $")
                .Ignore(n => n.IntValue);

            exporter.WorksheetName = "MyData";
            exporter.TableStyle = TableStyles.Medium2;

            var excelPackage = exporter.CreateExcelPackage();

            Program.SaveAndOpenDocument(excelPackage);
        }
    }
}
