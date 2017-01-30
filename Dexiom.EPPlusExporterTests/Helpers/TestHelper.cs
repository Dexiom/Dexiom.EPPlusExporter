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

namespace Dexiom.EPPlusExporterTests.Helpers
{
    public static class TestHelper
    {
        public static void OpenDocument(ExcelPackage excelPackage)
        {
            Console.WriteLine("Opening document");

            Directory.CreateDirectory("temp");
            var fileInfo = new FileInfo($"temp\\Test_{Guid.NewGuid():N}.xlsx");
            excelPackage.SaveAs(fileInfo);
            Process.Start(fileInfo.FullName);
        }
        
        public static ExcelPackage FakeAnExistingDocument()
        {
            var retVal = new ExcelPackage();
            var worksheet = retVal.Workbook.Worksheets.Add("IAmANormalSheet");
            worksheet.Cells[1, 1].Value = "I am a normal sheet";
            worksheet.Cells[2, 1].Value = "with completly irrelevant data in it!";

            return retVal;
        }
    }
}
