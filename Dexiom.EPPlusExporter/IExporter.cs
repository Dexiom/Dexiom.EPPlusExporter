using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace Dexiom.EPPlusExporter
{
    public interface IExporter
    {
        ExcelPackage CreateExcelPackage();
        ExcelWorksheet AddWorksheetToExistingPackage(ExcelPackage package);
    }
}
