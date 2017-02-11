using OfficeOpenXml;

namespace Dexiom.EPPlusExporter.Interfaces
{
    public interface IExporter
    {
        ExcelPackage CreateExcelPackage();
        ExcelWorksheet AddWorksheetToExistingPackage(ExcelPackage package);
    }
}
