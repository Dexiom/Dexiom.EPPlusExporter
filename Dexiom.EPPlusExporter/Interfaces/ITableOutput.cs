using OfficeOpenXml.Table;

namespace Dexiom.EPPlusExporter.Interfaces
{
    public interface ITableOutput
    {
        string WorksheetName { get; set; }
        TableStyles TableStyle { get; set; }
    }
}
