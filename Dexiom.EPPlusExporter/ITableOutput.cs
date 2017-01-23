using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml.Table;

namespace Dexiom.EPPlusExporter
{
    public interface ITableOutput
    {
        string WorksheetName { get; set; }
        TableStyles TableStyle { get; set; }
    }
}
