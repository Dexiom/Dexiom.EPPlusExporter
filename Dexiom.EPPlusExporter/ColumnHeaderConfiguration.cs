using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml.Style;

namespace Dexiom.EPPlusExporter
{
    public class ColumnHeaderConfiguration
    {
        public string Text { get; set; }

        public Action<ExcelStyle> SetStyle { get; set; } = style => { };
    }
}
