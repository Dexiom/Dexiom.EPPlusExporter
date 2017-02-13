using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Reflection;
using Dexiom.EPPlusExporter.Extensions;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace Dexiom.EPPlusExporter.Helpers
{
    public static class WorksheetHelper
    {
        #region Internal
        internal static void FormatAsTable(ExcelRangeBase range, TableStyles tableStyle, string tableName, bool autoFitColumns = true)
        {
            //format the table
            var table = range.Worksheet.Tables.Add(range, tableName);
            table.TableStyle = tableStyle;

            if (autoFitColumns)
                range.AutoFitColumns();
        }
        #endregion
    }
}
