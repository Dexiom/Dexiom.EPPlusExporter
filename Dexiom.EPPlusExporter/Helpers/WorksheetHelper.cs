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
        internal static ExcelRange AddDataWorksheet(ExcelPackage package, IEnumerable<object> data, string worksheetName)
        {
            //Avoid multiple enumeration
            var myData = data as IList<object> ?? data.ToList();

            if (data == null || !myData.Any())
                return null;
            
            var properties = myData.First().GetType().GetProperties();
            var worksheet = package.Workbook.Worksheets.Add(worksheetName);

            //Create table header
            var iCol = 0;
            foreach (var property in properties)
            {
                iCol++;
                worksheet.Cells[1, iCol].Value = ReflectionHelper.GetPropertyDisplayName(property);
            }

            //Add rows
            for (var iRow = 2; iRow < myData.Count + 2; iRow++)
            {
                var item = myData.ElementAt(iRow - 2);
                iCol = 0;

                foreach (var property in properties)
                {
                    iCol++;
                    worksheet.Cells[iRow, iCol].Value = ReflectionHelper.GetPropertyDisplayValue(property, item);
                }
            }
            
            return worksheet.Cells[1, 1, myData.Count + 1, iCol];
        }
         
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
