using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using Dexiom.EPPlusExporter.Extensions;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using System;

namespace Dexiom.EPPlusExporter.Helpers
{
    public static class WorksheetHelper
    {
        private const string ESCAPE_PREFIX = "__$";
        private const string SPACE_PLACEHOLDER = "__!";
        #region Internal
        internal static void FormatAsTable(ExcelRangeBase range, TableStyles tableStyle, string tableName, bool autoFitColumns = true)
        {
            string escapedTableName = FormatTableName(tableName);

            //format the table
            var table = range.Worksheet.Tables.Add(range, escapedTableName);
            table.TableStyle = tableStyle;

            if (autoFitColumns)
                range.AutoFitColumns();
        }
        #endregion

        private static string FormatTableName(string tableName)
        {
            string escapedTableName = tableName;

            char firstChar = tableName[0];

            if(Char.IsLetter(firstChar) == false)
            {
                escapedTableName =  $"{ESCAPE_PREFIX}{escapedTableName}";
            }

            escapedTableName = escapedTableName.Replace(" ", SPACE_PLACEHOLDER);

            return escapedTableName;
        }

    }
}
