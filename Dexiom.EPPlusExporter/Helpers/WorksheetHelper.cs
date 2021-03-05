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
        private const string InvalidCaracterPlaceholder = "_";

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
            var invalidChar = new [] { ' ', '-' };
            var retVal = Replace(tableName, invalidChar, InvalidCaracterPlaceholder);
            
            if (!char.IsLetter(retVal[0]))
                retVal = $"{InvalidCaracterPlaceholder}{retVal}";

            return retVal;
        }
        
        private static string Replace(string text, char[] invalidChars, string newVal)
        {
            var temp = text.Split(invalidChars, StringSplitOptions.RemoveEmptyEntries);
            return string.Join(newVal, temp);
        }
    }
}
