using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Dexiom.EPPlusExporter.Extensions;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.Table;

namespace Dexiom.EPPlusExporter
{
    public static class WorksheetHelper
    {
        #region Internal
        internal static ExcelWorksheet AddWorksheet(ExcelPackage package, IEnumerable<object> data, string worksheetName, TableStyles tableStyle)
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
                worksheet.Cells[1, iCol].Value = GetPropertyDisplayName(property);
            }

            //Add rows
            for (var iRow = 2; iRow < myData.Count + 2; iRow++)
            {
                var item = myData.ElementAt(iRow - 2);
                iCol = 0;

                for (var i = 0; i < properties.Length; i++)
                {
                    var property = properties.ElementAt(i);

                    iCol++;
                    worksheet.Cells[iRow, iCol].Value = GetPropertyValue(property, item);
                }
            }

            //Format as table
            using (var range = worksheet.Cells[1, 1, myData.Count + 1, iCol])
            {
                FormatAsTable(range, tableStyle, $"tb_{worksheetName.Replace(" ", string.Empty)}");
            }

            return worksheet;
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

        #region Private
        private static string GetPropertyDisplayName(MemberInfo property)
        {
            var displayNameAttribute = MemberInfoExtensions.GetCustomAttribute<DisplayNameAttribute>(property, true);
            if (displayNameAttribute != null)
                return displayNameAttribute.DisplayName;

            var displayAttribute = MemberInfoExtensions.GetCustomAttribute<DisplayAttribute>(property, true);
            if (displayAttribute != null)
                return displayAttribute.Name;

            //well, let's just take that property name then...
            return property.Name;
        }

        private static object GetPropertyValue(PropertyInfo property, object item)
        {
            var value = property.GetValue(item);

            //check for customization via attribute
            var displayFormatAttribute = MemberInfoExtensions.GetCustomAttribute<DisplayFormatAttribute>(property, true);
            if (displayFormatAttribute != null)
            {
                //handle NullDisplayText 
                if (value == null && !string.IsNullOrWhiteSpace(displayFormatAttribute.NullDisplayText))
                    return displayFormatAttribute.NullDisplayText;

                //handle display format
                if (value != null)
                    return string.Format(displayFormatAttribute.DataFormatString, value);
            }

            //value is null, nothing else to do
            if (value == null)
                return string.Empty;

            //simple type
            if (property.PropertyType.IsValueType)
                return value;

            //enumerable
            var enumerable = (value as IEnumerable<object>);
            if (enumerable != null)
                return enumerable.Count(); //just show the count

            //well, let's throw the value...
            return value.ToString();
        }
        #endregion
    }
}
