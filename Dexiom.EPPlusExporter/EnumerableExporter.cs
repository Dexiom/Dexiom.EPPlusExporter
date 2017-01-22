using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Dexiom.EPPlusExporter.Helpers;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace Dexiom.EPPlusExporter
{
    public class EnumerableExporter 
        : EnumerableExporter<object>
    {
        #region Constructors
        public EnumerableExporter(IEnumerable<object> data) 
            : base(data)
        {
        }
        #endregion
    }

    public class EnumerableExporter<T> : IExportFormat<T>
        where T : class
    {
        #region Constructors
        public EnumerableExporter(IEnumerable<T> data)
        {
            Data = data;
        }
        #endregion

        #region Public Functions
        public ExcelPackage CreateExcelPackage()
        {
            var retVal = new ExcelPackage();
            var excelRange = AddWorksheet(retVal);
            WorksheetHelper.FormatAsTable(excelRange, TableStyle, WorksheetName);

            return retVal;
        }

        public ExcelWorksheet AddWorksheetToExistingPackage(ExcelPackage package)
        {
            var excelRange = AddWorksheet(package);
            WorksheetHelper.FormatAsTable(excelRange, TableStyle, WorksheetName);

            return excelRange.Worksheet;
        }
        #endregion

        #region Private

        internal ExcelRange AddWorksheet(ExcelPackage package)
        {
            //Avoid multiple enumeration
            var myData = Data as IList<T> ?? Data.ToList();

            if (Data == null || !myData.Any())
                return null;

            var properties = myData.First().GetType().GetProperties();
            var worksheet = package.Workbook.Worksheets.Add(WorksheetName);

            //Create table header
            var iCol = 0;
            foreach (var property in properties)
            {
                if (IgnoredProperties.Contains(property.Name))
                    continue;

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
                    if (IgnoredProperties.Contains(property.Name))
                        continue;

                    iCol++;
                    worksheet.Cells[iRow, iCol].Value = GetPropertyValue(property, item);
                }
            }

            return worksheet.Cells[1, 1, myData.Count + 1, iCol];
        }

        internal object GetPropertyValue(PropertyInfo property, object item)
        {
            if (DisplayFormats.ContainsKey(property.Name))
            {
                var value = property.GetValue(item);
                if (value != null)
                    return string.Format(DisplayFormats[property.Name], value);
            }

            return ReflectionHelper.GetPropertyValue(property, item);
        }

        #endregion

        #region Properties

        public string WorksheetName { get; set; } = "Data";

        public TableStyles TableStyle { get; set; } = TableStyles.Medium4;

        public IEnumerable<T> Data { get; set; }

        #endregion

        #region IExportFormat<T>
        public Dictionary<string, string> DisplayFormats { get; set; } = new Dictionary<string, string>();
        public EnumerableExporter<T> DisplayFormatFor<TProperty>(Expression<Func<T, TProperty>> property, string format)
        {
            var propertyName = PropertyName.For(property);
            if (DisplayFormats.ContainsKey(propertyName))
                DisplayFormats[propertyName] = format;
            else
                DisplayFormats.Add(propertyName, format);

            return this;
        }

        public HashSet<string> IgnoredProperties { get; set; } = new HashSet<string>();
        public EnumerableExporter<T> Ignore<TProperty>(Expression<Func<T, TProperty>> property)
        {
            var propertyName = PropertyName.For(property);
            IgnoredProperties.Add(propertyName);

            return this;
        }
        #endregion
    }
}
