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
    #region Create Method (using type inference)
    public class EnumerableExporter
    {
        public static EnumerableExporter<T> Create<T>(IEnumerable<T> data) where T : class => new EnumerableExporter<T>(data);
    }
    #endregion

    public class EnumerableExporter<T> : TableExporter<T>
        where T : class
    {
        public IEnumerable<T> Data { get; set; }

        #region Constructors
        public EnumerableExporter(IEnumerable<T> data)
        {
            Data = data;
        }
        #endregion

        #region Protected
        protected override ExcelRange AddWorksheet(ExcelPackage package)
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
            var iRow = 2;
            foreach (var item in myData)
            {
                iCol = 0;
                foreach (var property in properties)
                {
                    if (IgnoredProperties.Contains(property.Name))
                        continue;

                    iCol++;
                    worksheet.Cells[iRow, iCol].Value = GetPropertyValue(property, item);
                }

                iRow++;
            }

            return worksheet.Cells[1, 1, myData.Count + 1, iCol];
        }

        #endregion

        #region Private
        private object GetPropertyValue(PropertyInfo property, object item)
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
    }
}
