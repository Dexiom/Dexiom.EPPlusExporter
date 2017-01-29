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
using OfficeOpenXml.Style;
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
            const int headerFirstRow = 1;
            const int headerFirstCol = 1;
            const int dataFirstRow = 2;
            const int dataFirstCol = 1;

            //let's avoid multiple enumeration
            var myData = Data as IList<T> ?? Data.ToList();

            if (Data == null || !myData.Any())
                return null;

            var properties = myData.First().GetType().GetProperties();
            var worksheet = package.Workbook.Worksheets.Add(WorksheetName);
            var displayedProperties = properties.Where(p => !IgnoredProperties.Contains(p.Name)).ToList();

            //Create table header
            {
                var col = headerFirstCol;
                foreach (var property in displayedProperties)
                {
                    worksheet.Cells[headerFirstRow, col].Value = ReflectionHelper.GetPropertyDisplayName(property);
                    col++;
                }
            }

            //Add rows
            var row = dataFirstRow;
            foreach (var item in myData)
            {
                var iCol = dataFirstCol;
                foreach (var property in displayedProperties)
                {
                    worksheet.Cells[row, iCol].Value = GetPropertyValue(property, item);
                    iCol++;
                }
                row++;
            }
            
            //get bottom & right bounds
            var dataLastCol = dataFirstCol + displayedProperties.Count - 1;
            var dataLastRow = dataFirstRow + myData.Count - 1;

            //apply styles
            {
                var iCol = dataFirstCol;
                foreach (var property in displayedProperties)
                {
                    var columnRange = worksheet.Cells[dataFirstRow, iCol, dataLastRow, iCol];

                    //apply default number format
                    if (DefaultNumberFormats.ContainsKey(property.PropertyType))
                        columnRange.Style.Numberformat.Format = DefaultNumberFormats[property.PropertyType];

                    //apply number format
                    if (NumberFormats.ContainsKey(property.Name))
                        columnRange.Style.Numberformat.Format = NumberFormats[property.Name];

                    //apply style
                    if (ColumnStyles.ContainsKey(property.Name))
                        ColumnStyles[property.Name](columnRange.Style);

                    iCol++;
                }
            }
        
            //return the entire grid range
            return worksheet.Cells[headerFirstRow, headerFirstCol, dataLastRow, dataLastCol]; 
        }

        #endregion
    }
}
