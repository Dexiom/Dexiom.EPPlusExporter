using System;
using System.Collections.Generic;
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
    public class ObjectExporter
    {
        public static ObjectExporter<T> Create<T>(T data) where T : class => new ObjectExporter<T>(data);
    }
    #endregion


    public class ObjectExporter<T> : TableExporter<T>
        where T : class
    {
        public T Data { get; set; }

        #region Constructors
        public ObjectExporter(T data)
        {
            Data = data;
        }
        #endregion
        
        #region Protected
        protected override ExcelRange AddWorksheet(ExcelPackage package)
        {
            if (Data == null)
                return null;

            var properties = Data.GetType().GetProperties();
            var worksheet = package.Workbook.Worksheets.Add(WorksheetName);

            //Create table header
            worksheet.Cells[1, 1].Value = "Item";
            worksheet.Cells[1, 2].Value = "Value";

            //Add rows
            var myData = properties.Select(property => new
            {
                Name = ReflectionHelper.GetPropertyDisplayName(property),
                Value = ReflectionHelper.GetPropertyValue(property, Data)
            });

            var iRow = 1;
            foreach (var item in myData)
            {
                iRow++;
                worksheet.Cells[iRow, 1].Value = item.Name;
                worksheet.Cells[iRow, 2].Value = item.Value;
            }

            return worksheet.Cells[1, 1, iRow, 2];
        }
        #endregion
    }
}
