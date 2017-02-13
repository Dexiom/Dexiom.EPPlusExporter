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
        public static ObjectExporter<T> Create<T>(T data, TableStyles tableStyles = TableStyles.Medium4) where T : class => new ObjectExporter<T>(data) { TableStyle = tableStyles };
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
            var displayedProperties = properties.Where(p => !IgnoredProperties.Contains(p.Name)).ToList();
            var columnConfigurations = GetColumnConfigurations(displayedProperties.Select(n => n.Name));

            //Create table header
            worksheet.Cells[1, 1].Value = "Item";
            worksheet.Cells[1, 2].Value = "Value";

            //Add rows
            var myData = displayedProperties.Select(property => new
            {
                Property = property,
                Value = GetPropertyValue(property, Data)
            });

            var iRow = 1;
            foreach (var item in myData)
            {
                var colConfig = columnConfigurations[item.Property.Name];

                iRow++;
                var nameCell = worksheet.Cells[iRow, 1];
                var valueCell = worksheet.Cells[iRow, 2];
                nameCell.Value = string.IsNullOrEmpty(colConfig.Header.Text) ? ReflectionHelper.GetPropertyDisplayName(item.Property) : colConfig.Header.Text;
                valueCell.Value = item.Value;


                //apply default number format
                if (DefaultNumberFormats.ContainsKey(item.Property.PropertyType))
                    valueCell.Style.Numberformat.Format = DefaultNumberFormats[item.Property.PropertyType];

                //apply number format
                if (colConfig.Content.NumberFormat != null)
                    valueCell.Style.Numberformat.Format = colConfig.Content.NumberFormat;

                //apply style
                colConfig.Header.SetStyle(nameCell.Style);
                colConfig.Content.SetStyle(valueCell.Style);

                //apply conditional styles
                if (ConditionalStyles.ContainsKey(item.Property.Name))
                    ConditionalStyles[item.Property.Name](Data, worksheet.Cells[iRow, 1, iRow, 2].Style);
            }

            return worksheet.Cells[1, 1, iRow, 2];
        }
        #endregion
    }
}
