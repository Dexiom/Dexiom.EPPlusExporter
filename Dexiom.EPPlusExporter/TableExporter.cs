using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Dexiom.EPPlusExporter.Extensions;
using Dexiom.EPPlusExporter.Helpers;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;

namespace Dexiom.EPPlusExporter
{
    public abstract class TableExporter<T> : IExporter, ITableOutput, ITableOutputCustomization<T>
        where T : class
    {
        protected abstract ExcelRange AddWorksheet(ExcelPackage package);

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

        #region ITableOutput
        public string WorksheetName { get; set; } = "Data";

        public TableStyles TableStyle { get; set; } = TableStyles.Medium4;
        #endregion

        #region ITableOutputCustomization<T>
        public TableExporter<T> TextFormatFor<TProperty>(Expression<Func<T, TProperty>> property, string format)
        {
            TextFormats.AddOrUpdate(PropertyName.For(property), format);
            return this;
        }

        public TableExporter<T> Ignore<TProperty>(Expression<Func<T, TProperty>> property)
        {
            var propertyName = PropertyName.For(property);
            IgnoredProperties.Add(propertyName);
            return this;
        }

        public TableExporter<T> StyleFor<TProperty>(Expression<Func<T, TProperty>> property, Action<ExcelStyle> setStyle)
        {
            ColumnStyles.AddOrUpdate(PropertyName.For(property), setStyle);
            return this;
        }

        public TableExporter<T> DefaultNumberFormat(Type type, string numberFormat)
        {
            DefaultNumberFormats.AddOrUpdate(type, numberFormat);
            return this;
        }

        public TableExporter<T> NumberFormatFor<TProperty>(Expression<Func<T, TProperty>> property, string numberFormat)
        {
            NumberFormats.AddOrUpdate(PropertyName.For(property), numberFormat);
            return this;
        }

        #endregion


        #region Protected 
        protected Dictionary<string, string> TextFormats { get; set; } = new Dictionary<string, string>();

        protected HashSet<string> IgnoredProperties { get; set; } = new HashSet<string>();

        protected Dictionary<string, Action<ExcelStyle>> ColumnStyles { get; set; } = new Dictionary<string, Action<ExcelStyle>>();

        protected Dictionary<Type, string> DefaultNumberFormats { get; set; } = new Dictionary<Type, string>
        {
            { typeof(DateTime), "yyyy-MM-dd HH:mm:ss" },
            { typeof(DateTime?), "yyyy-MM-dd HH:mm:ss" }
        };

        protected Dictionary<string, string> NumberFormats { get; set; } = new Dictionary<string, string>();

        protected object GetPropertyValue(PropertyInfo property, object item)
        {
            var value = property.GetValue(item);
            if (value != null && TextFormats.ContainsKey(property.Name))
                return string.Format(TextFormats[property.Name], value);

            return value;
        }
        #endregion
    }
}
