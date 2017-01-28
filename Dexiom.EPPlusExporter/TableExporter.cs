using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;
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
        public TableExporter<T> DisplayFormatFor<TProperty>(Expression<Func<T, TProperty>> property, string format)
        {
            var propertyName = PropertyName.For(property);
            if (DisplayFormats.ContainsKey(propertyName))
                DisplayFormats[propertyName] = format;
            else
                DisplayFormats.Add(propertyName, format);

            return this;
        }

        public TableExporter<T> Ignore<TProperty>(Expression<Func<T, TProperty>> property)
        {
            var propertyName = PropertyName.For(property);
            IgnoredProperties.Add(propertyName);

            return this;
        }

        public TableExporter<T> StyleFor<TProperty>(Expression<Func<T, TProperty>> property, Action<ExcelStyle> initStyle)
        {
            var propertyName = PropertyName.For(property);
            if (CellStyles.ContainsKey(propertyName))
                CellStyles[propertyName] = initStyle;
            else
                CellStyles.Add(propertyName, initStyle);

            return this;
        }

        #endregion

        #region Protected 
        protected Dictionary<string, string> DisplayFormats { get; set; } = new Dictionary<string, string>();

        protected HashSet<string> IgnoredProperties { get; set; } = new HashSet<string>();

        protected Dictionary<string, Action<ExcelStyle>> CellStyles { get; set; } = new Dictionary<string, Action<ExcelStyle>>();
        #endregion
    }
}
