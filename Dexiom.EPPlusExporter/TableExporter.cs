using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Dexiom.EPPlusExporter.Extensions;
using Dexiom.EPPlusExporter.Helpers;
using Dexiom.EPPlusExporter.Interfaces;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;

namespace Dexiom.EPPlusExporter
{
    public abstract class TableExporter<T> : IExporter, ITableOutput, ITableOutputCustomization<T>
        where T : class
    {
        private readonly List<KeyValuePair<string, Action<ColumnConfiguration>>> _columnAlterations = new List<KeyValuePair<string, Action<ColumnConfiguration>>>();
        private readonly List<Action<ExcelRange>> _tableCustomizations = new List<Action<ExcelRange>>();

        #region Abstract
        protected abstract ExcelRange AddWorksheet(ExcelPackage package);
        #endregion

        #region Public Functions
        public ExcelPackage CreateExcelPackage()
        {
            var retVal = new ExcelPackage();
            var excelRange = AddWorksheet(retVal);
            if (excelRange == null)
                return null;

            //apply table customizations
            foreach (var tableCustomization in _tableCustomizations)
                tableCustomization(excelRange);

            return retVal;
        }

        public ExcelWorksheet AppendToExcelPackage(ExcelPackage package)
        {
            if (package == null)
                throw new ArgumentNullException(nameof(package));

            var excelRange = AddWorksheet(package);
            if (excelRange == null)
                return null;

            //apply table customizations
            foreach (var tableCustomization in _tableCustomizations)
                tableCustomization(excelRange);

            return excelRange.Worksheet;
        }
        #endregion

        #region ITableOutputCustomization<T>

        public TableExporter<T> Configure<TProperty>(Expression<Func<T, TProperty>> property, Action<ColumnConfiguration> column)
        {
            foreach (var propName in PropertyNames.For(property))
                _columnAlterations.Add(new KeyValuePair<string, Action<ColumnConfiguration>>(propName, column));

            return this;
        }
        
        public TableExporter<T> CustomizeTable(Action<ExcelRange> applyCustomization)
        {
            _tableCustomizations.Add(applyCustomization);
            return this;
        }

        public TableExporter<T> StyleFor<TProperty>(Expression<Func<T, TProperty>> property, Action<ExcelStyle> setStyle)
        {
            Configure(property, c => c.Content.SetStyle = setStyle);
            return this;
        }

        public TableExporter<T> HeaderStyleFor<TProperty>(Expression<Func<T, TProperty>> property, Action<ExcelStyle> setStyle)
        {
            Configure(property, c => c.Header.SetStyle = setStyle);
            return this;
        }

        public TableExporter<T> NumberFormatFor<TProperty>(Expression<Func<T, TProperty>> property, string numberFormat)
        {
            Configure(property, c => c.Content.NumberFormat = numberFormat);
            return this;
        }

        public TableExporter<T> Display(IEnumerable<string> propertyNames)
        {
            if (DisplayedProperties == null)
                DisplayedProperties = new HashSet<string>();

            foreach (var propName in propertyNames)
                DisplayedProperties.Add(propName);

            return this;
        }
        
        public TableExporter<T> Display<TProperty>(Expression<Func<T, TProperty>> properties)
        {
            Display(PropertyNames.For(properties));

            return this;
        }

        public TableExporter<T> Ignore(IEnumerable<string> propertyNames)
        {
            foreach (var propName in propertyNames)
                IgnoredProperties.Add(propName);

            return this;
        }

        public TableExporter<T> Ignore<TProperty>(Expression<Func<T, TProperty>> properties)
        {
            Ignore(PropertyNames.For(properties));
            return this;
        }

        public TableExporter<T> DefaultNumberFormat(Type type, string numberFormat)
        {
            DefaultNumberFormats.AddOrUpdate(type, numberFormat);
            return this;
        }

        public TableExporter<T> TextFormatFor<TProperty>(Expression<Func<T, TProperty>> property, string format)
        {
            foreach (var propName in PropertyNames.For(property))
                TextFormats.AddOrUpdate(propName, format);

            return this;
        }
        
        public TableExporter<T> ConditionalStyleFor<TProperty>(Expression<Func<T, TProperty>> property, Action<T, ExcelStyle> setStyle)
        {
            foreach (var propName in PropertyNames.For(property))
                ConditionalStyles.AddOrUpdate(propName, setStyle);

            return this;
        }
        
        #endregion
        
        #region Protected
        protected Dictionary<string, string> TextFormats { get; } = new Dictionary<string, string>();

        protected HashSet<string> DisplayedProperties { get; private set; }

        protected HashSet<string> IgnoredProperties { get; } = new HashSet<string>();

        protected Dictionary<string, Action<T, ExcelStyle>> ConditionalStyles { get; } = new Dictionary<string, Action<T, ExcelStyle>>();

        protected Dictionary<Type, string> DefaultNumberFormats { get; } = new Dictionary<Type, string>
        {
            { typeof(DateTime), "yyyy-MM-dd HH:mm:ss" },
            { typeof(DateTime?), "yyyy-MM-dd HH:mm:ss" }
        };
        
        protected object GetPropertyValue(PropertyInfo property, object item)
        {
#if NET4
            var value = property.GetValue(item, null);
#endif
#if NET45 || NET46
            var value = property.GetValue(item);
#endif
            if (value != null && TextFormats.ContainsKey(property.Name))
                return string.Format(TextFormats[property.Name], value);

            return value;
        }

        protected Dictionary<string, ColumnConfiguration> GetColumnConfigurations(IEnumerable<string> columnNames)
        {
            var retVal = new Dictionary<string, ColumnConfiguration>();
            foreach (var colName in columnNames)
            {
                var newConfig = new ColumnConfiguration();

                //apply all the alterations to the column definition
                var alterations = _columnAlterations.Where(n => n.Key == colName);
                foreach (var alteration in alterations)
                    alteration.Value(newConfig);

                retVal.Add(colName, newConfig);
            }

            return retVal;
        }
        #endregion

        #region Properties

        public string WorksheetName { get; set; } = "Data";

        public TableStyles TableStyle { get; set; } = TableStyles.None;

        public bool AutoFitColumns { get; set; } = true;

        #endregion

    }
}
