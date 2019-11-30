using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace Dexiom.EPPlusExporter.Interfaces
{
    public interface ITableOutputCustomization<T> 
        where T : class
    {
        TableExporter<T> Configure<TProperty>(Expression<Func<T, TProperty>> properties, Action<ColumnConfiguration> column);
        TableExporter<T> Configure(IEnumerable<string> propertyNames, Action<ColumnConfiguration> column);
        
        TableExporter<T> CustomizeTable(Action<ExcelRange> applyCustomization);


        #region Shorthands for configure

        TableExporter<T> StyleFor<TProperty>(Expression<Func<T, TProperty>> properties, Action<ExcelStyle> setStyle);
        TableExporter<T> StyleFor(IEnumerable<string> propertyNames, Action<ExcelStyle> setStyle);

        TableExporter<T> HeaderStyleFor<TProperty>(Expression<Func<T, TProperty>> properties, Action<ExcelStyle> setStyle);
        TableExporter<T> HeaderStyleFor(IEnumerable<string> propertyNames, Action<ExcelStyle> setStyle);

        TableExporter<T> NumberFormatFor<TProperty>(Expression<Func<T, TProperty>> properties, string numberFormat);
        TableExporter<T> NumberFormatFor(IEnumerable<string> propertyNames, string numberFormat);

        #endregion


        #region Misc
        TableExporter<T> DefaultNumberFormat(Type type, string numberFormat);

        TableExporter<T> Display<TProperty>(Expression<Func<T, TProperty>> properties);
        TableExporter<T> Display(IEnumerable<string> propertyNames);

        TableExporter<T> Ignore<TProperty>(Expression<Func<T, TProperty>> properties);
        TableExporter<T> Ignore(IEnumerable<string> propertyNames);


        TableExporter<T> TextFormatFor<TProperty>(Expression<Func<T, TProperty>> properties, string format);
        TableExporter<T> TextFormatFor(IEnumerable<string> propertyNames, string format);

        TableExporter<T> ConditionalStyleFor<TProperty>(Expression<Func<T, TProperty>> properties, Action<T, ExcelStyle> setStyle);
        TableExporter<T> ConditionalStyleFor(IEnumerable<string> propertyNames, Action<T, ExcelStyle> setStyle);

        TableExporter<T> FormulaFor<TProperty>(Expression<Func<T, TProperty>> properties, Func<T, object, string> formulaFormat);
        TableExporter<T> FormulaFor(IEnumerable<string> propertyNames, Func<T, object, string> formulaFormat);

        #endregion
    }
}
