using System;
using System.Linq.Expressions;
using OfficeOpenXml.Style;

namespace Dexiom.EPPlusExporter.Interfaces
{
    public interface ITableOutputCustomization<T> 
        where T : class
    {
        TableExporter<T> Configure<TProperty>(Expression<Func<T, TProperty>> property, Action<ColumnConfiguration> column);

        #region Shorthands for configure
        TableExporter<T> StyleFor<TProperty>(Expression<Func<T, TProperty>> property, Action<ExcelStyle> setStyle);

        TableExporter<T> HeaderStyleFor<TProperty>(Expression<Func<T, TProperty>> property, Action<ExcelStyle> setStyle);

        TableExporter<T> NumberFormatFor<TProperty>(Expression<Func<T, TProperty>> property, string numberFormat);
        
        #endregion

        #region Misc
        TableExporter<T> Ignore<TProperty>(Expression<Func<T, TProperty>> property);

        TableExporter<T> DefaultNumberFormat(Type type, string numberFormat);

        TableExporter<T> TextFormatFor<TProperty>(Expression<Func<T, TProperty>> property, string format);
        
        TableExporter<T> ConditionalStyleFor<TProperty>(Expression<Func<T, TProperty>> property, Action<T, ExcelStyle> setStyle);

        #endregion
    }
}
