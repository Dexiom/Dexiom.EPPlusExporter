using System;
using System.Linq.Expressions;
using OfficeOpenXml.Style;

namespace Dexiom.EPPlusExporter.Interfaces
{
    public interface ITableOutputCustomization<T> 
        where T : class
    {
        TableExporter<T> TextFormatFor<TProperty>(Expression<Func<T, TProperty>> property, string format);
        
        TableExporter<T> Ignore<TProperty>(Expression<Func<T, TProperty>> property);
        
        TableExporter<T> StyleFor<TProperty>(Expression<Func<T, TProperty>> property, Action<ExcelStyle> setStyle);

        TableExporter<T> ConditionalStyleFor<TProperty>(Expression<Func<T, TProperty>> property, Action<T, ExcelStyle> setStyle);

        TableExporter<T> DefaultNumberFormat(Type type, string numberFormat);
        TableExporter<T> NumberFormatFor<TProperty>(Expression<Func<T, TProperty>> property, string numberFormat);
    }
}
