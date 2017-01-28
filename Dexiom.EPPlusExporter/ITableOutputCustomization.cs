using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml.Style;

namespace Dexiom.EPPlusExporter
{
    public interface ITableOutputCustomization<T> 
        where T : class
    {
        TableExporter<T> TextFormatFor<TProperty>(Expression<Func<T, TProperty>> property, string format);
        
        TableExporter<T> Ignore<TProperty>(Expression<Func<T, TProperty>> property);
        
        TableExporter<T> StyleFor<TProperty>(Expression<Func<T, TProperty>> property, Action<ExcelStyle> setStyle);

        TableExporter<T> DefaultNumberFormat(Type type, string numberFormat);
        TableExporter<T> NumberFormatFor<TProperty>(Expression<Func<T, TProperty>> property, string numberFormat);
    }
}
