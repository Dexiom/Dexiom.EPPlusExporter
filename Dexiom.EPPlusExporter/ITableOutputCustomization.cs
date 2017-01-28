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
        TableExporter<T> DisplayFormatFor<TProperty>(Expression<Func<T, TProperty>> property, string format);

        
        TableExporter<T> Ignore<TProperty>(Expression<Func<T, TProperty>> property);
        
        TableExporter<T> StyleFor<TProperty>(Expression<Func<T, TProperty>> property, Action<ExcelStyle> initStyle);
    }
}
