using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;

namespace Dexiom.EPPlusExporter
{
    public interface IExportFormat<T> 
        where T : class
    {
        Dictionary<string, string> DisplayFormats { get; set; }

        EnumerableExporter<T> DisplayFormatFor<TProperty>(Expression<Func<T, TProperty>> property, string format);


        HashSet<string> IgnoredProperties { get; set; }
        EnumerableExporter<T> Ignore<TProperty>(Expression<Func<T, TProperty>> property);
    }
}
