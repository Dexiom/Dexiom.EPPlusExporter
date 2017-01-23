using System.Collections.Generic;
using System.Reflection;

namespace Dexiom.EPPlusExporter
{
    internal static class EnumerableExtensions
    {
        public static EnumerableExporter<T> GetExcelExporter<T>(this IEnumerable<T> source) where T : class => new EnumerableExporter<T>(source);
    }
}
