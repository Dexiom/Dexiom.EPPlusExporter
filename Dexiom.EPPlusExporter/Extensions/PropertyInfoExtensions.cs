using System.Reflection;

namespace Dexiom.EPPlusExporter.Extensions
{
    internal static class PropertyInfoExtensions
    {
        public static object GetValue(this PropertyInfo element, object obj) => element.GetValue(obj, BindingFlags.Default, null, null, null);
    }
}
