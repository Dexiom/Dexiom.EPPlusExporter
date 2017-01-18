using System;
using System.Reflection;

namespace Dexiom.EPPlusExporter.Extensions
{
    internal static class MemberInfoExtensions
    {
        public static T GetCustomAttribute<T>(this MemberInfo element, bool inherit) where T : Attribute => (T)Attribute.GetCustomAttribute(element, typeof(T), inherit);

        public static T GetCustomAttribute<T>(this MemberInfo element) where T : Attribute => (T)Attribute.GetCustomAttribute(element, typeof(T));
    }
}
