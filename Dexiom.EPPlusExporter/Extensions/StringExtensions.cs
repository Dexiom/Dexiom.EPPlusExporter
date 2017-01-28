using System.Collections.Generic;
using System.Reflection;
using System.Text.RegularExpressions;

namespace Dexiom.EPPlusExporter.Extensions
{
    internal static class StringExtensions
    {
        public static string SplitCamelCase(this string source)
        {
            return Regex.Replace(
                Regex.Replace(
                    source,
                    @"(\P{Ll})(\P{Ll}\p{Ll})",
                    "$1 $2"
                ),
                @"(\p{Ll})(\P{Ll})",
                "$1 $2"
            );
        }
    }
}
