﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text.RegularExpressions;
using Dexiom.EPPlusExporter.Extensions;

namespace Dexiom.EPPlusExporter.Helpers
{
    internal static class ReflectionHelper
    {
        public static string GetPropertyDisplayName(MemberInfo property, bool splitCamelCase = true)
        {
            var displayNameAttribute = MemberInfoExtensions.GetCustomAttribute<DisplayNameAttribute>(property, true);
            if (displayNameAttribute != null)
                return displayNameAttribute.DisplayName;

#if NET45 || NET46 || NETSTANDARD2_1
            var displayAttribute = MemberInfoExtensions.GetCustomAttribute<System.ComponentModel.DataAnnotations.DisplayAttribute>(property, true);
            if (displayAttribute != null)
                return displayAttribute.Name;
#endif
            //well, no attribute found, let's just take that property's name then...
            return splitCamelCase ? SplitCamelCase(property.Name) : property.Name;
        }


        public static Type GetBaseTypeOfEnumerable(IEnumerable enumerable)
        {
            if (enumerable == null)
                throw new ArgumentNullException(nameof(enumerable));

            var genericEnumerableInterface = enumerable
                .GetType()
                .GetInterfaces()
                .FirstOrDefault(i => i.IsGenericType && i.GetGenericTypeDefinition() == typeof(IEnumerable<>));

            if (genericEnumerableInterface == null)
                throw new ArgumentException("IEnumerable<T> is not implemented", nameof(enumerable));

            var elementType = genericEnumerableInterface.GetGenericArguments()[0];
            return elementType.IsGenericType && elementType.GetGenericTypeDefinition() == typeof(Nullable<>)
                ? elementType.GetGenericArguments()[0]
                : elementType;
        }
        
#region Private
        private static string SplitCamelCase(string text)
        {
            return Regex.Replace(
                Regex.Replace(
                    text,
                    @"(\P{Ll})(\P{Ll}\p{Ll})",
                    "$1 $2"
                ),
                @"(\p{Ll})(\P{Ll})",
                "$1 $2"
            );
        }
#endregion
    }
}
