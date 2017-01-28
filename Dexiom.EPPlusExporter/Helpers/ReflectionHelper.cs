using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Dexiom.EPPlusExporter.Extensions;

namespace Dexiom.EPPlusExporter.Helpers
{
    internal static class ReflectionHelper
    {
        internal static string GetPropertyDisplayName(MemberInfo property)
        {
            var displayNameAttribute = MemberInfoExtensions.GetCustomAttribute<DisplayNameAttribute>(property, true);
            if (displayNameAttribute != null)
                return displayNameAttribute.DisplayName;

            var displayAttribute = MemberInfoExtensions.GetCustomAttribute<DisplayAttribute>(property, true);
            if (displayAttribute != null)
                return displayAttribute.Name;

            //well, no attribue found, let's just take that property's name then...
            return property.Name;
        }
        
        internal static object GetPropertyDisplayValue(PropertyInfo property, object item)
        {
            var value = property.GetValue(item);

            //check for customization via attribute
            var displayFormatAttribute = MemberInfoExtensions.GetCustomAttribute<DisplayFormatAttribute>(property, true);
            if (displayFormatAttribute != null)
            {
                //handle NullDisplayText 
                if (value == null && !string.IsNullOrWhiteSpace(displayFormatAttribute.NullDisplayText))
                    return displayFormatAttribute.NullDisplayText;

                //handle display format
                if (value != null)
                    return string.Format(displayFormatAttribute.DataFormatString, value);
            }

            //value is null, nothing else to do
            if (value == null)
                return string.Empty;

            //simple type
            if (property.PropertyType.IsValueType)
                return value;

            //enumerable
            var enumerable = (value as IEnumerable<object>);
            if (enumerable != null)
                return enumerable.Count(); //just show the count

            //well, let's throw the value...
            return value.ToString();
        }
    }
}
