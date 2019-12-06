using System;
using System.ComponentModel;
using System.Reflection;
using Dexiom.EPPlusExporter.Helpers;

namespace Dexiom.EPPlusExporter
{
    public class DisplayField<T> where T : class
    {
        private readonly PropertyInfo _propertyInfo;
        private readonly DynamicProperty<T> _dynamicProperty;

        public DisplayField(PropertyInfo propertyInfo)
        {
            _propertyInfo = propertyInfo;

            Name = _propertyInfo.Name;
            DisplayName = ReflectionHelper.GetPropertyDisplayName(_propertyInfo);
            Type = _propertyInfo.PropertyType;
        }

        public DisplayField(DynamicProperty<T> dynamicProperty)
        {
            _dynamicProperty = dynamicProperty;

            Name = _dynamicProperty.Name;
            DisplayName = _dynamicProperty.DisplayName;
            Type = _dynamicProperty.ValueType;
        }

        #region Properties
        public string Name { get; set; }
        public string DisplayName { get; set; }
        public Type Type { get; set; }
        #endregion

        public object GetValue(T item)
        {
            if (_propertyInfo != null)
            {
#if NET4
                return _propertyInfo.GetValue(item, null);
#endif
#if NET45 || NET46
            return  _propertyInfo.GetValue(item);
#endif
            }

            return _dynamicProperty.GetValue(item);
        }
    }
}