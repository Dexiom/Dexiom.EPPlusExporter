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
        }

        public DisplayField(DynamicProperty<T> dynamicProperty)
        {
            _dynamicProperty = dynamicProperty;
        }
        
        public string Name
        {
            get
            {
                if (_propertyInfo != null)
                    return _propertyInfo.Name;

                if (_dynamicProperty != null)
                    return _dynamicProperty.Name;

                throw new ArgumentException();
            }
        }

        public string DisplayName
        {
            get
            {
                if (_propertyInfo != null)
                    return ReflectionHelper.GetPropertyDisplayName(_propertyInfo);

                if (_dynamicProperty != null)
                    return _dynamicProperty.DisplayName;

                throw new ArgumentException();
            }
        }

        public Type Type
        {
            get
            {
                if (_propertyInfo != null)
                    return _propertyInfo.PropertyType;

                if (_dynamicProperty != null)
                    return _dynamicProperty.ValueType;

                throw new ArgumentException();
            }
        }

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

            if (_dynamicProperty != null)
                return _dynamicProperty.GetValue(item);

            throw new ArgumentException();
        }

    }
}