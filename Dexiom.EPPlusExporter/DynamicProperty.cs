using System;
using System.Collections.Generic;

namespace Dexiom.EPPlusExporter
{
    #region Create Method (using type inference)
    public static class DynamicProperty
    {
        public static DynamicProperty<T> Create<T>(IEnumerable<T> data, string name, string displayName, Type valueType, Func<T, object> getValue) where T : class
        {
            return new DynamicProperty<T>()
            {
                Name = name,
                DisplayName = displayName,
                ValueType = valueType,
                GetValue = getValue
            };
        }
    }
    #endregion

    public class DynamicProperty<T> 
        where T : class
    {
        public string Name { get; set; }
        public string DisplayName { get; set; }
        public Type ValueType { get; set; }
        public Func<T, object> GetValue { get; set; }
    }
}