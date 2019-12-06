using System;

namespace Dexiom.EPPlusExporter
{
    public static class DynamicProperty
    {
        public static DynamicProp<T> Create<T>(Func<T, object> getValue, Type valueType) where T : class => new DynamicProp<T>(getValue, valueType);
    }

    public class DynamicProp<T>
    {
        public DynamicProp(Func<T, object> getValue, Type valueType)
        {
            ValueType = valueType;
            GetValue = getValue;
        }


        public Type ValueType { get; set; }
        public Func<T, object> GetValue { get; set; }
    }
}