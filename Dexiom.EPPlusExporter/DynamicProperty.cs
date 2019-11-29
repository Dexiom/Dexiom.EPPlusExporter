using System;

namespace Dexiom.EPPlusExporter
{
    public class DynamicProperty<T>
    {
        public string Name { get; set; }
        public string DisplayName { get; set; }
        public Type ValueType { get; set; }
        public Func<T, object> GetValue { get; set; }
    }
}