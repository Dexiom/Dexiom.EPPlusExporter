using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;

namespace Dexiom.EPPlusExporter
{
    public class ColumnConfiguration
    {
        public ColumnHeaderConfiguration Header { get; } = new ColumnHeaderConfiguration();
        public ColumnContentConfiguration Content { get; } = new ColumnContentConfiguration();
    }
}
