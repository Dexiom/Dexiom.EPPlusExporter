using Microsoft.VisualStudio.TestTools.UnitTesting;
using Dexiom.EPPlusExporter.Extensions;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dexiom.EPPlusExporter.Extensions.Tests
{
    [TestClass()]
    public class MemberInfoExtensionsTests
    {
        [DisplayName("MyDisplayName")]
        public DateTime MyTestProperty => DateTime.Now;

        [TestMethod()]
        public void GetCustomAttributeTest()
        {
            var prop = typeof(MemberInfoExtensionsTests).GetProperty("MyTestProperty");
            var attr1 = prop.GetCustomAttribute<DisplayNameAttribute>();
            var attr2 = prop.GetCustomAttribute<DisplayNameAttribute>(true);

            Assert.IsTrue(attr1.DisplayName == "MyDisplayName" && attr2.DisplayName == "MyDisplayName");
        }
    }
}