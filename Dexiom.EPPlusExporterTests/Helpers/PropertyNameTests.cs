using Microsoft.VisualStudio.TestTools.UnitTesting;
using Dexiom.EPPlusExporter.Helpers;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;

namespace Dexiom.EPPlusExporter.Helpers.Tests
{
    [TestClass()]
    public class PropertyNameTests
    {
        [TestMethod()]
        public void ForTest()
        {
            Assert.IsTrue(PropertyName.For(() => new Dictionary<string, int>().Keys) == "Keys");
        }
    }
}