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
    public class PropertyNamesTests
    {
        [TestMethod()]
        public void ForTest()
        {
            Assert.IsTrue(PropertyNames.For(() => new Dictionary<string, int>().Keys)[0] == "Keys"); //UnaryExpression

            Assert.IsTrue(PropertyNames.For<Tuple<string, int, double>>(n => n.Item2)[0] == "Item2"); //MemberExpression

            var myPropNames = PropertyNames.For<Tuple<string, int, double>>(n => new { n.Item1, n.Item3 }); //NewExpression
            Assert.IsTrue(myPropNames[0] == "Item1");
            Assert.IsTrue(myPropNames[1] == "Item3");
        }
    }
}