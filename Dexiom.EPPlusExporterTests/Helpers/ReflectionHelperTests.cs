using Microsoft.VisualStudio.TestTools.UnitTesting;
using Dexiom.EPPlusExporter.Helpers;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dexiom.EPPlusExporter.Helpers.Tests
{
    [TestClass()]
    public class ReflectionHelperTests
    {
        #region WeirdType (for testing)
        private class WeirdType : IEnumerable
        {
            [Display(Name = "SomeProp Name")]
            public string SomeProp { get; set; }
            
            [DisplayName("AnotherProp Name")]
            public string AnotherProp { get; set; }

            public string MyCoolProp { get; set; }

            public IEnumerator GetEnumerator()
            {
                throw new NotImplementedException();
            }
        }
        #endregion

        [TestMethod()]
        public void GetPropertyDisplayNameTest()
        {
            var myType = typeof(WeirdType);
            Assert.IsTrue(ReflectionHelper.GetPropertyDisplayName(myType.GetMember("SomeProp").First()) == "SomeProp Name");
            Assert.IsTrue(ReflectionHelper.GetPropertyDisplayName(myType.GetMember("AnotherProp").First()) == "AnotherProp Name");
            Assert.IsTrue(ReflectionHelper.GetPropertyDisplayName(myType.GetMember("MyCoolProp").First()) == "My Cool Prop");
            Assert.IsTrue(ReflectionHelper.GetPropertyDisplayName(myType.GetMember("MyCoolProp").First(), false) == "MyCoolProp");
        }

        [TestMethod()]
        public void GetBaseTypeOfEnumerableTest()
        {
            try
            {
                //null paramter
                ReflectionHelper.GetBaseTypeOfEnumerable(null);
                Assert.Fail();
            }
            catch (ArgumentNullException) { }
            catch (Exception) { Assert.Fail(); }
            
            try
            {
                //wrong implementation
                ReflectionHelper.GetBaseTypeOfEnumerable(new WeirdType());
                Assert.Fail();
            }
            catch (ArgumentException) { }
            catch (Exception) { Assert.Fail(); }


            Assert.IsTrue(ReflectionHelper.GetBaseTypeOfEnumerable(new [] { "toto", "titi", "tata" }) == typeof(string));
            Assert.IsTrue(ReflectionHelper.GetBaseTypeOfEnumerable(new List<int>()) == typeof(int));
            Assert.IsTrue(ReflectionHelper.GetBaseTypeOfEnumerable(Enumerable.Empty<double>()) == typeof(double));
        }
    }
}