using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Tests
{
    [TestClass]
    public class MarioImportTests
    {
        [TestMethod]
        public void TestOrderLineModificationDistinction()
        {
            string[] ingredients = { "Ham", "Gorgonzola", "Spinazie", "Salami","Ham"};
            int expected = 4;
            int actual = 0;

            Dictionary<string, int> counts = ingredients.GroupBy(x => x)
                                      .ToDictionary(g => g.Key,
                                                    g => g.Count());
            actual = counts.Count;

            Assert.AreEqual(expected, actual);
        }
    }
}
