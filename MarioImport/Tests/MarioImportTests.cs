using System;
using System.Collections.Generic;
using System.Data.SqlClient;
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

        [TestMethod]
        public void TestGetStoresFromDatabase()
        {
            using (SqlConnection connection = new SqlConnection("Data Source = sql6009.site4now.net; Initial Catalog = DB_A2C9F3_MarioPizza; Persist Security Info = True; User ID = DB_A2C9F3_MarioPizza_admin; Password = Februarie2020!"))
            {
                List<String> columnData = new List<String>();
                connection.Open();
                string query = "SELECT DISTINCT [Name] FROM [Store-QL] UNION SELECT DISTINCT [Name] FROM Store ORDER BY [Name];";
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            columnData.Add(reader.GetString(0));
                        }
                    }

                }
                Assert.IsTrue(columnData != null);
            }
        }
    }
}
