using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using ExcelDataReader;
using MoreLinq;
using System.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using System.Text.RegularExpressions;

namespace MarioImport
{
    class Program
    {
        static void Main(string[] args)
        {
            var builder = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);

            IConfigurationRoot configuration = builder.Build();
            string basePath = configuration.GetSection("BasePath").Value;
            //Console.WriteLine(configuration.GetConnectionString("Storage"));
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            TestDennis(basePath);
            //TestJos(basePath);
            //TestChris(basePath);
        }

        private static void TestChris(string path)
        {
            CSVImport csvImport = new CSVImport(path);
            csvImport.importCSV("MarioOrderData01_10000.csv");
            csvImport.importCSV("MarioOrderData02_10000.csv");
            csvImport.importCSV("MarioOrderData03_10000.csv");
            csvImport.importCSV("MarioOrderData04_10000.csv");
            csvImport.importIngredients();
        }

        private static void TestDennis(string path)
        {
            var prod = GetProducts(path);
            var cat = GetCategories(path);
            WriteCategoriesToDB(cat);
            WriteProductsToDB(prod);
            CallSPForProduction();
            LinkCategories(prod);

            using (SqlConnection con = new SqlConnection("Data Source = sql6009.site4now.net; Initial Catalog = DB_A2C9F3_MarioPizza; Persist Security Info = True; User ID = DB_A2C9F3_MarioPizza_admin; Password = Februarie2020!"))
            using (SqlCommand cmd = new SqlCommand())
            {
                cmd.CommandText = "Insert into [Product_Combine](ProductID, PartID, Quantity, UOMID) Select a.ID, b.ID, @amount , (Select ID from UnitOfMeasure where name = 'Stuk') from product a, product b where a.Name = @prodName and b.Name = @ingredient";
                cmd.Connection = con;
                con.Open();
                foreach (Product p in prod)
                {
                    if (p.Ingredients != null)
                    {
                        foreach (Product ingr in p.Ingredients)
                        {
                            cmd.Parameters.Clear();
                            cmd.Parameters.AddWithValue("@prodname", p.Name);
                            cmd.Parameters.AddWithValue("@ingredient", ingr.Name);
                            cmd.Parameters.AddWithValue("@amount", ingr.amount);
                            cmd.ExecuteNonQuery();
                        }
                    }
                }
                con.Close();
            }

        }

        private static void CallSPForProduction()
        {
            SqlCommand sp = new SqlCommand("ProductsAndCategoriesToProduction", new SqlConnection("Data Source = sql6009.site4now.net; Initial Catalog = DB_A2C9F3_MarioPizza; Persist Security Info = True; User ID = DB_A2C9F3_MarioPizza_admin; Password = Februarie2020!"));
            sp.CommandType = CommandType.StoredProcedure;
            sp.Connection.Open();
            sp.ExecuteNonQuery();
            sp.Connection.Close();
        }

        private static void LinkCategories(List<Product> products)
        {
            using (SqlConnection con = new SqlConnection("Data Source = sql6009.site4now.net; Initial Catalog = DB_A2C9F3_MarioPizza; Persist Security Info = True; User ID = DB_A2C9F3_MarioPizza_admin; Password = Februarie2020!"))
            using (SqlCommand cmd = new SqlCommand())
            {
                cmd.CommandText = "Insert into [Product_ProductCategory_Connection](ProductID, ProductCatogoryID) Select product.ID, ProductCategory.ID FROM product , ProductCategory WHERE product.name = @prodname and ProductCategory.name = @catname";
                cmd.Connection = con;
                con.Open();
                foreach (Product p in products)
                {
                    foreach (string c in p.Categories)
                    {
                        cmd.Parameters.Clear();
                        cmd.Parameters.AddWithValue("@prodname", p.Name);
                        cmd.Parameters.AddWithValue("@catname", c);
                        cmd.ExecuteNonQuery();
                    }
                }
                con.Close();
            }
        }

        private static void TestJos(string path)
        {
            TXTImport import = new TXTImport(path);
            import.textImport();
            import.databasewrite();
        }

        private static List<string> GetCategories(string basePath, string productSpecific = "")
        {
            string[] files = { @"\pizza_ingredienten.xlsx", @"\Overige producten.xlsx" };
            List<string> result = new List<string>();
            foreach (string file in files)
            {
                using (FileStream stream = File.OpenRead(basePath + file))
                using (IExcelDataReader dr = ExcelReaderFactory.CreateOpenXmlReader(stream))
                {
                    DataSet data = dr.AsDataSet();
                    var table = data.Tables[0];
                    int categoryColumn = -1;
                    int subCatColumn = -1;
                    int prodNameColumn = -1;
                    for (int rowCount = 0; rowCount < table.Rows.Count; rowCount++)
                    {
                        if (rowCount == 0)
                        {
                            for (int columnCount = 0; columnCount < table.Rows[rowCount].ItemArray.Length; columnCount++)
                            {
                                if (table.Rows[rowCount].ItemArray[columnCount].ToString().ToLower() == "categorie")
                                {
                                    categoryColumn = columnCount;
                                }

                                if (table.Rows[rowCount].ItemArray[columnCount].ToString().ToLower() == "subcategorie")
                                {
                                    subCatColumn = columnCount;
                                }

                                if (table.Rows[rowCount].ItemArray[columnCount].ToString().ToLower() == "productnaam")
                                {
                                    prodNameColumn = columnCount;
                                }
                            }
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(productSpecific) && table.Rows[rowCount].ItemArray[prodNameColumn].ToString().ToLower() == productSpecific)
                            {
                                var specificResult = new List<string>();
                                if (table.Rows[rowCount].ItemArray[categoryColumn].ToString().ToLower().IndexOfAny(new char[] { '&', ',' }) > 0)
                                {
                                    specificResult.AddRange(table.Rows[rowCount].ItemArray[categoryColumn].ToString().Split(new char[] { '&', ',' }));
                                }
                                else
                                {
                                    specificResult.Add(table.Rows[rowCount].ItemArray[categoryColumn].ToString());
                                }

                                if (table.Rows[rowCount].ItemArray[subCatColumn].ToString().ToLower().IndexOfAny(new char[] { '&', ',' }) > 0)
                                {
                                    specificResult.AddRange(table.Rows[rowCount].ItemArray[subCatColumn].ToString().Split(new char[] { '&', ',' }));
                                }
                                else
                                {
                                    specificResult.Add(table.Rows[rowCount].ItemArray[subCatColumn].ToString());
                                }
                                return specificResult.Select(t => t.Trim()).Distinct().ToList();
                            }
                            else
                            {
                                if (table.Rows[rowCount].ItemArray[categoryColumn].ToString().ToLower().IndexOfAny(new char[] { '&', ',' }) > 0)
                                {
                                    result.AddRange(table.Rows[rowCount].ItemArray[categoryColumn].ToString().Split(new char[] { '&', ',' }));
                                }
                                else
                                {
                                    result.Add(table.Rows[rowCount].ItemArray[categoryColumn].ToString());
                                }

                                if (table.Rows[rowCount].ItemArray[subCatColumn].ToString().ToLower().IndexOfAny(new char[] { '&', ',' }) > 0)
                                {
                                    result.AddRange(table.Rows[rowCount].ItemArray[subCatColumn].ToString().Split(new char[] { '&', ',' }));
                                }
                                else
                                {
                                    result.Add(table.Rows[rowCount].ItemArray[subCatColumn].ToString());
                                }
                            }
                        }
                    }
                }
            }
            result.AddRange(new List<string> { "ingredient", "pizzabodem" });
            return result.Select(t => t.Trim()).Distinct().ToList();
        }

        private static List<Product> GetProducts(string basePath)
        {
            string[] files = { @"\pizza_ingredienten.xlsx", @"\Overige producten.xlsx", @"\pizzabodems.xlsx", @"\Extra Ingredienten.csv" };
            List<Product> result = new List<Product>();
            var pizzaList = GetProductWithDescription(basePath + files[0]);
            foreach (Product product in pizzaList)
            {
                product.Ingredients = GetIngredients(basePath + files[0], true, product.Name);
                product.Categories = GetCategories(basePath, product.Name);
            }
            var productList = GetProductWithDescription(basePath + files[1]);
            foreach (Product product in productList)
            {
                product.Categories = GetCategories(basePath, product.Name);
            }
            result.AddRange(pizzaList);
            result.AddRange(productList);
            result.AddRange(GetIngredients(basePath + files[3], false));
            result.AddRange(GetIngredients(basePath + files[0], true));
            result.AddRange(GetPizzaBottom(basePath + files[2]));
            return result.DistinctBy(p => p.Name).ToList();
        }

        private static List<Product> GetProductWithDescription(string pathToFile)
        {
            List<Product> result = new List<Product>();

            using (FileStream stream = File.OpenRead(pathToFile))
            using (IExcelDataReader dr = ExcelReaderFactory.CreateOpenXmlReader(stream))
            {
                int nameColumn = -1;
                int descriptionColumn = -1;
                int priceColumn = -1;
                int spicyColumn = -1;
                int vegetarianColumn = -1;
                int addChargeColumn = -1;
                DataSet data = dr.AsDataSet();
                var table = data.Tables[0];
                for (int rowCount = 0; rowCount < table.Rows.Count; rowCount++)
                {
                    if (rowCount == 0)
                    {
                        for (int columnCount = 0; columnCount < table.Rows[rowCount].ItemArray.Length; columnCount++)
                        {
                            string columnHeader = table.Rows[rowCount].ItemArray[columnCount].ToString().ToLower();
                            if (columnHeader.ToLower().Contains("productnaam"))
                            {
                                nameColumn = columnCount;
                            }
                            else if (columnHeader.ToLower().Contains("productomschrijving"))
                            {
                                descriptionColumn = columnCount;
                            }
                            else if (columnHeader.ToLower().Contains("prijs"))
                            {
                                priceColumn = columnCount;
                            }
                            else if (columnHeader.ToLower().Contains("spicy"))
                            {
                                spicyColumn = columnCount;
                            }
                            else if (columnHeader.ToLower().Contains("vegetarisch"))
                            {
                                vegetarianColumn = columnCount;
                            }
                            else if (columnHeader.ToLower().Contains("bezorgtoeslag"))
                            {
                                addChargeColumn = columnCount;
                            }
                        }
                    }
                    else
                    {
                        Product foundProduct = new Product
                        {
                            Name = nameColumn == -1 ? "" : table.Rows[rowCount].ItemArray[nameColumn].ToString().ToLower(),
                            Description = descriptionColumn == -1 ? "" : table.Rows[rowCount].ItemArray[descriptionColumn].ToString().ToLower(),
                            Price = priceColumn == -1 ? "0" : table.Rows[rowCount].ItemArray[priceColumn].ToString().ToLower(),
                            Spicy = spicyColumn == -1 ? false : table.Rows[rowCount].ItemArray[spicyColumn].ToString().ToLower() == "nee" ? false : true,
                            Vegetarisch = vegetarianColumn == -1 ? false : table.Rows[rowCount].ItemArray[vegetarianColumn].ToString().ToLower() == "nee" ? false : true,
                            AdditionalCharge = addChargeColumn == -(1) ? "0" : table.Rows[rowCount].ItemArray[addChargeColumn].ToString().ToLower(),
                        };
                        result.Add(foundProduct);
                    }
                }
            }
            return result.DistinctBy(p => p.Name).ToList();
        }

        private static List<Product> GetPizzaBottom(string pathToFile)
        {
            List<Product> result = new List<Product>();

            using (FileStream stream = File.OpenRead(pathToFile))
            using (IExcelDataReader dr = ExcelReaderFactory.CreateOpenXmlReader(stream))
            {
                int nameColumn = -1;
                int descriptionColumn = -1;
                int priceColumn = -1;
                int sizeColumn = -1;
                DataSet data = dr.AsDataSet();
                var table = data.Tables[0];
                for (int rowCount = 0; rowCount < table.Rows.Count; rowCount++)
                {
                    if (rowCount == 0)
                    {
                        for (int columnCount = 0; columnCount < table.Rows[rowCount].ItemArray.Length; columnCount++)
                        {
                            string columnHeader = table.Rows[rowCount].ItemArray[columnCount].ToString().ToLower();
                            if (columnHeader.ToLower().Contains("naam"))
                            {
                                nameColumn = columnCount;
                            }
                            else if (columnHeader.ToLower().Contains("omschrijving"))
                            {
                                descriptionColumn = columnCount;
                            }
                            else if (columnHeader.ToLower().Contains("toeslag"))
                            {
                                priceColumn = columnCount;
                            }
                            else if (columnHeader.ToLower().Contains("diameter"))
                            {
                                sizeColumn = columnCount;
                            }
                        }
                    }
                    else
                    {
                        Product foundProduct = new Product
                        {
                            Name = nameColumn == -1 ? "" : table.Rows[rowCount].ItemArray[nameColumn].ToString().ToLower(),
                            Description = descriptionColumn == -1 ? "" : table.Rows[rowCount].ItemArray[descriptionColumn].ToString().ToLower(),
                            Price = priceColumn == -1 ? "0" : table.Rows[rowCount].ItemArray[priceColumn].ToString().ToLower(),
                            UOM = "cm",
                            Size = sizeColumn == -1 ? 0 : Convert.ToInt32(table.Rows[rowCount].ItemArray[sizeColumn].ToString().ToLower()),
                            Categories = new List<string> { "pizzabodem" }
                        };
                        result.Add(foundProduct);
                    }
                }
            }
            return result.DistinctBy(p => p.Name).ToList();
        }

        private static List<Product> GetIngredients(string pathToFile, bool lookForSauce, string productName = "")
        {
            List<Product> result = new List<Product>();
            List<Product> ProductSpecificResult = new List<Product>();
            using (FileStream stream = File.OpenRead(pathToFile))
            using (var dr = pathToFile.Contains(".csv") ? ExcelReaderFactory.CreateCsvReader(stream) : ExcelReaderFactory.CreateOpenXmlReader(stream))
            {
                int productColumn = -1;
                int nameColumn = -1;
                int amountColumn = -1;
                int sauceColumn = -1;
                int priceColumn = -1;
                DataSet data = dr.AsDataSet();
                var table = data.Tables[0];
                for (int rowCount = 0; rowCount < table.Rows.Count; rowCount++)
                {
                    if (rowCount == 0 && !string.IsNullOrEmpty(productName))
                    {
                        for (int columnCount = 0; columnCount < table.Rows[rowCount].ItemArray.Length; columnCount++)
                        {
                            string columnHeader = table.Rows[rowCount].ItemArray[columnCount].ToString().ToLower();
                            if (columnHeader.ToLower().Contains("ingredientnaam"))
                            {
                                nameColumn = columnCount;
                            }
                            else if (columnHeader.ToLower().Contains("aantalkeer_ingredient"))
                            {
                                amountColumn = columnCount;
                            }
                            else if (columnHeader.ToLower().Contains("productnaam"))
                            {
                                productColumn = columnCount;
                            }
                            else if (columnHeader.ToLower().Contains("pizzasaus_standaard"))
                            {
                                sauceColumn = columnCount;
                            }
                            if (columnHeader.ToLower().Contains("price"))
                            {
                                priceColumn = columnCount;
                            }

                        }
                    }
                    else if (rowCount == 0 && string.IsNullOrEmpty(productName))
                    {
                        for (int columnCount = 0; columnCount < table.Rows[rowCount].ItemArray.Length; columnCount++)
                        {
                            string columnHeader = table.Rows[rowCount].ItemArray[columnCount].ToString().ToLower();
                            if (columnHeader.ToLower().Contains("ingredient"))
                            {
                                nameColumn = columnCount;
                            }
                            if (columnHeader.ToLower().Contains("price"))
                            {
                                priceColumn = columnCount;
                            }
                            if (columnHeader.ToLower().Contains("pizzasaus_standaard"))
                            {
                                sauceColumn = columnCount;
                            }
                        }

                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(productName) && table.Rows[rowCount].ItemArray[productColumn].ToString().ToLower() == productName)
                        {
                            ProductSpecificResult.Add(new Product
                            {
                                Name = table.Rows[rowCount].ItemArray[nameColumn].ToString().ToLower(),
                                amount = Convert.ToInt32(table.Rows[rowCount].ItemArray[amountColumn].ToString().ToLower()),
                                Categories = new List<string> { "ingredient" }
                            });
                            ProductSpecificResult.Add(new Product
                            {
                                Name = table.Rows[rowCount].ItemArray[sauceColumn].ToString().ToLower(),
                                amount = 1,
                                Categories = new List<string> { "ingredient" }
                            });
                        }
                        else
                        {
                            result.Add(new Product
                            {
                                Name = table.Rows[rowCount].ItemArray[nameColumn].ToString().ToLower(),
                                Categories = new List<string> { "ingredient" },
                                Price = priceColumn < 0 ? "" : table.Rows[rowCount].ItemArray[priceColumn].ToString().ToLower()

                            });

                        }
                        if (lookForSauce)
                        {
                            result.Add(new Product
                            {
                                Name = table.Rows[rowCount].ItemArray[sauceColumn].ToString().ToLower(),
                                Categories = new List<string> { "ingredient" }
                            });
                        }
                    }
                }
            }
            if (!string.IsNullOrEmpty(productName))
            {
                return ProductSpecificResult.DistinctBy(p => p.Name).ToList();
            }
            return result.DistinctBy(p => p.Name).ToList();
        }

        private static void WriteCategoriesToDB(List<String> categories)
        {
            var cnts = "Data Source = sql6009.site4now.net; Initial Catalog = DB_A2C9F3_MarioPizza; Persist Security Info = True; User ID = DB_A2C9F3_MarioPizza_admin; Password = Februarie2020!";

            using (SqlConnection cnx = new SqlConnection(cnts))
            {
                cnx.Open();
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = cnx;
                cmd.CommandText = "insert INTO [ProductCategory-QL](Name) VALUES (@name)";
                foreach (string str in categories)
                {
                    cmd.Parameters.Clear();
                    cmd.Parameters.AddWithValue("@name", str);
                    if (cmd.ExecuteNonQuery() > 0)
                    {
                        Console.WriteLine("Categorie written to db with success");
                    }
                }
            }
        }

        private static void WriteProductsToDB(List<Product> products)
        {
            var cnts = "Data Source = sql6009.site4now.net; Initial Catalog = DB_A2C9F3_MarioPizza; Persist Security Info = True; User ID = DB_A2C9F3_MarioPizza_admin; Password = Februarie2020!";
            using (SqlConnection cnx = new SqlConnection(cnts))
            {
                cnx.Open();
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = cnx;
                cmd.CommandText = "insert INTO [Product-QL](Name, Description, Size, [UOM ID], CostPriceID) VALUES (@name, @Description, @Size, @UOM, @Price)";

                foreach (Product str in products)
                {
                    cmd.Parameters.Clear();
                    cmd.Parameters.AddWithValue("@name", str.Name);
                    cmd.Parameters.AddWithValue("@Description", string.IsNullOrEmpty(str.Description) ? "" : str.Description);
                    cmd.Parameters.AddWithValue("@Size", str.Size);
                    cmd.Parameters.AddWithValue("@Price", string.IsNullOrEmpty(str.Price) ? "0" : Regex.Replace(str.Price, "[^0-9,.]", "").Replace('.', ',').Trim());
                    cmd.Parameters.AddWithValue("@UOM", str.Size > 0 ? "CM" : "");
                    if (cmd.ExecuteNonQuery() > 0)
                    {
                        Console.WriteLine("Product written to db with success");
                    }
                }
            }
        }
    }
}
