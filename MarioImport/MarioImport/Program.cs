using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using ExcelDataReader;

namespace MarioImport
{
    class Program
    {
        static void Main(string[] args)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            string basePath = @"C:\Users\tonyw\source\repos\MarioImport\Data";
            foreach (string str in GetCategories(basePath))
            {
                Console.WriteLine(str);
            }

        }
        private void test()
        {
            TXTImport import = new TXTImport();
            import.textImport();
            import.databasewrite();
        }
        private static List<string> GetCategories(string basePath)
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
                    int categoryColumn = 0;
                    int subCatColumn = 0;
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
                            }
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
            return result.Select(t => t.Trim()).Distinct().ToList();
        }

        private List<Product> GetProducts(string basePath)
        {
            string[] files = { @"\pizza_ingredienten.xlsx", @"\Overige producten.xlsx", @"\pizzabodems.xlsx" };
            List<Product> result = new List<Product>();
            foreach (string file in files)
            {
                using (FileStream stream = File.OpenRead(basePath + file))
                using (IExcelDataReader dr = ExcelReaderFactory.CreateOpenXmlReader(stream))
                {
                    List<int> ColumnsToRead = new List<int>();
                    DataSet data = dr.AsDataSet();
                    var table = data.Tables[0];
                    for (int rowCount = 0; rowCount < table.Rows.Count; rowCount++)
                    {
                        if (rowCount == 0)
                        {
                            for (int columnCount = 0; columnCount < table.Rows[rowCount].ItemArray.Length; columnCount++)
                            {
                                string columndHeader = table.Rows[rowCount].ItemArray[columnCount].ToString().ToLower();
                                if (columndHeader.Contains("naam") || columndHeader.Contains("pizzasaus_standaard"))
                                {
                                    ColumnsToRead.Add(columnCount);
                                }
                            }
                        }
                        else
                        {
                            foreach(int colIndex in ColumnsToRead)
                            {
                                result.Add(new Product
                                {
                                    Name = table.Rows[rowCount].ItemArray[colIndex].ToString().ToLower(),
                                });
                            }
                            
                        }
                    }
                }
            }
            return result;
        }
    }
}
