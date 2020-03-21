using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace CSV_MarioSplitter
{
    class CSVImport
    {
        
        static string sqlConn = "Data Source=sql6009.site4now.net;Initial Catalog=DB_A2C9F3_MarioPizza;Persist Security Info=True;User ID=DB_A2C9F3_MarioPizza_admin;Password=Februarie2020!";

        SqlConnection cnx = new SqlConnection(sqlConn);

        public void importCSV()
        {
            SqlCommand cmdOrderData = new SqlCommand();
            SqlCommand cmdAddress = new SqlCommand();
            SqlCommand cmdCustomer = new SqlCommand();
            SqlCommand cmdOrderLineData = new SqlCommand();
            SqlCommand cmdOrderLineModification = new SqlCommand();

            Boolean isHeader = false;
            int OrderId = 0;
            string Zipcode = "";
            List<string> ExtraIngredients = new List<string>();

            DataTable table = ConvertCSVtoDataTable(@"B:\Downloads\MarioData (1)\MarioOrderData02_10000.csv",true);

            DumpDataTable(table);

            cnx.Open();
            //accessConn.Open();
            cmdOrderData.Connection = cnx;
            cmdAddress.Connection = cnx;
            cmdCustomer.Connection = cnx;
            cmdOrderLineData.Connection = cnx;
            cmdOrderLineModification.Connection = cnx;

            //Import orders
            foreach (DataRow dataRow in table.Rows)
            {
                isHeader = false;
                ExtraIngredients.Clear();

                if (dataRow.ItemArray.GetValue(0) != "")
                {
                    OrderId++;

                    //Split address field into Streetname, house number and house number addition
                    List<string> addressInfo = new List<string>();
                    addressInfo = AdressSplitter(dataRow.ItemArray.GetValue(4).ToString());

                    //Get ZipCode from access database
                    if (addressInfo.Count > 1)
                    {
                        Zipcode = GetZipCode(addressInfo[0], addressInfo[1],dataRow.ItemArray.GetValue(5).ToString());
                    }


                    isHeader = true;

                    //Construct orderheader query
                    cmdOrderData.CommandText = "INSERT INTO [OrderHeader-QL] (" +
                        "ID," +
                        "CustomerID," +
                        "OrderDate," +
                        "StoreID," +
                        "StatusID," +
                        "AddressID," +
                        "CouponID," +
                        "Delivery," +
                        "ZipCode," +
                        "HousNumber," +
                        "HouseNumberAddition," +
                        "Deliverytime) " +
                        "VALUES (" +
                        "@ID," +
                        "@Customer," +
                        "@OrderDate," +
                        "@StoreID," +
                        "@StatusID," +
                        "@AddressID," +
                        "@CouponID," +
                        "@Delivery," +
                        "@ZipCode," +
                        "@HousNumber," +
                        "@HouseNumberAdditon," +
                        "@Deliverytime) ";

                    cmdOrderData.Parameters.AddWithValue("@ID", OrderId);
                    cmdOrderData.Parameters.AddWithValue("@Customer", OrderId); // Import Customer
                    cmdOrderData.Parameters.AddWithValue("@OrderDate", dataRow.ItemArray.GetValue(6));
                    cmdOrderData.Parameters.AddWithValue("@StoreID", dataRow.ItemArray.GetValue(0));
                    cmdOrderData.Parameters.AddWithValue("@StatusID", "");
                    cmdOrderData.Parameters.AddWithValue("@AddressID", OrderId); // Import Address
                    cmdOrderData.Parameters.AddWithValue("@CouponID", dataRow.ItemArray.GetValue(20));

                    if (dataRow.ItemArray.GetValue(7) == "Bezorgen")
                    {
                        cmdOrderData.Parameters.AddWithValue("@Delivery", 1);
                    }
                    else
                    {
                        cmdOrderData.Parameters.AddWithValue("@Delivery", 1);
                    }

                    cmdOrderData.Parameters.AddWithValue("@ZipCode", Zipcode);

                    if (addressInfo.Count > 1)
                    {
                        cmdOrderData.Parameters.AddWithValue("@HousNumber", addressInfo[1]);
                    }
                    else
                    {
                        cmdOrderData.Parameters.AddWithValue("@HousNumber", "");
                    }

                    if (addressInfo.Count > 2)
                    {
                        cmdOrderData.Parameters.AddWithValue("@HouseNumberAdditon", addressInfo[2]);
                    }
                    else
                    {
                        cmdOrderData.Parameters.AddWithValue("@HouseNumberAdditon", "");
                    }
                    cmdOrderData.Parameters.AddWithValue("@Deliverytime", dataRow.ItemArray.GetValue(8) + " " + dataRow.ItemArray.GetValue(9));

                    //Construct orderline query
                    cmdOrderLineData.CommandText = "INSERT INTO [OrderLine-QL] (Quantity,PricePaid,OrderHeaderID,ProductID) VALUES (@LineQuantity,@LinePricePaid,@OrderHeaderID,@ProductID)";
                    cmdOrderLineData.Parameters.AddWithValue("@LineQuantity", dataRow.ItemArray.GetValue(15));
                    cmdOrderLineData.Parameters.AddWithValue("@LinePricePaid", dataRow.ItemArray.GetValue(13));
                    cmdOrderLineData.Parameters.AddWithValue("@OrderHeaderID", OrderId);
                    cmdOrderLineData.Parameters.AddWithValue("@ProductID", dataRow.ItemArray.GetValue(10));

                    //if (dataRow.ItemArray.GetValue(16) != "")
                    //{
                    //    cmdOrderLineModification.CommandText = "INSERT INTO [Order_Line_Modification-QL] (ID,ProductID,Quantity,OrderLineID,UOMD) VALUES (@ID,@ProductID,@Quantity,OrderLineID,UOMID)";
                    //    cmdOrderLineModification.Parameters.AddWithValue("", "");
                    //    cmdOrderLineModification.Parameters.AddWithValue("", "");
                    //    cmdOrderLineModification.Parameters.AddWithValue("", "");
                    //    cmdOrderLineModification.Parameters.AddWithValue("", "");
                    //}

                    //Construct Customer query
                    cmdCustomer.CommandText = "INSERT INTO [Customer-QL] (ID,Name,Email,Phonenumber ) VALUES(@CustID,@CustName,@CustEmail,@CustPhoneNumber)";
                    cmdCustomer.Parameters.AddWithValue("@CustID", OrderId);
                    cmdCustomer.Parameters.AddWithValue("@CustName", dataRow.ItemArray.GetValue(1));
                    cmdCustomer.Parameters.AddWithValue("@CustEmail", dataRow.ItemArray.GetValue(3));
                    cmdCustomer.Parameters.AddWithValue("@CustPhoneNumber", dataRow.ItemArray.GetValue(2));

                    //Construct address query
                    cmdAddress.CommandText = "INSERT INTO [Address-QL] (ID,Zipcode,Housenumber,HouseNumberAddition,Streetname,City) VALUES(@AddressID,@ZipCode,@AddressHouseNumber,@AddressHouseNumberAdditon,@AddressStreetname,@AddressCity)";
                    cmdAddress.Parameters.AddWithValue("@AddressID", OrderId);
                    cmdAddress.Parameters.AddWithValue("@ZipCode", Zipcode);
                    if (addressInfo.Count > 0)
                    {
                        cmdAddress.Parameters.AddWithValue("@AddressHouseNumber", addressInfo[1]);
                    }
                    else
                    {
                        cmdAddress.Parameters.AddWithValue("@AddressHouseNumber", "");
                    }
                    if (addressInfo.Count > 2)
                    {
                        cmdAddress.Parameters.AddWithValue("@AddressHouseNumberAdditon", addressInfo[2]);
                    }
                    else
                    {
                        cmdAddress.Parameters.AddWithValue("@AddressHouseNumberAdditon", "");
                    }
                    if (addressInfo.Count > 1)
                    {
                        cmdAddress.Parameters.AddWithValue("@AddressStreetname", addressInfo[0]);
                    }
                    else
                    {
                        cmdAddress.Parameters.AddWithValue("@AddressStreetname", "");
                    }
                    cmdAddress.Parameters.AddWithValue("@AddressCity", dataRow.ItemArray.GetValue(5));

                }
                else
                {
                    isHeader = false;
                    //Construct orderline query
                    cmdOrderLineData.CommandText = "INSERT INTO [OrderLine-QL] (Quantity,PricePaid,OrderHeaderID,ProductID) VALUES (@LineQuantity,@LinePricePaid,@OrderHeaderID,@ProductID)";
                    cmdOrderLineData.Parameters.AddWithValue("@LineQuantity", dataRow.ItemArray.GetValue(15));
                    cmdOrderLineData.Parameters.AddWithValue("@LinePricePaid", dataRow.ItemArray.GetValue(13));
                    cmdOrderLineData.Parameters.AddWithValue("@OrderHeaderID", OrderId);
                    cmdOrderLineData.Parameters.AddWithValue("@ProductID", dataRow.ItemArray.GetValue(10));
                }

                try
                {
                    if (isHeader)
                    {
                        //Execute orderheader, customer and address queries
                        cmdOrderData.ExecuteNonQuery();
                        cmdCustomer.ExecuteNonQuery();
                        cmdAddress.ExecuteNonQuery();
                        cmdOrderLineData.ExecuteNonQuery();
                    }
                    else
                    {
                        //Execute orderline query
                        cmdOrderLineData.ExecuteNonQuery();
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                    break;
                }
                finally
                {
                    //Clear paramaters for the next row
                    cmdOrderData.Parameters.Clear();
                    cmdCustomer.Parameters.Clear();
                    cmdAddress.Parameters.Clear();
                    cmdOrderLineData.Parameters.Clear();
                }
            }
            cnx.Close();
            //accessConn.Close();
        }


        public List<string> AdressSplitter(string adress)
        {
            List<string> addressInfo = new List<string>();

            string[] numbers = Regex.Split(adress, @"^(.+)\s(\d+(\s*[^\d\s]+)*)$");
            foreach (string value in numbers)
            {
                if (!string.IsNullOrEmpty(value))
                {
                    addressInfo.Add(value.Trim());
                }
            }
            return addressInfo;
        }
        public string GetZipCode(string streetName, string houseNumber,string city)
        {
            string result = "";

            OleDbConnection accessConn = new OleDbConnection("Provider = Microsoft.Jet.OLEDB.4.0; Data Source = C:\\Postcode tabel.mdb");
            
            OleDbCommand cmd = new OleDbCommand
            {
                Connection = accessConn,
                CommandType = CommandType.Text,
                CommandText = "SELECT TOP 1 A13_POSTCODE FROM POSTCODES WHERE A13_WOONPLAATS = @City AND A13_STRAATNAAM = @StreetName AND  @HouseNumber BETWEEN A13_BREEKPUNT_VAN AND A13_BREEKPUNT_TEM"
            };

            cmd.Parameters.AddWithValue("@StreetName", streetName);
            cmd.Parameters.AddWithValue("@City", city.ToUpper());
            if (int.TryParse(houseNumber, out _))
            {
                cmd.Parameters.AddWithValue("@HouseNumber", houseNumber);
            }
            try
            {
                accessConn.Open();
                OleDbDataReader reader = cmd.ExecuteReader();
                if (!reader.HasRows)
                {
                    
                    result = "";
                }
                else
                {
                    while (reader.Read())
                    {
                        result = reader.GetString(0);
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            finally
            {
                cmd.Parameters.Clear();
                accessConn.Close();
            }

            return result;
        }
        public void importIngredients()
        {
            int CostPriceID = 0;
            SqlCommand cmdProducts = new SqlCommand();
            SqlCommand cmdPrice = new SqlCommand();

            DataTable productTable = ConvertCSVtoDataTable(@"B:\Downloads\MarioData (1)\Extra Ingredienten.csv",false);

            cnx.Open();

            cmdProducts.Connection = cnx;
            cmdPrice.Connection = cnx;

            foreach (DataRow dataRow in productTable.Rows)
            {
                CostPriceID++;
                cmdProducts.CommandText = "INSERT INTO [Product-QL](Name,ProductCategoryID,CostPriceID) VALUES (@Name, @ProductCategoryID, @CostPriceID)";
                cmdProducts.Parameters.AddWithValue("@Name", dataRow.ItemArray.GetValue(0));
                cmdProducts.Parameters.AddWithValue("@ProductCategoryID", "Ingredient");
                cmdProducts.Parameters.AddWithValue("@CostPriceID", CostPriceID);

                cmdPrice.CommandText = "INSERT INTO [CostPrice-QL] (ID,Amount) VALUES (@ID,@Amount)";
                cmdPrice.Parameters.AddWithValue("@ID", CostPriceID);
                cmdPrice.Parameters.AddWithValue("@Amount", dataRow.ItemArray.GetValue(1));

                try
                {
                    cmdProducts.ExecuteNonQuery();
                    cmdPrice.ExecuteNonQuery();
                }
                catch(Exception e)
                {
                    Console.WriteLine(e);
                    break;
                }
                finally
                {
                    cmdProducts.Parameters.Clear();
                    cmdPrice.Parameters.Clear();
                }
            }
            cnx.Close();
        }

        public DataTable ConvertCSVtoDataTable(string strFilePath,bool needSkipLine)
        {
            DataTable dt = new DataTable();
            using (StreamReader sr = new StreamReader(strFilePath))
            {
                if (needSkipLine)
                {
                    sr.ReadLine();
                    sr.ReadLine();
                    sr.ReadLine();
                    sr.ReadLine();
                }

                string[] headers = sr.ReadLine().Split(';');
                foreach (string header in headers)
                {
                    dt.Columns.Add(header);
                }
                while (!sr.EndOfStream)
                {
                    string[] rows = sr.ReadLine().Split(';');
                    DataRow dr = dt.NewRow();

                    for (int i = 0; i < headers.Length; i++)
                    {

                        if (rows.Length != 1)
                        {
                            dr[i] = rows[i];
                        }
                    }
                    if (!CheckRowEmpty(dr, headers.Length))
                    {
                        dt.Rows.Add(dr);
                    }
                }

            }

            return dt;
        }
        public  Boolean CheckRowEmpty(DataRow LocalDataRow, int Columns)
        {
            Boolean Empty = true;
            for (int i = 0; i < Columns; i++)
            {
                if (LocalDataRow[i] != "")
                {
                    Empty = false;
                }
                if (Empty = false)
                {
                    return Empty;
                }

            }
            return Empty;
        }
        public string DumpDataTable(DataTable table)
        {
            string data = string.Empty;
            StringBuilder sb = new StringBuilder();

            if (null != table && null != table.Rows)
            {
                foreach (DataRow dataRow in table.Rows)
                {
                    foreach (var item in dataRow.ItemArray)
                    {
                        sb.Append(item);
                        sb.Append(',');
                    }
                    sb.AppendLine();
                }

                data = sb.ToString();
            }
            return data;
        }


    }

}
