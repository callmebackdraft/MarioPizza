using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Data.OleDb;


namespace CSV_MarioSplitter
{
    class Program
    {
        static void Main(string[] args)
        {

            string connString = "Data Source=sql6009.site4now.net;Initial Catalog=DB_A2C9F3_MarioPizza;Persist Security Info=True;User ID=DB_A2C9F3_MarioPizza_admin;Password=Februarie2020!";
            

            SqlConnection cnx = new SqlConnection(connString);
            SqlCommand cmdOrderData = new SqlCommand();
            SqlCommand cmdAddress = new SqlCommand();
            SqlCommand cmdCustomer = new SqlCommand();

            Boolean isHeader = false;
            int OrderId = 1;
            int CustomerID = 1;
            int AdressID = 1;
            string Zipcode = "";

            System.Data.DataTable table = ConvertCSVtoDataTable(@"B:\Downloads\MarioData (1)\MarioOrderData01_10000.csv");

            cnx.Open();
            cmdOrderData.Connection = cnx;
            cmdAddress.Connection = cnx;
            cmdCustomer.Connection = cnx;

            //Import orders
            foreach (DataRow dataRow in table.Rows)
            {
                isHeader = false;

                if (dataRow.ItemArray.GetValue(0) != "")
                {
                    //Split address field into Streetname, house number and house number addition
                    List<string> addressInfo = new List<string>();
                    addressInfo = AdressSplitter(dataRow.ItemArray.GetValue(4).ToString());

                    //Get ZipCode from access database
                    if (addressInfo.Count > 1)
                    {
                        Zipcode  = GetZipCode(addressInfo[0], addressInfo[1]);
                    }

                    if (addressInfo == null)
                    {
                        Console.WriteLine("Het address splitten is niet gelukt");
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

                    cmdOrderData.Parameters.AddWithValue("@ID", 10);
                    cmdOrderData.Parameters.AddWithValue("@Customer", CustomerID); // Import Customer
                    cmdOrderData.Parameters.AddWithValue("@OrderDate", dataRow.ItemArray.GetValue(6));
                    cmdOrderData.Parameters.AddWithValue("@StoreID", dataRow.ItemArray.GetValue(0));
                    cmdOrderData.Parameters.AddWithValue("@StatusID", "");
                    cmdOrderData.Parameters.AddWithValue("@AddressID", AdressID); // Import Address
                    cmdOrderData.Parameters.AddWithValue("@CouponID", dataRow.ItemArray.GetValue(18));
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

                    //Construct Customer query
                    cmdCustomer.CommandText = "INSERT INTO [Customer-QL] (ID,Name,Email,Phonenumber ) VALUES(@CustID,@CustName,@CustEmail,@CustPhoneNumber)";
                    cmdCustomer.Parameters.AddWithValue("@CustID", CustomerID);
                    cmdCustomer.Parameters.AddWithValue("@CustName", dataRow.ItemArray.GetValue(1));
                    cmdCustomer.Parameters.AddWithValue("@CustEmail", dataRow.ItemArray.GetValue(4));
                    cmdCustomer.Parameters.AddWithValue("@CustPhoneNumber", dataRow.ItemArray.GetValue(2));

                    //Construct address query
                    cmdAddress.CommandText = "INSERT INTO [Address-QL] (ID,Zipcode,Housenumber,HouseNumberAddition,Streetname,City) VALUES(@AddressID,@ZipCode,@AddressHouseNumber,@AddressHouseNumberAdditon,@AddressStreetname,@AddressCity)";
                    cmdAddress.Parameters.AddWithValue("@AddressID", AdressID);
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
                    cmdOrderData.CommandText = "INSERT INTO [OrderLine-QL] (Quantity,PricePaid,OrderHeaderID,ProductID) VALUES (@LineQuantity,@LinePricePaid,@OrderHeaderID,@ProductID)";
                    cmdOrderData.Parameters.AddWithValue("@LineQuantity", dataRow.ItemArray.GetValue(17));
                    cmdOrderData.Parameters.AddWithValue("@LinePricePaid", dataRow.ItemArray.GetValue(15));
                    cmdOrderData.Parameters.AddWithValue("@OrderHeaderID", OrderId);
                    cmdOrderData.Parameters.AddWithValue("@ProductID", dataRow.ItemArray.GetValue(12));
                }

                try
                {
                    if (isHeader)
                    {
                        //Execute orderheader, customer and address queries
                        cmdOrderData.ExecuteNonQuery();
                        cmdCustomer.ExecuteNonQuery();
                        cmdAddress.ExecuteNonQuery();
                    }
                    else
                    {
                        //Execute orderline query
                        cmdOrderData.ExecuteNonQuery();
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
                    OrderId++;
                    CustomerID++;
                    AdressID++;
                }
            }
            cnx.Close();
        }

        public static List<string> AdressSplitter(string adress)
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
        public static string GetZipCode(string streetName, string houseNumber)
        {
            string result = "";
            OleDbConnection conn = new OleDbConnection("Provider = Microsoft.Jet.OLEDB.4.0; Data Source = C:\\Postcode tabel.mdb");

            OleDbCommand cmd = new OleDbCommand
            {
                Connection = conn,
                CommandType = CommandType.Text,
                CommandText = "SELECT [A13_POSTCODE] FROM [POSTCODES] WHERE " +
                    "[A13_STRAATNAAM] = @StreetName AND " +
                    "[A13_BREEKPUNT_VAN] <= @HouseNumber AND " +
                    "[A13_BREEKPUNT_TEM] >= @HouseNumber "
            };
            cmd.Parameters.AddWithValue("@StreetName", streetName);
            if (int.TryParse(houseNumber, out _))
            {
                cmd.Parameters.AddWithValue("@HouseNumber", houseNumber); 
            }
            try
            {
                conn.Open();
                OleDbDataReader reader = cmd.ExecuteReader();
                result = reader.GetString(0);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            finally
            {
                conn.Close();
            }

            return result;
        }

        public static System.Data.DataTable ConvertCSVtoDataTable(string strFilePath)
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            using (StreamReader sr = new StreamReader(strFilePath))
            {
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
                    if (CheckRowEmpty(dr,headers.Length) )
                    {
                        dt.Rows.Add(dr);
                    }
                }

            }

            return dt;
        }
        public static Boolean CheckRowEmpty(DataRow LocalDataRow, int Columns)
        {
            Boolean Empty = false;
            for (int i = 0; i < Columns; i++)
            {
                if (LocalDataRow[i] == "")
                {
                    Empty = true;
                }

            }
            return Empty;
        }
        public static string DumpDataTable(System.Data.DataTable table)
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
