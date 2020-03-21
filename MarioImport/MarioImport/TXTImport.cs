using System;
using System.IO;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Data.SqlClient;
using Newtonsoft.Json;

namespace MarioImport
{
    class TXTImport
    {
        private string path;
        string[] lines;
        List<Store> stores = new List<Store>();
        Dictionary<Store, string> storeDictoinary = new Dictionary<Store, string>();
        public TXTImport(string basePath)
        {
            path = basePath;
            lines = File.ReadAllLines(path + @"\Winkels Mario.txt");
        }

        public void textImport()
        {

            Console.WriteLine("contents of the stores.txt = ");

            for (int i = 3; i < lines.Length; i++)
            {
                String Name = lines[i];
                i++;
                String Street = lines[i];
                i++;
                String HomeNumber = lines[i];
                i++;
                String City = lines[i];
                i++;
                String Country = lines[i];
                i++;
                String ZipCode = lines[i];
                i++;
                String PhoneNumber = lines[i];

                string homeNumberSuffix = "";
                int homeNumberNumber = 0;
                string[] numbers = Regex.Split(HomeNumber, @"\D+");
                foreach (string value in numbers)
                {
                    if (!string.IsNullOrEmpty(value))
                    {
                        int hn = int.Parse(value);
                        homeNumberNumber = hn;
                        homeNumberSuffix = HomeNumber.Replace(value, "").Trim();
                    }
                }



                Store tempstore = new Store(Name, Street, homeNumberNumber.ToString(), homeNumberSuffix, City, Country, ZipCode, PhoneNumber);
                if (IsNLZipCode(ZipCode) && IsNLPhoennumber(PhoneNumber)) stores.Add(tempstore);
                else
                {
                    if (!IsNLZipCode(ZipCode))
                    {
                        storeDictoinary.Add(tempstore, ": Zip code incorrect.");
                    }
                    else if (!IsNLPhoennumber(PhoneNumber))
                    {
                        storeDictoinary.Add(tempstore, ": Phonenumber incorrect.");
                    }
                    else
                    {
                        storeDictoinary.Add(tempstore, ": Error unknow.");
                    }


                }
                i++;
            }

            Console.WriteLine("all the valid stores:");
            foreach (Store s in stores)
            {
                Console.WriteLine(s.ToString());
            }
            Console.WriteLine("all the nonvalid stores:");
            File.Delete(path + @"\WriteLines2.txt"); File.Delete(path + @"\WriteLines2.txt");
            using (System.IO.StreamWriter file =
            new System.IO.StreamWriter(path + @"\WriteLines2.txt"))
            {
                file.Write(JsonConvert.SerializeObject(storeDictoinary, Formatting.Indented));
            }
            foreach (KeyValuePair<Store, string> kvp in storeDictoinary)
            {
                Console.WriteLine("Store: {0} - Reason: {1}", kvp.Key, kvp.Value);
            }

        }

        private bool IsNLZipCode(string zipCode)
        {
            var _NLZipRegEx = @"[1-9][0-9]{3}?[A-Z]{2}$";

            var validZipCode = true;
            if (!Regex.Match(zipCode, _NLZipRegEx).Success) validZipCode = false;
            return validZipCode;
        }

        private bool IsNLPhoennumber(string PhoneNumber)
        {
            var _NLPhoneNumberEX = @"[0-9]{10}$";

            var validPhoneNumber = true;
            if (!Regex.Match(PhoneNumber, _NLPhoneNumberEX).Success) validPhoneNumber = false;
            return validPhoneNumber;
        }

        public void databasewrite()
        {
            var cnts = "Data Source = sql6009.site4now.net; Initial Catalog = DB_A2C9F3_MarioPizza; Persist Security Info = True; User ID = DB_A2C9F3_MarioPizza_admin; Password = Februarie2020!";
            SqlConnection cnx = new SqlConnection(cnts);
            SqlCommand cmd = new SqlCommand();
            SqlCommand cmd2 = new SqlCommand();

            cnx.Open();
            cmd.Connection = cnx;
            cmd2.Connection = cnx;
            Console.Clear();

            

            foreach (Store store in stores)
            {

                cmd.CommandText = "insert INTO [Store-QL](Name, Description, PhoneNumber) VALUES (@name, @discription, @phonenumber)";
                cmd.Parameters.AddWithValue("@name", store.name);
                cmd.Parameters.AddWithValue("@discription", store.name);
                cmd.Parameters.AddWithValue("@phonenumber", store.phoneNumber);
                cmd2.CommandText = "insert INTO [Address-QL](ZipCode, HouseNumber, HouseNumberAddition, Streetname, City, State, CountryID) VALUES (@ZipCode, @HouseNumber, @HouseNumberAddition, @Streetname, @City, @State, @Country)";
                cmd2.Parameters.AddWithValue("@ZipCode", store.zipCode);
                cmd2.Parameters.AddWithValue("@HouseNumber", store.homeNumber);
                cmd2.Parameters.AddWithValue("@HouseNumberAddition", store.homeNumberSuffix);
                cmd2.Parameters.AddWithValue("@Streetname", store.street);
                cmd2.Parameters.AddWithValue("@City", store.city);
                cmd2.Parameters.AddWithValue("@State", "NULL");
                cmd2.Parameters.AddWithValue("@Country", store.country);
                try
                {
                    cmd.ExecuteNonQuery();
                    cmd2.ExecuteNonQuery();
                }
                catch (InvalidCastException e)
                {
                    Console.WriteLine(e);
                }
                finally
                {
                    cmd.Parameters.Clear();
                    cmd2.Parameters.Clear();
                }
            }
            cnx.Close();
        }
    }
}
