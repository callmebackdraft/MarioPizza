using System;
using System.Data;
using System.IO;
using System.Text;

namespace CSV_MarioSplitter
{
    class Program
    {
        static void Main(string[] args)
        {

            DataTable table;
            table = ConvertCSVtoDataTable(@"B:\Downloads\MarioData (1)\MarioOrderData01_10000.csv");

            //foreach (DataRow dataRow in table.Rows)
            //{
            //    foreach (var item in dataRow.ItemArray)
            //    {
            //        Console.WriteLine(item);
            //    }
            //}

           Console.Write(DumpDataTable(table));

        }

        public static DataTable ConvertCSVtoDataTable(string strFilePath)
        {
            DataTable dt = new DataTable();
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
                        dt.Rows.Add(dr);
                }

            }


            return dt;
        }
        public static string DumpDataTable(DataTable table)
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
