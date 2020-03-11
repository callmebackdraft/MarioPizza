using System;
using System.IO;
using ExcelDataReader;

namespace MarioImport
{
    class Program
    {
        static void Main(string[] args)
        {
            string path = "";
            using (FileStream stream = File.OpenRead(path))
            using (IExcelDataReader dr = ExcelReaderFactory.CreateReader(stream))
            {
                while (dr.Read())
                {

                }
            }
        }
    }
}
