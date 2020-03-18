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
            CSVImport csvImport = new CSVImport();

            csvImport.importCSV();
            csvImport.importIngredients();

        }
    }
}
