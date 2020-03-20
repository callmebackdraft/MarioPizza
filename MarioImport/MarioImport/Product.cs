-using System;
using System.Collections.Generic;
using System.Text;

namespace MarioImport
{
    class Product
    {
        public string Name;
        public int amount;
        public string Description;
        public List<string> Categories;
        public List<Product> Ingredients;
        public decimal Vat;
        public decimal Size;
        public string UOM;
        public string Price;
        public bool Spicy;
        public bool Vegetarisch;
        public string AdditionalCharge;
    }
}
