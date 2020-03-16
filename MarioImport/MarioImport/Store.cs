using System;
using System.Collections.Generic;
using System.Text;

namespace MarioImport
{
    class Store
    {
        string Name;
        string Street;
        string HomeNumber;
        string HomeNumberSuffix;
        string City;
        string Country;
        string ZipCode;
        string PhoneNumber;

        public Store(string name, string street, string homeNumber, string homeNumberSuffix, string city, string country, string zipCode, string phoneNumber)
        {
            Name = name;
            Street = street;
            HomeNumber = homeNumber;
            HomeNumberSuffix = homeNumberSuffix;
            City = city;
            Country = country;
            ZipCode = zipCode;
            PhoneNumber = phoneNumber;
        }

        public string name
        {
            get { return Name; } // get method
            set { Name = value; } // set method
        }

        public string street
        {
            get { return Street; } // get method
            set { Street = value; } // set method
        }
        public string homeNumber
        {
            get { return HomeNumber; } // get method
            set { HomeNumber = value; } // set method
        }
        public string homeNumberSuffix
        {
            get { return HomeNumberSuffix; } // get method
            set { HomeNumberSuffix = value; } // set method
        }
        public string city
        {
            get { return City; } // get method
            set { City = value; } // set method
        }
        public string country
        {
            get { return Country; } // get method
            set { Country = value; } // set method
        }
        public string zipCode
        {
            get { return ZipCode; } // get method
            set { ZipCode = value; } // set method
        }

        public string phoneNumber
        {
            get { return PhoneNumber; } // get method
            set { PhoneNumber = value; } // set method
        }

        public override string ToString()
        {
            return String.Format("name: {0} Street: {1} Home number: {2} Home number suffix: {3} City: {4} Country: {5} Zip code: {6} Phone number: {7}",
                Name,
                Street,
                HomeNumber,
                HomeNumberSuffix,
                City, 
                Country,
                ZipCode,
                PhoneNumber);
        }
    }
}
