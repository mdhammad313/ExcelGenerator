using ExcelGenerator.Atrributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelGenerator.Sample.Models
{
    public class Model
    {
        [ExcelGenerator("First Name")]
        public string FirstName { get; set; }

        [ExcelGenerator("Last Name")]
        public string LastName { get; set; }


        [ExcelGenerator("Birth Date",dateFormat:Constants.DateFormat)]
        public DateTime DateOfBirth { get; set; }

        [ExcelGenerator]
        public string City { get; set; }

        public string County { get; set; }

        [ExcelGenerator]
        public string State { get; set; }

        [ExcelGenerator]
        public int Zip { get; set; }

        [ExcelGenerator("Phone Number")]
        public string Phone { get; set; }

        [ExcelGenerator("Email Address")]
        public string Email { get; set; }

        [ExcelGenerator(order: 3)]
        public string Address { get; set; }


    }
}
