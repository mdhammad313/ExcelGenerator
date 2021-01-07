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

        [ExcelGenerator(order: 3)]
        public string Address { get; set; }

        [ExcelGenerator]
        public string City { get; set; }

        public string County { get; set; }

        [ExcelGenerator("Email Address")]
        public string Email { get; set; }

        [ExcelGenerator("Birth Date",dateFormat:Constants.DateFormat)]
        public DateTime DateOfBirth { get; set; }
    }
}
