using ExcelGenerator.Sample.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Hosting;

namespace ExcelGenerator.Sample
{
    class Program
    {
        static void Main(string[] args)
        {
            //Get sample data to generate excel 
            var data = GetSampleData();

            //Method 1: Get excel file stream 
            var responseStream = ExcelHelper.GenerateExcel(data);

            //Method 2: Generate excel file on specified path 
            var currentDirectory = Directory.GetCurrentDirectory();
            ExcelHelper.GenerateExcel($"{currentDirectory}\\SampleFile.xlsx", data);

            Console.WriteLine("Excel file generated");
            Console.ReadKey();
        }


        public static List<Model> GetSampleData()
        {
            return new List<Model>
            {
                    new Model { FirstName = "James", Address="6649 N Blue Gum St", City="New Orleans" , County="Orleans",  Email="jbutt@gmail.com", DateOfBirth=DateTime.Now},
                    new Model { FirstName = "Josephine",  Address="4 B Blue Ridge Blvd", City="Brighton" , County="Livingston",  Email="josephine_darakjy@darakjy.org", DateOfBirth=DateTime.Now},
                    new Model { FirstName = "Art",  Address="8 W Cerritos Ave #54", City="Bridgeport" , County="Gloucester", Email="art@venere.org", DateOfBirth=DateTime.Now},
                    new Model { FirstName = "Lenna",  Address="639 Main St", City="Anchorage" , County="Anchorage", Email="lpaprocki@hotmail.com", DateOfBirth=DateTime.Now},
                    new Model { FirstName = "Donette", Address="34 Center St", City="Hamilton" , County="Butler", Email="donette.foller@cox.net", DateOfBirth=DateTime.Now},
                    new Model { FirstName = "Simona", Address="3 Mcauley Dr", City="Ashland" , County="Ashland", Email="simona@morasca.com", DateOfBirth=DateTime.Now},
                    new Model { FirstName = "Mitsue", Address="7 Eads St", City="Chicago" , County="Cook", Email="mitsue_tollner@yahoo.com", DateOfBirth=DateTime.Now},
                    new Model { FirstName = "Leota",  Address="7 W Jackson Blvd", City="San Jose" , County="Santa Clara", Email="leota@hotmail.com", DateOfBirth=DateTime.Now},
                    new Model { FirstName = "Sage", Address="5 Boston Ave #88", City="Sioux Falls" , County="Minnehaha", Email="sage_wieser@cox.net", DateOfBirth=DateTime.Now},
                    new Model { FirstName = "Kris",  Address="228 Runamuck Pl #2808", City="Baltimore" , County="Baltimore City", Email="kris@gmail.com", DateOfBirth=DateTime.Now},
            };
        }

    }
}
