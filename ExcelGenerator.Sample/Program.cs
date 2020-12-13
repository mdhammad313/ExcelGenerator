using ExcelGenerator.Library;
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

            //Get path of the current directory where excel file will be generated
            var currentDirectory = Directory.GetCurrentDirectory();


            //Method1: Get excel file in memory stream
            var responseStream = ExcelHelper.GenerateExcel(data);
            var fileStream = new FileStream($"{currentDirectory}\\SampleFileWithStream.xlsx", FileMode.Create, FileAccess.Write);
            responseStream.CopyTo(fileStream);
            fileStream.Dispose();

            //Method2: Generate excel file on specified path 
            //ExcelHelper.GenerateExcel($"{currentDirectory}\\SampleFileWithFilePath.xlsx",data);

            Console.WriteLine("Excel file generated");
            Console.ReadKey();
        }


        public static List<Model> GetSampleData()
        {
            return new List<Model>
            {
                    new Model { FirstName = "James",  LastName = "Butt",  Address="6649 N Blue Gum St", City="New Orleans" , County="Orleans",  State="LA" , Zip = 70116, Phone="504-621-8927" , Email="jbutt@gmail.com", DateOfBirth=DateTime.Now},
                    new Model { FirstName = "Josephine",  LastName = "Darakjy",  Address="4 B Blue Ridge Blvd", City="Brighton" , County="Livingston",  State="MI" , Zip = 48116, Phone="810-292-9388" , Email="josephine_darakjy@darakjy.org", DateOfBirth=DateTime.Now},
                    new Model { FirstName = "Art",  LastName = "Venere",  Address="8 W Cerritos Ave #54", City="Bridgeport" , County="Gloucester",  State="NJ" , Zip = 8014, Phone="856-636-8749" , Email="art@venere.org", DateOfBirth=DateTime.Now},
                    new Model { FirstName = "Lenna",  LastName = "Paprocki",  Address="639 Main St", City="Anchorage" , County="Anchorage",  State="AK" , Zip = 99501, Phone="907-385-4412" , Email="lpaprocki@hotmail.com", DateOfBirth=DateTime.Now},
                    new Model { FirstName = "Donette",  LastName = "Foller",  Address="34 Center St", City="Hamilton" , County="Butler",  State="OH" , Zip = 45011, Phone="513-570-1893" , Email="donette.foller@cox.net", DateOfBirth=DateTime.Now},
                    new Model { FirstName = "Simona",  LastName = "Morasca",  Address="3 Mcauley Dr", City="Ashland" , County="Ashland",  State="OH" , Zip = 44805, Phone="419-503-2484" , Email="simona@morasca.com", DateOfBirth=DateTime.Now},
                    new Model { FirstName = "Mitsue",  LastName = "Tollner",  Address="7 Eads St", City="Chicago" , County="Cook",  State="IL" , Zip = 60632, Phone="773-573-6914" , Email="mitsue_tollner@yahoo.com", DateOfBirth=DateTime.Now},
                    new Model { FirstName = "Leota",  LastName = "Dilliard",  Address="7 W Jackson Blvd", City="San Jose" , County="Santa Clara",  State="CA" , Zip = 95111, Phone="408-752-3500" , Email="leota@hotmail.com", DateOfBirth=DateTime.Now},
                    new Model { FirstName = "Sage",  LastName = "Wieser",  Address="5 Boston Ave #88", City="Sioux Falls" , County="Minnehaha",  State="SD" , Zip = 57105, Phone="605-414-2147" , Email="sage_wieser@cox.net", DateOfBirth=DateTime.Now},
                    new Model { FirstName = "Kris",  LastName = "Marrier",  Address="228 Runamuck Pl #2808", City="Baltimore" , County="Baltimore City",  State="MD" , Zip = 21224, Phone="410-655-8723" , Email="kris@gmail.com", DateOfBirth=DateTime.Now},
            };
        }

    }
}
