# ExcelGenerator
ExcelGenerator is a simple library to generate excel easily more than ever by just passing the list of records irrespective of model type, column mapping and data bindings. It will also take care of column name, column orders and date formats automatically or just define column name, order and date formats using attributes. It also provides multiple overloads for providing generating excel on specific location or having stream.

## Installation
ExcelGenerator is available as a NuGet Package. Type the following command into NuGet Package Manager Console window to install it:

```
PM> Install-Package ExcelGenerator.net
```

## Getting Started
Apply excel generator attributes to the properties that needs to be included in excel having multiple overloads for customizing name order and date formats. 
Here `County` property will not be added in excel file as it doesn't have excel generator attribute

```csharp
    public class Model
    {
        [ExcelGenerator("First Name")]
        public string FirstName { get; set; }

        [ExcelGenerator("Birth Date",dateFormat:"MMM dd, yyyy")]
        public DateTime DateOfBirth { get; set; }

        [ExcelGenerator]
        public string City { get; set; }

        public string County { get; set; }

        [ExcelGenerator("Email Address")]
        public string Email { get; set; }

        [ExcelGenerator(order: 3)]
        public string Address { get; set; }

    }
```

Now load sample data  
```csharp
    //Get sample data to generate excel 
    var data = GetSampleData();
            
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
```

Now generate excel using stream or generate excel on specific path
```csharp
            //Method 1: Get excel file stream 
            var responseStream = ExcelHelper.GenerateExcel(data);

            //Method 2: Generate excel file on specified path 
            var currentDirectory = Directory.GetCurrentDirectory();
            ExcelHelper.GenerateExcel($"{currentDirectory}\\SampleFile.xlsx", data);

```

You can also customize sheet name 
```csharp

   ExcelHelper.GenerateExcel($"{currentDirectory}\\SampleFile.xlsx", data,"TestSheet");
```


## Sample Project 
Sample project have also been added, named `ExcelGenerator.Sample` which consumes library with all possible methods  

## Nuget
Also available on Nuget 

[![Version](https://img.shields.io/nuget/vpre/ExcelGenerator.net.svg)](https://www.nuget.org/packages/ExcelGenerator.net)
