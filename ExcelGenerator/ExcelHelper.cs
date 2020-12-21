using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using ExcelGenerator.Atrributes;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExcelGenerator
{
    public class ExcelHelper
    {
        public static void GenerateExcel<T>(string filePath, IEnumerable<T> records) where T : class
        {
            using (var package = new ExcelPackage())
            {
                GeneratePackage(package, records);
                package.SaveAs(new FileInfo(filePath));
            }
        }

        public static MemoryStream GenerateExcel<T>(IEnumerable<T> records) where T : class
        {
            var stream = new MemoryStream();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(stream))
            {
                GeneratePackage(package, records);
                package.Save();
                stream.Position = 0;
            }

            return stream;
        }

        private static void GeneratePackage<T>(ExcelPackage package, IEnumerable<T> records) where T : class
        {
            var worksheetsName =  AppResource.ExcelSheetName;
            var workSheet = package.Workbook.Worksheets.Add(worksheetsName);

            workSheet.Row(1).Height = double.Parse(AppResource.ExcelRowHeight);
            workSheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Row(1).Style.Font.Bold = true;

            var propertyList = typeof(T).GetProperties().Where(x => x.GetCustomAttributes(typeof(ExcelGeneratorAttribute), false).Any())
                                                        .OrderBy(x => (x.GetCustomAttributes(typeof(ExcelGeneratorAttribute), false)
                                                                      .Single() as ExcelGeneratorAttribute).Order)
                                                        .Select((x, index) => new
                                                        {
                                                            Property = x,
                                                            ColumnIndex = index + 1
                                                        }).ToList();

            propertyList.ForEach(x =>
            {
                var headerRow = 1;
                var attributeValue = x.Property.GetCustomAttributes(false).FirstOrDefault(a => a is ExcelGeneratorAttribute) as ExcelGeneratorAttribute;
                workSheet.Cells[headerRow, x.ColumnIndex].Value = string.IsNullOrEmpty(attributeValue.ColumnName) ?  x.Property.Name : attributeValue.ColumnName;

                if (x.Property.PropertyType == typeof(DateTime?) || x.Property.PropertyType == typeof(DateTime))
                {
                    var dateFormat = string.Empty;

                    if (string.IsNullOrEmpty(attributeValue.DateFormat))
                        dateFormat = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;

                    else
                        dateFormat = attributeValue.DateFormat;

                    workSheet.Cells[headerRow + 1, x.ColumnIndex, headerRow + records.Count(), x.ColumnIndex].Style.Numberformat.Format = dateFormat;
                }
            });

            foreach (var (currentRow, rowIndex) in records.Select((object row, int index) => (row, index + 2)))
            {
                propertyList.ForEach(x =>
                {
                    workSheet.Cells[rowIndex, x.ColumnIndex].Value = x.Property.GetValue(currentRow);
                });
            }
        }
    }
}
