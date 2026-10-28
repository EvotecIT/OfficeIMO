using System;
using System.Collections.Generic;
using OfficeIMO.Excel;
using OfficeIMO.Drawing;
using OfficeIMO.Data;

namespace OfficeIMO.Examples.Excel
{
    public static class RowsFromObjects
    {
        private class Address {
            public string? City { get; set; }
            public string? Street { get; set; }
        }

        private class Person {
            public string Name { get; set; } = string.Empty;
            public int Age { get; set; }
            public Address? Address { get; set; }
        }

        public static void Example(string folderPath, bool openExcel)
        {
            Console.WriteLine("[*] Excel - Rows from objects");
            string filePath = System.IO.Path.Combine(folderPath, "RowsFromObjects.xlsx");
            var data = new List<Person>
            {
                new Person
                {
                    Name = "Alice",
                    Age = 30,
                    Address = new Address
                    {
                        City = "New York",
                        Street = "1st Ave"
                    }
                },
                new Person
                {
                    Name = "Bob",
                    Age = 25,
                    Address = new Address
                    {
                        City = "Los Angeles",
                        Street = "Main St"
                    }
                },
                new Person
                {
                    Name = "Charlie",
                    Age = 40,
                    Address = null
                }
            };
            using (var document = ExcelDocument.Create(filePath))
            {
                // Sheet 1: Basic transform with formatting and title-case headers
                {
                    var sheet1 = document.AddSheet("People");
                    int nextRow2 = 1;
                    var projection3 = new OfficeIMO.Data.ObjectFlattenerOptions();
                    {
                        (projection3).ExpandProperties.Add(nameof(Person.Address));
                        (projection3).HeaderPrefixTrimPaths = new[]
                        {
                            "Address."
                        };
                        (projection3).HeaderCase = HeaderCase.Title;
                        (projection3).NullPolicy = NullPolicy.EmptyString;
                        (projection3).Formatters["Age"] = v => $"{v} years";
                    }

                    var table4 = sheet1.AddTable(data, new ExcelTableOptions { Name = "People", Style = ExcelTableStyle.TableStyleMedium9, IncludeAutoFilter = true, Projection = projection3 }, row: nextRow2);
                    nextRow2 = table4.AsRange().LastRow + 1;
                    sheet1.AutoFitColumns();
                }

                // Sheet 2: Include/Exclude — keep Name/Age/Address.City; drop Address.Street
                {
                    var sheet5 = document.AddSheet("IncludeExclude");
                    int nextRow6 = 1;
                    var projection7 = new OfficeIMO.Data.ObjectFlattenerOptions();
                    {
                        (projection7).ExpandProperties.Add(nameof(Person.Address));
                        (projection7).HeaderCase = HeaderCase.Title;
                        (projection7).Include(nameof(Person.Name), nameof(Person.Age), "Address.City");
                        (projection7).Exclude("Address.Street");
                    }

                    var table8 = sheet5.AddTable(data, new ExcelTableOptions { Name = "People", Style = ExcelTableStyle.TableStyleMedium4, IncludeAutoFilter = true, Projection = projection7 }, row: nextRow6);
                    nextRow6 = table8.AsRange().LastRow + 1;
                    sheet5.AutoFitColumns();
                }

                // Sheet 3: Ordering — pin Name first, then Age, then Address.City, push Address.Street last
                {
                    var sheet9 = document.AddSheet("Ordering");
                    int nextRow10 = 1;
                    var projection11 = new OfficeIMO.Data.ObjectFlattenerOptions();
                    {
                        (projection11).ExpandProperties.Add(nameof(Person.Address));
                        (projection11).HeaderCase = HeaderCase.Title;
                        (projection11).Order(pinFirst: new[] { nameof(Person.Name) }, priority: new[] { nameof(Person.Age), "Address.City" }, pinLast: new[] { "Address.Street" });
                    }

                    var table12 = sheet9.AddTable(data, new ExcelTableOptions { Name = "People", Style = ExcelTableStyle.TableStyleMedium6, IncludeAutoFilter = true, Projection = projection11 }, row: nextRow10);
                    nextRow10 = table12.AsRange().LastRow + 1;
                    sheet9.AutoFitColumns();
                }

                document.Save();
                if (openExcel)
                    document.OpenInApplication();
            }
        }
    }
}