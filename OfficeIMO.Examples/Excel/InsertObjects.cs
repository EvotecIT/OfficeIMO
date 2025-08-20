using System;
using System.Collections.Generic;
using OfficeIMO.Excel;

namespace OfficeIMO.Examples.Excel {
    /// <summary>
    /// Demonstrates inserting objects into a worksheet.
    /// </summary>
    public static class InsertObjects {
        public static void Example(string folderPath, bool openExcel) {
            Console.WriteLine("[*] Excel - Insert objects");
            string filePath = System.IO.Path.Combine(folderPath, "InsertObjects.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                var people = new List<Person> {
                    new Person {
                        Name = "Alice",
                        Age = 30,
                        Address = new Address { City = "London", ZipCode = "SW1" }
                    },
                    new Person {
                        Name = "Bob",
                        Age = 40,
                        Address = new Address { City = "Paris", ZipCode = "75001" }
                    }
                };
                sheet.InsertObjects(people);
                document.Save(openExcel);
            }
        }

        private class Person {
            public string Name { get; set; } = string.Empty;
            public int Age { get; set; }
            public Address? Address { get; set; }
        }

        private class Address {
            public string City { get; set; } = string.Empty;
            public string ZipCode { get; set; } = string.Empty;
        }
    }
}
