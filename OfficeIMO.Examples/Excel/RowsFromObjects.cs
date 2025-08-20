using System;
using System.Collections.Generic;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Fluent;

namespace OfficeIMO.Examples.Excel {
    public static class RowsFromObjects {
        private class Address {
            public string? City { get; set; }
            public string? Street { get; set; }
        }

        private class Person {
            public string Name { get; set; } = string.Empty;
            public int Age { get; set; }
            public Address? Address { get; set; }
        }

        public static void Example(string folderPath, bool openExcel) {
            Console.WriteLine("[*] Excel - Rows from objects");
            string filePath = System.IO.Path.Combine(folderPath, "RowsFromObjects.xlsx");

            var data = new List<Person> {
                new Person { Name = "Alice", Age = 30, Address = new Address { City = "New York", Street = "1st Ave" } },
                new Person { Name = "Bob", Age = 25, Address = new Address { City = "Los Angeles", Street = "Main St" } },
                new Person { Name = "Charlie", Age = 40, Address = null }
            };

            using (var document = ExcelDocument.Create(filePath)) {
                document.AsFluent()
                    .Sheet("People", s => s
                        .RowsFrom(data, o => o.ExpandProperties.Add(nameof(Person.Address)))
                        .AutoFit(columns: true, rows: true))
                    .End()
                    .Save(openExcel);
            }
        }
    }
}
