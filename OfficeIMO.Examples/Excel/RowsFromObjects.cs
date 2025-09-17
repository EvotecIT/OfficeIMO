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
                // Sheet 1: Basic transform with formatting and title-case headers
                document.AsFluent()
                    .Sheet("People", s => s
                        .RowsFrom(data, o => {
                            o.ExpandProperties.Add(nameof(Person.Address));
                            o.HeaderPrefixTrimPaths = new[] { "Address." };
                            o.HeaderCase = HeaderCase.Title;
                            o.NullPolicy = NullPolicy.EmptyString;
                            o.Formatters["Age"] = v => $"{v} years";
                        })
                        .Table("People", t => t.Style(TableStyle.TableStyleMedium9))
                        .AutoFit(columns: true, rows: false))
                    .End();

                // Sheet 2: Include/Exclude — keep Name/Age/Address.City; drop Address.Street
                document.AsFluent()
                    .Sheet("IncludeExclude", s => s
                        .RowsFrom(data, o => {
                            o.ExpandProperties.Add(nameof(Person.Address));
                            o.HeaderCase = HeaderCase.Title;
                            o.Include(nameof(Person.Name), nameof(Person.Age), "Address.City");
                            o.Exclude("Address.Street");
                        })
                        .Table("People", t => t.Style(TableStyle.TableStyleMedium4))
                        .AutoFit(columns: true, rows: false))
                    .End();

                // Sheet 3: Ordering — pin Name first, then Age, then Address.City, push Address.Street last
                document.AsFluent()
                    .Sheet("Ordering", s => s
                        .RowsFrom(data, o => {
                            o.ExpandProperties.Add(nameof(Person.Address));
                            o.HeaderCase = HeaderCase.Title;
                            o.Order(
                                pinFirst: new[] { nameof(Person.Name) },
                                priority: new[] { nameof(Person.Age), "Address.City" },
                                pinLast: new[] { "Address.Street" }
                            );
                        })
                        .Table("People", t => t.Style(TableStyle.TableStyleMedium6))
                        .AutoFit(columns: true, rows: false))
                    .End();

                document.Save(openExcel);
            }
        }
    }
}
