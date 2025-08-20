using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        private class Person {
            public string Name { get; set; } = string.Empty;
            public int Age { get; set; }
            public Address? Address { get; set; }
        }

        private class Address {
            public string City { get; set; } = string.Empty;
            public string Zip { get; set; } = string.Empty;
        }

        [Fact]
        public void Test_InsertObjectsFromClass() {
            string filePath = Path.Combine(_directoryWithFiles, "InsertObjectsClass.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                var people = new List<Person> {
                    new Person {
                        Name = "Alice",
                        Age = 30,
                        Address = new Address { City = "London", Zip = "SW1" }
                    },
                    new Person {
                        Name = "Bob",
                        Age = 40,
                        Address = new Address { City = "Paris", Zip = "75001" }
                    }
                };
                sheet.InsertObjects(people);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                SharedStringTablePart shared = spreadsheet.WorkbookPart!.SharedStringTablePart!;

                Cell header1 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "A1");
                Cell header2 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "B1");
                Cell header3 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "C1");
                Cell header4 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "D1");
                Assert.Equal("Name", shared.SharedStringTable!.ElementAt(int.Parse(header1.CellValue!.Text)).InnerText);
                Assert.Equal("Age", shared.SharedStringTable!.ElementAt(int.Parse(header2.CellValue!.Text)).InnerText);
                Assert.Equal("Address.City", shared.SharedStringTable!.ElementAt(int.Parse(header3.CellValue!.Text)).InnerText);
                Assert.Equal("Address.Zip", shared.SharedStringTable!.ElementAt(int.Parse(header4.CellValue!.Text)).InnerText);

                Cell name1 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "A2");
                Cell age1 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "B2");
                Cell city1 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "C2");
                Cell zip1 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "D2");
                Assert.Equal("Alice", shared.SharedStringTable!.ElementAt(int.Parse(name1.CellValue!.Text)).InnerText);
                Assert.Equal("30", age1.CellValue!.Text);
                Assert.Equal("London", shared.SharedStringTable!.ElementAt(int.Parse(city1.CellValue!.Text)).InnerText);
                Assert.Equal("SW1", shared.SharedStringTable!.ElementAt(int.Parse(zip1.CellValue!.Text)).InnerText);

                Cell name2 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "A3");
                Cell age2 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "B3");
                Cell city2 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "C3");
                Cell zip2 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "D3");
                Assert.Equal("Bob", shared.SharedStringTable!.ElementAt(int.Parse(name2.CellValue!.Text)).InnerText);
                Assert.Equal("40", age2.CellValue!.Text);
                Assert.Equal("Paris", shared.SharedStringTable!.ElementAt(int.Parse(city2.CellValue!.Text)).InnerText);
                Assert.Equal("75001", shared.SharedStringTable!.ElementAt(int.Parse(zip2.CellValue!.Text)).InnerText);
            }
        }

        [Fact]
        public void Test_InsertObjectsFromDictionary() {
            string filePath = Path.Combine(_directoryWithFiles, "InsertObjectsDictionary.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                var items = new List<Dictionary<string, object>> {
                    new Dictionary<string, object> {
                        { "Name", "Alice" },
                        { "Age", 30 },
                        { "Address", new Dictionary<string, object> { { "City", "London" }, { "Zip", "SW1" } } }
                    },
                    new Dictionary<string, object> {
                        { "Name", "Bob" },
                        { "Age", 40 },
                        { "Address", new Dictionary<string, object> { { "City", "Paris" }, { "Zip", "75001" } } }
                    }
                };
                sheet.InsertObjects(items);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                SharedStringTablePart shared = spreadsheet.WorkbookPart!.SharedStringTablePart!;

                Cell header1 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "A1");
                Cell header2 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "B1");
                Cell header3 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "C1");
                Cell header4 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "D1");
                Assert.Equal("Name", shared.SharedStringTable!.ElementAt(int.Parse(header1.CellValue!.Text)).InnerText);
                Assert.Equal("Age", shared.SharedStringTable!.ElementAt(int.Parse(header2.CellValue!.Text)).InnerText);
                Assert.Equal("Address.City", shared.SharedStringTable!.ElementAt(int.Parse(header3.CellValue!.Text)).InnerText);
                Assert.Equal("Address.Zip", shared.SharedStringTable!.ElementAt(int.Parse(header4.CellValue!.Text)).InnerText);

                Cell name1 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "A2");
                Cell age1 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "B2");
                Cell city1 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "C2");
                Cell zip1 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "D2");
                Assert.Equal("Alice", shared.SharedStringTable!.ElementAt(int.Parse(name1.CellValue!.Text)).InnerText);
                Assert.Equal("30", age1.CellValue!.Text);
                Assert.Equal("London", shared.SharedStringTable!.ElementAt(int.Parse(city1.CellValue!.Text)).InnerText);
                Assert.Equal("SW1", shared.SharedStringTable!.ElementAt(int.Parse(zip1.CellValue!.Text)).InnerText);

                Cell name2 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "A3");
                Cell age2 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "B3");
                Cell city2 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "C3");
                Cell zip2 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "D3");
                Assert.Equal("Bob", shared.SharedStringTable!.ElementAt(int.Parse(name2.CellValue!.Text)).InnerText);
                Assert.Equal("40", age2.CellValue!.Text);
                Assert.Equal("Paris", shared.SharedStringTable!.ElementAt(int.Parse(city2.CellValue!.Text)).InnerText);
                Assert.Equal("75001", shared.SharedStringTable!.ElementAt(int.Parse(zip2.CellValue!.Text)).InnerText);
            }
        }
    }
}
