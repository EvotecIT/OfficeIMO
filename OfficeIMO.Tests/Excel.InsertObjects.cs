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
        }

        [Fact]
        public void Test_InsertObjectsFromClass() {
            string filePath = Path.Combine(_directoryWithFiles, "InsertObjectsClass.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                var people = new List<Person> {
                    new Person { Name = "Alice", Age = 30 },
                    new Person { Name = "Bob", Age = 40 }
                };
                sheet.InsertObjects(people);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                SharedStringTablePart shared = spreadsheet.WorkbookPart!.SharedStringTablePart!;

                Cell header1 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "A1");
                Cell header2 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "B1");
                Assert.Equal("Name", shared.SharedStringTable!.ElementAt(int.Parse(header1.CellValue!.Text)).InnerText);
                Assert.Equal("Age", shared.SharedStringTable!.ElementAt(int.Parse(header2.CellValue!.Text)).InnerText);

                Cell name1 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "A2");
                Cell age1 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "B2");
                Assert.Equal("Alice", shared.SharedStringTable!.ElementAt(int.Parse(name1.CellValue!.Text)).InnerText);
                Assert.Equal("30", age1.CellValue!.Text);

                Cell name2 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "A3");
                Cell age2 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "B3");
                Assert.Equal("Bob", shared.SharedStringTable!.ElementAt(int.Parse(name2.CellValue!.Text)).InnerText);
                Assert.Equal("40", age2.CellValue!.Text);
            }
        }

        [Fact]
        public void Test_InsertObjectsFromDictionary() {
            string filePath = Path.Combine(_directoryWithFiles, "InsertObjectsDictionary.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                var items = new List<Dictionary<string, object>> {
                    new Dictionary<string, object> { { "Name", "Alice" }, { "Age", 30 } },
                    new Dictionary<string, object> { { "Name", "Bob" }, { "Age", 40 } }
                };
                sheet.InsertObjects(items);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                SharedStringTablePart shared = spreadsheet.WorkbookPart!.SharedStringTablePart!;

                Cell header1 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "A1");
                Cell header2 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "B1");
                Assert.Equal("Name", shared.SharedStringTable!.ElementAt(int.Parse(header1.CellValue!.Text)).InnerText);
                Assert.Equal("Age", shared.SharedStringTable!.ElementAt(int.Parse(header2.CellValue!.Text)).InnerText);

                Cell name1 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "A2");
                Cell age1 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "B2");
                Assert.Equal("Alice", shared.SharedStringTable!.ElementAt(int.Parse(name1.CellValue!.Text)).InnerText);
                Assert.Equal("30", age1.CellValue!.Text);

                Cell name2 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "A3");
                Cell age2 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "B3");
                Assert.Equal("Bob", shared.SharedStringTable!.ElementAt(int.Parse(name2.CellValue!.Text)).InnerText);
                Assert.Equal("40", age2.CellValue!.Text);
            }
        }
    }
}
