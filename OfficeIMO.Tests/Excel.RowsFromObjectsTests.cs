using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Fluent;
using OfficeIMO.Excel.Utilities;
using Xunit;

namespace OfficeIMO.Tests {
    public class ExcelRowsFromObjectsTests {
        private static string GetCellValue(SpreadsheetDocument document, WorksheetPart worksheetPart, string cellReference) {
            var cell = worksheetPart.Worksheet.Descendants<Cell>().First(c => c.CellReference != null && c.CellReference.Value == cellReference);
            var value = cell.CellValue?.Text ?? string.Empty;
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString) {
                var table = document.WorkbookPart?.SharedStringTablePart?.SharedStringTable;
                if (table != null && int.TryParse(value, out int id)) {
                    return table.ChildElements[id].InnerText;
                }
            }
            return value;
        }

        private class Address {
            public string? City { get; set; }
            public string? Street { get; set; }
        }

        private class Person {
            public string Name { get; set; } = string.Empty;
            public int Age { get; set; }
            public Address? Address { get; set; }
            public List<string>? Tags { get; set; }
        }

        [Fact]
        public void RowsFrom_WritesHeadersAndValues_DeterministicOrder() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            var data = new[] {
                new Person { Name = "Alice", Age = 30, Address = new Address { City = "NY", Street = "1st" }, Tags = new List<string>{"a","b"} },
                new Person { Name = "Bob", Age = 25, Address = new Address { City = "LA", Street = "Main" }, Tags = new List<string>{"c"} }
            };

            using (var doc = ExcelDocument.Create(filePath)) {
                doc.AsFluent()
                    .Sheet("People", s => s.RowsFrom(data, o => {
                        o.ExpandProperties.Add(nameof(Person.Address));
                    }))
                    .End()
                    .Save();
            }

            using (var document = SpreadsheetDocument.Open(filePath, false)) {
                var workbookPart = document.WorkbookPart;
                Assert.NotNull(workbookPart);
                var wsPart = workbookPart.WorksheetParts.First();
                Assert.Equal("Name", GetCellValue(document, wsPart, "A1"));
                Assert.Equal("Age", GetCellValue(document, wsPart, "B1"));
                Assert.Equal("Address.City", GetCellValue(document, wsPart, "C1"));
                Assert.Equal("Address.Street", GetCellValue(document, wsPart, "D1"));
                Assert.Equal("Tags", GetCellValue(document, wsPart, "E1"));

                Assert.Equal("Alice", GetCellValue(document, wsPart, "A2"));
                Assert.Equal("30", GetCellValue(document, wsPart, "B2"));
                Assert.Equal("NY", GetCellValue(document, wsPart, "C2"));
                Assert.Equal("1st", GetCellValue(document, wsPart, "D2"));
                Assert.Equal("a,b", GetCellValue(document, wsPart, "E2"));
            }

            File.Delete(filePath);
        }

        [Fact]
        public void RowsFrom_NullPolicyAndFormatter() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            var data = new[] {
                new Person { Name = "Alice", Age = 30, Address = new Address { City = "NY", Street = null }, Tags = null },
                new Person { Name = "Bob", Age = 25, Address = null, Tags = null }
            };

            using (var doc = ExcelDocument.Create(filePath)) {
                doc.AsFluent()
                    .Sheet("People", s => s.RowsFrom(data, o => {
                        o.ExpandProperties.Add(nameof(Person.Address));
                        o.NullPolicy = NullPolicy.DefaultValue;
                        o.DefaultValues["Address.City"] = "N/A";
                        o.Formatters["Age"] = v => $"{v}y";
                    }))
                    .End()
                    .Save();
            }

            using (var document = SpreadsheetDocument.Open(filePath, false)) {
                var workbookPart = document.WorkbookPart;
                Assert.NotNull(workbookPart);
                var wsPart = workbookPart.WorksheetParts.First();
                Assert.Equal("N/A", GetCellValue(document, wsPart, "C3"));
                Assert.Equal("", GetCellValue(document, wsPart, "D2"));
                Assert.Equal("30y", GetCellValue(document, wsPart, "B2"));
            }

            File.Delete(filePath);
        }

        [Fact]
        public void RowsFrom_CollectionExpandRows_CreatesTable() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            var data = new[] {
                new Person { Name = "Alice", Age = 30, Tags = new List<string>{"a","b"} }
            };

            using (var doc = ExcelDocument.Create(filePath)) {
                doc.AsFluent()
                    .Sheet("People", s => s
                        .RowsFrom(data, o => {
                            o.ExpandProperties.Add(nameof(Person.Tags));
                            o.CollectionMode = CollectionMode.ExpandRows;
                        })
                        .Table("People", t => t.Style(OfficeIMO.Excel.TableStyle.TableStyleMedium9)))
                    .End()
                    .Save();
            }

            using (var document = SpreadsheetDocument.Open(filePath, false)) {
                var workbookPart = document.WorkbookPart;
                Assert.NotNull(workbookPart);
                var wsPart = workbookPart.WorksheetParts.First();
                // two rows for tags -> name repeats twice
                Assert.Equal("Alice", GetCellValue(document, wsPart, "A2"));
                Assert.Equal("Alice", GetCellValue(document, wsPart, "A3"));
                var table = wsPart.TableDefinitionParts.First();
                Assert.NotNull(table.Table);
                Assert.NotNull(table.Table!.DisplayName);
                Assert.Equal("People", table.Table.DisplayName!.Value);
                Assert.Equal("TableStyleMedium9", table.Table.TableStyleInfo?.Name?.Value);
            }

            File.Delete(filePath);
        }
    }
}
