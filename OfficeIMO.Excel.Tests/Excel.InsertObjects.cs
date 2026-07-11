using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
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

        private class LargeNode {
            public string Id { get; set; } = string.Empty;
            public NodeDetails Details { get; set; } = new NodeDetails();
            public List<string> Tags { get; set; } = new List<string>();
            public Dictionary<string, object> Attributes { get; set; } = new Dictionary<string, object>();
        }

        private class NodeDetails {
            public string Name { get; set; } = string.Empty;
            public NodeLevel Level { get; set; } = new NodeLevel();
        }

        private class NodeLevel {
            public string Stage { get; set; } = string.Empty;
            public Dictionary<string, string> Meta { get; set; } = new Dictionary<string, string>();
        }

        private class AttributeDetails {
            public int Score { get; set; }
            public int Weight { get; set; }
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
                Cell header1 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "A1");
                Cell header2 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "B1");
                Cell header3 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "C1");
                Cell header4 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "D1");
                Assert.Equal("Name", GetCellText(spreadsheet, header1));
                Assert.Equal("Age", GetCellText(spreadsheet, header2));
                Assert.Equal("Address.City", GetCellText(spreadsheet, header3));
                Assert.Equal("Address.Zip", GetCellText(spreadsheet, header4));

                Cell name1 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "A2");
                Cell age1 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "B2");
                Cell city1 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "C2");
                Cell zip1 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "D2");
                Assert.Equal("Alice", GetCellText(spreadsheet, name1));
                Assert.Equal("30", age1.CellValue!.Text);
                Assert.Equal("London", GetCellText(spreadsheet, city1));
                Assert.Equal("SW1", GetCellText(spreadsheet, zip1));

                Cell name2 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "A3");
                Cell age2 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "B3");
                Cell city2 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "C3");
                Cell zip2 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "D3");
                Assert.Equal("Bob", GetCellText(spreadsheet, name2));
                Assert.Equal("40", age2.CellValue!.Text);
                Assert.Equal("Paris", GetCellText(spreadsheet, city2));
                Assert.Equal("75001", GetCellText(spreadsheet, zip2));
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
                Cell header1 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "A1");
                Cell header2 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "B1");
                Cell header3 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "C1");
                Cell header4 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "D1");
                Assert.Equal("Name", GetCellText(spreadsheet, header1));
                Assert.Equal("Age", GetCellText(spreadsheet, header2));
                Assert.Equal("Address.City", GetCellText(spreadsheet, header3));
                Assert.Equal("Address.Zip", GetCellText(spreadsheet, header4));

                Cell name1 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "A2");
                Cell age1 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "B2");
                Cell city1 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "C2");
                Cell zip1 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "D2");
                Assert.Equal("Alice", GetCellText(spreadsheet, name1));
                Assert.Equal("30", age1.CellValue!.Text);
                Assert.Equal("London", GetCellText(spreadsheet, city1));
                Assert.Equal("SW1", GetCellText(spreadsheet, zip1));

                Cell name2 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "A3");
                Cell age2 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "B3");
                Cell city2 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "C3");
                Cell zip2 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "D3");
                Assert.Equal("Bob", GetCellText(spreadsheet, name2));
                Assert.Equal("40", age2.CellValue!.Text);
                Assert.Equal("Paris", GetCellText(spreadsheet, city2));
                Assert.Equal("75001", GetCellText(spreadsheet, zip2));
            }
        }

        [Fact]
        public void Test_InsertObjectsMaintainsHeaderOrderForLargeGraph() {
            string filePath = Path.Combine(_directoryWithFiles, "InsertObjectsLargeGraph.xlsx");

            var nodes = new List<LargeNode>();
            DateTime baseDate = new DateTime(2020, 1, 1);
            for (int i = 0; i < 150; i++) {
                var node = new LargeNode {
                    Id = $"Node-{i}",
                    Details = new NodeDetails {
                        Name = $"Node Name {i}",
                        Level = new NodeLevel {
                            Stage = $"Stage-{i % 4}",
                            Meta = new Dictionary<string, string> {
                                { "Category", $"Category-{i % 5}" },
                                { $"Extra-{i}", $"ExtraValue-{i}" }
                            }
                        }
                    },
                    Tags = new List<string> { "alpha", $"tag-{i % 3}" },
                    Attributes = new Dictionary<string, object> {
                        { "IsActive", i % 2 == 0 },
                        { "Created", baseDate.AddDays(i) },
                        { $"Dynamic-{i}", new AttributeDetails { Score = i, Weight = i * 2 } }
                    }
                };
                nodes.Add(node);
            }

            List<string> expectedHeaders = BuildExpectedHeaders(nodes.Cast<object>());

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.InsertObjects(nodes);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                SharedStringTablePart shared = spreadsheet.WorkbookPart!.SharedStringTablePart!;

                Row headerRow = wsPart.Worksheet.Descendants<Row>().First();
                List<string> headers = headerRow.Descendants<Cell>()
                    .Select(cell => GetCellText(cell, shared))
                    .ToList();

                Assert.Equal(expectedHeaders.Count, headers.Count);
                Assert.Equal(expectedHeaders, headers);
            }
        }

        private static List<string> BuildExpectedHeaders(IEnumerable<object> items) {
            MethodInfo? flattenMethod = typeof(ExcelSheet).GetMethod("FlattenObject", BindingFlags.NonPublic | BindingFlags.Static);
            Assert.NotNull(flattenMethod);

            var headers = new List<string>();
            var seen = new HashSet<string>();

            foreach (var item in items) {
                var dict = new Dictionary<string, object?>();
                flattenMethod!.Invoke(null, new object?[] { item, null, dict });
                foreach (var key in dict.Keys) {
                    if (seen.Add(key)) {
                        headers.Add(key);
                    }
                }
            }

            return headers;
        }

        private static string GetCellText(Cell cell, SharedStringTablePart shared) {
            if (cell.DataType?.Value == CellValues.InlineString) {
                return cell.InlineString?.InnerText ?? string.Empty;
            }
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString) {
                return shared.SharedStringTable!.ElementAt(int.Parse(cell.CellValue!.Text)).InnerText;
            }

            return cell.CellValue?.Text ?? string.Empty;
        }

        private static string GetCellText(SpreadsheetDocument spreadsheet, Cell cell) {
            if (cell.DataType?.Value == CellValues.InlineString) {
                return cell.InlineString?.InnerText ?? string.Empty;
            }
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString) {
                SharedStringTablePart shared = spreadsheet.WorkbookPart!.SharedStringTablePart!;
                return shared.SharedStringTable!.ElementAt(int.Parse(cell.CellValue!.Text)).InnerText;
            }

            return cell.CellValue?.Text ?? string.Empty;
        }
    }
}
