using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Data;
using Xunit;

namespace OfficeIMO.Tests
{
    public class ExcelRowsFromObjectsTests
    {
        private static string GetCellValue(SpreadsheetDocument document, WorksheetPart worksheetPart, string cellReference) {
            var cell = worksheetPart.Worksheet.Descendants<Cell>().First(c => c.CellReference != null && c.CellReference.Value == cellReference);
            if (cell.DataType?.Value == CellValues.InlineString) {
                return cell.InlineString?.InnerText ?? string.Empty;
            }
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
        public void RowsFrom_WritesHeadersAndValues_DeterministicOrder()
        {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            var data = new[]
            {
                new Person
                {
                    Name = "Alice",
                    Age = 30,
                    Address = new Address
                    {
                        City = "NY",
                        Street = "1st"
                    },
                    Tags = new List<string>
                    {
                        "a",
                        "b"
                    }
                },
                new Person
                {
                    Name = "Bob",
                    Age = 25,
                    Address = new Address
                    {
                        City = "LA",
                        Street = "Main"
                    },
                    Tags = new List<string>
                    {
                        "c"
                    }
                }
            };
            using (var doc = ExcelDocument.Create(filePath))
            {
                {
                    var sheet1 = doc.AddSheet("People");
                    int nextRow2 = 1;
                    var projection3 = new OfficeIMO.Data.ObjectFlattenerOptions();
                    {
                        (projection3).ExpandProperties.Add(nameof(Person.Address));
                    }

                    var table4 = sheet1.AddTable(data, new ExcelTableOptions { Name = null, Style = ExcelTableStyle.TableStyleMedium2, IncludeAutoFilter = true, Projection = projection3 }, row: nextRow2);
                    nextRow2 = table4.AsRange().LastRow + 1;
                    doc.Save();
                }
            }

            using (var document = SpreadsheetDocument.Open(filePath, false))
            {
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
        public void RowsFrom_DefaultOptions_JoinCollectionsBeforeDirectSaveShortcut()
        {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            var data = new[]
            {
                new Person
                {
                    Name = "Alice",
                    Age = 30,
                    Tags = new List<string>
                    {
                        "a",
                        "b"
                    }
                }
            };
            using (var doc = ExcelDocument.Create(filePath))
            {
                {
                    var sheet5 = doc.AddSheet("People");
                    int nextRow6 = 1;
                    var table7 = sheet5.AddTable(data, new ExcelTableOptions { Name = null, Style = ExcelTableStyle.TableStyleMedium2, IncludeAutoFilter = true, Projection = null }, row: nextRow6);
                    nextRow6 = table7.AsRange().LastRow + 1;
                    doc.Save();
                }
            }

            using (var document = SpreadsheetDocument.Open(filePath, false))
            {
                var workbookPart = document.WorkbookPart;
                Assert.NotNull(workbookPart);
                var wsPart = workbookPart.WorksheetParts.First();
                Assert.Equal("Tags", GetCellValue(document, wsPart, "D1"));
                Assert.Equal("a,b", GetCellValue(document, wsPart, "D2"));
            }

            File.Delete(filePath);
        }

        [Fact]
        public void NativeTable_NormalizesBlankHeadersAfterTrim()
        {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            var data = new[]
            {
                new Person
                {
                    Name = "Alice"
                }
            };
            using (var doc = ExcelDocument.Create(filePath))
            {
                {
                    var sheet8 = doc.AddSheet("People");
                    int nextRow9 = 1;
                    var projection10 = new OfficeIMO.Data.ObjectFlattenerOptions();
                    {
                        (projection10).Columns = new[]
                        {
                            nameof(Person.Name)
                        };
                        (projection10).HeaderPrefixTrimPaths = new[]
                        {
                            nameof(Person.Name)
                        };
                    }

                    var table11 = sheet8.AddTable(data, new ExcelTableOptions { Name = null, Style = ExcelTableStyle.TableStyleMedium2, IncludeAutoFilter = true, Projection = projection10 }, row: nextRow9);
                    nextRow9 = table11.AsRange().LastRow + 1;
                    doc.Save();
                }
            }

            using (var document = SpreadsheetDocument.Open(filePath, false))
            {
                var workbookPart = document.WorkbookPart;
                Assert.NotNull(workbookPart);
                var wsPart = workbookPart.WorksheetParts.First();
                Assert.Equal("Column1", GetCellValue(document, wsPart, "A1"));
                Assert.Equal("Alice", GetCellValue(document, wsPart, "A2"));
            }

            File.Delete(filePath);
        }

        [Fact]
        public void RowsFrom_AppliesSelectionRulesToExplicitColumnHeaders()
        {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            var data = new[]
            {
                new Person
                {
                    Name = "Alice",
                    Age = 30
                }
            };
            using (var doc = ExcelDocument.Create(filePath))
            {
                {
                    var sheet12 = doc.AddSheet("People");
                    int nextRow13 = 1;
                    var projection14 = new OfficeIMO.Data.ObjectFlattenerOptions();
                    {
                        (projection14).Columns = new[]
                        {
                            nameof(Person.Name),
                            nameof(Person.Age)
                        };
                        (projection14).ExcludeProperties = new[]
                        {
                            nameof(Person.Age)
                        };
                    }

                    var table15 = sheet12.AddTable(data, new ExcelTableOptions { Name = null, Style = ExcelTableStyle.TableStyleMedium2, IncludeAutoFilter = true, Projection = projection14 }, row: nextRow13);
                    nextRow13 = table15.AsRange().LastRow + 1;
                    doc.Save();
                }
            }

            using (var document = SpreadsheetDocument.Open(filePath, false))
            {
                WorksheetPart worksheet = document.WorkbookPart!.WorksheetParts.First();
                Assert.Equal("Name", GetCellValue(document, worksheet, "A1"));
                Assert.Equal("Alice", GetCellValue(document, worksheet, "A2"));
                Assert.DoesNotContain(worksheet.Worksheet.Descendants<Cell>(), cell => string.Equals(cell.CellReference?.Value, "B1", StringComparison.OrdinalIgnoreCase) || string.Equals(cell.CellReference?.Value, "B2", StringComparison.OrdinalIgnoreCase));
            }

            File.Delete(filePath);
        }

        [Fact]
        public void RowsFrom_HeaderCaseTransformsNestedPaths()
        {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            var data = new[]
            {
                new Person
                {
                    Name = "Alice",
                    Address = new Address
                    {
                        City = "NY",
                        Street = "1st"
                    }
                }
            };
            using (var doc = ExcelDocument.Create(filePath))
            {
                {
                    var sheet16 = doc.AddSheet("Pascal");
                    int nextRow17 = 1;
                    var projection18 = new OfficeIMO.Data.ObjectFlattenerOptions();
                    {
                        (projection18).ExpandProperties.Add(nameof(Person.Address));
                        (projection18).HeaderCase = HeaderCase.Pascal;
                    }

                    var table19 = sheet16.AddTable(data, new ExcelTableOptions { Name = null, Style = ExcelTableStyle.TableStyleMedium2, IncludeAutoFilter = true, Projection = projection18 }, row: nextRow17);
                    nextRow17 = table19.AsRange().LastRow + 1;
                    var sheet20 = doc.AddSheet("Title");
                    int nextRow21 = 1;
                    var projection22 = new OfficeIMO.Data.ObjectFlattenerOptions();
                    {
                        (projection22).ExpandProperties.Add(nameof(Person.Address));
                        (projection22).HeaderCase = HeaderCase.Title;
                    }

                    var table23 = sheet20.AddTable(data, new ExcelTableOptions { Name = null, Style = ExcelTableStyle.TableStyleMedium2, IncludeAutoFilter = true, Projection = projection22 }, row: nextRow21);
                    nextRow21 = table23.AsRange().LastRow + 1;
                    doc.Save();
                }
            }

            using (var document = SpreadsheetDocument.Open(filePath, false))
            {
                var workbookPart = document.WorkbookPart;
                Assert.NotNull(workbookPart);
                var sheets = workbookPart!.Workbook.Descendants<Sheet>().ToList();
                var pascalPart = (WorksheetPart)workbookPart.GetPartById(sheets.First(s => s.Name == "Pascal").Id!.Value!);
                var titlePart = (WorksheetPart)workbookPart.GetPartById(sheets.First(s => s.Name == "Title").Id!.Value!);
                Assert.Equal("AddressCity", GetCellValue(document, pascalPart, "C1"));
                Assert.Equal("Address City", GetCellValue(document, titlePart, "C1"));
            }

            File.Delete(filePath);
        }

        [Fact]
        public void RowsFrom_NullPolicyAndFormatter()
        {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            var data = new[]
            {
                new Person
                {
                    Name = "Alice",
                    Age = 30,
                    Address = new Address
                    {
                        City = "NY",
                        Street = null
                    },
                    Tags = null
                },
                new Person
                {
                    Name = "Bob",
                    Age = 25,
                    Address = null,
                    Tags = null
                }
            };
            using (var doc = ExcelDocument.Create(filePath))
            {
                {
                    var sheet24 = doc.AddSheet("People");
                    int nextRow25 = 1;
                    var projection26 = new OfficeIMO.Data.ObjectFlattenerOptions();
                    {
                        (projection26).ExpandProperties.Add(nameof(Person.Address));
                        (projection26).NullPolicy = NullPolicy.DefaultValue;
                        (projection26).DefaultValues["Address.City"] = "N/A";
                        (projection26).Formatters["Age"] = v => $"{v}y";
                    }

                    var table27 = sheet24.AddTable(data, new ExcelTableOptions { Name = null, Style = ExcelTableStyle.TableStyleMedium2, IncludeAutoFilter = true, Projection = projection26 }, row: nextRow25);
                    nextRow25 = table27.AsRange().LastRow + 1;
                    doc.Save();
                }
            }

            using (var document = SpreadsheetDocument.Open(filePath, false))
            {
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
        public void RowsFrom_CollectionExpandRows_CreatesTable()
        {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            var data = new[]
            {
                new Person
                {
                    Name = "Alice",
                    Age = 30,
                    Tags = new List<string>
                    {
                        "a",
                        "b"
                    }
                }
            };
            using (var doc = ExcelDocument.Create(filePath))
            {
                {
                    var sheet28 = doc.AddSheet("People");
                    int nextRow29 = 1;
                    var projection30 = new OfficeIMO.Data.ObjectFlattenerOptions();
                    {
                        (projection30).ExpandProperties.Add(nameof(Person.Tags));
                        (projection30).CollectionMode = CollectionMode.ExpandRows;
                    }

                    var table31 = sheet28.AddTable(data, new ExcelTableOptions { Name = "People", Style = OfficeIMO.Excel.ExcelTableStyle.TableStyleMedium9, IncludeAutoFilter = true, Projection = projection30 }, row: nextRow29);
                    nextRow29 = table31.AsRange().LastRow + 1;
                    doc.Save();
                }
            }

            using (var document = SpreadsheetDocument.Open(filePath, false))
            {
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

        [Fact]
        public void RowsFrom_CollectionExpandRows_StreamsNestedCollectionsOnce()
        {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            var tags = new SinglePassEnumerable<string>("a", "b");
            var data = new[]
            {
                new Person
                {
                    Name = "Alice",
                    Age = 30,
                    Tags = null
                }
            };
            using (var doc = ExcelDocument.Create(filePath))
            {
                {
                    var sheet32 = doc.AddSheet("People");
                    int nextRow33 = 1;
                    var projection34 = new OfficeIMO.Data.ObjectFlattenerOptions();
                    {
                        (projection34).ExpandProperties.Add(nameof(PersonWithEnumerableTags.Tags));
                        (projection34).CollectionMode = CollectionMode.ExpandRows;
                    }

                    var table35 = sheet32.AddTable(new[] { new PersonWithEnumerableTags { Name = data[0].Name, Age = data[0].Age, Tags = tags } }, new ExcelTableOptions { Name = null, Style = ExcelTableStyle.TableStyleMedium2, IncludeAutoFilter = true, Projection = projection34 }, row: nextRow33);
                    nextRow33 = table35.AsRange().LastRow + 1;
                    doc.Save();
                }
            }

            Assert.Equal(1, tags.EnumerationCount);
            using (var document = SpreadsheetDocument.Open(filePath, false))
            {
                var wsPart = document.WorkbookPart!.WorksheetParts.First();
                Assert.Equal("Alice", GetCellValue(document, wsPart, "A2"));
                Assert.Equal("a", GetCellValue(document, wsPart, "C2"));
                Assert.Equal("Alice", GetCellValue(document, wsPart, "A3"));
                Assert.Equal("b", GetCellValue(document, wsPart, "C3"));
            }

            File.Delete(filePath);
        }

        private sealed class PersonWithEnumerableTags {
            public string Name { get; set; } = string.Empty;

            public int Age { get; set; }

            public IEnumerable<string>? Tags { get; set; }
        }

        private sealed class SinglePassEnumerable<T> : IEnumerable<T> {
            private readonly T[] _items;

            public SinglePassEnumerable(params T[] items) {
                _items = items;
            }

            public int EnumerationCount { get; private set; }

            public IEnumerator<T> GetEnumerator() {
                EnumerationCount++;
                if (EnumerationCount > 1) {
                    throw new InvalidOperationException("Nested collection should be streamed once.");
                }

                return ((IEnumerable<T>)_items).GetEnumerator();
            }

            System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
        }
    }
}