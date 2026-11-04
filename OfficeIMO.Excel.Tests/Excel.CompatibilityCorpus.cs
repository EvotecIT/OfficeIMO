using System.Data;
using System.ComponentModel;
using System.Runtime.Serialization;
using System.Xml.Linq;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        private sealed class CompatibilityMessyHeaderRow {
            public string? Value { get; set; }
            public string? Value_2 { get; set; }
            public string? Column3 { get; set; }
            public string? value_3 { get; set; }
            public string? Value_2_2 { get; set; }
            public string? Value_2_3 { get; set; }
        }

        private sealed class CompatibilityFriendlyHeaderRow {
            public string? FirstName { get; set; }
            public string? FirstName_2 { get; set; }
            public int TotalAmount2 { get; set; }
        }

        private sealed class CompatibilityAttributedHeaderRow {
            [DisplayName("First Name")]
            public string? GivenName { get; set; }

            [DataMember(Name = "Status Code")]
            public string? Status { get; set; }

            [ExcelColumn("Total %", "Total Percent")]
            public int CompletionPercent { get; set; }
        }

        private class CompatibilityInheritedHeaderBaseRow {
            [ExcelColumn("Inherited Total", "Legacy Total")]
            public virtual int Total { get; set; }
        }

        private sealed class CompatibilityInheritedHeaderRow : CompatibilityInheritedHeaderBaseRow {
            public override int Total { get; set; }
        }

        private sealed class CompatibilityScoreRow {
            public string? Name { get; set; }
            public int Score { get; set; }
        }

        [Fact]
        public void Compatibility_Corpus_RecoverableContentTypes_StayInSyncAcrossPathAndStreamReaders() {
            string filePath = Path.Combine(_directoryWithFiles, "CompatibilityCorpus.RecoverableContentTypes.xlsx");

            try {
                ExcelCompatibilityCorpusBuilder.CreateWorkbook(filePath, document => {
                    var sheet = document.AddSheet("Data");
                    sheet.Cell(1, 1).SetValue("Region");
                    sheet.Cell(1, 2).SetValue("Amount");
                    sheet.Cell(1, 3).SetValue("Status");
                    sheet.Cell(2, 1).SetValue("North");
                    sheet.Cell(2, 2).SetValue(10d);
                    sheet.Cell(2, 3).SetValue("Open");
                    sheet.Cell(3, 1).SetValue("South");
                    sheet.Cell(3, 2).SetValue(20d);
                    sheet.Cell(3, 3).SetValue("Closed");
                }, rewriteAppContentTypeToXml: true);

                byte[] workbookBytes = File.ReadAllBytes(filePath);
                using var pathReader = ExcelDocumentReader.Open(filePath);
                using var streamReader = ExcelDocumentReader.Open(new MemoryStream(workbookBytes, writable: false));
                using var bytesReader = ExcelDocumentReader.Open(workbookBytes);

                var pathRows = pathReader.GetSheet("Data").ReadObjects("A1:C3").ToList();
                var streamRows = streamReader.GetSheet("Data").ReadObjects("A1:C3").ToList();
                var bytesRows = bytesReader.GetSheet("Data").ReadObjects("A1:C3").ToList();

                Assert.Equal(pathReader.GetSheetNames(), streamReader.GetSheetNames());
                Assert.Equal(pathReader.GetSheetNames(), bytesReader.GetSheetNames());
                Assert.Equal(pathRows.Count, streamRows.Count);
                Assert.Equal(pathRows.Count, bytesRows.Count);
                Assert.Equal(pathRows[0]["Region"], streamRows[0]["Region"]);
                Assert.Equal(pathRows[0]["Amount"], streamRows[0]["Amount"]);
                Assert.Equal(pathRows[1]["Status"], streamRows[1]["Status"]);
                Assert.Equal(pathRows[0]["Region"], bytesRows[0]["Region"]);
                Assert.Equal(pathRows[0]["Amount"], bytesRows[0]["Amount"]);
                Assert.Equal(pathRows[1]["Status"], bytesRows[1]["Status"]);

            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Compatibility_Corpus_MessyHeaders_PreserveAllColumnsAcrossReadSurfaces() {
            string filePath = Path.Combine(_directoryWithFiles, "CompatibilityCorpus.MessyHeaders.xlsx");

            try {
                ExcelCompatibilityCorpusBuilder.CreateWorkbook(filePath, document => {
                    var sheet = document.AddSheet("Data");
                    sheet.Cell(1, 1).SetValue("Value");
                    sheet.Cell(1, 2).SetValue("  Value  ");
                    sheet.Cell(1, 3).SetValue("");
                    sheet.Cell(1, 4).SetValue("value");
                    sheet.Cell(1, 5).SetValue("Value_2");
                    sheet.Cell(1, 6).SetValue("Value_2");
                    sheet.Cell(2, 1).SetValue("A");
                    sheet.Cell(2, 2).SetValue("B");
                    sheet.Cell(2, 3).SetValue("C");
                    sheet.Cell(2, 4).SetValue("D");
                    sheet.Cell(2, 5).SetValue("E");
                    sheet.Cell(2, 6).SetValue("F");
                });

                using var reader = ExcelDocumentReader.Open(filePath);
                var row = Assert.Single(reader.GetSheet("Data").ReadObjects("A1:F2"));
                DataTable table = reader.GetSheet("Data").ReadRangeAsDataTable("A1:F2");

                Assert.Equal(new[] { "Value", "Value_2", "Column3", "value_3", "Value_2_2", "Value_2_3" },
                    table.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToArray());
                Assert.Equal("A", row["Value"]);
                Assert.Equal("B", row["Value_2"]);
                Assert.Equal("C", row["Column3"]);
                Assert.Equal("D", row["value_3"]);
                Assert.Equal("E", row["Value_2_2"]);
                Assert.Equal("F", row["Value_2_3"]);

                using var document = ExcelDocument.Load(filePath);
                var sheet = document.GetSheet("Data");
                var headerMap = sheet.GetHeaderMap();
                var typed = Assert.Single(sheet.RowsAs<CompatibilityMessyHeaderRow>("A1:F2"));

                Assert.Equal(new[] { "Value", "Value_2", "Column3", "value_3", "Value_2_2", "Value_2_3" }, headerMap.Keys.ToArray());
                Assert.Equal("A", typed.Value);
                Assert.Equal("B", typed.Value_2);
                Assert.Equal("C", typed.Column3);
                Assert.Equal("D", typed.value_3);
                Assert.Equal("E", typed.Value_2_2);
                Assert.Equal("F", typed.Value_2_3);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Compatibility_Corpus_BlankFallbackHeaders_DoNotOverrideExplicitColumnNames() {
            string filePath = Path.Combine(_directoryWithFiles, "CompatibilityCorpus.BlankFallbackHeaders.xlsx");

            try {
                ExcelCompatibilityCorpusBuilder.CreateWorkbook(filePath, document => {
                    var sheet = document.AddSheet("Data");
                    sheet.Cell(1, 1).SetValue("");
                    sheet.Cell(1, 2).SetValue("Column1");
                    sheet.Cell(1, 3).SetValue("");
                    sheet.Cell(2, 1).SetValue("GeneratedLeft");
                    sheet.Cell(2, 2).SetValue("Explicit");
                    sheet.Cell(2, 3).SetValue("GeneratedRight");
                });

                using var reader = ExcelDocumentReader.Open(filePath);
                var row = Assert.Single(reader.GetSheet("Data").ReadObjects("A1:C2"));
                DataTable table = reader.GetSheet("Data").ReadRangeAsDataTable("A1:C2");

                Assert.Equal(new[] { "Column1_2", "Column1", "Column3" },
                    table.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToArray());
                Assert.Equal("GeneratedLeft", row["Column1_2"]);
                Assert.Equal("Explicit", row["Column1"]);
                Assert.Equal("GeneratedRight", row["Column3"]);

                using var document = ExcelDocument.Load(filePath);
                var sheet = document.GetSheet("Data");
                var headerMap = sheet.GetHeaderMap();
                Assert.Equal(1, headerMap["Column1_2"]);
                Assert.Equal(2, headerMap["Column1"]);
                Assert.Equal(3, headerMap["Column3"]);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Compatibility_Corpus_BlankFallbackHeaders_RespectExplicitSuffixedNames() {
            string filePath = Path.Combine(_directoryWithFiles, "CompatibilityCorpus.BlankFallbackHeadersReservedSuffix.xlsx");

            try {
                ExcelCompatibilityCorpusBuilder.CreateWorkbook(filePath, document => {
                    var sheet = document.AddSheet("Data");
                    sheet.Cell(1, 1).SetValue("");
                    sheet.Cell(1, 2).SetValue("Column1");
                    sheet.Cell(1, 3).SetValue("Column1_2");
                    sheet.Cell(2, 1).SetValue("Generated");
                    sheet.Cell(2, 2).SetValue("ExplicitBase");
                    sheet.Cell(2, 3).SetValue("ExplicitSuffix");
                });

                using var reader = ExcelDocumentReader.Open(filePath);
                var row = Assert.Single(reader.GetSheet("Data").ReadObjects("A1:C2"));
                DataTable table = reader.GetSheet("Data").ReadRangeAsDataTable("A1:C2");

                Assert.Equal(new[] { "Column1_3", "Column1", "Column1_2" },
                    table.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToArray());
                Assert.Equal("Generated", row["Column1_3"]);
                Assert.Equal("ExplicitBase", row["Column1"]);
                Assert.Equal("ExplicitSuffix", row["Column1_2"]);

                using var document = ExcelDocument.Load(filePath);
                var sheet = document.GetSheet("Data");
                var headerMap = sheet.GetHeaderMap();
                Assert.Equal(1, headerMap["Column1_3"]);
                Assert.Equal(2, headerMap["Column1"]);
                Assert.Equal(3, headerMap["Column1_2"]);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Compatibility_Corpus_ShiftedRanges_KeepLocalHeaderDisambiguationStable() {
            string filePath = Path.Combine(_directoryWithFiles, "CompatibilityCorpus.ShiftedRangeHeaders.xlsx");

            try {
                ExcelCompatibilityCorpusBuilder.CreateWorkbook(filePath, document => {
                    var sheet = document.AddSheet("Data");
                    sheet.Cell(2, 2).SetValue("");
                    sheet.Cell(2, 3).SetValue("Column1");
                    sheet.Cell(2, 4).SetValue("Column1_2");
                    sheet.Cell(3, 2).SetValue("Generated");
                    sheet.Cell(3, 3).SetValue("ExplicitBase");
                    sheet.Cell(3, 4).SetValue("ExplicitSuffix");
                });

                using var reader = ExcelDocumentReader.Open(filePath);
                var row = Assert.Single(reader.GetSheet("Data").ReadObjects("B2:D3"));
                DataTable table = reader.GetSheet("Data").ReadRangeAsDataTable("B2:D3");

                Assert.Equal(new[] { "Column1_3", "Column1", "Column1_2" },
                    table.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToArray());
                Assert.Equal("Generated", row["Column1_3"]);
                Assert.Equal("ExplicitBase", row["Column1"]);
                Assert.Equal("ExplicitSuffix", row["Column1_2"]);

            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Compatibility_Corpus_NormalizeHeadersFalse_PreservesWhitespaceDistinctHeaders() {
            string filePath = Path.Combine(_directoryWithFiles, "CompatibilityCorpus.NormalizeHeadersFalse.xlsx");

            try {
                ExcelCompatibilityCorpusBuilder.CreateWorkbook(filePath, document => {
                    var sheet = document.AddSheet("Data");
                    sheet.Cell(1, 1).SetValue("Value");
                    sheet.Cell(1, 2).SetValue("  Value  ");
                    sheet.Cell(1, 3).SetValue("");
                    sheet.Cell(2, 1).SetValue("A");
                    sheet.Cell(2, 2).SetValue("B");
                    sheet.Cell(2, 3).SetValue("C");
                });

                var options = new ExcelReaderOptions { NormalizeHeaders = false };

                using var reader = ExcelDocumentReader.Open(filePath, options);
                var row = Assert.Single(reader.GetSheet("Data").ReadObjects("A1:C2"));
                DataTable table = reader.GetSheet("Data").ReadRangeAsDataTable("A1:C2");

                Assert.Equal(new[] { "Value", "  Value  ", "Column3" },
                    table.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToArray());
                Assert.Equal("A", row["Value"]);
                Assert.Equal("B", row["  Value  "]);
                Assert.Equal("C", row["Column3"]);

                using var document = ExcelDocument.Load(filePath);
                var sheet = document.GetSheet("Data");
                var headerMap = sheet.GetHeaderMap(options);
                Assert.Equal(1, headerMap["Value"]);
                Assert.Equal(2, headerMap["  Value  "]);
                Assert.Equal(3, headerMap["Column3"]);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Compatibility_Corpus_TypedReads_MapFriendlyHeadersWithoutLosingDuplicateColumns() {
            string filePath = Path.Combine(_directoryWithFiles, "CompatibilityCorpus.FriendlyTypedHeaders.xlsx");

            try {
                ExcelCompatibilityCorpusBuilder.CreateWorkbook(filePath, document => {
                    var sheet = document.AddSheet("Data");
                    sheet.Cell(1, 1).SetValue("First Name");
                    sheet.Cell(1, 2).SetValue("First Name");
                    sheet.Cell(1, 3).SetValue("Total Amount 2");
                    sheet.Cell(2, 1).SetValue("Alice");
                    sheet.Cell(2, 2).SetValue("Bob");
                    sheet.Cell(2, 3).SetValue(42);
                });

                using var reader = ExcelDocumentReader.Open(filePath);
                var typedFromReader = Assert.Single(reader.GetSheet("Data").ReadObjects<CompatibilityFriendlyHeaderRow>("A1:C2"));

                Assert.Equal("Alice", typedFromReader.FirstName);
                Assert.Equal("Bob", typedFromReader.FirstName_2);
                Assert.Equal(42, typedFromReader.TotalAmount2);

                using var document = ExcelDocument.Load(filePath);
                var sheet = document.GetSheet("Data");
                var typedFromSheet = Assert.Single(sheet.RowsAs<CompatibilityFriendlyHeaderRow>("A1:C2"));

                Assert.Equal("Alice", typedFromSheet.FirstName);
                Assert.Equal("Bob", typedFromSheet.FirstName_2);
                Assert.Equal(42, typedFromSheet.TotalAmount2);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Compatibility_Corpus_TypedReads_MapAttributeBasedHeaderAliases() {
            string filePath = Path.Combine(_directoryWithFiles, "CompatibilityCorpus.AttributedTypedHeaders.xlsx");

            try {
                ExcelCompatibilityCorpusBuilder.CreateWorkbook(filePath, document => {
                    var sheet = document.AddSheet("Data");
                    sheet.Cell(1, 1).SetValue("First Name");
                    sheet.Cell(1, 2).SetValue("Status Code");
                    sheet.Cell(1, 3).SetValue("Total %");
                    sheet.Cell(2, 1).SetValue("Alice");
                    sheet.Cell(2, 2).SetValue("OK");
                    sheet.Cell(2, 3).SetValue(97);
                });

                using var reader = ExcelDocumentReader.Open(filePath);
                var typedFromReader = Assert.Single(reader.GetSheet("Data").ReadObjects<CompatibilityAttributedHeaderRow>("A1:C2"));

                Assert.Equal("Alice", typedFromReader.GivenName);
                Assert.Equal("OK", typedFromReader.Status);
                Assert.Equal(97, typedFromReader.CompletionPercent);

                using var document = ExcelDocument.Load(filePath);
                var sheet = document.GetSheet("Data");
                var typedFromSheet = Assert.Single(sheet.RowsAs<CompatibilityAttributedHeaderRow>("A1:C2"));

                Assert.Equal("Alice", typedFromSheet.GivenName);
                Assert.Equal("OK", typedFromSheet.Status);
                Assert.Equal(97, typedFromSheet.CompletionPercent);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Theory]
        [InlineData("Inherited Total")]
        [InlineData("Legacy Total")]
        public void Compatibility_Corpus_TypedReads_MapInheritedAttributeAliases(string header) {
            string filePath = Path.Combine(_directoryWithFiles, "CompatibilityCorpus.InheritedTypedHeaders.xlsx");

            try {
                ExcelCompatibilityCorpusBuilder.CreateWorkbook(filePath, document => {
                    var sheet = document.AddSheet("Data");
                    sheet.Cell(1, 1).SetValue(header);
                    sheet.Cell(2, 1).SetValue(97);
                });

                using var reader = ExcelDocumentReader.Open(filePath);
                var typedFromReader = Assert.Single(
                    reader.GetSheet("Data").ReadObjects<CompatibilityInheritedHeaderRow>("A1:A2"));
                Assert.Equal(97, typedFromReader.Total);

                using var document = ExcelDocument.Load(filePath);
                var typedFromSheet = Assert.Single(
                    document.GetSheet("Data").RowsAs<CompatibilityInheritedHeaderRow>("A1:A2"));
                Assert.Equal(97, typedFromSheet.Total);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Compatibility_Corpus_CustomFormats_PreserveIntendedReaderTypesWithinOneWorkbook() {
            string filePath = Path.Combine(_directoryWithFiles, "CompatibilityCorpus.CustomFormats.xlsx");

            try {
                ExcelCompatibilityCorpusBuilder.CreateWorkbook(filePath, document => {
                    var sheet = document.AddSheet("Data");
                    sheet.Cell(1, 1).SetValue("QuotedDays");
                    sheet.Cell(1, 2).SetValue("Duration");
                    sheet.Cell(1, 3).SetValue("EscapedHours");
                    sheet.Cell(2, 1).SetValue(3d);
                    sheet.Cell(2, 2).SetValue(1.5d);
                    sheet.Cell(2, 3).SetValue(7d);
                    sheet.ColumnStyleByHeader("QuotedDays").NumberFormat("0 \"days\"");
                    sheet.ColumnStyleByHeader("Duration").NumberFormat("[h]:mm");
                    sheet.ColumnStyleByHeader("EscapedHours").NumberFormat("0\\h");
                });

                using var reader = ExcelDocumentReader.Open(filePath);
                var row = Assert.Single(reader.GetSheet("Data").ReadObjects("A1:C2"));

                Assert.IsType<double>(row["QuotedDays"]);
                Assert.Equal(3d, (double)row["QuotedDays"]!);

                Assert.IsType<DateTime>(row["Duration"]);
                Assert.Equal(DateTime.FromOADate(1.5d), (DateTime)row["Duration"]!);

                Assert.IsType<double>(row["EscapedHours"]);
                Assert.Equal(7d, (double)row["EscapedHours"]!);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Compatibility_Corpus_NonCanonicalRowsAndMalformedReferences_ReadSafely() {
            string filePath = Path.Combine(_directoryWithFiles, "CompatibilityCorpus.NonCanonicalRows.xlsx");

            try {
                ExcelCompatibilityCorpusBuilder.CreateWorkbook(filePath, document => {
                    var sheet = document.AddSheet("Data");
                    sheet.Cell(1, 1).SetValue("Name");
                    sheet.Cell(1, 2).SetValue("Score");
                    sheet.Cell(2, 1).SetValue("First");
                    sheet.Cell(2, 2).SetValue(10);
                    sheet.Cell(3, 1).SetValue("Second");
                    sheet.Cell(3, 2).SetValue(20);
                });

                ExcelCompatibilityCorpusBuilder.RewriteWorksheetXml(filePath, "xl/worksheets/sheet1.xml", worksheet => {
                    XNamespace ns = worksheet.Root!.Name.Namespace;
                    var sheetData = worksheet.Root.Element(ns + "sheetData")!;
                    var rows = sheetData.Elements(ns + "row").ToList();
                    sheetData.RemoveNodes();

                    sheetData.Add(rows[0]);
                    sheetData.Add(rows[2]);
                    sheetData.Add(rows[1]);

                    var malformedCell = new XElement(ns + "c",
                        new XAttribute("r", "TOTAL"),
                        new XAttribute("t", "str"),
                        new XElement(ns + "v", "Bad"));
                    rows[1].Add(malformedCell);
                    return worksheet;
                });

                using var reader = ExcelDocumentReader.Open(filePath);
                object?[,] values = reader.GetSheet("Data").ReadRange("A1:B3", ExcelExecutionMode.Sequential);
                var objects = reader.GetSheet("Data").ReadObjects("A1:B3", ExcelExecutionMode.Sequential).ToList();
                var typedObjects = reader.GetSheet("Data").ReadObjects<CompatibilityScoreRow>("A1:B3", ExcelExecutionMode.Sequential).ToList();

                Assert.Equal("First", values[1, 0]);
                Assert.Equal(10d, values[1, 1]);
                Assert.Equal("Second", values[2, 0]);
                Assert.Equal(20d, values[2, 1]);
                Assert.Equal(2, objects.Count);
                Assert.Equal("First", objects[0]["Name"]);
                Assert.Equal("Second", objects[1]["Name"]);
                Assert.Equal(2, typedObjects.Count);
                Assert.Equal("First", typedObjects[0].Name);
                Assert.Equal(10, typedObjects[0].Score);
                Assert.Equal("Second", typedObjects[1].Name);
                Assert.Equal(20, typedObjects[1].Score);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Compatibility_Corpus_TextHeavyCellValues_SaveAsValidWorkbook() {
            string filePath = Path.Combine(_directoryWithFiles, "CompatibilityCorpus.TextHeavyCellValues.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddSheet("Strings");
                    var cells = Enumerable.Range(1, 300).SelectMany(row => new[] {
                        (row, 1, (object)("Repeated " + row % 10)),
                        (row, 2, (object)("Distinct " + row)),
                        (row, 3, (object)("Long " + new string((char)('A' + row % 26), 80)))
                    }).ToArray();
                    sheet.SetValues(cells, ExcelExecutionMode.Parallel);
                    sheet.AutoFitColumns();
                    document.Save();

                    var validationErrors = document.ValidateDocument();
                    Assert.Empty(validationErrors);
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                object?[,] values = reader.GetSheet("Strings").ReadRange("A1:C300");
                Assert.Equal(900, values.Length);
                Assert.Equal("Repeated 1", values[0, 0]);
                Assert.Equal("Distinct 300", values[299, 1]);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

    }
}
