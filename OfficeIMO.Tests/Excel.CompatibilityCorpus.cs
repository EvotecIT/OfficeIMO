using System.Data;
using System.ComponentModel;
using System.Runtime.Serialization;
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

        [Fact]
        public void Compatibility_Corpus_RecoverableContentTypes_StayInSyncAcrossPathAndStreamReaders() {
            string filePath = Path.Combine(_directoryWithFiles, "CompatibilityCorpus.RecoverableContentTypes.xlsx");

            try {
                ExcelCompatibilityCorpusBuilder.CreateWorkbook(filePath, document => {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Region");
                    sheet.CellValue(1, 2, "Amount");
                    sheet.CellValue(1, 3, "Status");
                    sheet.CellValue(2, 1, "North");
                    sheet.CellValue(2, 2, 10d);
                    sheet.CellValue(2, 3, "Open");
                    sheet.CellValue(3, 1, "South");
                    sheet.CellValue(3, 2, 20d);
                    sheet.CellValue(3, 3, "Closed");
                }, rewriteAppContentTypeToXml: true);

                using var pathReader = ExcelDocumentReader.Open(filePath);
                using var streamReader = ExcelDocumentReader.Open(new MemoryStream(File.ReadAllBytes(filePath), writable: false));

                var pathRows = pathReader.GetSheet("Data").ReadObjects("A1:C3").ToList();
                var streamRows = streamReader.GetSheet("Data").ReadObjects("A1:C3").ToList();

                Assert.Equal(pathReader.GetSheetNames(), streamReader.GetSheetNames());
                Assert.Equal(pathRows.Count, streamRows.Count);
                Assert.Equal(pathRows[0]["Region"], streamRows[0]["Region"]);
                Assert.Equal(pathRows[0]["Amount"], streamRows[0]["Amount"]);
                Assert.Equal(pathRows[1]["Status"], streamRows[1]["Status"]);

                using var pathDocument = ExcelDocument.Load(filePath);
                using var streamDocument = ExcelDocument.Load(new MemoryStream(File.ReadAllBytes(filePath), writable: false));

                var pathEditable = pathDocument.GetSheet("Data").RowsObjects("A1:C3").ToList();
                var streamEditable = streamDocument.GetSheet("Data").RowsObjects("A1:C3").ToList();

                Assert.Equal(pathEditable[0]["Region"].Value, streamEditable[0]["Region"].Value);
                Assert.Equal(pathEditable[0]["Amount"].Value, streamEditable[0]["Amount"].Value);
                Assert.Equal(pathEditable[1]["Status"].Value, streamEditable[1]["Status"].Value);
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
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Value");
                    sheet.CellValue(1, 2, "  Value  ");
                    sheet.CellValue(1, 3, "");
                    sheet.CellValue(1, 4, "value");
                    sheet.CellValue(1, 5, "Value_2");
                    sheet.CellValue(1, 6, "Value_2");
                    sheet.CellValue(2, 1, "A");
                    sheet.CellValue(2, 2, "B");
                    sheet.CellValue(2, 3, "C");
                    sheet.CellValue(2, 4, "D");
                    sheet.CellValue(2, 5, "E");
                    sheet.CellValue(2, 6, "F");
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
                var editable = Assert.Single(sheet.RowsObjects("A1:F2"));
                var typed = Assert.Single(sheet.RowsAs<CompatibilityMessyHeaderRow>("A1:F2"));

                Assert.Equal(new[] { "Value", "Value_2", "Column3", "value_3", "Value_2_2", "Value_2_3" }, headerMap.Keys.ToArray());
                Assert.Equal("A", editable["Value"].Value);
                Assert.Equal("F", editable["Value_2_3"].Value);
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
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "");
                    sheet.CellValue(1, 2, "Column1");
                    sheet.CellValue(1, 3, "");
                    sheet.CellValue(2, 1, "GeneratedLeft");
                    sheet.CellValue(2, 2, "Explicit");
                    sheet.CellValue(2, 3, "GeneratedRight");
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
                var editable = Assert.Single(sheet.RowsObjects("A1:C2"));

                Assert.Equal(1, headerMap["Column1_2"]);
                Assert.Equal(2, headerMap["Column1"]);
                Assert.Equal(3, headerMap["Column3"]);
                Assert.Equal("GeneratedLeft", editable["Column1_2"].Value);
                Assert.Equal("Explicit", editable["Column1"].Value);
                Assert.Equal("GeneratedRight", editable["Column3"].Value);
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
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "");
                    sheet.CellValue(1, 2, "Column1");
                    sheet.CellValue(1, 3, "Column1_2");
                    sheet.CellValue(2, 1, "Generated");
                    sheet.CellValue(2, 2, "ExplicitBase");
                    sheet.CellValue(2, 3, "ExplicitSuffix");
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
                var editable = Assert.Single(sheet.RowsObjects("A1:C2"));

                Assert.Equal(1, headerMap["Column1_3"]);
                Assert.Equal(2, headerMap["Column1"]);
                Assert.Equal(3, headerMap["Column1_2"]);
                Assert.Equal("Generated", editable["Column1_3"].Value);
                Assert.Equal("ExplicitBase", editable["Column1"].Value);
                Assert.Equal("ExplicitSuffix", editable["Column1_2"].Value);
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
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(2, 2, "");
                    sheet.CellValue(2, 3, "Column1");
                    sheet.CellValue(2, 4, "Column1_2");
                    sheet.CellValue(3, 2, "Generated");
                    sheet.CellValue(3, 3, "ExplicitBase");
                    sheet.CellValue(3, 4, "ExplicitSuffix");
                });

                using var reader = ExcelDocumentReader.Open(filePath);
                var row = Assert.Single(reader.GetSheet("Data").ReadObjects("B2:D3"));
                DataTable table = reader.GetSheet("Data").ReadRangeAsDataTable("B2:D3");

                Assert.Equal(new[] { "Column1_3", "Column1", "Column1_2" },
                    table.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToArray());
                Assert.Equal("Generated", row["Column1_3"]);
                Assert.Equal("ExplicitBase", row["Column1"]);
                Assert.Equal("ExplicitSuffix", row["Column1_2"]);

                using var document = ExcelDocument.Load(filePath);
                var sheet = document.GetSheet("Data");
                var editable = Assert.Single(sheet.RowsObjects("B2:D3"));

                Assert.Equal("Generated", editable["Column1_3"].Value);
                Assert.Equal("ExplicitBase", editable["Column1"].Value);
                Assert.Equal("ExplicitSuffix", editable["Column1_2"].Value);
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
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Value");
                    sheet.CellValue(1, 2, "  Value  ");
                    sheet.CellValue(1, 3, "");
                    sheet.CellValue(2, 1, "A");
                    sheet.CellValue(2, 2, "B");
                    sheet.CellValue(2, 3, "C");
                });

                var options = new ExcelReadOptions { NormalizeHeaders = false };

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
                var editable = Assert.Single(sheet.RowsObjects("A1:C2", options));

                Assert.Equal(1, headerMap["Value"]);
                Assert.Equal(2, headerMap["  Value  "]);
                Assert.Equal(3, headerMap["Column3"]);
                Assert.Equal("A", editable["Value"].Value);
                Assert.Equal("B", editable["  Value  "].Value);
                Assert.Equal("C", editable["Column3"].Value);
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
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "First Name");
                    sheet.CellValue(1, 2, "First Name");
                    sheet.CellValue(1, 3, "Total Amount 2");
                    sheet.CellValue(2, 1, "Alice");
                    sheet.CellValue(2, 2, "Bob");
                    sheet.CellValue(2, 3, 42);
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
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "First Name");
                    sheet.CellValue(1, 2, "Status Code");
                    sheet.CellValue(1, 3, "Total %");
                    sheet.CellValue(2, 1, "Alice");
                    sheet.CellValue(2, 2, "OK");
                    sheet.CellValue(2, 3, 97);
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

        [Fact]
        public void Compatibility_Corpus_CustomFormats_PreserveIntendedReaderTypesWithinOneWorkbook() {
            string filePath = Path.Combine(_directoryWithFiles, "CompatibilityCorpus.CustomFormats.xlsx");

            try {
                ExcelCompatibilityCorpusBuilder.CreateWorkbook(filePath, document => {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "QuotedDays");
                    sheet.CellValue(1, 2, "Duration");
                    sheet.CellValue(1, 3, "EscapedHours");
                    sheet.CellValue(2, 1, 3d);
                    sheet.CellValue(2, 2, 1.5d);
                    sheet.CellValue(2, 3, 7d);
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
    }
}
