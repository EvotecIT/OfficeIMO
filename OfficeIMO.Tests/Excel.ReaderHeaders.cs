using System;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        private sealed class DuplicateHeaderRow {
            public string? Value { get; set; }
            public string? Value_2 { get; set; }
        }

        private sealed class FriendlyHeaderRow {
            public string? FirstName { get; set; }
            public string? FirstName_2 { get; set; }
            public int TotalAmount2 { get; set; }
        }

        private sealed class ExactFriendlyPrecedenceRow {
            public string? FirstName { get; set; }
        }

        private sealed class AttributedHeaderRow {
            [DisplayName("First Name")]
            public string? GivenName { get; set; }

            [DataMember(Name = "Status Code")]
            public string? Status { get; set; }

            [ExcelColumn("Total %", "Total Percent")]
            public int CompletionPercent { get; set; }
        }

        private sealed class AmbiguousAttributedHeaderRow {
            [ExcelColumn("Status")]
            public string? PrimaryStatus { get; set; }

            [DisplayName("Status")]
            public string? SecondaryStatus { get; set; }
        }

        private sealed class StrictMappedRow {
            public string? Name { get; set; }
        }

        [Fact]
        public void Reader_ReadObjects_DisambiguatesDuplicateAndNormalizedHeaders() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderDuplicateHeaders.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Value");
                    sheet.CellValue(1, 2, "  Value  ");
                    sheet.CellValue(1, 3, "");
                    sheet.CellValue(2, 1, "Alpha");
                    sheet.CellValue(2, 2, "Beta");
                    sheet.CellValue(2, 3, "Gamma");
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                var row = Assert.Single(reader.GetSheet("Data").ReadObjects("A1:C2"));

                Assert.Equal("Alpha", row["Value"]);
                Assert.Equal("Beta", row["Value_2"]);
                Assert.Equal("Gamma", row["Column3"]);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_ReadRangeAsDataTable_DisambiguatesDuplicateHeaders() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderDuplicateHeadersDataTable.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Status");
                    sheet.CellValue(1, 2, "Status");
                    sheet.CellValue(1, 3, " status ");
                    sheet.CellValue(2, 1, "OK");
                    sheet.CellValue(2, 2, "Warning");
                    sheet.CellValue(2, 3, "Error");
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                DataTable table = reader.GetSheet("Data").ReadRangeAsDataTable("A1:C2");

                Assert.Equal(new[] { "Status", "Status_2", "status_3" }, table.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToArray());
                Assert.Equal("OK", table.Rows[0]["Status"]);
                Assert.Equal("Warning", table.Rows[0]["Status_2"]);
                Assert.Equal("Error", table.Rows[0]["status_3"]);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Sheet_ReadHelpers_ExposeDisambiguatedHeadersConsistently() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderDuplicateHeadersEditable.xlsx");

            try {
                using (var createdDocument = ExcelDocument.Create(filePath)) {
                    var createdSheet = createdDocument.AddWorkSheet("Data");
                    createdSheet.CellValue(1, 1, "Value");
                    createdSheet.CellValue(1, 2, "Value");
                    createdSheet.CellValue(2, 1, "Left");
                    createdSheet.CellValue(2, 2, "Right");
                    createdDocument.Save();
                }

                using var document = ExcelDocument.Load(filePath);
                var sheet = document.GetSheet("Data");

                var map = sheet.GetHeaderMap();
                Assert.Equal(1, map["Value"]);
                Assert.Equal(2, map["Value_2"]);

                var editable = Assert.Single(sheet.RowsObjects("A1:B2"));
                Assert.Equal("Left", editable["Value"].Value);
                Assert.Equal("Right", editable["Value_2"].Value);

                var typed = Assert.Single(sheet.RowsAs<DuplicateHeaderRow>("A1:B2"));
                Assert.Equal("Left", typed.Value);
                Assert.Equal("Right", typed.Value_2);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_BlankGeneratedHeaders_DoNotStealExplicitColumnNames() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderBlankGeneratedHeaders.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "");
                    sheet.CellValue(1, 2, "Column1");
                    sheet.CellValue(2, 1, "Generated");
                    sheet.CellValue(2, 2, "Explicit");
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                var row = Assert.Single(reader.GetSheet("Data").ReadObjects("A1:B2"));

                Assert.Equal("Generated", row["Column1_2"]);
                Assert.Equal("Explicit", row["Column1"]);

                DataTable table = reader.GetSheet("Data").ReadRangeAsDataTable("A1:B2");
                Assert.Equal(new[] { "Column1_2", "Column1" }, table.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToArray());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_ExplicitSuffixedHeaders_RemainStableWhenBaseHeaderRepeats() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderExplicitSuffixedHeaders.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Value");
                    sheet.CellValue(1, 2, "Value_2");
                    sheet.CellValue(1, 3, "Value");
                    sheet.CellValue(2, 1, "Alpha");
                    sheet.CellValue(2, 2, "Beta");
                    sheet.CellValue(2, 3, "Gamma");
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                var row = Assert.Single(reader.GetSheet("Data").ReadObjects("A1:C2"));

                Assert.Equal("Alpha", row["Value"]);
                Assert.Equal("Beta", row["Value_2"]);
                Assert.Equal("Gamma", row["Value_3"]);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_BlankGeneratedHeaders_DoNotStealExplicitSuffixedNames() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderBlankGeneratedHeadersReservedSuffix.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "");
                    sheet.CellValue(1, 2, "Column1");
                    sheet.CellValue(1, 3, "Column1_2");
                    sheet.CellValue(2, 1, "Generated");
                    sheet.CellValue(2, 2, "ExplicitBase");
                    sheet.CellValue(2, 3, "ExplicitSuffix");
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                var row = Assert.Single(reader.GetSheet("Data").ReadObjects("A1:C2"));
                DataTable table = reader.GetSheet("Data").ReadRangeAsDataTable("A1:C2");

                Assert.Equal(new[] { "Column1_3", "Column1", "Column1_2" }, table.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToArray());
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
        public void Sheet_HeaderMapCache_RebuildsAfterHeaderRenameWithinSameUsedRange() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderHeaderCacheRename.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Status");
                    sheet.CellValue(1, 2, "Value");
                    sheet.CellValue(2, 1, "Open");
                    sheet.CellValue(2, 2, "10");
                    document.Save();
                }

                using var loadedDocument = ExcelDocument.Load(filePath);
                var loadedSheet = loadedDocument.GetSheet("Data");

                var initialMap = loadedSheet.GetHeaderMap();
                Assert.Equal(1, initialMap["Status"]);

                loadedSheet.CellValue(1, 1, "State");

                var refreshedMap = loadedSheet.GetHeaderMap();
                Assert.False(refreshedMap.ContainsKey("Status"));
                Assert.Equal(1, refreshedMap["State"]);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Sheet_HeaderMapCache_RebuildsWhenUsedRangeShiftsAfterWrite() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderHeaderCacheUsedRangeShift.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(2, 2, "Region");
                    sheet.CellValue(2, 3, "Amount");
                    sheet.CellValue(3, 2, "North");
                    sheet.CellValue(3, 3, 10);
                    document.Save();
                }

                using var loadedDocument = ExcelDocument.Load(filePath);
                var loadedSheet = loadedDocument.GetSheet("Data");

                var initialMap = loadedSheet.GetHeaderMap();
                Assert.Equal(2, initialMap["Region"]);
                Assert.Equal(3, initialMap["Amount"]);

                loadedSheet.CellValue(1, 1, "Id");
                loadedSheet.CellValue(1, 2, "Region");
                loadedSheet.CellValue(1, 3, "Amount");

                var refreshedMap = loadedSheet.GetHeaderMap();
                Assert.Equal(1, refreshedMap["Id"]);
                Assert.Equal(2, refreshedMap["Region"]);
                Assert.Equal(3, refreshedMap["Amount"]);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Sheet_HeaderMapCache_RebuildsAfterParallelCellValuesOverwriteHeaders() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderHeaderCacheParallelCellValues.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "OldName");
                    sheet.CellValue(1, 2, "OldValue");
                    sheet.CellValue(2, 1, "Alpha");
                    sheet.CellValue(2, 2, "1");
                    document.Save();
                }

                using var loadedDocument = ExcelDocument.Load(filePath);
                var loadedSheet = loadedDocument.GetSheet("Data");

                var initialMap = loadedSheet.GetHeaderMap();
                Assert.Equal(1, initialMap["OldName"]);
                Assert.Equal(2, initialMap["OldValue"]);

                loadedSheet.CellValues(new[] {
                    (1, 1, (object)"Name"),
                    (1, 2, (object)"Value")
                }, ExecutionMode.Parallel);

                var refreshedMap = loadedSheet.GetHeaderMap();
                Assert.False(refreshedMap.ContainsKey("OldName"));
                Assert.False(refreshedMap.ContainsKey("OldValue"));
                Assert.Equal(1, refreshedMap["Name"]);
                Assert.Equal(2, refreshedMap["Value"]);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Sheet_HeaderMapCache_RebuildsAfterParallelInsertDataTableOverwriteHeaders() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderHeaderCacheParallelDataTable.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "OldStatus");
                    sheet.CellValue(1, 2, "OldOwner");
                    sheet.CellValue(2, 1, "Open");
                    sheet.CellValue(2, 2, "Alice");
                    document.Save();
                }

                using var loadedDocument = ExcelDocument.Load(filePath);
                var loadedSheet = loadedDocument.GetSheet("Data");

                var initialMap = loadedSheet.GetHeaderMap();
                Assert.Equal(1, initialMap["OldStatus"]);
                Assert.Equal(2, initialMap["OldOwner"]);

                var table = new DataTable();
                table.Columns.Add("Status");
                table.Columns.Add("Owner");
                table.Rows.Add("Closed", "Bob");

                loadedSheet.InsertDataTable(table, startRow: 1, startColumn: 1, includeHeaders: true, mode: ExecutionMode.Parallel);

                var refreshedMap = loadedSheet.GetHeaderMap();
                Assert.False(refreshedMap.ContainsKey("OldStatus"));
                Assert.False(refreshedMap.ContainsKey("OldOwner"));
                Assert.Equal(1, refreshedMap["Status"]);
                Assert.Equal(2, refreshedMap["Owner"]);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_ShiftedRange_DisambiguatesBlankAndExplicitHeadersLocally() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderShiftedRangeHeaders.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(2, 2, "");
                    sheet.CellValue(2, 3, "Column1");
                    sheet.CellValue(2, 4, "Column1_2");
                    sheet.CellValue(3, 2, "Generated");
                    sheet.CellValue(3, 3, "ExplicitBase");
                    sheet.CellValue(3, 4, "ExplicitSuffix");
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                var row = Assert.Single(reader.GetSheet("Data").ReadObjects("B2:D3"));
                DataTable table = reader.GetSheet("Data").ReadRangeAsDataTable("B2:D3");

                Assert.Equal(new[] { "Column1_3", "Column1", "Column1_2" }, table.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToArray());
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
        public void Reader_NormalizeHeadersFalse_PreservesWhitespaceDistinctHeadersAcrossReadSurfaces() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderNormalizeHeadersFalse.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Value");
                    sheet.CellValue(1, 2, "  Value  ");
                    sheet.CellValue(1, 3, "");
                    sheet.CellValue(2, 1, "Alpha");
                    sheet.CellValue(2, 2, "Beta");
                    sheet.CellValue(2, 3, "Gamma");
                    document.Save();
                }

                var options = new ExcelReadOptions { NormalizeHeaders = false };

                using var reader = ExcelDocumentReader.Open(filePath, options);
                var row = Assert.Single(reader.GetSheet("Data").ReadObjects("A1:C2"));
                DataTable table = reader.GetSheet("Data").ReadRangeAsDataTable("A1:C2");

                Assert.Equal(new[] { "Value", "  Value  ", "Column3" }, table.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToArray());
                Assert.Equal("Alpha", row["Value"]);
                Assert.Equal("Beta", row["  Value  "]);
                Assert.Equal("Gamma", row["Column3"]);

                using var loadedDocument = ExcelDocument.Load(filePath);
                var loadedSheet = loadedDocument.GetSheet("Data");
                var headerMap = loadedSheet.GetHeaderMap(options);
                var editable = Assert.Single(loadedSheet.RowsObjects("A1:C2", options));

                Assert.Equal(1, headerMap["Value"]);
                Assert.Equal(2, headerMap["  Value  "]);
                Assert.Equal(3, headerMap["Column3"]);
                Assert.Equal("Alpha", editable["Value"].Value);
                Assert.Equal("Beta", editable["  Value  "].Value);
                Assert.Equal("Gamma", editable["Column3"].Value);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_TypedObjects_MapFriendlyDuplicateHeadersToDisambiguatedProperties() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderFriendlyTypedHeaders.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "First Name");
                    sheet.CellValue(1, 2, "First Name");
                    sheet.CellValue(1, 3, "Total Amount 2");
                    sheet.CellValue(2, 1, "Alice");
                    sheet.CellValue(2, 2, "Bob");
                    sheet.CellValue(2, 3, 42);
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                var typedFromReader = Assert.Single(reader.GetSheet("Data").ReadObjects<FriendlyHeaderRow>("A1:C2"));

                Assert.Equal("Alice", typedFromReader.FirstName);
                Assert.Equal("Bob", typedFromReader.FirstName_2);
                Assert.Equal(42, typedFromReader.TotalAmount2);

                using var loadedDocument = ExcelDocument.Load(filePath);
                var typedFromSheet = Assert.Single(loadedDocument.GetSheet("Data").RowsAs<FriendlyHeaderRow>("A1:C2"));

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
        public void Reader_TypedObjects_PreferExactPropertyHeadersOverEarlierFriendlyMatches() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderFriendlyTypedHeadersExactPrecedence.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "First Name");
                    sheet.CellValue(1, 2, "FirstName");
                    sheet.CellValue(2, 1, "AliasValue");
                    sheet.CellValue(2, 2, "ExactValue");
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                var typed = Assert.Single(reader.GetSheet("Data").ReadObjects<ExactFriendlyPrecedenceRow>("A1:B2"));

                Assert.Equal("ExactValue", typed.FirstName);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_TypedObjects_MapAttributeBasedHeaderAliases() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderAttributedTypedHeaders.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "First Name");
                    sheet.CellValue(1, 2, "Status Code");
                    sheet.CellValue(1, 3, "Total %");
                    sheet.CellValue(2, 1, "Alice");
                    sheet.CellValue(2, 2, "OK");
                    sheet.CellValue(2, 3, 97);
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                var typedFromReader = Assert.Single(reader.GetSheet("Data").ReadObjects<AttributedHeaderRow>("A1:C2"));

                Assert.Equal("Alice", typedFromReader.GivenName);
                Assert.Equal("OK", typedFromReader.Status);
                Assert.Equal(97, typedFromReader.CompletionPercent);

                using var loadedDocument = ExcelDocument.Load(filePath);
                var typedFromSheet = Assert.Single(loadedDocument.GetSheet("Data").RowsAs<AttributedHeaderRow>("A1:C2"));

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
        public void Reader_TypedObjects_ReportAmbiguousAliasDiagnostics() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderAmbiguousTypedHeaders.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Status");
                    sheet.CellValue(2, 1, "Open");
                    document.Save();
                }

                var diagnostics = new List<string>();
                var options = new ExcelReadOptions();
                options.Execution.OnInfo = diagnostics.Add;

                using var reader = ExcelDocumentReader.Open(filePath, options);
                var typed = Assert.Single(reader.GetSheet("Data").ReadObjects<AmbiguousAttributedHeaderRow>("A1:A2"));

                Assert.Null(typed.PrimaryStatus);
                Assert.Null(typed.SecondaryStatus);
                Assert.Contains(diagnostics, message => message.Contains("TypedRead AmbiguousMapping", StringComparison.Ordinal)
                    && message.Contains("AmbiguousAttributedHeaderRow", StringComparison.Ordinal)
                    && message.Contains("Status", StringComparison.Ordinal));
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_TypedObjects_StrictMapping_ThrowsOnUnmappedHeaders() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderStrictTypedHeaders.Unmapped.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Name");
                    sheet.CellValue(1, 2, "Status");
                    sheet.CellValue(2, 1, "Alice");
                    sheet.CellValue(2, 2, "Open");
                    document.Save();
                }

                var options = new ExcelReadOptions { StrictTypedMapping = true };
                using var reader = ExcelDocumentReader.Open(filePath, options);
                var exception = Assert.Throws<InvalidOperationException>(() => reader.GetSheet("Data").ReadObjects<StrictMappedRow>("A1:B2").ToList());

                Assert.Contains("strict", exception.Message, StringComparison.OrdinalIgnoreCase);
                Assert.Contains("UnmappedHeader", exception.Message, StringComparison.Ordinal);
                Assert.Contains("Status", exception.Message, StringComparison.Ordinal);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_TypedObjects_StrictMapping_ThrowsOnAmbiguousAliases() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderStrictTypedHeaders.Ambiguous.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Status");
                    sheet.CellValue(2, 1, "Open");
                    document.Save();
                }

                var options = new ExcelReadOptions { StrictTypedMapping = true };
                using var reader = ExcelDocumentReader.Open(filePath, options);
                var exception = Assert.Throws<InvalidOperationException>(() => reader.GetSheet("Data").ReadObjects<AmbiguousAttributedHeaderRow>("A1:A2").ToList());

                Assert.Contains("strict", exception.Message, StringComparison.OrdinalIgnoreCase);
                Assert.Contains("AmbiguousMapping", exception.Message, StringComparison.Ordinal);
                Assert.Contains("Status", exception.Message, StringComparison.Ordinal);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }
    }
}
