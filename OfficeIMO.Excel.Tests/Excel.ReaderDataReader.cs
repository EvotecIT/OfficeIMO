using System;
using System.Data;
using System.Data.Common;
using System.IO;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Reader_ReadRangeAsDataReader_ExposesTypedSchemaAndValues() {
            var expectedDate = new DateTime(2026, 7, 8);
            using var memory = new MemoryStream();

            using (var document = ExcelDocument.Create(memory)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Id");
                sheet.CellValue(1, 2, "Amount");
                sheet.CellValue(1, 3, "Created");
                sheet.CellValue(1, 4, "Active");
                sheet.CellValue(1, 5, "Note");
                sheet.CellValue(2, 1, 1);
                sheet.CellValue(2, 2, 12.5d);
                sheet.CellValue(2, 3, expectedDate);
                sheet.CellValue(2, 4, true);
                sheet.CellValue(2, 5, "Alpha");
                sheet.CellValue(3, 1, 2);
                sheet.CellValue(3, 4, false);
            }

            using var reader = ExcelDocumentReader.Open(memory.ToArray());
            using var dataReader = reader.GetSheet("Data").ReadRangeAsDataReader("A1:E3", chunkRows: 1, schemaSampleRows: 2);

            Assert.True(Assert.IsAssignableFrom<DbDataReader>(dataReader).HasRows);
            Assert.Equal(5, dataReader.FieldCount);
            Assert.Equal(typeof(double), dataReader.GetFieldType(dataReader.GetOrdinal("Id")));
            Assert.Equal(typeof(double), dataReader.GetFieldType(dataReader.GetOrdinal("Amount")));
            Assert.Equal(typeof(DateTime), dataReader.GetFieldType(dataReader.GetOrdinal("Created")));
            Assert.Equal(typeof(bool), dataReader.GetFieldType(dataReader.GetOrdinal("Active")));
            Assert.Equal(typeof(string), dataReader.GetFieldType(dataReader.GetOrdinal("Note")));

            Assert.True(dataReader.Read());
            Assert.Equal(1d, dataReader.GetDouble(dataReader.GetOrdinal("Id")));
            Assert.Equal(12.5d, dataReader.GetDouble(dataReader.GetOrdinal("Amount")));
            Assert.Equal(expectedDate, dataReader.GetDateTime(dataReader.GetOrdinal("Created")));
            Assert.True(dataReader.GetBoolean(dataReader.GetOrdinal("Active")));
            Assert.Equal("Alpha", dataReader.GetString(dataReader.GetOrdinal("Note")));

            Assert.True(dataReader.Read());
            Assert.Equal(2d, dataReader.GetDouble(dataReader.GetOrdinal("Id")));
            Assert.True(dataReader.IsDBNull(dataReader.GetOrdinal("Amount")));
            Assert.False(dataReader.GetBoolean(dataReader.GetOrdinal("Active")));
            Assert.False(dataReader.Read());
        }

        [Fact]
        public void Reader_ReadRangeAsDataReader_LoadsDataTableWithDisambiguatedHeaders() {
            using var memory = new MemoryStream();

            using (var document = ExcelDocument.Create(memory)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Status");
                sheet.CellValue(1, 2, "Status");
                sheet.CellValue(1, 3, "");
                sheet.CellValue(2, 1, "OK");
                sheet.CellValue(2, 2, "Warning");
                sheet.CellValue(2, 3, "Generated");
            }

            using var reader = ExcelDocumentReader.Open(memory.ToArray());
            using var dataReader = reader.GetSheet("Data").ReadRangeAsDataReader("A1:C2");
            var table = new DataTable();
            table.Load(dataReader);

            Assert.Equal(new[] { "Status", "Status_2", "Column3" }, table.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToArray());
            DataRow row = Assert.Single(table.Rows.Cast<DataRow>());
            Assert.Equal("OK", row["Status"]);
            Assert.Equal("Warning", row["Status_2"]);
            Assert.Equal("Generated", row["Column3"]);
        }

        [Fact]
        public void Reader_ReadRangeAsDataReader_WithoutHeadersPreservesBlankRowsInsideRange() {
            using var memory = new MemoryStream();

            using (var document = ExcelDocument.Create(memory)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Alpha");
                sheet.CellValue(1, 3, 10);
                sheet.CellValue(3, 2, "Beta");
            }

            using var reader = ExcelDocumentReader.Open(memory.ToArray());
            using var dataReader = reader.GetSheet("Data").ReadRangeAsDataReader("A1:C3", headersInFirstRow: false, chunkRows: 1, schemaSampleRows: 1);

            Assert.Equal("Column1", dataReader.GetName(0));
            Assert.Equal("Column2", dataReader.GetName(1));
            Assert.Equal("Column3", dataReader.GetName(2));

            Assert.True(dataReader.Read());
            Assert.Equal("Alpha", dataReader.GetString(0));
            Assert.True(dataReader.IsDBNull(1));
            Assert.Equal(10d, dataReader.GetDouble(2));

            Assert.True(dataReader.Read());
            Assert.True(dataReader.IsDBNull(0));
            Assert.True(dataReader.IsDBNull(1));
            Assert.True(dataReader.IsDBNull(2));

            Assert.True(dataReader.Read());
            Assert.True(dataReader.IsDBNull(0));
            Assert.Equal("Beta", dataReader.GetString(1));
            Assert.True(dataReader.IsDBNull(2));
            Assert.False(dataReader.Read());
        }

        [Fact]
        public void Reader_ReadRangeAsDataReader_ReturnsSampledAndRemainingRows() {
            using var memory = new MemoryStream();

            using (var document = ExcelDocument.Create(memory)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Id");
                sheet.CellValue(1, 2, "Name");
                sheet.CellValue(2, 1, 1);
                sheet.CellValue(2, 2, "Alpha");
                sheet.CellValue(3, 1, 2);
                sheet.CellValue(3, 2, "Beta");
                sheet.CellValue(4, 1, 3);
                sheet.CellValue(4, 2, "Gamma");
            }

            using var reader = ExcelDocumentReader.Open(memory.ToArray());
            using var dataReader = reader.GetSheet("Data").ReadRangeAsDataReader("A1:B4", chunkRows: 1, schemaSampleRows: 2);

            Assert.True(dataReader.Read());
            Assert.Equal(1d, dataReader.GetDouble(0));
            Assert.Equal("Alpha", dataReader.GetString(1));

            Assert.True(dataReader.Read());
            Assert.Equal(2d, dataReader.GetDouble(0));
            Assert.Equal("Beta", dataReader.GetString(1));

            Assert.True(dataReader.Read());
            Assert.Equal(3d, dataReader.GetDouble(0));
            Assert.Equal("Gamma", dataReader.GetString(1));
            Assert.False(dataReader.Read());
        }
    }
}
