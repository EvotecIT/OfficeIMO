using System;

using System.Collections;

using System.Collections.Generic;

using System.Data;

using System.Globalization;

using System.IO;

using System.IO.Compression;

using System.Linq;

using System.Reflection;

using System.Text;

using System.Threading;

using System.Threading.Tasks;

using System.Xml;

using DocumentFormat.OpenXml;

using DocumentFormat.OpenXml.Packaging;

using DocumentFormat.OpenXml.Spreadsheet;

using DocumentFormat.OpenXml.Validation;

using OfficeIMO.Excel;

using Xunit;

using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

using Vt = DocumentFormat.OpenXml.VariantTypes;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void StreamingDataTable_SerializesNonNumericFormattableValuesAsText()
        {
            using var memory = new MemoryStream();
            var id = Guid.Parse("89f22c99-1d51-4de5-b3b4-b20c4a60164f");
            var dataSet = new DataSet();
            var table = new DataTable("Items");
            table.Columns.Add("Id", typeof(Guid));
            table.Rows.Add(id);
            dataSet.Tables.Add(table);
            using (var writer = ExcelDocument.CreateWriter(memory, leaveOpen: true))
            {
                foreach (DataTable source in dataSet.Tables)
                    writer.WriteTable(source.TableName, source);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var cell = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet.Descendants<Cell>().Single(c => c.CellReference?.Value == "A2");
            Assert.Equal(CellValues.InlineString, cell.DataType!.Value);
            Assert.Equal(id.ToString(), GetSpreadsheetCellText(spreadsheet, cell));
        }

        [Fact]
        public void StreamingDataTable_RejectsOversizedTextValues()
        {
            using var memory = new MemoryStream();
            var dataSet = new DataSet();
            var table = new DataTable("Items");
            table.Columns.Add("Notes", typeof(string));
            table.Rows.Add(new string ('A', 32768));
            dataSet.Tables.Add(table);
            var exception = Assert.Throws<ArgumentException>(() =>
            {
                using (var writer = ExcelDocument.CreateWriter(memory, leaveOpen: true))
                {
                    foreach (DataTable source in dataSet.Tables)
                        writer.WriteTable(source.TableName, source);
                }
            });
            Assert.Contains("32,767", exception.Message);
        }

        [Fact]
        public void StreamingDataTable_RejectsOversizedObjectStringValues()
        {
            using var memory = new MemoryStream();
            var dataSet = new DataSet();
            var table = new DataTable("Items");
            table.Columns.Add("Notes", typeof(object));
            table.Rows.Add(new string ('A', 32768));
            dataSet.Tables.Add(table);
            var exception = Assert.Throws<ArgumentException>(() =>
            {
                using (var writer = ExcelDocument.CreateWriter(memory, leaveOpen: true))
                {
                    foreach (DataTable source in dataSet.Tables)
                        writer.WriteTable(source.TableName, source);
                }
            });
            Assert.Contains("32,767", exception.Message);
        }

        [Fact]
        public void StreamingDataTable_FallsBackForOutOfRangeDateTimeOffset()
        {
            using var memory = new MemoryStream();
            var value = DateTimeOffset.MinValue;
            var dataSet = new DataSet();
            var table = new DataTable("Items");
            table.Columns.Add("When", typeof(DateTimeOffset));
            table.Rows.Add(value);
            dataSet.Tables.Add(table);
            using (var writer = ExcelDocument.CreateWriter(memory, leaveOpen: true))
            {
                foreach (DataTable source in dataSet.Tables)
                    writer.WriteTable(source.TableName, source);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var cell = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet.Descendants<Cell>().Single(c => c.CellReference?.Value == "A2");
            Assert.Equal(CellValues.InlineString, cell.DataType!.Value);
            Assert.Equal(value.ToString("o", CultureInfo.InvariantCulture), GetSpreadsheetCellText(spreadsheet, cell));
        }

        [Fact]
        public void StreamingDataTable_StreamsLargeWorksheetPackageReadable()
        {
            using var memory = new MemoryStream();
            var dataSet = new DataSet("Export");
            var table = new DataTable("Rows");
            table.Columns.Add("Index", typeof(int));
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("When", typeof(DateTime));
            table.Columns.Add("Duration", typeof(TimeSpan));
            for (int i = 0; i < 5000; i++)
            {
                table.Rows.Add(i, "Row " + i.ToString(CultureInfo.InvariantCulture), new DateTime(2026, 5, 18).AddMinutes(i), TimeSpan.FromSeconds(i));
            }

            dataSet.Tables.Add(table);
            using (var writer = ExcelDocument.CreateWriter(memory, leaveOpen: true))
            {
                foreach (DataTable source in dataSet.Tables)
                    writer.WriteTable(source.TableName, source);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheet = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet!;
            var cells = worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("4999", cells["A5001"].CellValue!.Text);
            Assert.Equal(CellValues.InlineString, cells["B5001"].DataType!.Value);
            Assert.Equal("Row 4999", GetSpreadsheetCellText(spreadsheet, cells["B5001"]));
            Assert.Equal(1U, cells["C5001"].StyleIndex!.Value);
            Assert.Equal(2U, cells["D5001"].StyleIndex!.Value);
        }

        [Fact]
        public void StreamingDataTable_SmallStringExportSkipsSharedStringTable()
        {
            using var memory = new MemoryStream();
            var dataSet = new DataSet("Export");
            var table = new DataTable("Rows");
            table.Columns.Add("Name", typeof(string));
            for (int i = 0; i < 20; i++)
            {
                table.Rows.Add("Row " + i.ToString(CultureInfo.InvariantCulture));
            }

            dataSet.Tables.Add(table);
            using (var writer = ExcelDocument.CreateWriter(memory, leaveOpen: true))
            {
                foreach (DataTable source in dataSet.Tables)
                    writer.WriteTable(source.TableName, source);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            Assert.Null(spreadsheet.WorkbookPart!.SharedStringTablePart);
            var worksheet = spreadsheet.WorkbookPart.WorksheetParts.First().Worksheet;
            var cells = worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal(CellValues.InlineString, cells["A2"].DataType!.Value);
            Assert.Equal("Row 0", GetSpreadsheetCellText(spreadsheet, cells["A2"]));
        }

        [Fact]
        public void EditableDataSet_DuplicateStringsBuildSharedStringIndexes()
        {
            using var memory = new MemoryStream();
            var dataSet = new DataSet("Export");
            var table = new DataTable("Rows");
            table.Columns.Add("Name", typeof(string));
            for (int i = 0; i < 600; i++)
            {
                table.Rows.Add("Repeated");
            }

            dataSet.Tables.Add(table);
            using (var editable = ExcelDocument.Create())
            {
                editable.InsertDataSet(dataSet);
                editable.Save(memory);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var sharedStrings = spreadsheet.WorkbookPart!.SharedStringTablePart!.SharedStringTable.Elements<SharedStringItem>().ToList();
            var worksheet = spreadsheet.WorkbookPart.WorksheetParts.First().Worksheet;
            var cells = worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal(new[] { "Name", "Repeated" }, sharedStrings.Select(item => item.InnerText).ToArray());
            Assert.Equal(CellValues.SharedString, cells["A2"].DataType!.Value);
            Assert.Equal("1", cells["A601"].CellValue!.Text);
        }

        [Fact]
        public void EditableDataSet_SharedStringPlannerRejectsOverLimitSharedValue()
        {
            using var memory = new MemoryStream();
            var dataSet = new DataSet("Export");
            var table = new DataTable("Rows");
            table.Columns.Add("Name", typeof(string));
            string tooLong = new('A', 32_768);
            for (int i = 0; i < 511; i++)
            {
                table.Rows.Add(tooLong);
            }

            dataSet.Tables.Add(table);
            var exception = Assert.Throws<ArgumentException>(() =>
            {
                using (var editable = ExcelDocument.Create())
                {
                    editable.InsertDataSet(dataSet);
                    editable.Save(memory);
                }
            });
            Assert.Contains("32,767", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void EditableDataSet_ForcedSharedHeaderCountsPriorDataOccurrence()
        {
            using var memory = new MemoryStream();
            var dataSet = new DataSet("Export");
            var first = new DataTable("First");
            first.Columns.Add("Value", typeof(string));
            first.Rows.Add("CrossSheetHeader");
            for (int i = 0; i < 600; i++)
            {
                first.Rows.Add("Repeated");
            }

            var second = new DataTable("Second");
            second.Columns.Add("CrossSheetHeader", typeof(string));
            second.Rows.Add("Other");
            dataSet.Tables.Add(first);
            dataSet.Tables.Add(second);
            using (var editable = ExcelDocument.Create())
            {
                editable.InsertDataSet(dataSet);
                editable.Save(memory);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var sharedStringTable = spreadsheet.WorkbookPart!.SharedStringTablePart!.SharedStringTable;
            int actualSharedCellReferences = spreadsheet.WorkbookPart.WorksheetParts.SelectMany(part => part.Worksheet.Descendants<Cell>()).Count(cell => cell.DataType?.Value == CellValues.SharedString);
            Assert.Contains(sharedStringTable.Elements<SharedStringItem>(), item => item.InnerText == "CrossSheetHeader");
            Assert.Equal((uint)actualSharedCellReferences, sharedStringTable.Count!.Value);
        }

        [Fact]
        public void EditableDataSet_SharedStringsSkipUniqueDataValues()
        {
            using var memory = new MemoryStream();
            var dataSet = new DataSet("Export");
            var table = new DataTable("Rows");
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Unique", typeof(string));
            for (int i = 0; i < 600; i++)
            {
                table.Rows.Add("Repeated", "Unique " + i.ToString(CultureInfo.InvariantCulture));
            }

            dataSet.Tables.Add(table);
            using (var editable = ExcelDocument.Create())
            {
                editable.InsertDataSet(dataSet);
                editable.Save(memory);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var sharedStrings = spreadsheet.WorkbookPart!.SharedStringTablePart!.SharedStringTable.Elements<SharedStringItem>().Select(item => item.InnerText).ToList();
            var worksheet = spreadsheet.WorkbookPart.WorksheetParts.First().Worksheet;
            var cells = worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal(new[] { "Name", "Unique", "Repeated" }, sharedStrings);
            Assert.Equal(CellValues.SharedString, cells["A2"].DataType!.Value);
            Assert.Equal("2", cells["A2"].CellValue!.Text);
            Assert.Equal(CellValues.InlineString, cells["B2"].DataType!.Value);
            Assert.Equal("Unique 0", GetSpreadsheetCellText(spreadsheet, cells["B2"]));
        }

        [Fact]
        public void EditableDataSet_PreservesMissingValuesFormatsAndEmptyStrings()
        {
            using var memory = new MemoryStream();
            var dataSet = new DataSet("Export");
            var table = new DataTable("Sparse");
            table.Columns.Add("Id", typeof(int));
            table.Columns.Add("OptionalCode", typeof(string));
            table.Columns.Add("ExplicitEmpty", typeof(string));
            table.Columns.Add("ReviewDate", typeof(DateTime));
            table.Rows.Add(1, DBNull.Value, string.Empty, new DateTime(2026, 5, 20));
            table.Rows.Add(2, "C2", DBNull.Value, DBNull.Value);
            dataSet.Tables.Add(table);
            using (var editable = ExcelDocument.Create())
            {
                editable.InsertDataSet(dataSet);
                editable.Save(memory);
            }

            memory.Position = 0;
            using (var spreadsheet = SpreadsheetDocument.Open(memory, false))
            {
                var worksheet = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet!;
                var cells = worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
                Assert.False(cells.ContainsKey("B2"));
                Assert.True(cells.ContainsKey("C2"));
                Assert.Equal(CellValues.InlineString, cells["C2"].DataType!.Value);
                Assert.Equal(string.Empty, GetSpreadsheetCellText(spreadsheet, cells["C2"]));
                Assert.True(cells.ContainsKey("D2"));
                Assert.True(cells.ContainsKey("A3"));
                Assert.True(cells.ContainsKey("B3"));
                Assert.False(cells.ContainsKey("C3"));
                AssertMissingSpreadsheetCell(cells, "D3");
                Assert.Equal(cells["D2"].StyleIndex!.Value, cells["D3"].StyleIndex!.Value);
                Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
            }

            using var reader = ExcelDocumentReader.Open(memory.ToArray());
            DataTable imported = reader.ReadTableAsDataTable("Sparse");
            Assert.Equal(2, imported.Rows.Count);
            Assert.Equal(DBNull.Value, imported.Rows[0]["OptionalCode"]);
            Assert.Equal(string.Empty, imported.Rows[0]["ExplicitEmpty"]);
            Assert.Equal(DBNull.Value, imported.Rows[1]["ExplicitEmpty"]);
            Assert.Equal(DBNull.Value, imported.Rows[1]["ReviewDate"]);
        }

#if NET6_0_OR_GREATER
        [Fact]
        public void StreamingDataTable_AppliesDateOnlyAndTimeOnlyStyles()
        {
            using var memory = new MemoryStream();
            var dataSet = new DataSet();
            var table = new DataTable("Items");
            table.Columns.Add("Date", typeof(DateOnly));
            table.Columns.Add("Time", typeof(TimeOnly));
            table.Rows.Add(new DateOnly(2026, 5, 17), new TimeOnly(14, 30, 0));
            dataSet.Tables.Add(table);
            using (var writer = ExcelDocument.CreateWriter(memory, leaveOpen: true))
            {
                foreach (DataTable source in dataSet.Tables)
                    writer.WriteTable(source.TableName, source);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var cells = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet.Descendants<Cell>().ToDictionary(c => c.CellReference!.Value!);
            Assert.Equal(1U, cells["A2"].StyleIndex!.Value);
            Assert.Equal(2U, cells["B2"].StyleIndex!.Value);
        }

#endif
        [Fact]
        public void StreamingDataTable_ObjectColumnDateAndTimeValuesUseValueStyles()
        {
            using var memory = new MemoryStream();
            var dataSet = new DataSet();
            var table = new DataTable("Mixed");
            table.Columns.Add("Kind", typeof(string));
            table.Columns.Add("Value", typeof(object));
            table.Rows.Add("When", new DateTime(2026, 5, 19, 8, 30, 0));
            table.Rows.Add("Duration", TimeSpan.FromMinutes(95));
            dataSet.Tables.Add(table);
            using (var writer = ExcelDocument.CreateWriter(memory, leaveOpen: true))
            {
                foreach (DataTable source in dataSet.Tables)
                    writer.WriteTable(source.TableName, source);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var cells = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet!.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal(1U, cells["B2"].StyleIndex!.Value);
            Assert.Equal(2U, cells["B3"].StyleIndex!.Value);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
        }

        [Fact]
        public void StreamingDataTable_RejectsTooManyColumns()
        {
            using var memory = new MemoryStream();
            var dataSet = new DataSet();
            var table = new DataTable("TooWide");
            for (int i = 0; i <= ExcelA1.MaxColumns; i++)
            {
                table.Columns.Add("Column" + i.ToString(CultureInfo.InvariantCulture));
            }

            dataSet.Tables.Add(table);
            var exception = Assert.Throws<ArgumentException>(() =>
            {
                using (var writer = ExcelDocument.CreateWriter(memory, leaveOpen: true))
                {
                    foreach (DataTable source in dataSet.Tables)
                        writer.WriteTable(source.TableName, source);
                }
            });
            Assert.Contains(ExcelA1.MaxColumns.ToString(CultureInfo.InvariantCulture), exception.Message);
        }

    }
}
