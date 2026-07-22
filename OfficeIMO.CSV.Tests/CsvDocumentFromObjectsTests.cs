using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using OfficeIMO.CSV;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public class CsvDocumentFromObjectsTests
{
    [Fact]
    public void FromObjects_UsesPropertiesAndDelimiter()
    {
        var items = new[]
        {
            new { Name = "A", Value = 1 },
            new { Name = "B", Value = 2 }
        };

        var doc = CsvDocument.FromObjects(items, ';');

        Assert.Equal(';', doc.Delimiter);
        Assert.Equal(new[] { "Name", "Value" }, doc.Header);

        var rows = doc.AsEnumerable().ToList();
        Assert.Equal(2, rows.Count);
        Assert.Equal("A", rows[0].AsString("Name"));
        Assert.Equal(1, rows[0].AsInt32("Value"));
        Assert.Equal("B", rows[1].AsString("Name"));
        Assert.Equal(2, rows[1].AsInt32("Value"));
    }

    [Fact]
    public void FromObjects_UsesDictionaryKeys()
    {
        var items = new List<Dictionary<string, object?>>
        {
            new() { ["Name"] = "C", ["Value"] = 3 },
            new() { ["Name"] = "D", ["Value"] = 4 }
        };

        var doc = CsvDocument.FromObjects(items);

        Assert.Contains("Name", doc.Header);
        Assert.Contains("Value", doc.Header);
        Assert.Equal(2, doc.AsEnumerable().Count());
        Assert.Equal("C", doc.AsEnumerable().ElementAt(0).AsString("Name"));    
        Assert.Equal(4, doc.AsEnumerable().ElementAt(1).AsInt32("Value"));      
    }

    [Fact]
    public void WriteObjects_WritesWithoutMaterializingDocument()
    {
        var items = new object?[]
        {
            new { Name = "A", Value = 1 },
            new { Name = "B, quoted", Value = 2 }
        };

        using var writer = new StringWriter();
        CsvDocument.WriteObjects(writer, items, new CsvSaveOptions { NewLine = "\n" });

        Assert.Equal("Name,Value\nA,1\n\"B, quoted\",2\n", writer.ToString());
    }

    [Fact]
    public void WriteObjects_QuotesBooleanValuesWhenCustomDelimiterAppearsInLiteral()
    {
        var items = new object?[]
        {
            new { Name = "A", Enabled = true },
            new { Name = "B", Enabled = false }
        };

        using var writer = new StringWriter();
        CsvDocument.WriteObjects(writer, items, new CsvSaveOptions { Delimiter = 'r', NewLine = "\n" });

        Assert.Equal("NamerEnabled\nAr\"True\"\nBrFalse\n", writer.ToString());
    }

    [Fact]
    public void WriteObjects_ProjectsDictionaryRowsByFirstRowColumns()
    {
        var items = new object?[]
        {
            new Dictionary<string, object?> { ["Name"] = "A", ["Value"] = 1 },
            new Dictionary<string, object?> { ["Value"] = 2, ["Name"] = "B" }
        };

        using var writer = new StringWriter();
        CsvDocument.WriteObjects(writer, items, new CsvSaveOptions { NewLine = "\n" });

        Assert.Equal("Name,Value\nA,1\nB,2\n", writer.ToString());
    }

    [Fact]
    public void CsvObjectWriter_StreamsRowsAndCanLeaveWriterOpen()
    {
        using var writer = new StringWriter();
        using (var csvWriter = new CsvObjectWriter(writer, new CsvSaveOptions { NewLine = "\n" }, leaveOpen: true))
        {
            csvWriter.WriteObject(new { Name = "A", Value = 1 });
            csvWriter.WriteObject(new { Name = "B", Value = 2 });
            Assert.True(csvWriter.HasRows);
        }

        writer.Write("#");
        Assert.Equal("Name,Value\nA,1\nB,2\n#", writer.ToString());
    }

    [Fact]
    public void CsvObjectWriter_FlushesWhenLeavingWriterOpen()
    {
        using var writer = new FlushTrackingWriter();
        using (var csvWriter = new CsvObjectWriter(writer, new CsvSaveOptions { NewLine = "\n" }, leaveOpen: true))
        {
            csvWriter.WriteObject(new { Name = "A", Value = 1 });
        }

        Assert.True(writer.WasFlushed);
        writer.Write("#");
        Assert.Equal("Name,Value\nA,1\n#", writer.ToString());
    }

    [Fact]
    public void CsvObjectWriter_WritesProjectedRows()
    {
        using var writer = new StringWriter();
        using (var csvWriter = new CsvObjectWriter(writer, new CsvSaveOptions { NewLine = "\n" }, leaveOpen: true))
        {
            csvWriter.WriteRow(new[] { "Name", "Value" }, new object?[] { "A", 1 });
            csvWriter.WriteRow(new[] { "Name", "Value" }, new object?[] { "B", 2 });
        }

        Assert.Equal("Name,Value\nA,1\nB,2\n", writer.ToString());
    }

    [Fact]
    public void CsvObjectWriter_WritesNarrowPlainTextRowsToGeneralTextWriter()
    {
        using var writer = new FlushTrackingWriter();
        using (var csvWriter = new CsvObjectWriter(writer, new CsvSaveOptions { NewLine = "\n" }, leaveOpen: true))
        {
            csvWriter.WriteTextRow(new[] { "Name", "Value" }, new string?[] { "Alpha", "1" });
            csvWriter.WriteTrustedTextRow(new string?[] { "Beta", "2" });
        }

        Assert.Equal("Name,Value\nAlpha,1\nBeta,2\n", writer.ToString());
    }

    [Fact]
    public void CsvObjectWriter_AlwaysQuotedProjectedRowsPreserveEscaping()
    {
        var created = new DateTime(2026, 1, 2, 3, 4, 5, DateTimeKind.Utc);
        using var writer = new StringWriter(CultureInfo.InvariantCulture);
        using (var csvWriter = new CsvObjectWriter(writer, new CsvSaveOptions { NewLine = "\n", QuoteMode = CsvQuoteMode.Always }, leaveOpen: true))
        {
            csvWriter.WriteRow(
                new[] { "Id", "Amount", "Enabled", "Created", "Name", "Missing" },
                new object?[] { 1, 12.5m, true, created, "A\"B", null });
        }

        var expectedCreated = created.ToString(CultureInfo.InvariantCulture);
        var expected =
            "\"Id\",\"Amount\",\"Enabled\",\"Created\",\"Name\",\"Missing\"\n" +
            $"\"1\",\"12.5\",\"True\",\"{expectedCreated}\",\"A\"\"B\",\"\"\n";

        Assert.Equal(expected, writer.ToString());
    }

    [Fact]
    public void CsvObjectWriter_WritesWideTextRowsWithEscaping()
    {
        var columns = Enumerable.Range(1, 21).Select(static index => $"C{index}").ToArray();
        var values = Enumerable.Range(1, 21).Select(static index => $"V{index}").ToArray();
        values[5] = "A,B";
        values[10] = "A\"B";

        using var writer = new StringWriter();
        using (var csvWriter = new CsvObjectWriter(writer, new CsvSaveOptions { NewLine = "\n" }, leaveOpen: true))
        {
            csvWriter.WriteTextRow(columns, values);
        }

        var output = writer.ToString();
        var lines = output.Split('\n');
        Assert.Equal(string.Join(",", columns), lines[0]);
        Assert.Contains("\"A,B\"", lines[1], StringComparison.Ordinal);
        Assert.Contains("\"A\"\"B\"", lines[1], StringComparison.Ordinal);
        Assert.EndsWith("\n", output, StringComparison.Ordinal);
    }

    [Fact]
    public void CsvObjectWriter_RejectsProjectedRowsWithDifferentColumns()
    {
        using var writer = new StringWriter();
        using var csvWriter = new CsvObjectWriter(writer, new CsvSaveOptions { NewLine = "\n" }, leaveOpen: true);
        csvWriter.WriteRow(new[] { "Name", "Value" }, new object?[] { "A", 1 });

        Assert.Throws<CsvException>(() => csvWriter.WriteRow(new[] { "Value", "Name" }, new object?[] { 2, "B" }));
    }

    [Fact]
    public void CsvObjectWriter_ValidatesProjectedRowWidthBeforeWritingHeader()
    {
        using var writer = new StringWriter();
        using var csvWriter = new CsvObjectWriter(writer, new CsvSaveOptions { NewLine = "\n" }, leaveOpen: true);

        Assert.Throws<CsvException>(() => csvWriter.WriteRow(new[] { "Name", "Value" }, new object?[] { "A" }));
        Assert.Equal(string.Empty, writer.ToString());
    }

    [Fact]
    public void CsvObjectWriter_WritesDataReaderSchemaAndRows()
    {
        using var reader = CreateReader();
        using var writer = new StringWriter();
        using (var csvWriter = new CsvObjectWriter(writer, new CsvSaveOptions { NewLine = "\n", NullValue = "<null>" }, leaveOpen: true))
        {
            csvWriter.WriteDataReader(reader);
        }

        Assert.Equal("Name,Score,Notes\nAlpha,1.5,\"A, quoted\"\nBeta,<null>,\n", writer.ToString());
    }

    [Fact]
    public void CsvObjectWriter_TreatsDBNullAsNullValue()
    {
        using var writer = new StringWriter();
        using (var csvWriter = new CsvObjectWriter(writer, new CsvSaveOptions { NewLine = "\n", NullValue = "<null>" }, leaveOpen: true))
        {
            csvWriter.WriteRow(new[] { "Name", "Score" }, new object?[] { "Alpha", DBNull.Value });
        }

        Assert.Equal("Name,Score\nAlpha,<null>\n", writer.ToString());
    }

    [Fact]
    public void WriteDataReader_WritesReaderWithoutMaterializingDocument()
    {
        using var reader = CreateReader();
        using var writer = new StringWriter();

        CsvDocument.WriteDataReader(writer, reader, new CsvSaveOptions { NewLine = "\n", IncludeHeader = false });

        Assert.Equal("Alpha,1.5,\"A, quoted\"\nBeta,,\n", writer.ToString());
    }

    [Fact]
    public void WriteDataReader_FallsBackWhenGetValuesIsUnsupported()
    {
        using var reader = new ThrowingGetValuesDataReader(
            new[] { "Name", "Score", "Notes" },
            new[]
            {
                new object?[] { "Alpha", 1.5m, "A, quoted" },
                new object?[] { "Beta", DBNull.Value, string.Empty }
            });
        using var writer = new StringWriter();

        CsvDocument.WriteDataReader(writer, reader, new CsvSaveOptions { NewLine = "\n", NullValue = "<null>" });

        Assert.Equal("Name,Score,Notes\nAlpha,1.5,\"A, quoted\"\nBeta,<null>,\n", writer.ToString());
    }

    [Fact]
    public void WriteDataReader_UsesReportedFieldTypesWithoutRequiringGetValues()
    {
        using var reader = new ThrowingGetValuesDataReader(
            new[] { "Name", "Score", "Notes" },
            new[]
            {
                new object?[] { "Alpha", 1.5m, "A, quoted" },
                new object?[] { "Beta", "reported type changed", DBNull.Value }
            });
        using var writer = new StringWriter(CultureInfo.InvariantCulture);

        CsvDocument.WriteDataReader(writer, reader, new CsvSaveOptions { NewLine = "\n" });

        Assert.Equal(
            "Name,Score,Notes\nAlpha,1.5,\"A, quoted\"\nBeta,reported type changed,\n",
            writer.ToString());
    }

    [Fact]
    public void WriteDataReader_SchemaPlanStillEscapesCultureSpecificValues()
    {
        var table = new DataTable();
        table.Columns.Add("Amount", typeof(decimal));
        table.Rows.Add(1.5m);
        using var reader = table.CreateDataReader();
        using var writer = new StringWriter(CultureInfo.InvariantCulture);

        CsvDocument.WriteDataReader(
            writer,
            reader,
            new CsvSaveOptions { NewLine = "\n", Culture = CultureInfo.GetCultureInfo("pl-PL") });

        Assert.Equal("Amount\n\"1,5\"\n", writer.ToString());
    }

    [Fact]
    public void WriteDataReader_SchemaPlanPreservesCommonScalarFormats()
    {
        var values = new object[]
        {
            42,
            4_294_967_296L,
            12.5m,
            3.25d,
            1.5f,
            new DateTime(2026, 7, 14, 12, 34, 56, DateTimeKind.Utc),
            new DateTimeOffset(2026, 7, 14, 12, 34, 56, TimeSpan.FromHours(2)),
            Guid.Parse("b6e2be52-3367-4bcd-84dc-c092a152ff73"),
            new TimeSpan(1, 2, 3),
            true
        };
        var table = new DataTable();
        for (var i = 0; i < values.Length; i++)
        {
            table.Columns.Add("Value" + i, values[i].GetType());
        }

        table.Rows.Add(values);
        using var reader = table.CreateDataReader();
        using var writer = new StringWriter(CultureInfo.InvariantCulture);

        CsvDocument.WriteDataReader(writer, reader, new CsvSaveOptions { NewLine = "\n" });

        var expectedHeader = string.Join(",", table.Columns.Cast<DataColumn>().Select(column => column.ColumnName));
        var expectedRow = string.Join(",", values.Select(value => Convert.ToString(value, CultureInfo.InvariantCulture)));
        Assert.Equal(expectedHeader + "\n" + expectedRow + "\n", writer.ToString());
    }

    private sealed class FlushTrackingWriter : StringWriter
    {
        public bool WasFlushed { get; private set; }

        public override void Flush()
        {
            WasFlushed = true;
            base.Flush();
        }
    }

    private static IDataReader CreateReader()
    {
        var table = new DataTable();
        table.Columns.Add("Name", typeof(string));
        table.Columns.Add("Score", typeof(decimal));
        table.Columns.Add("Notes", typeof(string));
        table.Rows.Add("Alpha", 1.5m, "A, quoted");
        table.Rows.Add("Beta", DBNull.Value, string.Empty);
        return table.CreateDataReader();
    }

    [Fact]
    public void ReadRows_StreamsHeaderAndRows()
    {
        using var reader = new StringReader("Name,Value\nA,1\nB,2\n");
        var rows = new List<string>();

        CsvDocument.ReadRows(reader, (header, values) =>
        {
            Assert.Equal(new[] { "Name", "Value" }, header);
            rows.Add(values[0] + ":" + values[1]);
        });

        Assert.Equal(new[] { "A:1", "B:2" }, rows);
    }

    [Fact]
    public void ReadRowsReusable_StreamsHeaderAndRows()
    {
        using var reader = new StringReader("Name,Value\nA,1\nB,2\n");
        var rows = new List<string>();

        CsvDocument.ReadRowsReusable(reader, (header, values) =>
        {
            Assert.Equal(new[] { "Name", "Value" }, header);
            rows.Add(values[0] + ":" + values[1]);
        });

        Assert.Equal(new[] { "A:1", "B:2" }, rows);
    }

    [Fact]
    public void ReadRowsReusable_ReusesUnquotedRowBuffer()
    {
        using var reader = new StringReader("Name,Value\nA,1\nB,2\n");
        IReadOnlyList<string>? first = null;
        IReadOnlyList<string>? second = null;

        CsvDocument.ReadRowsReusable(reader, (_, values) =>
        {
            if (first == null)
            {
                first = values;
            }
            else
            {
                second = values;
            }
        });

        Assert.NotNull(first);
        Assert.Same(first, second);
    }

    [Fact]
    public void ReadRowsReusable_HandlesQuotedRows()
    {
        using var reader = new StringReader("Name,Value\n\"A, quoted\",1\nB,2\n");
        var rows = new List<string>();

        CsvDocument.ReadRowsReusable(reader, (_, values) =>
        {
            rows.Add(values[0] + ":" + values[1]);
        });

        Assert.Equal(new[] { "A, quoted:1", "B:2" }, rows);
    }

    [Fact]
    public void ReadRecordsReusable_StreamsRawHeaderAndRows()
    {
        using var reader = new StringReader("Name,Value\nA,1\nB,2\n");
        var records = new List<string>();

        CsvDocument.ReadRecordsReusable(reader, values =>
        {
            records.Add(values[0] + ":" + values[1]);
        });

        Assert.Equal(new[] { "Name:Value", "A:1", "B:2" }, records);
    }

    [Fact]
    public void ReadRecordsReusable_PreservesExtraFields()
    {
        using var reader = new StringReader("Name\nA,1\n");
        IReadOnlyList<string>? row = null;

        CsvDocument.ReadRecordsReusable(reader, values =>
        {
            if (values.Count > 1)
            {
                row = values.ToArray();
            }
        });

        Assert.NotNull(row);
        Assert.Equal(new[] { "A", "1" }, row);
    }

    [Fact]
    public void ReadRecords_DetectDelimiter_UsesRawRecordCommentSemantics()
    {
        using var reader = new StringReader("#a,b,c\n1;2;3\n");

        var records = CsvDocument.ReadRecords(reader, new CsvLoadOptions { DetectDelimiter = true }).ToArray();

        Assert.Equal(new[] { "#a", "b", "c" }, records[0]);
        Assert.Equal(new[] { "1;2;3" }, records[1]);
    }

    [Fact]
    public void SaveObjects_WritesFile()
    {
        var path = Path.Combine(Path.GetTempPath(), "OfficeIMO.CSV.SaveObjects." + Guid.NewGuid().ToString("N") + ".csv");
        try
        {
            CsvDocument.SaveObjects(path, new object?[] { new { Name = "A", Value = 1 } }, new CsvSaveOptions { NewLine = "\n" });

            Assert.Equal("Name,Value\nA,1\n", File.ReadAllText(path));
        }
        finally
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public void Save_Load_RoundTripsGZipByExtension()
    {
        var path = Path.Combine(Path.GetTempPath(), "OfficeIMO.CSV.RoundTrip." + Guid.NewGuid().ToString("N") + ".csv.gz");
        try
        {
            new CsvDocument()
                .WithHeader("Name", "Value")
                .AddRow("Alpha", 1)
                .Save(path, new CsvSaveOptions { NewLine = "\n" });

            var parsed = CsvDocument.Load(path);
            var row = Assert.Single(parsed.AsEnumerable());

            Assert.Equal("Alpha", row.AsString("Name"));
            Assert.Equal(1, row.AsInt32("Value"));
        }
        finally
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public void ReadRowsReusable_ReadsGZipByExtension()
    {
        var path = Path.Combine(Path.GetTempPath(), "OfficeIMO.CSV.Stream." + Guid.NewGuid().ToString("N") + ".csv.gz");
        try
        {
            new CsvDocument()
                .WithHeader("Name", "Value")
                .AddRow("Alpha", 1)
                .AddRow("Beta", 2)
                .Save(path, new CsvSaveOptions { NewLine = "\n" });

            var rows = new List<string>();
            CsvDocument.ReadRowsReusable(path, (_, row) => rows.Add(row[0] + "|" + row[1]));

            Assert.Equal(new[] { "Alpha|1", "Beta|2" }, rows);
        }
        finally
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public void ReadRecords_ReadsGZipByExtension()
    {
        var path = Path.Combine(Path.GetTempPath(), "OfficeIMO.CSV.RawRecords." + Guid.NewGuid().ToString("N") + ".csv.gz");
        try
        {
            new CsvDocument()
                .WithHeader("Name", "Value")
                .AddRow("Alpha", 1)
                .Save(path, new CsvSaveOptions { NewLine = "\n" });

            var records = CsvDocument.ReadRecords(path).ToArray();

            Assert.Equal(2, records.Length);
            Assert.Equal(new[] { "Name", "Value" }, records[0]);
            Assert.Equal(new[] { "Alpha", "1" }, records[1]);
        }
        finally
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public void ReadRecordsReusable_ReadsGZipByExtension()
    {
        var path = Path.Combine(Path.GetTempPath(), "OfficeIMO.CSV.RawRecordsReusable." + Guid.NewGuid().ToString("N") + ".csv.gz");
        try
        {
            new CsvDocument()
                .WithHeader("Name", "Value")
                .AddRow("Alpha", 1)
                .AddRow("Beta", 2)
                .Save(path, new CsvSaveOptions { NewLine = "\n" });

            var records = new List<string[]>();
            CsvDocument.ReadRecordsReusable(path, values => records.Add(values.ToArray()));

            Assert.Equal(3, records.Count);
            Assert.Equal(new[] { "Name", "Value" }, records[0]);
            Assert.Equal(new[] { "Alpha", "1" }, records[1]);
            Assert.Equal(new[] { "Beta", "2" }, records[2]);
        }
        finally
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
    }

#if NET8_0_OR_GREATER
    [Fact]
    public void ReadFieldSpans_ReadsGZipByExtension()
    {
        var path = Path.Combine(Path.GetTempPath(), "OfficeIMO.CSV.FieldSpans." + Guid.NewGuid().ToString("N") + ".csv.gz");
        try
        {
            new CsvDocument()
                .WithHeader("Name", "Value")
                .AddRow("Alpha", 1)
                .Save(path, new CsvSaveOptions { NewLine = "\n" });

            var fields = new List<string>();
            CsvDocument.ReadFieldSpans(path, (recordIndex, fieldIndex, value) => fields.Add($"{recordIndex}:{fieldIndex}:{value.ToString()}"));

            Assert.Equal(new[] { "0:0:Name", "0:1:Value", "1:0:Alpha", "1:1:1" }, fields);
        }
        finally
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
    }
#endif

    [Fact]
    public void Load_EnforcesMaxDecompressedBytes()
    {
        var path = Path.Combine(Path.GetTempPath(), "OfficeIMO.CSV.Bounded." + Guid.NewGuid().ToString("N") + ".csv.gz");
        try
        {
            new CsvDocument()
                .WithHeader("Name", "Value")
                .AddRow("Alpha", 1)
                .Save(path, new CsvSaveOptions { NewLine = "\n" });

            Assert.Throws<InvalidOperationException>(() =>
                CsvDocument.Load(path, new CsvLoadOptions { MaxDecompressedBytes = 4 }));
        }
        finally
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public void LoadOptions_BoundsCompressedInputByDefault()
    {
        var options = new CsvLoadOptions();

        Assert.Equal(CsvLoadOptions.DefaultMaxInputBytes, options.MaxDecompressedBytes);
    }

    [Fact]
    public void Load_DoesNotApplyDecompressionLimitToPlainCsvFiles()
    {
        var path = Path.Combine(Path.GetTempPath(), "OfficeIMO.CSV.Plain." + Guid.NewGuid().ToString("N") + ".csv");
        try
        {
            File.WriteAllText(path, "Name,Value\nAlpha,1\n");

            CsvDocument document = CsvDocument.Load(path, new CsvLoadOptions { MaxDecompressedBytes = 1 });

            Assert.Equal("Alpha", Assert.Single(document.AsEnumerable()).AsString("Name"));
        }
        finally
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public void Save_InvalidCompressionLevelDoesNotTruncateExistingFile()
    {
        var path = Path.Combine(Path.GetTempPath(), "OfficeIMO.CSV.InvalidCompressionLevel." + Guid.NewGuid().ToString("N") + ".csv.gz");
        const string original = "Name,Value\nOriginal,1\n";
        try
        {
            File.WriteAllText(path, original);

            Assert.Throws<ArgumentOutOfRangeException>(() =>
                new CsvDocument()
                    .WithHeader("Name", "Value")
                    .AddRow("Replacement", 2)
                    .Save(path, new CsvSaveOptions
                    {
                        CompressionType = CsvCompressionType.GZip,
                        CompressionLevel = (CompressionLevel)123,
                        NewLine = "\n"
                    }));

            Assert.Equal(original, File.ReadAllText(path));
        }
        finally
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
    }

#if !NET8_0_OR_GREATER
    [Fact]
    public void Save_UnsupportedCompressionDoesNotTruncateExistingFile()
    {
        var path = Path.Combine(Path.GetTempPath(), "OfficeIMO.CSV.UnsupportedCompression." + Guid.NewGuid().ToString("N") + ".csv.br");
        const string original = "Name,Value\nOriginal,1\n";
        try
        {
            File.WriteAllText(path, original);

            Assert.Throws<PlatformNotSupportedException>(() =>
                new CsvDocument()
                    .WithHeader("Name", "Value")
                    .AddRow("Replacement", 2)
                    .Save(path, new CsvSaveOptions
                    {
                        CompressionType = CsvCompressionType.Brotli,
                        NewLine = "\n"
                    }));

            Assert.Equal(original, File.ReadAllText(path));
        }
        finally
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
    }
#endif

    [Fact]
    public void SaveObjects_UsesCompressionFromDestinationExtension()
    {
        var path = Path.Combine(Path.GetTempPath(), "OfficeIMO.CSV.SaveObjects." + Guid.NewGuid().ToString("N") + ".csv.gz");
        try
        {
            CsvDocument.SaveObjects(path, new object?[] { new { Name = "A", Value = 1 } }, new CsvSaveOptions { NewLine = "\n" });

            var parsed = CsvDocument.Load(path);
            var row = Assert.Single(parsed.AsEnumerable());

            Assert.Equal("A", row.AsString("Name"));
            Assert.Equal(1, row.AsInt32("Value"));
        }
        finally
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public void SaveObjects_DoesNotReplaceExistingFileWhenInputIsEmpty()
    {
        var path = Path.Combine(Path.GetTempPath(), "OfficeIMO.CSV.SaveObjects.Preserve." + Guid.NewGuid().ToString("N") + ".csv");
        File.WriteAllText(path, "existing");

        try
        {
            Assert.Throws<ArgumentException>(() => CsvDocument.SaveObjects(path, Array.Empty<object?>(), new CsvSaveOptions { NewLine = "\n" }));

            Assert.Equal("existing", File.ReadAllText(path));
        }
        finally
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public void SaveObjects_ReplacesExistingFileWithCompletedOutput()
    {
        var path = Path.Combine(Path.GetTempPath(), "OfficeIMO.CSV.SaveObjects.Replace." + Guid.NewGuid().ToString("N") + ".csv");
        File.WriteAllText(path, "existing");

        try
        {
            CsvDocument.SaveObjects(path, new object?[] { new { Name = "B", Value = 2 } }, new CsvSaveOptions { NewLine = "\n" });

            Assert.Equal("Name,Value\nB,2\n", File.ReadAllText(path));
        }
        finally
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public void FromObjects_ThrowsOnNullItems()
    {
        var ex = Assert.Throws<ArgumentNullException>(() => CsvDocument.FromObjects(null!));
        Assert.Equal("items", ex.ParamName);
    }

    [Fact]
    public void FromObjects_ThrowsOnEmptyItems()
    {
        var ex = Assert.Throws<ArgumentException>(() => CsvDocument.FromObjects(Array.Empty<object?>()));
        Assert.StartsWith("Provide at least one data row.", ex.Message);
        Assert.Equal("items", ex.ParamName);
    }

    [Fact]
    public void FromObjects_ThrowsOnNullFirstRow()
    {
        var ex = Assert.Throws<ArgumentException>(() => CsvDocument.FromObjects(new object?[] { null }));
        Assert.StartsWith("Data rows cannot be null.", ex.Message);
        Assert.Equal("items", ex.ParamName);
    }

    [Fact]
    public void FromObjects_ThrowsOnNullEntry()
    {
        var ex = Assert.Throws<InvalidOperationException>(() => CsvDocument.FromObjects(new object?[] { new { Name = "A", Value = 1 }, null }));
        Assert.Equal("Data rows cannot contain null entries.", ex.Message);
    }

    [Fact]
    public void FromObjects_ThrowsOnMissingColumns()
    {
        var ex = Assert.Throws<InvalidOperationException>(() => CsvDocument.FromObjects(new object?[] { new object() }));
        Assert.Equal("Unable to infer column names. Use objects with properties or dictionaries.", ex.Message);
    }
}
