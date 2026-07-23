using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.CSV;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public class CsvStreamingTests
{
    [Fact]
    public async Task LoadAsync_PathHonorsCompressionFromExtension()
    {
        string path = Path.Combine(Path.GetTempPath(), "OfficeIMO.CSV.Async." + Guid.NewGuid().ToString("N") + ".csv.gz");
        try
        {
            new CsvDocument()
                .WithHeader("Id", "Name")
                .AddRow(1, "Alice")
                .Save(path, new CsvSaveOptions { NewLine = "\n" });

            CsvDocument document = await CsvDocument.LoadAsync(path);

            Assert.Equal("Alice", Assert.Single(document.AsEnumerable()).AsString("Name"));
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public async Task LoadAsync_PathHonorsExplicitCompression()
    {
        string path = Path.Combine(Path.GetTempPath(), "OfficeIMO.CSV.Async." + Guid.NewGuid().ToString("N") + ".data");
        try
        {
            new CsvDocument()
                .WithHeader("Id", "Name")
                .AddRow(1, "Alice")
                .Save(path, new CsvSaveOptions { CompressionType = CsvCompressionType.GZip, NewLine = "\n" });

            CsvDocument document = await CsvDocument.LoadAsync(
                path,
                new CsvLoadOptions { CompressionType = CsvCompressionType.GZip });

            Assert.Equal("Alice", Assert.Single(document.AsEnumerable()).AsString("Name"));
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public async Task LoadAsync_PathEnforcesMaxDecompressedBytes()
    {
        string path = Path.Combine(Path.GetTempPath(), "OfficeIMO.CSV.AsyncBounded." + Guid.NewGuid().ToString("N") + ".csv.gz");
        try
        {
            new CsvDocument()
                .WithHeader("Id", "Name")
                .AddRow(1, "Alice")
                .Save(path, new CsvSaveOptions { NewLine = "\n" });

            await Assert.ThrowsAsync<InvalidOperationException>(() =>
                CsvDocument.LoadAsync(path, new CsvLoadOptions { MaxDecompressedBytes = 4 }));
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public async Task PlainPathLoads_EnforceMaxInputBytes()
    {
        string path = Path.Combine(Path.GetTempPath(), "OfficeIMO.CSV.PathBounded." + Guid.NewGuid().ToString("N") + ".csv");
        try
        {
            File.WriteAllText(path, "Id,Name\n1,Alice\n", Encoding.UTF8);
            var options = new CsvLoadOptions { MaxInputBytes = 8 };

            Assert.Throws<InvalidOperationException>(() => CsvDocument.Load(path, options));
            Assert.Throws<InvalidOperationException>(() =>
                CsvDocument.ReadRows(path, (_, _) => { }, options));
            Assert.Throws<InvalidOperationException>(() =>
            {
                using var reader = CsvDocument.CreateDataReader(
                    path,
                    options,
                    new CsvDataReaderOptions { InferSchema = true });
                reader.Read();
            });
            Assert.Throws<InvalidOperationException>(() =>
            {
                var visitor = new CapturingRowFieldSpanVisitor(new List<string>());
                CsvDocument.ReadRowFieldSpans(path, ref visitor, options);
            });
            await Assert.ThrowsAsync<InvalidOperationException>(() => CsvDocument.LoadAsync(path, options));

            Assert.Throws<ArgumentOutOfRangeException>(() =>
                CsvDocument.CreateDataReader(
                    path,
                    new CsvLoadOptions { Mode = CsvLoadMode.Stream, MaxInputBytes = 0 },
                    new CsvDataReaderOptions { InferSchema = true }));
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public async Task LoadAsync_SnapshotsCompleteCallerStreamAndRestoresPosition()
    {
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes("Id,Name\n1,Alice\n"));
        stream.Position = 5;

        CsvDocument document = await CsvDocument.LoadAsync(stream);

        Assert.Equal(5, stream.Position);
        Assert.Equal("Alice", Assert.Single(document.AsEnumerable()).AsString("Name"));
        stream.ReadByte();
    }

    [Fact]
    public async Task LoadAsync_HonorsPreCanceledTokenAndRestoresPosition()
    {
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes("Id\n1\n"));
        stream.Position = 2;
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            CsvDocument.LoadAsync(stream, cancellationToken: cancellation.Token));

        Assert.Equal(2, stream.Position);
    }

    [Fact]
    public async Task StreamLoads_EnforceCompleteInputLimitAndRestorePosition()
    {
        byte[] bytes = Encoding.UTF8.GetBytes("Id,Name\n1,Alice\n");
        using var syncStream = new MemoryStream(bytes);
        syncStream.Position = 3;
        Assert.Throws<InvalidDataException>(() =>
            CsvDocument.Load(syncStream, new CsvLoadOptions { MaxInputBytes = 8 }));
        Assert.Equal(3, syncStream.Position);

        using var asyncStream = new MemoryStream(bytes);
        asyncStream.Position = 4;
        await Assert.ThrowsAsync<InvalidDataException>(() =>
            CsvDocument.LoadAsync(asyncStream, new CsvLoadOptions { MaxInputBytes = 8 }));
        Assert.Equal(4, asyncStream.Position);
    }

    [Fact]
    public void LoadOptionsClone_PreservesCallerErrorSinkWithoutMutatingMissingSink()
    {
        var originalError = new CsvParseError(1, "original", new FormatException("original"));
        var options = new CsvLoadOptions {
            CollectParseErrors = true,
            ParseErrors = new List<CsvParseError> { originalError }
        };

        CsvLoadOptions clone = options.Clone();
        Assert.Same(options.ParseErrors, clone.ParseErrors);
        Assert.Same(originalError, Assert.Single(options.ParseErrors!));

        var withoutErrorCollection = new CsvLoadOptions { CollectParseErrors = true };
        CsvLoadOptions emptyClone = withoutErrorCollection.Clone();

        Assert.Null(withoutErrorCollection.ParseErrors);
        Assert.Empty(emptyClone.ParseErrors!);
    }

    [Fact]
    public void StreamingMode_ReadsLazily()
    {
        var tempPath = Path.GetTempFileName();
        try
        {
            File.WriteAllText(tempPath, "Id,Name\n1,Alice\n2,Bob\n3,Charlie\n");

            var options = new CsvLoadOptions { Mode = CsvLoadMode.Stream };
            var doc = CsvDocument.Load(tempPath, options);

            var rows = doc.AsEnumerable().ToList();
            Assert.Equal(3, rows.Count);
            Assert.Throws<InvalidOperationException>(() => doc.SortBy("Id"));

            doc.Materialize().SortBy("Id");
            Assert.Equal(1, doc.AsEnumerable().First().AsInt32("Id"));
        }
        finally
        {
            File.Delete(tempPath);
        }
    }

    [Fact]
    public void StreamingMode_ClonesFileOpenOptionsBeforeEnumeration()
    {
        var tempPath = Path.Combine(Path.GetTempPath(), "OfficeIMO.CSV.StreamingClone." + Guid.NewGuid().ToString("N") + ".csv.gz");
        try
        {
            new CsvDocument()
                .WithHeader("Id", "Name")
                .AddRow(1, "Alice")
                .Save(tempPath, new CsvSaveOptions { NewLine = "\n" });

            var options = new CsvLoadOptions { Mode = CsvLoadMode.Stream };
            var doc = CsvDocument.Load(tempPath, options);

            options.CompressionType = CsvCompressionType.None;
            options.MaxDecompressedBytes = 1;

            var row = Assert.Single(doc.AsEnumerable());
            Assert.Equal("Alice", row.AsString("Name"));
        }
        finally
        {
            if (File.Exists(tempPath))
            {
                File.Delete(tempPath);
            }
        }
    }

    [Fact]
    public void StreamingMode_ClonesStaticColumnsBeforeEnumeration()
    {
        var tempPath = Path.Combine(Path.GetTempPath(), "OfficeIMO.CSV.StreamingStaticClone." + Guid.NewGuid().ToString("N") + ".csv");
        try
        {
            File.WriteAllText(tempPath, "Id,Name\n1,Alice\n");
            var staticColumns = new Dictionary<string, object?> {
                ["Source"] = "original.csv"
            };
            var doc = CsvDocument.Load(
                tempPath,
                new CsvLoadOptions {
                    Mode = CsvLoadMode.Stream,
                    StaticColumns = staticColumns
                });

            staticColumns["Batch"] = "late";

            var row = Assert.Single(doc.AsEnumerable());
            Assert.Equal(new[] { "Id", "Name", "Source" }, doc.Header);
            Assert.Equal("original.csv", row.AsString("Source"));
        }
        finally
        {
            if (File.Exists(tempPath))
            {
                File.Delete(tempPath);
            }
        }
    }

    [Fact]
    public void StreamingMode_Disallows_FilterUntilMaterialized()
    {
        var csv = "Id,Value\n1,A\n2,B\n";
        var doc = CsvDocument.Parse(csv, new CsvLoadOptions { Mode = CsvLoadMode.Stream });
        Assert.Throws<InvalidOperationException>(() => doc.Filter(_ => true));
    }

    [Fact]
    public void ReadRowsReusable_PreservesDelimiterOnlyRows()
    {
        var rows = new List<string[]>();
        using var reader = new StringReader("Name,Value\n,\n\nAlpha,1\n");

        CsvDocument.ReadRowsReusable(reader, (_, values) => rows.Add(values.ToArray()));

        Assert.Equal(2, rows.Count);
        Assert.Equal(new[] { string.Empty, string.Empty }, rows[0]);
        Assert.Equal(new[] { "Alpha", "1" }, rows[1]);
    }

    [Fact]
    public void ReadRowsReusable_CanSkipInitialRecords()
    {
        var rows = new List<string[]>();
        using var reader = new StringReader("metadata\nName,Value\nAlpha,1\n");

        CsvDocument.ReadRowsReusable(
            reader,
            (_, values) => rows.Add(values.ToArray()),
            new CsvLoadOptions { SkipInitialRecords = 1 });

        var row = Assert.Single(rows);
        Assert.Equal(new[] { "Alpha", "1" }, row);
    }

    [Fact]
    public void ReadRowsReusable_SkipsInitialRecordsAfterLeadingComments()
    {
        var headers = new List<string[]>();
        var rows = new List<string[]>();
        using var reader = new StringReader("#note\nmetadata,with,commas\nName;Value\nAlpha;1\n");

        CsvDocument.ReadRowsReusable(
            reader,
            (header, values) =>
            {
                headers.Add(header.ToArray());
                rows.Add(values.ToArray());
            },
            new CsvLoadOptions {
                DetectDelimiter = true,
                SkipInitialRecords = 1
            });

        var header = Assert.Single(headers);
        var row = Assert.Single(rows);
        Assert.Equal(new[] { "Name", "Value" }, header);
        Assert.Equal(new[] { "Alpha", "1" }, row);
    }

    [Fact]
    public void ReadRowsReusable_DetectDelimiterSkipsQuotedMultilineInitialRecord()
    {
        var headers = new List<string[]>();
        var rows = new List<string[]>();
        using var reader = new StringReader("\"metadata\nstill,has,commas\"\nName;Value\nAlpha;1\n");

        CsvDocument.ReadRowsReusable(
            reader,
            (header, values) =>
            {
                headers.Add(header.ToArray());
                rows.Add(values.ToArray());
            },
            new CsvLoadOptions {
                DetectDelimiter = true,
                SkipInitialRecords = 1
            });

        var parsedHeader = Assert.Single(headers);
        var row = Assert.Single(rows);
        Assert.Equal(new[] { "Name", "Value" }, parsedHeader);
        Assert.Equal(new[] { "Alpha", "1" }, row);
    }

    [Fact]
    public void ReadRecordsReusable_CanSkipInitialRecords()
    {
        var records = new List<string[]>();
        using var reader = new StringReader("metadata\nName,Value\nAlpha,1\n");

        CsvDocument.ReadRecordsReusable(
            reader,
            values => records.Add(values.ToArray()),
            new CsvLoadOptions { SkipInitialRecords = 1 });

        Assert.Equal(2, records.Count);
        Assert.Equal(new[] { "Name", "Value" }, records[0]);
        Assert.Equal(new[] { "Alpha", "1" }, records[1]);
    }

#if NET8_0_OR_GREATER
    [Fact]
    public void ReadRowFieldSpans_StreamsHeaderAndRows()
    {
        var events = new List<string>();
        var visitor = new CapturingRowFieldSpanVisitor(events);
        using var reader = new StringReader("Name,Value\nA,1\nB,2\n");

        CsvDocument.ReadRowFieldSpans(reader, ref visitor);

        Assert.Equal(
            new[]
            {
                "begin:0:Name|Value",
                "field:0:0:A",
                "field:0:1:1",
                "end:0:2",
                "begin:1:Name|Value",
                "field:1:0:B",
                "field:1:1:2",
                "end:1:2"
            },
            events);
    }

    [Fact]
    public void ReadRowFieldSpansFromText_StreamsMultilineRows()
    {
        var events = new List<string>();
        var visitor = new CapturingRowFieldSpanVisitor(events);

        CsvDocument.ReadRowFieldSpansFromText("Name,Note\nAlpha,\"one\ntwo\"\n", ref visitor);

        Assert.Equal(
            new[]
            {
                "begin:0:Name|Note",
                "field:0:0:Alpha",
                "field:0:1:one\ntwo",
                "end:0:2"
            },
            events);
    }

    [Fact]
    public void ReadRowFieldSpansFromText_SkipsPreHeaderCommentsByDefault()
    {
        var events = new List<string>();
        var visitor = new CapturingRowFieldSpanVisitor(events);

        CsvDocument.ReadRowFieldSpansFromText("#Version: 1.0\nName,Value\nAlpha,1\n", ref visitor);

        Assert.Equal(
            new[]
            {
                "begin:0:Name|Value",
                "field:0:0:Alpha",
                "field:0:1:1",
                "end:0:2"
            },
            events);
    }

    [Fact]
    public void ReadRowFieldSpansFromText_SkipsCommentWithUnclosedQuoteBeforeHeader()
    {
        var events = new List<string>();
        var visitor = new CapturingRowFieldSpanVisitor(events);

        CsvDocument.ReadRowFieldSpansFromText("# generated \"by tool\nName,Value\nAlpha,1\n", ref visitor);

        Assert.Equal(
            new[]
            {
                "begin:0:Name|Value",
                "field:0:0:Alpha",
                "field:0:1:1",
                "end:0:2"
            },
            events);
    }

    [Fact]
    public void ReadRowFieldSpansFromText_SkipsMultilineCommentBeforeHeader()
    {
        var events = new List<string>();
        var visitor = new CapturingRowFieldSpanVisitor(events);

        CsvDocument.ReadRowFieldSpansFromText("#,\"note\ncontinued\"\nName,Value\nAlpha,1\n", ref visitor);

        Assert.Equal(
            new[]
            {
                "begin:0:Name|Value",
                "field:0:0:Alpha",
                "field:0:1:1",
                "end:0:2"
            },
            events);
    }

    [Fact]
    public void ReadRowFieldSpans_Appends_Static_Columns()
    {
        var events = new List<string>();
        var visitor = new CapturingRowFieldSpanVisitor(events);
        using var reader = new StringReader("Name,Value\nAlpha,1\n");

        CsvDocument.ReadRowFieldSpans(
            reader,
            ref visitor,
            new CsvLoadOptions {
                StaticColumns = new Dictionary<string, object?> {
                    ["SourceFile"] = "input.csv"
                }
            });

        Assert.Equal(
            new[]
            {
                "begin:0:Name|Value|SourceFile",
                "field:0:0:Alpha",
                "field:0:1:1",
                "field:0:2:input.csv",
                "end:0:3"
            },
            events);
    }

    [Fact]
    public void ReadRowFieldSpans_ReadsGZipByExtension()
    {
        var path = Path.Combine(Path.GetTempPath(), "OfficeIMO.CSV.RowFieldSpans." + Guid.NewGuid().ToString("N") + ".csv.gz");
        try
        {
            new CsvDocument()
                .WithHeader("Name", "Value")
                .AddRow("Alpha", 1)
                .Save(path, new CsvSaveOptions { NewLine = "\n" });

            var events = new List<string>();
            var visitor = new CapturingRowFieldSpanVisitor(events);

            CsvDocument.ReadRowFieldSpans(path, ref visitor);

            Assert.Equal(
                new[]
                {
                    "begin:0:Name|Value",
                    "field:0:0:Alpha",
                    "field:0:1:1",
                    "end:0:2"
                },
                events);
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
    public void ReadRowFieldSpans_ReadsSmallUncompressedFileWithMultilineField()
    {
        var path = Path.Combine(Path.GetTempPath(), "OfficeIMO.CSV.RowFieldSpans." + Guid.NewGuid().ToString("N") + ".csv");
        try
        {
            File.WriteAllText(path, "Name,Note,Value\nAlpha,\"one\ntwo\",1\nBeta,plain\n");
            var events = new List<string>();
            var visitor = new CapturingRowFieldSpanVisitor(events);

            CsvDocument.ReadRowFieldSpans(path, ref visitor);

            Assert.Equal(
                new[]
                {
                    "begin:0:Name|Note|Value",
                    "field:0:0:Alpha",
                    "field:0:1:one\ntwo",
                    "field:0:2:1",
                    "end:0:3",
                    "begin:1:Name|Note|Value",
                    "field:1:0:Beta",
                    "field:1:1:plain",
                    "field:1:2:",
                    "end:1:2"
                },
                events);
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
    public void ReadRowFieldSpansFromText_DetectDelimiterSkipsInitialRecordsAfterLeadingComments()
    {
        var events = new List<string>();
        var visitor = new CapturingRowFieldSpanVisitor(events);

        CsvDocument.ReadRowFieldSpansFromText(
            "#note\nmetadata,with,commas\nName;Value\nAlpha;1\n",
            ref visitor,
            new CsvLoadOptions {
                DetectDelimiter = true,
                SkipInitialRecords = 1
            });

        Assert.Equal(
            new[]
            {
                "begin:0:Name|Value",
                "field:0:0:Alpha",
                "field:0:1:1",
                "end:0:2"
            },
            events);
    }

    [Fact]
    public void ReadRowFieldSpansFromText_UsesExplicitHeader()
    {
        var events = new List<string>();
        var visitor = new CapturingRowFieldSpanVisitor(events);

        CsvDocument.ReadRowFieldSpansFromText(
            "Alpha;1\n",
            ref visitor,
            new CsvLoadOptions {
                Delimiter = ';',
                Header = new[] { "Name", "Value" }
            });

        Assert.Equal(
            new[]
            {
                "begin:0:Name|Value",
                "field:0:0:Alpha",
                "field:0:1:1",
                "end:0:2"
            },
            events);
    }

    [Fact]
    public void ReadRowFieldSpans_DetectDelimiterSkipsInitialRecordsAfterLeadingComments()
    {
        var events = new List<string>();
        var visitor = new CapturingRowFieldSpanVisitor(events);
        using var reader = new StringReader("#note\nmetadata,with,commas\nName;Value\nAlpha;1\n");

        CsvDocument.ReadRowFieldSpans(
            reader,
            ref visitor,
            new CsvLoadOptions {
                DetectDelimiter = true,
                SkipInitialRecords = 1
            });

        Assert.Equal(
            new[]
            {
                "begin:0:Name|Value",
                "field:0:0:Alpha",
                "field:0:1:1",
                "end:0:2"
            },
            events);
    }

    [Fact]
    public void ReadRowFieldSpans_RecognizesW3CFieldsHeader()
    {
        var events = new List<string>();
        var visitor = new CapturingRowFieldSpanVisitor(events);
        using var reader = new StringReader("#Version: 1.0\n#Fields: date time cs-uri\n2026-01-01 00:00 /index.html\n");

        CsvDocument.ReadRowFieldSpans(
            reader,
            ref visitor,
            new CsvLoadOptions { Delimiter = ' ' });

        Assert.Equal(
            new[]
            {
                "begin:0:date|time|cs-uri",
                "field:0:0:2026-01-01",
                "field:0:1:00:00",
                "field:0:2:/index.html",
                "end:0:3"
            },
            events);
    }

    [Fact]
    public void ReadRowFieldSpans_NoHeaderGeneratesDefaultHeaderAndPreservesFirstRow()
    {
        var events = new List<string>();
        var visitor = new CapturingRowFieldSpanVisitor(events);
        using var reader = new StringReader("A,1\nB,2\n");

        CsvDocument.ReadRowFieldSpans(
            reader,
            ref visitor,
            new CsvLoadOptions { HasHeaderRow = false });

        Assert.Equal(
            new[]
            {
                "begin:0:Column1|Column2",
                "field:0:0:A",
                "field:0:1:1",
                "end:0:2",
                "begin:1:Column1|Column2",
                "field:1:0:B",
                "field:1:1:2",
                "end:1:2"
            },
            events);
    }

    [Fact]
    public void ReadRowFieldSpans_StrictPolicyThrowsOnMismatchedDataRow()
    {
        var events = new List<string>();
        var visitor = new CapturingRowFieldSpanVisitor(events);
        using var reader = new StringReader("Name,Value\nAlpha,1,extra\n");

        Assert.Throws<CsvException>(() => CsvDocument.ReadRowFieldSpans(
            reader,
            ref visitor,
            new CsvLoadOptions { ColumnCountMismatchPolicy = CsvColumnCountMismatchPolicy.Strict }));
    }

    [Fact]
    public void ReadRowFieldSpans_PadsShortRowsByDefault()
    {
        var events = new List<string>();
        var visitor = new CapturingRowFieldSpanVisitor(events);
        using var reader = new StringReader("Name,Value\nAlpha\n");

        CsvDocument.ReadRowFieldSpans(reader, ref visitor);

        Assert.Equal(
            new[]
            {
                "begin:0:Name|Value",
                "field:0:0:Alpha",
                "field:0:1:",
                "end:0:1"
            },
            events);
    }

    [Fact]
    public void ReadFieldSpans_CanSkipInitialRecords()
    {
        var fields = new List<string>();
        using var reader = new StringReader("metadata\nName,Value\nAlpha,1\n");

        CsvDocument.ReadFieldSpans(
            reader,
            (recordIndex, fieldIndex, value) => fields.Add($"{recordIndex}:{fieldIndex}:{value.ToString()}"),
            new CsvLoadOptions { SkipInitialRecords = 1 });

        Assert.Equal(
            new[] { "0:0:Name", "0:1:Value", "1:0:Alpha", "1:1:1" },
            fields);
    }

    [Fact]
    public void ReadFieldSpans_ReadsWideUnquotedRecords()
    {
        var headers = string.Join(",", Enumerable.Range(0, 12).Select(index => $"H{index}"));
        var values = Enumerable.Range(0, 12).Select(index => new string((char)('A' + index), 40) + index).ToArray();
        var fields = new List<string>();
        using var reader = new StringReader(headers + "\n" + string.Join(",", values) + "\n");

        CsvDocument.ReadFieldSpans(
            reader,
            (recordIndex, fieldIndex, value) => fields.Add($"{recordIndex}:{fieldIndex}:{value.ToString()}"),
            new CsvLoadOptions { SkipInitialRecords = 1 });

        Assert.Equal(
            values.Select((value, index) => $"0:{index}:{value}").ToArray(),
            fields);
    }

    [Fact]
    public void ReadFieldSpansFromText_VectorizedUnquotedPathPreservesUnicodeCharacters()
    {
        const int rowCount = 5000;
        const string unicodeValue = "\u012C\u010A\u010D-value";
        var text = new StringBuilder("Name,Value\n");
        for (var i = 0; i < rowCount; i++)
        {
            text.Append(unicodeValue).Append(',').Append(i).Append('\n');
        }

        var fieldCount = 0;
        CsvDocument.ReadFieldSpansFromText(
            text.ToString(),
            (recordIndex, fieldIndex, value) =>
            {
                Assert.InRange(recordIndex, 0, rowCount - 1);
                if (fieldIndex == 0)
                {
                    Assert.Equal(unicodeValue, value.ToString());
                }
                else
                {
                    Assert.Equal(recordIndex.ToString(), value.ToString());
                }

                fieldCount++;
            },
            new CsvLoadOptions { SkipInitialRecords = 1 });

        Assert.Equal(rowCount * 2, fieldCount);
    }

    [Fact]
    public void ReadFieldSpansFromText_VectorizedQuotedPathGrowsForWideRecords()
    {
        var expected = Enumerable.Range(0, 40)
            .Select(index => index == 31 ? "line one\nline two" : $"value {index}")
            .ToArray();
        var text = string.Join(",", expected.Select(value => $"\"{value}\"")) + "\n";
        var actual = new List<string>();

        CsvDocument.ReadFieldSpansFromText(
            text,
            (recordIndex, fieldIndex, value) => actual.Add(value.ToString()));

        Assert.Equal(expected, actual);
    }

    [Fact]
    public void ReadFieldSpans_FallsBackForQuotedMultilineRecords()
    {
        var fields = new List<string>();
        using var reader = new StringReader("Name,Note\nAlpha,\"one\ntwo\"\n");

        CsvDocument.ReadFieldSpans(
            reader,
            (recordIndex, fieldIndex, value) => fields.Add($"{recordIndex}:{fieldIndex}:{value.ToString()}"),
            new CsvLoadOptions { SkipInitialRecords = 1 });

        Assert.Equal(new[] { "0:0:Alpha", "0:1:one\ntwo" }, fields);
    }

    [Fact]
    public void ReadFieldSpans_UnescapesQuotedFieldsWithoutMaterializedFallback()
    {
        var fields = new List<string>();
        using var reader = new StringReader("Name,Note\nAlpha,\"one \"\"quoted\"\" value\"\n");

        CsvDocument.ReadFieldSpans(
            reader,
            (recordIndex, fieldIndex, value) => fields.Add($"{recordIndex}:{fieldIndex}:{value.ToString()}"),
            new CsvLoadOptions { SkipInitialRecords = 1 });

        Assert.Equal(new[] { "0:0:Alpha", "0:1:one \"quoted\" value" }, fields);
    }

    [Fact]
    public void ReadFieldSpans_DoesNotEmitSkippedBlankLines()
    {
        var fields = new List<string>();
        using var reader = new StringReader("Name,Value\n\nAlpha,1\n");

        CsvDocument.ReadFieldSpans(
            reader,
            (recordIndex, fieldIndex, value) => fields.Add($"{recordIndex}:{fieldIndex}:{value.ToString()}"),
            new CsvLoadOptions { SkipInitialRecords = 1 });

        Assert.Equal(new[] { "0:0:Alpha", "0:1:1" }, fields);
    }

    [Fact]
    public void ReadFieldSpans_DoesNotEmitTrimmedWhitespaceOnlyLines()
    {
        var fields = new List<string>();
        using var reader = new StringReader("Name,Value\n   \t  \nAlpha,1\n");

        CsvDocument.ReadFieldSpans(
            reader,
            (recordIndex, fieldIndex, value) => fields.Add($"{recordIndex}:{fieldIndex}:{value.ToString()}"),
            new CsvLoadOptions {
                SkipInitialRecords = 1,
                TrimWhitespace = true
            });

        Assert.Equal(new[] { "0:0:Alpha", "0:1:1" }, fields);
    }

    [Fact]
    public void ReadFieldSpans_PreservesTrimmedWhitespaceDelimiterOnlyRows()
    {
        var fields = new List<string>();
        using var reader = new StringReader("Name\tValue\n \t \nAlpha\t1\n");

        CsvDocument.ReadFieldSpans(
            reader,
            (recordIndex, fieldIndex, value) => fields.Add($"{recordIndex}:{fieldIndex}:{value.ToString()}"),
            new CsvLoadOptions {
                Delimiter = '\t',
                SkipInitialRecords = 1,
                TrimWhitespace = true
            });

        Assert.Equal(
            new[] { "0:0:", "0:1:", "1:0:Alpha", "1:1:1" },
            fields);
    }

    [Fact]
    public void ReadFieldSpans_DetectsDelimiterAfterSkippedMetadata()
    {
        var fields = new List<string>();
        using var reader = new StringReader("metadata,with,commas\nName;Value\nAlpha;1\n");

        CsvDocument.ReadFieldSpans(
            reader,
            (recordIndex, fieldIndex, value) => fields.Add($"{recordIndex}:{fieldIndex}:{value.ToString()}"),
            new CsvLoadOptions {
                DetectDelimiter = true,
                SkipInitialRecords = 1
            });

        Assert.Equal(
            new[] { "0:0:Name", "0:1:Value", "1:0:Alpha", "1:1:1" },
            fields);
    }

    [Fact]
    public void ReadFieldSpansFromText_CanSkipInitialRecords()
    {
        var fields = new List<string>();

        CsvDocument.ReadFieldSpansFromText(
            "metadata\nName,Value\nAlpha,1\n",
            (recordIndex, fieldIndex, value) => fields.Add($"{recordIndex}:{fieldIndex}:{value.ToString()}"),
            new CsvLoadOptions { SkipInitialRecords = 1 });

        Assert.Equal(
            new[] { "0:0:Name", "0:1:Value", "1:0:Alpha", "1:1:1" },
            fields);
    }

    [Fact]
    public void ReadFieldSpansFromText_SkipsRawCommentsBeforeParsing()
    {
        var fields = new List<string>();

        CsvDocument.ReadFieldSpansFromText(
            "# generated \"by tool\nName,Value\n",
            (recordIndex, fieldIndex, value) => fields.Add($"{recordIndex}:{fieldIndex}:{value.ToString()}"),
            new CsvLoadOptions { SkipCommentRows = true });

        Assert.Equal(new[] { "0:0:Name", "0:1:Value" }, fields);
    }

    [Fact]
    public void ReadFieldSpansFromText_ReadsQuotedMultilineAndEscapedFields()
    {
        var fields = new List<string>();

        CsvDocument.ReadFieldSpansFromText(
            "Name,Note\nAlpha,\"one\n\"\"two\"\"\"\n",
            (recordIndex, fieldIndex, value) => fields.Add($"{recordIndex}:{fieldIndex}:{value.ToString()}"),
            new CsvLoadOptions { SkipInitialRecords = 1 });

        Assert.Equal(new[] { "0:0:Alpha", "0:1:one\n\"two\"" }, fields);
    }

    [Fact]
    public void ReadFieldSpansFromText_PreservesFlexibleMultilineQuotedParsing()
    {
        var fields = new List<string>();

        CsvDocument.ReadFieldSpansFromText(
            "Name,Note,Value\nA,b\"c\nd\",E\n",
            (recordIndex, fieldIndex, value) => fields.Add($"{recordIndex}:{fieldIndex}:{value.ToString()}"),
            new CsvLoadOptions { SkipInitialRecords = 1 });

        Assert.Equal(new[] { "0:0:A", "0:1:bc\nd", "0:2:E" }, fields);
    }

    [Fact]
    public void ReadFieldSpans_PreservesFlexibleMultilineQuotedParsing()
    {
        var fields = new List<string>();
        using var reader = new StringReader("Name,Note,Value\nA,b\"c\nd\",E\n");

        CsvDocument.ReadFieldSpans(
            reader,
            (recordIndex, fieldIndex, value) => fields.Add($"{recordIndex}:{fieldIndex}:{value.ToString()}"),
            new CsvLoadOptions { SkipInitialRecords = 1 });

        Assert.Equal(new[] { "0:0:A", "0:1:bc\nd", "0:2:E" }, fields);
    }

    [Fact]
    public void ReadFieldSpans_ReplaysContinuationsBeforeLenientFallback()
    {
        var fields = new List<string>();
        using var reader = new StringReader("Name,Note,Value\nA,\"b\nc\"x,D\n");

        CsvDocument.ReadFieldSpans(
            reader,
            (recordIndex, fieldIndex, value) => fields.Add($"{recordIndex}:{fieldIndex}:{value.ToString()}"),
            new CsvLoadOptions { SkipInitialRecords = 1 });

        Assert.Equal(new[] { "0:0:A", "0:1:b\ncx", "0:2:D" }, fields);
    }

    [Fact]
    public void ReadFieldSpansFromText_UsesLenientParsingAfterClosingQuotes()
    {
        var fields = new List<string>();

        CsvDocument.ReadFieldSpansFromText(
            "Name,Note\nAlpha,\"one\"two\n",
            (recordIndex, fieldIndex, value) => fields.Add($"{recordIndex}:{fieldIndex}:{value.ToString()}"),
            new CsvLoadOptions { SkipInitialRecords = 1 });

        Assert.Equal(new[] { "0:0:Alpha", "0:1:onetwo" }, fields);
    }

    [Fact]
    public void ReadFieldSpansFromText_CanVisitEscapedQuotedFieldsWithoutCompacting()
    {
        var fields = new List<string>();
        var visitor = new EscapedFieldCapturingVisitor(fields);

        CsvDocument.ReadFieldSpansFromText(
            "Name,Note\nAlpha,\"one \"\"quoted\"\" value\"\n",
            ref visitor,
            new CsvLoadOptions { SkipInitialRecords = 1 });

        Assert.Equal(
            new[] { "field:0:0:Alpha", "escaped:0:1:one \"\"quoted\"\" value:18" },
            fields);
    }

    [Fact]
    public void ReadFieldSpansFromText_DoesNotEmitSkippedBlankOrWhitespaceOnlyLines()
    {
        var fields = new List<string>();

        CsvDocument.ReadFieldSpansFromText(
            "Name,Value\n\n   \t  \nAlpha,1\n",
            (recordIndex, fieldIndex, value) => fields.Add($"{recordIndex}:{fieldIndex}:{value.ToString()}"),
            new CsvLoadOptions {
                SkipInitialRecords = 1,
                TrimWhitespace = true
            });

        Assert.Equal(new[] { "0:0:Alpha", "0:1:1" }, fields);
    }

    [Fact]
    public void ReadFieldSpansFromText_DetectsDelimiterAfterSkippedMetadata()
    {
        var fields = new List<string>();

        CsvDocument.ReadFieldSpansFromText(
            "metadata,with,commas\nName;Value\nAlpha;1\n",
            (recordIndex, fieldIndex, value) => fields.Add($"{recordIndex}:{fieldIndex}:{value.ToString()}"),
            new CsvLoadOptions {
                DetectDelimiter = true,
                SkipInitialRecords = 1
            });

        Assert.Equal(
            new[] { "0:0:Name", "0:1:Value", "1:0:Alpha", "1:1:1" },
            fields);
    }

    [Fact]
    public void ReadFieldSpansFromText_TrimWhitespaceAllowsSpacesAfterQuotedField()
    {
        var fields = new List<string>();

        CsvDocument.ReadFieldSpansFromText(
            "Name,Note,Value\nAlpha,\"one\"  ,1\n",
            (recordIndex, fieldIndex, value) => fields.Add($"{recordIndex}:{fieldIndex}:{value.ToString()}"),
            new CsvLoadOptions {
                SkipInitialRecords = 1,
                TrimWhitespace = true
            });

        Assert.Equal(new[] { "0:0:Alpha", "0:1:one", "0:2:1" }, fields);
    }

    [Fact]
    public void ReadFieldSpansFromText_ReadsEmptyFieldAfterQuotedField()
    {
        var fields = new List<string>();

        CsvDocument.ReadFieldSpansFromText(
            "Id,Name,Department,Region,Note,Empty,Value\n1,Alpha,Ops,EU,\"one\",,1\n",
            (recordIndex, fieldIndex, value) => fields.Add($"{recordIndex}:{fieldIndex}:{value.ToString()}"),
            new CsvLoadOptions { SkipInitialRecords = 1 });

        Assert.Equal(
            new[] { "0:0:1", "0:1:Alpha", "0:2:Ops", "0:3:EU", "0:4:one", "0:5:", "0:6:1" },
            fields);
    }

    [Fact]
    public void ReadFieldSpansFromText_RejectsUnexpectedCharacterAfterQuotedField()
    {
        Assert.Throws<CsvParseException>(() =>
            CsvDocument.ReadFieldSpansFromText(
                "Id,Name,Department,Region,Note,Value\n1,Alpha,Ops,EU,\"one\"x,1\n",
                static (_, _, _) => { },
                new CsvLoadOptions {
                    SkipInitialRecords = 1,
                    QuoteParsingMode = CsvQuoteParsingMode.Strict
                }));
    }

    [Fact]
    public void ReadFieldSpans_HonorsProjectedFieldVisitor()
    {
        var fields = new List<string>();
        var visitor = new ProjectedFieldCapturingVisitor(fields, projectedFieldIndex: 0);
        using var reader = new StringReader("Name,Note,Value\nAlpha,\"one \"\"quoted\"\" value\",1\nBeta,two,2\n");

        CsvDocument.ReadFieldSpans(
            reader,
            ref visitor,
            new CsvLoadOptions { SkipInitialRecords = 1 });

        Assert.Equal(new[] { "0:0:Alpha", "1:0:Beta" }, fields);
    }

    [Fact]
    public void ReadFieldSpansFromText_HonorsProjectedFieldVisitor()
    {
        var fields = new List<string>();
        var visitor = new ProjectedFieldCapturingVisitor(fields, projectedFieldIndex: 2);

        CsvDocument.ReadFieldSpansFromText(
            "Name,Note,Value\nAlpha,\"one \"\"quoted\"\" value\",1\nBeta,two,2\n",
            ref visitor,
            new CsvLoadOptions { SkipInitialRecords = 1 });

        Assert.Equal(new[] { "0:2:1", "1:2:2" }, fields);
    }

#endif

    [Fact]
    public void LoadFromStream_InMemoryMode_ParsesRows()
    {
        var bytes = Encoding.UTF8.GetBytes("Id,Name\n1,Alice\n2,Bob\n");
        using var stream = new MemoryStream(bytes, writable: false);

        var doc = CsvDocument.Load(stream, new CsvLoadOptions { Mode = CsvLoadMode.InMemory });
        var rows = doc.AsEnumerable().ToList();

        Assert.Equal(2, rows.Count);
        Assert.True(stream.CanRead);
    }

    [Fact]
    public void LoadFromStream_StreamMode_CanReenumerateRows()
    {
        var bytes = Encoding.UTF8.GetBytes("Id,Name\n1,Alice\n2,Bob\n");
        using var stream = new MemoryStream(bytes, writable: false);

        var doc = CsvDocument.Load(stream, new CsvLoadOptions { Mode = CsvLoadMode.Stream });
        var firstPass = doc.AsEnumerable().Select(r => r.AsString("Name")).ToArray();
        var secondPass = doc.AsEnumerable().Select(r => r.AsString("Name")).ToArray();

        Assert.Equal(new[] { "Alice", "Bob" }, firstPass);
        Assert.Equal(new[] { "Alice", "Bob" }, secondPass);
        Assert.True(stream.CanRead);
    }

    [Fact]
    public void LoadFromStream_LeavesSourceOpenAndRestoresSeekablePosition()
    {
        var bytes = Encoding.UTF8.GetBytes("Id,Name\n1,Alice\n");
        var stream = new MemoryStream(bytes, writable: false);
        stream.Position = stream.Length;

        var doc = CsvDocument.Load(stream, new CsvLoadOptions { Mode = CsvLoadMode.Stream });
        var rows = doc.AsEnumerable().ToList();

        Assert.Single(rows);
        Assert.True(stream.CanRead);
        Assert.Equal(stream.Length, stream.Position);
        stream.Dispose();
    }

    [Fact]
    public void LoadFromStream_StreamMode_SupportsNonSeekableSource()
    {
        var bytes = Encoding.UTF8.GetBytes("Id,Name\n1,Alice\n2,Bob\n");
        using var stream = new NonSeekableReadStream(bytes);

        var doc = CsvDocument.Load(stream, new CsvLoadOptions { Mode = CsvLoadMode.Stream });
        var firstPass = doc.AsEnumerable().Select(r => r.AsString("Name")).ToArray();
        var secondPass = doc.AsEnumerable().Select(r => r.AsString("Name")).ToArray();

        Assert.Equal(new[] { "Alice", "Bob" }, firstPass);
        Assert.Equal(new[] { "Alice", "Bob" }, secondPass);
    }

    private sealed class NonSeekableReadStream : Stream
    {
        private readonly Stream _inner;

        public NonSeekableReadStream(byte[] bytes)
        {
            _inner = new MemoryStream(bytes, writable: false);
        }

        public override bool CanRead => _inner.CanRead;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position
        {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        public override void Flush() => _inner.Flush();
        public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _inner.Dispose();
            }

            base.Dispose(disposing);
        }
    }

#if NET8_0_OR_GREATER
    private readonly struct CapturingRowFieldSpanVisitor : ICsvRowFieldSpanVisitor
    {
        private readonly List<string> _events;

        public CapturingRowFieldSpanVisitor(List<string> events)
        {
            _events = events;
        }

        public void BeginRow(IReadOnlyList<string> header, int rowIndex)
        {
            _events.Add($"begin:{rowIndex}:{string.Join("|", header)}");
        }

        public void VisitField(int rowIndex, int fieldIndex, ReadOnlySpan<char> value)
        {
            _events.Add($"field:{rowIndex}:{fieldIndex}:{value.ToString()}");
        }

        public void EndRow(int rowIndex, int fieldCount)
        {
            _events.Add($"end:{rowIndex}:{fieldCount}");
        }
    }

    private readonly struct EscapedFieldCapturingVisitor : ICsvFieldSpanVisitor
    {
        private readonly List<string> _events;

        public EscapedFieldCapturingVisitor(List<string> events)
        {
            _events = events;
        }

        public void VisitField(int recordIndex, int fieldIndex, ReadOnlySpan<char> value)
        {
            _events.Add($"field:{recordIndex}:{fieldIndex}:{value.ToString()}");
        }

        public bool TryVisitEscapedField(int recordIndex, int fieldIndex, ReadOnlySpan<char> escapedValue, int unescapedLength)
        {
            _events.Add($"escaped:{recordIndex}:{fieldIndex}:{escapedValue.ToString()}:{unescapedLength}");
            return true;
        }
    }

    private readonly struct ProjectedFieldCapturingVisitor : ICsvProjectedFieldSpanVisitor
    {
        private readonly List<string> _events;
        private readonly int _projectedFieldIndex;

        public ProjectedFieldCapturingVisitor(List<string> events, int projectedFieldIndex)
        {
            _events = events;
            _projectedFieldIndex = projectedFieldIndex;
        }

        public bool ShouldVisitField(int recordIndex, int fieldIndex) => fieldIndex == _projectedFieldIndex;

        public void VisitField(int recordIndex, int fieldIndex, ReadOnlySpan<char> value)
        {
            _events.Add($"{recordIndex}:{fieldIndex}:{value.ToString()}");
        }

        public bool TryVisitEscapedField(int recordIndex, int fieldIndex, ReadOnlySpan<char> escapedValue, int unescapedLength)
        {
            _events.Add($"{recordIndex}:{fieldIndex}:{escapedValue.ToString()}:{unescapedLength}");
            return true;
        }
    }
#endif
}
