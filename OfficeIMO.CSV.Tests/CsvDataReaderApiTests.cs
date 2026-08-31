using System;
using System.Data.Common;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
#if NET8_0_OR_GREATER
using System.Threading.Tasks;
#endif
using OfficeIMO.CSV;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public sealed class CsvDataReaderApiTests {
    [Fact]
    public void DataReaderApi_UsesOpenForSourcesAndCreateForLoadedDocuments() {
        Assert.DoesNotContain(
            typeof(CsvDocument).Assembly.GetExportedTypes(),
            static type => type.Name.EndsWith("DataReader", StringComparison.Ordinal));

        MethodInfo[] methods = typeof(CsvDocument).GetMethods(
            BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static | BindingFlags.DeclaredOnly);

        Assert.Contains(methods, static method =>
            method.Name == "OpenDataReader"
            && method.IsStatic
            && method.GetParameters().FirstOrDefault()?.ParameterType == typeof(string)
            && method.ReturnType == typeof(DbDataReader));
        Assert.Contains(methods, static method =>
            method.Name == "OpenDataReader"
            && method.IsStatic
            && method.GetParameters().FirstOrDefault()?.ParameterType == typeof(Stream)
            && method.ReturnType == typeof(DbDataReader));
        Assert.Contains(methods, static method =>
            method.Name == "CreateDataReader"
            && !method.IsStatic
            && method.ReturnType == typeof(DbDataReader));
        Assert.DoesNotContain(methods, static method =>
            method.Name == "CreateDataReader" && method.IsStatic);
        Assert.DoesNotContain(
            typeof(CsvDocument).Assembly.GetExportedTypes(),
            static type => type.Name == "CsvLoadMode");
        Assert.Null(typeof(CsvLoadOptions).GetProperty("Mode", BindingFlags.Public | BindingFlags.Instance));
        Assert.Null(typeof(CsvDocument).GetProperty("Mode", BindingFlags.Public | BindingFlags.Instance));
    }

#if NET8_0_OR_GREATER
    [Fact]
    public async Task OpenDataReaderAsyncUsesAsyncIoAndReturnsMemoryBackedAsyncCursor() {
        byte[] bytes = Encoding.UTF8.GetBytes("Id,Name\n1,Alpha\n2,Beta\n");
        await using var stream = new AsyncOnlyReadStream(bytes);

        using DbDataReader reader = await CsvDocument.OpenDataReaderAsync(
            stream,
            readerOptions: new CsvDataReaderOptions { InferSchema = true });

        Assert.True(stream.AsyncReadCount > 0);
        Assert.True(await reader.ReadAsync(CancellationToken.None));
        Assert.Equal(1, reader.GetInt32(0));
        Assert.Equal("Alpha", reader.GetString(1));
        Assert.True(await reader.ReadAsync(CancellationToken.None));
        Assert.Equal("Beta", reader.GetString(1));
        Assert.False(await reader.ReadAsync(CancellationToken.None));
        Assert.False(stream.IsDisposed);
    }

    [Fact]
    public async Task OpenDataReaderAsyncReadsFromCurrentSeekablePositionAndRestoresIt() {
        byte[] prefix = Encoding.UTF8.GetBytes("ignored-prefix");
        byte[] payload = Encoding.UTF8.GetBytes("Id,Name\n1,Ada\n");
        using var stream = new MemoryStream(prefix.Concat(payload).ToArray());
        stream.Position = prefix.Length;

        using DbDataReader reader = await CsvDocument.OpenDataReaderAsync(
            stream,
            new CsvLoadOptions { MaxInputBytes = payload.Length },
            new CsvDataReaderOptions { InferSchema = true });

        Assert.Equal(prefix.Length, stream.Position);
        Assert.Equal("Id", reader.GetName(0));
        Assert.Equal("Name", reader.GetName(1));
        Assert.True(await reader.ReadAsync(CancellationToken.None));
        Assert.Equal(1, reader.GetInt32(0));
        Assert.Equal("Ada", reader.GetString(1));
        Assert.False(await reader.ReadAsync(CancellationToken.None));
    }

    [Fact]
    public async Task OpenDataReaderAsyncDoesNotRetainOpeningCancellationInReturnedCursor() {
        using var openingCancellation = new CancellationTokenSource();
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes("Id,Name\n1,Ada\n"));

        using DbDataReader reader = await CsvDocument.OpenDataReaderAsync(
            stream,
            new CsvLoadOptions { CancellationToken = openingCancellation.Token });
        openingCancellation.Cancel();

        Assert.True(await reader.ReadAsync(CancellationToken.None));
        Assert.Equal("1", reader.GetString(0));
        Assert.Equal("Ada", reader.GetString(1));
    }
#endif

    [Fact]
    public void OpenDataReader_StreamSupportsTypedGettersAndLeavesStreamOpen() {
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes("Id,Name\n1,Ada\n"));
        using (DbDataReader reader = CsvDocument.OpenDataReader(
            stream,
            new CsvLoadOptions { Mode = CsvLoadMode.Stream },
            new CsvDataReaderOptions { InferSchema = true })) {
            Assert.True(reader.Read());
            Assert.Equal(1, reader.GetInt32(0));
            Assert.Equal("Ada", reader.GetString(1));
            Assert.False(reader.Read());
        }

        Assert.True(stream.CanRead);
    }

    [Fact]
    public void OpenDataReader_HandlesQuotedMultilineEscapedAndLongFields() {
        string longValue = new('x', 200_000);
        string csv = "Id,Description,LongValue\n"
            + "1,\"line one\nline \"\"two\"\"\",\"" + longValue + "\"\n";
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(csv));

        using DbDataReader reader = CsvDocument.OpenDataReader(stream);

        Assert.True(reader.Read());
        Assert.Equal(1, reader.GetInt32(0));
        Assert.Equal("line one\nline \"two\"", reader.GetString(1));
        Assert.Equal(longValue, reader.GetString(2));
        Assert.False(reader.Read());
    }

    [Fact]
    public void OpenDataReader_DoesNotLeakCrLfAcrossLargeUnquotedBufferBoundaries() {
        var csv = new StringBuilder(5_000_000);
        csv.Append("Id,DisplayName,Score,CreatedUtc\r\n");
        for (int id = 1; id <= 100_000; id++) {
            csv.Append(id)
                .Append(",Row ")
                .Append(id)
                .Append(',')
                .Append((id * 1.25m).ToString("F2", CultureInfo.InvariantCulture))
                .Append(",08/09/2026 09:31:27\r\n");
        }

        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.CSV.BufferBoundary.{Guid.NewGuid():N}.csv");
        try {
            File.WriteAllText(path, csv.ToString(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
            using DbDataReader reader = CsvDocument.OpenDataReader(path);

            int expectedId = 0;
            while (reader.Read()) {
                expectedId++;
                Assert.Equal(expectedId.ToString(), reader.GetString(0));
            }

            Assert.Equal(100_000, expectedId);
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OpenDataReader_PathPreservesUtf8FieldsWithOrWithoutPreamble(bool emitPreamble) {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.CSV.Utf8.{Guid.NewGuid():N}.csv");
        try {
            File.WriteAllText(
                path,
                "Id,Name,City\r\n1,Zażółć gęślą jaźń,東京\r\n2,😀,München\r\n",
                new UTF8Encoding(emitPreamble));

            using DbDataReader reader = CsvDocument.OpenDataReader(path);

            Assert.True(reader.Read());
            Assert.Equal("1", reader.GetString(0));
            Assert.Equal("Zażółć gęślą jaźń", reader.GetString(1));
            Assert.Equal("東京", reader.GetString(2));
            Assert.True(reader.Read());
            Assert.Equal("2", reader.GetString(0));
            Assert.Equal("😀", reader.GetString(1));
            Assert.Equal("München", reader.GetString(2));
            Assert.False(reader.Read());
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_PathFallsBackForUtf16Preamble() {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.CSV.Utf16.{Guid.NewGuid():N}.csv");
        try {
            File.WriteAllText(path, "Id,Name\r\n1,Łódź\r\n", Encoding.Unicode);

            using DbDataReader reader = CsvDocument.OpenDataReader(path);

            Assert.True(reader.Read());
            Assert.Equal("1", reader.GetString(0));
            Assert.Equal("Łódź", reader.GetString(1));
            Assert.False(reader.Read());
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_PathFallsBackAtOversizedUnquotedRecordBoundary() {
        string longValue = new('x', 600_000);
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.CSV.Oversized.{Guid.NewGuid():N}.csv");
        try {
            File.WriteAllText(
                path,
                "Id,Value\n1,short\n2," + longValue + "\n3,last\n",
                new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));

            using DbDataReader reader = CsvDocument.OpenDataReader(path);

            Assert.True(reader.Read());
            Assert.Equal("short", reader.GetString(1));
            Assert.True(reader.Read());
            Assert.Equal(longValue, reader.GetString(1));
            Assert.True(reader.Read());
            Assert.Equal("last", reader.GetString(1));
            Assert.False(reader.Read());
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_PathUtf8FastRowsEnforceStrictColumnCount() {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.CSV.Strict.{Guid.NewGuid():N}.csv");
        try {
            File.WriteAllText(path, "First,Second\n1,2\n3\n", new UTF8Encoding(false));
            using DbDataReader reader = CsvDocument.OpenDataReader(
                path,
                new CsvLoadOptions {
                    ColumnCountMismatchPolicy = CsvColumnCountMismatchPolicy.Strict
                });

            Assert.True(reader.Read());
            Assert.Equal("1", reader.GetString(0));
            CsvException exception = Assert.Throws<CsvException>(() => reader.Read());
            Assert.Contains("header defines 2 columns", exception.Message);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_PathFallbackDoesNotTreatLaterCommentAsPreHeaderComment() {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.CSV.CommentData.{Guid.NewGuid():N}.csv");
        try {
            File.WriteAllText(path, "Value\nfirst\n# literal data\nlast\n", new UTF8Encoding(false));

            using DbDataReader reader = CsvDocument.OpenDataReader(path);

            Assert.True(reader.Read());
            Assert.Equal("first", reader.GetString(0));
            Assert.True(reader.Read());
            Assert.Equal("# literal data", reader.GetString(0));
            Assert.True(reader.Read());
            Assert.Equal("last", reader.GetString(0));
            Assert.False(reader.Read());
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData('\n')]
    [InlineData('\r')]
    public void OpenDataReader_PathFallsBackForRecordSeparatorDelimiter(char delimiter) {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.CSV.RecordSeparatorDelimiter.{Guid.NewGuid():N}.csv");
        try {
            File.WriteAllText(
                path,
                $"Name{delimiter}Ada{delimiter}",
                new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));

            using DbDataReader reader = CsvDocument.OpenDataReader(
                path,
                new CsvLoadOptions { Delimiter = delimiter });

            Assert.Equal(1, reader.FieldCount);
            Assert.Equal("Name", reader.GetName(0));
            Assert.True(reader.Read());
            Assert.Equal("Ada", reader.GetString(0));
            Assert.False(reader.Read());
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_PathAlreadyCancelledReleasesFile() {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.CSV.CancelledPath.{Guid.NewGuid():N}.csv");
        File.WriteAllText(path, "Name\nAda\n", new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        try {
            Assert.Throws<OperationCanceledException>(() =>
                CsvDocument.OpenDataReader(
                    path,
                    new CsvLoadOptions { CancellationToken = cancellation.Token }));

            File.Delete(path);
            Assert.False(File.Exists(path));
        } finally {
            if (File.Exists(path)) {
                File.Delete(path);
            }
        }
    }

#if NET8_0_OR_GREATER
    [Theory]
    [InlineData("\n")]
    [InlineData("\r\n")]
    public void OpenDataReader_ReportsLogicalRecordsAndStreamingPhysicalLines(string newLine) {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.CSV.Position.{Guid.NewGuid():N}.csv");
        try {
            string csv = string.Join(newLine, new[] {
                "Id,Value",
                "1,one",
                "2,\"line one",
                "line two\"",
                "# ignored",
                "3,three"
            });
            File.WriteAllText(path, csv);
            using DbDataReader reader = CsvDocument.OpenDataReader(
                path,
                new CsvLoadOptions { SkipCommentRows = true });
            ICsvDataReaderPositionMetadata position =
                Assert.IsAssignableFrom<ICsvDataReaderPositionMetadata>(reader);

            Assert.Equal(0, position.RecordNumber);
            Assert.Null(position.PhysicalLineNumber);

            Assert.True(reader.Read());
            Assert.Equal(1, position.RecordNumber);
            Assert.Equal(2, position.PhysicalLineNumber);
            Assert.Equal(2, position.PhysicalEndLineNumber);

            Assert.True(reader.Read());
            Assert.Equal(2, position.RecordNumber);
            Assert.Equal(3, position.PhysicalLineNumber);
            Assert.Equal(4, position.PhysicalEndLineNumber);

            Assert.True(reader.Read());
            Assert.Equal(3, position.RecordNumber);
            Assert.Equal(6, position.PhysicalLineNumber);
            Assert.Equal(6, position.PhysicalEndLineNumber);

            Assert.False(reader.Read());
            Assert.Equal(0, position.RecordNumber);
            Assert.Null(position.PhysicalLineNumber);
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("\n")]
    [InlineData("\r\n")]
    public void OpenDataReader_ReportsMultilineEndPositionAtEndOfFile(string newLine) {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.CSV.PositionEof.{Guid.NewGuid():N}.csv");
        try {
            File.WriteAllText(path, $"Id,Value{newLine}1,\"first{newLine}second\"");
            using DbDataReader reader = CsvDocument.OpenDataReader(path);
            ICsvDataReaderPositionMetadata position =
                Assert.IsAssignableFrom<ICsvDataReaderPositionMetadata>(reader);

            Assert.True(reader.Read());
            Assert.Equal(2, position.PhysicalLineNumber);
            Assert.Equal(3, position.PhysicalEndLineNumber);
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("\n")]
    [InlineData("\r\n")]
    public void OpenDataReader_PreservesPhysicalLinesForReplayedCommentLookahead(string newLine) {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.CSV.PositionComment.{Guid.NewGuid():N}.csv");
        try {
            File.WriteAllText(path, $"Id,Value{newLine}# \"unterminated{newLine}1,one{newLine}");
            using DbDataReader reader = CsvDocument.OpenDataReader(
                path,
                new CsvLoadOptions { SkipCommentRows = true });
            ICsvDataReaderPositionMetadata position =
                Assert.IsAssignableFrom<ICsvDataReaderPositionMetadata>(reader);

            Assert.True(reader.Read());
            Assert.Equal("1", reader.GetString(0));
            Assert.Equal(3, position.PhysicalLineNumber);
            Assert.Equal(3, position.PhysicalEndLineNumber);
        } finally {
            File.Delete(path);
        }
    }
#endif

    [Fact]
    public void OpenDataReader_StringColumnsSupportTypedGettersWithoutSchemaInference() {
        Guid identifier = Guid.NewGuid();
        const string csv =
            "Boolean,Byte,Int16,Int32,Int64,Float,Double,Decimal,Date,Guid\n"
            + "true,7,-12,42,9876543210,1.5,2.75,165258.24,2026-07-29,";
        using var stream = new MemoryStream(
            Encoding.UTF8.GetBytes(csv + identifier.ToString("D") + "\n"));

        using DbDataReader reader = CsvDocument.OpenDataReader(stream);

        Assert.True(reader.Read());
        Assert.True(reader.GetBoolean(0));
        Assert.Equal((byte)7, reader.GetByte(1));
        Assert.Equal((short)-12, reader.GetInt16(2));
        Assert.Equal(42, reader.GetInt32(3));
        Assert.Equal(9_876_543_210L, reader.GetInt64(4));
        Assert.Equal(1.5f, reader.GetFloat(5));
        Assert.Equal(2.75d, reader.GetDouble(6));
        Assert.Equal(165258.24m, reader.GetDecimal(7));
        Assert.Equal(new DateTime(2026, 7, 29), reader.GetDateTime(8));
        Assert.Equal(identifier, reader.GetGuid(9));
    }

#if NET6_0_OR_GREATER
    [Fact]
    public void OpenDataReader_ExplicitDateOnlyAndTimeOnlyGettersPreserveStringSchema() {
        using DbDataReader reader = CsvDocument.OpenTextDataReader(
            "Date,Time\n2026-08-06,14:35:12\n");

        Assert.True(reader.Read());
        Assert.Equal(typeof(string), reader.GetFieldType(0));
        Assert.Equal(typeof(string), reader.GetFieldType(1));
        Assert.Equal(new DateOnly(2026, 8, 6), reader.GetFieldValue<DateOnly>(0));
        Assert.Equal(new TimeOnly(14, 35, 12), reader.GetFieldValue<TimeOnly>(1));
    }
#endif

    [Fact]
    public void OpenDataReader_SeekableFallbackStartsAtCallerPosition() {
        byte[] prefix = Encoding.UTF8.GetBytes("ignored prefix");
        byte[] payload = Encoding.UTF8.GetBytes("Id;Name\n1;Ada\n");
        using var stream = new MemoryStream(prefix.Concat(payload).ToArray());
        stream.Position = prefix.Length;

        using DbDataReader reader = CsvDocument.OpenDataReader(
            stream,
            new CsvLoadOptions {
                Mode = CsvLoadMode.Stream,
                DetectDelimiter = true
            });

        ICsvDataReaderMetadata metadata = Assert.IsAssignableFrom<ICsvDataReaderMetadata>(reader);
        Assert.Equal(';', metadata.Delimiter);
        Assert.Equal(prefix.Length, stream.Position);
        Assert.Equal("Id", reader.GetName(0));
        Assert.Equal("Name", reader.GetName(1));
        Assert.True(reader.Read());
        Assert.Equal(1, reader.GetInt32(0));
        Assert.Equal("Ada", reader.GetString(1));
        Assert.False(reader.Read());
    }

    [Fact]
    public void OpenDataReader_ObservesCancellationDuringTraversal() {
        using var cancellation = new CancellationTokenSource();
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes("Id\n1\n2\n"));
        using DbDataReader reader = CsvDocument.OpenDataReader(
            stream,
            new CsvLoadOptions {
                Mode = CsvLoadMode.Stream,
                CancellationToken = cancellation.Token
            });

        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() => reader.Read());
    }

    [Fact]
    public void OpenDataReader_ObservesCancellationWhileScanningLargeRecord() {
        using var cancellation = new CancellationTokenSource();
        byte[] bytes = Encoding.UTF8.GetBytes(
            "Value\n" + new string('x', 600_000) + "\n");
        using var stream = new CancelingSeekableReadStream(
            bytes,
            cancellation,
            cancelOnReadCount: 2);
        using DbDataReader reader = CsvDocument.OpenDataReader(
            stream,
            new CsvLoadOptions {
                Mode = CsvLoadMode.Stream,
                CancellationToken = cancellation.Token
            });

        Assert.False(cancellation.IsCancellationRequested);
        Assert.Throws<OperationCanceledException>(() => reader.Read());
        Assert.Equal(2, stream.ReadCount);
    }

#if NET8_0_OR_GREATER
    [Fact]
    public async Task OpenDataReader_ReadAsyncObservesPerCallCancellationWhileScanningLargeRecord() {
        using var cancellation = new CancellationTokenSource();
        byte[] bytes = Encoding.UTF8.GetBytes(
            "Value\n" + new string('x', 600_000) + "\n");
        using var stream = new CancelingSeekableReadStream(
            bytes,
            cancellation,
            cancelOnReadCount: 2);
        using DbDataReader reader = CsvDocument.OpenDataReader(
            stream,
            new CsvLoadOptions { Mode = CsvLoadMode.Stream });

        Assert.False(cancellation.IsCancellationRequested);
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            reader.ReadAsync(cancellation.Token));
        Assert.Equal(2, stream.ReadCount);
    }

    [Fact]
    public async Task OpenDataReader_ReadAsyncObservesPerCallCancellationOnGeneralParserFallback() {
        using var cancellation = new CancellationTokenSource();
        byte[] bytes = Encoding.UTF8.GetBytes(
            "Value\n" + new string('x', 600_000) + "\n");
        using var stream = new CancelingSeekableReadStream(
            bytes,
            cancellation,
            cancelOnReadCount: 2);
        using DbDataReader reader = CsvDocument.OpenDataReader(
            stream,
            new CsvLoadOptions {
                Mode = CsvLoadMode.Stream,
                NormalizeQuotes = true
            });

        Assert.False(cancellation.IsCancellationRequested);
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            reader.ReadAsync(cancellation.Token));
        Assert.Equal(2, stream.ReadCount);
    }

    [Fact]
    public async Task OpenDataReader_ParallelReadAsyncPropagatesPerCallCancellationToSource() {
        using var cancellation = new CancellationTokenSource();
        byte[] bytes = Encoding.UTF8.GetBytes(
            "Id,Value\n1," + new string('x', 600_000) + "\n");
        using var stream = new CancelingSeekableReadStream(
            bytes,
            cancellation,
            cancelOnReadCount: 2);
        CsvSchema schema = new CsvSchemaBuilder()
            .Column("Id").AsInt32()
            .Column("Value").AsString()
            .Done()
            .Build();
        using DbDataReader reader = CsvDocument.OpenDataReader(
            stream,
            new CsvLoadOptions { Mode = CsvLoadMode.Stream },
            new CsvDataReaderOptions {
                Schema = schema,
                ParallelProcessing = new CsvDataReaderParallelOptions {
                    MaxDegreeOfParallelism = 2,
                    BatchSize = 1
                }
            });

        Assert.False(cancellation.IsCancellationRequested);
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            reader.ReadAsync(cancellation.Token));
        Assert.Equal(2, stream.ReadCount);
    }

    [Fact]
    public void CreateDataReader_CancellationInterruptsSchemaInference() {
        var csv = new StringBuilder("Id,Value\n");
        const int rowCount = 250_000;
        for (int row = 0; row < rowCount; row++) {
            csv.Append(row).Append(',').Append("value-").Append(row).Append('\n');
        }
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(csv.ToString()));
        CsvDocument document = CsvDocument.Load(
            stream,
            new CsvLoadOptions { Mode = CsvLoadMode.InMemory });
        using var cancellation = new CancellationTokenSource();
        using var startCancellation = new ManualResetEventSlim();
        var cancellationThread = new Thread(() => {
            startCancellation.Wait();
            Thread.Sleep(1);
            cancellation.Cancel();
        });
        cancellationThread.Start();
        try {
            Assert.False(cancellation.IsCancellationRequested);
            startCancellation.Set();
            Assert.ThrowsAny<OperationCanceledException>(() =>
                document.CreateDataReader(
                    new CsvDataReaderOptions {
                        InferSchema = true,
                        SchemaSampleSize = rowCount
                    },
                    cancellation.Token));
        } finally {
            cancellationThread.Join();
        }
    }

#endif

    [Fact]
    public void OpenDataReader_NonSeekableFallbackObservesCancellationWhileBuffering() {
        using var cancellation = new CancellationTokenSource();
        using var stream = new CancelingNonSeekableReadStream(
            Encoding.UTF8.GetBytes("1,Ada\n2,Grace\n"),
            cancellation,
            maximumReadSize: 4);

        Assert.Throws<OperationCanceledException>(() =>
            CsvDocument.OpenDataReader(
                stream,
                new CsvLoadOptions {
                    Mode = CsvLoadMode.Stream,
                    HasHeaderRow = false,
                    CancellationToken = cancellation.Token
                }));
        Assert.Equal(1, stream.ReadCount);
        Assert.True(stream.CanRead);
    }

    [Fact]
    public void MemoryBackedReaderObservesCancellationBetweenReadChunks() {
        using var cancellation = new CancellationTokenSource();
        using var reader = new CancelingTextReader(
            "Id,Name\n1,Ada\n",
            cancellation,
            maximumReadSize: 4);

        Assert.Throws<OperationCanceledException>(() =>
            CsvDocument.ReadAllTextWithCancellation(
                reader,
                cancellation.Token));
        Assert.Equal(1, reader.ReadCount);
    }

    [Fact]
    public void OpenDataReader_PathRejectsInputBeyondConfiguredLimit() {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.CSV.Limit.{Guid.NewGuid():N}.csv");
        try {
            File.WriteAllText(path, "Id,Name\n1,Ada\n");

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
                CsvDocument.OpenDataReader(
                    path,
                    new CsvLoadOptions {
                        Mode = CsvLoadMode.Stream,
                        MaxInputBytes = 4
                    }));

            Assert.Contains("configured maximum", exception.Message, StringComparison.OrdinalIgnoreCase);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_SeekableStreamRejectsUnreadInputBeyondConfiguredLimit() {
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes("Id,Name\n1,Ada\n"));
        stream.Position = 3;

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            CsvDocument.OpenDataReader(
                stream,
                new CsvLoadOptions {
                    Mode = CsvLoadMode.Stream,
                    MaxInputBytes = 4
                }));

        Assert.Contains("configured maximum", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(3, stream.Position);
        Assert.True(stream.CanRead);
    }

    [Fact]
    public void OpenDataReader_DefaultSeekableStreamReturnsBeforeReadingToEnd() {
        var csv = new StringBuilder("Id,Name\n");
        for (int index = 0; index < 100_000; index++) {
            csv.Append(index).Append(",Value").Append(index).Append('\n');
        }
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(csv.ToString()));

        using DbDataReader reader = CsvDocument.OpenDataReader(stream);

        Assert.True(stream.Position < stream.Length);
        Assert.True(reader.Read());
        Assert.Equal(0, reader.GetInt32(0));
        Assert.Equal("Value0", reader.GetString(1));
    }

    [Fact]
    public void OpenDataReader_SuppliedLoadOptionsCannotDisableStreaming() {
        var csv = new StringBuilder("Id,Name\n");
        for (int index = 0; index < 100_000; index++) {
            csv.Append(index).Append(",Value").Append(index).Append('\n');
        }
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(csv.ToString()));

        using DbDataReader reader = CsvDocument.OpenDataReader(
            stream,
            new CsvLoadOptions {
                Mode = CsvLoadMode.InMemory,
                Culture = System.Globalization.CultureInfo.InvariantCulture
            });

        Assert.True(stream.Position < stream.Length);
        Assert.True(reader.Read());
        Assert.Equal(0, reader.GetInt32(0));
        Assert.Equal("Value0", reader.GetString(1));
    }

    [Fact]
    public void OpenDataReader_DefaultNonSeekableStreamReturnsBeforeReadingToEnd() {
        using var stream = new SingleChunkNonSeekableReadStream(
            Encoding.UTF8.GetBytes("# metadata\nId,Name\n1,Ada\n"));

        using DbDataReader reader = CsvDocument.OpenDataReader(stream);

        Assert.Equal(1, stream.ReadCount);
        Assert.True(reader.Read());
        Assert.Equal(1, reader.GetInt32(0));
        Assert.Equal("Ada", reader.GetString(1));
        Assert.Equal(1, stream.ReadCount);
    }

#if NET8_0_OR_GREATER
    private sealed class AsyncOnlyReadStream : Stream {
        private readonly byte[] _bytes;
        private int _position;

        internal AsyncOnlyReadStream(byte[] bytes) => _bytes = bytes;

        internal int AsyncReadCount { get; private set; }
        internal bool IsDisposed { get; private set; }
        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position {
            get => _position;
            set => throw new NotSupportedException();
        }

        public override void Flush() { }
        public override int Read(byte[] buffer, int offset, int count) =>
            throw new InvalidOperationException("The asynchronous reader must not use synchronous source I/O.");

        public override Task<int> ReadAsync(
            byte[] buffer,
            int offset,
            int count,
            CancellationToken cancellationToken) {
            cancellationToken.ThrowIfCancellationRequested();
            AsyncReadCount++;
            int copied = Math.Min(count, _bytes.Length - _position);
            if (copied > 0) {
                Array.Copy(_bytes, _position, buffer, offset, copied);
                _position += copied;
            }
            return Task.FromResult(copied);
        }

        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        protected override void Dispose(bool disposing) {
            IsDisposed = true;
            base.Dispose(disposing);
        }
    }
#endif

    private sealed class SingleChunkNonSeekableReadStream : Stream {
        private readonly byte[] _bytes;
        private bool _served;

        internal SingleChunkNonSeekableReadStream(byte[] bytes) {
            _bytes = bytes;
        }

        internal int ReadCount { get; private set; }

        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        public override void Flush() {
        }

        public override int Read(byte[] buffer, int offset, int count) {
            ReadCount++;
            if (_served) {
                throw new InvalidOperationException("The data reader attempted to drain the non-seekable source.");
            }

            _served = true;
            int copied = Math.Min(count, _bytes.Length);
            Array.Copy(_bytes, 0, buffer, offset, copied);
            return copied;
        }

        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }

    private sealed class CancelingNonSeekableReadStream : Stream {
        private readonly byte[] _bytes;
        private readonly CancellationTokenSource _cancellation;
        private readonly int _maximumReadSize;
        private int _position;

        internal CancelingNonSeekableReadStream(
            byte[] bytes,
            CancellationTokenSource cancellation,
            int maximumReadSize) {
            _bytes = bytes;
            _cancellation = cancellation;
            _maximumReadSize = maximumReadSize;
        }

        internal int ReadCount { get; private set; }

        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        public override void Flush() {
        }

        public override int Read(byte[] buffer, int offset, int count) {
            ReadCount++;
            int remaining = _bytes.Length - _position;
            int copied = Math.Min(Math.Min(count, _maximumReadSize), remaining);
            if (copied > 0) {
                Array.Copy(_bytes, _position, buffer, offset, copied);
                _position += copied;
                _cancellation.Cancel();
            }
            return copied;
        }

        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }

    private sealed class CancelingSeekableReadStream : Stream {
        private readonly MemoryStream _stream;
        private readonly CancellationTokenSource _cancellation;
        private readonly int _cancelOnReadCount;

        internal CancelingSeekableReadStream(
            byte[] bytes,
            CancellationTokenSource cancellation,
            int cancelOnReadCount) {
            _stream = new MemoryStream(bytes, writable: false);
            _cancellation = cancellation;
            _cancelOnReadCount = cancelOnReadCount;
        }

        internal int ReadCount { get; private set; }

        public override bool CanRead => true;
        public override bool CanSeek => true;
        public override bool CanWrite => false;
        public override long Length => _stream.Length;
        public override long Position {
            get => _stream.Position;
            set => _stream.Position = value;
        }

        public override void Flush() {
        }

        public override int Read(byte[] buffer, int offset, int count) {
            ReadCount++;
            int copied = _stream.Read(buffer, offset, count);
            if (ReadCount == _cancelOnReadCount) {
                _cancellation.Cancel();
            }

            return copied;
        }

        public override long Seek(long offset, SeekOrigin origin) =>
            _stream.Seek(offset, origin);

        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        protected override void Dispose(bool disposing) {
            if (disposing) {
                _stream.Dispose();
            }

            base.Dispose(disposing);
        }
    }

    private sealed class CancelingTextReader : TextReader {
        private readonly string _text;
        private readonly CancellationTokenSource _cancellation;
        private readonly int _maximumReadSize;
        private int _position;

        internal CancelingTextReader(
            string text,
            CancellationTokenSource cancellation,
            int maximumReadSize) {
            _text = text;
            _cancellation = cancellation;
            _maximumReadSize = maximumReadSize;
        }

        internal int ReadCount { get; private set; }

        public override int Read(char[] buffer, int index, int count) {
            ReadCount++;
            int copied = Math.Min(
                Math.Min(count, _maximumReadSize),
                _text.Length - _position);
            if (copied > 0) {
                _text.CopyTo(_position, buffer, index, copied);
                _position += copied;
                _cancellation.Cancel();
            }

            return copied;
        }
    }
}
