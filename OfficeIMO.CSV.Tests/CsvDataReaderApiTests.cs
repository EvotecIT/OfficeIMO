using System;
using System.Data.Common;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
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
    }

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
        string longValue = new('x', 70_000);
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
    public void OpenDataReader_NonSeekableStreamReturnsBeforeReadingToEnd() {
        using var stream = new SingleChunkNonSeekableReadStream(
            Encoding.UTF8.GetBytes("# metadata\nId,Name\n1,Ada\n"));

        using DbDataReader reader = CsvDocument.OpenDataReader(
            stream,
            new CsvLoadOptions { Mode = CsvLoadMode.Stream });

        Assert.Equal(1, stream.ReadCount);
        Assert.True(reader.Read());
        Assert.Equal(1, reader.GetInt32(0));
        Assert.Equal("Ada", reader.GetString(1));
        Assert.Equal(1, stream.ReadCount);
    }

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
