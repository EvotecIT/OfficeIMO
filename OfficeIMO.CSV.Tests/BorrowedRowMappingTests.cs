#if NET10_0_OR_GREATER
using System;
using System.Data;
using System.IO;
using System.Text;
using System.Threading;
using OfficeIMO.Data;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public sealed class BorrowedRowMappingTests {
    [Fact]
    public void FactoryMapsBorrowedUtf8AndNativeValuesInSourceOrder() {
        using var input = new MemoryStream(Encoding.UTF8.GetBytes(
            "Id,Value,Name,Date\n1,1.25,Zażółć 🚀,2024-01-02\n2,2.5,second,2024-01-03\n"));
        using var reader = CsvDocument.OpenDataReader(input);
        int name = reader.GetOrdinal("Name"), id = reader.GetOrdinal("Id");
        int date = reader.GetOrdinal("Date"), value = reader.GetOrdinal("Value");
        int count = 0;
        foreach (BorrowedRecord row in reader.RowsAsBorrowed<BorrowedRecord>(record => {
            Assert.True(record.TryGetUtf8Text(name, out ReadOnlySpan<byte> text));
            return new BorrowedRecord(text, record.GetInt32(id), record.GetDateTime(date), record.GetDouble(value));
        })) {
            count++;
            Assert.Equal(count == 1 ? "Zażółć 🚀" : "second", Encoding.UTF8.GetString(row.Name));
            Assert.Equal(count, row.Id);
            Assert.Equal(new DateTime(2024, 1, count + 1), row.Date);
            Assert.Equal(count == 1 ? 1.25 : 2.5, row.Value);
        }
        Assert.Equal(2, count);
        Assert.False(reader.IsClosed);
        Assert.True(input.CanRead);
    }

    [Fact]
    public void CurrentIsCachedAndEarlyDisposalKeepsCallerReaderOpen() {
        using var input = new MemoryStream(Encoding.UTF8.GetBytes("Id\n1\n2\n3\n"));
        using var reader = CsvDocument.OpenDataReader(input);
        int calls = 0;
        var enumerator = reader.RowsAsBorrowed<BorrowedRecord>(record => {
            calls++;
            return new BorrowedRecord(default, record.GetInt32(0), default, default);
        }).GetEnumerator();
        Assert.IsType<InvalidOperationException>(CurrentFailure(ref enumerator));
        Assert.True(enumerator.MoveNext());
        Assert.Equal(1, enumerator.Current.Id);
        Assert.Equal(1, enumerator.Current.Id);
        Assert.Equal(1, calls);
        Assert.True(enumerator.MoveNext());
        Assert.Equal(2, enumerator.Current.Id);
        Assert.Equal(2, enumerator.Current.Id);
        Assert.Equal(2, calls);
        enumerator.Dispose();
        Assert.IsType<InvalidOperationException>(CurrentFailure(ref enumerator));
        Assert.False(enumerator.MoveNext());
        Assert.Equal(2, calls);
        Assert.False(reader.IsClosed);
        Assert.True(reader.Read());
        Assert.Equal(3, reader.GetInt32(0));
        Assert.True(input.CanRead);
    }

    [Fact]
    public void EnumerationStartsAtNextUnreadRowAndInvalidatesCurrentAtEnd() {
        using var reader = CsvDocument.OpenTextDataReader("Id\n1\n2\n");
        Assert.True(reader.Read());
        int calls = 0;
        var enumerator = reader.RowsAsBorrowed<BorrowedRecord>(record => {
            calls++;
            return new BorrowedRecord(default, record.GetInt32(0), default, default);
        }).GetEnumerator();
        Assert.True(enumerator.MoveNext());
        Assert.Equal(2, enumerator.Current.Id);
        Assert.False(enumerator.MoveNext());
        Assert.False(enumerator.MoveNext());
        Assert.IsType<InvalidOperationException>(CurrentFailure(ref enumerator));
        Assert.Equal(1, calls);
        Assert.False(reader.IsClosed);
    }

    [Fact]
    public void ZeroColumnReaderDoesNotInvokeFactory() {
        using var reader = CsvDocument.OpenTextDataReader(string.Empty);
        int calls = 0;
        var enumerator = reader.RowsAsBorrowed<BorrowedRecord>(_ => {
            calls++;
            return default;
        }).GetEnumerator();
        Assert.Equal(0, reader.FieldCount);
        Assert.False(enumerator.MoveNext());
        Assert.IsType<InvalidOperationException>(CurrentFailure(ref enumerator));
        Assert.Equal(0, calls);
        Assert.False(reader.IsClosed);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void CancellationStopsBeforeAnotherSourceAdvanceOrFactoryCall(bool beforeFirstRow) {
        using var reader = CsvDocument.OpenTextDataReader("Id\n1\n2\n");
        using var cancellation = new CancellationTokenSource();
        int calls = 0;
        var enumerator = reader.RowsAsBorrowed<BorrowedRecord>(record => {
            calls++;
            return new BorrowedRecord(default, record.GetInt32(0), default, default);
        }, cancellation.Token).GetEnumerator();
        if (!beforeFirstRow) {
            Assert.True(enumerator.MoveNext());
            Assert.Equal(1, enumerator.Current.Id);
        }
        cancellation.Cancel();
        var failure = Assert.IsType<OperationCanceledException>(MoveFailure(ref enumerator));
        Assert.Equal(cancellation.Token, failure.CancellationToken);
        Assert.Equal(beforeFirstRow ? 0 : 1, calls);
        Assert.IsType<InvalidOperationException>(CurrentFailure(ref enumerator));
        Assert.False(enumerator.MoveNext());
        Assert.False(reader.IsClosed);
        Assert.True(reader.Read());
        Assert.Equal(beforeFirstRow ? 1 : 2, reader.GetInt32(0));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FactoryFailurePreservesOriginalExceptionAndEndsOnlyThisEnumeration(bool cancellationFailure) {
        using var reader = CsvDocument.OpenTextDataReader("Id\n1\n2\n");
        using var otherCancellation = new CancellationTokenSource();
        Exception original = cancellationFailure
            ? new OperationCanceledException("caller factory cancellation", new Exception("inner"), otherCancellation.Token)
            : new InvalidOperationException("caller factory failure", new Exception("inner"));
        int calls = 0;
        var enumerator = reader.RowsAsBorrowed<BorrowedRecord>(_ => {
            calls++;
            throw original;
        }).GetEnumerator();
        Assert.Same(original, MoveFailure(ref enumerator));
        Assert.IsType<InvalidOperationException>(CurrentFailure(ref enumerator));
        Assert.False(enumerator.MoveNext());
        Assert.Equal(1, calls);
        Assert.False(otherCancellation.IsCancellationRequested);
        Assert.False(reader.IsClosed);
        Assert.True(reader.Read());
        Assert.Equal(2, reader.GetInt32(0));
    }

    [Fact]
    public void CallerFactoryKeepsCanonicalSchemaDefaultsAndConvertedTextFallback() {
        CsvSchema schema = new CsvSchemaBuilder()
            .Column("Name").AsString().ConvertUsing(value => "mapped:" + value).Done()
            .Column("Id").AsInt32().WithDefault(42).Done().Build();
        using var reader = CsvDocument.OpenTextDataReader("Name,Id\nfirst,-1\n",
            new CsvLoadOptions { NullValue = "-1" }, new CsvDataReaderOptions { Schema = schema });
        int count = 0;
        foreach (BorrowedTextRecord row in reader.RowsAsBorrowed<BorrowedTextRecord>(record => {
            Assert.False(record.TryGetUtf8Text(0, out _));
            Assert.False(record.IsDBNull(1));
            return new BorrowedTextRecord(record.GetString(0).AsSpan(), record.GetInt32(1));
        })) {
            count++;
            Assert.Equal("mapped:first", row.Name.ToString());
            Assert.Equal(42, row.Id);
        }
        Assert.Equal(1, count);
    }

    private static Exception? MoveFailure<T>(ref BorrowedRowEnumerable<T>.Enumerator enumerator) where T : allows ref struct {
        try { enumerator.MoveNext(); return null; }
        catch (Exception exception) { return exception; }
    }

    private static Exception? CurrentFailure<T>(ref BorrowedRowEnumerable<T>.Enumerator enumerator) where T : allows ref struct {
        try { _ = enumerator.Current; return null; }
        catch (Exception exception) { return exception; }
    }

    private readonly ref struct BorrowedRecord {
        internal BorrowedRecord(ReadOnlySpan<byte> name, int id, DateTime date, double value) {
            Name = name; Id = id; Date = date; Value = value;
        }
        internal ReadOnlySpan<byte> Name { get; }
        internal int Id { get; }
        internal DateTime Date { get; }
        internal double Value { get; }
    }

    private readonly ref struct BorrowedTextRecord {
        internal BorrowedTextRecord(ReadOnlySpan<char> name, int id) { Name = name; Id = id; }
        internal ReadOnlySpan<char> Name { get; }
        internal int Id { get; }
    }
}
#endif
