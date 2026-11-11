using System.Text;
using Xunit;

namespace OfficeIMO.Toml.Tests;

// Loading contracts outside toml-test: limits, cancellation, encodings and the codes and positions of errors.
public class TomlLoadingTests {
    [Theory]
    [InlineData("a = [[[[1]]]]\n", nameof(TomlLoadOptions.MaxDepth))]
    [InlineData("a.b.c.d.e = 1\n", nameof(TomlLoadOptions.MaxDepth))]
    [InlineData("[a.b.c.d.e]\n", nameof(TomlLoadOptions.MaxDepth))]
    [InlineData("a = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]\n", nameof(TomlLoadOptions.MaxNodes))]
    [InlineData("a = \"0123456789\"\n", nameof(TomlLoadOptions.MaxStringLength))]
    [InlineData("a = 1\nb = 2\nc = 3\n", nameof(TomlLoadOptions.MaxInputBytes))]
    public void LimitsStopLoadingWithTheirName(string toml, string limit) {
        var options = new TomlLoadOptions { MaxDepth = 4, MaxNodes = 8, MaxStringLength = 8, MaxInputBytes = limit == nameof(TomlLoadOptions.MaxInputBytes) ? 16 : null };
        TomlLimitExceededException ex = Assert.Throws<TomlLimitExceededException>(() => TomlDocument.Load(new MemoryStream(Encoding.UTF8.GetBytes(toml)), options));
        Assert.Equal(limit, ex.LimitName);
        Assert.Equal(TomlErrorCodes.LimitExceeded, ex.Code);
    }

    [Fact]
    public void UntrustedDefaultsAreBoundedAndIndependent() {
        TomlLoadOptions options = TomlLoadOptions.UntrustedDefaults;
        Assert.Equal(16L * 1024 * 1024, options.MaxInputBytes);
        Assert.Equal(64, options.MaxDepth);
        options.MaxDepth = 2;
        Assert.Equal(64, TomlLoadOptions.UntrustedDefaults.MaxDepth);
    }

    [Fact]
    public async Task LoadingObservesCancellation() {
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => TomlDocument.LoadAsync(new MemoryStream(Encoding.UTF8.GetBytes("a = 1\n")), null, cancelled.Token));
        var reader = new TomlReader(string.Concat(Enumerable.Repeat("a = [1]\n", 3000)).Replace("a =", "k ="), null, cancelled.Token);
        Assert.Throws<OperationCanceledException>(() => { while (reader.Read()) { } });
    }

    [Fact]
    public void AByteOrderMarkIsKeptOnSave() {
        byte[] bytes = { 0xEF, 0xBB, 0xBF, (byte)'a', (byte)' ', (byte)'=', (byte)' ', (byte)'1', (byte)'\n' };
        using TomlDocument doc = TomlDocument.Load(new MemoryStream(bytes));
        doc.Root.Add("b", 2);
        var output = new MemoryStream();
        doc.Save(output);
        Assert.Equal(new byte[] { 0xEF, 0xBB, 0xBF }, output.ToArray().Take(3));
        Assert.Equal("a = 1\nb = 2\n", Encoding.UTF8.GetString(output.ToArray(), 3, (int)output.Length - 3));
    }

    [Fact]
    public void InvalidUtf8ReportsTheLineOfTheBadByte() {
        byte[] bytes = Encoding.UTF8.GetBytes("a = 1\nb = \"x\"\n").Concat(new byte[] { (byte)'c', (byte)' ', (byte)'=', (byte)' ', (byte)'"', 0xC3, 0x28, (byte)'"', (byte)'\n' }).ToArray();
        TomlException ex = Assert.Throws<TomlException>(() => TomlDocument.Load(new MemoryStream(bytes)));
        Assert.Equal(TomlErrorCodes.InvalidEncoding, ex.Code);
        Assert.Equal(3, ex.Line);
    }

    [Theory]
    [InlineData("a = 1\na = 2\n", TomlErrorCodes.DuplicateKey, 2, 1)]
    [InlineData("[t]\nx = 1\n[t]\n", TomlErrorCodes.InvalidTable, 3, 2)]
    [InlineData("a = 1\n[[a]]\n", TomlErrorCodes.InvalidArrayOfTables, 2, 3)]
    [InlineData("a = { b = 1 }\na.c = 2\n", TomlErrorCodes.InvalidTable, 2, 1)]
    [InlineData("a = 01\n", TomlErrorCodes.InvalidInteger, 1, 5)]
    [InlineData("s = \"\\q\"\n", TomlErrorCodes.InvalidEscape, 1, 6)]
    [InlineData("d = 2021-02-30\n", TomlErrorCodes.InvalidDateTime, 1, 10)]
    [InlineData("a = 1 b = 2\n", TomlErrorCodes.UnexpectedToken, 1, 7)]
    [InlineData("a = [\n  1,\n  2\n", TomlErrorCodes.UnexpectedToken, 4, 1)]
    public void ErrorsCarryAStableCodeAndThePosition(string toml, string code, int line, int column) {
        TomlException ex = Assert.Throws<TomlException>(() => TomlDocument.Parse(toml));
        Assert.Equal(code, ex.Code);
        Assert.Equal((line, column), (ex.Line, ex.Column));
    }

    [Fact]
    public void TypedReadsConvertAcrossNumericAndDateTypes() {
        using TomlDocument doc = TomlDocument.Parse("i = 42\nf = 2.0\nh = 2.5\nd = 1979-05-27\nt = 07:32:00.5\no = 1979-05-27T07:32:00-07:00\n");
        Assert.Equal(42.0, doc.Root["i"].GetValue<double>());
        Assert.Equal(2, doc.Root["f"].GetValue<int>());
        Assert.False(doc.Root["h"].TryGetValue(out int _));
        Assert.Equal(new DateTime(1979, 5, 27), doc.Root["d"].GetValue<DateTime>());
        Assert.Equal(new TimeSpan(0, 7, 32, 0, 500), doc.Root["t"].GetValue<TimeSpan>());
        Assert.Equal(new DateTimeOffset(1979, 5, 27, 14, 32, 0, TimeSpan.Zero), doc.Root["o"].GetValue<DateTimeOffset>());
        Assert.Equal("1979-05-27", doc.Root["d"].AsString());
        Assert.Throws<FormatException>(() => doc.Root["d"].GetValue<DateTimeOffset>());
        Assert.Throws<NotSupportedException>(() => doc.Root["i"].GetValue<Guid>());
    }
}
