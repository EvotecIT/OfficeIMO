using System.Security.Cryptography;
using Xunit;

namespace OfficeIMO.Odbc.Tests;

public sealed class SnapshotTests {
    [Fact]
    public void OversizedWorkbookIsRejectedBeforeParsing() {
        using var fixture = new WorkbookFixture();
        using (var file = File.OpenWrite(fixture.FilePath)) file.SetLength(64L * 1024 * 1024 + 1);
        Assert.Throws<InvalidDataException>(() => new WorkbookSnapshot(fixture.FilePath, CancellationToken.None));
    }

    [Fact]
    public void QueriesPreserveUnicodeNumbersMixedTextAndNulls() {
        using var fixture = new WorkbookFixture();
        byte[] original = SHA256.HashData(File.ReadAllBytes(fixture.FilePath));
        using var snapshot = new WorkbookSnapshot(fixture.FilePath, CancellationToken.None);
        Assert.Single(snapshot.Tables);
        Assert.Equal(typeof(double), snapshot.Tables[0].Columns[2].Type);
        Assert.Equal(typeof(string), snapshot.Tables[0].Columns[3].Type);
        using var reader = snapshot.Query("SELECT Name, Amount, Mixed, Optional FROM Customers ORDER BY Amount", new Dictionary<string, object?>(), CancellationToken.None);
        Assert.True(reader.Read());
        Assert.Equal("Alicja Żółć", reader.GetString(0));
        Assert.Equal(12.5, reader.GetDouble(1));
        Assert.Equal("7", reader.GetString(2));
        Assert.Equal("", reader.GetString(3));
        Assert.True(reader.Read());
        Assert.Equal("001", reader.GetString(2));
        Assert.True(reader.IsDBNull(3));
        Assert.False(reader.Read());
        Assert.Equal(original, SHA256.HashData(File.ReadAllBytes(fixture.FilePath)));
    }

    [Fact]
    public void IndependentConnectionsKeepTheirOwnImmutableSnapshots() {
        using var fixture = new WorkbookFixture();
        using var first = new WorkbookSnapshot(fixture.FilePath, CancellationToken.None);
        fixture.Write("Changed");
        using var second = new WorkbookSnapshot(fixture.FilePath, CancellationToken.None);
        using var original = first.Query("SELECT Name FROM Customers WHERE Country = @p1", new Dictionary<string, object?> { ["@p1"] = "Poland" }, CancellationToken.None);
        using var changed = second.Query("SELECT Name FROM Customers WHERE Country = @p1", new Dictionary<string, object?> { ["@p1"] = "Poland" }, CancellationToken.None);
        Assert.True(original.Read());
        Assert.True(changed.Read());
        Assert.Equal("Alicja Żółć", original.GetString(0));
        Assert.Equal("Changed", changed.GetString(0));
    }

    [Theory]
    [InlineData("DELETE FROM Customers")]
    [InlineData("PRAGMA query_only=OFF")]
    [InlineData("ATTACH DATABASE 'file.db' AS other")]
    [InlineData("SELECT * FROM Customers; DELETE FROM Customers")]
    [InlineData("WITH x AS (SELECT 1) SELECT * FROM x")]
    public void ApplicationCommandsCannotMutateOrAttach(string sql) {
        using var fixture = new WorkbookFixture();
        using var snapshot = new WorkbookSnapshot(fixture.FilePath, CancellationToken.None);
        Assert.Throws<NotSupportedException>(() => snapshot.Query(sql, new Dictionary<string, object?>(), CancellationToken.None));
        using var count = snapshot.Query("SELECT COUNT(*) FROM Customers", new Dictionary<string, object?>(), CancellationToken.None);
        Assert.True(count.Read());
        Assert.Equal(2, count.GetInt64(0));
    }

    [Fact]
    public void CancellationInterruptsNativeQueryExecution() {
        using var fixture = new WorkbookFixture();
        using var snapshot = new WorkbookSnapshot(fixture.FilePath, CancellationToken.None);
        using var cancellation = new CancellationTokenSource(TimeSpan.FromMilliseconds(100));
        // Cross-joining 2 rows 30 times ensures the native engine is still working when cancellation arrives.
        string from = string.Join(",", Enumerable.Range(0, 30).Select(i => "Customers c" + i));
        Assert.ThrowsAny<OperationCanceledException>(() => snapshot.Query("SELECT COUNT(*) FROM " + from, new Dictionary<string, object?>(), cancellation.Token));
    }
}
