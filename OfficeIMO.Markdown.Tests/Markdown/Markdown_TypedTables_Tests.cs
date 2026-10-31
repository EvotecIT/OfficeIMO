using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public sealed class Markdown_TypedTables_Tests {
    [Fact]
    public void TypedTableReturnsItsBlockAndPreservesProjectionControls() {
        using var document = MarkdownDocument.Create();
        var options = new MarkdownTableOptions();
        options.Include.UnionWith(new[] { "Name", "Amount", "Date" });
        options.Exclude.Add("Date");
        options.Order.AddRange(new[] { "Amount", "Name" });
        options.HeaderTransform = name => "Display " + name;
        options.HeaderRenames["Amount"] = "Total";
        options.Formatters["Amount"] = value => string.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0.00}", value);
        options.Alignments.Add(OfficeAlignment.Center);
        var table = document.AddTable(new[] { new Row { Name = "First", Amount = 12.5m, Date = "2025-01-01" } }, options);

        Assert.Same(table, Assert.Single(document.Blocks));
        Assert.Equal(new[] { "Total", "Display Name" }, table.Headers);
        Assert.Equal(new[] { "12.50", "First" }, Assert.Single(table.Rows));
        Assert.Equal(MarkdownColumnAlignment.Center, table.Alignments[0]);
        table.Rows.Add(new[] { "8.00", "Second" });
        Assert.Contains("Second", document.ToMarkdown());
    }

    [Fact]
    public void TypedDictionaryRowsUseTheirKeysAndEnumerateTheSourceOnce() {
        int enumerations = 0;
        IEnumerable<IReadOnlyDictionary<string, object?>> Rows() {
            if (++enumerations != 1) throw new InvalidOperationException("The source was enumerated twice.");
            yield return new System.Collections.ObjectModel.ReadOnlyDictionary<string, object?>(new Dictionary<string, object?> {
                ["Name"] = "First", ["Amount"] = 2
            });
            yield return new System.Collections.ObjectModel.ReadOnlyDictionary<string, object?>(new Dictionary<string, object?> {
                ["Name"] = "Second"
            });
        }
        using var document = MarkdownDocument.Create();
        var table = document.AddTable(Rows());
        Assert.Equal(1, enumerations);
        Assert.Equal(new[] { "Name", "Amount" }, table.Headers);
        Assert.Equal(new[] { "Second", "" }, table.Rows[1]);
    }

    [Fact]
    public void FailedTypedProjectionDoesNotAddAPartialTable() {
        using var document = MarkdownDocument.Create();
        var options = new OfficeTableOptions();
        options.Formatters["Name"] = _ => throw new InvalidOperationException("Invalid source value.");
        Assert.Throws<InvalidOperationException>(() => document.AddTable(new[] { new Row() }, options));
        Assert.Empty(document.Blocks);
        Assert.Throws<InvalidOperationException>(() => document.AddTable(Enumerable.Repeat(new Row(), 10001)));
        Assert.Empty(document.Blocks);
        Assert.Throws<ArgumentException>(() => document.AddTable(Array.Empty<IReadOnlyDictionary<string, object?>>()));
        Assert.Empty(document.Blocks);
        options.Formatters.Clear();
        options.Alignments.Add(OfficeAlignment.Justify);
        Assert.Throws<ArgumentOutOfRangeException>(() => document.AddTable(new[] { new Row() }, options));
        Assert.Empty(document.Blocks);
    }

    [Fact]
    public void HeaderlessTablesAndRuntimeDataRetainTheirValues() {
        using var document = MarkdownDocument.Create();
        var headerless = document.AddTable(new[] { new Row { Name = "First", Amount = 2 } },
            new OfficeTableOptions { IncludeHeaders = false });
        Assert.Empty(headerless.Headers);
        Assert.Equal("First", Assert.Single(headerless.Rows)[0]);
        var metadata = document.AddTable((object)new { Project = "OfficeIMO", Version = "3.6" });
        Assert.Equal(new[] { "Property", "Value" }, metadata.Headers);
        Assert.Equal(new[] { "Project", "OfficeIMO" }, metadata.Rows[0]);
    }

    [Fact]
    public void DisplayHeadersDoNotChangeTheSelectedMemberSchema() {
        using var document = MarkdownDocument.Create();
        var options = new OfficeTableOptions { HeaderTransform = _ => "Value" };
        var table = document.AddTable(new[] { new Row { Name = "First", Amount = 2 } }, options);
        Assert.Equal(new[] { "Value", "Value", "Value" }, table.Headers);
        Assert.Equal("First", Assert.Single(table.Rows)[0]);
    }

    private sealed class Row {
        public string Name { get; set; } = "";
        public decimal Amount;
        public string Date { get; set; } = "";
    }
}
