using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed class BookOnixComplexityTests {
    private static readonly XNamespace Ns = BookProject.OnixNamespace;

    [Theory]
    [InlineData(BookOnixComplexityScheme.FryReadability, "03", "15")]
    [InlineData(BookOnixComplexityScheme.IoeBookBand, "04", "Pink A")]
    [InlineData(BookOnixComplexityScheme.FountasAndPinnell, "05", "Z+")]
    [InlineData(BookOnixComplexityScheme.Lexile, "06", "AD0L")]
    [InlineData(BookOnixComplexityScheme.Atos, "07", "17.0")]
    [InlineData(BookOnixComplexityScheme.FleschKincaid, "08", "-2.5")]
    [InlineData(BookOnixComplexityScheme.GuidedReading, "09", "J")]
    [InlineData(BookOnixComplexityScheme.ReadingRecovery, "10", "20")]
    [InlineData(BookOnixComplexityScheme.Lix, "11", "42")]
    [InlineData(BookOnixComplexityScheme.LexileAudio, "12", "600L")]
    [InlineData(BookOnixComplexityScheme.LexileSpanish, "13", "880L")]
    public void ComplexityOnlyAssertionsRetainSchemeAndValue(BookOnixComplexityScheme scheme, string code, string value) {
        var project = BookOnixTests.Project();
        byte[] before = project.ToProjectBytes();
        var baseline = project.ExportOnix(BookOnixTests.Options(), BookOnixTests.TestSchema());
        var result = project.ExportOnix(BookOnixTests.Options() with { Audience = new() {
            Complexities = [new(scheme, value)]
        } }, BookOnixTests.TestSchema());
        var xml = XDocument.Load(new MemoryStream(result.Bytes));
        var entry = Assert.Single(xml.Descendants(Ns + "Complexity"));
        Assert.Equal(code, entry.Element(Ns + "ComplexitySchemeIdentifier")!.Value);
        Assert.Equal(value, entry.Element(Ns + "ComplexityCode")!.Value);
        Assert.Empty(xml.Descendants(Ns + "Audience"));
        Assert.Equal(baseline.Publication.Bytes, result.Publication.Bytes);
        Assert.Equal(before, project.ToProjectBytes());
        Assert.Equal(result.Bytes, BookOnixMessage.Create([result], BookOnixTests.TestSchema()).Bytes);
    }

    [Fact]
    public void ComplexityFollowsAudienceDescriptionsAndAllowsDistinctValuesInOneScheme() {
        var result = BookOnixTests.Project().ExportOnix(BookOnixTests.Options() with { Audience = new() {
            Categories = [new(BookOnixAudienceType.Children)], Descriptions = [new("Independent assertions")],
            Complexities = [new(BookOnixComplexityScheme.Atos, "0"), new(BookOnixComplexityScheme.Atos, "17"),
                new(BookOnixComplexityScheme.FleschKincaid, "25.5"), new(BookOnixComplexityScheme.FryReadability, "1"),
                new(BookOnixComplexityScheme.ReadingRecovery, "1"), new(BookOnixComplexityScheme.FountasAndPinnell, "A")]
        } }, BookOnixTests.TestSchema());
        var children = XDocument.Load(new MemoryStream(result.Bytes)).Descendants(Ns + "DescriptiveDetail").Single().Elements().ToArray();
        int description = Array.FindIndex(children, e => e.Name == Ns + "AudienceDescription");
        Assert.True(description >= 0);
        Assert.All(children.Skip(description + 1), e => Assert.Equal(Ns + "Complexity", e.Name));
        Assert.Equal(new[] { "0", "17", "25.5", "1", "1", "A" }, children.Skip(description + 1).Select(e => e.Element(Ns + "ComplexityCode")!.Value));
    }

    [Theory]
    [InlineData(BookOnixComplexityScheme.FryReadability, "0")]
    [InlineData(BookOnixComplexityScheme.FryReadability, "16")]
    [InlineData(BookOnixComplexityScheme.FryReadability, "1.5")]
    [InlineData(BookOnixComplexityScheme.ReadingRecovery, "0")]
    [InlineData(BookOnixComplexityScheme.ReadingRecovery, "21")]
    [InlineData(BookOnixComplexityScheme.Atos, "-0.1")]
    [InlineData(BookOnixComplexityScheme.Atos, "17.1")]
    [InlineData(BookOnixComplexityScheme.Atos, "1,5")]
    [InlineData(BookOnixComplexityScheme.FleschKincaid, "NaN")]
    [InlineData(BookOnixComplexityScheme.FleschKincaid, "1e2")]
    [InlineData(BookOnixComplexityScheme.FountasAndPinnell, "AA")]
    [InlineData(BookOnixComplexityScheme.FountasAndPinnell, "a")]
    [InlineData(BookOnixComplexityScheme.Lexile, " 880L")]
    [InlineData(BookOnixComplexityScheme.IoeBookBand, "Pink\tA")]
    [InlineData(BookOnixComplexityScheme.Lix, "")]
    [InlineData(BookOnixComplexityScheme.Lix, null)]
    [InlineData(BookOnixComplexityScheme.Lix, "123456789012345678901")]
    [InlineData((BookOnixComplexityScheme)99, "1")]
    public void InvalidValuesAreRejectedWithoutMutation(BookOnixComplexityScheme scheme, string? value) {
        var project = BookOnixTests.Project();
        byte[] before = project.ToProjectBytes();
        Assert.ThrowsAny<ArgumentException>(() => project.ExportOnix(BookOnixTests.Options() with { Audience = new() {
            Complexities = [new(scheme, value!)]
        } }, BookOnixTests.TestSchema()));
        Assert.Equal(before, project.ToProjectBytes());
    }

    [Theory]
    [InlineData("null-list")]
    [InlineData("null-entry")]
    [InlineData("duplicate")]
    [InlineData("count")]
    public void InvalidListsAreRejected(string kind) {
        var entries = kind switch {
            "null-list" => null!, "null-entry" => new BookOnixComplexity[] { null! },
            "duplicate" => new[] { new BookOnixComplexity(BookOnixComplexityScheme.Lexile, "880L"), new(BookOnixComplexityScheme.Lexile, "880L") },
            _ => Enumerable.Range(0, 65).Select(n => new BookOnixComplexity(BookOnixComplexityScheme.FleschKincaid, n.ToString(System.Globalization.CultureInfo.InvariantCulture))).ToArray()
        };
        Assert.ThrowsAny<ArgumentException>(() => BookOnixTests.Project().ExportOnix(BookOnixTests.Options() with {
            Audience = new() { Complexities = entries }
        }, BookOnixTests.TestSchema()));
    }
}
