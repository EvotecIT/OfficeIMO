using OfficeIMO.Epub;
using System.Text;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubAccessibilityMetadataContracts {
    private static readonly XNamespace Opf = "http://www.idpf.org/2007/opf";

    [Fact]
    public void DiscoveryMetadataRoundTripsModeCombinationsAndReplacesOldValues() {
        var book = Book();
        book.DeclareVocabularyPrefix("discovery", "http://schema.org/");
        book.AddMetadataProperty("discovery:accessMode", "visual");
        book.AddMetadataProperty("schema:accessibilityFeature", "alternativeText");
        book.AddMetadataProperty("schema:accessibilityFeature", "index");
        book.AddMetadataProperty("schema:copyrightYear", "2026");
        var claims = Claims();
        claims.AccessModes = new[] { "textual", "visual", "textual" };
        claims.SufficientAccessModes = new IReadOnlyList<string>[] { new[] { "textual", "visual" }, new[] { "auditory" } };
        claims.Summary = "Reviewed text and navigation.";
        book.SetAccessibilityMetadata(claims);
        var reopened = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        XElement[] metadata = reopened.GetPackageXml().Descendants(Opf + "meta").ToArray();
        Assert.Equal(new[] { "textual", "visual" }, metadata.Where(e => ((string?)e.Attribute("property"))?.EndsWith(":accessMode") == true).Select(e => e.Value));
        Assert.Equal(new[] { "textual,visual", "auditory" }, Values(metadata, "accessModeSufficient"));
        Assert.Equal(new[] { "structuralNavigation" }, Values(metadata, "accessibilityFeature"));
        Assert.Equal(new[] { "2026" }, Values(metadata, "copyrightYear"));
        Assert.Equal(EpubPreflightStatus.Passed, Discovery(reopened).Status);
        Assert.Empty(Discovery(reopened).Diagnostics);
        reopened.SetAccessibilityMetadata(Claims());
        Assert.Empty(Values(reopened.GetPackageXml().Descendants(Opf + "meta"), "accessModeSufficient"));
        Assert.Empty(Values(reopened.GetPackageXml().Descendants(Opf + "meta"), "accessibilitySummary"));
    }

    [Fact]
    public void InvalidClaimsAndReferencedDeclarationRemovalAreAtomic() {
        var book = Book();
        book.SetAccessibilityMetadata(Claims());
        byte[] before = book.Write().Bytes;
        var invalid = Claims();
        invalid.Hazards = new[] { "bad value" };
        Assert.Throws<ArgumentException>(() => book.SetAccessibilityMetadata(invalid));
        Assert.Equal(before, book.Write().Bytes);
        invalid = Claims();
        invalid.SufficientAccessModes = new IReadOnlyList<string>[] { Array.Empty<string>() };
        Assert.Throws<ArgumentException>(() => book.SetAccessibilityMetadata(invalid));
        Assert.Equal(before, book.Write().Bytes);

        XDocument package = book.GetPackageXml();
        package.Root!.Element(Opf + "metadata")!.Add(
            new XElement(Opf + "meta", new XAttribute("property", "schema:accessibilitySummary"), new XAttribute("id", "summary"), "Retain me"),
            new XElement(Opf + "meta", new XAttribute("property", "alternate-script"), new XAttribute("refines", "#summary"), new XAttribute(XNamespace.Xml + "lang", "fr"), "Résumé"));
        byte[] input = EpubIntegrityFixtures.ReplaceEntry(before, book.PackagePath, Encoding.UTF8.GetBytes(package.ToString()));
        var imported = EpubPublication.Load(new MemoryStream(input));
        Assert.Throws<InvalidOperationException>(() => imported.SetAccessibilityMetadata(Claims()));
        Assert.Equal(input, imported.Write().Bytes);
        var replacement = Claims();
        replacement.Summary = "Replacement summary";
        imported.SetAccessibilityMetadata(replacement);
        XElement[] output = EpubPublication.Load(new MemoryStream(imported.Write().Bytes)).GetPackageXml().Descendants(Opf + "meta").ToArray();
        Assert.Contains(output, e => (string?)e.Attribute("id") == "summary" && e.Value == replacement.Summary);
        Assert.Contains(output, e => (string?)e.Attribute("refines") == "#summary" && e.Value == "Résumé");
    }

    [Fact]
    public void DiscoveryDistinguishesRequiredClaimsRecommendationsAndUnsupportedVersion() {
        var book = Book();
        Assert.Equal(3, Discovery(book).Diagnostics.Count(d => d.Severity == EpubDiagnosticSeverity.Error));
        Assert.Equal(2, Discovery(book).Diagnostics.Count(d => d.Severity == EpubDiagnosticSeverity.Warning));
        book.SetAccessibilityMetadata(Claims());
        Assert.Equal(EpubPreflightStatus.Passed, Discovery(book).Status);
        Assert.Equal(2, Discovery(book).Diagnostics.Count);
        var legacy = EpubPublication.Create("Legacy", "en", version: EpubVersion.Epub2);
        Assert.Equal(EpubPreflightStatus.NotChecked, Discovery(legacy).Status);
        Assert.Throws<NotSupportedException>(() => legacy.SetAccessibilityMetadata(Claims()));
    }

    private static IEnumerable<string> Values(IEnumerable<XElement> elements, string property) =>
        elements.Where(e => (string?)e.Attribute("property") == "schema:" + property).Select(e => e.Value);
    private static EpubPreflightCheck Discovery(EpubPublication book) => book.Preflight().Checks.Single(c => c.Code == "accessibility-discovery");
    private static EpubAccessibilityMetadata Claims() => new EpubAccessibilityMetadata {
        AccessModes = new[] { "textual" }, Features = new[] { "structuralNavigation" }, Hazards = new[] { "unknown" }
    };
    private static EpubPublication Book() {
        var book = EpubPublication.Create("Metadata", "en");
        book.AddChapter("c", "EPUB/c.xhtml", "Chapter", "<h1>Chapter</h1><p>Text.</p>");
        return book;
    }
}
