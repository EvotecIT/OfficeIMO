using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed class BookOnixResourceTests {
    private static readonly XNamespace Ns = BookProject.OnixNamespace;
    private static BookOnixResourceVersion Version() => new() { Form = BookOnixResourceForm.Downloadable, Links = [new("https://example.org/cover?a=1&b=2")] };
    private static BookOnixSupportingResource Resource() => new() {
        Type = BookOnixResourceContentType.FrontCover, Mode = BookOnixResourceMode.Image,
        Audiences = [BookOnixContentAudience.EndCustomers], Versions = [Version()]
    };
    private static BookOnixExportOptions Options(params BookOnixSupportingResource[] resources) => BookOnixTests.Options() with { SupportingResources = resources };

    [Fact]
    public void ResourceOnlyRecordKeepsVersionMetadataTranslationsAndPublication() {
        var project = BookOnixTests.Project(); byte[] before = project.ToProjectBytes();
        var baseline = project.ExportOnix(BookOnixTests.Options(), BookOnixTests.TestSchema());
        var result = project.ExportOnix(Options(Resource() with {
            Credits = [new("Publisher & artist", "eng"), new("Wydawca i artysta", "pol")], AlternativeTexts = [new("A blue cover")],
            Versions = [Version() with { FileFormatCode = "D502", ImageWidth = 1200, ImageHeight = 1800,
                FileName = "cover.jpg", ByteLength = 12345, Sha256 = new string('a', 64),
                UsageConstraints = [new(BookOnixUsageType.TextAndDataMining, BookOnixUsageStatus.Prohibited)],
                Licenses = [new() { Names = [new("Cover terms")] }],
                UsableFrom = new(2026, 1, 1), UsableUntil = new(2027, 1, 1), UpdatedOn = new(2026, 10, 6) },
                Version() with { Form = BookOnixResourceForm.Linkable, Links = [new("https://example.org/thumb.jpg")] }]
        }), BookOnixTests.TestSchema());
        var doc = XDocument.Load(new MemoryStream(result.Bytes)); var resource = doc.Descendants(Ns + "SupportingResource").Single();
        Assert.Empty(doc.Descendants(Ns + "TextContent"));
        Assert.Equal(new[] { "SequenceNumber", "ResourceContentType", "ContentAudience", "ResourceMode", "ResourceFeature", "ResourceFeature", "ResourceVersion", "ResourceVersion" }, resource.Elements().Select(e => e.Name.LocalName));
        Assert.Equal(new[] { "eng", "pol" }, resource.Elements(Ns + "ResourceFeature").First().Elements(Ns + "FeatureNote").Select(e => (string?)e.Attribute("language")));
        Assert.Equal("Publisher & artist", resource.Descendants(Ns + "FeatureNote").First().Value);
        var version = resource.Elements(Ns + "ResourceVersion").First();
        Assert.Equal(new[] { "D502", "1800", "1200", "cover.jpg", "12345", new string('a', 64) }, version.Elements(Ns + "ResourceVersionFeature").Select(e => e.Element(Ns + "FeatureValue")!.Value));
        Assert.Equal("https://example.org/cover?a=1&b=2", version.Element(Ns + "ResourceLink")!.Value);
        Assert.Equal(new[] { "14", "15", "17" }, version.Descendants(Ns + "ContentDateRole").Select(e => e.Value));
        Assert.Equal(new[] { "EpubUsageConstraint", "EpubLicense", "ContentDate", "ContentDate", "ContentDate" }, version.Element(Ns + "ResourceLink")!.ElementsAfterSelf().Select(e => e.Name.LocalName));
        Assert.Equal(before, project.ToProjectBytes()); Assert.Equal(baseline.Publication.Bytes, result.Publication.Bytes);
        Assert.Equal(result.Bytes, BookOnixMessage.Create([result], BookOnixTests.TestSchema()).Bytes);
    }

    [Fact]
    public void AlternateUrlsMayShareALanguageAndParallelLanguagesRemainExplicit() {
        var result = BookOnixTests.Project().ExportOnix(Options(Resource() with { Versions = [Version() with {
            Links = [new("https://example.org/one", "eng"), new("https://example.org/two", "eng"), new("https://example.org/three", "pol")]
        }, Version() with { Links = [new("https://example.org/four"), new("https://example.org/five")] }] }), BookOnixTests.TestSchema());
        Assert.Equal(new[] { "eng", "eng", "pol", null, null }, XDocument.Load(new MemoryStream(result.Bytes)).Descendants(Ns + "ResourceLink").Select(e => (string?)e.Attribute("language")));
    }

    [Fact]
    public void ResourcesFollowTextAndResourceSequenceIsIndependent() {
        var options = Options(Resource(), Resource() with { Type = BookOnixResourceContentType.BackCover }) with {
            CollateralTexts = [new() { Type = BookOnixTextType.Description, Audiences = [BookOnixContentAudience.Unrestricted], Texts = [new("Description")] }]
        };
        var result = BookOnixTests.Project().ExportOnix(options, BookOnixTests.TestSchema());
        var collateral = XDocument.Load(new MemoryStream(result.Bytes)).Descendants(Ns + "CollateralDetail").Single();
        Assert.Equal(new[] { "TextContent", "SupportingResource", "SupportingResource" }, collateral.Elements().Select(e => e.Name.LocalName));
        Assert.Equal(new[] { "1", "2" }, collateral.Elements(Ns + "SupportingResource").Select(e => e.Element(Ns + "SequenceNumber")!.Value));
    }

    [Theory]
    [InlineData("null-resources")][InlineData("null-resource")][InlineData("too-many-resources")]
    [InlineData("unknown-type")][InlineData("legacy-type")][InlineData("unknown-mode")][InlineData("no-audience")][InlineData("mixed-audience")]
    [InlineData("no-versions")][InlineData("null-versions")][InlineData("null-version")][InlineData("too-many-versions")]
    [InlineData("no-links")][InlineData("null-links")][InlineData("null-link")][InlineData("too-many-links")]
    [InlineData("link-scheme")][InlineData("link-credentials")][InlineData("duplicate-link")][InlineData("mixed-language")][InlineData("bad-language")]
    [InlineData("unknown-form")][InlineData("unknown-format")][InlineData("width-zero")][InlineData("height-negative")][InlineData("nonimage-dimensions")]
    [InlineData("size-negative")][InlineData("hash-length")][InlineData("hash-character")][InlineData("filename-path")][InlineData("filename-long")]
    [InlineData("dates")][InlineData("negative-minutes")][InlineData("image-minutes")][InlineData("credit-language")][InlineData("credit-xhtml")][InlineData("null-credit")]
    public void InvalidResourceAssertionsFailWithoutMutation(string kind) {
        var version = kind switch {
            "no-links" => Version() with { Links = [] }, "null-links" => Version() with { Links = null! },
            "null-link" => Version() with { Links = [null!] }, "too-many-links" => Version() with { Links = Enumerable.Repeat(new BookOnixResourceLink("https://example.org"), 17).ToArray() },
            "link-scheme" => Version() with { Links = [new("file:///tmp/a")] }, "link-credentials" => Version() with { Links = [new("https://u:p@example.org")] },
            "duplicate-link" => Version() with { Links = [new("https://example.org"), new("https://example.org")] },
            "mixed-language" => Version() with { Links = [new("https://example.org/one", "eng"), new("https://example.org/two")] },
            "bad-language" => Version() with { Links = [new("https://example.org", "en")] },
            "unknown-form" => Version() with { Form = (BookOnixResourceForm)99 }, "unknown-format" => Version() with { FileFormatCode = "image/jpeg" },
            "width-zero" => Version() with { ImageWidth = 0 }, "height-negative" => Version() with { ImageHeight = -1 },
            "nonimage-dimensions" => Version() with { ImageWidth = 100 }, "size-negative" => Version() with { ByteLength = -1 },
            "hash-length" => Version() with { Sha256 = "abc" }, "hash-character" => Version() with { Sha256 = new string('z', 64) },
            "filename-path" => Version() with { FileName = "../cover.jpg" }, "filename-long" => Version() with { FileName = new string('a', 256) },
            "dates" => Version() with { UsableFrom = new(2027, 1, 1), UsableUntil = new(2026, 1, 1) }, _ => Version()
        };
        var resource = Resource() with { Versions = [version] };
        resource = kind switch {
            "unknown-type" => resource with { Type = (BookOnixResourceContentType)99 }, "legacy-type" => resource with { Type = (BookOnixResourceContentType)27 },
            "unknown-mode" => resource with { Mode = (BookOnixResourceMode)99 }, "no-audience" => resource with { Audiences = [] },
            "mixed-audience" => resource with { Audiences = [BookOnixContentAudience.Unrestricted, BookOnixContentAudience.Press] },
            "no-versions" => resource with { Versions = [] }, "null-versions" => resource with { Versions = null! },
            "null-version" => resource with { Versions = [null!] }, "too-many-versions" => resource with { Versions = Enumerable.Repeat(version, 17).ToArray() },
            "nonimage-dimensions" => resource with { Mode = BookOnixResourceMode.Audio }, "negative-minutes" => resource with { Mode = BookOnixResourceMode.Audio, LengthMinutes = -1 },
            "image-minutes" => resource with { LengthMinutes = 1 }, "credit-language" => resource with { Credits = [new("One", "eng"), new("Two")] },
            "credit-xhtml" => resource with { Credits = [new("<p>Credit</p>") { Format = BookOnixCollateralTextFormat.Xhtml }] },
            "null-credit" => resource with { Credits = null! }, _ => resource
        };
        BookOnixSupportingResource[] resources = kind switch {
            "null-resources" => null!, "null-resource" => [null!], "too-many-resources" => Enumerable.Repeat(resource, 65).ToArray(), _ => [resource]
        };
        var project = BookOnixTests.Project(); byte[] before = project.ToProjectBytes();
        Assert.ThrowsAny<ArgumentException>(() => project.ExportOnix(Options(resources), BookOnixTests.TestSchema()));
        Assert.Equal(before, project.ToProjectBytes());
    }

    [Fact]
    public void ResourceTextAndLinksShareTheExistingCollateralBudget() {
        var text = new BookOnixCollateralText { Type = BookOnixTextType.Description, Audiences = [BookOnixContentAudience.Unrestricted], Texts = [new(new string('a', 65536))] };
        Assert.Throws<ArgumentException>(() => BookOnixTests.Project().ExportOnix(Options(Resource()) with {
            CollateralTexts = Enumerable.Repeat(text, 8).ToArray()
        }, BookOnixTests.TestSchema()));
    }
}
