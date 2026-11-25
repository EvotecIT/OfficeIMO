using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>
/// Loading a plain Open XML presentation and saving it unchanged does not open the Open XML SDK package, and writes
/// what a save through the SDK writes: the same parts and relationships, the same document properties, and every
/// other part as it was loaded.
/// </summary>
public class PowerPointUnchangedSaveTests {
    private const string ApplicationPart = "docProps/app.xml";
    private const string CorePart = "docProps/core.xml";
    private const string PresentationPart = "ppt/presentation.xml";
    private static readonly XNamespace DublinCoreTerms = "http://purl.org/dc/terms/";

    public static IEnumerable<object[]> CorpusPresentations() => OpenXmlSlideTextReaderTests.CorpusPresentations();

    [Theory]
    [MemberData(nameof(CorpusPresentations))]
    public void SavesAnUnchangedCorpusPresentationAsTheSdkPathDoes(string relativePath) {
        byte[] source = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, relativePath));
        byte[] throughSdk = SaveThroughSdk(source);

        byte[] withoutSdk;
        bool sdkLoaded;
        using (var presentation = PowerPointDocument.Load(new MemoryStream(source))) {
            withoutSdk = presentation.ToBytes();
            sdkLoaded = presentation.IsSdkLoaded;
        }

        AssertSamePackage(source, throughSdk, withoutSdk, keptAsLoaded: !sdkLoaded);
    }

    // Opt-in: OFFICEIMO_PPTX_CORPUS names a folder of presentations, such as the public corpora the real-world corpus
    // runner fetches. A deck the deferred load accepts must open in the SDK, and save as the SDK path saves it.
    [Fact]
    public void ExternalCorpusPresentationsLoadAndSaveAsTheSdkPathDoes() {
        string? folder = Environment.GetEnvironmentVariable("OFFICEIMO_PPTX_CORPUS");
        if (string.IsNullOrEmpty(folder) || !Directory.Exists(folder)) return;
        var failures = new List<string>();
        foreach (string path in Directory.EnumerateFiles(folder, "*.pptx", SearchOption.AllDirectories).OrderBy(file => file, StringComparer.Ordinal)) {
            Exception? failure = Record.Exception(() => CheckCorpusPresentation(path));
            if (failure != null) failures.Add(path + ": " + failure.Message);
        }
        Assert.True(failures.Count == 0, string.Join(Environment.NewLine, failures));
    }

    private static void CheckCorpusPresentation(string path) {
        byte[] source = File.ReadAllBytes(path);
        PowerPointDocument presentation;
        try {
            presentation = PowerPointDocument.Load(new MemoryStream(source));
        } catch (Exception) {
            // The load failed before any deferral could hide it.
            return;
        }
        using (presentation) {
            if (presentation.IsSdkLoaded) return;
            byte[]? throughSdk = null;
            Exception? sdkFailure = null;
            try {
                throughSdk = SaveThroughSdk(source);
            } catch (Exception exception) {
                sdkFailure = exception;
            }
            using (var opened = PowerPointDocument.Load(new MemoryStream(source))) {
                Exception? openFailure = Record.Exception(() => opened.OpenXmlDocument);
                Assert.True(openFailure == null, "The deferred load accepted a package the SDK rejects: " + openFailure);
            }
            Exception? saveFailure = Record.Exception(() => presentation.ToBytes());
            if (sdkFailure != null) {
                Assert.True(saveFailure != null, "The SDK save fails, the save without it did not: " + sdkFailure.Message);
                return;
            }
            Assert.True(saveFailure == null, "The save without the SDK failed: " + saveFailure);
            byte[] saved = presentation.ToBytes();
            AssertSamePackage(source, throughSdk!, saved, keptAsLoaded: !presentation.IsSdkLoaded);
        }
    }

    [Fact]
    public void SavesWithoutOpeningTheSdkAndRefreshesTheDocumentProperties() {
        byte[] source = CreateDeckWithStaleProperties();
        DateTime before = DateTime.UtcNow.AddSeconds(-1);

        byte[] saved;
        using (var presentation = PowerPointDocument.Load(new MemoryStream(source))) {
            saved = presentation.ToBytes();
            Assert.False(presentation.IsSdkLoaded);
        }

        XDocument application = ReadXml(saved, ApplicationPart);
        XNamespace ap = application.Root!.Name.Namespace;
        Assert.Equal("2", application.Root.Element(ap + "Slides")!.Value);
        Assert.Equal("1", application.Root.Element(ap + "HiddenSlides")!.Value);
        Assert.Equal("1", application.Root.Element(ap + "Notes")!.Value);
        Assert.Equal("Standard", application.Root.Element(ap + "PresentationFormat")!.Value);
        DateTime modified = DateTime.Parse(ReadXml(saved, CorePart).Root!.Element(DublinCoreTerms + "modified")!.Value,
            System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.AdjustToUniversal);
        Assert.True(modified >= before, modified.ToString("O"));
        AssertSamePackage(source, SaveThroughSdk(source), saved);
        Assert.Equal(0, CountValidationErrors(saved));
    }

    [Fact]
    public void OpensTheSdkOnFirstUseAndSavesThroughIt() {
        byte[] source = CreateDeckWithStaleProperties();
        using var presentation = PowerPointDocument.Load(new MemoryStream(source));
        Assert.False(presentation.IsSdkLoaded);

        Assert.Equal(2, presentation.Slides.Count);
        Assert.False(presentation.IsSdkLoaded);
        PowerPointSlide first = presentation.Slides[0];
        Assert.False(presentation.IsSdkLoaded);
        _ = presentation.OpenXmlDocument;
        Assert.True(presentation.IsSdkLoaded);
        Assert.Equal(1, presentation.DeferredSdkOpenCount);
        Assert.Same(first, presentation.Slides[0]);
        OfficeDiagnostic fallback = Assert.Single(presentation.Diagnostics);
        Assert.Equal(OfficeDiagnosticCodes.SdkFallback, fallback.Code);
        Assert.Equal("PowerPointDocument.OpenXmlDocument", fallback.Location);
        first.AddTextBox("Added");

        using var reopened = PowerPointDocument.Load(new MemoryStream(presentation.ToBytes()));
        Assert.Contains(reopened.Slides[0].TextBoxes, box => box.Text == "Added");
    }

    [Fact]
    public void ReadOnlyLoadsCopyTheLoadedPackageWithoutTheSdk() {
        byte[] source = CreateDeckWithStaleProperties();
        using var presentation = PowerPointDocument.Load(new MemoryStream(source),
            new PowerPointLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });

        byte[] saved = presentation.ToBytes();

        Assert.False(presentation.IsSdkLoaded);
        Assert.Throws<InvalidOperationException>(() => presentation.Save(new MemoryStream()));
        AssertSamePackage(source, SaveThroughSdk(source), saved);
    }

    [Fact]
    public void SaveOnDisposeWritesTheUnchangedPresentationWithoutTheSdk() {
        byte[] source = CreateDeckWithStaleProperties();
        var destination = new MemoryStream();
        destination.Write(source, 0, source.Length);
        destination.Position = 0;

        var presentation = PowerPointDocument.Load(destination, new PowerPointLoadOptions { PersistenceMode = DocumentPersistenceMode.SaveOnDispose });
        presentation.Dispose();

        Assert.False(presentation.IsSdkLoaded);
        AssertSamePackage(source, SaveThroughSdk(source), destination.ToArray());
    }

    [Theory]
    [InlineData("footer")]
    [InlineData("legacy-show")]
    [InlineData("missing-app")]
    [InlineData("signed")]
    public void UsesTheSdkWhenTheSaveWouldChangeMoreThanTheProperties(string variant) {
        byte[] source = CreateDeckWithStaleProperties();
        source = variant switch {
            "footer" => Rewrite(source, PresentationPart, xml => xml.Replace("</p:presentation>",
                "<p:extLst><p:ext uri=\"urn:officeimo:powerpoint:footer-definitions:1\"/></p:extLst></p:presentation>")),
            "legacy-show" => Rewrite(source, PresentationPart, xml => xml.Replace("<p:sldId ", "<p:sldId show=\"0\" ")),
            "missing-app" => Rewrite(source, ApplicationPart, xml => Regex.Replace(xml, "<(\\w+:)?TotalTime>.*?</(\\w+:)?TotalTime>|<(\\w+:)?TotalTime/>", string.Empty)),
            _ => Rewrite(source, ApplicationPart, xml => Regex.Replace(xml, "</(\\w+:)?Properties>\\s*$", match => "<ap:DigSig xmlns:ap=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\"><vt:blob xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">AA==</vt:blob></ap:DigSig>" + match.Value))
        };

        using var presentation = PowerPointDocument.Load(new MemoryStream(source));
        // A signed package opens the SDK at load, where its fingerprint is taken, and a save is blocked as before.
        if (variant == "signed") {
            Assert.True(presentation.IsSdkLoaded);
            Assert.Throws<PowerPointSignedDocumentMutationException>(() => presentation.ToBytes());
            Assert.Throws<PowerPointSignedDocumentMutationException>(() => SaveThroughSdk(source));
            return;
        }
        byte[] saved = presentation.ToBytes();

        Assert.True(presentation.IsSdkLoaded);
        AssertSamePackage(source, SaveThroughSdk(source), saved, keptAsLoaded: false);
    }

    // The SDK reads markup compatibility as written by default, so a wrapped element is not the one it updates or counts.
    [Fact]
    public void ReadsTheSaveValuesWithMarkupCompatibilityAsWritten() {
        const string compatibility = "http://schemas.openxmlformats.org/markup-compatibility/2006";
        byte[] source = Rewrite(CreateDeckWithStaleProperties(), PresentationPart, xml => Regex.Replace(xml, "<p:sldSz [^>]*/>",
            match => "<mc:AlternateContent xmlns:mc=\"" + compatibility + "\"><mc:Choice Requires=\"p\">" + match.Value + "</mc:Choice></mc:AlternateContent>"));
        Assert.Contains("mc:AlternateContent", System.Text.Encoding.UTF8.GetString(ReadParts(source)[PresentationPart]));

        byte[] saved;
        using (var presentation = PowerPointDocument.Load(new MemoryStream(source))) {
            saved = presentation.ToBytes();
            Assert.False(presentation.IsSdkLoaded);
        }

        AssertSamePackage(source, SaveThroughSdk(source), saved);
    }

    [Fact]
    public void UsesTheSdkWhenACounterIsWrappedInMarkupCompatibility() {
        const string compatibility = "http://schemas.openxmlformats.org/markup-compatibility/2006";
        byte[] source = Rewrite(CreateDeckWithStaleProperties(), ApplicationPart, xml => Regex.Replace(xml, "<(\\w+:)?Slides>[^<]*</(\\w+:)?Slides>",
            match => "<mc:AlternateContent xmlns:mc=\"" + compatibility + "\"><mc:Fallback>" + match.Value + "</mc:Fallback></mc:AlternateContent>"));
        Assert.Contains("mc:Fallback", System.Text.Encoding.UTF8.GetString(ReadParts(source)[ApplicationPart]));

        using var presentation = PowerPointDocument.Load(new MemoryStream(source));
        byte[] saved = presentation.ToBytes();

        Assert.True(presentation.IsSdkLoaded);
        AssertSamePackage(source, SaveThroughSdk(source), saved, keptAsLoaded: false);
    }

    // Only the entries of the root's slide list are slides; an entry in an extension is not one to the SDK either.
    [Fact]
    public void CountsOnlyTheEntriesOfTheSlideList() {
        byte[] source = Rewrite(CreateDeckWithStaleProperties(), PresentationPart, xml => xml.Replace("</p:presentation>",
            "<x:metadata xmlns:x=\"urn:example\"><p:sldId xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" id=\"999\" r:id=\"rIdOther\"/></x:metadata></p:presentation>"));

        byte[] saved;
        using (var presentation = PowerPointDocument.Load(new MemoryStream(source))) {
            saved = presentation.ToBytes();
            Assert.False(presentation.IsSdkLoaded);
        }

        Assert.Equal("2", ReadXml(saved, ApplicationPart).Root!.Elements().Single(element => element.Name.LocalName == "Slides").Value);
        AssertSamePackage(source, SaveThroughSdk(source), saved);
    }

    [Fact]
    public void TheReturnedBytesBelongToTheCaller() {
        using var presentation = PowerPointDocument.Load(new MemoryStream(CreateDeckWithStaleProperties()));
        byte[] saved = presentation.ToBytes();
        Assert.False(presentation.IsSdkLoaded);

        Array.Clear(saved, 0, saved.Length);

        Assert.Equal(2, presentation.Slides.Count);
    }

    [Fact]
    public void AnSdkOpenedAfterTheSaveReadsWhatWasSaved() {
        using var presentation = PowerPointDocument.Load(new MemoryStream(CreateDeckWithStaleProperties()));
        presentation.ToBytes();
        Assert.False(presentation.IsSdkLoaded);

        Assert.Equal("2", presentation.OpenXmlDocument.ExtendedFilePropertiesPart!.Properties!.Slides!.Text);
        Assert.Equal("Standard", presentation.ApplicationProperties.PresentationFormat);
        Assert.True(presentation.OpenXmlDocument.PackageProperties.Modified > new DateTime(2001, 1, 2, 0, 0, 0, DateTimeKind.Utc));
    }

    [Fact]
    public void TheSdkOpensWithTheSettingsAsTheyWereAtLoad() {
        byte[] source = CreateDeckWithStaleProperties();
        var settings = new OfficeOpenXmlLoadSettings { MaxCharactersInPart = 10_000_000 };
        using var presentation = PowerPointDocument.Load(new MemoryStream(source), new PowerPointLoadOptions { OpenSettings = settings });
        Assert.False(presentation.IsSdkLoaded);

        settings.MaxCharactersInPart = 1;

        Assert.Equal(2, presentation.Slides.Count);
    }

    [Fact]
    public void LoadsThroughTheSdkWhenAListedSlideIsNotASlidePart() {
        byte[] source = Rewrite(CreateDeckWithStaleProperties(), "[Content_Types].xml", xml => xml.Replace(
            "PartName=\"/ppt/slides/slide1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slide+xml\"",
            "PartName=\"/ppt/slides/slide1.xml\" ContentType=\"application/xml\""));
        Assert.Contains("PartName=\"/ppt/slides/slide1.xml\" ContentType=\"application/xml\"", System.Text.Encoding.UTF8.GetString(ReadParts(source)["[Content_Types].xml"]));

        Exception? failure = Record.Exception(() => {
            using var presentation = PowerPointDocument.Load(new MemoryStream(source));
            Assert.True(presentation.IsSdkLoaded);
        });
        Exception? sdkFailure = Record.Exception(() => SaveThroughSdk(source));
        Assert.Equal(sdkFailure?.GetType(), failure?.GetType());
    }

    [Theory]
    [InlineData(ApplicationPart, true)]
    [InlineData(CorePart, false)]
    public void PropertyPartsOverThePartSizeLimitGoThroughTheSdk(string part, bool atLoad) {
        byte[] source = CreateDeckWithStaleProperties();
        long limit = ReadParts(source)[PresentationPart].Length + 64;
        source = Rewrite(source, part, xml => xml.Replace("?>", "?>" + new string(' ', (int)limit)));
        var options = new PowerPointLoadOptions { OpenSettings = new OfficeOpenXmlLoadSettings { MaxCharactersInPart = limit } };

        Exception? failure = Record.Exception(() => {
            using var presentation = PowerPointDocument.Load(new MemoryStream(source), options);
            Assert.Equal(atLoad, presentation.IsSdkLoaded);
            presentation.ToBytes();
            Assert.True(presentation.IsSdkLoaded);
        });
        Assert.True(failure is null or not Xunit.Sdk.XunitException, failure?.ToString());
    }

    // A slide list entry the SDK cannot resolve fails the load, as it did before.
    [Fact]
    public void ALoadTheSdkRejectsStillFails() {
        byte[] source = Rewrite(CreateDeckWithStaleProperties(), PresentationPart, xml => Regex.Replace(xml, "r:id=\"rId\\d+\"", "r:id=\"rIdMissing\"", RegexOptions.None, TimeSpan.FromSeconds(1)));

        Assert.ThrowsAny<Exception>(() => PowerPointDocument.Load(new MemoryStream(source)));
    }

    internal static byte[] SaveThroughSdk(byte[] source) {
        using var presentation = PowerPointDocument.Load(new MemoryStream(source));
        _ = presentation.OpenXmlDocument;
        return presentation.ToBytes();
    }

    // Two decks are the same when their parts, content types and relationships match; docProps/app.xml and
    // presentation.xml hold the same XML; core properties differ only in their modification time; and every other part
    // has the loaded bytes.
    internal static void AssertSamePackage(byte[] source, byte[] throughSdk, byte[] withoutSdk, bool keptAsLoaded = true) {
        Dictionary<string, byte[]> expected = ReadParts(throughSdk);
        Dictionary<string, byte[]> actual = ReadParts(withoutSdk);
        Dictionary<string, byte[]> loaded = ReadParts(source);
        Assert.Equal(expected.Keys.OrderBy(name => name, StringComparer.Ordinal), actual.Keys.OrderBy(name => name, StringComparer.Ordinal));
        foreach (KeyValuePair<string, byte[]> part in actual) {
            if (part.Key == CorePart) {
                List<string> expectedCore = CoreValues(expected[part.Key]);
                List<string> actualCore = CoreValues(part.Value);
                Assert.True(expectedCore.SequenceEqual(actualCore), "Core properties differ. SDK only: "
                    + string.Join(" | ", expectedCore.Except(actualCore)) + "; without the SDK only: " + string.Join(" | ", actualCore.Except(expectedCore)));
                Assert.True(Regex.IsMatch(Modified(part.Value), ModifiedFormat(expected[part.Key])),
                    "SDK '" + Modified(expected[part.Key]) + "', without the SDK '" + Modified(part.Value) + "'.");
            } else if (part.Key == "_rels/.rels") {
                // Nothing in a package refers to its own relationship ids, and .NET Framework's packaging writes them
                // and their targets in its own spelling; what matters is which parts they reach.
                Assert.Equal(PackageRelationships(expected[part.Key]), PackageRelationships(part.Value));
            } else if (part.Key == "[Content_Types].xml") {
                // .NET Framework's packaging lays the catalog out differently; what matters is each part's content type.
                Assert.Equal(ContentTypes(throughSdk), ContentTypes(withoutSdk));
            } else if (part.Key.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) || part.Key.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)) {
                if (!XNode.DeepEquals(Normalize(expected[part.Key]), Normalize(part.Value)) && Environment.GetEnvironmentVariable("OFFICEIMO_DUMP_DIFF") is { } dump) {
                    File.WriteAllText(Path.Combine(dump, "expected.xml"), Normalize(expected[part.Key]).ToString());
                    File.WriteAllText(Path.Combine(dump, "actual.xml"), Normalize(part.Value).ToString());
                }
                XDocument expectedXml = Normalize(expected[part.Key], part.Key, RelationshipTargets(expected, part.Key));
                XDocument actualXml = Normalize(part.Value, part.Key, RelationshipTargets(actual, part.Key));
                Assert.True(XNode.DeepEquals(expectedXml, actualXml), part.Key + " differs from the SDK save."
                    + Environment.NewLine + Shorten(expectedXml) + Environment.NewLine + Shorten(actualXml));
                if (keptAsLoaded && part.Key != ApplicationPart && loaded.TryGetValue(part.Key, out byte[]? original)) {
                    Assert.True(original.AsSpan().SequenceEqual(part.Value), part.Key + " was not kept as loaded.");
                }
            } else {
                Assert.Equal(expected[part.Key], part.Value);
            }
        }
    }

    private static string Shorten(XDocument document) {
        string text = document.ToString(SaveOptions.DisableFormatting);
        return text.Length <= 2000 ? text : text.Substring(0, 2000) + "...";
    }

    private static List<string> PackageRelationships(byte[] relationships) =>
        XDocument.Load(new MemoryStream(relationships)).Root!.Elements()
            .Where(element => !((string?)element.Attribute("Target") ?? string.Empty).Contains("package/services/metadata/core-properties/"))
            .Select(element => (string?)element.Attribute("Type") + " " + ((string?)element.Attribute("Target") ?? string.Empty).TrimStart('/')
                + " " + ((string?)element.Attribute("TargetMode") ?? "Internal"))
            .OrderBy(value => value, StringComparer.Ordinal)
            .ToList();

    private static List<string> ContentTypes(byte[] package) {
        using var archive = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        XDocument catalog;
        using (Stream stream = archive.GetEntry("[Content_Types].xml")!.Open()) catalog = XDocument.Load(stream);
        XNamespace types = "http://schemas.openxmlformats.org/package/2006/content-types";
        var overrides = catalog.Root!.Elements(types + "Override").GroupBy(element => ((string)element.Attribute("PartName")!).ToUpperInvariant())
            .ToDictionary(group => group.Key, group => (string)group.First().Attribute("ContentType")!);
        var defaults = catalog.Root.Elements(types + "Default").GroupBy(element => ((string)element.Attribute("Extension")!).ToUpperInvariant())
            .ToDictionary(group => group.Key, group => (string)group.First().Attribute("ContentType")!);
        return archive.Entries.Where(entry => entry.FullName != "[Content_Types].xml" && !entry.FullName.EndsWith("/", StringComparison.Ordinal)
                && !entry.FullName.StartsWith("package/services/metadata/core-properties/", StringComparison.Ordinal))
            .Select(entry => {
                string name = "/" + entry.FullName;
                string extension = Path.GetExtension(entry.FullName).TrimStart('.').ToUpperInvariant();
                string? type = overrides.TryGetValue(name.ToUpperInvariant(), out string? found) ? found
                    : defaults.TryGetValue(extension, out string? fallback) ? fallback : null;
                return name + "=" + type;
            })
            .OrderBy(value => value, StringComparer.Ordinal)
            .ToList();
    }

    private static string Modified(byte[] core) => XDocument.Load(new MemoryStream(core)).Root!.Element(DublinCoreTerms + "modified")!.Value;

    // The SDK's spelling of the time, with digits generalised.
    private static string ModifiedFormat(byte[] core) => "^" + Regex.Replace(Regex.Escape(Modified(core)), "[0-9]", "[0-9]") + "$";

    // The SDK writes the core properties again in its own order and date spelling (created, modified and lastPrinted);
    // the values are what must match.
    private static List<string> CoreValues(byte[] core) =>
        XDocument.Load(new MemoryStream(core)).Root!.Elements()
            .Where(element => element.Name != DublinCoreTerms + "modified")
            .Select(element => element.Name + "=" + (element.Name.Namespace == DublinCoreTerms || element.Name.LocalName == "lastPrinted"
                ? DateTime.Parse(element.Value, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.AdjustToUniversal).ToString("O")
                : element.Value))
            .OrderBy(value => value, StringComparer.Ordinal)
            .ToList();

    // The internal targets of a part's relationships by id, resolved to part names; empty when it has none.
    private static Dictionary<string, string> RelationshipTargets(Dictionary<string, byte[]> parts, string partName) {
        var targets = new Dictionary<string, string>(StringComparer.Ordinal);
        int slash = partName.LastIndexOf('/');
        string folder = slash < 0 ? string.Empty : partName.Substring(0, slash + 1);
        if (!parts.TryGetValue(folder + "_rels/" + partName.Substring(slash + 1) + ".rels", out byte[]? relationships)) return targets;
        var source = new Uri("http://package/" + folder);
        foreach (XElement relationship in XDocument.Load(new MemoryStream(relationships)).Root!.Elements()) {
            string? id = (string?)relationship.Attribute("Id"), target = (string?)relationship.Attribute("Target");
            if (id == null || target == null || (string?)relationship.Attribute("TargetMode") == "External") continue;
            targets[id] = new Uri(source, target).AbsolutePath;
        }
        return targets;
    }

    // Declarations and indentation are the SDK's formatting, not content.
    private static XDocument Normalize(byte[] xml, string? partName = null, Dictionary<string, string>? relationshipTargets = null) {
        XDocument document = XDocument.Load(new MemoryStream(xml));
        // .NET Framework's packaging writes internal targets as absolute part names; compare the parts they reach.
        if (partName != null && partName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)) {
            string folder = partName.Contains("_rels/") ? partName.Substring(0, partName.LastIndexOf("_rels/", StringComparison.Ordinal)) : string.Empty;
            var source = new Uri("http://package/" + folder);
            foreach (XElement relationship in document.Root!.Elements().Where(element => (string?)element.Attribute("TargetMode") != "External")) {
                if (relationship.Attribute("Target") is { } target) target.Value = new Uri(source, target.Value).AbsolutePath;
            }
        }
        foreach (XElement element in document.Descendants()) element.Attributes().Where(attribute => attribute.IsNamespaceDeclaration).Remove();
        // A core properties part the SDK creates gets a random name, so its relationship differs between any two saves,
        // and so does the id the SDK gives any relationship it adds.
        document.Descendants().Where(element => element.Name.LocalName == "Relationship"
            && ((string?)element.Attribute("Target") ?? string.Empty).Contains("package/services/metadata/core-properties/")).Remove();
        // .NET Framework's packaging writes external targets as System.Uri spells them, so compare them as URIs.
        foreach (XElement relationship in document.Descendants().Where(element => element.Name.LocalName == "Relationship"
            && (string?)element.Attribute("TargetMode") == "External").ToList()) {
            if (relationship.Attribute("Target") is { } target && Uri.TryCreate(target.Value, UriKind.Absolute, out Uri? uri)) target.Value = uri.AbsoluteUri;
        }
        foreach (XAttribute id in document.Descendants().Where(element => element.Name.LocalName == "Relationship")
            .Select(element => element.Attribute("Id")).OfType<XAttribute>().Where(id => Regex.IsMatch(id.Value, "^R[0-9a-f]{16}$")).ToList()) {
            id.Value = "R";
        }
        // Markup refers to those ids through relationship attributes, such as a picture's r:embed; each compares by the
        // part its relationship reaches, so references to different parts stay different.
        foreach (XAttribute reference in document.Descendants().Attributes()
            .Where(attribute => attribute.Name.NamespaceName == "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                && Regex.IsMatch(attribute.Value, "^R[0-9a-f]{16}$")).ToList()) {
            reference.Value = relationshipTargets != null && relationshipTargets.TryGetValue(reference.Value, out string? target) ? "R:" + target : "R";
        }
        foreach (XText text in document.DescendantNodes().OfType<XText>().Where(text => string.IsNullOrWhiteSpace(text.Value)).ToList()) text.Remove();
        // The SDK writes attributes in schema order and empty elements in short form.
        foreach (XElement element in document.Descendants().ToList()) {
            element.ReplaceAttributes(element.Attributes().OrderBy(attribute => attribute.Name.ToString(), StringComparer.Ordinal).ToList());
            if (!element.Nodes().Any()) element.Value = string.Empty;
        }
        return document;
    }

    internal static Dictionary<string, byte[]> ReadParts(byte[] package) {
        var parts = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        using var archive = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        foreach (ZipArchiveEntry entry in archive.Entries) {
            // System.IO.Packaging names a core properties part it creates at random, on every save.
            if (entry.FullName.StartsWith("package/services/metadata/core-properties/", StringComparison.Ordinal)) continue;
            using Stream stream = entry.Open();
            var buffer = new MemoryStream();
            stream.CopyTo(buffer);
            parts[entry.FullName] = buffer.ToArray();
        }
        return parts;
    }

    internal static XDocument ReadXml(byte[] package, string name) => XDocument.Load(new MemoryStream(ReadParts(package)[name]));

    internal static byte[] Rewrite(byte[] package, string member, Func<string, string> change) {
        var stream = new MemoryStream();
        stream.Write(package, 0, package.Length);
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true)) {
            ZipArchiveEntry entry = archive.GetEntry(member)!;
            string xml;
            using (var reader = new StreamReader(entry.Open())) xml = reader.ReadToEnd();
            entry.Delete();
            using var writer = new StreamWriter(archive.CreateEntry(member).Open(), new System.Text.UTF8Encoding(false));
            writer.Write(change(xml));
        }
        return stream.ToArray();
    }

    // Two slides, the second hidden and with notes, on a 4:3 slide, saved and then given stale counts.
    internal static byte[] CreateDeckWithStaleProperties() {
        byte[] deck;
        using (var presentation = PowerPointDocument.Create(new MemoryStream())) {
            presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen4x3);
            presentation.AddSlide().AddTextBox("First");
            PowerPointSlide second = presentation.AddSlide();
            second.AddTextBox("Second");
            second.Hidden = true;
            second.Notes.Text = "Notes";
            deck = presentation.ToBytes();
        }
        deck = Rewrite(deck, CorePart, xml => Regex.Replace(xml, "(<dcterms:modified[^>]*>)[^<]*(</dcterms:modified>)", "${1}2001-01-01T00:00:00Z${2}"));
        return Rewrite(deck, ApplicationPart, xml => Regex.Replace(Regex.Replace(Regex.Replace(xml,
            "<Slides>\\d+</Slides>", "<Slides>9</Slides>"), "<HiddenSlides>\\d+</HiddenSlides>", "<HiddenSlides>0</HiddenSlides>"),
            "<PresentationFormat>[^<]*</PresentationFormat>", "<PresentationFormat>On-screen Show (4:3)</PresentationFormat>"));
    }

    private static int CountValidationErrors(byte[] package) {
        using var stream = new MemoryStream(package);
        using PresentationDocument document = PresentationDocument.Open(stream, false);
        return new OpenXmlValidator().Validate(document).Count();
    }
}
