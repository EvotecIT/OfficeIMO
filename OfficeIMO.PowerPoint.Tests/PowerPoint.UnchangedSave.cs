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
        Assert.True(modified >= before.AddSeconds(-1), modified.ToString("O"));
        AssertSamePackage(source, SaveThroughSdk(source), saved);
        Assert.Equal(0, CountValidationErrors(saved));
    }

    [Fact]
    public void OpensTheSdkOnFirstUseAndSavesThroughIt() {
        byte[] source = CreateDeckWithStaleProperties();
        using var presentation = PowerPointDocument.Load(new MemoryStream(source));
        Assert.False(presentation.IsSdkLoaded);

        Assert.Equal(2, presentation.Slides.Count);
        Assert.True(presentation.IsSdkLoaded);
        Assert.Equal(1, presentation.DeferredSdkOpenCount);
        presentation.Slides[0].AddTextBox("Added");

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
            _ => Rewrite(source, ApplicationPart, xml => Regex.Replace(xml, "</(\\w+:)?Properties>\\s*$", match => "<ap:DigitalSignature xmlns:ap=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\"><vt:blob xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">AA==</vt:blob></ap:DigitalSignature>" + match.Value))
        };

        using var presentation = PowerPointDocument.Load(new MemoryStream(source));
        // A signed package opens the SDK at load, where its fingerprint is taken.
        if (variant == "signed") Assert.True(presentation.IsSdkLoaded);
        byte[] saved = presentation.ToBytes();

        Assert.True(presentation.IsSdkLoaded);
        AssertSamePackage(source, SaveThroughSdk(source), saved, keptAsLoaded: false);
    }

    // A slide list entry the SDK cannot resolve fails the load, as it did before.
    [Fact]
    public void ALoadTheSdkRejectsStillFails() {
        byte[] source = Rewrite(CreateDeckWithStaleProperties(), PresentationPart, xml => Regex.Replace(xml, "r:id=\"rId\\d+\"", "r:id=\"rIdMissing\"", RegexOptions.None, TimeSpan.FromSeconds(1)));

        Assert.ThrowsAny<Exception>(() => PowerPointDocument.Load(new MemoryStream(source)));
    }

    private static byte[] SaveThroughSdk(byte[] source) {
        using var presentation = PowerPointDocument.Load(new MemoryStream(source));
        _ = presentation.OpenXmlDocument;
        return presentation.ToBytes();
    }

    // Two decks are the same when their parts, content types and relationships match; docProps/app.xml and
    // presentation.xml hold the same XML; core properties differ only in their modification time; and every other part
    // has the loaded bytes.
    private static void AssertSamePackage(byte[] source, byte[] throughSdk, byte[] withoutSdk, bool keptAsLoaded = true) {
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
            } else if (part.Key.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) || part.Key.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)) {
                if (!XNode.DeepEquals(Normalize(expected[part.Key]), Normalize(part.Value)) && Environment.GetEnvironmentVariable("OFFICEIMO_DUMP_DIFF") is { } dump) {
                    File.WriteAllText(Path.Combine(dump, "expected.xml"), Normalize(expected[part.Key]).ToString());
                    File.WriteAllText(Path.Combine(dump, "actual.xml"), Normalize(part.Value).ToString());
                }
                Assert.True(XNode.DeepEquals(Normalize(expected[part.Key]), Normalize(part.Value)), part.Key + " differs from the SDK save.");
                if (keptAsLoaded && part.Key != ApplicationPart && loaded.TryGetValue(part.Key, out byte[]? original)) {
                    Assert.True(original.AsSpan().SequenceEqual(part.Value), part.Key + " was not kept as loaded.");
                }
            } else {
                Assert.Equal(expected[part.Key], part.Value);
            }
        }
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

    // Declarations and indentation are the SDK's formatting, not content.
    private static XDocument Normalize(byte[] xml) {
        XDocument document = XDocument.Load(new MemoryStream(xml));
        foreach (XElement element in document.Descendants()) element.Attributes().Where(attribute => attribute.IsNamespaceDeclaration).Remove();
        // A core properties part the SDK creates gets a random name, so its relationship differs between any two saves,
        // and so does the id the SDK gives any relationship it adds.
        document.Descendants().Where(element => element.Name.LocalName == "Relationship"
            && ((string?)element.Attribute("Target") ?? string.Empty).Contains("package/services/metadata/core-properties/")).Remove();
        foreach (XAttribute id in document.Descendants().Where(element => element.Name.LocalName == "Relationship")
            .Select(element => element.Attribute("Id")).OfType<XAttribute>().Where(id => Regex.IsMatch(id.Value, "^R[0-9a-f]{16}$")).ToList()) {
            id.Value = "R";
        }
        foreach (XText text in document.DescendantNodes().OfType<XText>().Where(text => string.IsNullOrWhiteSpace(text.Value)).ToList()) text.Remove();
        // The SDK writes attributes in schema order and empty elements in short form.
        foreach (XElement element in document.Descendants().ToList()) {
            element.ReplaceAttributes(element.Attributes().OrderBy(attribute => attribute.Name.ToString(), StringComparer.Ordinal).ToList());
            if (!element.Nodes().Any()) element.Value = string.Empty;
        }
        return document;
    }

    private static Dictionary<string, byte[]> ReadParts(byte[] package) {
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

    private static XDocument ReadXml(byte[] package, string name) => XDocument.Load(new MemoryStream(ReadParts(package)[name]));

    private static byte[] Rewrite(byte[] package, string member, Func<string, string> change) {
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
    private static byte[] CreateDeckWithStaleProperties() {
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
