using System.IO.Compression;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>
/// Documents that load without the Open XML SDK read and edit paragraphs, runs, text and basic formatting on OfficeIMO's
/// own Open XML layer. Each test runs the same steps through that path and through the SDK path and expects the same
/// result, or checks how the two meet when a caller mixes OfficeIMO with direct SDK edits.
/// </summary>
public sealed class WordWithoutSdkTests {
    private static readonly XNamespace Wn = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly string Documents = Path.Combine(AppContext.BaseDirectory, "Documents");

    [Fact]
    public void PlainDocumentLoadsReadsEditsAndSavesWithoutTheSdk() {
        byte[] source = CreatePackage(BodyXml(
            "<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space=\"preserve\">Bold </w:t></w:r><w:r><w:t>plain</w:t></w:r></w:p>"
            + "<w:p><w:r><w:t xml:space=\"preserve\">  edges  </w:t></w:r><w:r><w:t>  trimmed  </w:t></w:r></w:p>"));

        using WordDocument document = WordDocument.Load(new MemoryStream(source));
        Assert.True(document.IsSdkDeferred);

        List<WordParagraph> paragraphs = document.Paragraphs;
        Assert.Equal(new[] { "Bold ", "plain", "  edges  ", "trimmed" }, paragraphs.Select(paragraph => paragraph.Text));
        Assert.Equal(new[] { true, false, false, false }, paragraphs.Select(paragraph => paragraph.Bold));

        WordParagraph added = document.AddParagraph("Added");
        added.Italic = true;
        added.FontSize = OfficeLength.Points(12);
        using var saved = new MemoryStream();
        document.Save(saved);

        Assert.True(document.IsSdkDeferred);
        Assert.Empty(document.OpenXmlSdkFallbacks);
        using WordDocument reopened = WordDocument.Load(new MemoryStream(saved.ToArray()));
        WordParagraph last = reopened.Paragraphs.Last();
        Assert.Equal("Added", last.Text);
        Assert.True(last.Italic);
        Assert.Equal(12D, last.FontSize?.ToPoints());
    }

    [Fact]
    public void ReadsMatchTheSdkPathOnTheTestDocuments() {
        int compared = 0;
        foreach (string path in Directory.EnumerateFiles(Documents, "*.docx").OrderBy(file => file, StringComparer.Ordinal)) {
            byte[] bytes = File.ReadAllBytes(path);
            List<string>? sdk = ReadParagraphs(bytes, withoutSdk: false);
            List<string>? owned = ReadParagraphs(bytes, withoutSdk: true);
            if (sdk == null) {
                Assert.Null(owned);
                continue;
            }
            Assert.Equal(sdk, owned);
            compared++;
        }
        Assert.True(compared > 20, $"Only {compared} documents were compared.");
    }

    [Fact]
    public void EditsAndSavesMatchTheSdkPathOnTheTestDocuments() {
        int compared = 0;
        foreach (string path in Directory.EnumerateFiles(Documents, "*.docx").OrderBy(file => file, StringComparer.Ordinal)) {
            byte[] bytes = File.ReadAllBytes(path);
            byte[]? sdk = EditAndSave(bytes, withoutSdk: false);
            byte[]? owned = EditAndSave(bytes, withoutSdk: true);
            if (sdk == null) {
                Assert.Null(owned);
                continue;
            }
            AssertSamePackage(sdk, owned!, Path.GetFileName(path));
            compared++;
        }
        Assert.True(compared > 20, $"Only {compared} documents were compared.");
    }

    [Fact]
    public void TextSetterKeepsPageBreaksAndWritesTabsAndLineBreaksLikeTheSdkPath() {
        byte[] source = CreatePackage(BodyXml(
            "<w:p><w:r><w:t>one</w:t><w:br w:type=\"page\"/><w:t>two</w:t><w:tab/><w:t>three</w:t></w:r></w:p>"));
        void Edit(WordDocument document) {
            WordParagraph paragraph = Assert.Single(document.Paragraphs);
            Assert.Equal("one\u2028two\tthree", paragraph.Text);
            paragraph.Text = "first\u2028second\tthird\nfourth";
            paragraph.AddText("a\nb");
        }
        AssertSamePackage(EditAndSave(source, false, Edit)!, EditAndSave(source, true, Edit)!, "breaks");
    }

    [Fact]
    public void FormattingMembersMatchTheSdkPath() {
        byte[] source = CreatePackage(BodyXml(
            "<w:p><w:r><w:rPr><w:rFonts w:ascii=\"Calibri\"/><w:b w:val=\"0\"/><w:u w:val=\"single\" w:color=\"FF0000\"/><w:sz w:val=\"30\"/>"
            + "<w:highlight w:val=\"green\"/><w:caps/></w:rPr><w:t>formatted</w:t></w:r></w:p><w:p/>"));
        string Read(WordDocument document) => string.Join("|", document.Paragraphs.Select(paragraph =>
            $"{paragraph.Bold},{paragraph.Italic},{paragraph.Underline},{paragraph.Strikethrough},{paragraph.DoubleStrikethrough},"
            + $"{paragraph.FontSize},{paragraph.Color},{paragraph.FontFamily},{paragraph.Highlight},{paragraph.CapsStyle}"));
        void Edit(WordDocument document) {
            Assert.Equal(Read(Open(source, false)), Read(document));
            List<WordParagraph> paragraphs = document.Paragraphs;
            WordParagraph first = paragraphs[0];
            first.Bold = true;
            first.Underline = WordUnderlineStyle.Wave;
            first.DoubleStrikethrough = true;
            first.FontSize = null;
            first.Highlight = null;
            first.CapsStyle = WordCapsStyle.SmallCaps;
            first.FontFamily = "Arial";
            Assert.True(OfficeColor.TryParseHex("336699", out OfficeColor color));
            first.Color = color;
            WordParagraph empty = paragraphs[1];
            empty.Highlight = null;
            empty.Strikethrough = true;
            empty.Text = "now with text";
        }
        AssertSamePackage(EditAndSave(source, false, Edit)!, EditAndSave(source, true, Edit)!, "formatting");
    }

    [Fact]
    public void OfficeImoAndSdkEditsMixAndSurviveSave() {
        byte[] source = CreatePackage(BodyXml("<w:p><w:r><w:t>Original</w:t></w:r></w:p>"));
        using var output = new MemoryStream();
        using (WordDocument document = WordDocument.Load(new MemoryStream(source))) {
            WordParagraph original = document.Paragraphs[0];
            original.Bold = true;
            WordParagraph before = document.AddParagraph("Added before the SDK opened");
            Assert.True(document.IsSdkDeferred);

            WordprocessingDocument package = document.OpenXmlDocument;
            Assert.False(document.IsSdkDeferred);
            Assert.Equal(new[] { "WordDocument.OpenXmlDocument" }, document.OpenXmlSdkFallbacks);
            Body body = package.MainDocumentPart!.Document!.Body!;
            Assert.Equal("Original", body.Elements<Paragraph>().First().InnerText);
            body.InsertBefore(new Paragraph(new Run(new Text("Added through the SDK"))), body.GetFirstChild<SectionProperties>());

            // Paragraphs handed out before the SDK opened now edit the SDK package.
            before.Italic = true;
            original.Text = "Original, edited";
            document.AddParagraph("Added after the SDK opened").Underline = WordUnderlineStyle.Single;
            document.Save(output);
        }

        using WordDocument reopened = WordDocument.Load(new MemoryStream(output.ToArray()));
        List<WordParagraph> paragraphs = reopened.Paragraphs;
        Assert.Equal(new[] { "Original, edited", "Added before the SDK opened", "Added through the SDK", "Added after the SDK opened" },
            paragraphs.Select(paragraph => paragraph.Text));
        Assert.True(paragraphs[0].Bold);
        Assert.True(paragraphs[1].Italic);
        Assert.Equal(WordUnderlineStyle.Single, paragraphs[3].Underline);
    }

    [Fact]
    public void MembersTheOwnedLayerLacksOpenTheSdkAndSayWhy() {
        byte[] source = CreatePackage(BodyXml("<w:p><w:r><w:t>Text</w:t></w:r></w:p>"));
        using WordDocument document = WordDocument.Load(new MemoryStream(source));
        Assert.True(document.IsSdkDeferred);

        WordParagraph paragraph = document.Paragraphs[0];
        paragraph.ParagraphAlignment = WordParagraphAlignment.Center;

        Assert.False(document.IsSdkDeferred);
        string reason = Assert.Single(document.OpenXmlSdkFallbacks);
        Assert.Contains("ParagraphAlignment", reason, StringComparison.Ordinal);
        Assert.Equal(WordParagraphAlignment.Center, document.Paragraphs[0].ParagraphAlignment);
    }

    [Fact]
    public void ContentTheOwnedLayerDoesNotModelIsWrittenBackUntouched() {
        const string kept = "<w:customXml w:element=\"kept\"><w:r><w:t>inside</w:t></w:r></w:customXml>";
        byte[] source = CreatePackage(BodyXml("<w:p>" + kept + "<w:r><w:t>after</w:t></w:r></w:p>"));
        using WordDocument document = WordDocument.Load(new MemoryStream(source));
        Assert.Equal(new[] { "inside", "after" }, document.Paragraphs.Select(paragraph => paragraph.Text));
        document.Paragraphs[1].Bold = true;
        using var saved = new MemoryStream();
        document.Save(saved);

        Assert.True(document.IsSdkDeferred);
        XElement paragraphElement = ReadPart(saved.ToArray(), "word/document.xml").Descendants(Wn + "p").Single();
        XElement customXml = paragraphElement.Element(Wn + "customXml")!;
        Assert.Equal("kept", (string?)customXml.Attribute(Wn + "element"));
        Assert.Equal("inside", customXml.Value);
    }

    [Fact]
    public void StrictAndSignedPackagesKeepLoadingThroughTheSdk() {
        byte[] strict = CreatePackage(BodyXml("<w:p/>").Replace(
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "http://purl.oclc.org/ooxml/wordprocessingml/main"));
        Assert.Null(TryOpenOwned(strict));

        byte[] signed = CreatePackage(BodyXml("<w:p/>"), extraPackageRelationship:
            "<Relationship Id=\"rSig\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin\" Target=\"_xmlsignatures/origin.sigs\"/>");
        Assert.Null(TryOpenOwned(signed));
    }

    [Fact]
    public void ValueTokensMatchTheSdk() {
        foreach (WordUnderlineStyle style in Enum.GetValues(typeof(WordUnderlineStyle))) {
            Assert.Equal(((IEnumValue)style.ToOpenXml()).Value, WordUnderlineStyleTokens.ToToken(style));
            Assert.True(WordUnderlineStyleTokens.TryParse(WordUnderlineStyleTokens.ToToken(style), out WordUnderlineStyle parsed));
            Assert.Equal(style, parsed);
        }
        foreach (WordHighlightColor color in Enum.GetValues(typeof(WordHighlightColor))) {
            Assert.Equal(((IEnumValue)color.ToOpenXml()).Value, color.ToOpenXmlToken());
        }
    }

    private static OfficeIMO.Word.OpenXml.WordPackageModel? TryOpenOwned(byte[] bytes) {
        OfficeIMO.Word.OpenXml.WordPackageModel? model = OfficeIMO.Word.OpenXml.WordPackageModel.TryOpen(bytes, 0);
        model?.Dispose();
        return model;
    }

    private static WordDocument Open(byte[] bytes, bool withoutSdk) {
        bool previous = WordDocument.DisableOwnedOpenXmlForTesting.Value;
        WordDocument.DisableOwnedOpenXmlForTesting.Value = !withoutSdk;
        try {
            return WordDocument.Load(new MemoryStream(bytes));
        } finally {
            WordDocument.DisableOwnedOpenXmlForTesting.Value = previous;
        }
    }

    private static List<string>? ReadParagraphs(byte[] bytes, bool withoutSdk) {
        WordDocument document;
        try {
            document = Open(bytes, withoutSdk);
        } catch (Exception exception) when (exception is not OutOfMemoryException) {
            return null;
        }
        using (document) {
            return document.Paragraphs.Select(paragraph =>
                $"{paragraph.Text}|{paragraph.Bold}|{paragraph.Italic}|{paragraph.Underline}|{paragraph.Strikethrough}|{paragraph.FontSize}"
                + $"|{paragraph.Color}|{paragraph.FontFamily}|{paragraph.Highlight}|{paragraph.CapsStyle}").ToList();
        }
    }

    private static void StandardEdit(WordDocument document) {
        WordParagraph added = document.AddParagraph("Hello world");
        added.Bold = true;
        added.Underline = WordUnderlineStyle.Double;
        added.FontFamily = "Arial";
        added.AddText(" second").Strikethrough = true;
        document.AddParagraph("tab\there\nnext line");
        List<WordParagraph> paragraphs = document.Paragraphs;
        if (paragraphs.Count > 0) {
            paragraphs[0].Italic = true;
            paragraphs[0].Text = paragraphs[0].Text + " edited";
        }
    }

    private static byte[]? EditAndSave(byte[] bytes, bool withoutSdk, Action<WordDocument>? edit = null) {
        WordDocument document;
        try {
            document = Open(bytes, withoutSdk);
        } catch (Exception exception) when (exception is not OutOfMemoryException) {
            return null;
        }
        using (document) {
            (edit ?? StandardEdit)(document);
            using var output = new MemoryStream();
            document.Save(output);
            return output.ToArray();
        }
    }

    private static void AssertSamePackage(byte[] expected, byte[] actual, string name) {
        Dictionary<string, byte[]> expectedParts = ReadParts(expected);
        Dictionary<string, byte[]> actualParts = ReadParts(actual);
        Assert.Equal(expectedParts.Keys.OrderBy(key => key, StringComparer.Ordinal), actualParts.Keys.OrderBy(key => key, StringComparer.Ordinal));
        foreach (KeyValuePair<string, byte[]> part in expectedParts) {
            byte[] other = actualParts[part.Key];
            if (part.Value.AsSpan().SequenceEqual(other)) continue;
            if (!part.Key.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) && !part.Key.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)) {
                Assert.Fail($"{name}: binary part {part.Key} differs.");
            }
            Assert.True(Canonical(part.Value) == Canonical(other), $"{name}: {part.Key} differs.\nSDK:   {Canonical(part.Value)}\nOwned: {Canonical(other)}");
        }
    }

    // Element and attribute names with their namespaces, values, and text, in document order; attribute order, prefixes,
    // namespace declarations and whitespace-only text between elements don't count. Revision identifiers are random on
    // both paths.
    private static string Canonical(byte[] xml) {
        XDocument document;
        using (var stream = new MemoryStream(xml)) document = XDocument.Load(stream);
        var builder = new System.Text.StringBuilder();
        void Write(XElement element) {
            builder.Append('<').Append(element.Name.ToString());
            foreach (XAttribute attribute in element.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration)
                         .OrderBy(attribute => attribute.Name.ToString(), StringComparer.Ordinal)) {
                bool revision = attribute.Name.Namespace == Wn && attribute.Name.LocalName.StartsWith("rsid", StringComparison.Ordinal);
                builder.Append(' ').Append(attribute.Name.ToString()).Append("=\"").Append(revision ? "*" : attribute.Value).Append('"');
            }
            builder.Append('>');
            foreach (XNode node in element.Nodes()) {
                if (node is XElement child) Write(child);
                else if (node is XText text && !string.IsNullOrWhiteSpace(text.Value)) builder.Append(text.Value);
                else if (node is XText significant && element.Name == Wn + "t") builder.Append(significant.Value);
            }
            builder.Append("</>");
        }
        Write(document.Root!);
        return builder.ToString();
    }

    private static Dictionary<string, byte[]> ReadParts(byte[] package) {
        using var archive = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        return archive.Entries.ToDictionary(entry => entry.FullName, entry => {
            using Stream stream = entry.Open();
            using var buffer = new MemoryStream();
            stream.CopyTo(buffer);
            return buffer.ToArray();
        }, StringComparer.Ordinal);
    }

    private static XDocument ReadPart(byte[] package, string name) {
        using var stream = new MemoryStream(ReadParts(package)[name]);
        return XDocument.Load(stream);
    }

    private static string BodyXml(string body) =>
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        + "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>" + body
        + "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr></w:body></w:document>";

    private static byte[] CreatePackage(string documentXml, string extraPackageRelationship = "") {
        using var output = new MemoryStream();
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
            void Add(string name, string content) {
                using var writer = new StreamWriter(archive.CreateEntry(name).Open());
                writer.Write(content);
            }
            Add("[Content_Types].xml", "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                + "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/></Types>");
            Add("_rels/.rels", "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>"
                + extraPackageRelationship + "</Relationships>");
            Add("word/document.xml", documentXml);
        }
        return output.ToArray();
    }
}
