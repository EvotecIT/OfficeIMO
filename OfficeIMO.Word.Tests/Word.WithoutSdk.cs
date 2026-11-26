using System.IO.Compression;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;
#if NET

/// <summary>
/// Documents that load without the Open XML SDK read and edit paragraphs, runs, text and basic formatting on OfficeIMO's
/// own Open XML layer. Each test runs the same steps through that path and through the SDK path and expects the same
/// result, or checks how the two meet when a caller mixes OfficeIMO with direct SDK edits.
/// </summary>
public sealed class WordWithoutSdkTests {
    private static readonly XNamespace Wn = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace Mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    private static readonly string Documents = Path.Combine(AppContext.BaseDirectory, "Documents");
    // Whether the last read or save through the helpers below finished without opening the SDK.
    private static bool _lastRunStayedOwned;

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
    public void CreatedDocumentsMatchTheSdkPath() {
        var edits = new Dictionary<string, Action<WordDocument>> {
            ["empty"] = _ => { },
            ["standard"] = StandardEdit,
            ["styled"] = document => {
                WordParagraph heading = document.AddParagraph("Heading");
                heading.Style = WordParagraphStyles.Heading1;
                heading.ParagraphAlignment = WordParagraphAlignment.Center;
                document.AddParagraph("Title").Style = WordParagraphStyles.Title;
                document.AddParagraph("Body").Style = WordParagraphStyles.Normal;
            },
            ["many"] = document => {
                for (int index = 0; index < 50; index++) {
                    WordParagraph paragraph = document.AddParagraph("Paragraph " + index);
                    paragraph.Bold = index % 2 == 0;
                    paragraph.FontSize = OfficeLength.Points(10 + index % 5);
                    paragraph.Color = OfficeColor.FromRgb(0x1F, 0x4E, 0x79);
                }
            },
        };
        foreach (KeyValuePair<string, Action<WordDocument>> edit in edits) {
            byte[] sdk = CreateAndSave(withoutSdk: false, edit.Value, out _);
            byte[] owned = CreateAndSave(withoutSdk: true, edit.Value, out bool stayedOwned);
            Assert.True(stayedOwned, edit.Key);
            AssertSamePackage(sdk, owned, "created, " + edit.Key);
        }
    }

    private static byte[] CreateAndSave(bool withoutSdk, Action<WordDocument> edit, out bool stayedOwned) {
        bool previous = WordDocument.DisableOwnedOpenXmlForTesting.Value;
        WordDocument.DisableOwnedOpenXmlForTesting.Value = !withoutSdk;
        try {
            using WordDocument document = WordDocument.Create();
            edit(document);
            using var output = new MemoryStream();
            document.Save(output);
            stayedOwned = document.IsSdkDeferred;
            return output.ToArray();
        } finally {
            WordDocument.DisableOwnedOpenXmlForTesting.Value = previous;
        }
    }

    [Fact]
    public void StylesResolveToTheDocumentsOwnDefinitionsLikeTheSdkPath() {
        const string styles = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><w:styles xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
            + "<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Standard\"><w:name w:val=\"Normal\"/></w:style>"
            + "<w:style w:type=\"paragraph\" w:styleId=\"berschrift1\"><w:name w:val=\"heading 1\"/><w:basedOn w:val=\"Standard\"/><w:next w:val=\"Standard\"/></w:style>"
            + "</w:styles>";
        const string settings = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><w:settings xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"/>";
        byte[] source = CreatePackage(BodyXml("<w:p><w:pPr><w:pStyle w:val=\"berschrift1\"/></w:pPr><w:r><w:t>Kapitel</w:t></w:r></w:p>"),
            extraParts: new Dictionary<string, string> { ["word/styles.xml"] = styles, ["word/settings.xml"] = settings },
            extraContentTypes: "<Override PartName=\"/word/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml\"/>"
                + "<Override PartName=\"/word/settings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml\"/>",
            documentRelationships: "<Relationship Id=\"rStyles\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>"
                + "<Relationship Id=\"rSettings\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings\" Target=\"settings.xml\"/>");
        void Edit(WordDocument document) {
            Assert.Equal(WordParagraphStyles.Heading1, document.Paragraphs[0].Style);
            document.AddParagraph("Weiter").Style = WordParagraphStyles.Heading1;
            document.AddParagraph("Text").Style = WordParagraphStyles.Normal;
        }
        byte[] sdk = EditAndSave(source, false, Edit)!;
        byte[] owned = EditAndSave(source, true, Edit)!;
        Assert.True(_lastRunStayedOwned);
        AssertSamePackage(sdk, owned, "localized styles");
        Assert.Equal(new[] { "berschrift1", "berschrift1", "Standard" },
            ReadPart(owned, "word/document.xml").Descendants(Wn + "pStyle").Select(style => (string?)style.Attribute(Wn + "val")));
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
            if (_lastRunStayedOwned) compared++;
        }
        Assert.True(compared > 20, $"Only {compared} documents were read without the SDK.");
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
            if (_lastRunStayedOwned) compared++;
        }
        // Documents with tables still save through the SDK, so fewer of them count here than for reads.
        Assert.True(compared > 15, $"Only {compared} documents were edited and saved without the SDK.");
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
    public void ParagraphsHandedOutBeforeTheSdkOpenedFollowTheirElementsThroughLaterEdits() {
        byte[] source = CreatePackage(BodyXml(
            "<w:p><w:r><w:t>A</w:t></w:r></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p><w:p><w:r><w:t>C</w:t></w:r></w:p>"));
        using WordDocument document = WordDocument.Load(new MemoryStream(source));
        List<WordParagraph> paragraphs = document.Paragraphs;
        Assert.True(document.IsSdkDeferred);

        // Removing the first paragraph opens the SDK; the others must still reach their own elements afterwards.
        paragraphs[0].Remove();
        Assert.False(document.IsSdkDeferred);
        Assert.Equal("B", paragraphs[1].Text);
        paragraphs[1].Remove();
        paragraphs[2].Bold = true;

        WordParagraph remaining = Assert.Single(document.Paragraphs);
        Assert.Equal("C", remaining.Text);
        Assert.True(remaining.Bold);

        // A paragraph the SDK inserts in front doesn't take the place of one handed out before.
        Body body = document.OpenXmlDocument.MainDocumentPart!.Document!.Body!;
        body.PrependChild(new Paragraph(new Run(new Text("Inserted"))));
        paragraphs[2].Text = "C, edited";
        Assert.Equal(new[] { "Inserted", "C, edited" }, body.Elements<Paragraph>().Select(paragraph => paragraph.InnerText));
    }

    [Fact]
    public void BookmarkParagraphsKeepTheirBookmarkWhenTheyGetARun() {
        byte[] source = CreatePackage(BodyXml(
            "<w:p><w:bookmarkStart w:id=\"0\" w:name=\"Mark\"/><w:r><w:t>Text</w:t></w:r><w:bookmarkEnd w:id=\"0\"/></w:p>"));
        foreach (bool withoutSdk in new[] { false, true }) {
            using WordDocument document = Open(source, withoutSdk);
            WordParagraph bookmark = document.Paragraphs.Single(paragraph => paragraph.IsBookmark);
            bookmark.Bold = true;
            Assert.True(bookmark.IsBookmark, withoutSdk ? "owned" : "SDK");
            Assert.Equal("Mark", bookmark.Bookmark!.Name);
        }
        void Edit(WordDocument document) => document.Paragraphs.Single(paragraph => paragraph.IsBookmark).Text = "Added";
        AssertSamePackage(EditAndSave(source, false, Edit)!, EditAndSave(source, true, Edit)!, "bookmark");
    }

    [Fact]
    public void SavesWithTablesInNotesGoThroughTheSdk() {
        const string footnotes = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><w:footnotes xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
            + "<w:footnote w:id=\"1\"><w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl><w:p/></w:footnote></w:footnotes>";
        byte[] source = CreatePackage(BodyXml("<w:p><w:r><w:t>Text</w:t></w:r></w:p>"),
            extraParts: new Dictionary<string, string> { ["word/footnotes.xml"] = footnotes },
            extraContentTypes: "<Override PartName=\"/word/footnotes.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml\"/>",
            documentRelationships: "<Relationship Id=\"rNotes\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes\" Target=\"footnotes.xml\"/>");
        using WordDocument document = WordDocument.Load(new MemoryStream(source));
        document.AddParagraph("Added");
        document.Save(new MemoryStream());
        Assert.Contains("Save: tables in a header, footer, footnote or endnote", document.OpenXmlSdkFallbacks);
    }

    [Fact]
    public void SaveKeepsCleanupNumberingAfterTheNumberingInstances() {
        const string numbering = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><w:numbering xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
            + "<w:abstractNum w:abstractNumId=\"0\"/><w:numIdMacAtCleanup w:val=\"1\"/><w:num w:numId=\"1\"><w:abstractNumId w:val=\"0\"/></w:num></w:numbering>";
        byte[] source = CreatePackage(BodyXml("<w:p><w:r><w:t>Text</w:t></w:r></w:p>"),
            extraParts: new Dictionary<string, string> { ["word/numbering.xml"] = numbering },
            extraContentTypes: "<Override PartName=\"/word/numbering.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml\"/>",
            documentRelationships: "<Relationship Id=\"rNumbering\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering\" Target=\"numbering.xml\"/>");
        byte[] saved = EditAndSave(source, true, document => document.AddParagraph("Added"))!;
        Assert.Equal(new[] { "abstractNum", "num", "numIdMacAtCleanup" },
            ReadPart(saved, "word/numbering.xml").Root!.Elements().Select(element => element.Name.LocalName));
    }

    [Fact]
    public void SaveKeepsExternalRelationshipTargets() {
        const string external = "<Relationship Id=\"rLink\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"/word/guide.pdf\" TargetMode=\"External\"/>";
        byte[] source = CreatePackage(BodyXml("<w:p><w:r><w:t>Text</w:t></w:r></w:p>"), documentRelationships: external);
        void Edit(WordDocument document) => document.AddParagraph("Added");
        foreach (bool withoutSdk in new[] { false, true }) {
            XElement link = ReadPart(EditAndSave(source, withoutSdk, Edit)!, "word/_rels/document.xml.rels").Root!.Elements()
                .Single(relationship => (string?)relationship.Attribute("Id") == "rLink");
            Assert.Equal("/word/guide.pdf", (string?)link.Attribute("Target"));
        }
    }

    [Fact]
    public void CharacterLimitCoversEveryXmlPart() {
        string header = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><w:hdr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:p><w:r><w:t>"
            + new string('x', 4000) + "</w:t></w:r></w:p></w:hdr>";
        byte[] source = CreatePackage(BodyXml("<w:p><w:r><w:t>Text</w:t></w:r></w:p>").Replace("<w:sectPr>",
                "<w:sectPr><w:headerReference xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" w:type=\"default\" r:id=\"rHeader\"/>"),
            extraParts: new Dictionary<string, string> { ["word/header1.xml"] = header },
            extraContentTypes: "<Override PartName=\"/word/header1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml\"/>",
            documentRelationships: "<Relationship Id=\"rHeader\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header1.xml\"/>");
        // The SDK path rejects the header when it loads it; the owned path must not accept the document instead.
        WordDocument Load(long limit) =>
            WordDocument.Load(new MemoryStream(source), new WordLoadOptions { OpenSettings = new OfficeOpenXmlLoadSettings { MaxCharactersInPart = limit } });
        Assert.Throws<System.Xml.XmlException>(() => Load(2000));
        using WordDocument permitted = Load(8000);
        Assert.True(permitted.IsSdkDeferred);
    }

    [Fact]
    public void ParagraphsStayReadableAfterTheDocumentIsDisposed() {
        byte[] source = CreatePackage(BodyXml(
            "<w:p><w:pPr><w:jc w:val=\"center\"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>First</w:t></w:r></w:p><w:p><w:r><w:t>Second</w:t></w:r></w:p>"));
        List<WordParagraph> paragraphs;
        WordParagraph added, sameRun;
        using (WordDocument document = WordDocument.Load(new MemoryStream(source))) {
            paragraphs = document.Paragraphs;
            sameRun = document.Paragraphs[1];
            paragraphs[1].Text = "Second, edited";
            added = document.AddParagraph("Added");
            Assert.True(document.IsSdkDeferred);
        }

        // PowerShell formats cmdlet output after the cmdlet disposed its document, and using blocks hand paragraphs out.
        // Formatting read first finds the copy the same way text does.
        Assert.True(paragraphs[0].Bold);
        Assert.Equal("First", paragraphs[0].Text);
        Assert.Equal("Second, edited", paragraphs[1].Text);
        Assert.Equal("Added", added.Text);
        Assert.Equal(WordParagraphAlignment.Center, paragraphs[0].ParagraphAlignment);
        Assert.False(paragraphs[0].KeepWithNext);
        Assert.Equal("Second, edited", paragraphs[1].Text);

        // KeepWithNext opened the SDK over the copy; edits through another object for the same run land there too.
        sameRun.Bold = true;
        _ = paragraphs[1].KeepWithNext;
        Assert.True(paragraphs[1].Bold);
    }

    [Fact]
    public void CustomPropertiesSaveThroughTheSdkAsBefore() {
        const string custom = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/custom-properties\" "
            + "xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\"><property fmtid=\"{D5CDD505-2E9C-101B-9397-08002B2CF9AE}\" pid=\"5\" name=\"Ratio\"><vt:r8>1.5</vt:r8></property></Properties>";
        byte[] source = CreatePackage(BodyXml("<w:p><w:r><w:t>Text</w:t></w:r></w:p>"), extraPackageRelationship:
            "<Relationship Id=\"rCustom\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties\" Target=\"docProps/custom.xml\"/>",
            extraParts: new Dictionary<string, string> { ["docProps/custom.xml"] = custom },
            extraContentTypes: "<Override PartName=\"/docProps/custom.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.custom-properties+xml\"/>");
        void Edit(WordDocument document) => document.AddParagraph("Added");
        byte[] sdk = EditAndSave(source, false, Edit)!;
        byte[] owned = EditAndSave(source, true, Edit)!;
        AssertSamePackage(sdk, owned, "custom properties");

        using WordDocument loaded = WordDocument.Load(new MemoryStream(source));
        loaded.AddParagraph("Added");
        loaded.Save(new MemoryStream());
        Assert.Contains("Save: custom document properties", loaded.OpenXmlSdkFallbacks);
    }

    [Fact]
    public void MembersTheOwnedLayerLacksOpenTheSdkAndSayWhy() {
        byte[] source = CreatePackage(BodyXml("<w:p><w:r><w:t>Text</w:t></w:r></w:p>"));
        using WordDocument document = WordDocument.Load(new MemoryStream(source));
        Assert.True(document.IsSdkDeferred);

        WordParagraph paragraph = document.Paragraphs[0];
        paragraph.KeepWithNext = true;

        Assert.False(document.IsSdkDeferred);
        string reason = Assert.Single(document.OpenXmlSdkFallbacks);
        Assert.Contains("KeepWithNext", reason, StringComparison.Ordinal);
        Assert.True(document.Paragraphs[0].KeepWithNext);
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
            List<string> paragraphs = document.Paragraphs.Select(paragraph =>
                $"{paragraph.Text}|{paragraph.Bold}|{paragraph.Italic}|{paragraph.Underline}|{paragraph.Strikethrough}|{paragraph.FontSize}"
                + $"|{paragraph.Color}|{paragraph.FontFamily}|{paragraph.Highlight}|{paragraph.CapsStyle}"
                + $"|{paragraph.StyleId}|{paragraph.Style}|{paragraph.ParagraphAlignment}"
                + $"|{paragraph.IndentationBefore}|{paragraph.IndentationAfter}|{paragraph.IndentationFirstLine}|{paragraph.IndentationHanging}"
                + $"|{paragraph.LineSpacingBefore}|{paragraph.LineSpacingAfter}").ToList();
            _lastRunStayedOwned = document.IsSdkDeferred;
            return paragraphs;
        }
    }

    private static void StandardEdit(WordDocument document) {
        WordParagraph added = document.AddParagraph("Hello world");
        added.Bold = true;
        added.Underline = WordUnderlineStyle.Double;
        added.FontFamily = "Arial";
        added.AddText(" second").Strikethrough = true;
        document.AddParagraph("tab\there\nnext line");
        WordParagraph centered = document.AddParagraph("Centered");
        centered.ParagraphAlignment = WordParagraphAlignment.Center;
        centered.IndentationBefore = OfficeLength.Points(36);
        centered.IndentationFirstLine = OfficeLength.Points(12);
        centered.LineSpacingBefore = OfficeLength.Points(6);
        centered.LineSpacingAfter = OfficeLength.Points(10);
        centered.IndentationAfter = null;
        List<WordParagraph> paragraphs = document.Paragraphs;
        if (paragraphs.Count > 0) {
            paragraphs[0].Italic = true;
            paragraphs[0].Text = paragraphs[0].Text + " edited";
            paragraphs[0].ParagraphAlignment = WordParagraphAlignment.Right;
            paragraphs[0].Style = WordParagraphStyles.Normal;
        }
        if (paragraphs.Count > 1) {
            paragraphs[1].ParagraphAlignment = null;
            paragraphs[1].IndentationHanging = OfficeLength.Points(18);
            paragraphs[1].LineSpacingAfter = null;
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
            _lastRunStayedOwned = document.IsSdkDeferred;
            return output.ToArray();
        }
    }

    private static void AssertSamePackage(byte[] expected, byte[] actual, string name) {
        Dictionary<string, byte[]> expectedParts = ReadParts(expected);
        Dictionary<string, byte[]> actualParts = ReadParts(actual);
        Assert.Equal(expectedParts.Keys.OrderBy(key => key, StringComparer.Ordinal), actualParts.Keys.OrderBy(key => key, StringComparer.Ordinal));
        foreach (KeyValuePair<string, byte[]> part in expectedParts) {
            byte[] other = actualParts[part.Key];
            bool xml = part.Key.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) || part.Key.EndsWith(".rels", StringComparison.OrdinalIgnoreCase);
            Dictionary<string, string> expectedIds = GeneratedRelationshipIds(expectedParts, part.Key);
            Dictionary<string, string> actualIds = GeneratedRelationshipIds(actualParts, part.Key);
            // Identical bytes still differ when the same generated identifiers point elsewhere.
            bool sameIds = expectedIds.Count == actualIds.Count && expectedIds.All(id => actualIds.TryGetValue(id.Key, out string? target) && target == id.Value);
            if (part.Value.AsSpan().SequenceEqual(other) && (sameIds || !xml)) continue;
            if (!xml) Assert.Fail($"{name}: binary part {part.Key} differs.");
            string expectedXml = Canonical(part.Value, expectedIds);
            string actualXml = Canonical(other, actualIds);
            if (expectedXml == actualXml) continue;
            int at = 0;
            while (at < expectedXml.Length && at < actualXml.Length && expectedXml[at] == actualXml[at]) at++;
            int from = Math.Max(0, at - 200);
            Assert.Fail($"{name}: {part.Key} differs at {at}.\nSDK:   {Excerpt(expectedXml)}\nOwned: {Excerpt(actualXml)}");
            string Excerpt(string xml) => xml.Substring(Math.Min(from, xml.Length), Math.Min(500, xml.Length - Math.Min(from, xml.Length)));
        }
    }

    // Element and attribute names with their namespaces, values, and text, in document order; attribute order, prefixes,
    // namespace declarations and whitespace-only text between child elements don't count. Revision identifiers are random
    // on both paths, and so are the relationship identifiers the SDK generates for parts it adds: each of those stands for
    // its relationship's type and target, in the relationships part and in every reference from its source part, so a
    // consistent renaming matches and two identifiers that swap targets don't.
    private static readonly System.Text.RegularExpressions.Regex GeneratedRelationshipId = new("^R[0-9a-f]{16}$");
    private static readonly XNamespace PackageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";

    private static Dictionary<string, string> GeneratedRelationshipIds(Dictionary<string, byte[]> parts, string partName) {
        var ids = new Dictionary<string, string>(StringComparer.Ordinal);
        string relationshipsPart = partName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase) ? partName
            : partName.Substring(0, partName.LastIndexOf('/') + 1) + "_rels/" + partName.Substring(partName.LastIndexOf('/') + 1) + ".rels";
        if (!parts.TryGetValue(relationshipsPart, out byte[]? bytes)) return ids;
        using var stream = new MemoryStream(bytes);
        foreach (XElement relationship in XDocument.Load(stream).Root!.Elements(PackageRelationships + "Relationship")) {
            string id = (string?)relationship.Attribute("Id") ?? string.Empty;
            if (GeneratedRelationshipId.IsMatch(id)) {
                ids[id] = "rel(" + (string?)relationship.Attribute("Type") + " " + (string?)relationship.Attribute("Target") + ")";
            }
        }
        return ids;
    }

    private static string Canonical(byte[] xml, IReadOnlyDictionary<string, string> generatedIds) {
        XDocument document;
        using (var stream = new MemoryStream(xml)) document = XDocument.Load(stream);
        var builder = new System.Text.StringBuilder();
        void Write(XElement element) {
            builder.Append('<').Append(element.Name.ToString());
            foreach (XAttribute attribute in element.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration)
                         .OrderBy(attribute => attribute.Name.ToString(), StringComparer.Ordinal)) {
                bool revision = attribute.Name.Namespace == Wn && attribute.Name.LocalName.StartsWith("rsid", StringComparison.Ordinal);
                string value = revision ? "*"
                    : generatedIds.TryGetValue(attribute.Value, out string? relationship) ? relationship
                    : attribute.Name.Namespace == Mc ? PrefixesAsUris(element, attribute.Value) : attribute.Value;
                builder.Append(' ').Append(attribute.Name.ToString()).Append("=\"").Append(System.Security.SecurityElement.Escape(value)).Append('"');
            }
            builder.Append('>');
            bool layout = element.HasElements;
            foreach (XNode node in element.Nodes()) {
                if (node is XElement child) Write(child);
                else if (node is XText text && (!layout || !string.IsNullOrWhiteSpace(text.Value))) builder.Append(System.Security.SecurityElement.Escape(text.Value));
            }
            builder.Append("</>");
        }
        Write(document.Root!);
        return builder.ToString();
    }

    // Markup compatibility attributes list prefixes; what counts is the namespaces they name.
    private static string PrefixesAsUris(XElement element, string value) => string.Join(" ", value
        .Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
        .Select(token => {
            int colon = token.IndexOf(':');
            string prefix = colon < 0 ? token : token.Substring(0, colon);
            string uri = element.GetNamespaceOfPrefix(prefix)?.NamespaceName ?? prefix;
            return colon < 0 ? uri : uri + token.Substring(colon);
        }));

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

    private static byte[] CreatePackage(string documentXml, string extraPackageRelationship = "",
        Dictionary<string, string>? extraParts = null, string extraContentTypes = "", string documentRelationships = "") {
        using var output = new MemoryStream();
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
            void Add(string name, string content) {
                using var writer = new StreamWriter(archive.CreateEntry(name).Open());
                writer.Write(content);
            }
            Add("[Content_Types].xml", "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                + "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>"
                + extraContentTypes + "</Types>");
            Add("_rels/.rels", "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>"
                + extraPackageRelationship + "</Relationships>");
            Add("word/document.xml", documentXml);
            if (documentRelationships.Length > 0) {
                Add("word/_rels/document.xml.rels", "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                    + documentRelationships + "</Relationships>");
            }
            foreach (KeyValuePair<string, string> part in extraParts ?? new Dictionary<string, string>()) Add(part.Key, part.Value);
        }
        return output.ToArray();
    }
}
#endif
