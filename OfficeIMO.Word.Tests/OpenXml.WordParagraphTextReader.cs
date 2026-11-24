using System.IO.Compression;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Core.Internal.OpenXml;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>
/// The SDK-free Word paragraph text read must return what the Open XML SDK returns for the body paragraphs'
/// InnerText, with markup compatibility processed as an Office 2007 consumer.
/// </summary>
public class OpenXmlWordParagraphTextReaderTests {
    private const string WordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    public static IEnumerable<object[]> CorpusDocuments() {
        string root = AppContext.BaseDirectory;
        foreach (string folder in new[] { "Documents", "WordLayoutFixtures" }) {
            string path = Path.Combine(root, folder);
            if (!Directory.Exists(path)) continue;
            foreach (string file in Directory.EnumerateFiles(path, "*.docx", SearchOption.AllDirectories).OrderBy(file => file, StringComparer.Ordinal)) {
                yield return new object[] { file.Substring(root.Length).TrimStart(Path.DirectorySeparatorChar) };
            }
        }
    }

    [Theory]
    [MemberData(nameof(CorpusDocuments))]
    public void ReadsTheSameParagraphTextAsTheSdkForCorpusDocuments(string relativePath) {
        byte[] package = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, relativePath));
        List<string>? expected = TryReadWithSdk(package);
        if (expected == null) return;

        Assert.Equal(expected, ReadWithPackageReader(package));
    }

    [Fact]
    public void ReadsTheBenchmarkParagraphFixtureLikeTheSdk() {
        byte[] package = CreateSdkParagraphFixture(2_000);

        List<string> actual = ReadWithPackageReader(package);

        Assert.Equal(2_000, actual.Count);
        Assert.Equal(TryReadWithSdk(package), actual);
    }

    [Theory]
    [InlineData("entities", "<w:p><w:r><w:t>a &amp; b &lt;c&gt; &quot;d&quot; &apos;e&apos; &#169; &#x1F600;</w:t></w:r></w:p>")]
    [InlineData("strict and transitional attribute names", "<w:p xmlns:t=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:s=\"http://purl.oclc.org/ooxml/officeDocument/relationships\" t:id=\"1\" s:id=\"2\"><w:r><w:t>both</w:t></w:r></w:p>")]
    [InlineData("zero-padded references", "<w:p><w:r><w:t>&#0000000000000065;&#x0000000000000041;</w:t></w:r></w:p>")]
    [InlineData("line ends", "<w:p><w:r><w:t xml:space=\"preserve\">one\r\ntwo\rthree</w:t></w:r></w:p>")]
    [InlineData("whitespace", "<w:p>\n  <w:r>\n    <w:rPr>\n    </w:rPr>\n    <w:tab>  </w:tab>\n    <w:t>   </w:t></w:r><w:r><w:t xml:space=\"preserve\">   </w:t></w:r><w:r><w:t> x </w:t></w:r></w:p>")]
    [InlineData("cdata and comments", "<w:p><w:r><w:t>a<!-- note --> <![CDATA[<b>]]> c<?pi data?></w:t></w:r></w:p>")]
    [InlineData("deleted and field text", "<w:p><w:del w:id=\"1\" w:author=\"A\"><w:r><w:delText>gone</w:delText></w:r></w:del><w:r><w:instrText xml:space=\"preserve\"> PAGE </w:instrText></w:r><w:r><w:tab/><w:t>kept</w:t><w:br/></w:r></w:p>")]
    [InlineData("per-element prefixes", "<w:p><x:r xmlns:x=\"" + WordNamespace + "\"><x:t>declared here</x:t></x:r><r xmlns=\"" + WordNamespace + "\"><t>default namespace</t></r></w:p>")]
    [InlineData("nested content", "<w:p><w:hyperlink w:anchor=\"a\"><w:r><w:t>link</w:t></w:r></w:hyperlink><w:sdt><w:sdtContent><w:r><w:t>control</w:t></w:r></w:sdtContent></w:sdt></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>")]
    [InlineData("alternate content", "<w:p><w:r><mc:AlternateContent><mc:Choice xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\" Requires=\"wps\"><w:t>choice</w:t></mc:Choice><mc:Fallback><w:t>fallback</w:t></mc:Fallback></mc:AlternateContent></w:r></w:p>")]
    [InlineData("understood choice", "<w:p><w:r><mc:AlternateContent><mc:Choice xmlns:m2=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" Requires=\"m2\"><w:t>first</w:t></mc:Choice><mc:Choice Requires=\"w\"><w:t>second</w:t></mc:Choice><mc:Fallback><w:t>fallback</w:t></mc:Fallback></mc:AlternateContent></w:r></w:p>")]
    [InlineData("body-level alternate content", "<mc:AlternateContent><mc:Choice xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" Requires=\"w14\"><w:p><w:r><w:t>new</w:t></w:r></w:p></mc:Choice><mc:Fallback><w:p><w:r><w:t>old</w:t></w:r></w:p></mc:Fallback></mc:AlternateContent>")]
    [InlineData("ignorable content", "<w:p xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:x9=\"urn:example:x9\" mc:Ignorable=\"w14 x9\" mc:ProcessContent=\"x9:wrap\" w14:paraId=\"1\"><w14:dropped><w:r><w:t>dropped</w:t></w:r></w14:dropped><x9:wrap><w:r><w:t>kept</w:t></w:r></x9:wrap></w:p>")]
    public void ReadsMarkupEdgeCasesLikeTheSdk(string name, string body) {
        byte[] package = CreatePackage(body);

        List<string> expected = TryReadWithSdk(package) ?? throw new InvalidOperationException("The SDK did not open the '" + name + "' case.");
        Assert.Equal(expected, ReadWithPackageReader(package));
    }

    [Fact]
    public void ReadsStrictDocumentsAsTheirTransitionalEquivalent() {
        string body = "<w:p><w:r><w:t>strict</w:t></w:r><w:r><mc:AlternateContent><mc:Choice Requires=\"w\"><w:t> branch</w:t></mc:Choice><mc:Fallback><w:t> other</w:t></mc:Fallback></mc:AlternateContent></w:r></w:p>";
        byte[] transitional = CreatePackage(body);
        byte[] strict = CreatePackage(body, strict: true);

        Assert.Equal(TryReadWithSdk(transitional), ReadWithPackageReader(strict));
    }

    [Fact]
    public void ReadsUtf16Parts() {
        string body = "<w:p><w:r><w:t>UTF-16 é中</w:t></w:r></w:p>";
        byte[] package = CreatePackage(body, encoding: new UnicodeEncoding(bigEndian: false, byteOrderMark: true));

        Assert.Equal(new[] { "UTF-16 é中" }, ReadWithPackageReader(package));
    }

    [Theory]
    [InlineData("<!DOCTYPE w:document []>")]
    [InlineData("<w:p><w:r><w:t>&custom;</w:t></w:r></w:p>")]
    [InlineData("<w:p><w:r><w:t>&#0;</w:t></w:r></w:p>")]
    [InlineData("<w:p><w:r><w:t>open</w:r></w:p>")]
    [InlineData("<w:p><q:r><q:t>unbound</q:t></q:r></w:p>")]
    [InlineData("<w:p xmlns:x9=\"urn:example:x9\" mc:MustUnderstand=\"x9\"><w:r><w:t>required</w:t></w:r></w:p>")]
    [InlineData("<w:p w:rsidR=\"1\" w:rsidR=\"2\"/>")]
    [InlineData("<w:p xmlns:w2=\"" + WordNamespace + "\" w:rsidR=\"1\" w2:rsidR=\"2\"/>")]
    [InlineData("SECOND-ROOT")]
    [InlineData("<w:p><w:r><w:t>a\u0001b</w:t></w:r></w:p>")]
    [InlineData("<w:p w14:x=\"\uFFFE\" xmlns:w14=\"urn:example\"/>")]
    public void RejectsMarkupItCannotReadSafely(string body) {
        byte[] package = body.StartsWith("<!DOCTYPE", StringComparison.Ordinal)
            ? CreatePackage("<w:p/>", prolog: body)
            : body == "SECOND-ROOT"
                ? CreatePackage("<w:p><w:r><w:t>first</w:t></w:r></w:p>", epilog: "<w:document xmlns:w=\"" + WordNamespace + "\"><w:body><w:p/></w:body></w:document>")
                : CreatePackage(body);

        // Malformed XML is an XmlException; a namespace the markup requires but the reader does not know is not a syntax error.
        Type expected = body.Contains("MustUnderstand") ? typeof(InvalidDataException) : typeof(System.Xml.XmlException);
        Assert.IsType(expected, Record.Exception(() => ReadWithPackageReader(package)));
    }

    [Fact]
    public void RequiresMustUnderstandNamespacesOnlyOnTheSelectedBranch() {
        byte[] package = CreatePackage("<w:p><w:r><mc:AlternateContent><mc:Choice xmlns:x9=\"urn:example:x9\" Requires=\"x9\" mc:MustUnderstand=\"x9\"><w:t>choice</w:t></mc:Choice><mc:Fallback><w:t>fallback</w:t></mc:Fallback></mc:AlternateContent></w:r></w:p>");

        Assert.Equal(new[] { "fallback" }, ReadWithPackageReader(package));
    }

    [Theory]
    [InlineData(new byte[] { 0x3C, 0x61, 0x2F, 0x3E }, 4)]
    [InlineData(new byte[] { 0xEF, 0xBB, 0xBF, 0x3C, 0x61, 0x2F, 0x3E }, 4)]
    [InlineData(new byte[] { 0xFF, 0xFE, 0x3C, 0x00, 0x61, 0x00, 0x2F, 0x00, 0x3E, 0x00 }, 4)]
    [InlineData(new byte[] { 0x3C, 0xC3, 0xA9, 0xF0, 0x9F, 0x98, 0x80, 0x3E }, 5)]
    public void CountsXmlCharactersInTheirEncoding(byte[] part, long expected) {
        Assert.Equal(expected, OfficeXmlPullReader.CountCharacters(part, part.Length));
    }

    [Fact]
    public void RejectsMalformedUtf8OutsideDecodedText() {
        byte[] part = Encoding.ASCII.GetBytes("<a><!-- x --></a>");
        part[8] = 0x80;

        Assert.Throws<System.Xml.XmlException>(() => new OfficeXmlPullReader(part, part.Length));
    }

    [Fact]
    public void MatchesNoteReferencesToNotesByNumber() {
        const string footnotes = "<w:footnotes xmlns:w=\"" + WordNamespace + "\"><w:footnote w:id=\"1\"><w:p><w:r><w:t>Numbered note</w:t></w:r></w:p></w:footnote></w:footnotes>";
        byte[] package = CreatePackage("<w:p><w:r><w:t>Body</w:t></w:r><w:r><w:footnoteReference w:id=\"01\"/></w:r></w:p>",
            extraPart: ("word/footnotes.xml", footnotes, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"));
        using var stream = new MemoryStream(package, writable: false);

        List<string>? lines = OfficeIMO.Word.WordStreamedBodyText.TryReadLines(stream, 8000, 0, null, includeNotes: true, CancellationToken.None, out string? reason);

        Assert.Equal(new[] { "Body", "Numbered note" }, lines);
        Assert.Null(reason);
    }

    [Fact]
    public void DeclinesAPartWhoseDeclaredSizeIsOverTheLimitBeforeInflatingIt() {
        // The member claims more bytes than it holds; reading it would fail its length check, so a decline shows the
        // limit was applied to the declared size first.
        byte[] package = DeclareUncompressedSize(CreatePackage("<w:p><w:r><w:t>small</w:t></w:r></w:p>"),
            "word/document.xml", 10_000);
        using var stream = new MemoryStream(package, writable: false);

        List<string>? lines = OfficeIMO.Word.WordStreamedBodyText.TryReadLines(stream, 8000, 0, 1_000, includeNotes: true, CancellationToken.None, out string? reason);

        Assert.Null(lines);
        Assert.Equal("part over the character limit", reason);
    }

    [Theory]
    [InlineData(new byte[] { 0xFF, 0xFE, 0x3C, 0x00, 0x61, 0x00, 0x00, 0xD8, 0x2F, 0x00, 0x3E, 0x00 })]
    [InlineData(new byte[] { 0xFF, 0xFE, 0x3C, 0x00, 0x61, 0x00, 0x2F, 0x00, 0x3E })]
    public void RejectsMalformedUtf16(byte[] part) {
        Assert.Throws<System.Xml.XmlException>(() => new OfficeXmlPullReader(part, part.Length));
    }

    [Fact]
    public void DeclinesWhenAComplexFieldSeparatesInsideASkippedContentControl() {
        // The field's result in the last paragraph shows only if its separate marker, inside the control, is seen.
        byte[] package = CreatePackage(
            "<w:p><w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText>PAGE</w:instrText></w:r></w:p>"
            + "<w:sdt><w:sdtContent><w:p><w:r><w:fldChar w:fldCharType=\"separate\"/></w:r></w:p></w:sdtContent></w:sdt>"
            + "<w:p><w:r><w:t>Result</w:t></w:r><w:r><w:fldChar w:fldCharType=\"end\"/></w:r></w:p>");
        using var stream = new MemoryStream(package, writable: false);

        List<string>? lines = OfficeIMO.Word.WordStreamedBodyText.TryReadLines(stream, 8000, 0, null, includeNotes: true, CancellationToken.None, out string? reason);

        Assert.Null(lines);
        Assert.Equal("complex field marker in a skipped block", reason);
    }

    [Theory]
    [InlineData("")]
    [InlineData("<Types xmlns=\"urn:other\"/>")]
    public void RejectsAPackageWithoutAReadableContentTypesPart(string contentTypes) {
        byte[] package = CreatePackage("<w:p><w:r><w:t>Body</w:t></w:r></w:p>", contentTypesXml: contentTypes);

        Assert.Throws<InvalidDataException>(() => OfficeOpcPackage.Open(new MemoryStream(package, writable: false)));
    }

    [Fact]
    public void RejectsAMainPartWithAnotherContentType() {
        byte[] package = CreatePackage("<w:p><w:r><w:t>Body</w:t></w:r></w:p>",
            contentTypesXml: "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                + "<Default Extension=\"xml\" ContentType=\"application/xml\"/></Types>");
        using var stream = new MemoryStream(package, writable: false);

        Assert.Throws<InvalidDataException>(() => OfficeIMO.Word.WordStreamedBodyText.TryReadLines(stream, 8000, 0, null,
            includeNotes: true, CancellationToken.None, out _));
    }

    [Fact]
    public void FindsContentTypeOverridesSpelledWithDifferentPercentEncoding() {
        byte[] package = CreatePackage("<w:p/>", documentPath: "word/pr%C3%A9sentation.xml", documentTarget: "word/pr\u00E9sentation.xml");
        using OfficeOpcPackage opened = OfficeOpcPackage.Open(new MemoryStream(package, writable: false));

        Assert.Equal("application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml", opened.GetContentType("/word/pr\u00E9sentation.xml"));
    }

    [Fact]
    public void KeepsDistinctIdsForManyDeclaredNamespaces() {
        var table = new OfficeXmlNamespaces.Table();
        var ids = new int[2_000];
        for (int index = 0; index < ids.Length; index++) ids[index] = table.GetId("urn:custom:" + index.ToString("D5", System.Globalization.CultureInfo.InvariantCulture));

        Assert.Equal(ids.Length, ids.Distinct().Count());
        for (int index = 0; index < ids.Length; index++) {
            string uri = "urn:custom:" + index.ToString("D5", System.Globalization.CultureInfo.InvariantCulture);
            Assert.Equal(ids[index], table.GetId(uri));
            Assert.Equal(uri, table.GetUri(ids[index]));
        }
    }

    [Fact]
    public void RejectsPartsWithoutARootElement() {
        Assert.Throws<System.Xml.XmlException>(() => ReadWithPackageReader(CreatePackage(null, documentXml: "<?xml version=\"1.0\"?>\n  ")));
    }

    [Fact]
    public void RejectsRepeatedRelationshipIds() {
        byte[] package = CreatePackage("<w:p/>", extraRelationship: "<Relationship Id=\"rId1\" Type=\"urn:example\" Target=\"other.xml\"/>");

        Assert.Throws<InvalidDataException>(() => ReadWithPackageReader(package));
    }

    [Fact]
    public void ResolvesNonAsciiTargetsToPercentEncodedMembers() {
        byte[] package = CreatePackage("<w:p><w:r><w:t>encoded</w:t></w:r></w:p>", documentPath: "word/pr%C3%A9sentation.xml", documentTarget: "word/pr\u00E9sentation.xml");

        Assert.Equal(new[] { "encoded" }, ReadWithPackageReader(package));
    }

    [Fact]
    public void RejectsMembersThatOnlyDifferInCase() {
        byte[] package = CreatePackage("<w:p/>", extraMember: "WORD/document.xml");

        Assert.Throws<InvalidDataException>(() => ReadWithPackageReader(package));
    }

    [Fact]
    public void RejectsMembersThatFailTheirChecksum() {
        byte[] package = CreatePackage("<w:p><w:r><w:t>checksum</w:t></w:r></w:p>", compression: CompressionLevel.NoCompression);
        int text = IndexOf(package, Encoding.ASCII.GetBytes("checksum"));
        package[text] = (byte)'C';

        Assert.Throws<InvalidDataException>(() => ReadWithPackageReader(package));
    }

    [Fact]
    public void ResolvesPercentEncodedAndRelativeTargets() {
        byte[] package = CreatePackage("<w:p><w:r><w:t>moved</w:t></w:r></w:p>", documentPath: "word/main body.xml", documentTarget: "./word/main%20body.xml");

        Assert.Equal(new[] { "moved" }, ReadWithPackageReader(package));
    }

    [Theory]
    [InlineData("/word/_rels/document.xml.rels", "media/image1.png", "/word/media/image1.png")]
    [InlineData("/word/_rels/document.xml.rels", "./styles.xml", "/word/styles.xml")]
    [InlineData("/word/_rels/document.xml.rels", "../customXml/item1.xml", "/customXml/item1.xml")]
    [InlineData("/word/_rels/document.xml.rels", "/docProps/app.xml", "/docProps/app.xml")]
    [InlineData("/word/_rels/document.xml.rels", "media/a%20b.png", "/word/media/a%20b.png")]
    [InlineData("/word/_rels/document.xml.rels", "media\\image1.png#frag", "/word/media/image1.png")]
    [InlineData("/_rels/.rels", "word/document.xml", "/word/document.xml")]
    [InlineData("/_rels/.rels", "../../word/document.xml", null)]
    [InlineData("/word/_rels/document.xml.rels", "https://example.com/a.png", null)]
    [InlineData("/word/_rels/document.xml.rels", "C:\\a.png", null)]
    [InlineData("/word/_rels/document.xml.rels", "//server/share/a.png", null)]
    [InlineData("/word/_rels/document.xml.rels", "media/", null)]
    [InlineData("/word/_rels/document.xml.rels", "media/%2E%2E/a.png", null)]
    public void ResolvesRelationshipTargetsToPartNames(string relationshipsPart, string target, string? expected) {
        Assert.Equal(expected, OfficeOpcPartNames.ResolveRelationshipTarget(relationshipsPart, target));
    }

    private static List<string> ReadWithPackageReader(byte[] package) {
        using var stream = new MemoryStream(package, writable: false);
        return OfficeOpenXmlTextReader.ReadWordBodyParagraphs(stream).ToList();
    }

    private static List<string>? TryReadWithSdk(byte[] package) {
        var settings = new OpenSettings {
            MarkupCompatibilityProcessSettings = new MarkupCompatibilityProcessSettings(
                MarkupCompatibilityProcessMode.ProcessAllParts, FileFormatVersions.Office2007)
        };
        try {
            using var stream = new MemoryStream(package, writable: false);
            using WordprocessingDocument document = WordprocessingDocument.Open(stream, false, settings);
            Body? body = document.MainDocumentPart?.Document?.Body;
            return body?.Elements<Paragraph>().Select(paragraph => paragraph.InnerText).ToList();
        } catch (InvalidMCContentException) {
            // Undeclared prefixes in mc:Ignorable stop the SDK's compatibility processing; it reads them unprocessed.
            using var stream = new MemoryStream(package, writable: false);
            using WordprocessingDocument document = WordprocessingDocument.Open(stream, false);
            return document.MainDocumentPart?.Document?.Body?.Elements<Paragraph>().Select(paragraph => paragraph.InnerText).ToList();
        } catch (Exception exception) when (exception is OpenXmlPackageException || exception is InvalidDataException
            || exception is FormatException || exception is IOException || exception is NotSupportedException) {
            // Encrypted, legacy or deliberately broken corpus files are outside this comparison.
            return null;
        }
    }

    private static byte[] CreateSdkParagraphFixture(int count) {
        using var stream = new MemoryStream();
        using (WordprocessingDocument document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document)) {
            MainDocumentPart main = document.AddMainDocumentPart();
            var body = new Body();
            for (int index = 0; index < count; index++) {
                body.Append(new Paragraph(new Run(new Text("Record " + index.ToString("D6", System.Globalization.CultureInfo.InvariantCulture) + " & <owner>"))));
            }
            main.Document = new Document(body);
        }
        return stream.ToArray();
    }

    // Rewrites the uncompressed size a stored member declares in its local header and the central directory.
    private static byte[] DeclareUncompressedSize(byte[] package, string member, uint size) {
        byte[] patched = (byte[])package.Clone();
        byte[] name = Encoding.UTF8.GetBytes(member);
        for (int offset = 0; offset + 46 < patched.Length; offset++) {
            uint signature = BitConverter.ToUInt32(patched, offset);
            if (signature == 0x04034B50 && Matches(patched, offset + 30, BitConverter.ToUInt16(patched, offset + 26), name)) {
                BitConverter.GetBytes(size).CopyTo(patched, offset + 22);
            } else if (signature == 0x02014B50 && Matches(patched, offset + 46, BitConverter.ToUInt16(patched, offset + 28), name)) {
                BitConverter.GetBytes(size).CopyTo(patched, offset + 24);
            }
        }
        return patched;

        static bool Matches(byte[] data, int start, int length, byte[] expected) =>
            length == expected.Length && start + length <= data.Length && data.AsSpan(start, length).SequenceEqual(expected);
    }

    private static byte[] CreatePackage(string? body, bool strict = false, Encoding? encoding = null, string prolog = "",
        string? extraMember = null, CompressionLevel compression = CompressionLevel.Optimal,
        string documentPath = "word/document.xml", string? documentTarget = null, string epilog = "",
        string? documentXml = null, string extraRelationship = "", (string Path, string Xml, string RelationshipType)? extraPart = null,
        string? contentTypesXml = null) {
        string wordNamespace = strict ? "http://purl.oclc.org/ooxml/wordprocessingml/main" : WordNamespace;
        string relationshipType = strict
            ? "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument"
            : "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
        encoding ??= new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
        string declaration = "<?xml version=\"1.0\" encoding=\"" + (encoding is UnicodeEncoding ? "UTF-16" : "UTF-8") + "\" standalone=\"yes\"?>";
        string document = declaration + prolog
            + "<w:document xmlns:w=\"" + wordNamespace + "\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"><w:body>"
            + body + "</w:body></w:document>" + epilog;
        if (documentXml != null) document = documentXml;
        string partName = "/" + documentPath.Replace(" ", "%20").Replace("\u00E9", "%C3%A9");
        string contentTypes = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
            + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
            + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
            + "<Override PartName=\"" + partName + "\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/></Types>";
        if (contentTypesXml != null) contentTypes = contentTypesXml;
        string relationships = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
            + "<Relationship Id=\"rId1\" Type=\"" + relationshipType + "\" Target=\"" + (documentTarget ?? documentPath) + "\"/>" + extraRelationship + "</Relationships>";

        using var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true)) {
            if (contentTypes.Length > 0) Write(archive, "[Content_Types].xml", Encoding.UTF8.GetBytes(contentTypes), compression);
            Write(archive, "_rels/.rels", Encoding.UTF8.GetBytes(relationships), compression);
            byte[] documentBytes = encoding.GetPreamble().Concat(encoding.GetBytes(document)).ToArray();
            Write(archive, documentPath, documentBytes, compression);
            if (extraMember != null) Write(archive, extraMember, documentBytes, compression);
            if (extraPart is { } part) {
                Write(archive, part.Path, Encoding.UTF8.GetBytes(part.Xml), compression);
                string documentRelationships = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                    + "<Relationship Id=\"rIdExtra\" Type=\"" + part.RelationshipType + "\" Target=\"" + Path.GetFileName(part.Path) + "\"/></Relationships>";
                Write(archive, Path.GetDirectoryName(documentPath)!.Replace('\\', '/') + "/_rels/" + Path.GetFileName(documentPath) + ".rels",
                    Encoding.UTF8.GetBytes(documentRelationships), compression);
            }
        }
        return stream.ToArray();
    }

    private static void Write(ZipArchive archive, string name, byte[] content, CompressionLevel compression) {
        using Stream entry = archive.CreateEntry(name, compression).Open();
        entry.Write(content, 0, content.Length);
    }

    private static int IndexOf(byte[] data, byte[] value) {
        for (int index = 0; index <= data.Length - value.Length; index++) {
            int offset = 0;
            while (offset < value.Length && data[index + offset] == value[offset]) offset++;
            if (offset == value.Length) return index;
        }
        throw new InvalidOperationException("The value was not found.");
    }
}
