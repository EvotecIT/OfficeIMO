using System.IO.Compression;
using System.IO.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Experimental;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word;
using Xunit;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, false, false)]
    [InlineData(false, true, false)]
    [InlineData(true, false, false)]
    [InlineData(true, true, false)]
    [InlineData(false, false, true)]
    [InlineData(false, true, true)]
    [InlineData(true, false, true)]
    [InlineData(true, true, true)]
    public void SharedPackageCloning_PublicWorkflowsPreserveGraphAndPendingTableState(
        bool redline, bool sharedFirst, bool authored) {
        using var input = new MemoryStream(CreateSharedCloneInput(sharedFirst));
        using WordDocument target = WordDocument.Load(input);
        target.Paragraphs[0].Text = "Pending target text";
        WordTable table = target.AddTable(1, 2);
        table.ColumnWidth = new List<WordTableWidth?> {
            WordTableWidth.Length(OfficeLength.Points(72)), WordTableWidth.Length(OfficeLength.Points(144))
        };
        if (authored) table.GridColumnWidth = new List<OfficeLength> { OfficeLength.Points(105), OfficeLength.Points(135) };
        // Leave the physical grid stale: the clone must retain whether normalization may derive it.
        table.Rows[0].Cells[0].Width = WordTableWidth.Length(OfficeLength.Points(90));
        table.Rows[0].Cells[1].Width = WordTableWidth.Length(OfficeLength.Points(150));
        WordprocessingDocument live = target.OpenXmlDocument;
        ExtendedFilePropertiesPart lazyProperties = live.ExtendedFilePropertiesPart!;
        Assert.False(lazyProperties.IsRootElementLoaded);
        string before = live.MainDocumentPart!.Document.OuterXml;
        string[] partsBefore = SharedClonePartUris(live);

        if (redline) {
            using WordDocument source = WordDocument.Create();
            source.AddParagraph("Original text");
            string output = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".docx");
            WordComparisonResult comparison = WordDocumentComparer.CreateRedlineDocument(source, target, output,
                new WordComparisonRedlineOptions { Mode = WordComparisonRedlineMode.InPlaceTarget });
            Assert.NotEmpty(comparison.Findings);
            using WordprocessingDocument saved = WordprocessingDocument.Open(output, false);
            AssertSharedData(saved);
            Assert.Equal(partsBefore, SharedClonePartUris(saved));
            Assert.Contains("Pending target text", saved.MainDocumentPart!.Document.InnerText);
            Assert.NotEmpty(saved.MainDocumentPart.Document.Descendants<W.InsertedRun>());
            W.Table savedTable = Assert.Single(saved.MainDocumentPart.Document.Body!.Elements<W.Table>());
            Assert.Equal(authored ? new[] { "2100", "2700" } : new[] { "1800", "3000" },
                savedTable.GetFirstChild<W.TableGrid>()!.Elements<W.GridColumn>().Select(column => column.Width!.Value));
            using var archive = ZipFile.OpenRead(output);
            Assert.Equal(archive.Entries.Count, archive.Entries.Select(entry => entry.FullName.ToUpperInvariant()).Distinct().Count());
            Assert.Single(archive.Entries, entry => entry.FullName.Equals("udata/data.bin", StringComparison.OrdinalIgnoreCase));
            Assert.Single(archive.Entries, entry => entry.FullName.EndsWith(".bin", StringComparison.OrdinalIgnoreCase));
        } else {
            LegacyDocWriteAssessment assessment = target.AssessLegacyDocWrite();
            Assert.True(assessment.IsSupported, assessment.Message);
            Assert.True(assessment.EncodedByteCount > 0);
        }

        Assert.Equal(before, live.MainDocumentPart.Document.OuterXml);
        Assert.False(lazyProperties.IsRootElementLoaded);
        AssertSharedData(live);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void SharedPackageCloning_ReadOnlyInputsRemainUntouched(bool redline, bool sharedFirst) {
        byte[] bytes = CreateSharedCloneInput(sharedFirst);
        string input = Path.Combine(_directoryWithFiles, "target.docx");
        File.WriteAllBytes(input, bytes);
        if (redline) {
            string source = Path.Combine(_directoryWithFiles, "source.docx");
            using (WordDocument document = WordDocument.Create()) {
                document.AddParagraph("Before text");
                document.Save(source);
            }
            string output = Path.Combine(_directoryWithFiles, "redline.docx");
            WordDocumentComparer.CreateRedlineDocument(source, input, output,
                new WordComparisonRedlineOptions { Mode = WordComparisonRedlineMode.InPlaceTarget });
            using WordprocessingDocument saved = WordprocessingDocument.Open(output, false);
            AssertSharedData(saved);
            Assert.NotEmpty(saved.MainDocumentPart!.Document.Descendants<W.InsertedRun>());
        } else {
            using WordDocument document = WordDocument.Load(input, new WordLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });
            Assert.True(document.AssessLegacyDocWrite().IsSupported);
            AssertSharedData(document.OpenXmlDocument);
        }
        Assert.Equal(bytes, File.ReadAllBytes(input));
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void SharedPackageCloning_StrictInputsKeepLazyXmlReadable(bool redline, bool sharedFirst) {
        using var input = new MemoryStream(CreateStrictSharedCloneInput(sharedFirst));
        using WordDocument target = WordDocument.Load(input);
        WordprocessingDocument live = target.OpenXmlDocument;
        Assert.True(live.StrictRelationshipFound);
        Assert.False(live.ExtendedFilePropertiesPart!.IsRootElementLoaded);
        if (redline) {
            using WordDocument source = WordDocument.Create();
            source.AddParagraph("Before text");
            string output = Path.Combine(_directoryWithFiles, "strict-redline.docx");
            WordDocumentComparer.CreateRedlineDocument(source, target, output,
                new WordComparisonRedlineOptions { Mode = WordComparisonRedlineMode.InPlaceTarget });
            using WordprocessingDocument saved = WordprocessingDocument.Open(output, false);
            AssertSharedData(saved);
            Assert.NotNull(saved.ExtendedFilePropertiesPart!.Properties);
            Assert.NotEmpty(saved.MainDocumentPart!.Document.Descendants<W.InsertedRun>());
        } else {
            LegacyDocWriteAssessment result = target.AssessLegacyDocWrite();
            Assert.True(result.IsSupported, result.Message);
        }
        Assert.False(live.ExtendedFilePropertiesPart.IsRootElementLoaded);
        AssertSharedData(live);
    }

    private static byte[] CreateStrictSharedCloneInput(bool sharedFirst, int propertiesPadding = 0) {
        using var input = new MemoryStream();
        byte[] original = CreateSharedCloneInput(sharedFirst);
        input.Write(original, 0, original.Length);
        using (var archive = new ZipArchive(input, ZipArchiveMode.Update, leaveOpen: true)) {
            foreach (ZipArchiveEntry entry in archive.Entries.ToArray()) {
                if (!entry.FullName.EndsWith(".xml", StringComparison.Ordinal) &&
                    !entry.FullName.EndsWith(".rels", StringComparison.Ordinal)) continue;
                string xml;
                using (var reader = new StreamReader(entry.Open())) xml = reader.ReadToEnd();
                if (entry.FullName == "docProps/app.xml" && propertiesPadding > 0) {
                    xml = xml.Replace("</Properties>", new string(' ', propertiesPadding) + "</Properties>");
                }
                xml = xml.Replace("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "http://purl.oclc.org/ooxml/wordprocessingml/main")
                    .Replace("http://schemas.openxmlformats.org/officeDocument/2006/extended-properties", "http://purl.oclc.org/ooxml/officeDocument/extendedProperties")
                    .Replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "http://purl.oclc.org/ooxml/officeDocument/relationships")
                    .Replace("http://purl.oclc.org/ooxml/officeDocument/relationships/extended-properties", "http://purl.oclc.org/ooxml/officeDocument/relationships/extendedProperties");
                using Stream output = entry.Open();
                output.SetLength(0);
                using var writer = new StreamWriter(output);
                writer.Write(xml);
            }
        }
        return input.ToArray();
    }

    private static byte[] CreateSharedCloneInput(bool sharedFirst) {
        using var input = new MemoryStream();
        using (WordprocessingDocument package = WordprocessingDocument.Create(input, WordprocessingDocumentType.Document)) {
            package.PackageProperties.Creator = "Shared custom-data fixture";
            ExtendedPart? shared = sharedFirst ? AddSharedData(package) : null;
            MainDocumentPart main = package.AddMainDocumentPart();
            main.Document = new W.Document(new W.Body(new W.Paragraph(new W.Run(new W.Text("Original text")))));
            main.AddNewPart<DocumentSettingsPart>("rIdSettings").Settings = new W.Settings();
            shared ??= AddSharedData(package);
            main.AddPart(shared, "rIdSharedChild");
            package.AddExternalRelationship("urn:officeimo:test:external", new Uri("https://example.invalid/inert"), "rIdExternalRoot");
            main.AddHyperlinkRelationship(new Uri("https://example.invalid/link"), true, "rIdExternalChild");
            ExtendedFilePropertiesPart properties = package.AddNewPart<ExtendedFilePropertiesPart>("rIdLazyProperties");
            using var propertiesBytes = new MemoryStream(Encoding.UTF8.GetBytes(
                "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\"><Application>Inert custom-data fixture</Application></Properties>"));
            properties.FeedData(propertiesBytes);
        }
        return input.ToArray();
    }

    private static ExtendedPart AddSharedData(WordprocessingDocument package) {
        ExtendedPart shared = package.AddExtendedPart("urn:officeimo:test:shared", "application/octet-stream", ".bin", "rIdSharedRoot");
        using var bytes = new MemoryStream(new byte[] { 0, 1, 2, 3, 255 });
        shared.FeedData(bytes);
        return shared;
    }

#pragma warning disable OOXML0001
    private static string[] SharedClonePartUris(WordprocessingDocument package) => package.GetPackage().GetParts()
        .Where(part => !PackUriHelper.IsRelationshipPartUri(part.Uri))
        .Select(part => part.Uri.ToString().ToUpperInvariant()).OrderBy(uri => uri, StringComparer.Ordinal).ToArray();
#pragma warning restore OOXML0001

    private static void AssertSharedData(WordprocessingDocument package) {
        OpenXmlPart root = package.GetPartById("rIdSharedRoot");
        Assert.Same(root, package.MainDocumentPart!.GetPartById("rIdSharedChild"));
        Assert.Equal("/udata/data.bin", root.Uri.ToString(), ignoreCase: true);
        using Stream content = root.GetStream();
        using var copy = new MemoryStream();
        content.CopyTo(copy);
        Assert.Equal(new byte[] { 0, 1, 2, 3, 255 }, copy.ToArray());
        Assert.Single(package.Parts, pair => pair.OpenXmlPart is ExtendedPart);
        Assert.Single(package.MainDocumentPart.Parts, pair => pair.OpenXmlPart is ExtendedPart);
        ExternalRelationship external = Assert.Single(package.ExternalRelationships);
        Assert.Equal("rIdExternalRoot", external.Id);
        Assert.Equal("urn:officeimo:test:external", external.RelationshipType);
        Assert.Equal("https://example.invalid/inert", external.Uri.ToString());
        HyperlinkRelationship link = Assert.Single(package.MainDocumentPart.HyperlinkRelationships);
        Assert.Equal("rIdExternalChild", link.Id);
        Assert.Equal("https://example.invalid/link", link.Uri.ToString());
    }
}
