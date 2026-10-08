using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Xps;
using ReaderOptions = OfficeIMO.Reader.ReaderOptions;
using Xunit;
using static OfficeIMO.Xps.Tests.XpsLogicalStructureTests;

namespace OfficeIMO.Xps.Tests;

public sealed class XpsReaderTests {
    [Theory]
    [InlineData("<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Type='http://schemas.microsoft.com/xps/2005/06/fixedrepresentation' Target='https://example.test/input' TargetMode='External'/></Relationships>")]
    [InlineData("<!DOCTYPE Relationships [<!ENTITY text 'untrusted'>]><Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Type='http://schemas.microsoft.com/xps/2005/06/fixedrepresentation' Target='&text;'/></Relationships>")]
    [InlineData("<Relationships xmlns='urn:foreign'><Relationship Type='http://schemas.microsoft.com/xps/2005/06/fixedrepresentation' Target='sequence.fdseq'/></Relationships>")]
    [InlineData("<Relationships")]
    public async Task ContentDetectionDoesNotPromoteUnsafeOrMalformedRelationships(string xml) {
        using var output = new MemoryStream();
        using (var zip = new ZipArchive(output, ZipArchiveMode.Create, true)) {
            using var writer = new StreamWriter(zip.CreateEntry("_rels/.rels").Open());
            writer.Write(xml);
        }
        byte[] bytes = output.ToArray();
        var reader = new OfficeDocumentReaderBuilder().AddXpsHandler().Build();
        Assert.Equal(ReaderInputKind.Zip, reader.Detect(bytes, "unknown.bin").Kind);
        Assert.Equal(ReaderInputKind.Zip, (await reader.DetectAsync(bytes, "unknown.bin")).Kind);
    }

    [Fact]
    public void InterleavedPackageUsesTheSameNativeReader() {
        var document = Create(XpsFormat.OpenXps, new[] { "Piece text" });
        var reader = new OfficeDocumentReaderBuilder().AddXpsHandler().Build();
        var result = reader.ReadDocument(XpsInterleavingTests.Interleave(document.Save()), "input.oxps");
        Assert.Equal("Piece text", Assert.Single(result.Chunks).Text);
    }

    [Theory]
    [InlineData(XpsFormat.Xps)]
    [InlineData(XpsFormat.OpenXps)]
    public async Task ContentDetectionUsesTheFixedRepresentationRelationship(XpsFormat format) {
        byte[] bytes = Create(format, new[] { "Text" }).Save();
        var reader = new OfficeDocumentReaderBuilder().AddXpsHandler().Build();
        Assert.Equal(ReaderInputKind.Xps, reader.Detect(bytes, "mislabeled.zip").Kind);
        using var stream = new MemoryStream(bytes, false);
        var result = await reader.ReadDocumentAsync(stream, "unknown.bin", new ReaderOptions { DetectionMode = ReaderDetectionMode.PreferContent });
        Assert.Equal(ReaderInputKind.Xps, result.Kind); Assert.Equal("Text", Assert.Single(result.Chunks).Text);
    }

    [Fact]
    public void LegacyNestedSchemaIsRetainedAndCannotHideAnXpsChild() {
        var parent = new OfficeDocumentReadResult { SchemaVersion = 9, Kind = ReaderInputKind.Zip,
            NestedDocuments = new[] { new OfficeDocumentNestedResult { Path = "child", Document = new OfficeDocumentReadResult { Kind = ReaderInputKind.Text } } } };
        string serialized = OfficeDocumentReadResultJson.Serialize(parent);
        using var json = System.Text.Json.JsonDocument.Parse(serialized);
        Assert.Equal(9, json.RootElement.GetProperty("schemaVersion").GetInt32());
        Assert.Equal(9, json.RootElement.GetProperty("nestedDocuments")[0].GetProperty("document").GetProperty("schemaVersion").GetInt32());
        var restored = OfficeDocumentReadResultJson.Deserialize(serialized);
        Assert.Equal(OfficeDocumentReadResultSchema.CurrentVersion, restored.SchemaVersion);
        parent.NestedDocuments[0].Document.Kind = ReaderInputKind.Xps;
        Assert.Throws<System.Text.Json.JsonException>(() => OfficeDocumentReadResultJson.Serialize(parent));
    }

    [Fact]
    public void NativeLinksRetainExternalUrisAndInternalPageIdentity() {
        var doc = Create(XpsFormat.OpenXps, new[] { "External", "Internal" }, new[] { "Target" });
        var page = doc.Pages[0]; var xml = page.GetMarkup();
        xml.Elements().First().SetAttributeValue("FixedPage.NavigateUri", "https://example.test/native");
        xml.Elements().Last().SetAttributeValue("FixedPage.NavigateUri", "/" + doc.Pages[1].PartName + "#Target"); page.ReplaceMarkup(xml);
        var result = doc.ToOfficeDocumentReadResult();
        Assert.Equal(2, result.Links.Count); Assert.Equal("https://example.test/native", result.Links[0].Uri);
        Assert.Equal(2, result.Links[1].DestinationPageNumber); Assert.Equal("Target", result.Links[1].DestinationName);
        Assert.Equal(1, result.Links[1].Location.Page); Assert.Equal("Internal", result.Links[1].Text);
        Assert.Equal(2, result.Pages[0].Links.Count);
    }

    [Fact]
    public void OverlappingNativeTextOwnershipIsRejected() {
        var doc = Create(XpsFormat.Xps, new[] { "Text" }); var ns = doc.StructureNamespace;
        doc.Pages[0].ReplaceStoryFragmentsMarkup(Fragments(doc, null, null, Paragraph(ns, "Text", "Text")));
        Assert.Throws<NotSupportedException>(() => doc.ToOfficeDocumentReadResult());
    }

    [Theory]
    [InlineData(XpsFormat.Xps)]
    [InlineData(XpsFormat.OpenXps)]
    public void NativeOrderAndPageIdentitySurviveReaderTransport(XpsFormat format) {
        var doc = Create(format, new[] { "Alpha", "Extra" }, new[] { "Beta" }); var ns = doc.StructureNamespace;
        doc.Pages[0].ReplaceStoryFragmentsMarkup(Fragments(doc, "body", "one", Paragraph(ns, "Alpha")));
        doc.Pages[1].ReplaceStoryFragmentsMarkup(Fragments(doc, "body", "two", Paragraph(ns, "Beta")));
        doc.Documents[0].ReplaceDocumentStructureMarkup(Story(doc, "body", (2, "two"), (1, "one")));
        var native = doc.ToOfficeDocumentModel("ordered.oxps");
        Assert.Equal(OfficeDocumentFormat.Xps, native.Format);
        Assert.Equal(new[] { "Beta", "Alpha", "Extra" }, native.Blocks.Select(b => b.Text));
        var reader = new OfficeDocumentReaderBuilder().AddXpsHandler().Build();
        var result = reader.ReadDocument(doc.Save(), "ordered.oxps");
        Assert.Equal(ReaderInputKind.Xps, result.Kind); Assert.Equal(OfficeDocumentPageProvenance.Native, result.GetPageProvenance());
        Assert.Equal(new int?[] { 2, 1, 1 }, result.Chunks.Select(c => c.Location.Page));
        Assert.Equal(new[] { "Beta", "Alpha", "Extra" }, result.EnumerateBlocks().Select(b => b.Text));
        var restored = OfficeDocumentReadResultJson.Deserialize(OfficeDocumentReadResultJson.Serialize(result));
        Assert.Equal(new[] { "Beta", "Alpha", "Extra" }, restored.EnumerateBlocks().Select(b => b.Text));
        Assert.NotNull(result.Source.SourceHash); Assert.All(result.Chunks, c => Assert.Equal(result.Source.SourceHash, c.SourceHash));
        result.SchemaVersion = 9; Assert.Throws<System.Text.Json.JsonException>(() => OfficeDocumentReadResultJson.Serialize(result));
    }

    [Fact]
    public void StreamSnapshotPreservesPositionOwnershipAndCapturedHash() {
        var doc = Create(XpsFormat.OpenXps, new[] { "Text" }); byte[] bytes = doc.Save();
        using var stream = new MemoryStream(bytes, false); stream.Position = 3;
        var reader = new OfficeDocumentReaderBuilder().AddXpsHandler().Build();
        var result = reader.ReadDocument(stream, "input.xps");
        Assert.Equal(3, stream.Position); Assert.True(stream.CanRead);
        Assert.Equal(bytes.LongLength, result.Source.LengthBytes);
        Assert.Equal(OfficeDocumentAssetHash.ComputeSha256Hex(bytes), result.Source.SourceHash);
        Assert.Equal("Text", Assert.Single(result.Chunks).Text);
        var loaded = doc.ToOfficeDocumentReadResult(); Assert.Null(loaded.Source.SourceHash); Assert.Null(Assert.Single(loaded.Chunks).SourceHash);
    }

    [Fact]
    public void NonSeekableInputIsBoundedAndLeftOpen() {
        var doc = Create(XpsFormat.Xps, new[] { "Text" }); using var stream = new NonSeekable(doc.Save());
        var reader = new OfficeDocumentReaderBuilder().AddXpsHandler().Build();
        Assert.Throws<IOException>(() => reader.ReadDocument(stream, "input.xps", new ReaderOptions { MaxInputBytes = 8 }));
        Assert.True(stream.CanRead); Assert.False(stream.Disposed);
    }

    [Fact]
    public void RegisteredLimitsAreSnapshottedAndCannotBeRaisedByReaderOptions() {
        var doc = Create(XpsFormat.OpenXps, new[] { "Text" });
        var options = new ReaderXpsOptions { ReadOptions = new XpsReadOptions { MaximumInputBytes = 8 } };
        var reader = new OfficeDocumentReaderBuilder().AddXpsHandler(options).Build();
        options.ReadOptions.MaximumInputBytes = int.MaxValue;
        Assert.Throws<IOException>(() => reader.ReadDocument(doc.Save(), "input.oxps", new ReaderOptions { MaxInputBytes = int.MaxValue }));
    }

    [Fact]
    public void NativeTableSpansAndListMarkersRemainAvailable() {
        var doc = Create(XpsFormat.OpenXps, new[] { "Marker", "Body", "Cell", "Other" }); var ns = doc.StructureNamespace;
        doc.Pages[0].ReplaceStoryFragmentsMarkup(Fragments(doc, null, null,
            new XElement(ns + "ListStructure", new XElement(ns + "ListItemStructure", new XAttribute("Marker", "Marker"), Paragraph(ns, "Body"))),
            new XElement(ns + "TableStructure", new XElement(ns + "TableRowGroupStructure",
                new XElement(ns + "TableRowStructure", new XElement(ns + "TableCellStructure", new XAttribute("RowSpan", "2"), new XAttribute("ColumnSpan", "2"), Paragraph(ns, "Cell"))),
                new XElement(ns + "TableRowStructure", new XElement(ns + "TableCellStructure", Paragraph(ns, "Other")))))));
        var native = doc.ToOfficeDocumentModel();
        var table = Assert.Single(native.Tables); Assert.Equal(new[] { "Cell", "", "" }, table.Rows[0]); Assert.Equal(new[] { "", "", "Other" }, table.Rows[1]);
        var result = doc.ToOfficeDocumentReadResult(readerOptions: new ReaderOptions { MaxTableRows = 1 });
        var projected = Assert.Single(result.Tables); Assert.Single(projected.Rows); Assert.True(projected.Truncated); Assert.Equal(2, projected.TotalRowCount);
        Assert.Equal(new[] { "Marker", "Body", "Cell", "Other" }, result.Chunks.Select(c => c.Text));
        Assert.Equal("list-marker", result.Blocks[0].Kind);
    }

    [Fact]
    public void PageMarkdownEmitsTableTextOnceAndRetainsRowsOutsideTheGridPrefix() {
        var doc = Create(XpsFormat.Xps, new[] { "FirstCell", "LastCell" }); var ns = doc.StructureNamespace;
        doc.Pages[0].ReplaceStoryFragmentsMarkup(Fragments(doc, null, null,
            new XElement(ns + "TableStructure", new XElement(ns + "TableRowGroupStructure",
                new XElement(ns + "TableRowStructure", new XElement(ns + "TableCellStructure", Paragraph(ns, "FirstCell"))),
                new XElement(ns + "TableRowStructure", new XElement(ns + "TableCellStructure", Paragraph(ns, "LastCell")))))));
        foreach (int rows in new[] { 1, 2 }) {
            var result = doc.ToOfficeDocumentReadResult(readerOptions: new ReaderOptions { MaxTableRows = rows });
            string markdown = Assert.Single(result.GetPageMarkdown()).Markdown;
            Assert.Equal(1, markdown.Split(new[] { "FirstCell" }, StringSplitOptions.None).Length - 1);
            Assert.Equal(1, markdown.Split(new[] { "LastCell" }, StringSplitOptions.None).Length - 1);
            Assert.Equal("FirstCellLastCell", string.Concat(result.EnumerateBlocks().Select(b => b.Text)));
            Assert.Equal("1", Assert.Single(result.Metadata, m => m.Id == "reader-table-count").Value);
        }
    }

    [Fact]
    public void SummaryCountsDescribeNativeBlocksRatherThanChunkSplits() {
        var result = Create(XpsFormat.OpenXps, new[] { new string('x', 9000) }).ToOfficeDocumentReadResult();
        Assert.Single(result.Blocks); Assert.Equal(2, result.Chunks.Count);
        Assert.Equal("1", Assert.Single(result.Metadata, m => m.Id == "reader-block-count").Value);
        Assert.Equal("2", Assert.Single(result.Metadata, m => m.Id == "reader-chunk-count").Value);
        Assert.Equal("1", Assert.Single(result.Metadata, m => m.Id == "reader-page-count").Value);
    }

    [Fact]
    public void NestedTableMarkdownUsesTheRepresentedOuterCellOnce() {
        var doc = Create(XpsFormat.OpenXps, new[] { "Before", "Inner", "After" }); var ns = doc.StructureNamespace;
        XElement Table(params XElement[] children) => new(ns + "TableStructure", new XElement(ns + "TableRowGroupStructure",
            new XElement(ns + "TableRowStructure", new XElement(ns + "TableCellStructure", children))));
        doc.Pages[0].ReplaceStoryFragmentsMarkup(Fragments(doc, null, null, Table(Paragraph(ns, "Before"), Table(Paragraph(ns, "Inner")), Paragraph(ns, "After"))));
        var result = doc.ToOfficeDocumentReadResult();
        Assert.Equal(2, result.Tables.Count); Assert.Single(result.Pages[0].Tables);
        foreach (var transport in new[] { result, OfficeDocumentReadResultJson.Deserialize(OfficeDocumentReadResultJson.Serialize(result)) }) {
            string markdown = Assert.Single(transport.GetPageMarkdown()).Markdown;
            foreach (string text in new[] { "Before", "Inner", "After" })
                Assert.Equal(1, markdown.Split(new[] { text }, StringSplitOptions.None).Length - 1);
        }
    }

    [Fact]
    public void CrossPageTableKeepsTheLogicalGridAndPhysicalPageTextSeparate() {
        var doc = Create(XpsFormat.Xps, new[] { "Alpha" }, new[] { "Beta" }); var ns = doc.StructureNamespace;
        XElement Table(string text) => new(ns + "TableStructure", new XElement(ns + "TableRowGroupStructure",
            new XElement(ns + "TableRowStructure", new XElement(ns + "TableCellStructure", Paragraph(ns, text)))));
        doc.Pages[0].ReplaceStoryFragmentsMarkup(Fragments(doc, "body", "one", Table("Alpha")));
        doc.Pages[1].ReplaceStoryFragmentsMarkup(Fragments(doc, "body", "two", Table("Beta")));
        doc.Documents[0].ReplaceDocumentStructureMarkup(Story(doc, "body", (1, "one"), (2, "two")));
        var result = doc.ToOfficeDocumentReadResult();
        Assert.Single(result.Tables); Assert.All(result.Pages, page => Assert.Empty(page.Tables));
        var markdown = result.GetPageMarkdown();
        Assert.Contains("Alpha", markdown[0].Markdown); Assert.DoesNotContain("Beta", markdown[0].Markdown);
        Assert.Contains("Beta", markdown[1].Markdown); Assert.DoesNotContain("Alpha", markdown[1].Markdown);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "XpsCrossPageTable");
    }

    [Fact]
    public void DanglingRowSpanRetainsNativeStructureWithoutInventingGridRows() {
        var doc = Create(XpsFormat.Xps, new[] { "First", "Last" }); var ns = doc.StructureNamespace;
        doc.Pages[0].ReplaceStoryFragmentsMarkup(Fragments(doc, null, null,
            new XElement(ns + "TableStructure", new XElement(ns + "TableRowGroupStructure",
                new XElement(ns + "TableRowStructure", new XElement(ns + "TableCellStructure", Paragraph(ns, "First"))),
                new XElement(ns + "TableRowStructure", new XElement(ns + "TableCellStructure", new XAttribute("RowSpan", "2"), Paragraph(ns, "Last")))))));
        var model = doc.ToOfficeDocumentModel();
        Assert.Empty(model.Tables);
        Assert.Contains(model.Diagnostics, d => d.Code == "XpsTableGridUnavailable");
        Assert.Equal("FirstLast", string.Concat(model.Blocks.Select(b => b.Text)));
        var cell = model.Structure[0].Children[0].Children[0].Children[1].Children[0];
        Assert.Equal("2", cell.Attributes["rowSpan"]);
    }

    [Fact]
    public void TextProjectionExcludesResourcesAndReportsMissingUnicode() {
        var doc = Create(XpsFormat.Xps, new[] { "Visible", "NoUnicode" }); var page = doc.Pages[0]; var xml = page.GetMarkup();
        xml.Elements().Last().Attribute("UnicodeString")!.Remove();
        xml.Add(new XElement(xml.Name.Namespace + "FixedPage.Resources", new XElement(xml.Name.Namespace + "ResourceDictionary",
            new XElement(xml.Name.Namespace + "VisualBrush", new XElement(xml.Name.Namespace + "VisualBrush.Visual",
                new XElement(xml.Name.Namespace + "Glyphs", new XAttribute("UnicodeString", "ResourceOnly")))))));
        page.ReplaceMarkup(xml);
        Assert.DoesNotContain("ResourceOnly", page.ExtractText());
        var result = doc.ToOfficeDocumentReadResult(); Assert.Equal("Visible", Assert.Single(result.Chunks).Text);
        Assert.Contains(result.Diagnostics, d => d.Code == "XpsGlyphTextUnavailable");
        Assert.Contains(result.Diagnostics, d => d.Code == "XpsMarkupTextOrder");
    }

    [Fact]
    public void BoundedChunksKeepAllUnicodeScalars() {
        string text = new string('a', 255) + "😀b";
        var doc = Create(XpsFormat.OpenXps, new[] { text });
        var result = doc.ToOfficeDocumentReadResult(readerOptions: new ReaderOptions { MaxChars = 256 });
        Assert.Equal(2, result.Chunks.Count);
        Assert.Equal(text, string.Concat(result.Chunks.Select(c => c.Text)));
        Assert.All(result.Chunks, c => Assert.False(char.IsHighSurrogate(c.Text[c.Text.Length - 1])));
        Assert.All(result.Chunks, c => Assert.False(char.IsLowSurrogate(c.Text[0])));
    }

    [Fact]
    public void PreviewUsesNativeSvgAndPreCanceledProjectionStops() {
        var doc = Create(XpsFormat.OpenXps, new[] { "Text" });
        var result = doc.ToOfficeDocumentReadResult(xpsOptions: new ReaderXpsOptions { IncludeSvgPreviewAssets = true });
        var asset = Assert.Single(result.Assets); Assert.Equal("image/svg+xml", asset.MediaType); Assert.NotEmpty(asset.PayloadBytes!);
        using var canceled = new CancellationTokenSource(); canceled.Cancel();
        Assert.Throws<OperationCanceledException>(() => doc.ToOfficeDocumentReadResult(cancellationToken: canceled.Token));
    }

    private sealed class NonSeekable : Stream {
        private readonly MemoryStream _source;
        internal NonSeekable(byte[] bytes) => _source = new MemoryStream(bytes, false);
        internal bool Disposed { get; private set; }
        public override bool CanRead => !Disposed;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override int Read(byte[] buffer, int offset, int count) => _source.Read(buffer, offset, count);
        protected override void Dispose(bool disposing) { Disposed = true; if (disposing) _source.Dispose(); base.Dispose(disposing); }
        public override void Flush() => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }
}
