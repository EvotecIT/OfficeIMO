using OfficeIMO.Word;
using OfficeIMO.Word.LegacyDoc.Model;
using System.IO.Compression;
using Xunit;

namespace OfficeIMO.Tests {

    public partial class Word {
        [Theory]
        [InlineData(false, "\n")]
        [InlineData(false, "\r")]
        [InlineData(false, "\r\n")]
        [InlineData(true, "\n")]
        [InlineData(true, "\r")]
        [InlineData(true, "\r\n")]
        public void NativeDocHyperlinkMetadataRejectsFieldBoundariesWithoutChangingDocx(bool targetFrame, string boundary) {
            using WordDocument source = WordDocument.Create();
            WordParagraph paragraph = source.AddParagraph().AddHyperLink("Guide", new Uri("https://example.test/guide"));
            if (targetFrame) {
                paragraph._paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Hyperlink>().Single().TargetFrame = "Window" + boundary + "Name";
            } else {
                paragraph.Hyperlink!.Tooltip = "Line1" + boundary + "Line2";
            }
            byte[] docx = source.ToBytes();
            NotSupportedException error = Assert.Throws<NotSupportedException>(() => source.ToBytes(WordFileFormat.Doc));
            Assert.Contains("line breaks", error.Message, StringComparison.Ordinal);
            using var expected = new ZipArchive(new MemoryStream(docx), ZipArchiveMode.Read);
            using var actual = new ZipArchive(new MemoryStream(source.ToBytes()), ZipArchiveMode.Read);
            Assert.Equal(expected.Entries.Select(entry => entry.FullName).OrderBy(name => name, StringComparer.Ordinal),
                actual.Entries.Select(entry => entry.FullName).OrderBy(name => name, StringComparer.Ordinal));
            foreach (ZipArchiveEntry entry in expected.Entries) {
                using Stream expectedPart = entry.Open();
                using Stream actualPart = actual.GetEntry(entry.FullName)!.Open();
                using var expectedBytes = new MemoryStream();
                using var actualBytes = new MemoryStream();
                expectedPart.CopyTo(expectedBytes);
                actualPart.CopyTo(actualBytes);
                Assert.Equal(expectedBytes.ToArray(), actualBytes.ToArray());
            }
            Assert.Empty(source.ValidateDocument());
        }

        [Theory]
        [InlineData(false, WordHyperlinkTargetFrame._blank)]
        [InlineData(false, WordHyperlinkTargetFrame._self)]
        [InlineData(false, WordHyperlinkTargetFrame._parent)]
        [InlineData(false, WordHyperlinkTargetFrame._top)]
        [InlineData(true, WordHyperlinkTargetFrame._blank)]
        [InlineData(true, WordHyperlinkTargetFrame._self)]
        public void NativeDocHyperlinkMetadataPreservesTooltipAndTarget(bool internalLink, WordHyperlinkTargetFrame frame) {
            using WordDocument source = WordDocument.Create();
            const string tooltip = "Read \"Guide\" \\l literal — zażółć";
            WordParagraph paragraph = internalLink
                ? source.AddParagraph().AddHyperLink("Guide", "Destination", tooltip: tooltip)
                : source.AddParagraph().AddHyperLink("Guide", new Uri("https://example.test/guide"), tooltip: tooltip);
            paragraph.Hyperlink!.TargetFrame = frame;
            source.AddParagraph("Destination").AddBookmark("Destination");
            string sourceXml = source._wordprocessingDocument!.MainDocumentPart!.Document.OuterXml;

            using WordDocument native = WordDocument.Load(new MemoryStream(source.ToBytes(WordFileFormat.Doc)));
            WordHyperLink link = Assert.Single(native.HyperLinks);
            Assert.Equal("Guide", link.Text);
            Assert.Equal(tooltip, link.Tooltip);
            Assert.Equal(frame, link.TargetFrame);
            Assert.Equal(internalLink ? "Destination" : null, link.Anchor);
            Assert.Equal(internalLink ? null : "https://example.test/guide", link.Uri?.ToString());
            Assert.Equal(sourceXml, source._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
            using WordDocument docx = WordDocument.Load(new MemoryStream(native.ToBytes()));
            Assert.Equal(tooltip, Assert.Single(docx.HyperLinks).Tooltip);
            Assert.Equal(frame, Assert.Single(docx.HyperLinks).TargetFrame);
            Assert.Empty(docx.ValidateDocument());
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData(" ")]
        public void NativeDocHyperlinkMetadataPreservesAuthoredEmptyTooltip(string? tooltip) {
            using WordDocument source = WordDocument.Create();
            WordParagraph paragraph = source.AddParagraph().AddHyperLink("Guide", new Uri("https://example.test/guide"));
            paragraph.Hyperlink!.Tooltip = tooltip;
            using WordDocument loaded = WordDocument.Load(new MemoryStream(source.ToBytes(WordFileFormat.Doc)));
            Assert.Equal(tooltip, Assert.Single(loaded.HyperLinks).Tooltip);
            Assert.Null(Assert.Single(loaded.HyperLinks).TargetFrame);
            Assert.Empty(loaded.ValidateDocument());
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void NativeDocHyperlinkMetadataKeepsAdjacentDifferentLinks(bool differentFrames) {
            using WordDocument source = WordDocument.Create();
            WordParagraph first = source.AddParagraph().AddHyperLink("First", new Uri("https://example.test/guide"), tooltip: "First tip");
            first.Hyperlink!.TargetFrame = WordHyperlinkTargetFrame._self;
            WordParagraph second = first.AddHyperLink("Second", new Uri("https://example.test/guide"), tooltip: differentFrames ? "First tip" : "Second tip");
            second.Hyperlink!.TargetFrame = differentFrames ? WordHyperlinkTargetFrame._blank : WordHyperlinkTargetFrame._self;
            using WordDocument loaded = WordDocument.Load(new MemoryStream(source.ToBytes(WordFileFormat.Doc)));
            WordHyperLink[] links = loaded.HyperLinks.ToArray();
            Assert.Equal(2, links.Length);
            Assert.Equal("First", links[0].Text);
            Assert.Equal("Second", links[1].Text);
            Assert.Equal("First tip", links[0].Tooltip);
            Assert.Equal(differentFrames ? "First tip" : "Second tip", links[1].Tooltip);
            Assert.Equal(differentFrames ? WordHyperlinkTargetFrame._blank : WordHyperlinkTargetFrame._self, links[1].TargetFrame);
        }

        [Theory]
        [InlineData(" HYPERLINK \"https://example.test/guide\" \\o \"Guide tip\" \\t \"_blank\" ", false)]
        [InlineData(" HYPERLINK \\l \"Destination\" \\t \"_self\" \\o \"Guide tip\" ", true)]
        public void NativeDocHyperlinkMetadataImportsFieldSwitches(string instruction, bool internalLink) {
            byte[] bytes = LegacyDocTestBuilder.CreateSimpleDoc(LegacyDocField.Begin + instruction + LegacyDocField.Separator + "Guide" + LegacyDocField.End);
            using WordDocument loaded = WordDocument.Load(new MemoryStream(bytes));
            WordHyperLink link = Assert.Single(loaded.HyperLinks);
            Assert.Equal("Guide tip", link.Tooltip);
            Assert.Equal(internalLink ? WordHyperlinkTargetFrame._self : WordHyperlinkTargetFrame._blank, link.TargetFrame);
            Assert.Equal(internalLink ? "Destination" : null, link.Anchor);
            Assert.Equal(internalLink ? null : "https://example.test/guide", link.Uri?.ToString());
        }

        [Fact]
        public void NativeDocHyperlinkMetadataPreservesSeparateStories() {
            using WordDocument source = WordDocument.Create();
            WordParagraph[] paragraphs = {
                source.AddParagraph(),
                source.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0],
                source.HeaderDefaultOrCreate.AddParagraph(),
                source.FooterDefaultOrCreate.AddParagraph(),
                source.AddParagraph("Reference").AddFootNote("Footnote prefix").FootNote!.Paragraphs!.Single(p => p.Text == "Footnote prefix"),
                source.AddParagraph("Reference").AddEndNote("Endnote prefix").EndNote!.Paragraphs!.Single(p => p.Text == "Endnote prefix")
            };
            for (int index = 0; index < paragraphs.Length; index++) {
                WordParagraph link = paragraphs[index].AddHyperLink("Guide " + index, new Uri("https://example.test/guide"), tooltip: "Tip " + index);
                link.Hyperlink!.TargetFrame = WordHyperlinkTargetFrame._blank;
            }
            using WordDocument loaded = WordDocument.Load(new MemoryStream(source.ToBytes(WordFileFormat.Doc)));
            IEnumerable<WordParagraph>[] stories = {
                loaded.Paragraphs,
                loaded.Tables[0].Rows[0].Cells[0].Paragraphs,
                loaded.Sections[0].Header.Default!.Paragraphs,
                loaded.Sections[0].Footer.Default!.Paragraphs,
                Assert.Single(loaded.FootNotes).Paragraphs!,
                Assert.Single(loaded.EndNotes).Paragraphs!
            };
            for (int index = 0; index < stories.Length; index++) {
                WordHyperLink link = Assert.Single(DistinctHyperlinks(stories[index].SelectMany(p => p.GetRuns()).Select(p => p.Hyperlink)));
                Assert.Equal("Guide " + index, link.Text);
                Assert.Equal("Tip " + index, link.Tooltip);
                Assert.Equal(WordHyperlinkTargetFrame._blank, link.TargetFrame);
            }
        }
    }
}
