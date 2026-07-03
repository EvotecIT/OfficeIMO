using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void CoverPageTemplates_AssignPositiveSdtIds() {
            using var document = WordDocument.Create();

            document.AddCoverPage(CoverPageTemplate.Austin);

            var coverBlocks = FindDocPartBlocks(document, "Cover Pages");
            Assert.NotEmpty(coverBlocks);

            AssertSdtIdsArePositiveAndUnique(GetAllSdtIds(document));
        }

        [Fact]
        public void TableOfContentTemplates_AssignPositiveSdtIds() {
            using var document = WordDocument.Create();

            document.AddTableOfContent(TableOfContentStyle.Template2);

            var tocBlocks = FindDocPartBlocks(document, "Table of Contents");
            Assert.NotEmpty(tocBlocks);

            AssertSdtIdsArePositiveAndUnique(GetAllSdtIds(document));
        }

        [Fact]
        public void WatermarkTemplates_AssignPositiveSdtIds() {
            using var document = WordDocument.Create();

            document.AddParagraph("Section");
            document.AddHeadersAndFooters();
            document.Sections[0].AddWatermark(WordWatermarkStyle.Text, "Draft");

            var watermarkBlocks = FindDocPartBlocks(document, "Watermarks");
            Assert.NotEmpty(watermarkBlocks);

            AssertSdtIdsArePositiveAndUnique(GetAllSdtIds(document));
        }

        [Fact]
        public void PageNumberTemplates_AssignPositiveSdtIds() {
            using var document = WordDocument.Create();

            document.AddParagraph("Section");
            document.AddHeadersAndFooters();
            var footer = document.Sections[0].Footer.Default
                ?? throw new InvalidOperationException("Default footer was not initialized.");
            footer.AddPageNumber(WordPageNumberStyle.PlainNumber);

            var pageNumberBlocks = FindDocPartBlocks(document, "Page Numbers");
            Assert.NotEmpty(pageNumberBlocks);

            AssertSdtIdsArePositiveAndUnique(GetAllSdtIds(document));
        }

        private static IReadOnlyList<SdtBlock> FindDocPartBlocks(WordDocument document, string galleryPrefix) {
            var mainPart = document._wordprocessingDocument!.MainDocumentPart
                ?? throw new InvalidOperationException("Main document part is missing.");

            IEnumerable<SdtBlock> Enumerate(OpenXmlElement? root) {
                return root?.Descendants<SdtBlock>() ?? Enumerable.Empty<SdtBlock>();
            }

            return Enumerate(mainPart.Document.Body)
                .Concat(mainPart.HeaderParts.SelectMany(part => Enumerate(part.RootElement)))
                .Concat(mainPart.FooterParts.SelectMany(part => Enumerate(part.RootElement)))
                .Where(block => HasDocPartGallery(block, galleryPrefix))
                .ToList();
        }

        private static bool HasDocPartGallery(SdtBlock block, string galleryPrefix) {
            var gallery = block.SdtProperties?
                .GetFirstChild<SdtContentDocPartObject>()?
                .GetFirstChild<DocPartGallery>();

            return gallery?.Val?.Value != null && gallery.Val.Value.StartsWith(galleryPrefix, StringComparison.Ordinal);
        }

        private static IReadOnlyList<int> GetAllSdtIds(WordDocument document) {
            var mainPart = document._wordprocessingDocument!.MainDocumentPart
                ?? throw new InvalidOperationException("Main document part is missing.");

            IEnumerable<SdtId> EnumerateIds(OpenXmlElement? root) {
                return root?.Descendants<SdtId>() ?? Enumerable.Empty<SdtId>();
            }

            return EnumerateIds(mainPart.Document.Body)
                .Concat(mainPart.HeaderParts.SelectMany(part => EnumerateIds(part.RootElement)))
                .Concat(mainPart.FooterParts.SelectMany(part => EnumerateIds(part.RootElement)))
                .Where(id => id.Val?.HasValue == true)
                .Select(id => id.Val!.Value)
                .ToList();
        }

        private static void AssertSdtIdsArePositiveAndUnique(IReadOnlyList<int> ids) {
            Assert.NotEmpty(ids);
            Assert.All(ids, id => Assert.InRange(id, 1, int.MaxValue - 1));
            Assert.Equal(ids.Count, ids.Distinct().Count());
        }
    }
}
