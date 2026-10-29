using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.Pdf;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word {
    /// <summary>
    /// Page numbers of body bookmarks taken from OfficeIMO's own page layout (the native PDF layout engine).
    /// </summary>
    internal sealed class WordLayoutPagination {
        private readonly Dictionary<string, int> _bookmarkPages;

        private WordLayoutPagination(Dictionary<string, int> bookmarkPages, int pageCount) {
            _bookmarkPages = bookmarkPages;
            PageCount = pageCount;
        }

        /// <summary>Gets the number of laid out pages.</summary>
        internal int PageCount { get; }

        internal bool TryGetBookmarkPage(string bookmarkName, out int pageNumber) {
            return _bookmarkPages.TryGetValue(bookmarkName, out pageNumber);
        }

        /// <summary>
        /// Lays out a copy of the document and records the displayed page number of every body bookmark.
        /// </summary>
        internal static WordLayoutPagination Create(WordDocument document) {
            using var cloneStream = new MemoryStream();
            using (document._wordprocessingDocument.Clone(cloneStream, true)) {
            }

            cloneStream.Position = 0;
            using WordDocument clone = WordDocument.Load(cloneStream, new WordLoadOptions {
                MaxInputBytes = Math.Max(WordLoadOptions.DefaultMaxInputBytes, cloneStream.Length)
            });

            Body? body = clone._wordprocessingDocument.MainDocumentPart?.Document?.Body;
            if (body == null) {
                return new WordLayoutPagination(new Dictionary<string, int>(StringComparer.Ordinal), 1);
            }

            UnwrapTablesOfContents(body);
            MoveBookmarksIntoParagraphs(body);

            PdfCore.PdfDocument pdf = clone.ToPdfDocument();
            PdfCore.PdfReadDocument read = pdf.GetReadDocument();
            IReadOnlyList<PdfCore.PdfPageLabel> labels = read.PageLabels;

            var pages = new Dictionary<string, int>(StringComparer.Ordinal);
            foreach (PdfCore.PdfNamedDestination destination in read.NamedDestinations) {
                if (destination.PageNumber is int physicalPage && !pages.ContainsKey(destination.Name)) {
                    pages.Add(destination.Name, GetDisplayedPageNumber(physicalPage, labels));
                }
            }

            // Bookmarks the layout does not project as destinations (for example a second bookmark in a paragraph,
            // or one inside a table cell) take the page of their paragraph or of the closest preceding projected one.
            int currentPage = 1;
            foreach (Paragraph paragraph in body.Descendants<Paragraph>()) {
                List<string> names = paragraph.Descendants<BookmarkStart>()
                    .Select(_ => _.Name?.Value)
                    .OfType<string>()
                    .Where(_ => _.Length > 0)
                    .ToList();
                foreach (string name in names) {
                    if (pages.TryGetValue(name, out int page)) {
                        currentPage = page;
                        break;
                    }
                }

                foreach (string name in names) {
                    if (!pages.ContainsKey(name)) {
                        pages.Add(name, currentPage);
                    }
                }
            }

            return new WordLayoutPagination(pages, Math.Max(1, read.Pages.Count));
        }

        /// <summary>
        /// The layout engine draws table-of-contents controls with its own generated entries. Unwrapping them lays out
        /// the entry paragraphs that are actually stored in the document, so the pages after the table match.
        /// </summary>
        private static void UnwrapTablesOfContents(Body body) {
            foreach (SdtBlock block in body.Descendants<SdtBlock>().Where(IsTableOfContentBlock).ToList()) {
                SdtContentBlock? content = block.SdtContentBlock;
                if (content == null || block.Parent == null) {
                    continue;
                }

                foreach (var child in content.ChildElements.ToList()) {
                    child.Remove();
                    block.Parent.InsertBefore(child, block);
                }

                block.Remove();
            }
        }

        private static bool IsTableOfContentBlock(SdtBlock block) {
            string? gallery = block.SdtProperties?.GetFirstChild<SdtContentDocPartObject>()?.GetFirstChild<DocPartGallery>()?.Val?.Value;
            return string.Equals(gallery, "Table of Contents", StringComparison.Ordinal);
        }

        /// <summary>
        /// Bookmarks that start between paragraphs are moved into the following paragraph so the layout records them.
        /// </summary>
        private static void MoveBookmarksIntoParagraphs(Body body) {
            foreach (BookmarkStart bookmark in body.Descendants<BookmarkStart>().ToList()) {
                if (bookmark.Ancestors<Paragraph>().Any()) {
                    continue;
                }

                Paragraph? next = bookmark.ElementsAfter().OfType<Paragraph>().FirstOrDefault() ??
                    bookmark.ElementsAfter().SelectMany(_ => _.Descendants<Paragraph>()).FirstOrDefault();
                if (next == null) {
                    continue;
                }

                bookmark.Remove();
                if (next.ParagraphProperties != null) {
                    next.InsertAfter(bookmark, next.ParagraphProperties);
                } else {
                    next.PrependChild(bookmark);
                }
            }
        }

        private static int GetDisplayedPageNumber(int physicalPage, IReadOnlyList<PdfCore.PdfPageLabel> labels) {
            PdfCore.PdfPageLabel? label = labels
                .Where(_ => _.StartPageIndex <= physicalPage - 1)
                .OrderByDescending(_ => _.StartPageIndex)
                .FirstOrDefault();
            if (label == null) {
                return physicalPage;
            }

            return (label.StartNumber ?? 1) + (physicalPage - 1 - label.StartPageIndex);
        }
    }
}
