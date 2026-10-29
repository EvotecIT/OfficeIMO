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
            // Saving normalizes table grids from the table and cell widths; the layout reads those grids.
            document.NormalizeTablesForOnline();
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

            int pageCount = Math.Max(1, read.Pages.Count);
            AddUnprojectedBookmarks(body, pages, GetDisplayedPageNumber(pageCount, labels));
            return new WordLayoutPagination(pages, pageCount);
        }

        /// <summary>
        /// The layout does not record every bookmark as a destination: a second bookmark in a paragraph, or one in a
        /// table cell, has none. Such a bookmark takes the page of its own paragraph when another bookmark there was
        /// recorded, or a page that the recorded bookmarks before and after it pin down (both on the same page, or the
        /// last page with none after it, or page 1 with none before it). Any other bookmark stays unplaced.
        /// </summary>
        private static void AddUnprojectedBookmarks(Body body, Dictionary<string, int> pages, int lastPage) {
            var paragraphs = new List<(List<string> Names, int? Page)>();
            foreach (Paragraph paragraph in body.Descendants<Paragraph>()) {
                List<string> names = paragraph.Descendants<BookmarkStart>()
                    .Select(_ => _.Name?.Value)
                    .OfType<string>()
                    .Where(_ => _.Length > 0)
                    .ToList();
                if (names.Count == 0) {
                    continue;
                }

                int? page = null;
                foreach (string name in names) {
                    if (pages.TryGetValue(name, out int recorded)) {
                        page = recorded;
                        break;
                    }
                }

                paragraphs.Add((names, page));
            }

            for (int index = 0; index < paragraphs.Count; index++) {
                (List<string> names, int? page) = paragraphs[index];
                page ??= GetBracketedPage(paragraphs, index, lastPage);
                if (page == null) {
                    continue;
                }

                foreach (string name in names) {
                    if (!pages.ContainsKey(name)) {
                        pages.Add(name, page.Value);
                    }
                }
            }
        }

        private static int? GetBracketedPage(List<(List<string> Names, int? Page)> paragraphs, int index, int lastPage) {
            int? before = null;
            for (int previous = index - 1; previous >= 0 && before == null; previous--) {
                before = paragraphs[previous].Page;
            }

            int? after = null;
            for (int next = index + 1; next < paragraphs.Count && after == null; next++) {
                after = paragraphs[next].Page;
            }

            if (before != null && after != null) {
                if (before == after) {
                    return before;
                }

                return null;
            }

            if (before != null && before == lastPage) {
                return before;
            }

            if (after != null && after == 1) {
                return after;
            }

            return null;
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

                // The first paragraph in document order: a following paragraph, or the first one inside a following table.
                Paragraph? next = bookmark.ElementsAfter()
                    .Select(_ => _ as Paragraph ?? _.Descendants<Paragraph>().FirstOrDefault())
                    .FirstOrDefault(_ => _ != null);
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
