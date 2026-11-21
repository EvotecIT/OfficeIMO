using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    /// <summary>
    /// Page numbers of body bookmarks taken from OfficeIMO's own page layout (the native PDF layout engine).
    /// </summary>
    internal static class WordPdfLayoutPagination {
        /// <summary>
        /// Lays out a copy of the document with the PDF export options and records the displayed page number of every
        /// body bookmark. The copy goes through the same conversion as PDF export with the same options, so fonts,
        /// substitutions, fallbacks and the resource policy resolve the same way.
        /// </summary>
        internal static WordLayoutPagination Create(WordDocument document, WordPdfExportOptions layoutOptions) {
            using var cloneStream = new MemoryStream();
            using (document._wordprocessingDocument.Clone(cloneStream, true)) {
            }

            cloneStream.Position = 0;
            using WordDocument clone = WordDocument.Load(cloneStream, new WordLoadOptions {
                MaxInputBytes = Math.Max(WordLoadOptions.DefaultMaxInputBytes, cloneStream.Length)
            });

            // Preserve the caller's pending normalization state on the private layout copy.
            if (document.TableNormalizationRequired) {
                clone.MarkTableNormalizationRequired();
                clone.NormalizeTablesForOnline();
            }

            Body? body = clone._wordprocessingDocument.MainDocumentPart?.Document?.Body;
            if (body == null) {
                return new WordLayoutPagination(new Dictionary<string, (int Number, string? Label)>(StringComparer.Ordinal), 1);
            }

            MoveBookmarksIntoParagraphs(body);

            WordPdfExportOptions options = layoutOptions.Clone();
            options.IncludeLayoutPageLabels = true;
            // A field update reports through its own result: loss findings must not stop it, and DATE/TIME belong to the update.
            options.FailOnLoss = false;
            options.DateTimeFieldUpdateOptions = null;
            using PdfCore.PdfDocument pdf = clone.ToPdf(options);
            PdfCore.PdfReadDocument read = pdf.GetReadDocument();
            IReadOnlyList<PdfCore.PdfPageLabel> labels = read.PageLabels;

            var pages = new Dictionary<string, int>(StringComparer.Ordinal);
            foreach (PdfCore.PdfNamedDestination destination in read.NamedDestinations) {
                if (destination.PageNumber is int physicalPage && !pages.ContainsKey(destination.Name)) {
                    pages.Add(destination.Name, physicalPage);
                }
            }

            // A PAGE field marker is a link drawn on the field's own line; that line's page wins over the paragraph's first line.
            var fieldLines = new HashSet<string>(StringComparer.Ordinal);
            for (int pageIndex = 0; pageIndex < read.Pages.Count; pageIndex++) {
                foreach (PdfCore.PdfLinkAnnotation link in read.Pages[pageIndex].GetLinkAnnotations()) {
                    if (link.DestinationName is string name && name.StartsWith(WordFieldUpdater.PageFieldBookmarkPrefix, StringComparison.Ordinal)
                        && fieldLines.Add(name)) {
                        pages[name] = link.PageNumber ?? pageIndex + 1;
                    }
                }
            }

            int pageCount = Math.Max(1, read.Pages.Count);
            AddUnprojectedBookmarks(body, pages, pageCount);
            var displayedPages = pages.ToDictionary(pair => pair.Key, pair => GetDisplayedPageNumber(pair.Value, labels), StringComparer.Ordinal);
            return new WordLayoutPagination(displayedPages, pageCount);
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

        private static (int Number, string? Label) GetDisplayedPageNumber(int physicalPage, IReadOnlyList<PdfCore.PdfPageLabel> labels) {
            PdfCore.PdfPageLabel? label = labels
                .Where(_ => _.StartPageIndex <= physicalPage - 1)
                .OrderByDescending(_ => _.StartPageIndex)
                .FirstOrDefault();
            if (label == null) {
                return (physicalPage, physicalPage.ToString(System.Globalization.CultureInfo.InvariantCulture));
            }

            int number = (label.StartNumber ?? 1) + (physicalPage - 1 - label.StartPageIndex);
            if (label.Prefix == PdfCore.PdfCreateOptions.UnsupportedLayoutPageNumberPrefix) return (number, null);
            PdfCore.PdfPageNumberStyle style = label.Style switch {
                "r" => PdfCore.PdfPageNumberStyle.LowerRoman,
                "R" => PdfCore.PdfPageNumberStyle.UpperRoman,
                "a" => PdfCore.PdfPageNumberStyle.LowerLetter,
                "A" => PdfCore.PdfPageNumberStyle.UpperLetter,
                _ => PdfCore.PdfPageNumberStyle.Arabic
            };
            return (number, PdfCore.PdfPageNumberFormatter.Format(number, style));
        }
    }
}
