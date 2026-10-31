using System.Collections.Generic;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Pdf.Word {
    internal static partial class PdfWordConverter {
        private static ImportNavigationMap BuildNavigationMap(PdfCore.PdfDocumentReadResult source, WordPdfImportOptions options) {
            if (!options.ImportInternalLinks) {
                return ImportNavigationMap.Empty;
            }

            bool hasInternalLinks = source.Links.Any(link => link.IsInternalDestinationLink);
            bool hasNamedDestinations = source.NamedDestinations.Count > 0;
            if (!hasInternalLinks && !hasNamedDestinations) {
                return ImportNavigationMap.Empty;
            }

            string prefix = string.IsNullOrWhiteSpace(options.BookmarkPrefix) ? "OfficeIMO_Pdf" : options.BookmarkPrefix;
            var used = new HashSet<string>(StringComparer.Ordinal);
            var pageAnchors = new Dictionary<int, string>();
            var anchorsByPage = new Dictionary<int, List<string>>();
            var namedAnchors = new Dictionary<string, string>(StringComparer.Ordinal);
            var destinationTops = new Dictionary<string, double?>(StringComparer.Ordinal);
            var importedPageNumbers = new HashSet<int>(source.Pages.Select(page => page.PageNumber));

            foreach (int pageNumber in importedPageNumbers.OrderBy(pageNumber => pageNumber)) {
                string pageAnchor = CreateUniqueBookmarkName(prefix + "_Page_" + pageNumber.ToString(CultureInfo.InvariantCulture), used);
                pageAnchors[pageNumber] = pageAnchor;
                anchorsByPage[pageNumber] = new List<string> { pageAnchor };
            }

            for (int i = 0; i < source.NamedDestinations.Count; i++) {
                PdfCore.PdfNamedDestination destination = source.NamedDestinations[i];
                if (!destination.PageNumber.HasValue || !importedPageNumbers.Contains(destination.PageNumber.Value)) {
                    continue;
                }

                if (namedAnchors.ContainsKey(destination.Name)) continue;
                string namedAnchor = CreateUniqueBookmarkName(prefix + "_Dest_" + destination.Name, used);
                namedAnchors[destination.Name] = namedAnchor;
                anchorsByPage[destination.PageNumber.Value].Add(namedAnchor);
                destinationTops[namedAnchor] = destination.DestinationTop;
            }

            return new ImportNavigationMap(pageAnchors, namedAnchors, anchorsByPage, destinationTops);
        }

        /// <summary>
        /// Places every page and named-destination anchor on the imported paragraph at its source position.
        /// Anchors that land on the same paragraph share one Word bookmark, so each paragraph carries at most one.
        /// </summary>
        private static ImportNavigationMap PlaceNavigationAnchors(
            PdfCore.PdfDocumentReadResult source,
            IReadOnlyList<List<ImportItem>> pageItems,
            ImportNavigationMap navigation) {
            if (navigation.IsEmpty) return navigation;
            var primaryByAnchor = new Dictionary<string, string>(StringComparer.Ordinal);
            var itemAnchors = new Dictionary<(int PageNumber, int ItemIndex), string>();
            var pageStartAnchors = new Dictionary<int, string>();
            for (int pageIndex = 0; pageIndex < source.Pages.Count; pageIndex++) {
                int pageNumber = source.Pages[pageIndex].PageNumber;
                IReadOnlyList<string> anchors = navigation.GetAnchorsForPage(pageNumber);
                if (anchors.Count == 0) continue;
                List<ImportItem> items = pageItems[pageIndex];
                int firstAnchorable = items.FindIndex(static item => CanCarryBookmark(item.Kind));
                for (int anchorIndex = 0; anchorIndex < anchors.Count; anchorIndex++) {
                    string anchor = anchors[anchorIndex];
                    int slot = firstAnchorable;
                    if (slot >= 0 && navigation.TryGetDestinationTop(anchor, out double top)) {
                        slot = FindDestinationItem(items, top, firstAnchorable);
                    }
                    if (slot < 0) {
                        if (pageStartAnchors.TryGetValue(pageNumber, out string? pageStart)) {
                            primaryByAnchor[anchor] = pageStart;
                        } else {
                            pageStartAnchors[pageNumber] = anchor;
                            primaryByAnchor[anchor] = anchor;
                        }
                        continue;
                    }
                    if (itemAnchors.TryGetValue((pageNumber, slot), out string? primary)) {
                        primaryByAnchor[anchor] = primary;
                    } else {
                        itemAnchors[(pageNumber, slot)] = anchor;
                        primaryByAnchor[anchor] = anchor;
                    }
                }
            }
            return navigation.WithPlacement(primaryByAnchor, itemAnchors, pageStartAnchors);
        }

        private static bool CanCarryBookmark(ImportItemKind kind) => kind is
            ImportItemKind.Heading or ImportItemKind.Paragraph or ImportItemKind.TextBlock or
            ImportItemKind.ListItem or ImportItemKind.Link;

        private static int FindDestinationItem(List<ImportItem> items, double destinationTop, int fallback) {
            // PDF destinations name the top of the target area; choose the first paragraph that starts at or below it.
            const double tolerance = 4D;
            int best = -1;
            double bestDistance = double.MaxValue;
            for (int index = 0; index < items.Count; index++) {
                ImportItem item = items[index];
                if (!CanCarryBookmark(item.Kind) || item.Y > destinationTop + tolerance) continue;
                double distance = destinationTop - item.Y;
                if (distance < bestDistance) {
                    best = index;
                    bestDistance = distance;
                }
            }
            return best >= 0 ? best : fallback;
        }

        private static bool AddNavigationBookmarks(WordDocument document, PdfCore.PdfLogicalPage page, ImportNavigationMap navigation) {
            if (!navigation.TryGetPageStartAnchor(page.PageNumber, out string? anchor)) {
                return false;
            }

            document.AddParagraph().AddBookmark(anchor!);
            return true;
        }

        private static bool TryResolveWordLinkTarget(
            PdfCore.PdfLogicalLinkAnnotation link,
            WordPdfImportOptions options,
            ImportNavigationMap navigation,
            out WordLinkTarget target) {
            if (options.ImportUriLinks &&
                !string.IsNullOrWhiteSpace(link.Uri) &&
                TryCreateWordHyperlinkUri(link, options, out Uri? uri)) {
                target = WordLinkTarget.ForUri(uri!);
                return true;
            }

            if (options.ImportInternalLinks && link.IsInternalDestinationLink) {
                if (!string.IsNullOrWhiteSpace(link.DestinationName) &&
                    navigation.NamedAnchors.TryGetValue(link.DestinationName!, out string? namedAnchor)) {
                    target = WordLinkTarget.ForAnchor(namedAnchor);
                    return true;
                }

                if (link.DestinationPageNumber.HasValue &&
                    navigation.PageAnchors.TryGetValue(link.DestinationPageNumber.Value, out string? pageAnchor)) {
                    target = WordLinkTarget.ForAnchor(pageAnchor);
                    return true;
                }
            }

            target = default;
            return false;
        }

        private static string GetInternalLinkDisplayText(PdfCore.PdfLogicalLinkAnnotation link) {
            if (!string.IsNullOrWhiteSpace(link.Contents)) {
                return link.Contents!;
            }

            if (!string.IsNullOrWhiteSpace(link.DestinationName)) {
                return link.DestinationName!;
            }

            if (link.DestinationPageNumber.HasValue) {
                return "Page " + link.DestinationPageNumber.Value.ToString(CultureInfo.InvariantCulture);
            }

            return "PDF internal link";
        }

        private static string CreateUniqueBookmarkName(string candidate, HashSet<string> used) {
            string normalized = NormalizeBookmarkName(candidate);
            string unique = normalized;
            int suffix = 1;
            while (!used.Add(unique)) {
                string suffixText = "_" + suffix.ToString(CultureInfo.InvariantCulture);
                int maxBaseLength = Math.Max(1, 40 - suffixText.Length);
                unique = normalized.Length > maxBaseLength
                    ? normalized.Substring(0, maxBaseLength) + suffixText
                    : normalized + suffixText;
                suffix++;
            }

            return unique;
        }

        private static string NormalizeBookmarkName(string value) {
            string seed = string.IsNullOrWhiteSpace(value) ? "OfficeIMO_Pdf_Link" : value;
            var builder = new StringBuilder(seed.Length);
            for (int i = 0; i < seed.Length; i++) {
                char ch = seed[i];
                builder.Append(char.IsLetterOrDigit(ch) || ch == '_' ? ch : '_');
            }

            if (builder.Length == 0 || (!char.IsLetter(builder[0]) && builder[0] != '_')) {
                builder.Insert(0, 'B');
            }

            string normalized = builder.ToString();
            if (normalized.Length <= 40) {
                return normalized;
            }

            string hash = ComputeShortHash(normalized);
            return normalized.Substring(0, 31) + "_" + hash;
        }

        private static string ComputeShortHash(string value) {
            using SHA256 sha = SHA256.Create();
            byte[] hash = sha.ComputeHash(Encoding.UTF8.GetBytes(value));
            var builder = new StringBuilder(8);
            for (int i = 0; i < 4; i++) {
                builder.Append(hash[i].ToString("x2", CultureInfo.InvariantCulture));
            }

            return builder.ToString();
        }

        private readonly struct WordLinkTarget {
            private WordLinkTarget(Uri? uri, string? anchor) {
                Uri = uri;
                Anchor = anchor;
            }

            public Uri? Uri { get; }

            public string? Anchor { get; }

            public bool IsUri => Uri is not null;

            public static WordLinkTarget ForUri(Uri uri) => new WordLinkTarget(uri, null);

            public static WordLinkTarget ForAnchor(string anchor) => new WordLinkTarget(null, anchor);
        }

        private sealed class ImportNavigationMap {
            public static readonly ImportNavigationMap Empty = new ImportNavigationMap(
                new Dictionary<int, string>(),
                new Dictionary<string, string>(StringComparer.Ordinal),
                new Dictionary<int, List<string>>(),
                new Dictionary<string, double?>(StringComparer.Ordinal));

            private readonly IReadOnlyDictionary<string, double?> _destinationTops;
            private readonly IReadOnlyDictionary<(int PageNumber, int ItemIndex), string>? _itemAnchors;
            private readonly IReadOnlyDictionary<int, string>? _pageStartAnchors;

            public ImportNavigationMap(
                IReadOnlyDictionary<int, string> pageAnchors,
                IReadOnlyDictionary<string, string> namedAnchors,
                IReadOnlyDictionary<int, List<string>> anchorsByPage,
                IReadOnlyDictionary<string, double?> destinationTops,
                IReadOnlyDictionary<(int PageNumber, int ItemIndex), string>? itemAnchors = null,
                IReadOnlyDictionary<int, string>? pageStartAnchors = null) {
                PageAnchors = pageAnchors;
                NamedAnchors = namedAnchors;
                AnchorsByPage = anchorsByPage;
                _destinationTops = destinationTops;
                _itemAnchors = itemAnchors;
                _pageStartAnchors = pageStartAnchors;
            }

            public IReadOnlyDictionary<int, string> PageAnchors { get; }

            public IReadOnlyDictionary<string, string> NamedAnchors { get; }

            private IReadOnlyDictionary<int, List<string>> AnchorsByPage { get; }

            public bool IsEmpty => AnchorsByPage.Count == 0;

            public bool HasAnchorsForPage(int pageNumber) =>
                AnchorsByPage.TryGetValue(pageNumber, out List<string>? anchors) && anchors.Count > 0;

            public IReadOnlyList<string> GetAnchorsForPage(int pageNumber) =>
                AnchorsByPage.TryGetValue(pageNumber, out List<string>? anchors)
                    ? anchors.AsReadOnly()
                    : Array.Empty<string>();

            public bool TryGetDestinationTop(string anchor, out double top) {
                if (_destinationTops.TryGetValue(anchor, out double? value) && value.HasValue &&
                    !double.IsNaN(value.Value) && !double.IsInfinity(value.Value)) {
                    top = value.Value;
                    return true;
                }
                top = 0D;
                return false;
            }

            public bool TryGetItemAnchor(int pageNumber, int itemIndex, out string? anchor) {
                anchor = null;
                return _itemAnchors != null && _itemAnchors.TryGetValue((pageNumber, itemIndex), out anchor);
            }

            public bool TryGetPageStartAnchor(int pageNumber, out string? anchor) {
                if (_pageStartAnchors == null) {
                    // Without placement every page anchor starts the page.
                    return PageAnchors.TryGetValue(pageNumber, out anchor);
                }
                return _pageStartAnchors.TryGetValue(pageNumber, out anchor);
            }

            public ImportNavigationMap WithPlacement(
                IReadOnlyDictionary<string, string> primaryByAnchor,
                IReadOnlyDictionary<(int PageNumber, int ItemIndex), string> itemAnchors,
                IReadOnlyDictionary<int, string> pageStartAnchors) {
                string Resolve(string anchor) => primaryByAnchor.TryGetValue(anchor, out string? primary) ? primary : anchor;
                var pageAnchors = PageAnchors.ToDictionary(static pair => pair.Key, pair => Resolve(pair.Value));
                var namedAnchors = NamedAnchors.ToDictionary(static pair => pair.Key, pair => Resolve(pair.Value), StringComparer.Ordinal);
                return new ImportNavigationMap(pageAnchors, namedAnchors, AnchorsByPage, _destinationTops, itemAnchors, pageStartAnchors);
            }
        }
    }
}
