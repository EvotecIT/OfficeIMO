using System.Collections.Generic;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    internal static partial class PdfWordConverter {
        private static ImportNavigationMap BuildNavigationMap(PdfCore.PdfLogicalDocument source, PdfWordReadOptions options) {
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

                string namedAnchor = CreateUniqueBookmarkName(prefix + "_Dest_" + destination.Name, used);
                namedAnchors[destination.Name] = namedAnchor;
                anchorsByPage[destination.PageNumber.Value].Add(namedAnchor);
            }

            return new ImportNavigationMap(pageAnchors, namedAnchors, anchorsByPage);
        }

        private static bool AddNavigationBookmarks(WordDocument document, PdfCore.PdfLogicalPage page, ImportNavigationMap navigation) {
            IReadOnlyList<string> anchors = navigation.GetAnchorsForPage(page.PageNumber);
            if (anchors.Count == 0) {
                return false;
            }

            WordParagraph paragraph = document.AddParagraph();
            for (int i = 0; i < anchors.Count; i++) {
                paragraph.AddBookmark(anchors[i]);
            }

            return true;
        }

        private static bool TryResolveWordLinkTarget(
            PdfCore.PdfLogicalLinkAnnotation link,
            PdfWordReadOptions options,
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
                new Dictionary<int, List<string>>());

            public ImportNavigationMap(
                IReadOnlyDictionary<int, string> pageAnchors,
                IReadOnlyDictionary<string, string> namedAnchors,
                IReadOnlyDictionary<int, List<string>> anchorsByPage) {
                PageAnchors = pageAnchors;
                NamedAnchors = namedAnchors;
                AnchorsByPage = anchorsByPage;
            }

            public IReadOnlyDictionary<int, string> PageAnchors { get; }

            public IReadOnlyDictionary<string, string> NamedAnchors { get; }

            private IReadOnlyDictionary<int, List<string>> AnchorsByPage { get; }

            public bool HasAnchorsForPage(int pageNumber) =>
                AnchorsByPage.TryGetValue(pageNumber, out List<string>? anchors) && anchors.Count > 0;

            public IReadOnlyList<string> GetAnchorsForPage(int pageNumber) =>
                AnchorsByPage.TryGetValue(pageNumber, out List<string>? anchors)
                    ? anchors.AsReadOnly()
                    : Array.Empty<string>();
        }
    }
}
