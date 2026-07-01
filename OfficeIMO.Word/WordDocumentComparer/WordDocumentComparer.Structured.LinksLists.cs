using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public static partial class WordDocumentComparer {
        private static void AnalyzeBookmarks(WordDocument source, WordDocument target, WordComparisonResult result, WordComparisonOptions options) {
            AddFeatureRangeFindings(
                GetBookmarkSnapshots(source, options),
                GetBookmarkSnapshots(target, options),
                WordComparisonScope.Bookmark,
                "bookmark",
                "Bookmark",
                result);
        }

        private static void AnalyzeHyperlinks(WordDocument source, WordDocument target, WordComparisonResult result, WordComparisonOptions options) {
            AddFeatureRangeFindings(
                GetHyperlinkSnapshots(source, options),
                GetHyperlinkSnapshots(target, options),
                WordComparisonScope.Hyperlink,
                "hyperlink",
                "Hyperlink",
                result);
        }

        private static void AnalyzeLists(WordDocument source, WordDocument target, WordComparisonResult result, WordComparisonOptions options) {
            AddFeatureRangeFindings(
                GetListSnapshots(source, options),
                GetListSnapshots(target, options),
                WordComparisonScope.List,
                "list",
                "List item",
                result);
        }

        private static List<BookmarkSnapshot> GetBookmarkSnapshots(WordDocument document, WordComparisonOptions options) {
            MainDocumentPart? mainPart = document._wordprocessingDocument.MainDocumentPart;
            if (mainPart == null) {
                return new List<BookmarkSnapshot>();
            }

            var snapshots = new List<BookmarkSnapshot>();
            foreach (WordFieldInventory.FieldRoot root in WordFieldInventory.EnumerateFieldRoots(mainPart)) {
                foreach (OrderedElement ordered in EnumerateDescendantsWithOrder(root.Root, GetFeatureOrderBase(root.LocationKind))) {
                    if (ordered.Element is not BookmarkStart bookmarkStart) {
                        continue;
                    }

                    string name = bookmarkStart.Name?.Value ?? string.Empty;
                    string id = bookmarkStart.Id?.Value ?? string.Empty;
                    string text = GetBookmarkDisplayText(bookmarkStart);
                    string displayText = "name=" + name + "; id=" + id + "; text=" + text;
                    string matchSignature = string.Join(
                        "|",
                        NormalizeComparisonText(name, options),
                        NormalizeComparisonText(text, options),
                        IsInTableFeature(bookmarkStart) ? "table" : string.Empty,
                        IsInTextBoxFeature(bookmarkStart) ? "text-box" : string.Empty);
                    string signature = string.Join(
                        "|",
                        NormalizeComparisonText(name, options),
                        options.CompareGeneratedIds ? NormalizeComparisonText(id, options) : string.Empty,
                        NormalizeComparisonText(text, options),
                        root.LocationKind.ToString(),
                        root.PartUri,
                        GetStableElementPath(bookmarkStart),
                        IsInTableFeature(bookmarkStart) ? "table" : string.Empty,
                        IsInTextBoxFeature(bookmarkStart) ? "text-box" : string.Empty);

                    snapshots.Add(new BookmarkSnapshot(
                        snapshots.Count,
                        GetFeatureMatchKey(root, matchSignature, bookmarkStart.LocalName),
                        signature,
                        displayText,
                        JoinFeatureLocation(
                            root.LocationKind.ToString(),
                            root.PartUri,
                            FeatureLocation("bookmark", snapshots.Count),
                            IsInTableFeature(bookmarkStart) ? "table" : string.Empty,
                            IsInContentControlFeature(bookmarkStart) ? "content-control" : string.Empty,
                            IsInTextBoxFeature(bookmarkStart) ? "text-box" : string.Empty),
                        ordered.DocumentOrder));
                }
            }

            return snapshots;
        }

        private static List<HyperlinkSnapshot> GetHyperlinkSnapshots(WordDocument document, WordComparisonOptions options) {
            MainDocumentPart? mainPart = document._wordprocessingDocument.MainDocumentPart;
            if (mainPart == null) {
                return new List<HyperlinkSnapshot>();
            }

            Dictionary<string, OpenXmlPart> partsByUri = GetMainPartDescendantsByUri(mainPart);
            var snapshots = new List<HyperlinkSnapshot>();
            foreach (WordFieldInventory.FieldRoot root in WordFieldInventory.EnumerateFieldRoots(mainPart)) {
                partsByUri.TryGetValue(root.PartUri, out OpenXmlPart? part);
                foreach (OrderedElement ordered in EnumerateDescendantsWithOrder(root.Root, GetFeatureOrderBase(root.LocationKind))) {
                    if (ordered.Element is not Hyperlink hyperlink) {
                        continue;
                    }

                    string relationshipId = hyperlink.Id?.Value ?? string.Empty;
                    string anchor = hyperlink.Anchor?.Value ?? string.Empty;
                    string target = GetHyperlinkTarget(part, relationshipId);
                    string text = NormalizeFeatureText(hyperlink.InnerText ?? string.Empty);
                    string displayText = "text=" + text + "; anchor=" + anchor + "; target=" + target + "; relationshipId=" + relationshipId;
                    string matchSignature = string.Join(
                        "|",
                        NormalizeComparisonText(text, options),
                        NormalizeComparisonText(anchor, options),
                        NormalizeComparisonText(target, options),
                        IsInTableFeature(hyperlink) ? "table" : string.Empty,
                        IsInTextBoxFeature(hyperlink) ? "text-box" : string.Empty);
                    string signature = string.Join(
                        "|",
                        NormalizeComparisonText(text, options),
                        NormalizeComparisonText(anchor, options),
                        NormalizeComparisonText(target, options),
                        options.CompareGeneratedIds ? NormalizeComparisonText(relationshipId, options) : string.Empty,
                        IsInTableFeature(hyperlink) ? "table" : string.Empty,
                        IsInTextBoxFeature(hyperlink) ? "text-box" : string.Empty);

                    snapshots.Add(new HyperlinkSnapshot(
                        snapshots.Count,
                        GetFeatureMatchKey(root, matchSignature, hyperlink.LocalName),
                        signature,
                        displayText,
                        JoinFeatureLocation(
                            root.LocationKind.ToString(),
                            root.PartUri,
                            FeatureLocation("hyperlink", snapshots.Count),
                            IsInTableFeature(hyperlink) ? "table" : string.Empty,
                            IsInContentControlFeature(hyperlink) ? "content-control" : string.Empty,
                            IsInTextBoxFeature(hyperlink) ? "text-box" : string.Empty),
                        ordered.DocumentOrder));
                }
            }

            return snapshots;
        }

        private static List<ListSnapshot> GetListSnapshots(WordDocument document, WordComparisonOptions options) {
            MainDocumentPart? mainPart = document._wordprocessingDocument.MainDocumentPart;
            if (mainPart == null) {
                return new List<ListSnapshot>();
            }

            var snapshots = new List<ListSnapshot>();
            foreach (WordFieldInventory.FieldRoot root in WordFieldInventory.EnumerateFieldRoots(mainPart)) {
                foreach (OrderedElement ordered in EnumerateDescendantsWithOrder(root.Root, GetFeatureOrderBase(root.LocationKind))) {
                    if (ordered.Element is not Paragraph paragraph) {
                        continue;
                    }

                    NumberingProperties? numbering = paragraph.ParagraphProperties?.NumberingProperties;
                    if (numbering == null) {
                        continue;
                    }

                    string numberId = numbering.NumberingId?.Val?.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty;
                    string level = numbering.NumberingLevelReference?.Val?.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty;
                    string text = NormalizeFeatureText(paragraph.InnerText ?? string.Empty);
                    string displayText = "numId=" + numberId + "; level=" + level + "; text=" + text;
                    string matchSignature = string.Join(
                        "|",
                        NormalizeComparisonText(level, options),
                        NormalizeComparisonText(text, options),
                        IsInTableFeature(paragraph) ? "table" : string.Empty,
                        IsInContentControlFeature(paragraph) ? "content-control" : string.Empty,
                        IsInTextBoxFeature(paragraph) ? "text-box" : string.Empty);
                    string signature = string.Join(
                        "|",
                        options.CompareGeneratedIds ? NormalizeComparisonText(numberId, options) : string.Empty,
                        NormalizeComparisonText(level, options),
                        NormalizeComparisonText(text, options),
                        IsInTableFeature(paragraph) ? "table" : string.Empty,
                        IsInContentControlFeature(paragraph) ? "content-control" : string.Empty,
                        IsInTextBoxFeature(paragraph) ? "text-box" : string.Empty);

                    snapshots.Add(new ListSnapshot(
                        snapshots.Count,
                        GetFeatureMatchKey(root, matchSignature, paragraph.LocalName),
                        signature,
                        displayText,
                        JoinFeatureLocation(
                            root.LocationKind.ToString(),
                            root.PartUri,
                            FeatureLocation("list", snapshots.Count),
                            IsInTableFeature(paragraph) ? "table" : string.Empty,
                            IsInContentControlFeature(paragraph) ? "content-control" : string.Empty,
                            IsInTextBoxFeature(paragraph) ? "text-box" : string.Empty),
                        ordered.DocumentOrder));
                }
            }

            return snapshots;
        }

        private static string GetBookmarkDisplayText(BookmarkStart bookmarkStart) {
            string rangeText = GetBookmarkRangeText(bookmarkStart);
            if (!string.IsNullOrEmpty(rangeText)) {
                return NormalizeFeatureText(rangeText);
            }

            Paragraph? paragraph = bookmarkStart.Ancestors<Paragraph>().FirstOrDefault();
            return paragraph == null ? string.Empty : NormalizeFeatureText(paragraph.InnerText ?? string.Empty);
        }

        private static string GetBookmarkRangeText(BookmarkStart bookmarkStart) {
            string? bookmarkId = bookmarkStart.Id?.Value;
            if (string.IsNullOrWhiteSpace(bookmarkId)) {
                return string.Empty;
            }

            OpenXmlElement root = bookmarkStart.Ancestors().LastOrDefault() ?? bookmarkStart;
            var text = new List<string>();
            bool inRange = false;

            foreach (OpenXmlElement element in root.Descendants()) {
                if (ReferenceEquals(element, bookmarkStart)) {
                    inRange = true;
                    continue;
                }

                if (element is BookmarkEnd bookmarkEnd && bookmarkEnd.Id?.Value == bookmarkId) {
                    break;
                }

                if (inRange && element is Text textElement) {
                    text.Add(textElement.Text);
                }
            }

            return string.Concat(text);
        }

        private static string GetHyperlinkTarget(OpenXmlPart? part, string relationshipId) {
            if (part == null || string.IsNullOrWhiteSpace(relationshipId)) {
                return string.Empty;
            }

            HyperlinkRelationship? relationship = part.HyperlinkRelationships
                .FirstOrDefault(item => string.Equals(item.Id, relationshipId, StringComparison.Ordinal));
            return relationship?.Uri.ToString() ?? string.Empty;
        }

        private static Dictionary<string, OpenXmlPart> GetMainPartDescendantsByUri(MainDocumentPart mainPart) {
            var parts = new Dictionary<string, OpenXmlPart>(StringComparer.OrdinalIgnoreCase) {
                [mainPart.Uri.ToString()] = mainPart
            };

            foreach (OpenXmlPart part in mainPart.Parts.Select(pair => pair.OpenXmlPart)) {
                parts[part.Uri.ToString()] = part;
            }

            return parts;
        }

        private static string GetFeatureMatchKey(WordFieldInventory.FieldRoot root, string signature, string elementName) {
            return string.Join(
                "|",
                root.LocationKind.ToString(),
                root.PartUri,
                elementName,
                signature);
        }

        private static bool IsInContentControlFeature(OpenXmlElement element) => element.Ancestors<SdtElement>().Any();

        private sealed class BookmarkSnapshot : IFeatureSnapshot {
            internal BookmarkSnapshot(int index, string matchKey, string signature, string displayText, string detailedLocation, int documentOrder) {
                Index = index;
                MatchKey = matchKey;
                Signature = signature;
                DisplayText = displayText;
                DetailedLocation = detailedLocation;
                DocumentOrder = documentOrder;
            }

            public int Index { get; }

            public string MatchKey { get; }

            public string Signature { get; }

            public string DisplayText { get; }

            public string DetailedLocation { get; }

            public int DocumentOrder { get; }
        }

        private sealed class HyperlinkSnapshot : IFeatureSnapshot {
            internal HyperlinkSnapshot(int index, string matchKey, string signature, string displayText, string detailedLocation, int documentOrder) {
                Index = index;
                MatchKey = matchKey;
                Signature = signature;
                DisplayText = displayText;
                DetailedLocation = detailedLocation;
                DocumentOrder = documentOrder;
            }

            public int Index { get; }

            public string MatchKey { get; }

            public string Signature { get; }

            public string DisplayText { get; }

            public string DetailedLocation { get; }

            public int DocumentOrder { get; }
        }

        private sealed class ListSnapshot : IFeatureSnapshot {
            internal ListSnapshot(int index, string matchKey, string signature, string displayText, string detailedLocation, int documentOrder) {
                Index = index;
                MatchKey = matchKey;
                Signature = signature;
                DisplayText = displayText;
                DetailedLocation = detailedLocation;
                DocumentOrder = documentOrder;
            }

            public int Index { get; }

            public string MatchKey { get; }

            public string Signature { get; }

            public string DisplayText { get; }

            public string DetailedLocation { get; }

            public int DocumentOrder { get; }
        }
    }
}
