using OfficeIMO.Word.OpenXml;

namespace OfficeIMO.Word {
    public static partial class WordDocumentComparer {
        private static void AnalyzeBookmarks(WordPackageView source, WordPackageView target, WordComparisonResult result, WordComparisonOptions options) {
            AddFeatureRangeFindings(
                GetBookmarkSnapshots(source, options),
                GetBookmarkSnapshots(target, options),
                WordComparisonScope.Bookmark,
                "bookmark",
                "Bookmark",
                result);
        }

        private static void AnalyzeHyperlinks(WordPackageView source, WordPackageView target, WordComparisonResult result, WordComparisonOptions options) {
            AddFeatureRangeFindings(
                GetHyperlinkSnapshots(source, options),
                GetHyperlinkSnapshots(target, options),
                WordComparisonScope.Hyperlink,
                "hyperlink",
                "Hyperlink",
                result);
        }

        private static void AnalyzeLists(
            WordPackageView source,
            WordPackageView target,
            WordComparisonResult result,
            WordComparisonOptions options,
            ParagraphNumberingStyleCatalogCache numberingStyleCatalogs) {
            AddFeatureRangeFindings(
                GetListSnapshots(source, options, numberingStyleCatalogs),
                GetListSnapshots(target, options, numberingStyleCatalogs),
                WordComparisonScope.List,
                "list",
                "List item",
                result);
        }

        private static List<BookmarkSnapshot> GetBookmarkSnapshots(WordPackageView document, WordComparisonOptions options) {
            WordPartView? mainPart = document.Main;
            if (mainPart == null) {
                return new List<BookmarkSnapshot>();
            }

            var snapshots = new List<BookmarkSnapshot>();
            foreach (WordFieldInventory.NodeFieldRoot root in WordFieldInventory.EnumerateFieldRoots(mainPart)) {
                foreach (OrderedElement ordered in EnumerateDescendantsWithOrder(root.Root, GetFeatureOrderBase(root.LocationKind))) {
                    if (!ordered.Element.Is(W, "bookmarkStart")) {
                        continue;
                    }

                    OfficeXmlNode bookmarkStart = ordered.Element;
                    string name = bookmarkStart.Attribute(W, "name") ?? string.Empty;
                    string id = bookmarkStart.Attribute(W, "id") ?? string.Empty;
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
                        GetBookmarkLocationSignature(bookmarkStart, options),
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

        private static List<HyperlinkSnapshot> GetHyperlinkSnapshots(WordPackageView document, WordComparisonOptions options) {
            WordPartView? mainPart = document.Main;
            if (mainPart == null) {
                return new List<HyperlinkSnapshot>();
            }

            Dictionary<string, WordPartView> partsByUri = GetMainPartDescendantsByUri(mainPart);
            var snapshots = new List<HyperlinkSnapshot>();
            foreach (WordFieldInventory.NodeFieldRoot root in WordFieldInventory.EnumerateFieldRoots(mainPart)) {
                partsByUri.TryGetValue(root.PartUri, out WordPartView? part);
                foreach (OrderedElement ordered in EnumerateDescendantsWithOrder(root.Root, GetFeatureOrderBase(root.LocationKind))) {
                    if (!ordered.Element.Is(W, "hyperlink")) {
                        continue;
                    }

                    OfficeXmlNode hyperlink = ordered.Element;
                    string relationshipId = hyperlink.Attribute(OfficeIMO.Core.Internal.OpenXml.OfficeXmlNamespaces.Relationships, "id") ?? string.Empty;
                    string anchor = hyperlink.Attribute(W, "anchor") ?? string.Empty;
                    string target = GetHyperlinkTarget(part, relationshipId);
                    string metadataSignature = GetHyperlinkMetadataSignature(hyperlink, options);
                    string text = NormalizeFeatureText(InnerText(hyperlink));
                    string displayText = "text=" + text + "; anchor=" + anchor + "; target=" + target + "; relationshipId=" + relationshipId + "; metadata=" + metadataSignature;
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
                        metadataSignature,
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

        private static List<ListSnapshot> GetListSnapshots(
            WordPackageView document,
            WordComparisonOptions options,
            ParagraphNumberingStyleCatalogCache numberingStyleCatalogs) {
            WordPartView? mainPart = document.Main;
            if (mainPart == null) {
                return new List<ListSnapshot>();
            }

            var snapshots = new List<ListSnapshot>();
            WordPartView? numberingDefinitions = mainPart.FirstPart(WordPackageView.NumberingRelationshipType);
            ParagraphNumberingStyleCatalog styleCatalog = numberingStyleCatalogs.GetOrCreate(mainPart);
            var numberingSignatures = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (WordFieldInventory.NodeFieldRoot root in WordFieldInventory.EnumerateFieldRoots(mainPart)) {
                foreach (OrderedElement ordered in EnumerateDescendantsWithOrder(root.Root, GetFeatureOrderBase(root.LocationKind))) {
                    if (!ordered.Element.Is(W, "p")) {
                        continue;
                    }

                    OfficeXmlNode paragraph = ordered.Element;
                    OfficeXmlNode numbering = ResolveParagraphNumberingProperties(paragraph, styleCatalog);
                    if (numbering.IsNone) {
                        continue;
                    }

                    string numberId = Int32Text(ChildVal(numbering, "numId")) ?? string.Empty;
                    string level = Int32Text(ChildVal(numbering, "ilvl")) ?? string.Empty;
                    string signatureKey = numberId + "\u001f" + level;
                    if (!numberingSignatures.TryGetValue(signatureKey, out string? numberingSignature)) {
                        numberingSignature = GetListNumberingDefinitionSignature(numberingDefinitions, numberId, level, options);
                        numberingSignatures.Add(signatureKey, numberingSignature);
                    }

                    string text = NormalizeFeatureText(InnerText(paragraph));
                    string displayText = "numId=" + numberId + "; level=" + level + "; text=" + text + "; numbering=" + numberingSignature;
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
                        numberingSignature,
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

        private static string GetBookmarkDisplayText(OfficeXmlNode bookmarkStart) {
            string rangeText = GetBookmarkRangeText(bookmarkStart);
            if (!string.IsNullOrEmpty(rangeText)) {
                return NormalizeFeatureText(rangeText);
            }

            OfficeXmlNode paragraph = AncestorW(bookmarkStart, "p");
            return paragraph.IsNone ? string.Empty : NormalizeFeatureText(InnerText(paragraph));
        }

        private static string GetBookmarkRangeText(OfficeXmlNode bookmarkStart) {
            string? bookmarkId = bookmarkStart.Attribute(W, "id");
            if (string.IsNullOrWhiteSpace(bookmarkId)) {
                return string.Empty;
            }

            OfficeXmlNode root = WordXmlNodes.Root(bookmarkStart);
            var text = new List<string>();
            bool inRange = false;
            OfficeXmlNode previousParagraph = default;
            OfficeXmlNode previousCell = default;

            foreach (OfficeXmlNode element in TypedDescendants(root)) {
                if (element == bookmarkStart) {
                    inRange = true;
                    continue;
                }

                if (element.Is(W, "bookmarkEnd") && element.Attribute(W, "id") == bookmarkId) {
                    break;
                }

                if (inRange && element.Is(W, "t")) {
                    AddRangeTextSeparatorIfNeeded(text, element, ref previousParagraph, ref previousCell);
                    text.Add(element.InnerText);
                }
            }

            return string.Concat(text);
        }

        private static void AddRangeTextSeparatorIfNeeded(List<string> parts, OfficeXmlNode current, ref OfficeXmlNode previousParagraph, ref OfficeXmlNode previousCell) {
            OfficeXmlNode paragraph = AncestorW(current, "p");
            OfficeXmlNode cell = AncestorW(current, "tc");
            if (parts.Count > 0 &&
                (paragraph != previousParagraph || cell != previousCell)) {
                parts.Add(" ");
            }

            previousParagraph = paragraph;
            previousCell = cell;
        }

        private static string GetBookmarkLocationSignature(OfficeXmlNode bookmarkStart, WordComparisonOptions options) {
            OfficeXmlNode paragraph = AncestorW(bookmarkStart, "p");
            if (paragraph.IsNone) {
                return GetStableElementPath(bookmarkStart);
            }

            string paragraphText = NormalizeComparisonText(NormalizeFeatureText(InnerText(paragraph)), options);
            int occurrence = GetParagraphTextOccurrence(paragraph, paragraphText, options);
            return string.Join(
                "|",
                "paragraph",
                paragraphText,
                occurrence.ToString(System.Globalization.CultureInfo.InvariantCulture),
                IsInTableFeature(paragraph) ? "table" : string.Empty,
                IsInContentControlFeature(paragraph) ? "content-control" : string.Empty,
                IsInTextBoxFeature(paragraph) ? "text-box" : string.Empty);
        }

        private static int GetParagraphTextOccurrence(OfficeXmlNode paragraph, string normalizedText, WordComparisonOptions options) {
            OfficeXmlNode root = WordXmlNodes.Root(paragraph);
            int occurrence = 0;
            foreach (OfficeXmlNode candidate in DescendantsW(root, "p")) {
                if (candidate == paragraph) {
                    return occurrence;
                }

                string candidateText = NormalizeComparisonText(NormalizeFeatureText(InnerText(candidate)), options);
                if (string.Equals(candidateText, normalizedText, StringComparison.Ordinal)) {
                    occurrence++;
                }
            }

            return occurrence;
        }

        private static string GetHyperlinkTarget(WordPartView? part, string relationshipId) {
            if (part == null || string.IsNullOrWhiteSpace(relationshipId)) {
                return string.Empty;
            }

            return part.HyperlinkTarget(relationshipId) ?? string.Empty;
        }

        private static string GetHyperlinkMetadataSignature(OfficeXmlNode hyperlink, WordComparisonOptions options) {
            bool? history = ParseOnOff(hyperlink.Attribute(W, "history"));
            return string.Join(
                "|",
                NormalizeComparisonText(hyperlink.Attribute(W, "docLocation") ?? string.Empty, options),
                history == null ? string.Empty : (history.Value ? "history=true" : "history=false"),
                NormalizeComparisonText(hyperlink.Attribute(W, "tooltip") ?? string.Empty, options),
                NormalizeComparisonText(hyperlink.Attribute(W, "tgtFrame") ?? string.Empty, options));
        }

        private static string GetListNumberingDefinitionSignature(WordPartView? numberingPart, string numberId, string level, WordComparisonOptions options) {
            // The report text carries the definitions' markup as the SDK wrote it, so that markup is written the same way.
            return NormalizeComparisonText(WordPackageView.SdkListNumberingSignature(numberingPart, numberId, level, options.CompareGeneratedIds), options);
        }

        private static Dictionary<string, WordPartView> GetMainPartDescendantsByUri(WordPartView mainPart) {
            var parts = new Dictionary<string, WordPartView>(StringComparer.OrdinalIgnoreCase) {
                [mainPart.PartName] = mainPart
            };

            foreach (WordPartView part in mainPart.ChildParts()) {
                parts[part.PartName] = part;
            }

            return parts;
        }

        private static string GetFeatureMatchKey(WordFieldInventory.NodeFieldRoot root, string signature, string elementName) {
            return string.Join(
                "|",
                root.LocationKind.ToString(),
                root.PartUri,
                elementName,
                signature);
        }

        private static bool IsInContentControlFeature(OfficeXmlNode element) => HasAncestorW(element, "sdt");

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

            ulong IComparisonFingerprint.ComparisonFingerprint => GetOrdinalTextFingerprint(MatchKey);
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

            ulong IComparisonFingerprint.ComparisonFingerprint => GetOrdinalTextFingerprint(MatchKey);
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

            ulong IComparisonFingerprint.ComparisonFingerprint => GetOrdinalTextFingerprint(MatchKey);
        }
    }
}
