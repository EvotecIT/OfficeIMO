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
                    string metadataSignature = GetHyperlinkMetadataSignature(hyperlink, options);
                    string text = NormalizeFeatureText(hyperlink.InnerText ?? string.Empty);
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

        private static List<ListSnapshot> GetListSnapshots(WordDocument document, WordComparisonOptions options) {
            MainDocumentPart? mainPart = document._wordprocessingDocument.MainDocumentPart;
            if (mainPart == null) {
                return new List<ListSnapshot>();
            }

            var snapshots = new List<ListSnapshot>();
            Numbering? numberingDefinitions = mainPart.NumberingDefinitionsPart?.Numbering;
            foreach (WordFieldInventory.FieldRoot root in WordFieldInventory.EnumerateFieldRoots(mainPart)) {
                foreach (OrderedElement ordered in EnumerateDescendantsWithOrder(root.Root, GetFeatureOrderBase(root.LocationKind))) {
                    if (ordered.Element is not Paragraph paragraph) {
                        continue;
                    }

                    NumberingProperties? numbering = ResolveParagraphNumberingProperties(paragraph, mainPart);
                    if (numbering == null) {
                        continue;
                    }

                    string numberId = numbering.NumberingId?.Val?.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty;
                    string level = numbering.NumberingLevelReference?.Val?.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty;
                    string numberingSignature = GetListNumberingDefinitionSignature(numberingDefinitions, numberId, level, options);
                    string text = NormalizeFeatureText(paragraph.InnerText ?? string.Empty);
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
            Paragraph? previousParagraph = null;
            TableCell? previousCell = null;

            foreach (OpenXmlElement element in root.Descendants()) {
                if (ReferenceEquals(element, bookmarkStart)) {
                    inRange = true;
                    continue;
                }

                if (element is BookmarkEnd bookmarkEnd && bookmarkEnd.Id?.Value == bookmarkId) {
                    break;
                }

                if (inRange && element is Text textElement) {
                    AddRangeTextSeparatorIfNeeded(text, element, ref previousParagraph, ref previousCell);
                    text.Add(textElement.Text);
                }
            }

            return string.Concat(text);
        }

        private static void AddRangeTextSeparatorIfNeeded(List<string> parts, OpenXmlElement current, ref Paragraph? previousParagraph, ref TableCell? previousCell) {
            Paragraph? paragraph = current.Ancestors<Paragraph>().FirstOrDefault();
            TableCell? cell = current.Ancestors<TableCell>().FirstOrDefault();
            if (parts.Count > 0 &&
                (!ReferenceEquals(paragraph, previousParagraph) || !ReferenceEquals(cell, previousCell))) {
                parts.Add(" ");
            }

            previousParagraph = paragraph;
            previousCell = cell;
        }

        private static NumberingProperties? ResolveParagraphNumberingProperties(Paragraph paragraph, MainDocumentPart mainPart) {
            NumberingProperties? directNumbering = paragraph.ParagraphProperties?.NumberingProperties;
            if (directNumbering != null) {
                return directNumbering;
            }

            string? styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (string.IsNullOrWhiteSpace(styleId)) {
                return null;
            }

            Styles? styles = mainPart.StyleDefinitionsPart?.Styles;
            if (styles == null) {
                return null;
            }

            var visited = new HashSet<string>(StringComparer.Ordinal);
            while (!string.IsNullOrWhiteSpace(styleId)) {
                string currentStyleId = styleId!;
                if (!visited.Add(currentStyleId)) {
                    break;
                }

                Style? style = styles.Elements<Style>()
                    .FirstOrDefault(item => string.Equals(item.StyleId?.Value, currentStyleId, StringComparison.Ordinal));
                NumberingProperties? numbering = style?.StyleParagraphProperties?.NumberingProperties;
                if (numbering != null) {
                    return numbering;
                }

                styleId = style?.BasedOn?.Val?.Value;
            }

            return null;
        }

        private static string GetBookmarkLocationSignature(BookmarkStart bookmarkStart, WordComparisonOptions options) {
            Paragraph? paragraph = bookmarkStart.Ancestors<Paragraph>().FirstOrDefault();
            if (paragraph == null) {
                return GetStableElementPath(bookmarkStart);
            }

            string paragraphText = NormalizeComparisonText(NormalizeFeatureText(paragraph.InnerText ?? string.Empty), options);
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

        private static int GetParagraphTextOccurrence(Paragraph paragraph, string normalizedText, WordComparisonOptions options) {
            OpenXmlElement root = paragraph.Ancestors().LastOrDefault() ?? paragraph;
            int occurrence = 0;
            foreach (Paragraph candidate in root.Descendants<Paragraph>()) {
                if (ReferenceEquals(candidate, paragraph)) {
                    return occurrence;
                }

                string candidateText = NormalizeComparisonText(NormalizeFeatureText(candidate.InnerText ?? string.Empty), options);
                if (string.Equals(candidateText, normalizedText, StringComparison.Ordinal)) {
                    occurrence++;
                }
            }

            return occurrence;
        }

        private static string GetHyperlinkTarget(OpenXmlPart? part, string relationshipId) {
            if (part == null || string.IsNullOrWhiteSpace(relationshipId)) {
                return string.Empty;
            }

            HyperlinkRelationship? relationship = part.HyperlinkRelationships
                .FirstOrDefault(item => string.Equals(item.Id, relationshipId, StringComparison.Ordinal));
            return relationship?.Uri.ToString() ?? string.Empty;
        }

        private static string GetHyperlinkMetadataSignature(Hyperlink hyperlink, WordComparisonOptions options) {
            return string.Join(
                "|",
                NormalizeComparisonText(hyperlink.DocLocation?.Value ?? string.Empty, options),
                hyperlink.History?.Value == null ? string.Empty : (hyperlink.History.Value ? "history=true" : "history=false"),
                NormalizeComparisonText(hyperlink.Tooltip?.Value ?? string.Empty, options),
                NormalizeComparisonText(hyperlink.TargetFrame?.Value ?? string.Empty, options));
        }

        private static string GetListNumberingDefinitionSignature(Numbering? numbering, string numberId, string level, WordComparisonOptions options) {
            if (numbering == null || string.IsNullOrWhiteSpace(numberId)) {
                return string.Empty;
            }

            NumberingInstance? instance = numbering.Elements<NumberingInstance>()
                .FirstOrDefault(item => string.Equals(
                    item.NumberID?.Value.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    numberId,
                    StringComparison.Ordinal));
            if (instance == null) {
                return string.Empty;
            }

            int? abstractNumberId = instance.AbstractNumId?.Val?.Value;
            AbstractNum? abstractNum = abstractNumberId.HasValue
                ? numbering.Elements<AbstractNum>().FirstOrDefault(item => item.AbstractNumberId?.Value == abstractNumberId.Value)
                : null;

            OpenXmlElement instanceClone = instance.CloneNode(true);
            OpenXmlElement? abstractClone = abstractNum?.CloneNode(true);
            NormalizeNumberingVolatileIds(abstractClone);
            if (!options.CompareGeneratedIds) {
                NormalizeNumberingGeneratedIds(instanceClone, abstractClone);
            }

            Level? resolvedLevel = abstractClone?
                .Elements<Level>()
                .FirstOrDefault(item => string.Equals(
                    item.LevelIndex?.Value.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    level,
                    StringComparison.Ordinal));
            LevelOverride? resolvedOverride = instanceClone
                .Elements<LevelOverride>()
                .FirstOrDefault(item => string.Equals(
                    item.LevelIndex?.Value.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    level,
                    StringComparison.Ordinal));

            return NormalizeComparisonText(string.Join(
                "|",
                instanceClone.OuterXml,
                abstractClone?.OuterXml ?? string.Empty,
                resolvedLevel?.OuterXml ?? string.Empty,
                resolvedOverride?.OuterXml ?? string.Empty), options);
        }

        private static void NormalizeNumberingGeneratedIds(OpenXmlElement instanceClone, OpenXmlElement? abstractClone) {
            if (instanceClone is NumberingInstance numberingInstance) {
                numberingInstance.NumberID = 0;
                AbstractNumId? abstractNumId = numberingInstance.GetFirstChild<AbstractNumId>();
                if (abstractNumId != null) {
                    abstractNumId.Val = 0;
                }
            }

            if (abstractClone is AbstractNum abstractNum) {
                abstractNum.AbstractNumberId = 0;
            }
        }

        private static void NormalizeNumberingVolatileIds(OpenXmlElement? abstractClone) {
            if (abstractClone is not AbstractNum abstractNum) {
                return;
            }

            abstractNum.RemoveAllChildren<Nsid>();
            abstractNum.RemoveAllChildren<TemplateCode>();
            foreach (Level level in abstractNum.Descendants<Level>()) {
                level.TemplateCode = null;
            }
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
