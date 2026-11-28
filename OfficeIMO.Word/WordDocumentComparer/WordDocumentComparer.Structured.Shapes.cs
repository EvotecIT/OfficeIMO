using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using OfficeIMO.Word.OpenXml;

namespace OfficeIMO.Word {
    public static partial class WordDocumentComparer {
        private const string WordprocessingShapeUri = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
        private const string WordprocessingGroupUri = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";

        private static void AnalyzeShapes(
            WordPackageView source,
            WordPackageView target,
            WordComparisonResult result,
            WordComparisonOptions options) {
            IReadOnlyList<ShapeSnapshot> sourceShapes = GetShapeSnapshots(source, options);
            IReadOnlyList<ShapeSnapshot> targetShapes = GetShapeSnapshots(target, options);
            IReadOnlyList<MatchedIndexPair> matches = FindMatchingIndexes(
                sourceShapes,
                targetShapes,
                ShapeSnapshotContentComparer.Instance);

            int sourceStart = 0;
            int targetStart = 0;
            foreach (MatchedIndexPair match in matches) {
                AddShapeRangeFindings(sourceShapes, targetShapes, sourceStart, match.SourceIndex, targetStart, match.TargetIndex, result);
                AddMatchedShapeLayoutFinding(sourceShapes[match.SourceIndex], targetShapes[match.TargetIndex], match, result);
                sourceStart = match.SourceIndex + 1;
                targetStart = match.TargetIndex + 1;
            }
            AddShapeRangeFindings(sourceShapes, targetShapes, sourceStart, sourceShapes.Count, targetStart, targetShapes.Count, result);
        }

        private static void AddShapeRangeFindings(
            IReadOnlyList<ShapeSnapshot> source,
            IReadOnlyList<ShapeSnapshot> target,
            int sourceStart,
            int sourceEnd,
            int targetStart,
            int targetEnd,
            WordComparisonResult result) {
            int sourceIndex = sourceStart;
            int targetIndex = targetStart;
            while (sourceIndex < sourceEnd && targetIndex < targetEnd) {
                if (targetEnd - targetIndex > sourceEnd - sourceIndex &&
                    targetIndex + 1 < targetEnd &&
                    GetShapeSimilarity(source[sourceIndex], target[targetIndex + 1]) >
                    GetShapeSimilarity(source[sourceIndex], target[targetIndex])) {
                    AddInsertedShapeFinding(target, targetIndex++, result);
                    continue;
                }
                if (sourceEnd - sourceIndex > targetEnd - targetIndex &&
                    sourceIndex + 1 < sourceEnd &&
                    GetShapeSimilarity(source[sourceIndex + 1], target[targetIndex]) >
                    GetShapeSimilarity(source[sourceIndex], target[targetIndex])) {
                    AddDeletedShapeFinding(source, sourceIndex++, result);
                    continue;
                }
                if (!string.Equals(source[sourceIndex].PartKey, target[targetIndex].PartKey, StringComparison.Ordinal)) {
                    AddDeletedShapeFinding(source, sourceIndex++, result);
                    AddInsertedShapeFinding(target, targetIndex++, result);
                    continue;
                }

                ShapeSnapshot sourceShape = source[sourceIndex];
                ShapeSnapshot targetShape = target[targetIndex];
                result.Add(new WordComparisonFinding(
                    WordComparisonScope.Shape,
                    WordComparisonChangeKind.Modified,
                    ShapeLocation(targetIndex),
                    sourceIndex,
                    targetIndex,
                    sourceShape.DisplayText,
                    targetShape.DisplayText,
                    "DrawingML shape content changed.",
                    targetShape.DetailedLocation),
                    targetShape.DocumentOrder);
                sourceIndex++;
                targetIndex++;
            }
            while (targetIndex < targetEnd) AddInsertedShapeFinding(target, targetIndex++, result);
            while (sourceIndex < sourceEnd) AddDeletedShapeFinding(source, sourceIndex++, result);
        }

        private static void AddMatchedShapeLayoutFinding(
            ShapeSnapshot source,
            ShapeSnapshot target,
            MatchedIndexPair match,
            WordComparisonResult result) {
            if (string.Equals(source.LayoutFingerprint, target.LayoutFingerprint, StringComparison.Ordinal) &&
                string.Equals(source.PartKey, target.PartKey, StringComparison.Ordinal)) return;
            string message = string.Equals(source.PartKey, target.PartKey, StringComparison.Ordinal)
                ? "DrawingML shape placement or frame layout changed."
                : "DrawingML shape moved to a different document part; this is a structural position finding, not Word move-range semantics.";
            result.Add(new WordComparisonFinding(
                WordComparisonScope.Shape,
                WordComparisonChangeKind.Modified,
                ShapeLocation(match.TargetIndex),
                match.SourceIndex,
                match.TargetIndex,
                source.DisplayText,
                target.DisplayText,
                message,
                target.DetailedLocation),
                target.DocumentOrder);
        }

        private static void AddInsertedShapeFinding(IReadOnlyList<ShapeSnapshot> shapes, int index, WordComparisonResult result) {
            ShapeSnapshot shape = shapes[index];
            result.Add(new WordComparisonFinding(
                WordComparisonScope.Shape,
                WordComparisonChangeKind.Inserted,
                ShapeLocation(index),
                null,
                index,
                null,
                shape.DisplayText,
                "DrawingML shape inserted.",
                shape.DetailedLocation),
                shape.DocumentOrder);
        }

        private static void AddDeletedShapeFinding(IReadOnlyList<ShapeSnapshot> shapes, int index, WordComparisonResult result) {
            ShapeSnapshot shape = shapes[index];
            result.Add(new WordComparisonFinding(
                WordComparisonScope.Shape,
                WordComparisonChangeKind.Deleted,
                ShapeLocation(index),
                index,
                null,
                shape.DisplayText,
                null,
                "DrawingML shape deleted.",
                shape.DetailedLocation),
                shape.DocumentOrder);
        }

        private static List<ShapeSnapshot> GetShapeSnapshots(WordPackageView document, WordComparisonOptions options) {
            var snapshots = new List<ShapeSnapshot>();
            WordPartView? mainPart = document.Main;
            AddShapeSnapshots(snapshots, BodyOf(mainPart), BodyPartKey, BodyPartOrderBase, options);
            if (mainPart == null) return snapshots;

            int headerIndex = 0;
            foreach (KeyValuePair<WordPartView, string> part in CreateOrderedHeaderPartKeys(mainPart)) {
                AddShapeSnapshots(snapshots, HeaderRoot(part.Key), part.Value, HeaderPartOrderBase + (headerIndex++ * RelatedPartOrderStride), options);
            }
            int footerIndex = 0;
            foreach (KeyValuePair<WordPartView, string> part in CreateOrderedFooterPartKeys(mainPart)) {
                AddShapeSnapshots(snapshots, FooterRoot(part.Key), part.Value, FooterPartOrderBase + (footerIndex++ * RelatedPartOrderStride), options);
            }
            List<OfficeXmlNode> footnotes = GetReferencedFootnotes(mainPart);
            for (int index = 0; index < footnotes.Count; index++) {
                string id = GetNotePartKeyId(footnotes[index], index);
                AddShapeSnapshots(snapshots, footnotes[index], FootnotePartKeyPrefix + id, FootnotePartOrderBase + (index * RelatedPartOrderStride), options);
            }
            List<OfficeXmlNode> endnotes = GetReferencedEndnotes(mainPart);
            for (int index = 0; index < endnotes.Count; index++) {
                string id = GetNotePartKeyId(endnotes[index], index);
                AddShapeSnapshots(snapshots, endnotes[index], EndnotePartKeyPrefix + id, EndnotePartOrderBase + (index * RelatedPartOrderStride), options);
            }
            return snapshots;
        }

        private static void AddShapeSnapshots(
            List<ShapeSnapshot> snapshots,
            OfficeXmlNode container,
            string partKey,
            int orderBase,
            WordComparisonOptions options) {
            if (container.IsNone) return;
            int partIndex = 0;
            foreach (OrderedElement ordered in EnumerateDescendantsWithOrder(container, orderBase)) {
                if (!ordered.Element.Is(W, "drawing") || !TryGetShapeGraphicData(ordered.Element, out OfficeXmlNode graphicData)) continue;
                snapshots.Add(CreateShapeSnapshot(ordered.Element, graphicData, partKey, partIndex++, ordered.DocumentOrder, options));
            }
        }

        private static bool TryGetShapeGraphicData(OfficeXmlNode drawing, out OfficeXmlNode graphicData) {
            graphicData = DescendantsOf(drawing, Dml, "graphicData").FirstOrDefault(data => {
                string uri = data.Attribute("uri") ?? string.Empty;
                return string.Equals(uri, WordprocessingShapeUri, StringComparison.Ordinal) ||
                       string.Equals(uri, WordprocessingGroupUri, StringComparison.Ordinal);
            });
            return !graphicData.IsNone;
        }

        private static ShapeSnapshot CreateShapeSnapshot(
            OfficeXmlNode drawing,
            OfficeXmlNode graphicData,
            string partKey,
            int partIndex,
            int documentOrder,
            WordComparisonOptions options) {
            bool group = TypedDescendants(graphicData).Any(element => element.Is(WordprocessingGroupUri, "wgp"));
            List<string> presets = DescendantsOf(graphicData, Dml, "prstGeom")
                .Select(geometry => geometry.Attribute("prst") ?? string.Empty)
                .Where(value => value.Length > 0)
                .ToList();
            string kind = group ? "shape-group" : "shape";
            string geometryText = presets.Count == 0 ? "custom" : string.Join(",", presets);
            WordDrawingLayoutReader.TryRead(drawing, out WordDrawingLayoutInfo? layout);
            string size = layout == null
                ? "unknown"
                : layout.WidthPoints.ToString("0.###", CultureInfo.InvariantCulture) + "x" +
                  layout.HeightPoints.ToString("0.###", CultureInfo.InvariantCulture) + "pt";
            string name = layout?.Name ?? string.Empty;
            string display = kind + "; geometry=" + geometryText + "; name=" + name +
                             "; placement=" + (layout?.Placement.ToString() ?? "Unknown") + "; size=" + size;
            return new ShapeSnapshot(
                partKey,
                partIndex,
                documentOrder,
                kind,
                HashShapeContent(graphicData, options.CompareGeneratedIds),
                GetShapeLayoutFingerprint(drawing, layout, options.CompareGeneratedIds),
                display);
        }

        private static string HashShapeContent(OfficeXmlNode graphicData, bool compareGeneratedIds) {
            string signature = GetElementSignature(graphicData, compareGeneratedIds ? SignatureRules.None : GeneratedShapeAttributeRules.Instance);
            using SHA256 sha = SHA256.Create();
            return Convert.ToBase64String(sha.ComputeHash(Encoding.UTF8.GetBytes(signature)));
        }

        private static string GetShapeLayoutFingerprint(OfficeXmlNode drawing, WordDrawingLayoutInfo? layout, bool compareGeneratedIds) {
            string generated = string.Empty;
            if (compareGeneratedIds) {
                OfficeXmlNode properties = FirstDescendant(drawing, Wp, "docPr");
                generated = ";id=" + (properties.IsNone ? string.Empty : ParseUInt32(properties.Attribute("id"))?.ToString(CultureInfo.InvariantCulture) ?? string.Empty);
            }
            return (layout?.Placement.ToString() ?? string.Empty) + ";" +
                   (layout?.Name ?? string.Empty) + ";" +
                   (layout?.WidthPoints.ToString("R", CultureInfo.InvariantCulture) ?? string.Empty) + ";" +
                   (layout?.HeightPoints.ToString("R", CultureInfo.InvariantCulture) ?? string.Empty) + ";" +
                   (layout?.HorizontalRelativeFrom ?? string.Empty) + ";" +
                   (layout?.HorizontalOffsetPoints?.ToString("R", CultureInfo.InvariantCulture) ?? string.Empty) + ";" +
                   (layout?.VerticalRelativeFrom ?? string.Empty) + ";" +
                   (layout?.VerticalOffsetPoints?.ToString("R", CultureInfo.InvariantCulture) ?? string.Empty) + ";" +
                   (layout?.HorizontalAlignment ?? string.Empty) + ";" +
                   (layout?.VerticalAlignment ?? string.Empty) + ";" +
                   (layout?.UsesSimplePosition.ToString() ?? string.Empty) + ";" +
                   (layout?.Wrap.ToString() ?? string.Empty) + ";" +
                   (layout?.BehindDocument.ToString() ?? string.Empty) + ";" +
                   (layout?.LayoutInCell.ToString() ?? string.Empty) + ";" +
                   (layout?.AllowOverlap.ToString() ?? string.Empty) + ";" +
                   (layout?.RelativeHeight.ToString(CultureInfo.InvariantCulture) ?? string.Empty) + ";" +
                   (layout?.IsGroup.ToString() ?? string.Empty) + generated;
        }

        /// <summary>Shape markup without editing ids and the generated ids of shape and group drawing properties.</summary>
        private sealed class GeneratedShapeAttributeRules : SignatureRules {
            internal static readonly GeneratedShapeAttributeRules Instance = new();

            internal override void Rewrite(OfficeXmlNode element, List<WordXmlAttribute> attributes) {
                bool drawingProperties = element.Is(WordprocessingGroupUri, "cNvPr") || element.Is(WordprocessingShapeUri, "cNvPr");
                RemoveAll(attributes, attribute =>
                    attribute.LocalName.Equals("anchorId", StringComparison.OrdinalIgnoreCase) ||
                    attribute.LocalName.Equals("editId", StringComparison.OrdinalIgnoreCase) ||
                    attribute.LocalName.Equals("id", StringComparison.OrdinalIgnoreCase) && drawingProperties);
            }
        }

        private static double GetShapeSimilarity(ShapeSnapshot source, ShapeSnapshot target) {
            double score = string.Equals(source.Kind, target.Kind, StringComparison.Ordinal) ? 0.5D : 0D;
            if (string.Equals(source.PartKey, target.PartKey, StringComparison.Ordinal)) score += 0.25D;
            if (string.Equals(source.LayoutFingerprint, target.LayoutFingerprint, StringComparison.Ordinal)) score += 0.25D;
            return score;
        }

        private static string ShapeLocation(int index) => "shape[" + index.ToString(CultureInfo.InvariantCulture) + "]";

        private sealed class ShapeSnapshot : IComparisonFingerprint {
            internal ShapeSnapshot(string partKey, int partIndex, int documentOrder, string kind, string contentFingerprint, string layoutFingerprint, string displayText) {
                PartKey = partKey;
                PartIndex = partIndex;
                DocumentOrder = documentOrder;
                Kind = kind;
                ContentFingerprint = contentFingerprint;
                LayoutFingerprint = layoutFingerprint;
                DisplayText = displayText;
            }
            internal string PartKey { get; }
            internal int PartIndex { get; }
            internal int DocumentOrder { get; }
            internal string Kind { get; }
            internal string ContentFingerprint { get; }
            internal string LayoutFingerprint { get; }
            internal string DisplayText { get; }
            internal string DetailedLocation => PartKey + "/shape[" + PartIndex.ToString(CultureInfo.InvariantCulture) + "]";
            public ulong ComparisonFingerprint => GetOrdinalTextFingerprint(ContentFingerprint);
        }

        private sealed class ShapeSnapshotContentComparer : IEqualityComparer<ShapeSnapshot> {
            internal static readonly ShapeSnapshotContentComparer Instance = new();
            public bool Equals(ShapeSnapshot? x, ShapeSnapshot? y) =>
                ReferenceEquals(x, y) || x != null && y != null &&
                string.Equals(x.ContentFingerprint, y.ContentFingerprint, StringComparison.Ordinal);
            public int GetHashCode(ShapeSnapshot obj) => StringComparer.Ordinal.GetHashCode(obj.ContentFingerprint);
        }
    }
}
