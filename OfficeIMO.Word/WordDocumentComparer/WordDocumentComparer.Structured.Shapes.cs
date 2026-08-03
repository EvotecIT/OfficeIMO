using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using WordDrawing = DocumentFormat.OpenXml.Wordprocessing.Drawing;
using Wpg = DocumentFormat.OpenXml.Office2010.Word.DrawingGroup;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;

namespace OfficeIMO.Word {
    public static partial class WordDocumentComparer {
        private const string WordprocessingShapeUri = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
        private const string WordprocessingGroupUri = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";

        private static void AnalyzeShapes(
            WordDocument source,
            WordDocument target,
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

        private static List<ShapeSnapshot> GetShapeSnapshots(WordDocument document, WordComparisonOptions options) {
            var snapshots = new List<ShapeSnapshot>();
            MainDocumentPart? mainPart = document._wordprocessingDocument.MainDocumentPart;
            AddShapeSnapshots(snapshots, mainPart?.Document?.Body, BodyPartKey, BodyPartOrderBase, options);
            if (mainPart == null) return snapshots;

            int headerIndex = 0;
            foreach (KeyValuePair<HeaderPart, string> part in CreateOrderedHeaderPartKeys(mainPart)) {
                AddShapeSnapshots(snapshots, part.Key.Header, part.Value, HeaderPartOrderBase + (headerIndex++ * RelatedPartOrderStride), options);
            }
            int footerIndex = 0;
            foreach (KeyValuePair<FooterPart, string> part in CreateOrderedFooterPartKeys(mainPart)) {
                AddShapeSnapshots(snapshots, part.Key.Footer, part.Value, FooterPartOrderBase + (footerIndex++ * RelatedPartOrderStride), options);
            }
            List<Footnote> footnotes = GetReferencedFootnotes(mainPart);
            for (int index = 0; index < footnotes.Count; index++) {
                string id = GetNotePartKeyId(footnotes[index], index);
                AddShapeSnapshots(snapshots, footnotes[index], FootnotePartKeyPrefix + id, FootnotePartOrderBase + (index * RelatedPartOrderStride), options);
            }
            List<Endnote> endnotes = GetReferencedEndnotes(mainPart);
            for (int index = 0; index < endnotes.Count; index++) {
                string id = GetNotePartKeyId(endnotes[index], index);
                AddShapeSnapshots(snapshots, endnotes[index], EndnotePartKeyPrefix + id, EndnotePartOrderBase + (index * RelatedPartOrderStride), options);
            }
            return snapshots;
        }

        private static void AddShapeSnapshots(
            List<ShapeSnapshot> snapshots,
            OpenXmlElement? container,
            string partKey,
            int orderBase,
            WordComparisonOptions options) {
            if (container == null) return;
            int partIndex = 0;
            foreach (OrderedElement ordered in EnumerateDescendantsWithOrder(container, orderBase)) {
                if (ordered.Element is not WordDrawing drawing || !TryGetShapeGraphicData(drawing, out A.GraphicData? graphicData)) continue;
                snapshots.Add(CreateShapeSnapshot(drawing, graphicData!, partKey, partIndex++, ordered.DocumentOrder, options));
            }
        }

        private static bool TryGetShapeGraphicData(WordDrawing drawing, out A.GraphicData? graphicData) {
            graphicData = drawing.Descendants<A.GraphicData>().FirstOrDefault(data => {
                string uri = data.Uri?.Value ?? string.Empty;
                return string.Equals(uri, WordprocessingShapeUri, StringComparison.Ordinal) ||
                       string.Equals(uri, WordprocessingGroupUri, StringComparison.Ordinal);
            });
            return graphicData != null;
        }

        private static ShapeSnapshot CreateShapeSnapshot(
            WordDrawing drawing,
            A.GraphicData graphicData,
            string partKey,
            int partIndex,
            int documentOrder,
            WordComparisonOptions options) {
            bool group = graphicData.Descendants<Wpg.WordprocessingGroup>().Any();
            List<string> presets = graphicData.Descendants<A.PresetGeometry>()
                .Select(geometry => geometry.Preset?.InnerText ?? string.Empty)
                .Where(value => value.Length > 0)
                .ToList();
            string kind = group ? "shape-group" : "shape";
            string geometryText = presets.Count == 0 ? "custom" : string.Join(",", presets);
            WordDrawingLayoutReader.TryRead(drawing, out WordDrawingLayoutSnapshot? layout);
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

        private static string HashShapeContent(A.GraphicData graphicData, bool compareGeneratedIds) {
            A.GraphicData clone = (A.GraphicData)graphicData.CloneNode(true);
            if (!compareGeneratedIds) RemoveGeneratedShapeAttributes(clone);
            using SHA256 sha = SHA256.Create();
            return Convert.ToBase64String(sha.ComputeHash(Encoding.UTF8.GetBytes(clone.OuterXml)));
        }

        private static string GetShapeLayoutFingerprint(WordDrawing drawing, WordDrawingLayoutSnapshot? layout, bool compareGeneratedIds) {
            string generated = string.Empty;
            if (compareGeneratedIds) {
                DW.DocProperties? properties = drawing.Descendants<DW.DocProperties>().FirstOrDefault();
                generated = ";id=" + (properties?.Id?.Value.ToString(CultureInfo.InvariantCulture) ?? string.Empty);
            }
            return (layout?.Placement.ToString() ?? string.Empty) + ";" +
                   (layout?.Name ?? string.Empty) + ";" +
                   (layout?.WidthPoints.ToString("R", CultureInfo.InvariantCulture) ?? string.Empty) + ";" +
                   (layout?.HeightPoints.ToString("R", CultureInfo.InvariantCulture) ?? string.Empty) + ";" +
                   (layout?.HorizontalRelativeFrom ?? string.Empty) + ";" +
                   (layout?.HorizontalOffsetPoints?.ToString("R", CultureInfo.InvariantCulture) ?? string.Empty) + ";" +
                   (layout?.VerticalRelativeFrom ?? string.Empty) + ";" +
                   (layout?.VerticalOffsetPoints?.ToString("R", CultureInfo.InvariantCulture) ?? string.Empty) + ";" +
                   (layout?.UsesSimplePosition.ToString() ?? string.Empty) + ";" +
                   (layout?.Wrap.ToString() ?? string.Empty) + generated;
        }

        private static void RemoveGeneratedShapeAttributes(OpenXmlElement root) {
            foreach (OpenXmlElement element in root.Descendants().Prepend(root)) {
                foreach (OpenXmlAttribute attribute in element.GetAttributes()
                    .Where(attribute => attribute.LocalName.Equals("anchorId", StringComparison.OrdinalIgnoreCase) ||
                                        attribute.LocalName.Equals("editId", StringComparison.OrdinalIgnoreCase) ||
                                        attribute.LocalName.Equals("id", StringComparison.OrdinalIgnoreCase) &&
                                        (element is Wpg.NonVisualDrawingProperties ||
                                         element is Wps.NonVisualDrawingProperties))
                    .ToList()) {
                    element.RemoveAttribute(attribute.LocalName, attribute.NamespaceUri);
                }
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
