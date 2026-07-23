using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using V = DocumentFormat.OpenXml.Vml;

namespace OfficeIMO.Word {
    public static partial class WordDocumentComparer {
        private static void AnalyzeImages(WordDocument source, WordDocument target, WordComparisonResult result) {
            IReadOnlyList<ImageSnapshot> sourceImages = GetImageSnapshots(source);
            IReadOnlyList<ImageSnapshot> targetImages = GetImageSnapshots(target);
            IReadOnlyList<MatchedIndexPair> matchedImages = FindMatchingIndexes(
                sourceImages,
                targetImages,
                ImageSnapshotEqualityComparer.Instance);

            int sourceStart = 0;
            int targetStart = 0;

            foreach (MatchedIndexPair match in matchedImages) {
                AddImageRangeFindings(sourceImages, targetImages, sourceStart, match.SourceIndex, targetStart, match.TargetIndex, result);
                sourceStart = match.SourceIndex + 1;
                targetStart = match.TargetIndex + 1;
            }

            AddImageRangeFindings(sourceImages, targetImages, sourceStart, sourceImages.Count, targetStart, targetImages.Count, result);
            AddImagePositionFindings(sourceImages, targetImages, matchedImages, result);
        }

        private static void AddImagePositionFindings(
            IReadOnlyList<ImageSnapshot> sourceImages,
            IReadOnlyList<ImageSnapshot> targetImages,
            IReadOnlyList<MatchedIndexPair> matchedImages,
            WordComparisonResult result) {
            if (sourceImages.Count != targetImages.Count || matchedImages.Count != sourceImages.Count) {
                return;
            }

            foreach (MatchedIndexPair match in matchedImages) {
                ImageSnapshot sourceImage = sourceImages[match.SourceIndex];
                ImageSnapshot targetImage = targetImages[match.TargetIndex];
                if (!ImageSnapshotEqualityComparer.Instance.Equals(sourceImage, targetImage)) {
                    continue;
                }

                if (!string.Equals(sourceImage.PositionKey, targetImage.PositionKey, StringComparison.Ordinal)) {
                    result.Add(new WordComparisonFinding(
                        WordComparisonScope.Image,
                        WordComparisonChangeKind.Modified,
                        ImageLocation(match.TargetIndex),
                        match.SourceIndex,
                        match.TargetIndex,
                        sourceImage.DisplayText,
                        targetImage.DisplayText,
                        "Image position changed."),
                        targetImage.DocumentOrder);
                }

                if (!string.Equals(sourceImage.VisualSignature, targetImage.VisualSignature, StringComparison.Ordinal)) {
                    result.Add(new WordComparisonFinding(
                        WordComparisonScope.Image,
                        WordComparisonChangeKind.Modified,
                        ImageLocation(match.TargetIndex),
                        match.SourceIndex,
                        match.TargetIndex,
                        sourceImage.DisplayText,
                        targetImage.DisplayText,
                        "Image layout changed."),
                        targetImage.DocumentOrder);
                }
            }
        }

        private static void AddImageRangeFindings(
            IReadOnlyList<ImageSnapshot> sourceImages,
            IReadOnlyList<ImageSnapshot> targetImages,
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
                    GetImageSimilarity(sourceImages[sourceIndex], targetImages[targetIndex + 1]) >
                    GetImageSimilarity(sourceImages[sourceIndex], targetImages[targetIndex])) {
                    AddInsertedImageFinding(targetImages, targetIndex, result);
                    targetIndex++;
                    continue;
                }

                if (sourceEnd - sourceIndex > targetEnd - targetIndex &&
                    sourceIndex + 1 < sourceEnd &&
                    GetImageSimilarity(sourceImages[sourceIndex + 1], targetImages[targetIndex]) >
                    GetImageSimilarity(sourceImages[sourceIndex], targetImages[targetIndex])) {
                    AddDeletedImageFinding(sourceImages, sourceIndex, result);
                    sourceIndex++;
                    continue;
                }

                if (!string.Equals(sourceImages[sourceIndex].PartKey, targetImages[targetIndex].PartKey, StringComparison.Ordinal)) {
                    AddDeletedImageFinding(sourceImages, sourceIndex, result);
                    AddInsertedImageFinding(targetImages, targetIndex, result);
                    sourceIndex++;
                    targetIndex++;
                    continue;
                }

                bool samePayload = HasSameImagePayload(sourceImages[sourceIndex], targetImages[targetIndex]);
                result.Add(new WordComparisonFinding(
                    WordComparisonScope.Image,
                    WordComparisonChangeKind.Modified,
                    ImageLocation(targetIndex),
                    sourceIndex,
                    targetIndex,
                    sourceImages[sourceIndex].DisplayText,
                    targetImages[targetIndex].DisplayText,
                    samePayload ? "Image layout changed." : "Image payload changed."),
                    targetImages[targetIndex].DocumentOrder);

                sourceIndex++;
                targetIndex++;
            }

            while (targetIndex < targetEnd) {
                AddInsertedImageFinding(targetImages, targetIndex, result);
                targetIndex++;
            }

            while (sourceIndex < sourceEnd) {
                AddDeletedImageFinding(sourceImages, sourceIndex, result);
                sourceIndex++;
            }
        }

        private static void AddInsertedImageFinding(IReadOnlyList<ImageSnapshot> targetImages, int imageIndex, WordComparisonResult result) {
            result.Add(new WordComparisonFinding(
                WordComparisonScope.Image,
                WordComparisonChangeKind.Inserted,
                ImageLocation(imageIndex),
                null,
                imageIndex,
                null,
                targetImages[imageIndex].DisplayText,
                "Image inserted."),
                targetImages[imageIndex].DocumentOrder);
        }

        private static void AddDeletedImageFinding(IReadOnlyList<ImageSnapshot> sourceImages, int imageIndex, WordComparisonResult result) {
            result.Add(new WordComparisonFinding(
                WordComparisonScope.Image,
                WordComparisonChangeKind.Deleted,
                ImageLocation(imageIndex),
                imageIndex,
                null,
                sourceImages[imageIndex].DisplayText,
                null,
                "Image deleted."),
                sourceImages[imageIndex].DocumentOrder);
        }

        private static List<ImageSnapshot> GetImageSnapshots(WordDocument document) {
            var snapshots = new List<ImageSnapshot>();
            MainDocumentPart? mainPart = document._wordprocessingDocument.MainDocumentPart;
            AddImageSnapshots(snapshots, mainPart, mainPart?.Document?.Body, BodyPartKey, BodyPartOrderBase);

            if (mainPart != null) {
                int headerIndex = 0;
                foreach (KeyValuePair<HeaderPart, string> headerPartKey in CreateOrderedHeaderPartKeys(mainPart)) {
                    AddImageSnapshots(snapshots, headerPartKey.Key, headerPartKey.Key.Header, headerPartKey.Value, HeaderPartOrderBase + (headerIndex * RelatedPartOrderStride));
                    headerIndex++;
                }

                int footerIndex = 0;
                foreach (KeyValuePair<FooterPart, string> footerPartKey in CreateOrderedFooterPartKeys(mainPart)) {
                    AddImageSnapshots(snapshots, footerPartKey.Key, footerPartKey.Key.Footer, footerPartKey.Value, FooterPartOrderBase + (footerIndex * RelatedPartOrderStride));
                    footerIndex++;
                }

                List<Footnote> footnotes = GetReferencedFootnotes(mainPart);
                for (int footnoteIndex = 0; footnoteIndex < footnotes.Count; footnoteIndex++) {
                    string noteId = footnotes[footnoteIndex].Id?.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) ??
                        footnoteIndex.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    AddImageSnapshots(snapshots, mainPart.FootnotesPart, footnotes[footnoteIndex], FootnotePartKeyPrefix + noteId, FootnotePartOrderBase + (footnoteIndex * RelatedPartOrderStride));
                }

                List<Endnote> endnotes = GetReferencedEndnotes(mainPart);
                for (int endnoteIndex = 0; endnoteIndex < endnotes.Count; endnoteIndex++) {
                    string noteId = endnotes[endnoteIndex].Id?.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) ??
                        endnoteIndex.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    AddImageSnapshots(snapshots, mainPart.EndnotesPart, endnotes[endnoteIndex], EndnotePartKeyPrefix + noteId, EndnotePartOrderBase + (endnoteIndex * RelatedPartOrderStride));
                }
            }

            return snapshots;
        }

        private static void AddImageSnapshots(List<ImageSnapshot> snapshots, OpenXmlPart? part, OpenXmlElement? container, string partKey, int orderBase) {
            if (part == null || container == null) {
                return;
            }

            foreach (OrderedElement ordered in EnumerateDescendantsWithOrder(container, orderBase)) {
                switch (ordered.Element) {
                    case DocumentFormat.OpenXml.Wordprocessing.Drawing drawing:
                        DocumentFormat.OpenXml.Drawing.Blip? blip = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().FirstOrDefault();
                        if (blip == null) {
                            break;
                        }

                        string drawingVisualSignature = GetDrawingVisualSignature(part, drawing);
                        string drawingPositionKey = GetImagePositionKey(partKey, drawing);
                        if (blip.Embed?.Value is string embeddedRelationshipId) {
                            AddEmbeddedImageSnapshot(snapshots, part, embeddedRelationshipId, drawingVisualSignature, partKey, ordered.DocumentOrder, drawingPositionKey);
                        } else if (blip.Link?.Value is string externalRelationshipId) {
                            AddExternalImageSnapshot(snapshots, part, externalRelationshipId, drawingVisualSignature, partKey, ordered.DocumentOrder, drawingPositionKey);
                        }

                        break;
                    case V.ImageData imageData when imageData.RelationshipId?.Value is string relationshipId:
                        string vmlVisualSignature = GetVmlVisualSignature(part, imageData);
                        string vmlPositionKey = GetImagePositionKey(partKey, imageData);
                        if (part.ExternalRelationships.Any(item => item.Id == relationshipId)) {
                            AddExternalImageSnapshot(snapshots, part, relationshipId, vmlVisualSignature, partKey, ordered.DocumentOrder, vmlPositionKey);
                        } else {
                            AddEmbeddedImageSnapshot(snapshots, part, relationshipId, vmlVisualSignature, partKey, ordered.DocumentOrder, vmlPositionKey);
                        }

                        break;
                }
            }
        }

        private static void AddEmbeddedImageSnapshot(List<ImageSnapshot> snapshots, OpenXmlPart part, string relationshipId, string visualSignature, string partKey, int documentOrder, string positionKey) {
            OpenXmlPart relatedPart;
            try {
                relatedPart = part.GetPartById(relationshipId);
            } catch (ArgumentOutOfRangeException) {
                return;
            }

            if (relatedPart is not ImagePart imagePart) {
                return;
            }

            using Stream stream = imagePart.GetStream(FileMode.Open, FileAccess.Read);
            snapshots.Add(ImageSnapshot.FromEmbedded(CreateImageFingerprint(stream), visualSignature, partKey, documentOrder, positionKey));
        }

        private static void AddExternalImageSnapshot(List<ImageSnapshot> snapshots, OpenXmlPart part, string relationshipId, string visualSignature, string partKey, int documentOrder, string positionKey) {
            ExternalRelationship? relationship = part.ExternalRelationships.FirstOrDefault(item => item.Id == relationshipId);
            snapshots.Add(ImageSnapshot.FromExternal(relationship?.Uri?.ToString() ?? relationshipId, visualSignature, partKey, documentOrder, positionKey));
        }

        private static string GetDrawingVisualSignature(OpenXmlPart part, DocumentFormat.OpenXml.Wordprocessing.Drawing drawing) {
            OpenXmlElement clone = drawing.CloneNode(true);
            foreach (DocumentFormat.OpenXml.Drawing.Blip blip in clone.Descendants<DocumentFormat.OpenXml.Drawing.Blip>()) {
                blip.Embed = null;
                blip.Link = null;
            }

            foreach (OpenXmlElement element in new[] { clone }.Concat(clone.Descendants())) {
                element.RemoveAttribute("embed", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                element.RemoveAttribute("link", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                RemoveVolatileDrawingAttributes(element);
            }

            foreach (DW.DocProperties properties in clone.Descendants<DW.DocProperties>()) {
                properties.Id = 0U;
                properties.Name = string.Empty;
            }

            foreach (PIC.NonVisualDrawingProperties properties in clone.Descendants<PIC.NonVisualDrawingProperties>()) {
                properties.Id = 0U;
                properties.Name = string.Empty;
            }

            return clone.OuterXml + GetImageHyperlinkSignature(part, drawing);
        }

        private static string GetVmlVisualSignature(OpenXmlPart part, V.ImageData imageData) {
            OpenXmlElement clone = (imageData.Parent ?? imageData).CloneNode(true);
            if (clone is V.ImageData clonedImageData) {
                clonedImageData.RelationshipId = null;
            }

            foreach (V.ImageData descendant in clone.Descendants<V.ImageData>()) {
                descendant.RelationshipId = null;
            }

            foreach (OpenXmlElement element in new[] { clone }.Concat(clone.Descendants())) {
                if (element is V.Shape shape) {
                    shape.Id = null;
                }

                RemoveVolatileVmlAttributes(element);
            }

            return clone.OuterXml + GetImageHyperlinkSignature(part, imageData);
        }

        private static string GetImageHyperlinkSignature(OpenXmlPart part, OpenXmlElement imageElement) {
            var tokens = new List<string>();
            Hyperlink? hyperlink = imageElement.Ancestors<Hyperlink>().FirstOrDefault();
            if (hyperlink != null) {
                tokens.Add("word:" + GetHyperlinkSignature(part, hyperlink));
            }

            foreach (A.HyperlinkOnClick drawingHyperlink in imageElement.Descendants<A.HyperlinkOnClick>()) {
                tokens.Add("drawing:" + GetDrawingHyperlinkSignature(part, drawingHyperlink));
            }

            return tokens.Count == 0 ? string.Empty : "|hyperlink:" + string.Join("|", tokens.ToArray());
        }

        private static string GetDrawingHyperlinkSignature(OpenXmlPart part, A.HyperlinkOnClick hyperlink) {
            return string.Join(
                "|",
                hyperlink.GetAttributes()
                    .OrderBy(attribute => attribute.NamespaceUri, StringComparer.Ordinal)
                    .ThenBy(attribute => attribute.LocalName, StringComparer.Ordinal)
                    .Select(attribute => attribute.LocalName == "id"
                        ? attribute.LocalName + "=" + GetRelationshipTarget(part, attribute.Value ?? string.Empty)
                        : attribute.LocalName + "=" + (attribute.Value ?? string.Empty))
                    .ToArray());
        }

        private static string GetImagePositionKey(string partKey, OpenXmlElement imageElement) {
            OpenXmlElement block = imageElement.Ancestors<Paragraph>().FirstOrDefault() ??
                                   imageElement.Ancestors<Table>().FirstOrDefault() ??
                                   imageElement;
            return partKey + ":" + GetImageBlockPath(partKey, block) +
                   ":image:" + GetImageOrdinalWithinBlock(block, imageElement).ToString(System.Globalization.CultureInfo.InvariantCulture) +
                   ":offset:" + GetImageInlineOffsetWithinBlock(block, imageElement).ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        private static string GetImageBlockPath(string partKey, OpenXmlElement block) {
            OpenXmlElement? noteRoot = null;
            if (partKey.StartsWith(FootnotePartKeyPrefix, StringComparison.Ordinal)) {
                noteRoot = block.Ancestors<Footnote>().FirstOrDefault();
            } else if (partKey.StartsWith(EndnotePartKeyPrefix, StringComparison.Ordinal)) {
                noteRoot = block.Ancestors<Endnote>().FirstOrDefault();
            }

            return noteRoot == null ? GetStableElementPath(block) : GetStableElementPathRelativeTo(block, noteRoot);
        }

        private static string GetStableElementPathRelativeTo(OpenXmlElement element, OpenXmlElement root) {
            var segments = new Stack<string>();
            OpenXmlElement? current = element;
            while (current != null && current.Parent != null && !ReferenceEquals(current, root)) {
                OpenXmlElement parent = current.Parent;
                int ordinal = parent.Elements()
                    .Where(item => item.GetType() == current.GetType())
                    .TakeWhile(item => !ReferenceEquals(item, current))
                    .Count();
                segments.Push(current.GetType().Name + "[" + ordinal.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]");
                current = parent;
            }

            return string.Join("/", segments.ToArray());
        }

        private static void RemoveVolatileDrawingAttributes(OpenXmlElement element) {
            foreach (OpenXmlAttribute attribute in element.GetAttributes().ToList()) {
                if (attribute.LocalName == "editId" || attribute.LocalName == "anchorId") {
                    element.RemoveAttribute(attribute.LocalName, attribute.NamespaceUri);
                }
            }
        }

        private static void RemoveVolatileVmlAttributes(OpenXmlElement element) {
            foreach (OpenXmlAttribute attribute in element.GetAttributes().ToList()) {
                if (attribute.LocalName == "id" ||
                    attribute.LocalName == "spid" ||
                    attribute.LocalName == "connectortype") {
                    element.RemoveAttribute(attribute.LocalName, attribute.NamespaceUri);
                }
            }
        }

        private static int GetImageOrdinalWithinBlock(OpenXmlElement block, OpenXmlElement imageElement) {
            int ordinal = 0;
            foreach (OpenXmlElement element in EnumerateComparableDescendants(block)) {
                if (ReferenceEquals(element, imageElement)) {
                    return ordinal;
                }

                if (element is DocumentFormat.OpenXml.Wordprocessing.Drawing || element is V.ImageData) {
                    ordinal++;
                }
            }

            return ordinal;
        }

        private static int GetImageInlineOffsetWithinBlock(OpenXmlElement block, OpenXmlElement imageElement) {
            int offset = 0;
            foreach (OpenXmlElement element in EnumerateComparableDescendants(block)) {
                if (ReferenceEquals(element, imageElement)) {
                    return offset;
                }

                offset += GetInlinePositionLength(element);
            }

            return offset;
        }

        private static int GetInlinePositionLength(OpenXmlElement element) {
            return element switch {
                Text text => text.Text?.Length ?? 0,
                TabChar => 1,
                Break => 1,
                SymbolChar => 1,
                NoBreakHyphen => 1,
                SoftHyphen => 1,
                CarriageReturn => 1,
                FootnoteReference => 1,
                EndnoteReference => 1,
                _ => 0
            };
        }

        private static string ImageLocation(int imageIndex) {
            return "image[" + imageIndex.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]";
        }

        private static bool HasSameImagePayload(ImageSnapshot sourceImage, ImageSnapshot targetImage) {
            if (!string.Equals(sourceImage.PartKey, targetImage.PartKey, StringComparison.Ordinal)) {
                return false;
            }

            if (sourceImage.ExternalUri != null || targetImage.ExternalUri != null) {
                return string.Equals(sourceImage.ExternalUri, targetImage.ExternalUri, StringComparison.Ordinal);
            }

            return sourceImage.EmbeddedFingerprint != null &&
                   targetImage.EmbeddedFingerprint != null &&
                   sourceImage.EmbeddedFingerprint.Equals(targetImage.EmbeddedFingerprint);
        }

        private static ImageFingerprint CreateImageFingerprint(Stream stream) {
            using System.Security.Cryptography.SHA256 sha256 = System.Security.Cryptography.SHA256.Create();
            byte[] buffer = new byte[81920];
            long length = 0;
            int bytesRead;
            while ((bytesRead = stream.Read(buffer, 0, buffer.Length)) > 0) {
                sha256.TransformBlock(buffer, 0, bytesRead, null, 0);
                length += bytesRead;
            }

            sha256.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
            return new ImageFingerprint(length, Convert.ToBase64String(sha256.Hash ?? Array.Empty<byte>()));
        }

        private sealed class ImageSnapshot : IComparisonFingerprint {
            private ImageSnapshot(ImageFingerprint? embeddedFingerprint, string? externalUri, string visualSignature, string partKey, int documentOrder, string positionKey) {
                EmbeddedFingerprint = embeddedFingerprint;
                ExternalUri = externalUri;
                VisualSignature = visualSignature;
                PartKey = partKey;
                DocumentOrder = documentOrder;
                PositionKey = positionKey;
            }

            internal ImageFingerprint? EmbeddedFingerprint { get; }

            internal string? ExternalUri { get; }

            internal string VisualSignature { get; }

            internal string PartKey { get; }

            internal int DocumentOrder { get; }

            internal string PositionKey { get; }

            internal string DisplayText => ExternalUri == null ? "[Image]" : "[Image: " + ExternalUri + "]";

            public ulong ComparisonFingerprint {
                get {
                    ulong fingerprint = CombineComparisonFingerprints(
                        GetOrdinalTextFingerprint(PartKey),
                        GetOrdinalTextFingerprint(VisualSignature));
                    if (ExternalUri != null) {
                        return CombineComparisonFingerprints(
                            fingerprint,
                            CombineComparisonFingerprints(0x45585445524E414CUL, GetOrdinalTextFingerprint(ExternalUri)));
                    }

                    if (EmbeddedFingerprint != null) {
                        ulong embedded = CombineComparisonFingerprints(
                            unchecked((ulong)EmbeddedFingerprint.Value.Length),
                            GetOrdinalTextFingerprint(EmbeddedFingerprint.Value.Sha256));
                        return CombineComparisonFingerprints(
                            fingerprint,
                            CombineComparisonFingerprints(0x454D424544444544UL, embedded));
                    }

                    return fingerprint;
                }
            }

            internal static ImageSnapshot FromEmbedded(ImageFingerprint embeddedFingerprint, string visualSignature, string partKey, int documentOrder, string positionKey) {
                return new ImageSnapshot(embeddedFingerprint, null, visualSignature, partKey, documentOrder, positionKey);
            }

            internal static ImageSnapshot FromExternal(string externalUri, string visualSignature, string partKey, int documentOrder, string positionKey) {
                return new ImageSnapshot(null, externalUri, visualSignature, partKey, documentOrder, positionKey);
            }
        }

        private sealed class ImageSnapshotEqualityComparer : IEqualityComparer<ImageSnapshot> {
            internal static readonly ImageSnapshotEqualityComparer Instance = new();

            public bool Equals(ImageSnapshot? x, ImageSnapshot? y) {
                if (ReferenceEquals(x, y)) {
                    return true;
                }

                if (x == null || y == null) {
                    return false;
                }

                if (x.ExternalUri != null || y.ExternalUri != null) {
                    return string.Equals(x.PartKey, y.PartKey, StringComparison.Ordinal) &&
                           string.Equals(x.ExternalUri, y.ExternalUri, StringComparison.Ordinal) &&
                           string.Equals(x.VisualSignature, y.VisualSignature, StringComparison.Ordinal);
                }

                return x.EmbeddedFingerprint != null &&
                       y.EmbeddedFingerprint != null &&
                       string.Equals(x.PartKey, y.PartKey, StringComparison.Ordinal) &&
                       x.EmbeddedFingerprint.Equals(y.EmbeddedFingerprint) &&
                       string.Equals(x.VisualSignature, y.VisualSignature, StringComparison.Ordinal);
            }

            public int GetHashCode(ImageSnapshot obj) {
                if (obj.ExternalUri != null) {
                    int externalHash = StringComparer.Ordinal.GetHashCode(obj.PartKey);
                    externalHash = (externalHash * 397) ^ StringComparer.Ordinal.GetHashCode(obj.ExternalUri);
                    return (externalHash * 397) ^ StringComparer.Ordinal.GetHashCode(obj.VisualSignature);
                }

                if (obj.EmbeddedFingerprint == null) {
                    return StringComparer.Ordinal.GetHashCode(obj.PartKey);
                }

                unchecked {
                    int hashCode = StringComparer.Ordinal.GetHashCode(obj.PartKey);
                    hashCode = (hashCode * 397) ^ obj.EmbeddedFingerprint.GetHashCode();
                    return (hashCode * 397) ^ StringComparer.Ordinal.GetHashCode(obj.VisualSignature);
                }
            }
        }

        private readonly struct ImageFingerprint : IEquatable<ImageFingerprint> {
            internal ImageFingerprint(long length, string sha256) {
                Length = length;
                Sha256 = sha256;
            }

            internal long Length { get; }

            internal string Sha256 { get; }

            public bool Equals(ImageFingerprint other) {
                return Length == other.Length &&
                       string.Equals(Sha256, other.Sha256, StringComparison.Ordinal);
            }

            public override bool Equals(object? obj) {
                return obj is ImageFingerprint other && Equals(other);
            }

            public override int GetHashCode() {
                unchecked {
                    return (Length.GetHashCode() * 397) ^ StringComparer.Ordinal.GetHashCode(Sha256);
                }
            }
        }
    }
}
