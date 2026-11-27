using OfficeIMO.Word.OpenXml;

namespace OfficeIMO.Word {
    public static partial class WordDocumentComparer {
        private static void AnalyzeImages(WordPackageView source, WordPackageView target, WordComparisonResult result) {
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

        private static List<ImageSnapshot> GetImageSnapshots(WordPackageView document) {
            var snapshots = new List<ImageSnapshot>();
            WordPartView? mainPart = document.Main;
            AddImageSnapshots(snapshots, mainPart, BodyOf(mainPart), BodyPartKey, BodyPartOrderBase);

            if (mainPart != null) {
                int headerIndex = 0;
                foreach (KeyValuePair<WordPartView, string> headerPartKey in CreateOrderedHeaderPartKeys(mainPart)) {
                    AddImageSnapshots(snapshots, headerPartKey.Key, HeaderRoot(headerPartKey.Key), headerPartKey.Value, HeaderPartOrderBase + (headerIndex * RelatedPartOrderStride));
                    headerIndex++;
                }

                int footerIndex = 0;
                foreach (KeyValuePair<WordPartView, string> footerPartKey in CreateOrderedFooterPartKeys(mainPart)) {
                    AddImageSnapshots(snapshots, footerPartKey.Key, FooterRoot(footerPartKey.Key), footerPartKey.Value, FooterPartOrderBase + (footerIndex * RelatedPartOrderStride));
                    footerIndex++;
                }

                List<OfficeXmlNode> footnotes = GetReferencedFootnotes(mainPart);
                for (int footnoteIndex = 0; footnoteIndex < footnotes.Count; footnoteIndex++) {
                    string noteId = GetNotePartKeyId(footnotes[footnoteIndex], footnoteIndex);
                    AddImageSnapshots(snapshots, FootnotesPart(mainPart), footnotes[footnoteIndex], FootnotePartKeyPrefix + noteId, FootnotePartOrderBase + (footnoteIndex * RelatedPartOrderStride));
                }

                List<OfficeXmlNode> endnotes = GetReferencedEndnotes(mainPart);
                for (int endnoteIndex = 0; endnoteIndex < endnotes.Count; endnoteIndex++) {
                    string noteId = GetNotePartKeyId(endnotes[endnoteIndex], endnoteIndex);
                    AddImageSnapshots(snapshots, EndnotesPart(mainPart), endnotes[endnoteIndex], EndnotePartKeyPrefix + noteId, EndnotePartOrderBase + (endnoteIndex * RelatedPartOrderStride));
                }
            }

            return snapshots;
        }

        private static void AddImageSnapshots(List<ImageSnapshot> snapshots, WordPartView? part, OfficeXmlNode container, string partKey, int orderBase) {
            if (part == null || container.IsNone) {
                return;
            }

            foreach (OrderedElement ordered in EnumerateDescendantsWithOrder(container, orderBase)) {
                OfficeXmlNode element = ordered.Element;
                if (element.Is(W, "drawing")) {
                    OfficeXmlNode blip = FirstDescendant(element, Dml, "blip");
                    if (blip.IsNone) {
                        continue;
                    }

                    string drawingVisualSignature = GetDrawingVisualSignature(part, element);
                    string drawingPositionKey = GetImagePositionKey(partKey, element);
                    if (blip.Attribute(OfficeIMO.Core.Internal.OpenXml.OfficeXmlNamespaces.Relationships, "embed") is string embeddedRelationshipId) {
                        AddEmbeddedImageSnapshot(snapshots, part, embeddedRelationshipId, drawingVisualSignature, partKey, ordered.DocumentOrder, drawingPositionKey);
                    } else if (blip.Attribute(OfficeIMO.Core.Internal.OpenXml.OfficeXmlNamespaces.Relationships, "link") is string externalRelationshipId) {
                        AddExternalImageSnapshot(snapshots, part, externalRelationshipId, drawingVisualSignature, partKey, ordered.DocumentOrder, drawingPositionKey);
                    }
                } else if (element.Is(Vml, "imagedata") &&
                           element.Attribute(OfficeIMO.Core.Internal.OpenXml.OfficeXmlNamespaces.Relationships, "id") is string relationshipId) {
                    string vmlVisualSignature = GetVmlVisualSignature(part, element);
                    string vmlPositionKey = GetImagePositionKey(partKey, element);
                    if (part.HasExternalRelationship(relationshipId)) {
                        AddExternalImageSnapshot(snapshots, part, relationshipId, vmlVisualSignature, partKey, ordered.DocumentOrder, vmlPositionKey);
                    } else {
                        AddEmbeddedImageSnapshot(snapshots, part, relationshipId, vmlVisualSignature, partKey, ordered.DocumentOrder, vmlPositionKey);
                    }
                }
            }
        }

        private static void AddEmbeddedImageSnapshot(List<ImageSnapshot> snapshots, WordPartView part, string relationshipId, string visualSignature, string partKey, int documentOrder, string positionKey) {
            WordPartView? relatedPart = part.GetPartById(relationshipId);
            if (relatedPart == null || !IsImagePartRelationship(part, relationshipId)) {
                return;
            }

            snapshots.Add(ImageSnapshot.FromEmbedded(CreateImageFingerprint(relatedPart), visualSignature, partKey, documentOrder, positionKey));
        }

        private static void AddExternalImageSnapshot(List<ImageSnapshot> snapshots, WordPartView part, string relationshipId, string visualSignature, string partKey, int documentOrder, string positionKey) {
            string externalUri = part.TryGetExternalRelationship(relationshipId, out _, out string uri) ? uri : relationshipId;
            snapshots.Add(ImageSnapshot.FromExternal(externalUri, visualSignature, partKey, documentOrder, positionKey));
        }

        private static string GetDrawingVisualSignature(WordPartView part, OfficeXmlNode drawing) {
            return GetElementSignature(drawing, DrawingVisualSignatureRules.Instance) + GetImageHyperlinkSignature(part, drawing);
        }

        /// <summary>A drawing without its relationship ids, editing ids and generated identity.</summary>
        private sealed class DrawingVisualSignatureRules : SignatureRules {
            internal static readonly DrawingVisualSignatureRules Instance = new();

            internal override void Rewrite(OfficeXmlNode element, List<WordXmlAttribute> attributes) {
                RemoveAll(attributes, attribute =>
                    attribute.NamespaceUri == RelationshipsNamespaceUri && (attribute.LocalName == "embed" || attribute.LocalName == "link") ||
                    attribute.LocalName == "editId" || attribute.LocalName == "anchorId");
                ResetDrawingProperties(element, attributes);
            }
        }

        private static string GetVmlVisualSignature(WordPartView part, OfficeXmlNode imageData) {
            OfficeXmlNode shape = imageData.Parent.IsNone ? imageData : imageData.Parent;
            return GetElementSignature(shape, VmlVisualSignatureRules.Instance) + GetImageHyperlinkSignature(part, imageData);
        }

        /// <summary>A VML shape without its relationship id and generated shape identifiers.</summary>
        private sealed class VmlVisualSignatureRules : SignatureRules {
            internal static readonly VmlVisualSignatureRules Instance = new();

            internal override void Rewrite(OfficeXmlNode element, List<WordXmlAttribute> attributes) {
                RemoveAll(attributes, attribute => attribute.LocalName == "id" || attribute.LocalName == "spid" || attribute.LocalName == "connectortype");
            }
        }

        private static string GetImageHyperlinkSignature(WordPartView part, OfficeXmlNode imageElement) {
            var tokens = new List<string>();
            OfficeXmlNode hyperlink = AncestorW(imageElement, "hyperlink");
            if (!hyperlink.IsNone) {
                tokens.Add("word:" + GetHyperlinkSignature(part, hyperlink));
            }

            foreach (OfficeXmlNode drawingHyperlink in DescendantsOf(imageElement, Dml, "hlinkClick")) {
                tokens.Add("drawing:" + GetDrawingHyperlinkSignature(part, drawingHyperlink));
            }

            return tokens.Count == 0 ? string.Empty : "|hyperlink:" + string.Join("|", tokens.ToArray());
        }

        private static string GetDrawingHyperlinkSignature(WordPartView part, OfficeXmlNode hyperlink) {
            return string.Join(
                "|",
                WordXmlNodes.Attributes(hyperlink)
                    .OrderBy(attribute => attribute.NamespaceUri, StringComparer.Ordinal)
                    .ThenBy(attribute => attribute.LocalName, StringComparer.Ordinal)
                    .Select(attribute => attribute.LocalName == "id"
                        ? attribute.LocalName + "=" + GetRelationshipTarget(part, attribute.Value)
                        : attribute.LocalName + "=" + attribute.Value)
                    .ToArray());
        }

        private static string GetImagePositionKey(string partKey, OfficeXmlNode imageElement) {
            OfficeXmlNode block = AncestorW(imageElement, "p");
            if (block.IsNone) block = AncestorW(imageElement, "tbl");
            if (block.IsNone) block = imageElement;
            return partKey + ":" + GetImageBlockPath(partKey, block) +
                   ":image:" + GetImageOrdinalWithinBlock(block, imageElement).ToString(System.Globalization.CultureInfo.InvariantCulture) +
                   ":offset:" + GetImageInlineOffsetWithinBlock(block, imageElement).ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        private static string GetImageBlockPath(string partKey, OfficeXmlNode block) {
            OfficeXmlNode noteRoot = default;
            if (partKey.StartsWith(FootnotePartKeyPrefix, StringComparison.Ordinal)) {
                noteRoot = AncestorW(block, "footnote");
            } else if (partKey.StartsWith(EndnotePartKeyPrefix, StringComparison.Ordinal)) {
                noteRoot = AncestorW(block, "endnote");
            }

            return noteRoot.IsNone ? GetStableElementPath(block) : GetStableElementPathRelativeTo(block, noteRoot);
        }

        private static int GetImageOrdinalWithinBlock(OfficeXmlNode block, OfficeXmlNode imageElement) {
            int ordinal = 0;
            foreach (OfficeXmlNode element in EnumerateComparableDescendants(block)) {
                if (element == imageElement) {
                    return ordinal;
                }

                if (element.Is(W, "drawing") || element.Is(Vml, "imagedata")) {
                    ordinal++;
                }
            }

            return ordinal;
        }

        private static int GetImageInlineOffsetWithinBlock(OfficeXmlNode block, OfficeXmlNode imageElement) {
            int offset = 0;
            foreach (OfficeXmlNode element in EnumerateComparableDescendants(block)) {
                if (element == imageElement) {
                    return offset;
                }

                offset += GetInlinePositionLength(element);
            }

            return offset;
        }

        private static int GetInlinePositionLength(OfficeXmlNode element) {
            if (element.NamespaceId != W) {
                return 0;
            }

            switch (element.LocalName) {
                case "t":
                    return element.InnerText.Length;
                case "tab":
                    return IsTabChar(element) ? 1 : 0;
                case "br":
                case "sym":
                case "noBreakHyphen":
                case "softHyphen":
                case "cr":
                case "footnoteReference":
                case "endnoteReference":
                    return 1;
                default:
                    return 0;
            }
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

        private static ImageFingerprint CreateImageFingerprint(WordPartView imagePart) {
            byte[] content = imagePart.ReadBytes();
            using System.Security.Cryptography.SHA256 sha256 = System.Security.Cryptography.SHA256.Create();
            return new ImageFingerprint(content.LongLength, Convert.ToBase64String(sha256.ComputeHash(content)));
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
