using OfficeIMO.Word.OpenXml;

namespace OfficeIMO.Word {
    public static partial class WordDocumentComparer {
        private const string OfficeRelationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        private static void ApplyImageFindings(WordPackageView sourceDocument, WordPackageView targetDocument, WordComparisonResult result, WordComparisonRedlineOptions options) {
            List<RedlineImageContainer> targetContainers = GetRedlineImageContainers(targetDocument);
            List<RedlineImageEntry> sourceImages = GetRedlineImageEntries(sourceDocument);
            List<RedlineImageEntry> targetImages = GetRedlineImageEntries(targetDocument);
            var rewrittenTargetImages = new HashSet<int>();
            var insertedSourceImages = new HashSet<int>();
            var insertedBeforeImageAnchors = new Dictionary<OfficeXmlNode, OfficeXmlNode>();
            var insertedAfterImageAnchors = new Dictionary<OfficeXmlNode, OfficeXmlNode>();

            List<WordComparisonFinding> imageFindings = result.Findings
                .Where(finding => ShouldTrackFinding(finding, options) && finding.Scope == WordComparisonScope.Image)
                .ToList();

            foreach (WordComparisonFinding finding in imageFindings.Where(finding => finding.ChangeKind != WordComparisonChangeKind.Deleted)) {
                if (!ShouldTrackFinding(finding, options) ||
                    finding.Scope != WordComparisonScope.Image) {
                    continue;
                }

                switch (finding.ChangeKind) {
                    case WordComparisonChangeKind.Inserted:
                        int insertedTargetIndex = GetTargetImageIndex(finding);
                        if (insertedTargetIndex >= 0 &&
                            insertedTargetIndex < targetImages.Count &&
                            !rewrittenTargetImages.Contains(insertedTargetIndex) &&
                            WrapImageRunAsInserted(targetImages[insertedTargetIndex], options)) {
                            rewrittenTargetImages.Add(insertedTargetIndex);
                        }

                        break;
                    case WordComparisonChangeKind.Modified:
                        int modifiedSourceIndex = GetSourceImageIndex(finding);
                        int modifiedTargetIndex = GetTargetImageIndex(finding);
                        if (modifiedSourceIndex >= 0 &&
                            modifiedSourceIndex < sourceImages.Count &&
                            modifiedTargetIndex >= 0 &&
                            modifiedTargetIndex < targetImages.Count &&
                            !rewrittenTargetImages.Contains(modifiedTargetIndex) &&
                            WrapImageRunAsChanged(sourceImages[modifiedSourceIndex], targetImages[modifiedTargetIndex], targetDocument, options)) {
                            rewrittenTargetImages.Add(modifiedTargetIndex);
                            insertedSourceImages.Add(modifiedSourceIndex);
                        }

                        break;
                }
            }

            foreach (WordComparisonFinding finding in imageFindings
                .Where(finding => finding.ChangeKind == WordComparisonChangeKind.Deleted)
                .OrderBy(finding => GetSourceImageIndex(finding))) {
                int deletedSourceIndex = GetSourceImageIndex(finding);
                if (deletedSourceIndex >= 0 &&
                    deletedSourceIndex < sourceImages.Count &&
                    !insertedSourceImages.Contains(deletedSourceIndex) &&
                    InsertDeletedImageRun(sourceImages, sourceImages[deletedSourceIndex], targetContainers, targetImages, targetDocument, options, insertedBeforeImageAnchors, insertedAfterImageAnchors)) {
                    insertedSourceImages.Add(deletedSourceIndex);
                }
            }
        }

        private static List<RedlineImageEntry> GetRedlineImageEntries(WordPackageView document) {
            var entries = new List<RedlineImageEntry>();
            foreach (RedlineImageContainer container in GetRedlineImageContainers(document)) {
                AddRedlineImageEntries(entries, container);
            }

            return entries;
        }

        private static List<RedlineImageContainer> GetRedlineImageContainers(WordPackageView document) {
            WordPartView? mainPart = document.Main;
            var containers = new List<RedlineImageContainer>();
            OfficeXmlNode body = BodyOf(mainPart);
            if (!body.IsNone) {
                containers.Add(new RedlineImageContainer(containers.Count, BodyPartKey, mainPart!, body, BodyPartOrderBase));
            }

            if (mainPart == null) {
                return containers;
            }

            int headerIndex = 0;
            foreach (KeyValuePair<WordPartView, string> headerPartKey in CreateOrderedHeaderPartKeys(mainPart)) {
                OfficeXmlNode header = HeaderRoot(headerPartKey.Key);
                if (!header.IsNone) {
                    containers.Add(new RedlineImageContainer(containers.Count, headerPartKey.Value, headerPartKey.Key, header, HeaderPartOrderBase + (headerIndex * RelatedPartOrderStride)));
                }

                headerIndex++;
            }

            int footerIndex = 0;
            foreach (KeyValuePair<WordPartView, string> footerPartKey in CreateOrderedFooterPartKeys(mainPart)) {
                OfficeXmlNode footer = FooterRoot(footerPartKey.Key);
                if (!footer.IsNone) {
                    containers.Add(new RedlineImageContainer(containers.Count, footerPartKey.Value, footerPartKey.Key, footer, FooterPartOrderBase + (footerIndex * RelatedPartOrderStride)));
                }

                footerIndex++;
            }

            List<OfficeXmlNode> footnotes = GetReferencedFootnotes(mainPart);
            WordPartView? footnotesPart = FootnotesPart(mainPart);
            for (int footnoteIndex = 0; footnoteIndex < footnotes.Count; footnoteIndex++) {
                if (footnotesPart != null) {
                    string noteId = GetNotePartKeyId(footnotes[footnoteIndex], footnoteIndex);
                    containers.Add(new RedlineImageContainer(containers.Count, FootnotePartKeyPrefix + noteId, footnotesPart, footnotes[footnoteIndex], FootnotePartOrderBase + (footnoteIndex * RelatedPartOrderStride)));
                }
            }

            List<OfficeXmlNode> endnotes = GetReferencedEndnotes(mainPart);
            WordPartView? endnotesPart = EndnotesPart(mainPart);
            for (int endnoteIndex = 0; endnoteIndex < endnotes.Count; endnoteIndex++) {
                if (endnotesPart != null) {
                    string noteId = GetNotePartKeyId(endnotes[endnoteIndex], endnoteIndex);
                    containers.Add(new RedlineImageContainer(containers.Count, EndnotePartKeyPrefix + noteId, endnotesPart, endnotes[endnoteIndex], EndnotePartOrderBase + (endnoteIndex * RelatedPartOrderStride)));
                }
            }

            return containers;
        }

        private static void AddRedlineImageEntries(List<RedlineImageEntry> entries, RedlineImageContainer container) {
            int containerImageIndex = 0;
            foreach (OrderedElement ordered in EnumerateDescendantsWithOrder(container.Container, container.OrderBase)) {
                if (!IsRedlineImageElement(ordered.Element)) {
                    continue;
                }

                OfficeXmlNode run = AncestorW(ordered.Element, "r");
                if (run.IsNone) {
                    continue;
                }

                OfficeXmlNode imageElement = GetImageRedlineElement(ordered.Element);
                entries.Add(new RedlineImageEntry(entries.Count, container.Index, container.PartKey, containerImageIndex, container.Part, container.Container, run, imageElement));
                containerImageIndex++;
            }
        }

        private static OfficeXmlNode GetImageRedlineElement(OfficeXmlNode element) {
            if (element.Is(Vml, "imagedata")) {
                OfficeXmlNode picture = AncestorW(element, "pict");
                return picture.IsNone ? element : picture;
            }

            return element;
        }

        private static bool IsRedlineImageElement(OfficeXmlNode element) {
            return element.Is(Vml, "imagedata") ||
                element.Is(W, "drawing") && HasDescendant(element, Dml, "blip");
        }

        private static bool IsInTrackedRun(OfficeXmlNode run) => HasAncestorW(run, "ins") || HasAncestorW(run, "del");

        private static bool WrapImageRunAsInserted(RedlineImageEntry entry, WordComparisonRedlineOptions options) {
            OfficeXmlNode imageRun = entry.Run;
            OfficeXmlNode parent = imageRun.Parent;
            if (parent.IsNone || IsInTrackedRun(imageRun)) {
                return false;
            }

            OfficeXmlNode inserted = parent.CreateElement(W, "ins");
            SetRevisionAttributes(inserted, options.Author, options.DateTime ?? DateTime.Now);
            OfficeXmlNode insertedRun = CloneImageOnlyRun(inserted, entry);
            inserted.Append(insertedRun);
            parent.InsertBefore(inserted, imageRun);
            RemoveImageElementFromRun(entry);
            return true;
        }

        private static bool WrapImageRunAsChanged(RedlineImageEntry sourceEntry, RedlineImageEntry targetEntry, WordPackageView targetDocument, WordComparisonRedlineOptions options) {
            OfficeXmlNode targetRun = targetEntry.Run;
            OfficeXmlNode parent = targetRun.Parent;
            if (parent.IsNone || IsInTrackedRun(targetRun)) {
                return false;
            }

            OfficeXmlNode deletedRun = CloneImageRunForPart(parent, sourceEntry, targetEntry.Part, targetDocument);
            if (deletedRun.IsNone) {
                return false;
            }

            parent.InsertBefore(WrapRunAsDeleted(parent, deletedRun, options), targetRun);
            parent.InsertBefore(WrapRunAsInserted(parent, CloneImageOnlyRun(parent, targetEntry), options), targetRun);
            RemoveImageElementFromRun(targetEntry);
            return true;
        }

        private static bool InsertDeletedImageRun(
            IReadOnlyList<RedlineImageEntry> sourceImages,
            RedlineImageEntry sourceEntry,
            IReadOnlyList<RedlineImageContainer> targetContainers,
            IReadOnlyList<RedlineImageEntry> targetImages,
            WordPackageView targetDocument,
            WordComparisonRedlineOptions options,
            Dictionary<OfficeXmlNode, OfficeXmlNode> insertedBeforeImageAnchors,
            Dictionary<OfficeXmlNode, OfficeXmlNode> insertedAfterImageAnchors) {
            RedlineImageContainer targetContainer = GetTargetImageContainer(sourceEntry, targetContainers);
            // Runs and their revision wrappers are made as a paragraph holds them, whatever the container is.
            OfficeXmlNode context = targetContainer.Container.CreateElement(W, "p");
            OfficeXmlNode deletedRun = CloneImageRunForPart(context, sourceEntry, targetContainer.Part, targetDocument);
            if (deletedRun.IsNone) {
                return false;
            }

            OfficeXmlNode deleted = WrapRunAsDeleted(context, deletedRun, options);
            List<RedlineImageEntry> sourcePartImages = sourceImages
                .Where(image => string.Equals(image.PartKey, sourceEntry.PartKey, StringComparison.Ordinal))
                .ToList();
            List<RedlineImageEntry> targetPartImages = targetImages
                .Where(image => string.Equals(image.PartKey, targetContainer.PartKey, StringComparison.Ordinal))
                .ToList();
            int targetImageIndex = FindTargetGapByNeighborIdentity(
                sourcePartImages,
                targetPartImages,
                sourceEntry.ContainerImageIndex,
                GetRedlineImageIdentity,
                GetRedlineImageIdentity);
            RedlineImageEntry? nextImage = targetImages
                .Where(image => string.Equals(image.PartKey, targetContainer.PartKey, StringComparison.Ordinal) && image.ContainerImageIndex >= targetImageIndex)
                .OrderBy(image => image.ContainerImageIndex)
                .FirstOrDefault();
            if (nextImage != null && !nextImage.Run.Parent.IsNone) {
                InsertBeforeImageAnchor(nextImage.Run, deleted, insertedBeforeImageAnchors);
                return true;
            }

            RedlineImageEntry? previousImage = targetImages
                .Where(image => string.Equals(image.PartKey, targetContainer.PartKey, StringComparison.Ordinal) && image.ContainerImageIndex < targetImageIndex)
                .OrderByDescending(image => image.ContainerImageIndex)
                .FirstOrDefault();
            if (previousImage != null && !previousImage.Run.Parent.IsNone) {
                InsertAfterImageAnchor(previousImage.Run, deleted, insertedAfterImageAnchors);
                return true;
            }

            if (TryInsertDeletedImageAtSourceParagraphGap(sourceEntry, targetContainer, deleted)) {
                return true;
            }

            OfficeXmlNode paragraph = targetContainer.Container.CreateElement(W, "p");
            paragraph.Append(deleted);
            AppendImageFallbackParagraph(targetContainer.Container, paragraph);
            return true;
        }

        private static bool TryInsertDeletedImageAtSourceParagraphGap(RedlineImageEntry sourceEntry, RedlineImageContainer targetContainer, OfficeXmlNode deleted) {
            OfficeXmlNode sourceParagraph = AncestorW(sourceEntry.Run, "p");
            if (sourceParagraph.IsNone) {
                return false;
            }

            List<OfficeXmlNode> sourceParagraphs = DescendantsW(sourceEntry.Container, "p");
            List<OfficeXmlNode> targetParagraphs = DescendantsW(targetContainer.Container, "p");
            int sourceParagraphIndex = sourceParagraphs.FindIndex(paragraph => paragraph == sourceParagraph);
            if (sourceParagraphIndex < 0 || targetParagraphs.Count == 0) {
                return false;
            }

            int targetParagraphIndex = FindTargetGapByNeighborIdentity(
                sourceParagraphs,
                targetParagraphs,
                sourceParagraphIndex,
                paragraph => GetParagraphText(paragraph),
                paragraph => GetParagraphText(paragraph));
            targetParagraphIndex = Math.Max(0, Math.Min(targetParagraphIndex, targetParagraphs.Count - 1));
            targetParagraphs[targetParagraphIndex].Append(deleted);
            return true;
        }

        private static void AppendImageFallbackParagraph(OfficeXmlNode container, OfficeXmlNode paragraph) {
            if (container.Is(W, "body")) {
                OfficeXmlNode sectionProperties = container.Children(W, "sectPr").Last();
                if (!sectionProperties.IsNone) {
                    container.InsertBefore(paragraph, sectionProperties);
                    return;
                }
            }

            container.Append(paragraph);
        }

        private static RedlineImageContainer GetTargetImageContainer(RedlineImageEntry sourceEntry, IReadOnlyList<RedlineImageContainer> targetContainers) {
            RedlineImageContainer? matchingContainer = targetContainers.FirstOrDefault(container => string.Equals(container.PartKey, sourceEntry.PartKey, StringComparison.Ordinal));
            if (matchingContainer != null) {
                return matchingContainer;
            }

            if (targetContainers.Count > 0) {
                return targetContainers[0];
            }

            throw new InvalidOperationException("Target document has no part that can receive a deleted image redline.");
        }

        private static void InsertBeforeImageAnchor(OfficeXmlNode anchorRun, OfficeXmlNode deleted, Dictionary<OfficeXmlNode, OfficeXmlNode> insertedBeforeImageAnchors) {
            if (insertedBeforeImageAnchors.TryGetValue(anchorRun, out OfficeXmlNode previousInserted)) {
                deleted = previousInserted.InsertAfterSelf(deleted);
            } else {
                deleted = anchorRun.Parent.InsertBefore(deleted, anchorRun);
            }

            insertedBeforeImageAnchors[anchorRun] = deleted;
        }

        private static void InsertAfterImageAnchor(OfficeXmlNode anchorRun, OfficeXmlNode deleted, Dictionary<OfficeXmlNode, OfficeXmlNode> insertedAfterImageAnchors) {
            if (insertedAfterImageAnchors.TryGetValue(anchorRun, out OfficeXmlNode previousInserted)) {
                deleted = previousInserted.InsertAfterSelf(deleted);
            } else {
                deleted = anchorRun.InsertAfterSelf(deleted);
            }

            insertedAfterImageAnchors[anchorRun] = deleted;
        }

        /// <summary>A copy of the source image run whose relationships target copies in <paramref name="targetPart"/>, or none when one can't be copied.</summary>
        private static OfficeXmlNode CloneImageRunForPart(OfficeXmlNode context, RedlineImageEntry sourceEntry, WordPartView targetPart, WordPackageView targetDocument) {
            OfficeXmlNode clonedRun = CloneImageOnlyRun(context, sourceEntry);
            foreach (OfficeXmlNode blip in DescendantsOf(clonedRun, Dml, "blip")) {
                if (blip.Attribute(OfficeIMO.Core.Internal.OpenXml.OfficeXmlNamespaces.Relationships, "embed") is string embeddedRelationshipId) {
                    string? copiedRelationshipId = CopyEmbeddedImageRelationship(sourceEntry.Part, targetPart, embeddedRelationshipId);
                    if (copiedRelationshipId == null) {
                        return default;
                    }

                    blip.SetAttribute(OfficeIMO.Core.Internal.OpenXml.OfficeXmlNamespaces.Relationships, "embed", copiedRelationshipId);
                }

                if (blip.Attribute(OfficeIMO.Core.Internal.OpenXml.OfficeXmlNamespaces.Relationships, "link") is string externalRelationshipId) {
                    string? copiedRelationshipId = CopyExternalImageRelationship(sourceEntry.Part, targetPart, externalRelationshipId);
                    if (copiedRelationshipId == null) {
                        return default;
                    }

                    blip.SetAttribute(OfficeIMO.Core.Internal.OpenXml.OfficeXmlNamespaces.Relationships, "link", copiedRelationshipId);
                }
            }

            foreach (OfficeXmlNode imageData in DescendantsOf(clonedRun, Vml, "imagedata")) {
                if (imageData.Attribute(OfficeIMO.Core.Internal.OpenXml.OfficeXmlNamespaces.Relationships, "id") is not string relationshipId) {
                    continue;
                }

                string? copiedRelationshipId = sourceEntry.Part.HasExternalRelationship(relationshipId)
                    ? CopyExternalImageRelationship(sourceEntry.Part, targetPart, relationshipId)
                    : CopyEmbeddedImageRelationship(sourceEntry.Part, targetPart, relationshipId);
                if (copiedRelationshipId == null) {
                    return default;
                }

                imageData.SetAttribute(OfficeIMO.Core.Internal.OpenXml.OfficeXmlNamespaces.Relationships, "id", copiedRelationshipId);
            }

            if (!CopyAdditionalRunRelationships(sourceEntry.Part, targetPart, clonedRun)) {
                return default;
            }

            WordPackageView.ReassignDrawingIds(targetDocument.SdkDocument!, clonedRun);
            RefreshClonedVmlIds(targetDocument, clonedRun);
            return clonedRun;
        }

        private static OfficeXmlNode CloneImageOnlyRun(OfficeXmlNode context, RedlineImageEntry entry) {
            OfficeXmlNode clonedRun = context.CreateElement(W, "r");
            OfficeXmlNode properties = entry.Run.Child(W, "rPr");
            if (!properties.IsNone) {
                clonedRun.Append(properties.Clone());
            }

            clonedRun.Append(entry.ImageElement.Clone());
            return clonedRun;
        }

        private static void RemoveImageElementFromRun(RedlineImageEntry entry) {
            entry.ImageElement.Remove();
            if (!entry.Run.Children().ToList().Any(child => !child.Is(W, "rPr"))) {
                entry.Run.Remove();
            }
        }

        private static string GetRedlineImageIdentity(RedlineImageEntry entry) {
            OfficeXmlNode drawing = FirstDescendant(entry.Run, W, "drawing");
            if (!drawing.IsNone) {
                OfficeXmlNode blip = FirstDescendant(drawing, Dml, "blip");
                if (blip.Attribute(OfficeIMO.Core.Internal.OpenXml.OfficeXmlNamespaces.Relationships, "embed") is string embeddedRelationshipId) {
                    return entry.PartKey + "|embedded|" + GetImageFingerprintText(entry.Part, embeddedRelationshipId) + "|" + GetDrawingVisualSignature(entry.Part, drawing);
                }

                if (blip.Attribute(OfficeIMO.Core.Internal.OpenXml.OfficeXmlNamespaces.Relationships, "link") is string externalRelationshipId) {
                    return entry.PartKey + "|external|" + GetRelationshipTarget(entry.Part, externalRelationshipId) + "|" + GetDrawingVisualSignature(entry.Part, drawing);
                }
            }

            OfficeXmlNode imageData = FirstDescendant(entry.Run, Vml, "imagedata");
            if (imageData.Attribute(OfficeIMO.Core.Internal.OpenXml.OfficeXmlNamespaces.Relationships, "id") is string relationshipId) {
                string relationshipIdentity = entry.Part.HasExternalRelationship(relationshipId)
                    ? "external|" + GetRelationshipTarget(entry.Part, relationshipId)
                    : "embedded|" + GetImageFingerprintText(entry.Part, relationshipId);
                return entry.PartKey + "|" + relationshipIdentity + "|" + GetVmlVisualSignature(entry.Part, imageData);
            }

            return entry.PartKey + "|run|" + GetElementSignature(entry.Run);
        }

        private static string GetImageFingerprintText(WordPartView part, string relationshipId) {
            WordPartView? imagePart = part.GetPartById(relationshipId);
            if (imagePart != null && IsImagePartRelationship(part, relationshipId)) {
                ImageFingerprint fingerprint = CreateImageFingerprint(imagePart);
                return fingerprint.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + ":" + fingerprint.Sha256;
            }

            return relationshipId;
        }

        private static bool CopyAdditionalRunRelationships(WordPartView sourcePart, WordPartView targetPart, OfficeXmlNode clonedRun) {
            foreach (OfficeXmlNode element in Descendants(clonedRun)) {
                foreach (WordXmlAttribute attribute in WordXmlNodes.Attributes(element)) {
                    if (!string.Equals(attribute.NamespaceUri, OfficeRelationshipNamespace, StringComparison.Ordinal) ||
                        string.IsNullOrWhiteSpace(attribute.Value) ||
                        IsHandledImageRelationshipAttribute(element, attribute)) {
                        continue;
                    }

                    string? copiedRelationshipId = CopyRelationship(sourcePart, targetPart, attribute.Value);
                    if (copiedRelationshipId == null) {
                        return false;
                    }

                    WordXmlNodes.SetAttribute(element, attribute.Prefix, attribute.NamespaceUri, attribute.LocalName, copiedRelationshipId);
                }
            }

            return true;
        }

        private static bool IsHandledImageRelationshipAttribute(OfficeXmlNode element, WordXmlAttribute attribute) {
            return element.Is(Dml, "blip") &&
                   (string.Equals(attribute.LocalName, "embed", StringComparison.Ordinal) ||
                    string.Equals(attribute.LocalName, "link", StringComparison.Ordinal)) ||
                   element.Is(Vml, "imagedata") &&
                   string.Equals(attribute.LocalName, "id", StringComparison.Ordinal);
        }

        private static string? CopyRelationship(WordPartView sourcePart, WordPartView targetPart, string relationshipId) {
            if (sourcePart.HyperlinkTarget(relationshipId, out bool isExternal) != null) {
                return targetPart.AddHyperlinkRelationship(sourcePart.RawTarget(relationshipId) ?? string.Empty, isExternal);
            }

            if (sourcePart.TryGetExternalRelationship(relationshipId, out string relationshipType, out _)) {
                return targetPart.AddExternalRelationship(relationshipType, sourcePart.RawTarget(relationshipId) ?? string.Empty);
            }

            return sourcePart.GetPartById(relationshipId) != null && IsImagePartRelationship(sourcePart, relationshipId)
                ? CopyEmbeddedImageRelationship(sourcePart, targetPart, relationshipId)
                : null;
        }

        private static string? CopyEmbeddedImageRelationship(WordPartView sourcePart, WordPartView targetPart, string relationshipId) {
            WordPartView? sourceImagePart = sourcePart.GetPartById(relationshipId);
            if (sourceImagePart == null || !IsImagePartRelationship(sourcePart, relationshipId)) {
                return null;
            }

            return targetPart.AddImagePart(sourceImagePart.ContentType ?? string.Empty, sourceImagePart.ReadBytes());
        }

        private static string? CopyExternalImageRelationship(WordPartView sourcePart, WordPartView targetPart, string relationshipId) {
            if (!sourcePart.TryGetExternalRelationship(relationshipId, out string relationshipType, out _)) {
                return null;
            }

            return targetPart.AddExternalRelationship(relationshipType, sourcePart.RawTarget(relationshipId) ?? string.Empty);
        }

        private static void RefreshClonedVmlIds(WordPackageView targetDocument, OfficeXmlNode clonedRun) {
            HashSet<string> usedIds = GetVmlShapeIds(targetDocument);
            int nextId = usedIds.Count + 1;
            foreach (OfficeXmlNode shape in DescendantsOf(clonedRun, Vml, "shape")) {
                string candidate;
                do {
                    candidate = "OfficeIMOImageRedline" + nextId.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    nextId++;
                } while (usedIds.Contains(candidate));

                shape.SetAttribute("id", candidate);
                usedIds.Add(candidate);
            }
        }

        private static HashSet<string> GetVmlShapeIds(WordPackageView document) {
            var ids = new HashSet<string>(StringComparer.Ordinal);
            WordPartView? mainPart = document.Main;
            AddVmlShapeIds(ids, DocumentRoot(mainPart));

            if (mainPart != null) {
                foreach (WordPartView headerPart in HeaderParts(mainPart)) {
                    AddVmlShapeIds(ids, HeaderRoot(headerPart));
                }

                foreach (WordPartView footerPart in FooterParts(mainPart)) {
                    AddVmlShapeIds(ids, FooterRoot(footerPart));
                }

                AddVmlShapeIds(ids, FootnotesRoot(FootnotesPart(mainPart)));
                AddVmlShapeIds(ids, EndnotesRoot(EndnotesPart(mainPart)));
                AddVmlShapeIds(ids, mainPart.FirstPart(WordPackageView.CommentsRelationshipType)?.RootNamed(W, "comments") ?? default);
            }

            return ids;
        }

        private static void AddVmlShapeIds(HashSet<string> ids, OfficeXmlNode root) {
            if (root.IsNone) {
                return;
            }

            foreach (OfficeXmlNode shape in DescendantsOf(root, Vml, "shape")) {
                string? shapeId = shape.Attribute("id");
                if (!string.IsNullOrEmpty(shapeId)) {
                    ids.Add(shapeId!);
                }
            }
        }

        private static OfficeXmlNode WrapRunAsInserted(OfficeXmlNode context, OfficeXmlNode run, WordComparisonRedlineOptions options) {
            OfficeXmlNode inserted = context.CreateElement(W, "ins");
            SetRevisionAttributes(inserted, options.Author, options.DateTime ?? DateTime.Now);
            inserted.Append(run);
            return inserted;
        }

        private static OfficeXmlNode WrapRunAsDeleted(OfficeXmlNode context, OfficeXmlNode run, WordComparisonRedlineOptions options) {
            OfficeXmlNode deleted = context.CreateElement(W, "del");
            SetRevisionAttributes(deleted, options.Author, options.DateTime ?? DateTime.Now);
            deleted.Append(run);
            return deleted;
        }

        private static int GetSourceImageIndex(WordComparisonFinding finding) {
            return finding.SourceIndex ?? -1;
        }

        private static int GetTargetImageIndex(WordComparisonFinding finding) {
            int targetIndex = finding.TargetIndex ?? -1;
            if (targetIndex < 0) {
                TryParseIndexedLocation(finding.Location, "image", out targetIndex);
            }

            return targetIndex;
        }

        private sealed class RedlineImageContainer {
            internal RedlineImageContainer(int index, string partKey, WordPartView part, OfficeXmlNode container, int orderBase) {
                Index = index;
                PartKey = partKey;
                Part = part;
                Container = container;
                OrderBase = orderBase;
            }

            internal int Index { get; }

            internal string PartKey { get; }

            internal WordPartView Part { get; }

            internal OfficeXmlNode Container { get; }

            internal int OrderBase { get; }
        }

        private sealed class RedlineImageEntry {
            internal RedlineImageEntry(int index, int containerIndex, string partKey, int containerImageIndex, WordPartView part, OfficeXmlNode container, OfficeXmlNode run, OfficeXmlNode imageElement) {
                Index = index;
                ContainerIndex = containerIndex;
                PartKey = partKey;
                ContainerImageIndex = containerImageIndex;
                Part = part;
                Container = container;
                Run = run;
                ImageElement = imageElement;
            }

            internal int Index { get; }

            internal int ContainerIndex { get; }

            internal string PartKey { get; }

            internal int ContainerImageIndex { get; }

            internal WordPartView Part { get; }

            internal OfficeXmlNode Container { get; }

            internal OfficeXmlNode Run { get; }

            internal OfficeXmlNode ImageElement { get; }
        }
    }
}
