using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Wordprocessing;
using V = DocumentFormat.OpenXml.Vml;

namespace OfficeIMO.Word {
    public static partial class WordDocumentComparer {
        private const string OfficeRelationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        private static void ApplyImageFindings(WordprocessingDocument sourceDocument, WordprocessingDocument targetDocument, WordComparisonResult result, WordComparisonRedlineOptions options) {
            List<RedlineImageContainer> targetContainers = GetRedlineImageContainers(targetDocument);
            List<RedlineImageEntry> sourceImages = GetRedlineImageEntries(sourceDocument);
            List<RedlineImageEntry> targetImages = GetRedlineImageEntries(targetDocument);
            var rewrittenTargetImages = new HashSet<int>();
            var insertedSourceImages = new HashSet<int>();
            var insertedBeforeImageAnchors = new Dictionary<Run, OpenXmlElement>();
            var insertedAfterImageAnchors = new Dictionary<Run, OpenXmlElement>();

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

        private static List<RedlineImageEntry> GetRedlineImageEntries(WordprocessingDocument document) {
            var entries = new List<RedlineImageEntry>();
            foreach (RedlineImageContainer container in GetRedlineImageContainers(document)) {
                AddRedlineImageEntries(entries, container);
            }

            return entries;
        }

        private static List<RedlineImageContainer> GetRedlineImageContainers(WordprocessingDocument document) {
            MainDocumentPart? mainPart = document.MainDocumentPart;
            var containers = new List<RedlineImageContainer>();
            if (mainPart?.Document?.Body != null) {
                containers.Add(new RedlineImageContainer(containers.Count, BodyPartKey, mainPart, mainPart.Document.Body, BodyPartOrderBase));
            }

            if (mainPart == null) {
                return containers;
            }

            int headerIndex = 0;
            foreach (KeyValuePair<HeaderPart, string> headerPartKey in CreateOrderedHeaderPartKeys(mainPart)) {
                if (headerPartKey.Key.Header != null) {
                    containers.Add(new RedlineImageContainer(containers.Count, headerPartKey.Value, headerPartKey.Key, headerPartKey.Key.Header, HeaderPartOrderBase + (headerIndex * RelatedPartOrderStride)));
                }

                headerIndex++;
            }

            int footerIndex = 0;
            foreach (KeyValuePair<FooterPart, string> footerPartKey in CreateOrderedFooterPartKeys(mainPart)) {
                if (footerPartKey.Key.Footer != null) {
                    containers.Add(new RedlineImageContainer(containers.Count, footerPartKey.Value, footerPartKey.Key, footerPartKey.Key.Footer, FooterPartOrderBase + (footerIndex * RelatedPartOrderStride)));
                }

                footerIndex++;
            }

            List<Footnote> footnotes = GetReferencedFootnotes(mainPart);
            for (int footnoteIndex = 0; footnoteIndex < footnotes.Count; footnoteIndex++) {
                if (mainPart.FootnotesPart != null) {
                    string noteId = GetNotePartKeyId(footnotes[footnoteIndex], footnoteIndex);
                    containers.Add(new RedlineImageContainer(containers.Count, FootnotePartKeyPrefix + noteId, mainPart.FootnotesPart, footnotes[footnoteIndex], FootnotePartOrderBase + (footnoteIndex * RelatedPartOrderStride)));
                }
            }

            List<Endnote> endnotes = GetReferencedEndnotes(mainPart);
            for (int endnoteIndex = 0; endnoteIndex < endnotes.Count; endnoteIndex++) {
                if (mainPart.EndnotesPart != null) {
                    string noteId = GetNotePartKeyId(endnotes[endnoteIndex], endnoteIndex);
                    containers.Add(new RedlineImageContainer(containers.Count, EndnotePartKeyPrefix + noteId, mainPart.EndnotesPart, endnotes[endnoteIndex], EndnotePartOrderBase + (endnoteIndex * RelatedPartOrderStride)));
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

                Run? run = ordered.Element.Ancestors<Run>().FirstOrDefault();
                if (run == null) {
                    continue;
                }

                OpenXmlElement imageElement = GetImageRedlineElement(ordered.Element);
                entries.Add(new RedlineImageEntry(entries.Count, container.Index, container.PartKey, containerImageIndex, container.Part, container.Container, run, imageElement));
                containerImageIndex++;
            }
        }

        private static OpenXmlElement GetImageRedlineElement(OpenXmlElement element) {
            if (element is V.ImageData) {
                return element.Ancestors<Picture>().FirstOrDefault() ?? element;
            }

            return element;
        }

        private static bool IsRedlineImageElement(OpenXmlElement element) {
            return element is V.ImageData ||
                element is DocumentFormat.OpenXml.Wordprocessing.Drawing drawing &&
                drawing.Descendants<A.Blip>().Any();
        }

        private static bool WrapImageRunAsInserted(RedlineImageEntry entry, WordComparisonRedlineOptions options) {
            Run imageRun = entry.Run;
            OpenXmlElement? parent = imageRun.Parent;
            if (parent == null || imageRun.Ancestors<InsertedRun>().Any() || imageRun.Ancestors<DeletedRun>().Any()) {
                return false;
            }

            var inserted = new InsertedRun {
                Author = options.Author,
                Date = options.DateTime ?? DateTime.Now,
                Id = WordHeadersAndFooters.GenerateRevisionId()
            };
            Run insertedRun = CloneImageOnlyRun(entry);
            inserted.Append(insertedRun);
            parent.InsertBefore(inserted, imageRun);
            RemoveImageElementFromRun(entry);
            return true;
        }

        private static bool WrapImageRunAsChanged(RedlineImageEntry sourceEntry, RedlineImageEntry targetEntry, WordprocessingDocument targetDocument, WordComparisonRedlineOptions options) {
            Run targetRun = targetEntry.Run;
            OpenXmlElement? parent = targetRun.Parent;
            if (parent == null || targetRun.Ancestors<InsertedRun>().Any() || targetRun.Ancestors<DeletedRun>().Any()) {
                return false;
            }

            Run? deletedRun = CloneImageRunForPart(sourceEntry, targetEntry.Part, targetDocument);
            if (deletedRun == null) {
                return false;
            }

            parent.InsertBefore(WrapRunAsDeleted(deletedRun, options), targetRun);
            parent.InsertBefore(WrapRunAsInserted(CloneImageOnlyRun(targetEntry), options), targetRun);
            RemoveImageElementFromRun(targetEntry);
            return true;
        }

        private static bool InsertDeletedImageRun(
            IReadOnlyList<RedlineImageEntry> sourceImages,
            RedlineImageEntry sourceEntry,
            IReadOnlyList<RedlineImageContainer> targetContainers,
            IReadOnlyList<RedlineImageEntry> targetImages,
            WordprocessingDocument targetDocument,
            WordComparisonRedlineOptions options,
            Dictionary<Run, OpenXmlElement> insertedBeforeImageAnchors,
            Dictionary<Run, OpenXmlElement> insertedAfterImageAnchors) {
            RedlineImageContainer targetContainer = GetTargetImageContainer(sourceEntry, targetContainers);
            Run? deletedRun = CloneImageRunForPart(sourceEntry, targetContainer.Part, targetDocument);
            if (deletedRun == null) {
                return false;
            }

            DeletedRun deleted = WrapRunAsDeleted(deletedRun, options);
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
            if (nextImage?.Run.Parent != null) {
                InsertBeforeImageAnchor(nextImage.Run, deleted, insertedBeforeImageAnchors);
                return true;
            }

            RedlineImageEntry? previousImage = targetImages
                .Where(image => string.Equals(image.PartKey, targetContainer.PartKey, StringComparison.Ordinal) && image.ContainerImageIndex < targetImageIndex)
                .OrderByDescending(image => image.ContainerImageIndex)
                .FirstOrDefault();
            if (previousImage?.Run.Parent != null) {
                InsertAfterImageAnchor(previousImage.Run, deleted, insertedAfterImageAnchors);
                return true;
            }

            if (TryInsertDeletedImageAtSourceParagraphGap(sourceEntry, targetContainer, deleted)) {
                return true;
            }

            var paragraph = new Paragraph(deleted);
            AppendImageFallbackParagraph(targetContainer.Container, paragraph);
            return true;
        }

        private static bool TryInsertDeletedImageAtSourceParagraphGap(RedlineImageEntry sourceEntry, RedlineImageContainer targetContainer, DeletedRun deleted) {
            Paragraph? sourceParagraph = sourceEntry.Run.Ancestors<Paragraph>().FirstOrDefault();
            if (sourceParagraph == null) {
                return false;
            }

            List<Paragraph> sourceParagraphs = sourceEntry.Container.Descendants<Paragraph>().ToList();
            List<Paragraph> targetParagraphs = targetContainer.Container.Descendants<Paragraph>().ToList();
            int sourceParagraphIndex = sourceParagraphs.FindIndex(paragraph => ReferenceEquals(paragraph, sourceParagraph));
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

        private static void AppendImageFallbackParagraph(OpenXmlElement container, Paragraph paragraph) {
            if (container is Body body && body.Elements<SectionProperties>().LastOrDefault() is SectionProperties sectionProperties) {
                body.InsertBefore(paragraph, sectionProperties);
                return;
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

        private static void InsertBeforeImageAnchor(Run anchorRun, OpenXmlElement deleted, Dictionary<Run, OpenXmlElement> insertedBeforeImageAnchors) {
            if (insertedBeforeImageAnchors.TryGetValue(anchorRun, out OpenXmlElement? previousInserted)) {
                previousInserted.InsertAfterSelf(deleted);
            } else {
                anchorRun.Parent!.InsertBefore(deleted, anchorRun);
            }

            insertedBeforeImageAnchors[anchorRun] = deleted;
        }

        private static void InsertAfterImageAnchor(Run anchorRun, OpenXmlElement deleted, Dictionary<Run, OpenXmlElement> insertedAfterImageAnchors) {
            if (insertedAfterImageAnchors.TryGetValue(anchorRun, out OpenXmlElement? previousInserted)) {
                previousInserted.InsertAfterSelf(deleted);
            } else {
                anchorRun.Parent!.InsertAfter(deleted, anchorRun);
            }

            insertedAfterImageAnchors[anchorRun] = deleted;
        }

        private static Run? CloneImageRunForPart(RedlineImageEntry sourceEntry, OpenXmlPart targetPart, WordprocessingDocument targetDocument) {
            var clonedRun = CloneImageOnlyRun(sourceEntry);
            foreach (A.Blip blip in clonedRun.Descendants<A.Blip>()) {
                if (blip.Embed?.Value is string embeddedRelationshipId) {
                    string? copiedRelationshipId = CopyEmbeddedImageRelationship(sourceEntry.Part, targetPart, embeddedRelationshipId);
                    if (copiedRelationshipId == null) {
                        return null;
                    }

                    blip.Embed = copiedRelationshipId;
                }

                if (blip.Link?.Value is string externalRelationshipId) {
                    string? copiedRelationshipId = CopyExternalImageRelationship(sourceEntry.Part, targetPart, externalRelationshipId);
                    if (copiedRelationshipId == null) {
                        return null;
                    }

                    blip.Link = copiedRelationshipId;
                }
            }

            foreach (V.ImageData imageData in clonedRun.Descendants<V.ImageData>()) {
                if (imageData.RelationshipId?.Value is not string relationshipId) {
                    continue;
                }

                string? copiedRelationshipId = sourceEntry.Part.ExternalRelationships.Any(relationship => relationship.Id == relationshipId)
                    ? CopyExternalImageRelationship(sourceEntry.Part, targetPart, relationshipId)
                    : CopyEmbeddedImageRelationship(sourceEntry.Part, targetPart, relationshipId);
                if (copiedRelationshipId == null) {
                    return null;
                }

                imageData.RelationshipId = copiedRelationshipId;
            }

            if (!CopyAdditionalRunRelationships(sourceEntry.Part, targetPart, clonedRun)) {
                return null;
            }

            RefreshClonedDrawingIds(targetDocument, clonedRun);
            RefreshClonedVmlIds(targetDocument, clonedRun);
            return clonedRun;
        }

        private static Run CloneImageOnlyRun(RedlineImageEntry entry) {
            var clonedRun = new Run();
            RunProperties? properties = entry.Run.GetFirstChild<RunProperties>();
            if (properties != null) {
                clonedRun.Append((RunProperties)properties.CloneNode(true));
            }

            clonedRun.Append(entry.ImageElement.CloneNode(true));
            return clonedRun;
        }

        private static void RemoveImageElementFromRun(RedlineImageEntry entry) {
            entry.ImageElement.Remove();
            if (!entry.Run.ChildElements.Any(child => child is not RunProperties)) {
                entry.Run.Remove();
            }
        }

        private static string GetRedlineImageIdentity(RedlineImageEntry entry) {
            DocumentFormat.OpenXml.Wordprocessing.Drawing? drawing = entry.Run.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>().FirstOrDefault();
            if (drawing != null) {
                A.Blip? blip = drawing.Descendants<A.Blip>().FirstOrDefault();
                if (blip?.Embed?.Value is string embeddedRelationshipId) {
                    return entry.PartKey + "|embedded|" + GetImageFingerprintText(entry.Part, embeddedRelationshipId) + "|" + GetDrawingVisualSignature(entry.Part, drawing);
                }

                if (blip?.Link?.Value is string externalRelationshipId) {
                    return entry.PartKey + "|external|" + GetRelationshipTarget(entry.Part, externalRelationshipId) + "|" + GetDrawingVisualSignature(entry.Part, drawing);
                }
            }

            V.ImageData? imageData = entry.Run.Descendants<V.ImageData>().FirstOrDefault();
            if (imageData?.RelationshipId?.Value is string relationshipId) {
                string relationshipIdentity = entry.Part.ExternalRelationships.Any(relationship => relationship.Id == relationshipId)
                    ? "external|" + GetRelationshipTarget(entry.Part, relationshipId)
                    : "embedded|" + GetImageFingerprintText(entry.Part, relationshipId);
                return entry.PartKey + "|" + relationshipIdentity + "|" + GetVmlVisualSignature(entry.Part, imageData);
            }

            return entry.PartKey + "|run|" + entry.Run.OuterXml;
        }

        private static string GetImageFingerprintText(OpenXmlPart part, string relationshipId) {
            try {
                if (part.GetPartById(relationshipId) is ImagePart imagePart) {
                    using Stream stream = imagePart.GetStream(FileMode.Open, FileAccess.Read);
                    ImageFingerprint fingerprint = CreateImageFingerprint(stream);
                    return fingerprint.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + ":" + fingerprint.Sha256;
                }
            } catch (ArgumentOutOfRangeException) {
            }

            return relationshipId;
        }

        private static bool CopyAdditionalRunRelationships(OpenXmlPart sourcePart, OpenXmlPart targetPart, Run clonedRun) {
            foreach (OpenXmlElement element in clonedRun.Descendants<OpenXmlElement>()) {
                foreach (OpenXmlAttribute attribute in element.GetAttributes()) {
                    if (!string.Equals(attribute.NamespaceUri, OfficeRelationshipNamespace, StringComparison.Ordinal) ||
                        string.IsNullOrWhiteSpace(attribute.Value) ||
                        IsHandledImageRelationshipAttribute(element, attribute)) {
                        continue;
                    }

                    string relationshipId = attribute.Value!;
                    string? copiedRelationshipId = CopyRelationship(sourcePart, targetPart, relationshipId);
                    if (copiedRelationshipId == null) {
                        return false;
                    }

                    element.SetAttribute(new OpenXmlAttribute(attribute.Prefix, attribute.LocalName, attribute.NamespaceUri, copiedRelationshipId));
                }
            }

            return true;
        }

        private static bool IsHandledImageRelationshipAttribute(OpenXmlElement element, OpenXmlAttribute attribute) {
            return element is A.Blip &&
                   (string.Equals(attribute.LocalName, "embed", StringComparison.Ordinal) ||
                    string.Equals(attribute.LocalName, "link", StringComparison.Ordinal)) ||
                   element is V.ImageData &&
                   string.Equals(attribute.LocalName, "id", StringComparison.Ordinal);
        }

        private static string? CopyRelationship(OpenXmlPart sourcePart, OpenXmlPart targetPart, string relationshipId) {
            HyperlinkRelationship? hyperlinkRelationship = sourcePart.HyperlinkRelationships.FirstOrDefault(relationship => relationship.Id == relationshipId);
            if (hyperlinkRelationship != null) {
                HyperlinkRelationship targetRelationship = targetPart.AddHyperlinkRelationship(hyperlinkRelationship.Uri, hyperlinkRelationship.IsExternal);
                return targetRelationship.Id;
            }

            ExternalRelationship? externalRelationship = sourcePart.ExternalRelationships.FirstOrDefault(relationship => relationship.Id == relationshipId);
            if (externalRelationship != null) {
                ExternalRelationship targetRelationship = targetPart.AddExternalRelationship(externalRelationship.RelationshipType, externalRelationship.Uri);
                return targetRelationship.Id;
            }

            try {
                return sourcePart.GetPartById(relationshipId) is ImagePart
                    ? CopyEmbeddedImageRelationship(sourcePart, targetPart, relationshipId)
                    : null;
            } catch (ArgumentOutOfRangeException) {
                return null;
            }
        }

        private static string? CopyEmbeddedImageRelationship(OpenXmlPart sourcePart, OpenXmlPart targetPart, string relationshipId) {
            OpenXmlPart sourceRelatedPart;
            try {
                sourceRelatedPart = sourcePart.GetPartById(relationshipId);
            } catch (ArgumentOutOfRangeException) {
                return null;
            }

            if (sourceRelatedPart is not ImagePart sourceImagePart) {
                return null;
            }

            ImagePart targetImagePart = AddImagePart(targetPart, sourceImagePart.ContentType);
            using Stream sourceStream = sourceImagePart.GetStream(FileMode.Open, FileAccess.Read);
            targetImagePart.FeedData(sourceStream);
            return targetPart.GetIdOfPart(targetImagePart);
        }

        private static string? CopyExternalImageRelationship(OpenXmlPart sourcePart, OpenXmlPart targetPart, string relationshipId) {
            ExternalRelationship? sourceRelationship = sourcePart.ExternalRelationships.FirstOrDefault(relationship => relationship.Id == relationshipId);
            if (sourceRelationship == null) {
                return null;
            }

            ExternalRelationship targetRelationship = targetPart.AddExternalRelationship(sourceRelationship.RelationshipType, sourceRelationship.Uri);
            return targetRelationship.Id;
        }

        private static void RefreshClonedDrawingIds(WordprocessingDocument targetDocument, Run clonedRun) {
            uint nextId = GetNextDrawingDocPropertiesId(targetDocument);
            foreach (DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties properties in clonedRun.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties>()) {
                properties.Id = nextId;
                nextId++;
            }

            foreach (DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties properties in clonedRun.Descendants<DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties>()) {
                properties.Id = nextId;
                nextId++;
            }
        }

        private static uint GetNextDrawingDocPropertiesId(WordprocessingDocument document) {
            uint max = 0U;
            MainDocumentPart? mainPart = document.MainDocumentPart;
            UpdateMaxDrawingDocPropertiesId(mainPart?.Document, ref max);

            if (mainPart != null) {
                foreach (HeaderPart headerPart in mainPart.HeaderParts) {
                    UpdateMaxDrawingDocPropertiesId(headerPart.Header, ref max);
                }

                foreach (FooterPart footerPart in mainPart.FooterParts) {
                    UpdateMaxDrawingDocPropertiesId(footerPart.Footer, ref max);
                }

                UpdateMaxDrawingDocPropertiesId(mainPart.FootnotesPart?.Footnotes, ref max);
                UpdateMaxDrawingDocPropertiesId(mainPart.EndnotesPart?.Endnotes, ref max);
                UpdateMaxDrawingDocPropertiesId(mainPart.WordprocessingCommentsPart?.Comments, ref max);
            }

            return max + 1U;
        }

        private static void UpdateMaxDrawingDocPropertiesId(OpenXmlElement? root, ref uint max) {
            if (root == null) {
                return;
            }

            foreach (DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties properties in root.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties>()) {
                if (properties.Id != null && properties.Id.Value > max) {
                    max = properties.Id.Value;
                }
            }

            foreach (DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties properties in root.Descendants<DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties>()) {
                if (properties.Id != null && properties.Id.Value > max) {
                    max = properties.Id.Value;
                }
            }
        }

        private static void RefreshClonedVmlIds(WordprocessingDocument targetDocument, Run clonedRun) {
            HashSet<string> usedIds = GetVmlShapeIds(targetDocument);
            int nextId = usedIds.Count + 1;
            foreach (V.Shape shape in clonedRun.Descendants<V.Shape>()) {
                string candidate;
                do {
                    candidate = "OfficeIMOImageRedline" + nextId.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    nextId++;
                } while (usedIds.Contains(candidate));

                shape.Id = candidate;
                usedIds.Add(candidate);
            }
        }

        private static HashSet<string> GetVmlShapeIds(WordprocessingDocument document) {
            var ids = new HashSet<string>(StringComparer.Ordinal);
            MainDocumentPart? mainPart = document.MainDocumentPart;
            AddVmlShapeIds(ids, mainPart?.Document);

            if (mainPart != null) {
                foreach (HeaderPart headerPart in mainPart.HeaderParts) {
                    AddVmlShapeIds(ids, headerPart.Header);
                }

                foreach (FooterPart footerPart in mainPart.FooterParts) {
                    AddVmlShapeIds(ids, footerPart.Footer);
                }

                AddVmlShapeIds(ids, mainPart.FootnotesPart?.Footnotes);
                AddVmlShapeIds(ids, mainPart.EndnotesPart?.Endnotes);
                AddVmlShapeIds(ids, mainPart.WordprocessingCommentsPart?.Comments);
            }

            return ids;
        }

        private static void AddVmlShapeIds(HashSet<string> ids, OpenXmlElement? root) {
            if (root == null) {
                return;
            }

            foreach (V.Shape shape in root.Descendants<V.Shape>()) {
                string? shapeId = shape.Id?.Value;
                if (!string.IsNullOrEmpty(shapeId)) {
                    ids.Add(shapeId!);
                }
            }
        }

        private static ImagePart AddImagePart(OpenXmlPart part, string contentType) {
            return part switch {
                MainDocumentPart mainDocumentPart => mainDocumentPart.AddImagePart(contentType),
                HeaderPart headerPart => headerPart.AddImagePart(contentType),
                FooterPart footerPart => footerPart.AddImagePart(contentType),
                FootnotesPart footnotesPart => footnotesPart.AddImagePart(contentType),
                EndnotesPart endnotesPart => endnotesPart.AddImagePart(contentType),
                _ => throw new InvalidOperationException("Images cannot be copied into this document part.")
            };
        }

        private static InsertedRun WrapRunAsInserted(Run run, WordComparisonRedlineOptions options) {
            var inserted = new InsertedRun {
                Author = options.Author,
                Date = options.DateTime ?? DateTime.Now,
                Id = WordHeadersAndFooters.GenerateRevisionId()
            };
            inserted.Append(run);
            return inserted;
        }

        private static DeletedRun WrapRunAsDeleted(Run run, WordComparisonRedlineOptions options) {
            var deleted = new DeletedRun {
                Author = options.Author,
                Date = options.DateTime ?? DateTime.Now,
                Id = WordHeadersAndFooters.GenerateRevisionId()
            };
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
            internal RedlineImageContainer(int index, string partKey, OpenXmlPart part, OpenXmlElement container, int orderBase) {
                Index = index;
                PartKey = partKey;
                Part = part;
                Container = container;
                OrderBase = orderBase;
            }

            internal int Index { get; }

            internal string PartKey { get; }

            internal OpenXmlPart Part { get; }

            internal OpenXmlElement Container { get; }

            internal int OrderBase { get; }
        }

        private sealed class RedlineImageEntry {
            internal RedlineImageEntry(int index, int containerIndex, string partKey, int containerImageIndex, OpenXmlPart part, OpenXmlElement container, Run run, OpenXmlElement imageElement) {
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

            internal OpenXmlPart Part { get; }

            internal OpenXmlElement Container { get; }

            internal Run Run { get; }

            internal OpenXmlElement ImageElement { get; }
        }
    }
}
