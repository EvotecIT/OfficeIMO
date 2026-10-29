using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Threaded = DocumentFormat.OpenXml.Office2019.Excel.ThreadedComments;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private const int VmlShapeIdBlockSize = 1024;

        private static readonly Regex VmlIdMapDataPattern = new Regex(
            @"(<(?:[A-Za-z_][\w.-]*:)?idmap\b[^>]*?\bdata\s*=\s*[""'])([^""']*)([""'])",
            RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

        private static readonly Regex VmlShapeIdPattern = new Regex(
            @"_x0000_s(\d+)",
            RegexOptions.CultureInvariant);

        // ISO-8859-1 maps every byte to one character, so VML (which is often not
        // well-formed XML) round-trips byte for byte while ASCII ids are rewritten.
        private static readonly Encoding VmlByteTransparentEncoding = Encoding.GetEncoding(28591);

        /// <summary>
        /// Clones the parts behind drawings, legacy and threaded comments, header/footer VML,
        /// background pictures and external hyperlinks into the copied worksheet part and
        /// points the copied worksheet XML at the new relationship ids.
        /// References whose source relationship cannot be resolved are removed.
        /// </summary>
        private void CopyRelationshipBackedWorksheetFeatures(ExcelDocument sourceDocument, ExcelSheet sourceSheet, WorksheetPart copiedPart) {
            WorksheetPart sourcePart = sourceSheet.WorksheetPart;
            Worksheet worksheet = copiedPart.Worksheet ?? throw new InvalidOperationException("Worksheet is missing.");
            sourceSheet.CommitCommentVmlChanges();
            var cloner = new RelatedPartCloner();

            foreach (DocumentFormat.OpenXml.Spreadsheet.Drawing drawing in worksheet.Elements<DocumentFormat.OpenXml.Spreadsheet.Drawing>().ToList()) {
                string? id = CloneWorksheetRelatedPart<DrawingsPart>(sourcePart, copiedPart, drawing.Id?.Value, cloner);
                if (id == null) {
                    drawing.Remove();
                } else {
                    drawing.Id = id;
                }
            }

            foreach (LegacyDrawing legacyDrawing in worksheet.Elements<LegacyDrawing>().ToList()) {
                string? id = CloneWorksheetRelatedPart<VmlDrawingPart>(sourcePart, copiedPart, legacyDrawing.Id?.Value, cloner);
                if (id == null) {
                    legacyDrawing.Remove();
                } else {
                    legacyDrawing.Id = id;
                }
            }

            foreach (LegacyDrawingHeaderFooter legacyDrawing in worksheet.Elements<LegacyDrawingHeaderFooter>().ToList()) {
                string? id = CloneWorksheetRelatedPart<VmlDrawingPart>(sourcePart, copiedPart, legacyDrawing.Id?.Value, cloner);
                if (id == null) {
                    legacyDrawing.Remove();
                } else {
                    legacyDrawing.Id = id;
                }
            }

            foreach (Picture picture in worksheet.Elements<Picture>().ToList()) {
                string? id = CloneWorksheetRelatedPart<ImagePart>(sourcePart, copiedPart, picture.Id?.Value, cloner);
                if (id == null) {
                    picture.Remove();
                } else {
                    picture.Id = id;
                }
            }

            WorksheetCommentsPart? sourceCommentsPart = sourcePart.WorksheetCommentsPart;
            if (sourceCommentsPart != null) {
                cloner.Clone(sourceCommentsPart, copiedPart, MakeUnusedRelationshipId(copiedPart));
            }

            foreach (WorksheetThreadedCommentsPart sourceThreadedCommentsPart in sourcePart.WorksheetThreadedCommentsParts) {
                cloner.Clone(sourceThreadedCommentsPart, copiedPart, MakeUnusedRelationshipId(copiedPart));
            }

            CopyWorksheetHyperlinkRelationships(sourcePart, copiedPart, worksheet);
            RemoveCopiedSlicerAndTimelineFrames(copiedPart);
            IReadOnlyDictionary<VmlDrawingPart, Dictionary<int, int>> shapeBlockMaps = RenumberCopiedVmlShapeIds(copiedPart);
            RemapCopiedDrawingVmlShapeReferences(copiedPart, worksheet, shapeBlockMaps);
            RemapCopiedThreadedComments(sourceDocument, copiedPart);
        }

        /// <summary>
        /// Removes drawing anchors for slicers and timelines, together with any
        /// <c>mc:AlternateContent</c> wrapper, because the copy does not carry the
        /// slicer and timeline parts those frames point to.
        /// </summary>
        private static void RemoveCopiedSlicerAndTimelineFrames(WorksheetPart copiedPart) {
            DocumentFormat.OpenXml.Drawing.Spreadsheet.WorksheetDrawing? drawing = copiedPart.DrawingsPart?.WorksheetDrawing;
            if (drawing == null) {
                return;
            }

            bool changed = false;
            foreach (OpenXmlElement anchor in drawing.ChildElements.ToList()) {
                // Match by name: frames inside mc:AlternateContent may be unknown elements rather than typed GraphicData.
                bool isSlicerOrTimeline = anchor.Descendants<OpenXmlElement>()
                    .Where(_ => string.Equals(_.LocalName, "graphicData", StringComparison.Ordinal))
                    .Select(_ => _.GetAttributes().FirstOrDefault(_ => string.Equals(_.LocalName, "uri", StringComparison.Ordinal)).Value)
                    .Any(_ => string.Equals(_, SlicerGraphicDataUri, StringComparison.OrdinalIgnoreCase)
                        || string.Equals(_, TimelineGraphicDataUri, StringComparison.OrdinalIgnoreCase));
                if (isSlicerOrTimeline) {
                    anchor.Remove();
                    changed = true;
                }
            }

            if (changed) {
                drawing.Save();
            }
        }

        /// <summary>
        /// Keeps drawing shapes that point at legacy VML shapes (<c>a14:compatExt spid</c>)
        /// in step with the renumbered VML shape ids.
        /// </summary>
        private static void RemapCopiedDrawingVmlShapeReferences(
            WorksheetPart copiedPart,
            Worksheet worksheet,
            IReadOnlyDictionary<VmlDrawingPart, Dictionary<int, int>> shapeBlockMaps) {
            DocumentFormat.OpenXml.Drawing.Spreadsheet.WorksheetDrawing? drawing = copiedPart.DrawingsPart?.WorksheetDrawing;
            string? legacyDrawingId = worksheet.GetFirstChild<LegacyDrawing>()?.Id?.Value;
            if (drawing == null
                || string.IsNullOrWhiteSpace(legacyDrawingId)
                || TryGetRelatedPart(copiedPart, legacyDrawingId!) is not VmlDrawingPart vmlPart
                || !shapeBlockMaps.TryGetValue(vmlPart, out Dictionary<int, int>? blockMap)
                || blockMap.Count == 0) {
                return;
            }

            bool changed = false;
            foreach (OpenXmlElement element in drawing.Descendants<OpenXmlElement>()) {
                foreach (OpenXmlAttribute attribute in element.GetAttributes()) {
                    if (!string.Equals(attribute.LocalName, "spid", StringComparison.Ordinal) || attribute.Value == null) {
                        continue;
                    }

                    Match match = VmlShapeIdPattern.Match(attribute.Value);
                    if (!match.Success
                        || match.Index != 0
                        || match.Length != attribute.Value.Length
                        || !int.TryParse(match.Groups[1].Value, NumberStyles.None, CultureInfo.InvariantCulture, out int shapeId)
                        || !blockMap.TryGetValue(shapeId / VmlShapeIdBlockSize, out int targetBlock)) {
                        continue;
                    }

                    int targetShapeId = (targetBlock * VmlShapeIdBlockSize) + (shapeId % VmlShapeIdBlockSize);
                    element.SetAttribute(new OpenXmlAttribute(
                        attribute.Prefix,
                        attribute.LocalName,
                        attribute.NamespaceUri,
                        "_x0000_s" + targetShapeId.ToString(CultureInfo.InvariantCulture)));
                    changed = true;
                }
            }

            if (changed) {
                drawing.Save();
            }
        }

        private static string? CloneWorksheetRelatedPart<TPart>(
            WorksheetPart sourcePart,
            WorksheetPart copiedPart,
            string? sourceRelationshipId,
            RelatedPartCloner cloner)
            where TPart : OpenXmlPart {
            if (string.IsNullOrWhiteSpace(sourceRelationshipId)
                || TryGetRelatedPart(sourcePart, sourceRelationshipId!) is not TPart related) {
                return null;
            }

            string relationshipId = MakeUnusedRelationshipId(copiedPart);
            cloner.Clone(related, copiedPart, relationshipId);
            return relationshipId;
        }

        private static OpenXmlPart? TryGetRelatedPart(OpenXmlPartContainer container, string relationshipId) {
            foreach (IdPartPair pair in container.Parts) {
                if (string.Equals(pair.RelationshipId, relationshipId, StringComparison.Ordinal)) {
                    return pair.OpenXmlPart;
                }
            }

            return null;
        }

        private static void CopyWorksheetHyperlinkRelationships(WorksheetPart sourcePart, WorksheetPart copiedPart, Worksheet worksheet) {
            var copiedRelationshipIds = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (Hyperlinks hyperlinks in worksheet.Elements<Hyperlinks>().ToList()) {
                foreach (Hyperlink hyperlink in hyperlinks.Elements<Hyperlink>().Where(_ => _.Id != null).ToList()) {
                    string sourceRelationshipId = hyperlink.Id!.Value ?? string.Empty;
                    if (!copiedRelationshipIds.TryGetValue(sourceRelationshipId, out string? relationshipId)) {
                        HyperlinkRelationship? sourceRelationship = sourcePart.HyperlinkRelationships
                            .FirstOrDefault(_ => string.Equals(_.Id, sourceRelationshipId, StringComparison.Ordinal));
                        if (sourceRelationship == null) {
                            hyperlink.Remove();
                            continue;
                        }

                        relationshipId = MakeUnusedRelationshipId(copiedPart);
                        copiedPart.AddHyperlinkRelationship(sourceRelationship.Uri, sourceRelationship.IsExternal, relationshipId);
                        copiedRelationshipIds.Add(sourceRelationshipId, relationshipId);
                    }

                    hyperlink.Id = relationshipId;
                }

                if (!hyperlinks.Elements<Hyperlink>().Any()) {
                    hyperlinks.Remove();
                }
            }
        }

        /// <summary>
        /// Moves each copied VML drawing onto shape-id blocks no other VML drawing in this
        /// workbook uses, so comment and header/footer shapes keep workbook-unique ids.
        /// </summary>
        private IReadOnlyDictionary<VmlDrawingPart, Dictionary<int, int>> RenumberCopiedVmlShapeIds(WorksheetPart copiedPart) {
            var blockMaps = new Dictionary<VmlDrawingPart, Dictionary<int, int>>();
            foreach (VmlDrawingPart vmlPart in copiedPart.VmlDrawingParts.ToList()) {
                string text = ReadVmlText(vmlPart);
                HashSet<int> usedBlocks = CollectUsedVmlShapeBlocks(vmlPart);
                int nextBlock = usedBlocks.Count == 0 ? 1 : usedBlocks.Max() + 1;
                var blockMap = new Dictionary<int, int>();
                blockMaps[vmlPart] = blockMap;
                int MapBlock(int sourceBlock) {
                    if (!blockMap.TryGetValue(sourceBlock, out int targetBlock)) {
                        targetBlock = nextBlock++;
                        blockMap.Add(sourceBlock, targetBlock);
                    }

                    return targetBlock;
                }

                string MapBlockList(string blockList) {
                    var mapped = new List<string>();
                    foreach (string item in blockList.Split(',')) {
                        string trimmed = item.Trim();
                        if (trimmed.Length == 0) {
                            continue;
                        }

                        if (int.TryParse(trimmed, NumberStyles.None, CultureInfo.InvariantCulture, out int block)) {
                            mapped.Add(MapBlock(block).ToString(CultureInfo.InvariantCulture));
                        } else {
                            mapped.Add(trimmed);
                        }
                    }

                    return string.Join(",", mapped);
                }

                string renumbered = VmlIdMapDataPattern.Replace(text, _ => _.Groups[1].Value + MapBlockList(_.Groups[2].Value) + _.Groups[3].Value);
                renumbered = VmlShapeIdPattern.Replace(renumbered, _ => {
                    if (!int.TryParse(_.Groups[1].Value, NumberStyles.None, CultureInfo.InvariantCulture, out int shapeId)) {
                        return _.Value;
                    }

                    int targetShapeId = (MapBlock(shapeId / VmlShapeIdBlockSize) * VmlShapeIdBlockSize) + (shapeId % VmlShapeIdBlockSize);
                    return "_x0000_s" + targetShapeId.ToString(CultureInfo.InvariantCulture);
                });

                if (string.Equals(renumbered, text, StringComparison.Ordinal)) {
                    continue;
                }

                using Stream stream = vmlPart.GetStream(FileMode.Create, FileAccess.Write);
                byte[] bytes = VmlByteTransparentEncoding.GetBytes(renumbered);
                stream.Write(bytes, 0, bytes.Length);
            }

            return blockMaps;
        }

        private HashSet<int> CollectUsedVmlShapeBlocks(VmlDrawingPart excludedPart) {
            var blocks = new HashSet<int>();
            var visited = new HashSet<VmlDrawingPart>();
            foreach (WorksheetPart worksheetPart in WorkbookPartRoot.WorksheetParts) {
                foreach (VmlDrawingPart vmlPart in worksheetPart.VmlDrawingParts) {
                    if (ReferenceEquals(vmlPart, excludedPart) || !visited.Add(vmlPart)) {
                        continue;
                    }

                    string text = ReadVmlText(vmlPart);
                    foreach (Match match in VmlIdMapDataPattern.Matches(text)) {
                        foreach (string item in match.Groups[2].Value.Split(',')) {
                            if (int.TryParse(item.Trim(), NumberStyles.None, CultureInfo.InvariantCulture, out int block)) {
                                blocks.Add(block);
                            }
                        }
                    }

                    foreach (Match match in VmlShapeIdPattern.Matches(text)) {
                        if (int.TryParse(match.Groups[1].Value, NumberStyles.None, CultureInfo.InvariantCulture, out int shapeId)) {
                            blocks.Add(shapeId / VmlShapeIdBlockSize);
                        }
                    }
                }
            }

            return blocks;
        }

        private static string ReadVmlText(VmlDrawingPart vmlPart) {
            using Stream stream = vmlPart.GetStream(FileMode.Open, FileAccess.Read);
            using var reader = new StreamReader(stream, VmlByteTransparentEncoding, detectEncodingFromByteOrderMarks: false);
            return reader.ReadToEnd();
        }

        /// <summary>
        /// Gives copied threaded comments new workbook-unique ids, keeps legacy comment
        /// placeholders linked to them and brings the people they reference into this workbook.
        /// </summary>
        private void RemapCopiedThreadedComments(ExcelDocument sourceDocument, WorksheetPart copiedPart) {
            var commentParts = copiedPart.WorksheetThreadedCommentsParts
                .Where(_ => _.ThreadedComments != null)
                .ToList();
            if (commentParts.Count == 0) {
                return;
            }

            var idMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (WorksheetThreadedCommentsPart part in commentParts) {
                foreach (Threaded.ThreadedComment comment in part.ThreadedComments!.Elements<Threaded.ThreadedComment>()) {
                    string? id = comment.Id?.Value;
                    if (!string.IsNullOrWhiteSpace(id) && !idMap.ContainsKey(id!)) {
                        idMap.Add(id!, "{" + Guid.NewGuid().ToString().ToUpperInvariant() + "}");
                    }
                }
            }

            var personIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (WorksheetThreadedCommentsPart part in commentParts) {
                foreach (Threaded.ThreadedComment comment in part.ThreadedComments!.Elements<Threaded.ThreadedComment>()) {
                    if (comment.Id?.Value is string id && idMap.TryGetValue(id, out string? newId)) {
                        comment.Id = newId;
                    }

                    if (comment.ParentId?.Value is string parentId && idMap.TryGetValue(parentId, out string? newParentId)) {
                        comment.ParentId = newParentId;
                    }

                    if (!string.IsNullOrWhiteSpace(comment.PersonId?.Value)) {
                        personIds.Add(comment.PersonId!.Value!);
                    }

                    foreach (Threaded.Mention mention in comment.Descendants<Threaded.Mention>()) {
                        if (!string.IsNullOrWhiteSpace(mention.MentionpersonId?.Value)) {
                            personIds.Add(mention.MentionpersonId!.Value!);
                        }
                    }
                }

                part.ThreadedComments!.Save();
            }

            RemapLegacyThreadedCommentPlaceholders(copiedPart.WorksheetCommentsPart, idMap);
            if (!ReferenceEquals(sourceDocument.WorkbookPartRoot, WorkbookPartRoot)) {
                CopyReferencedWorkbookPersons(sourceDocument.WorkbookPartRoot, personIds);
            }
        }

        private static void RemapLegacyThreadedCommentPlaceholders(WorksheetCommentsPart? commentsPart, IReadOnlyDictionary<string, string> idMap) {
            Comments? comments = commentsPart?.Comments;
            if (comments == null) {
                return;
            }

            const string threadedAuthorPrefix = "tc=";
            bool changed = false;
            foreach (Author author in comments.Authors?.Elements<Author>() ?? Enumerable.Empty<Author>()) {
                string text = author.Text ?? string.Empty;
                if (text.StartsWith(threadedAuthorPrefix, StringComparison.OrdinalIgnoreCase)
                    && idMap.TryGetValue(text.Substring(threadedAuthorPrefix.Length), out string? newId)) {
                    author.Text = threadedAuthorPrefix + newId;
                    changed = true;
                }
            }

            foreach (Comment comment in comments.CommentList?.Elements<Comment>() ?? Enumerable.Empty<Comment>()) {
                foreach (OpenXmlAttribute attribute in comment.GetAttributes()) {
                    if (string.Equals(attribute.LocalName, "uid", StringComparison.Ordinal)
                        && attribute.Value != null
                        && idMap.TryGetValue(attribute.Value, out string? newId)) {
                        comment.SetAttribute(new OpenXmlAttribute(attribute.Prefix, attribute.LocalName, attribute.NamespaceUri, newId));
                        changed = true;
                    }
                }
            }

            if (changed) {
                comments.Save();
            }
        }

        private void CopyReferencedWorkbookPersons(WorkbookPart sourceWorkbookPart, IReadOnlyCollection<string> personIds) {
            if (personIds.Count == 0) {
                return;
            }

            var existing = new HashSet<string>(
                WorkbookPartRoot.WorkbookPersonParts
                    .SelectMany(_ => _.PersonList?.Elements<Threaded.Person>() ?? Enumerable.Empty<Threaded.Person>())
                    .Select(_ => _.Id?.Value)
                    .Where(_ => !string.IsNullOrWhiteSpace(_))
                    .Select(_ => _!),
                StringComparer.OrdinalIgnoreCase);
            var sourcePersons = new Dictionary<string, Threaded.Person>(StringComparer.OrdinalIgnoreCase);
            foreach (Threaded.Person person in sourceWorkbookPart.WorkbookPersonParts
                .SelectMany(_ => _.PersonList?.Elements<Threaded.Person>() ?? Enumerable.Empty<Threaded.Person>())) {
                string? id = person.Id?.Value;
                if (!string.IsNullOrWhiteSpace(id) && !sourcePersons.ContainsKey(id!)) {
                    sourcePersons.Add(id!, person);
                }
            }

            Threaded.PersonList? targetList = null;
            foreach (string personId in personIds) {
                if (existing.Contains(personId) || !sourcePersons.TryGetValue(personId, out Threaded.Person? sourcePerson)) {
                    continue;
                }

                if (targetList == null) {
                    WorkbookPersonPart personPart = WorkbookPartRoot.WorkbookPersonParts.FirstOrDefault()
                        ?? WorkbookPartRoot.AddNewPart<WorkbookPersonPart>();
                    targetList = personPart.PersonList ??= new Threaded.PersonList();
                }

                targetList.Append(sourcePerson.CloneNode(true));
                existing.Add(personId);
            }

            targetList?.Save();
        }

        /// <summary>
        /// Clones a related-part graph into a target container. Each source part is copied once
        /// per operation; later references to it share the copy, mirroring the source graph.
        /// Child relationship ids are kept because every copied part starts with no relationships.
        /// </summary>
        private sealed class RelatedPartCloner {
            private readonly Dictionary<OpenXmlPart, OpenXmlPart> _copied = new Dictionary<OpenXmlPart, OpenXmlPart>();

            internal OpenXmlPart Clone(OpenXmlPart sourcePart, OpenXmlPartContainer owner, string relationshipId) {
                if (_copied.TryGetValue(sourcePart, out OpenXmlPart? existing)) {
                    owner.AddPart(existing, relationshipId);
                    return existing;
                }

                OpenXmlPart targetPart = AddPartLike(owner, sourcePart, relationshipId);
                _copied.Add(sourcePart, targetPart);
                CopyPartData(sourcePart, targetPart);

                foreach (HyperlinkRelationship relationship in sourcePart.HyperlinkRelationships) {
                    targetPart.AddHyperlinkRelationship(relationship.Uri, relationship.IsExternal, relationship.Id);
                }

                foreach (ExternalRelationship relationship in sourcePart.ExternalRelationships) {
                    targetPart.AddExternalRelationship(relationship.RelationshipType, relationship.Uri, relationship.Id);
                }

                foreach (IdPartPair child in sourcePart.Parts.ToList()) {
                    Clone(child.OpenXmlPart, targetPart, child.RelationshipId);
                }

                return targetPart;
            }

            private static void CopyPartData(OpenXmlPart sourcePart, OpenXmlPart targetPart) {
                // A loaded root may hold edits that were never written back to the part stream.
                if (sourcePart.IsRootElementLoaded && sourcePart.RootElement is OpenXmlPartRootElement root) {
                    using Stream targetStream = targetPart.GetStream(FileMode.Create, FileAccess.Write);
                    root.Save(targetStream);
                    return;
                }

                using Stream sourceStream = sourcePart.GetStream(FileMode.Open, FileAccess.Read);
                targetPart.FeedData(sourceStream);
            }

            private static OpenXmlPart AddPartLike(OpenXmlPartContainer owner, OpenXmlPart sourcePart, string relationshipId) {
                string contentType = sourcePart.ContentType;
                switch (sourcePart) {
                    case DrawingsPart:
                        return owner.AddNewPart<DrawingsPart>(contentType, relationshipId);
                    case ChartPart:
                        return owner.AddNewPart<ChartPart>(contentType, relationshipId);
                    case ExtendedChartPart:
                        return owner.AddNewPart<ExtendedChartPart>(contentType, relationshipId);
                    case ChartDrawingPart:
                        return owner.AddNewPart<ChartDrawingPart>(contentType, relationshipId);
                    case ChartStylePart:
                        return owner.AddNewPart<ChartStylePart>(contentType, relationshipId);
                    case ChartColorStylePart:
                        return owner.AddNewPart<ChartColorStylePart>(contentType, relationshipId);
                    case ThemeOverridePart:
                        return owner.AddNewPart<ThemeOverridePart>(contentType, relationshipId);
                    case ImagePart:
                        return owner.AddNewPart<ImagePart>(contentType, relationshipId);
                    case EmbeddedPackagePart:
                        return owner.AddNewPart<EmbeddedPackagePart>(contentType, relationshipId);
                    case EmbeddedObjectPart:
                        return owner.AddNewPart<EmbeddedObjectPart>(contentType, relationshipId);
                    case DiagramColorsPart:
                        return owner.AddNewPart<DiagramColorsPart>(contentType, relationshipId);
                    case DiagramDataPart:
                        return owner.AddNewPart<DiagramDataPart>(contentType, relationshipId);
                    case DiagramLayoutDefinitionPart:
                        return owner.AddNewPart<DiagramLayoutDefinitionPart>(contentType, relationshipId);
                    case DiagramPersistLayoutPart:
                        return owner.AddNewPart<DiagramPersistLayoutPart>(contentType, relationshipId);
                    case DiagramStylePart:
                        return owner.AddNewPart<DiagramStylePart>(contentType, relationshipId);
                    case VmlDrawingPart:
                        return owner.AddNewPart<VmlDrawingPart>(contentType, relationshipId);
                    case WorksheetCommentsPart:
                        return owner.AddNewPart<WorksheetCommentsPart>(contentType, relationshipId);
                    case WorksheetThreadedCommentsPart:
                        return owner.AddNewPart<WorksheetThreadedCommentsPart>(contentType, relationshipId);
                    case WebExtensionPart:
                        return owner.AddNewPart<WebExtensionPart>(contentType, relationshipId);
                    case CustomXmlPart:
                        return owner.AddNewPart<CustomXmlPart>(contentType, relationshipId);
                    default:
                        string extension = Path.GetExtension(sourcePart.Uri.OriginalString);
                        return owner.AddExtendedPart(
                            sourcePart.RelationshipType,
                            contentType,
                            string.IsNullOrEmpty(extension) ? ".bin" : extension,
                            relationshipId);
                }
            }
        }
    }
}
