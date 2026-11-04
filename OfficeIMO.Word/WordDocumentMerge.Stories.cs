using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2013.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    internal sealed partial class WordDocumentMerge {
        /// <summary>
        /// Copies the footnotes and endnotes referenced by the appended content and renumbers the references.
        /// </summary>
        private List<OpenXmlElement> ImportNotes(OpenXmlElement root) {
            var imported = new List<OpenXmlElement>();
            ImportNoteStory<FootnoteReference>(
                root,
                _sourceMain.FootnotesPart,
                _sourceMain.FootnotesPart?.Footnotes,
                () => {
                    FootnotesPart part = _destinationMain.FootnotesPart ?? _destinationMain.AddNewPart<FootnotesPart>();
                    if (part.Footnotes == null) {
                        WordDocument.GenerateFootNotesPart1Content(part);
                    }

                    return (part, part.Footnotes!);
                },
                imported);
            ImportNoteStory<EndnoteReference>(
                root,
                _sourceMain.EndnotesPart,
                _sourceMain.EndnotesPart?.Endnotes,
                () => {
                    EndnotesPart part = _destinationMain.EndnotesPart ?? _destinationMain.AddNewPart<EndnotesPart>();
                    if (part.Endnotes == null) {
                        WordDocument.GenerateEndNotesPart1Content(part);
                    }

                    return (part, part.Endnotes!);
                },
                imported);
            return imported;
        }

        private void ImportNoteStory<TReference>(
            OpenXmlElement root,
            OpenXmlPart? sourcePart,
            OpenXmlElement? sourceNotes,
            Func<(OpenXmlPart Part, OpenXmlElement Notes)> ensureDestination,
            List<OpenXmlElement> imported)
            where TReference : FootnoteEndnoteReferenceType {
            List<TReference> references = root.Descendants<TReference>().ToList();
            if (references.Count == 0 || sourcePart == null || sourceNotes == null) {
                return;
            }

            (OpenXmlPart destinationPart, OpenXmlElement destinationNotes) = ensureDestination();
            long next = Math.Max(1L, destinationNotes.Elements<FootnoteEndnoteType>()
                .Select(_ => _.Id?.Value ?? 0L)
                .DefaultIfEmpty(0L)
                .Max() + 1L);
            var map = new Dictionary<long, long>();

            foreach (TReference reference in references) {
                if (reference.Id?.Value is not long oldId) {
                    continue;
                }

                if (!map.TryGetValue(oldId, out long newId)) {
                    FootnoteEndnoteType? note = sourceNotes.Elements<FootnoteEndnoteType>().FirstOrDefault(_ => _.Id?.Value == oldId);
                    if (note == null) {
                        continue;
                    }

                    var clone = (FootnoteEndnoteType)note.CloneNode(true);
                    newId = next++;
                    clone.Id = newId;
                    map[oldId] = newId;
                    destinationNotes.Append(clone);
                    RemapRelationshipIds(clone, sourcePart, destinationPart);
                    imported.Add(clone);
                }

                reference.Id = newId;
            }
        }

        /// <summary>
        /// Copies the comments anchored in the appended content and renumbers their ranges and references.
        /// </summary>
        private List<OpenXmlElement> ImportComments(IEnumerable<OpenXmlElement> roots) {
            var imported = new List<OpenXmlElement>();
            WordprocessingCommentsPart? sourcePart = _sourceMain.WordprocessingCommentsPart;
            Comments? sourceComments = sourcePart?.Comments;
            if (sourcePart == null || sourceComments == null) {
                return imported;
            }

            List<OpenXmlElement> anchors = roots
                .SelectMany(_ => _.Descendants())
                .Where(_ => _ is CommentRangeStart or CommentRangeEnd or CommentReference)
                .ToList();
            if (anchors.Count == 0) {
                return imported;
            }

            Comments destinationComments = WordComment.GetCommentsPart(_destination);
            WordprocessingCommentsPart destinationPart = _destinationMain.WordprocessingCommentsPart!;
            int next = destinationComments.Elements<Comment>()
                .Select(_ => int.TryParse(_.Id?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed) ? parsed : 0)
                .DefaultIfEmpty(0)
                .Max() + 1;
            var map = new Dictionary<string, string>(StringComparer.Ordinal);

            foreach (OpenXmlElement anchor in anchors) {
                string? oldId = anchor switch {
                    CommentRangeStart start => start.Id?.Value,
                    CommentRangeEnd end => end.Id?.Value,
                    CommentReference reference => reference.Id?.Value,
                    _ => null
                };
                if (oldId == null) {
                    continue;
                }

                if (!map.TryGetValue(oldId, out string? newId)) {
                    Comment? comment = sourceComments.Elements<Comment>().FirstOrDefault(_ => _.Id?.Value == oldId);
                    if (comment == null) {
                        continue;
                    }

                    var clone = (Comment)comment.CloneNode(true);
                    newId = (next++).ToString(CultureInfo.InvariantCulture);
                    clone.Id = newId;
                    map[oldId] = newId;
                    destinationComments.Append(clone);
                    RemapRelationshipIds(clone, sourcePart, destinationPart);
                    imported.Add(clone);
                }

                switch (anchor) {
                    case CommentRangeStart start:
                        start.Id = newId;
                        break;
                    case CommentRangeEnd end:
                        end.Id = newId;
                        break;
                    case CommentReference reference:
                        reference.Id = newId;
                        break;
                }
            }

            // Replies may have no anchor of their own; they are threaded to their parent through commentsExtended.
            CommentsEx? extensions = _sourceMain.WordprocessingCommentsExPart?.CommentsEx;
            if (extensions != null) {
                var importedParagraphIds = new HashSet<string>(
                    imported.SelectMany(_ => _.Elements<Paragraph>()).Select(_ => _.ParagraphId?.Value).OfType<string>(),
                    StringComparer.OrdinalIgnoreCase);
                bool added;
                do {
                    added = false;
                    foreach (Comment comment in sourceComments.Elements<Comment>()) {
                        string? oldId = comment.Id?.Value;
                        if (oldId == null || map.ContainsKey(oldId)) {
                            continue;
                        }

                        var paragraphIds = comment.Elements<Paragraph>().Select(_ => _.ParagraphId?.Value).OfType<string>().ToList();
                        bool isReply = extensions.Elements<CommentEx>().Any(_ =>
                            _.ParaId?.Value is string paraId &&
                            paragraphIds.Contains(paraId, StringComparer.OrdinalIgnoreCase) &&
                            _.ParaIdParent?.Value is string parent &&
                            importedParagraphIds.Contains(parent));
                        if (!isReply) {
                            continue;
                        }

                        var clone = (Comment)comment.CloneNode(true);
                        string newId = (next++).ToString(CultureInfo.InvariantCulture);
                        clone.Id = newId;
                        map[oldId] = newId;
                        destinationComments.Append(clone);
                        RemapRelationshipIds(clone, sourcePart, destinationPart);
                        imported.Add(clone);
                        importedParagraphIds.UnionWith(paragraphIds);
                        added = true;
                    }
                } while (added);
            }

            return imported;
        }

        /// <summary>
        /// Copies reply threading and resolved state (commentsExtended) for imported comments.
        /// </summary>
        private void ImportCommentExtensions(List<OpenXmlElement> comments, Dictionary<string, string> paragraphIds) {
            CommentsEx? sourceExtensions = _sourceMain.WordprocessingCommentsExPart?.CommentsEx;
            if (sourceExtensions == null || comments.Count == 0) {
                return;
            }

            var sourceByParagraphId = new Dictionary<string, CommentEx>(StringComparer.OrdinalIgnoreCase);
            foreach (CommentEx extension in sourceExtensions.Elements<CommentEx>()) {
                string? paragraphId = extension.ParaId?.Value;
                if (paragraphId != null && !sourceByParagraphId.ContainsKey(paragraphId)) {
                    sourceByParagraphId.Add(paragraphId, extension);
                }
            }

            var originalIds = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (KeyValuePair<string, string> pair in paragraphIds) {
                originalIds[pair.Value] = pair.Key;
            }

            CommentsEx? destinationExtensions = null;
            foreach (Comment comment in comments.OfType<Comment>()) {
                foreach (Paragraph paragraph in comment.Elements<Paragraph>()) {
                    string? currentId = paragraph.ParagraphId?.Value;
                    if (currentId == null) {
                        continue;
                    }

                    string originalId = originalIds.TryGetValue(currentId, out string? renamed) ? renamed : currentId;
                    if (!sourceByParagraphId.TryGetValue(originalId, out CommentEx? extension)) {
                        continue;
                    }

                    var clone = (CommentEx)extension.CloneNode(true);
                    clone.ParaId = currentId;
                    if (clone.ParaIdParent?.Value is string parent && paragraphIds.TryGetValue(parent, out string? mappedParent)) {
                        clone.ParaIdParent = mappedParent;
                    }

                    destinationExtensions ??= WordComment.GetCommentsExPart(_destination);
                    destinationExtensions.Append(clone);
                }
            }
        }
    }
}
