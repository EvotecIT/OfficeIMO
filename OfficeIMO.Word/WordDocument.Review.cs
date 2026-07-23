using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2013.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        private const string WordprocessingNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        private const int ReviewBodyOrderBase = 0;
        private const int ReviewHeaderOrderBase = 1_000_000;
        private const int ReviewFooterOrderBase = 2_000_000;
        private const int ReviewFootnoteOrderBase = 3_000_000;
        private const int ReviewEndnoteOrderBase = 4_000_000;
        private const int ReviewRelatedPartOrderStride = 100_000;

        /// <summary>
        /// Inspects comments, comment thread metadata, and tracked revisions without mutating the document.
        /// </summary>
        public WordReviewInfo InspectReview() {
            MainDocumentPart mainPart = _wordprocessingDocument.MainDocumentPart
                ?? throw new InvalidOperationException("MainDocumentPart is missing.");

            ReviewRoot[] roots = GetReviewRoots(mainPart);

            Dictionary<string, CommentTargetInfo> commentTargets = CollectCommentTargets(roots);
            IReadOnlyList<WordCommentInfo> comments = InspectComments(mainPart, commentTargets);
            IReadOnlyList<WordRevisionInfo> revisions = InspectRevisions(roots);
            IReadOnlyList<string> unsupportedMetadata = InspectUnsupportedReviewMetadata(mainPart);

            return new WordReviewInfo(comments, revisions, unsupportedMetadata);
        }

        /// <summary>
        /// Creates a structured review report for comments, tracked revisions, unsupported review metadata, and optional accept/reject actions.
        /// </summary>
        /// <param name="actions">Optional accept/reject operation reports to include in the report.</param>
        public WordReviewReport InspectReviewReport(params WordRevisionOperationReport[] actions) {
            return WordReviewReport.From(InspectReview(), actions);
        }

        private static IReadOnlyList<WordCommentInfo> InspectComments(MainDocumentPart mainPart, Dictionary<string, CommentTargetInfo> commentTargets) {
            WordprocessingCommentsPart? commentsPart = mainPart.WordprocessingCommentsPart;
            if (commentsPart?.Comments == null) {
                return Array.Empty<WordCommentInfo>();
            }

            List<Comment> commentElements = commentsPart.Comments.Elements<Comment>().ToList();
            List<CommentEx> commentExElements = mainPart.WordprocessingCommentsExPart?.CommentsEx?.Elements<CommentEx>().ToList()
                ?? new List<CommentEx>();
            Dictionary<string, CommentEx> commentExByParagraphId =
                WordComment.IndexCommentExByParagraphId(commentExElements);
            var comments = new List<WordCommentInfo>();

            for (int i = 0; i < commentElements.Count; i++) {
                Comment comment = commentElements[i];
                CommentEx? commentEx = WordComment.FindCommentExForComment(comment, commentExElements,
                    commentExByParagraphId, i);
                string? id = comment.Id?.Value;
                commentTargets.TryGetValue(id ?? string.Empty, out CommentTargetInfo? target);

                comments.Add(new WordCommentInfo(
                    index: i,
                    id: id,
                    author: comment.Author?.Value,
                    initials: comment.Initials?.Value,
                    dateTime: comment.Date?.Value,
                    text: NormalizeText(comment.InnerText),
                    paraId: commentEx?.ParaId?.Value,
                    parentParaId: commentEx?.ParaIdParent?.Value,
                    isResolved: commentEx?.Done?.Value,
                    targetText: target?.Text ?? string.Empty,
                    targetLocationKind: target?.LocationKind,
                    targetPartUri: target?.PartUri,
                    isInTable: target?.IsInTable ?? false,
                    isInContentControl: target?.IsInContentControl ?? false,
                    isInTextBox: target?.IsInTextBox ?? false,
                    documentOrder: target?.DocumentOrder ?? GetReviewDocumentOrder(null, i)));
            }

            return comments;
        }

        private static Dictionary<string, CommentTargetInfo> CollectCommentTargets(IEnumerable<ReviewRoot> roots) {
            var targets = new Dictionary<string, CommentTargetInfo>(StringComparer.Ordinal);

            foreach (ReviewRoot root in roots) {
                int order = root.OrderBase;
                foreach (OpenXmlElement element in root.Root.Descendants()) {
                    if (element is CommentRangeStart start) {
                        AddCommentRangeTarget(targets, root, start, order);
                    } else if (element is CommentReference reference) {
                        AddCommentReferenceTarget(targets, root, reference, order);
                    }

                    order++;
                }
            }

            return targets;
        }

        private static void AddCommentRangeTarget(Dictionary<string, CommentTargetInfo> targets, ReviewRoot root, CommentRangeStart start, int documentOrder) {
            string? id = start.Id?.Value;
            string commentId = id ?? string.Empty;
            if (commentId.Length == 0 || targets.ContainsKey(commentId)) {
                return;
            }

            targets[commentId] = new CommentTargetInfo(
                ReadCommentRangeText(start, commentId),
                root.LocationKind,
                root.PartUri,
                IsInTable(start),
                IsInContentControl(start),
                IsInTextBox(start),
                documentOrder);
        }

        private static void AddCommentReferenceTarget(Dictionary<string, CommentTargetInfo> targets, ReviewRoot root, CommentReference reference, int documentOrder) {
            string? id = reference.Id?.Value;
            string commentId = id ?? string.Empty;
            if (commentId.Length == 0 || targets.ContainsKey(commentId)) {
                return;
            }

            Paragraph? paragraph = reference.Ancestors<Paragraph>().FirstOrDefault();
            targets[commentId] = new CommentTargetInfo(
                NormalizeText(paragraph?.InnerText ?? string.Empty),
                root.LocationKind,
                root.PartUri,
                IsInTable(reference),
                IsInContentControl(reference),
                IsInTextBox(reference),
                documentOrder);
        }

        private static string ReadCommentRangeText(CommentRangeStart start, string id) {
            var parts = new List<string>();
            OpenXmlElement? current = GetNextElementInDocumentOrder(start);
            Paragraph? previousParagraph = null;
            TableCell? previousCell = null;

            while (current != null) {
                if (current is CommentRangeEnd end && string.Equals(end.Id?.Value, id, StringComparison.Ordinal)) {
                    break;
                }

                if (current is Text text) {
                    AddRangeTextSeparatorIfNeeded(parts, current, ref previousParagraph, ref previousCell);
                    parts.Add(text.Text);
                } else if (current is DeletedText deletedText) {
                    AddRangeTextSeparatorIfNeeded(parts, current, ref previousParagraph, ref previousCell);
                    parts.Add(deletedText.Text);
                }

                current = GetNextElementInDocumentOrder(current);
            }

            return NormalizeText(string.Concat(parts));
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

        private static OpenXmlElement? GetNextElementInDocumentOrder(OpenXmlElement element) {
            if (element.FirstChild != null) {
                return element.FirstChild;
            }

            OpenXmlElement? current = element;
            while (current != null) {
                OpenXmlElement? next = current.NextSibling();
                if (next != null) {
                    return next;
                }

                current = current.Parent;
            }

            return null;
        }

        private static IReadOnlyList<WordRevisionInfo> InspectRevisions(IEnumerable<ReviewRoot> roots) {
            var revisions = new List<WordRevisionInfo>();

            foreach (ReviewRoot root in roots) {
                int order = root.OrderBase;
                foreach (OpenXmlElement element in root.Root.Descendants()) {
                    if (!TryGetRevisionType(element, out WordReviewRevisionType revisionType)) {
                        order++;
                        continue;
                    }

                    revisions.Add(new WordRevisionInfo(
                        index: revisions.Count,
                        revisionType: revisionType,
                        elementName: element.LocalName,
                        id: GetWordprocessingAttribute(element, "id"),
                        author: GetWordprocessingAttribute(element, "author"),
                        dateTime: ParseDate(GetWordprocessingAttribute(element, "date")),
                        affectedText: NormalizeText(element.InnerText),
                        locationText: GetLocationText(element),
                        locationKind: root.LocationKind,
                        partUri: root.PartUri,
                        isInTable: IsInTable(element),
                        isInContentControl: IsInContentControl(element),
                        isInTextBox: IsInTextBox(element),
                        documentOrder: order));
                    order++;
                }
            }

            return revisions;
        }

        private static IReadOnlyList<string> InspectUnsupportedReviewMetadata(MainDocumentPart mainPart) {
            var details = new List<string>();

            if (mainPart.WordprocessingCommentsIdsPart != null) {
                details.Add($"{mainPart.WordprocessingCommentsIdsPart.Uri}: durable comment ids are preserved but not yet mapped to comment readback.");
            }

            if (mainPart.WordprocessingPeoplePart != null) {
                details.Add($"{mainPart.WordprocessingPeoplePart.Uri}: people metadata is preserved but not yet mapped to authors.");
            }

            foreach (OpenXmlPart part in mainPart.Parts
                .Select(pair => pair.OpenXmlPart)
                .Where(part => IsModernReviewMetadataPart(part))
                .OrderBy(part => part.Uri.ToString(), StringComparer.OrdinalIgnoreCase)) {
                string detail = IsCommentsExtensibleMetadataPart(part)
                    ? $"{part.Uri}: extensible comment metadata is preserved but not yet parsed. ({part.ContentType})"
                    : $"{part.Uri} ({part.ContentType})";
                if (!details.Contains(detail, StringComparer.OrdinalIgnoreCase)) {
                    details.Add(detail);
                }
            }

            return details;
        }

        private static bool IsModernReviewMetadataPart(OpenXmlPart part) {
            string uri = part.Uri.OriginalString;
            string contentType = part.ContentType;
            return IsCommentsExtensibleMetadataPart(part)
                || ContainsIgnoreCase(uri, "commentsIds")
                || ContainsIgnoreCase(contentType, "commentsIds")
                || ContainsIgnoreCase(uri, "people")
                || ContainsIgnoreCase(contentType, "people");
        }

        private static bool IsCommentsExtensibleMetadataPart(OpenXmlPart part) {
            string uri = part.Uri.OriginalString;
            string contentType = part.ContentType;
            return ContainsIgnoreCase(uri, "commentsExtensible")
                || ContainsIgnoreCase(contentType, "commentsExtensible");
        }

        private static bool TryGetRevisionType(OpenXmlElement element, out WordReviewRevisionType revisionType) {
            switch (element.LocalName) {
                case "ins":
                    revisionType = WordReviewRevisionType.Insertion;
                    return true;
                case "del":
                    revisionType = WordReviewRevisionType.Deletion;
                    return true;
                case "moveFrom":
                    revisionType = WordReviewRevisionType.MoveFrom;
                    return true;
                case "moveTo":
                    revisionType = WordReviewRevisionType.MoveTo;
                    return true;
                case "pPrChange":
                    revisionType = WordReviewRevisionType.ParagraphFormatting;
                    return true;
                case "rPrChange":
                    revisionType = WordReviewRevisionType.RunFormatting;
                    return true;
                case "tblPrChange":
                case "tblGridChange":
                    revisionType = WordReviewRevisionType.TableFormatting;
                    return true;
                case "tblPrExChange":
                case "trPrChange":
                    revisionType = WordReviewRevisionType.TableRowFormatting;
                    return true;
                case "tcPrChange":
                    revisionType = WordReviewRevisionType.TableCellFormatting;
                    return true;
                case "sectPrChange":
                    revisionType = WordReviewRevisionType.SectionFormatting;
                    return true;
                default:
                    revisionType = WordReviewRevisionType.Unknown;
                    return false;
            }
        }

        private static string? GetWordprocessingAttribute(OpenXmlElement element, string localName) {
            foreach (OpenXmlAttribute attribute in element.GetAttributes()) {
                if (string.Equals(attribute.LocalName, localName, StringComparison.Ordinal)
                    && string.Equals(attribute.NamespaceUri, WordprocessingNamespace, StringComparison.Ordinal)) {
                    return string.IsNullOrWhiteSpace(attribute.Value) ? null : attribute.Value;
                }
            }

            return null;
        }

        private static DateTime? ParseDate(string? value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return null;
            }

            return DateTime.TryParse(value, out DateTime parsed) ? parsed : null;
        }

        private static string GetLocationText(OpenXmlElement element) {
            Paragraph? paragraph = element.Ancestors<Paragraph>().FirstOrDefault();
            if (paragraph != null) {
                return NormalizeText(paragraph.InnerText);
            }

            TableCell? tableCell = element.Ancestors<TableCell>().FirstOrDefault();
            if (tableCell != null) {
                return NormalizeText(tableCell.InnerText);
            }

            Table? table = element.Ancestors<Table>().FirstOrDefault();
            if (table != null) {
                return NormalizeText(table.InnerText);
            }

            string parentText = NormalizeText(element.Parent?.InnerText ?? element.InnerText);
            if (!string.IsNullOrWhiteSpace(parentText)) {
                return parentText;
            }

            return GetNearestPreviousText(element.Parent ?? element);
        }

        private static string GetNearestPreviousText(OpenXmlElement element) {
            OpenXmlElement? current = element;
            while (current != null) {
                OpenXmlElement? previous = current.PreviousSibling();
                while (previous != null) {
                    string text = NormalizeText(previous.InnerText);
                    if (!string.IsNullOrWhiteSpace(text)) {
                        return text;
                    }

                    previous = previous.PreviousSibling();
                }

                current = current.Parent;
            }

            return string.Empty;
        }

        private static WordReviewLocationKind MapLocationKind(WordFieldLocationKind locationKind) {
            switch (locationKind) {
                case WordFieldLocationKind.Header:
                    return WordReviewLocationKind.Header;
                case WordFieldLocationKind.Footer:
                    return WordReviewLocationKind.Footer;
                case WordFieldLocationKind.Footnote:
                    return WordReviewLocationKind.Footnote;
                case WordFieldLocationKind.Endnote:
                    return WordReviewLocationKind.Endnote;
                default:
                    return WordReviewLocationKind.Body;
            }
        }

        private static ReviewRoot[] GetReviewRoots(MainDocumentPart mainPart) {
            var roots = new List<ReviewRoot>();
            int headerIndex = 0;
            int footerIndex = 0;

            foreach (WordFieldInventory.FieldRoot root in WordFieldInventory.EnumerateFieldRoots(mainPart)) {
                WordReviewLocationKind locationKind = MapLocationKind(root.LocationKind);
                int orderBase;
                switch (locationKind) {
                    case WordReviewLocationKind.Header:
                        orderBase = ReviewHeaderOrderBase + (headerIndex++ * ReviewRelatedPartOrderStride);
                        break;
                    case WordReviewLocationKind.Footer:
                        orderBase = ReviewFooterOrderBase + (footerIndex++ * ReviewRelatedPartOrderStride);
                        break;
                    case WordReviewLocationKind.Footnote:
                        orderBase = ReviewFootnoteOrderBase;
                        break;
                    case WordReviewLocationKind.Endnote:
                        orderBase = ReviewEndnoteOrderBase;
                        break;
                    default:
                        orderBase = ReviewBodyOrderBase;
                        break;
                }

                roots.Add(new ReviewRoot(root.Root, locationKind, root.PartUri, orderBase));
            }

            return roots.ToArray();
        }

        private static int GetReviewDocumentOrder(WordReviewLocationKind? locationKind, int index) {
            int baseOrder;
            switch (locationKind) {
                case WordReviewLocationKind.Header:
                    baseOrder = ReviewHeaderOrderBase;
                    break;
                case WordReviewLocationKind.Footer:
                    baseOrder = ReviewFooterOrderBase;
                    break;
                case WordReviewLocationKind.Footnote:
                    baseOrder = ReviewFootnoteOrderBase;
                    break;
                case WordReviewLocationKind.Endnote:
                    baseOrder = ReviewEndnoteOrderBase;
                    break;
                default:
                    baseOrder = ReviewBodyOrderBase;
                    break;
            }

            return baseOrder + index;
        }

        private static bool IsInTable(OpenXmlElement element) => element.Ancestors<Table>().Any();

        private static bool IsInContentControl(OpenXmlElement element) => element.Ancestors<SdtElement>().Any();

        private static bool IsInTextBox(OpenXmlElement element) =>
            element.Ancestors().Any(ancestor =>
                string.Equals(ancestor.LocalName, "txbxContent", StringComparison.Ordinal) ||
                string.Equals(ancestor.LocalName, "textbox", StringComparison.Ordinal));

        private static bool ContainsIgnoreCase(string value, string search) =>
            value.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0;

        private static string NormalizeText(string value) =>
            string.IsNullOrEmpty(value) ? string.Empty : string.Join(" ", value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));

        private sealed class ReviewRoot {
            internal ReviewRoot(OpenXmlCompositeElement root, WordReviewLocationKind locationKind, string partUri, int orderBase) {
                Root = root;
                LocationKind = locationKind;
                PartUri = partUri;
                OrderBase = orderBase;
            }

            internal OpenXmlCompositeElement Root { get; }

            internal WordReviewLocationKind LocationKind { get; }

            internal string PartUri { get; }

            internal int OrderBase { get; }
        }

        private sealed class CommentTargetInfo {
            internal CommentTargetInfo(
                string text,
                WordReviewLocationKind locationKind,
                string partUri,
                bool isInTable,
                bool isInContentControl,
                bool isInTextBox,
                int documentOrder) {
                Text = text;
                LocationKind = locationKind;
                PartUri = partUri;
                IsInTable = isInTable;
                IsInContentControl = isInContentControl;
                IsInTextBox = isInTextBox;
                DocumentOrder = documentOrder;
            }

            internal string Text { get; }

            internal WordReviewLocationKind LocationKind { get; }

            internal string PartUri { get; }

            internal bool IsInTable { get; }

            internal bool IsInContentControl { get; }

            internal bool IsInTextBox { get; }

            internal int DocumentOrder { get; }
        }
    }
}
