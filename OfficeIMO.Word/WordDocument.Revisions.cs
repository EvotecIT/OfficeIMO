using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Handles revisions within the document.
    /// </summary>
    public partial class WordDocument {

        /// <summary>
        /// Given author name, accept all revisions by given Author.
        /// </summary>
        /// <param name="authorName">Author whose revisions should be accepted.</param>
        public void AcceptRevisions(string authorName) {
            AcceptRevisions(new WordRevisionFilter { Author = authorName });
        }

        /// <summary>
        /// Accept all revisions in the document.
        /// </summary>
        public void AcceptRevisions() {
            AcceptRevisions(WordRevisionFilter.All());
        }

        /// <summary>
        /// Accepts revisions that match the provided filter and returns a deterministic operation report.
        /// </summary>
        /// <param name="filter">Filter describing revisions to accept.</param>
        public WordRevisionOperationReport AcceptRevisions(WordRevisionFilter filter) {
            return ApplyRevisionOperation(WordRevisionOperationKind.Accept, filter, scope: null);
        }

        /// <summary>
        /// Accepts revisions inside the specified paragraph.
        /// </summary>
        /// <param name="paragraph">Paragraph that scopes the operation.</param>
        /// <param name="filter">Optional additional filter.</param>
        public WordRevisionOperationReport AcceptRevisionsInParagraph(WordParagraph paragraph, WordRevisionFilter? filter = null) {
            if (paragraph == null) throw new ArgumentNullException(nameof(paragraph));
            return ApplyRevisionOperation(WordRevisionOperationKind.Accept, filter ?? WordRevisionFilter.All(), paragraph._paragraph);
        }

        /// <summary>
        /// Converts tracked revisions into visible markup by replacing revision
        /// elements with formatted runs. Inserted text is underlined and colored
        /// blue, while deleted text is displayed with red strikethrough.
        /// </summary>
        public void ConvertRevisionsToMarkup() {
            var body = this._document.Body ?? throw new InvalidOperationException("Document body is missing.");

            // Process insertions
            foreach (var insertion in body.Descendants<InsertedRun>().ToList()) {
                var parent = insertion.Parent ?? throw new InvalidOperationException("Insertion has no parent.");
                OpenXmlElement last = insertion;
                foreach (var run in insertion.Elements<Run>().Select(r => (Run)r.CloneNode(true))) {
                    var rPr = run.RunProperties ?? new RunProperties();
                    rPr.Color = new Color() { Val = "0000FF" };
                    rPr.Underline = new Underline() { Val = UnderlineValues.Single };
                    run.RunProperties = rPr;
                    parent.InsertAfter(run, last);
                    last = run;
                }
                insertion.Remove();
            }

            // Process deletions
            foreach (var deletion in body.Descendants<DeletedRun>().ToList()) {
                var parent = deletion.Parent ?? throw new InvalidOperationException("Deletion has no parent.");
                OpenXmlElement last = deletion;
                foreach (var run in deletion.Elements<Run>().Select(r => (Run)r.CloneNode(true))) {
                    var rPr = run.RunProperties ?? new RunProperties();
                    rPr.Color = new Color() { Val = "FF0000" };
                    rPr.Strike = new Strike();
                    run.RunProperties = rPr;
                    parent.InsertAfter(run, last);
                    last = run;
                }
                deletion.Remove();
            }
        }

        /// <summary>
        /// Reject all revisions by given author.
        /// </summary>
        /// <param name="authorName">Author whose revisions should be rejected.</param>
        public void RejectRevisions(string authorName) {
            RejectRevisions(new WordRevisionFilter { Author = authorName });
        }

        /// <summary>
        /// Reject all revisions in the document.
        /// </summary>
        public void RejectRevisions() {
            RejectRevisions(WordRevisionFilter.All());
        }

        /// <summary>
        /// Rejects revisions that match the provided filter and returns a deterministic operation report.
        /// </summary>
        /// <param name="filter">Filter describing revisions to reject.</param>
        public WordRevisionOperationReport RejectRevisions(WordRevisionFilter filter) {
            return ApplyRevisionOperation(WordRevisionOperationKind.Reject, filter, scope: null);
        }

        /// <summary>
        /// Rejects revisions inside the specified paragraph.
        /// </summary>
        /// <param name="paragraph">Paragraph that scopes the operation.</param>
        /// <param name="filter">Optional additional filter.</param>
        public WordRevisionOperationReport RejectRevisionsInParagraph(WordParagraph paragraph, WordRevisionFilter? filter = null) {
            if (paragraph == null) throw new ArgumentNullException(nameof(paragraph));
            return ApplyRevisionOperation(WordRevisionOperationKind.Reject, filter ?? WordRevisionFilter.All(), paragraph._paragraph);
        }

        private WordRevisionOperationReport ApplyRevisionOperation(WordRevisionOperationKind operation, WordRevisionFilter filter, OpenXmlElement? scope) {
            if (filter == null) throw new ArgumentNullException(nameof(filter));

            List<RevisionCandidate> matches = CollectRevisionCandidates()
                .Where(candidate => MatchesFilter(candidate, filter))
                .Where(candidate => scope == null || IsWithinScope(candidate.Element, scope))
                .ToList();

            WordRevisionInfo[] matchedInfo = matches
                .Select((candidate, index) => candidate.ToInfo(index))
                .ToArray();

            foreach (RevisionCandidate match in matches) {
                ApplyRevisionOperation(operation, match);
            }

            return new WordRevisionOperationReport(operation, matchedInfo);
        }

        private List<RevisionCandidate> CollectRevisionCandidates() {
            MainDocumentPart mainPart = _wordprocessingDocument.MainDocumentPart
                ?? throw new InvalidOperationException("MainDocumentPart is missing.");
            var candidates = new List<RevisionCandidate>();

            foreach (WordFieldInventory.FieldRoot root in WordFieldInventory.EnumerateFieldRoots(mainPart)) {
                WordReviewLocationKind locationKind = MapLocationKind(root.LocationKind);
                foreach (OpenXmlElement element in root.Root.Descendants()) {
                    if (!TryGetRevisionType(element, out WordReviewRevisionType revisionType)) {
                        continue;
                    }

                    candidates.Add(new RevisionCandidate(element, revisionType, locationKind, root.PartUri));
                }
            }

            return candidates;
        }

        private static bool MatchesFilter(RevisionCandidate candidate, WordRevisionFilter filter) {
            if (!string.IsNullOrWhiteSpace(filter.Author)
                && !string.Equals(candidate.Author, filter.Author, StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            if (!string.IsNullOrWhiteSpace(filter.RevisionId)
                && !string.Equals(candidate.Id, filter.RevisionId, StringComparison.Ordinal)) {
                return false;
            }

            if (filter.RevisionType != null && candidate.RevisionType != filter.RevisionType.Value) {
                return false;
            }

            if (filter.DateFrom != null && (candidate.DateTime == null || candidate.DateTime.Value < filter.DateFrom.Value)) {
                return false;
            }

            if (filter.DateTo != null && (candidate.DateTime == null || candidate.DateTime.Value > filter.DateTo.Value)) {
                return false;
            }

            if (filter.LocationKind != null && candidate.LocationKind != filter.LocationKind.Value) {
                return false;
            }

            if (!string.IsNullOrWhiteSpace(filter.PartUri)
                && !string.Equals(candidate.PartUri, filter.PartUri, StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            if (filter.IsInTable != null && candidate.IsInTable != filter.IsInTable.Value) {
                return false;
            }

            if (filter.IsInContentControl != null && candidate.IsInContentControl != filter.IsInContentControl.Value) {
                return false;
            }

            if (filter.IsInTextBox != null && candidate.IsInTextBox != filter.IsInTextBox.Value) {
                return false;
            }

            return true;
        }

        private static bool IsWithinScope(OpenXmlElement element, OpenXmlElement scope) {
            if (ReferenceEquals(element, scope)) {
                return true;
            }

            return element.Ancestors().Any(ancestor => ReferenceEquals(ancestor, scope));
        }

        private static void ApplyRevisionOperation(WordRevisionOperationKind operation, RevisionCandidate candidate) {
            bool accept = operation == WordRevisionOperationKind.Accept;
            switch (candidate.RevisionType) {
                case WordReviewRevisionType.Insertion:
                case WordReviewRevisionType.MoveTo:
                    if (accept) {
                        PromoteRevisionRuns(candidate.Element, "rsidR", restoreDeletedText: false);
                    } else {
                        candidate.Element.Remove();
                    }
                    break;
                case WordReviewRevisionType.Deletion:
                case WordReviewRevisionType.MoveFrom:
                    if (accept) {
                        candidate.Element.Remove();
                    } else {
                        PromoteRevisionRuns(candidate.Element, "rsidDel", restoreDeletedText: true);
                    }
                    break;
                case WordReviewRevisionType.RunFormatting:
                case WordReviewRevisionType.ParagraphFormatting:
                case WordReviewRevisionType.TableFormatting:
                case WordReviewRevisionType.TableRowFormatting:
                case WordReviewRevisionType.TableCellFormatting:
                case WordReviewRevisionType.SectionFormatting:
                    if (accept) {
                        candidate.Element.Remove();
                    } else {
                        RestorePreviousProperties(candidate.Element);
                    }
                    break;
                default:
                    candidate.Element.Remove();
                    break;
            }
        }

        private static void PromoteRevisionRuns(OpenXmlElement revisionElement, string revisionAttributeName, bool restoreDeletedText) {
            OpenXmlElement next = revisionElement;
            foreach (Run run in revisionElement.Elements<Run>().ToList()) {
                Run promoted = new Run(run.OuterXml);
                if (restoreDeletedText) {
                    RestoreDeletedText(promoted);
                }

                next.InsertAfterSelf(promoted);
                next = next.NextSibling() ?? throw new InvalidOperationException("Revision has no next sibling.");
            }

            revisionElement.RemoveAttribute(revisionAttributeName, WordprocessingNamespace);
            revisionElement.RemoveAttribute("rsidRPr", WordprocessingNamespace);
            revisionElement.Remove();
        }

        private static void RestoreDeletedText(Run run) {
            foreach (DeletedText deletedText in run.Descendants<DeletedText>().ToList()) {
                var restored = new Text(deletedText.Text);
                if (deletedText.Space != null) {
                    restored.Space = deletedText.Space.Value;
                }

                deletedText.InsertAfterSelf(restored);
                deletedText.Remove();
            }
        }

        private static void RestorePreviousProperties(OpenXmlElement revisionElement) {
            if (revisionElement.Parent is not OpenXmlCompositeElement propertyContainer) {
                revisionElement.Remove();
                return;
            }

            OpenXmlElement? previousProperties = revisionElement.ChildElements.FirstOrDefault();
            propertyContainer.RemoveAllChildren();

            if (previousProperties == null) {
                return;
            }

            foreach (OpenXmlElement child in previousProperties.ChildElements) {
                propertyContainer.Append(child.CloneNode(true));
            }
        }

        private sealed class RevisionCandidate {
            internal RevisionCandidate(OpenXmlElement element, WordReviewRevisionType revisionType, WordReviewLocationKind locationKind, string partUri) {
                Element = element;
                RevisionType = revisionType;
                LocationKind = locationKind;
                PartUri = partUri;
            }

            internal OpenXmlElement Element { get; }

            internal WordReviewRevisionType RevisionType { get; }

            internal WordReviewLocationKind LocationKind { get; }

            internal string PartUri { get; }

            internal string ElementName => Element.LocalName;

            internal string? Id => GetWordprocessingAttribute(Element, "id");

            internal string? Author => GetWordprocessingAttribute(Element, "author");

            internal DateTime? DateTime => ParseDate(GetWordprocessingAttribute(Element, "date"));

            internal bool IsInTable => WordDocument.IsInTable(Element);

            internal bool IsInContentControl => WordDocument.IsInContentControl(Element);

            internal bool IsInTextBox => WordDocument.IsInTextBox(Element);

            internal WordRevisionInfo ToInfo(int index) {
                return new WordRevisionInfo(
                    index: index,
                    revisionType: RevisionType,
                    elementName: ElementName,
                    id: Id,
                    author: Author,
                    dateTime: DateTime,
                    affectedText: NormalizeText(Element.InnerText),
                    locationText: GetLocationText(Element),
                    locationKind: LocationKind,
                    partUri: PartUri,
                    isInTable: IsInTable,
                    isInContentControl: IsInContentControl,
                    isInTextBox: IsInTextBox,
                    documentOrder: GetReviewDocumentOrder(LocationKind, index));
            }
        }
    }
}
