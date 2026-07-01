using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public static partial class WordDocumentComparer {
        private static void ApplyContentControlFindings(WordprocessingDocument sourceDocument, WordprocessingDocument targetDocument, WordComparisonResult result, WordComparisonRedlineOptions options) {
            List<RedlineContentControlEntry> sourceControls = GetRedlineContentControlEntries(sourceDocument);
            List<RedlineContentControlEntry> targetControls = GetRedlineContentControlEntries(targetDocument);
            HashSet<int> redlineTargetIndexes = GetContentControlRedlineTargetIndexes(result, options, targetControls);
            var rewrittenControls = new HashSet<int>();

            foreach (WordComparisonFinding finding in result.Findings) {
                if (TryApplyDeletedContentControlFinding(finding, options, sourceControls, targetDocument)) {
                    continue;
                }

                if (!TryGetContentControlRedlineTargetIndex(finding, options, targetControls, out int targetIndex)) {
                    continue;
                }

                if (rewrittenControls.Contains(targetIndex)) {
                    continue;
                }

                RedlineContentControlEntry entry = targetControls[targetIndex];
                if (HasDescendantContentControlRedlineFinding(entry.ContentControl, targetControls, redlineTargetIndexes)) {
                    continue;
                }

                string? sourceText = ExtractContentControlFindingText(finding.SourceText);
                string? targetText = ExtractContentControlFindingText(finding.TargetText);
                if (string.Equals(sourceText, targetText, StringComparison.Ordinal)) {
                    continue;
                }

                switch (finding.ChangeKind) {
                    case WordComparisonChangeKind.Modified:
                        if (RewriteContentControlWithTrackedText(entry.ContentControl, sourceText, targetText, options)) {
                            rewrittenControls.Add(targetIndex);
                        }

                        break;
                    case WordComparisonChangeKind.Inserted:
                        if (RewriteContentControlWithTrackedText(entry.ContentControl, null, targetText, options)) {
                            rewrittenControls.Add(targetIndex);
                        }

                        break;
                }
            }
        }

        private static bool TryApplyDeletedContentControlFinding(
            WordComparisonFinding finding,
            WordComparisonRedlineOptions options,
            IReadOnlyList<RedlineContentControlEntry> sourceControls,
            WordprocessingDocument targetDocument) {
            if (!ShouldTrackFinding(finding, options) ||
                finding.Scope != WordComparisonScope.ContentControl ||
                finding.ChangeKind != WordComparisonChangeKind.Deleted ||
                !HasTrackedText(finding)) {
                return false;
            }

            int sourceIndex = finding.SourceIndex ?? -1;
            if (sourceIndex < 0 && !TryParseIndexedLocation(finding.Location, "content-control", out sourceIndex)) {
                return false;
            }

            if (sourceIndex < 0 || sourceIndex >= sourceControls.Count) {
                return false;
            }

            SdtElement deletedControl = (SdtElement)sourceControls[sourceIndex].ContentControl.CloneNode(true);
            if (!RewriteContentControlWithTrackedText(deletedControl, ExtractContentControlFindingText(finding.SourceText), null, options)) {
                return false;
            }

            InsertDeletedContentControl(targetDocument, sourceControls[sourceIndex], deletedControl);
            return true;
        }

        private static List<RedlineContentControlEntry> GetRedlineContentControlEntries(WordprocessingDocument document) {
            MainDocumentPart? mainPart = document.MainDocumentPart;
            if (mainPart == null) {
                return new List<RedlineContentControlEntry>();
            }

            var entries = new List<RedlineContentControlEntry>();
            foreach (WordFieldInventory.FieldRoot root in WordFieldInventory.EnumerateFieldRoots(mainPart)) {
                foreach (OrderedElement ordered in EnumerateDescendantsWithOrder(root.Root, GetFeatureOrderBase(root.LocationKind))) {
                    if (ordered.Element is SdtElement contentControl) {
                        entries.Add(new RedlineContentControlEntry(
                            entries.Count,
                            root.PartUri,
                            root.LocationKind,
                            GetNoteId(contentControl),
                            contentControl));
                    }
                }
            }

            return entries;
        }

        private static HashSet<int> GetContentControlRedlineTargetIndexes(
            WordComparisonResult result,
            WordComparisonRedlineOptions options,
            IReadOnlyList<RedlineContentControlEntry> targetControls) {
            var targetIndexes = new HashSet<int>();

            foreach (WordComparisonFinding finding in result.Findings) {
                if (TryGetContentControlRedlineTargetIndex(finding, options, targetControls, out int targetIndex)) {
                    RedlineContentControlEntry entry = targetControls[targetIndex];
                    if (HasChangedContentControlFindingText(finding)) {
                        targetIndexes.Add(targetIndex);
                    }
                }
            }

            return targetIndexes;
        }

        private static bool TryGetContentControlRedlineTargetIndex(
            WordComparisonFinding finding,
            WordComparisonRedlineOptions options,
            IReadOnlyList<RedlineContentControlEntry> targetControls,
            out int targetIndex) {
            targetIndex = finding.TargetIndex ?? -1;
            if (!ShouldTrackFinding(finding, options) ||
                finding.Scope != WordComparisonScope.ContentControl ||
                !HasTrackedText(finding)) {
                return false;
            }

            if (finding.ChangeKind != WordComparisonChangeKind.Modified &&
                finding.ChangeKind != WordComparisonChangeKind.Inserted) {
                return false;
            }

            if (targetIndex < 0 &&
                !TryParseIndexedLocation(finding.Location, "content-control", out targetIndex)) {
                return false;
            }

            return targetIndex >= 0 && targetIndex < targetControls.Count;
        }

        private static bool HasChangedContentControlFindingText(WordComparisonFinding finding) {
            string? sourceText = ExtractContentControlFindingText(finding.SourceText);
            string? targetText = ExtractContentControlFindingText(finding.TargetText);
            return !string.Equals(sourceText, targetText, StringComparison.Ordinal);
        }

        private static bool HasDescendantContentControlRedlineFinding(
            SdtElement contentControl,
            IReadOnlyList<RedlineContentControlEntry> targetControls,
            ISet<int> redlineTargetIndexes) {
            if (redlineTargetIndexes.Count == 0) {
                return false;
            }

            foreach (int targetIndex in redlineTargetIndexes) {
                SdtElement candidate = targetControls[targetIndex].ContentControl;
                if (!ReferenceEquals(candidate, contentControl) &&
                    candidate.Ancestors<SdtElement>().Any(ancestor => ReferenceEquals(ancestor, contentControl))) {
                    return true;
                }
            }

            return false;
        }

        private static bool RewriteContentControlWithTrackedText(SdtElement contentControl, string? sourceText, string? targetText, WordComparisonRedlineOptions options) {
            if (contentControl is SdtRun runControl) {
                runControl.SdtContentRun ??= new SdtContentRun();
                runControl.SdtContentRun.RemoveAllChildren();
                AppendTrackedInlineContent(runControl.SdtContentRun, sourceText, targetText, options);
                return true;
            }

            if (contentControl is SdtBlock blockControl) {
                blockControl.SdtContentBlock ??= new SdtContentBlock();
                blockControl.SdtContentBlock.RemoveAllChildren();

                var paragraph = new Paragraph();
                AppendTrackedInlineContent(paragraph, sourceText, targetText, options);
                blockControl.SdtContentBlock.Append(paragraph);
                return true;
            }

            if (contentControl is SdtCell cellControl) {
                cellControl.SdtContentCell ??= new SdtContentCell();
                TableCell? templateCell = cellControl.SdtContentCell.Elements<TableCell>().FirstOrDefault();
                TableCell trackedCell = CreateTrackedTableCell(templateCell, sourceText, targetText, options);

                cellControl.SdtContentCell.RemoveAllChildren();
                cellControl.SdtContentCell.Append(trackedCell);
                return true;
            }

            if (contentControl is SdtRow rowControl) {
                rowControl.SdtContentRow ??= new SdtContentRow();
                TableRow? templateRow = rowControl.SdtContentRow.Elements<TableRow>().FirstOrDefault();
                TableRow trackedRow = CreateTrackedTableRow(templateRow, sourceText, targetText, options);

                rowControl.SdtContentRow.RemoveAllChildren();
                rowControl.SdtContentRow.Append(trackedRow);
                return true;
            }

            return false;
        }

        private static TableRow CreateTrackedTableRow(TableRow? templateRow, string? sourceText, string? targetText, WordComparisonRedlineOptions options) {
            var row = new TableRow();
            TableRowProperties? rowProperties = templateRow?.GetFirstChild<TableRowProperties>()?.CloneNode(true) as TableRowProperties;
            if (rowProperties != null) {
                row.Append(rowProperties);
            }

            List<TableCell> templateCells = templateRow?.Elements<TableCell>().ToList() ?? new List<TableCell>();
            if (templateCells.Count == 0) {
                row.Append(CreateTrackedTableCell(null, sourceText, targetText, options));
                return row;
            }

            for (int index = 0; index < templateCells.Count; index++) {
                TableCell templateCell = templateCells[index];
                row.Append(index == 0
                    ? CreateTrackedTableCell(templateCell, sourceText, targetText, options)
                    : CreateEmptyTableCell(templateCell));
            }

            return row;
        }

        private static TableCell CreateTrackedTableCell(TableCell? templateCell, string? sourceText, string? targetText, WordComparisonRedlineOptions options) {
            var cell = new TableCell();
            TableCellProperties? cellProperties = templateCell?.GetFirstChild<TableCellProperties>()?.CloneNode(true) as TableCellProperties;
            if (cellProperties != null) {
                cell.Append(cellProperties);
            }

            var paragraph = new Paragraph();
            AppendTrackedInlineContent(paragraph, sourceText, targetText, options);
            cell.Append(paragraph);
            return cell;
        }

        private static TableCell CreateEmptyTableCell(TableCell templateCell) {
            var cell = new TableCell();
            TableCellProperties? cellProperties = templateCell.GetFirstChild<TableCellProperties>()?.CloneNode(true) as TableCellProperties;
            if (cellProperties != null) {
                cell.Append(cellProperties);
            }

            cell.Append(new Paragraph());
            return cell;
        }

        private static void AppendTrackedInlineContent(OpenXmlCompositeElement container, string? sourceText, string? targetText, WordComparisonRedlineOptions options) {
            if (!string.IsNullOrEmpty(sourceText)) {
                container.Append(CreateDeletedRun(sourceText!, options));
            }

            if (!string.IsNullOrEmpty(targetText)) {
                container.Append(CreateInsertedRun(targetText!, options));
            }

            if (!container.ChildElements.Any()) {
                container.Append(new Run());
            }
        }

        private static string? ExtractContentControlFindingText(string? displayText) {
            if (displayText == null) {
                return null;
            }

            const string textMarker = "; text=";
            int markerIndex = displayText.IndexOf(textMarker, StringComparison.Ordinal);
            if (markerIndex < 0) {
                return displayText;
            }

            return displayText.Substring(markerIndex + textMarker.Length);
        }

        private static void InsertDeletedContentControl(WordprocessingDocument targetDocument, RedlineContentControlEntry entry, SdtElement deletedControl) {
            OpenXmlCompositeElement? targetContainer = GetRedlineContentControlContainer(targetDocument, entry);
            if (targetContainer == null) {
                return;
            }

            OpenXmlElement block = deletedControl switch {
                SdtBlock => deletedControl,
                SdtRun => new Paragraph(deletedControl),
                SdtCell => new Table(new TableRow(deletedControl)),
                SdtRow => new Table(deletedControl),
                _ => new Paragraph()
            };

            InsertContentControlRedlineElement(targetContainer, block);
        }

        private static OpenXmlCompositeElement? GetRedlineContentControlContainer(WordprocessingDocument targetDocument, RedlineContentControlEntry entry) {
            MainDocumentPart? mainPart = targetDocument.MainDocumentPart;
            if (mainPart == null) {
                return null;
            }

            foreach (WordFieldInventory.FieldRoot root in WordFieldInventory.EnumerateFieldRoots(mainPart)) {
                if (!string.Equals(root.PartUri, entry.PartUri, StringComparison.Ordinal)) {
                    continue;
                }

                if (entry.LocationKind == WordFieldLocationKind.Footnote) {
                    return FindNoteContainer<Footnote>(root.Root, entry.NoteId);
                }

                if (entry.LocationKind == WordFieldLocationKind.Endnote) {
                    return FindNoteContainer<Endnote>(root.Root, entry.NoteId);
                }

                return root.Root;
            }

            return mainPart.Document?.Body;
        }

        private static OpenXmlCompositeElement? FindNoteContainer<TNote>(OpenXmlCompositeElement root, string? noteId)
            where TNote : OpenXmlCompositeElement {
            if (string.IsNullOrWhiteSpace(noteId)) {
                return null;
            }

            return root.Elements<TNote>().FirstOrDefault(note => {
                OpenXmlAttribute id = note.GetAttribute("id", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                return string.Equals(id.Value, noteId, StringComparison.Ordinal);
            });
        }

        private static string? GetNoteId(SdtElement contentControl) {
            Footnote? footnote = contentControl.Ancestors<Footnote>().FirstOrDefault();
            if (footnote?.Id?.Value != null) {
                return footnote.Id.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
            }

            Endnote? endnote = contentControl.Ancestors<Endnote>().FirstOrDefault();
            return endnote?.Id?.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        private static void InsertContentControlRedlineElement(OpenXmlCompositeElement container, OpenXmlElement element) {
            if (container is Body body) {
                SectionProperties? sectionProperties = body.Elements<SectionProperties>().LastOrDefault();
                if (sectionProperties != null) {
                    body.InsertBefore(element, sectionProperties);
                    return;
                }
            }

            if (container is Footnotes footnotes) {
                footnotes.Append(element);
                return;
            }

            if (container is Endnotes endnotes) {
                endnotes.Append(element);
                return;
            }

            container.Append(element);
        }

        private static bool TryParseIndexedLocation(string location, string prefix, out int index) {
            index = -1;
            string start = prefix + "[";
            if (!location.StartsWith(start, StringComparison.Ordinal) ||
                !location.EndsWith("]", StringComparison.Ordinal)) {
                return false;
            }

            string indexText = location.Substring(start.Length, location.Length - start.Length - 1);
            return int.TryParse(indexText, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out index);
        }

        private sealed class RedlineContentControlEntry {
            internal RedlineContentControlEntry(int index, string partUri, WordFieldLocationKind locationKind, string? noteId, SdtElement contentControl) {
                Index = index;
                PartUri = partUri;
                LocationKind = locationKind;
                NoteId = noteId;
                ContentControl = contentControl;
            }

            internal int Index { get; }

            internal string PartUri { get; }

            internal WordFieldLocationKind LocationKind { get; }

            internal string? NoteId { get; }

            internal SdtElement ContentControl { get; }
        }
    }
}
