using OfficeIMO.Word.OpenXml;

namespace OfficeIMO.Word {
    public static partial class WordDocumentComparer {
        private static void ApplyContentControlFindings(WordPackageView sourceDocument, WordPackageView targetDocument, WordComparisonResult result, WordComparisonRedlineOptions options) {
            List<RedlineContentControlEntry> sourceControls = GetRedlineContentControlEntries(sourceDocument);
            List<RedlineContentControlEntry> targetControls = GetRedlineContentControlEntries(targetDocument);
            HashSet<int> deletedSourceIndexes = GetDeletedContentControlRedlineSourceIndexes(result, options, sourceControls);
            HashSet<int> redlineTargetIndexes = GetContentControlRedlineTargetIndexes(result, options, targetControls);
            var rewrittenControls = new HashSet<int>();

            foreach (WordComparisonFinding finding in result.Findings) {
                if (TryApplyDeletedContentControlFinding(finding, options, sourceControls, targetControls, deletedSourceIndexes, targetDocument)) {
                    continue;
                }

                if (!TryGetContentControlRedlineTargetIndex(finding, options, targetControls, out int targetIndex)) {
                    continue;
                }

                RedlineContentControlEntry entry = targetControls[targetIndex];
                if (rewrittenControls.Contains(targetIndex) ||
                    HasAncestorContentControlRedlineFinding(entry.ContentControl, targetControls, rewrittenControls)) {
                    continue;
                }

                if (HasDescendantContentControlRedlineFinding(entry.ContentControl, targetControls, redlineTargetIndexes) &&
                    HasUnchangedDirectContentControlText(finding, sourceControls, entry.ContentControl)) {
                    continue;
                }

                string? sourceText = GetContentControlRedlineText(sourceControls, finding.SourceIndex, finding.SourceText);
                string? targetText = GetContentControlRedlineText(targetControls, targetIndex, finding.TargetText);
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
            IReadOnlyList<RedlineContentControlEntry> targetControls,
            ISet<int> deletedSourceIndexes,
            WordPackageView targetDocument) {
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

            if (HasAncestorContentControlRedlineFinding(sourceControls[sourceIndex].ContentControl, sourceControls, deletedSourceIndexes)) {
                return true;
            }

            OfficeXmlNode deletedControl = sourceControls[sourceIndex].ContentControl.Clone();
            WordPackageView.CopyGridNormalizationState(sourceControls[sourceIndex].ContentControl, deletedControl);
            SdtKind deletedKind = GetSdtKind(sourceControls[sourceIndex].ContentControl);
            if (!RewriteContentControlWithTrackedText(deletedControl, deletedKind, GetContentControlRedlineText(sourceControls, sourceIndex, finding.SourceText), null, options)) {
                return false;
            }

            InsertDeletedContentControl(targetDocument, sourceControls, targetControls, sourceIndex, deletedControl, deletedKind);
            return true;
        }

        private static List<RedlineContentControlEntry> GetRedlineContentControlEntries(WordPackageView document) {
            WordPartView? mainPart = document.Main;
            if (mainPart == null) {
                return new List<RedlineContentControlEntry>();
            }

            var entries = new List<RedlineContentControlEntry>();
            foreach (WordFieldInventory.NodeFieldRoot root in WordFieldInventory.EnumerateFieldRoots(mainPart)) {
                foreach (OrderedElement ordered in EnumerateDescendantsWithOrder(root.Root, GetFeatureOrderBase(root.LocationKind))) {
                    if (ordered.Element.Is(W, "sdt")) {
                        entries.Add(new RedlineContentControlEntry(
                            entries.Count,
                            root.PartUri,
                            root.LocationKind,
                            GetContentControlNoteId(ordered.Element),
                            ordered.Element));
                    }
                }
            }

            return entries;
        }

        private static HashSet<int> GetDeletedContentControlRedlineSourceIndexes(
            WordComparisonResult result,
            WordComparisonRedlineOptions options,
            IReadOnlyList<RedlineContentControlEntry> sourceControls) {
            var sourceIndexes = new HashSet<int>();

            foreach (WordComparisonFinding finding in result.Findings) {
                if (!ShouldTrackFinding(finding, options) ||
                    finding.Scope != WordComparisonScope.ContentControl ||
                    finding.ChangeKind != WordComparisonChangeKind.Deleted ||
                    !HasTrackedText(finding) ||
                    !HasChangedContentControlFindingText(finding)) {
                    continue;
                }

                int sourceIndex = finding.SourceIndex ?? -1;
                if (sourceIndex < 0 &&
                    !TryParseIndexedLocation(finding.Location, "content-control", out sourceIndex)) {
                    continue;
                }

                if (sourceIndex >= 0 && sourceIndex < sourceControls.Count) {
                    sourceIndexes.Add(sourceIndex);
                }
            }

            return sourceIndexes;
        }

        private static HashSet<int> GetContentControlRedlineTargetIndexes(
            WordComparisonResult result,
            WordComparisonRedlineOptions options,
            IReadOnlyList<RedlineContentControlEntry> targetControls) {
            var targetIndexes = new HashSet<int>();

            foreach (WordComparisonFinding finding in result.Findings) {
                if (TryGetContentControlRedlineTargetIndex(finding, options, targetControls, out int targetIndex) &&
                    HasChangedContentControlFindingText(finding)) {
                    targetIndexes.Add(targetIndex);
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

        private static bool HasUnchangedDirectContentControlText(
            WordComparisonFinding finding,
            IReadOnlyList<RedlineContentControlEntry> sourceControls,
            OfficeXmlNode targetControl) {
            int sourceIndex = finding.SourceIndex ?? -1;
            if (sourceIndex < 0 &&
                !TryParseIndexedLocation(finding.Location, "content-control", out sourceIndex)) {
                return false;
            }

            if (sourceIndex < 0 || sourceIndex >= sourceControls.Count) {
                return false;
            }

            string sourceDirectText = GetDirectContentControlText(sourceControls[sourceIndex].ContentControl);
            string targetDirectText = GetDirectContentControlText(targetControl);
            return string.Equals(sourceDirectText, targetDirectText, StringComparison.Ordinal);
        }

        private static string GetDirectContentControlText(OfficeXmlNode contentControl) {
            return string.Concat(DescendantsW(contentControl, "t")
                .Where(text => !HasOtherContentControlAncestor(text, contentControl))
                .Select(text => text.InnerText));
        }

        private static bool HasOtherContentControlAncestor(OfficeXmlNode element, OfficeXmlNode contentControl) {
            for (OfficeXmlNode ancestor = element.Parent; !ancestor.IsNone; ancestor = ancestor.Parent) {
                if (ancestor.Is(W, "sdt") && ancestor != contentControl) {
                    return true;
                }
            }

            return false;
        }

        private static bool HasAncestorContentControlRedlineFinding(
            OfficeXmlNode contentControl,
            IReadOnlyList<RedlineContentControlEntry> contentControls,
            ISet<int> redlineIndexes) {
            if (redlineIndexes.Count == 0) {
                return false;
            }

            for (OfficeXmlNode ancestor = contentControl.Parent; !ancestor.IsNone; ancestor = ancestor.Parent) {
                if (!ancestor.Is(W, "sdt")) {
                    continue;
                }

                foreach (int redlineIndex in redlineIndexes) {
                    if (redlineIndex >= 0 &&
                        redlineIndex < contentControls.Count &&
                        contentControls[redlineIndex].ContentControl == ancestor) {
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool HasDescendantContentControlRedlineFinding(
            OfficeXmlNode contentControl,
            IReadOnlyList<RedlineContentControlEntry> targetControls,
            ISet<int> redlineTargetIndexes) {
            if (redlineTargetIndexes.Count == 0) {
                return false;
            }

            foreach (int targetIndex in redlineTargetIndexes) {
                OfficeXmlNode candidate = targetControls[targetIndex].ContentControl;
                if (candidate != contentControl && candidate.Parent.IsInside(contentControl) && !candidate.Parent.IsNone) {
                    return true;
                }
            }

            return false;
        }

        /// <summary>The typed class the SDK reads a <c>w:sdt</c> as, which depends on where it sits.</summary>
        private enum SdtKind {
            Block,
            Run,
            Cell,
            Row
        }

        private const string Word2010NamespaceUri = "http://schemas.microsoft.com/office/word/2010/wordml";

        /// <summary>A <c>w14:conflictIns</c> or <c>w14:conflictDel</c>, which holds run-level content.</summary>
        private static bool IsConflictRevision(OfficeXmlNode element) =>
            (string.Equals(element.LocalName, "conflictIns", StringComparison.Ordinal) ||
             string.Equals(element.LocalName, "conflictDel", StringComparison.Ordinal)) &&
            string.Equals(element.NamespaceUri, Word2010NamespaceUri, StringComparison.Ordinal);

        private static SdtKind GetSdtKind(OfficeXmlNode contentControl) {
            OfficeXmlNode parent = contentControl.Parent;
            for (int depth = 0; !parent.IsNone && depth < 256; depth++) {
                if (parent.NamespaceId != W) {
                    // The SDK read a control in a Word 2010 conflict revision as a run-level control.
                    return IsConflictRevision(parent) ? SdtKind.Run : SdtKind.Block;
                }

                switch (parent.LocalName) {
                    case "sdtContent":
                        // The content of a control holds what the control itself holds.
                        parent = parent.Parent.Parent;
                        continue;
                    case "customXml":
                    case "smartTag":
                        parent = parent.Parent;
                        if (parent.Is(W, "p")) return SdtKind.Run;
                        continue;
                    case "p":
                    case "hyperlink":
                    case "fldSimple":
                    case "ins":
                    case "del":
                    case "moveFrom":
                    case "moveTo":
                    case "dir":
                    case "bdo":
                        return SdtKind.Run;
                    case "tr":
                        return SdtKind.Cell;
                    case "tbl":
                        return SdtKind.Row;
                    default:
                        return SdtKind.Block;
                }
            }

            return SdtKind.Block;
        }

        private static string GetSdtKindName(SdtKind kind) => kind switch {
            SdtKind.Run => "SdtRun",
            SdtKind.Cell => "SdtCell",
            SdtKind.Row => "SdtRow",
            _ => "SdtBlock"
        };

        private static bool RewriteContentControlWithTrackedText(OfficeXmlNode contentControl, string? sourceText, string? targetText, WordComparisonRedlineOptions options) =>
            RewriteContentControlWithTrackedText(contentControl, GetSdtKind(contentControl), sourceText, targetText, options);

        private static bool RewriteContentControlWithTrackedText(OfficeXmlNode contentControl, SdtKind kind, string? sourceText, string? targetText, WordComparisonRedlineOptions options) {
            OfficeXmlNode content = contentControl.Child(W, "sdtContent");
            if (content.IsNone) {
                content = contentControl.InsertOrdered(contentControl.CreateElement(W, "sdtContent"), SdtOrder);
            }

            switch (kind) {
                case SdtKind.Run:
                    RemoveAllChildren(content);
                    AppendTrackedInlineContent(content, sourceText, targetText, options);
                    return true;
                case SdtKind.Block: {
                    RemoveAllChildren(content);
                    OfficeXmlNode paragraph = content.CreateElement(W, "p");
                    AppendTrackedInlineContent(paragraph, sourceText, targetText, options);
                    content.Append(paragraph);
                    return true;
                }
                case SdtKind.Cell: {
                    OfficeXmlNode templateCell = content.Child(W, "tc");
                    OfficeXmlNode trackedCell = CreateTrackedTableCell(content, templateCell, sourceText, targetText, options);
                    RemoveAllChildren(content);
                    content.Append(trackedCell);
                    return true;
                }
                case SdtKind.Row: {
                    OfficeXmlNode templateRow = content.Child(W, "tr");
                    OfficeXmlNode trackedRow = CreateTrackedTableRow(content, templateRow, sourceText, targetText, options);
                    RemoveAllChildren(content);
                    content.Append(trackedRow);
                    return true;
                }
                default:
                    return false;
            }
        }

        // w:sdt holds its properties, end properties and content in that order.
        private static readonly OfficeIMO.Core.Internal.OpenXml.OfficeXmlChildOrder SdtOrder = new(
            new[] { W, W, W },
            new[] { "sdtPr", "sdtEndPr", "sdtContent" },
            new[] { 0, 1, 2 });

        private static OfficeXmlNode CreateTrackedTableRow(OfficeXmlNode context, OfficeXmlNode templateRow, string? sourceText, string? targetText, WordComparisonRedlineOptions options) {
            OfficeXmlNode row = context.CreateElement(W, "tr");
            OfficeXmlNode rowProperties = templateRow.IsNone ? default : templateRow.Child(W, "trPr");
            if (!rowProperties.IsNone) {
                row.Append(rowProperties.Clone());
            }

            List<OfficeXmlNode> templateCells = templateRow.IsNone ? new List<OfficeXmlNode>() : ChildrenW(templateRow, "tc");
            if (templateCells.Count == 0) {
                row.Append(CreateTrackedTableCell(row, default, sourceText, targetText, options));
                return row;
            }

            for (int index = 0; index < templateCells.Count; index++) {
                OfficeXmlNode templateCell = templateCells[index];
                OfficeXmlNode cell = index == 0
                    ? CreateTrackedTableCell(row, templateCell, sourceText, targetText, options)
                    : templateCell.Clone();
                if (index != 0) WordPackageView.CopyGridNormalizationState(templateCell, cell);
                row.Append(cell);
            }

            return row;
        }

        private static OfficeXmlNode CreateTrackedTableCell(OfficeXmlNode context, OfficeXmlNode templateCell, string? sourceText, string? targetText, WordComparisonRedlineOptions options) {
            OfficeXmlNode cell = context.CreateElement(W, "tc");
            OfficeXmlNode cellProperties = templateCell.IsNone ? default : templateCell.Child(W, "tcPr");
            if (!cellProperties.IsNone) {
                cell.Append(cellProperties.Clone());
            }

            OfficeXmlNode paragraph = cell.CreateElement(W, "p");
            AppendTrackedInlineContent(paragraph, sourceText, targetText, options);
            cell.Append(paragraph);
            return cell;
        }

        private static void AppendTrackedInlineContent(OfficeXmlNode container, string? sourceText, string? targetText, WordComparisonRedlineOptions options) {
            if (!string.IsNullOrEmpty(sourceText)) {
                container.Append(CreateDeletedRun(container, sourceText!, options));
            }

            if (!string.IsNullOrEmpty(targetText)) {
                container.Append(CreateInsertedRun(container, targetText!, options));
            }

            if (container.FirstChild.IsNone) {
                container.Append(container.CreateElement(W, "r"));
            }
        }

        private static string? ExtractContentControlFindingText(string? displayText) {
            if (displayText == null) {
                return null;
            }

            const string textMarker = "; text=";
            int markerIndex = displayText.LastIndexOf(textMarker, StringComparison.Ordinal);
            if (markerIndex < 0) {
                return displayText;
            }

            return displayText.Substring(markerIndex + textMarker.Length);
        }

        private static string? GetContentControlRedlineText(IReadOnlyList<RedlineContentControlEntry> entries, int? index, string? fallbackDisplayText) {
            if (index.HasValue &&
                index.Value >= 0 &&
                index.Value < entries.Count) {
                return NormalizeFeatureText(InnerText(entries[index.Value].ContentControl));
            }

            return ExtractContentControlFindingText(fallbackDisplayText);
        }

        private static void InsertDeletedContentControl(
            WordPackageView targetDocument,
            IReadOnlyList<RedlineContentControlEntry> sourceControls,
            IReadOnlyList<RedlineContentControlEntry> targetControls,
            int sourceIndex,
            OfficeXmlNode deletedControl,
            SdtKind deletedKind) {
            RedlineContentControlEntry entry = sourceControls[sourceIndex];
            if (TryInsertDeletedContentControlAtTargetGap(sourceControls, targetControls, sourceIndex, deletedControl, deletedKind)) {
                return;
            }

            OfficeXmlNode targetContainer = GetRedlineContentControlContainer(targetDocument, entry);
            if (targetContainer.IsNone) {
                return;
            }

            OfficeXmlNode block = CreateDeletedContentControlRedlineBlock(targetContainer, deletedControl, deletedKind);
            InsertContentControlRedlineElement(targetContainer, block);
        }

        /// <summary>A block that can hold a deleted control where blocks go: the control itself, or a paragraph or table around it.</summary>
        private static OfficeXmlNode CreateDeletedContentControlRedlineBlock(OfficeXmlNode context, OfficeXmlNode deletedControl, SdtKind kind) {
            switch (kind) {
                case SdtKind.Block:
                    return deletedControl;
                case SdtKind.Run: {
                    OfficeXmlNode paragraph = context.CreateElement(W, "p");
                    paragraph.Append(deletedControl);
                    return paragraph;
                }
                case SdtKind.Cell: {
                    OfficeXmlNode table = context.CreateElement(W, "tbl");
                    OfficeXmlNode row = table.CreateElement(W, "tr");
                    row.Append(deletedControl);
                    table.Append(row);
                    return table;
                }
                case SdtKind.Row: {
                    OfficeXmlNode table = context.CreateElement(W, "tbl");
                    table.Append(deletedControl);
                    return table;
                }
                default:
                    return context.CreateElement(W, "p");
            }
        }

        /// <summary>How a block made by <see cref="CreateDeletedContentControlRedlineBlock"/> or a deleted control reads where it is inserted.</summary>
        private enum RedlineBlockKind {
            Paragraph,
            Table,
            BlockControl,
            RunControl,
            Other
        }

        private static bool TryInsertDeletedContentControlAtTargetGap(
            IReadOnlyList<RedlineContentControlEntry> sourceControls,
            IReadOnlyList<RedlineContentControlEntry> targetControls,
            int sourceIndex,
            OfficeXmlNode deletedControl,
            SdtKind deletedKind) {
            if (sourceIndex < 0 || sourceIndex >= sourceControls.Count || targetControls.Count == 0) {
                return false;
            }

            RedlineContentControlEntry sourceEntry = sourceControls[sourceIndex];
            for (int index = sourceIndex + 1; index < sourceControls.Count; index++) {
                RedlineContentControlEntry? targetEntry = FindMatchingTargetContentControl(sourceControls[index], targetControls, sourceEntry);
                if (TryInsertDeletedContentControlNearTarget(targetEntry, deletedControl, deletedKind, beforeTarget: true)) {
                    return true;
                }
            }

            for (int index = sourceIndex - 1; index >= 0; index--) {
                RedlineContentControlEntry? targetEntry = FindMatchingTargetContentControl(sourceControls[index], targetControls, sourceEntry);
                if (TryInsertDeletedContentControlNearTarget(targetEntry, deletedControl, deletedKind, beforeTarget: false)) {
                    return true;
                }
            }

            return false;
        }

        private static bool TryInsertDeletedContentControlNearTarget(RedlineContentControlEntry? targetEntry, OfficeXmlNode deletedControl, SdtKind deletedKind, bool beforeTarget) {
            if (targetEntry == null || targetEntry.ContentControl.Parent.IsNone) {
                return false;
            }

            RedlineBlockKind controlKind = deletedKind switch {
                SdtKind.Block => RedlineBlockKind.BlockControl,
                SdtKind.Run => RedlineBlockKind.RunControl,
                _ => RedlineBlockKind.Other
            };
            if (CanInsertContentControlGapBlockBeside(targetEntry.ContentControl, controlKind)) {
                if (beforeTarget) {
                    targetEntry.ContentControl.Parent.InsertBefore(deletedControl, targetEntry.ContentControl);
                } else {
                    targetEntry.ContentControl.InsertAfterSelf(deletedControl);
                }

                return true;
            }

            RedlineBlockKind blockKind = deletedKind switch {
                SdtKind.Block => RedlineBlockKind.BlockControl,
                SdtKind.Run => RedlineBlockKind.Paragraph,
                SdtKind.Cell => RedlineBlockKind.Table,
                SdtKind.Row => RedlineBlockKind.Table,
                _ => RedlineBlockKind.Paragraph
            };
            OfficeXmlNode anchor = GetContentControlGapAnchor(targetEntry.ContentControl, blockKind);
            if (anchor.IsNone || anchor.Parent.IsNone) {
                return false;
            }

            OfficeXmlNode block = CreateDeletedContentControlRedlineBlock(anchor.Parent, deletedControl, deletedKind);
            if (beforeTarget) {
                anchor.Parent.InsertBefore(block, anchor);
            } else {
                anchor.InsertAfterSelf(block);
            }

            return true;
        }

        private static OfficeXmlNode GetContentControlGapAnchor(OfficeXmlNode targetControl, RedlineBlockKind block) {
            if (CanInsertContentControlGapBlockBeside(targetControl, block)) {
                return targetControl;
            }

            OfficeXmlNode paragraph = AncestorW(targetControl, "p");
            if (!paragraph.IsNone && CanInsertContentControlGapBlockBeside(paragraph, block)) {
                return paragraph;
            }

            OfficeXmlNode table = AncestorW(targetControl, "tbl");
            if (!table.IsNone && CanInsertContentControlGapBlockBeside(table, block)) {
                return table;
            }

            return default;
        }

        private static bool CanInsertContentControlGapBlockBeside(OfficeXmlNode anchor, RedlineBlockKind block) {
            OfficeXmlNode parent = anchor.Parent;
            if (parent.NamespaceId != W) {
                return false;
            }

            switch (parent.LocalName) {
                case "body":
                case "hdr":
                case "ftr":
                case "footnote":
                case "endnote":
                case "tc":
                    return block is RedlineBlockKind.Paragraph or RedlineBlockKind.Table or RedlineBlockKind.BlockControl;
                case "sdtContent":
                    // The content of a block control holds blocks; the content of a run control holds runs.
                    SdtKind owner = parent.Parent.Is(W, "sdt") ? GetSdtKind(parent.Parent) : SdtKind.Block;
                    if (owner == SdtKind.Block) return block is RedlineBlockKind.Paragraph or RedlineBlockKind.Table or RedlineBlockKind.BlockControl;
                    if (owner == SdtKind.Run) return block is RedlineBlockKind.RunControl;
                    return false;
                case "p":
                    return block is RedlineBlockKind.RunControl;
                default:
                    return false;
            }
        }

        private static RedlineContentControlEntry? FindMatchingTargetContentControl(
            RedlineContentControlEntry sourceCandidate,
            IReadOnlyList<RedlineContentControlEntry> targetControls,
            RedlineContentControlEntry deletedSourceEntry) {
            string sourceKey = GetContentControlPlacementKey(sourceCandidate.ContentControl);
            if (string.IsNullOrEmpty(sourceKey)) {
                return null;
            }

            return targetControls.FirstOrDefault(target =>
                string.Equals(target.PartUri, deletedSourceEntry.PartUri, StringComparison.Ordinal) &&
                target.LocationKind == deletedSourceEntry.LocationKind &&
                string.Equals(target.NoteId, deletedSourceEntry.NoteId, StringComparison.Ordinal) &&
                string.Equals(GetContentControlPlacementKey(target.ContentControl), sourceKey, StringComparison.Ordinal));
        }

        private static string GetContentControlPlacementKey(OfficeXmlNode contentControl) {
            OfficeXmlNode properties = contentControl.Child(W, "sdtPr");
            string alias = ChildVal(properties, "alias") ?? string.Empty;
            string tag = ChildVal(properties, "tag") ?? string.Empty;
            OfficeXmlNode dataBinding = properties.IsNone ? default : properties.Child(W, "dataBinding");
            string binding = dataBinding.IsNone
                ? string.Empty
                : string.Join("|", dataBinding.Attribute(W, "storeItemID") ?? string.Empty, dataBinding.Attribute(W, "xpath") ?? string.Empty, dataBinding.Attribute(W, "prefixMappings") ?? string.Empty);

            string typeName = GetSdtKindName(GetSdtKind(contentControl));
            if (!string.IsNullOrWhiteSpace(alias) ||
                !string.IsNullOrWhiteSpace(tag) ||
                !string.IsNullOrWhiteSpace(binding)) {
                return string.Join("|", typeName, alias, tag, binding);
            }

            return typeName + "|" + NormalizeFeatureText(InnerText(contentControl));
        }

        private static OfficeXmlNode GetRedlineContentControlContainer(WordPackageView targetDocument, RedlineContentControlEntry entry) {
            WordPartView? mainPart = targetDocument.Main;
            if (mainPart == null) {
                return default;
            }

            foreach (WordFieldInventory.NodeFieldRoot root in WordFieldInventory.EnumerateFieldRoots(mainPart)) {
                if (!string.Equals(root.PartUri, entry.PartUri, StringComparison.Ordinal)) {
                    continue;
                }

                if (entry.LocationKind == WordFieldLocationKind.Footnote) {
                    return FindNoteContainer(root.Root, "footnote", entry.NoteId);
                }

                if (entry.LocationKind == WordFieldLocationKind.Endnote) {
                    return FindNoteContainer(root.Root, "endnote", entry.NoteId);
                }

                return root.Root;
            }

            return BodyOf(mainPart);
        }

        private static OfficeXmlNode FindNoteContainer(OfficeXmlNode root, string noteName, string? noteId) {
            if (string.IsNullOrWhiteSpace(noteId)) {
                return default;
            }

            foreach (OfficeXmlNode note in root.Children(W, noteName)) {
                if (string.Equals(note.Attribute(W, "id"), noteId, StringComparison.Ordinal)) {
                    return note;
                }
            }

            return default;
        }

        private static string? GetContentControlNoteId(OfficeXmlNode contentControl) {
            OfficeXmlNode footnote = AncestorW(contentControl, "footnote");
            long? footnoteId = footnote.IsNone ? null : GetNoteId(footnote);
            if (footnoteId != null) {
                return footnoteId.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
            }

            OfficeXmlNode endnote = AncestorW(contentControl, "endnote");
            return endnote.IsNone ? null : GetNoteId(endnote)?.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        private static void InsertContentControlRedlineElement(OfficeXmlNode container, OfficeXmlNode element) {
            if (container.Is(W, "body")) {
                OfficeXmlNode sectionProperties = container.Children(W, "sectPr").Last();
                if (!sectionProperties.IsNone) {
                    container.InsertBefore(element, sectionProperties);
                    return;
                }
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
            internal RedlineContentControlEntry(int index, string partUri, WordFieldLocationKind locationKind, string? noteId, OfficeXmlNode contentControl) {
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

            internal OfficeXmlNode ContentControl { get; }
        }
    }
}
