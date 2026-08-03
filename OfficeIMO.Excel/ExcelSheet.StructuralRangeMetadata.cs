using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using Xnsv = DocumentFormat.OpenXml.Office2021.Excel.NamedSheetViews;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private void RemapShiftedProtectedRanges(
            int firstAffectedRow,
            int rowDelta,
            int? lastDeletedRow) {
            ProtectedRanges? protectedRanges = WorksheetRoot.GetFirstChild<ProtectedRanges>();
            if (protectedRanges == null) {
                return;
            }

            foreach (ProtectedRange protectedRange in protectedRanges.Elements<ProtectedRange>().ToList()) {
                if (protectedRange.SequenceOfReferences?.InnerText is not string references
                    || !TryRemapShiftedReferenceListRows(
                        references,
                        firstAffectedRow,
                        rowDelta,
                        lastDeletedRow,
                        out List<string> remapped)) {
                    continue;
                }

                if (remapped.Count == 0) {
                    protectedRange.Remove();
                } else {
                    protectedRange.SequenceOfReferences = new ListValue<StringValue> {
                        InnerText = string.Join(" ", remapped)
                    };
                }
            }

            if (!protectedRanges.Elements<ProtectedRange>().Any()) {
                protectedRanges.Remove();
            }
        }

        private void RemapShiftedIgnoredErrors(
            int firstAffectedRow,
            int rowDelta,
            int? lastDeletedRow) {
            IgnoredErrors? ignoredErrors = WorksheetRoot.GetFirstChild<IgnoredErrors>();
            if (ignoredErrors != null) {
                foreach (IgnoredError ignoredError in ignoredErrors.Elements<IgnoredError>().ToList()) {
                    if (ignoredError.SequenceOfReferences?.InnerText is not string references
                        || !TryRemapShiftedReferenceListRows(
                            references,
                            firstAffectedRow,
                            rowDelta,
                            lastDeletedRow,
                            out List<string> remapped)) {
                        continue;
                    }

                    if (remapped.Count == 0) {
                        ignoredError.Remove();
                    } else {
                        ignoredError.SequenceOfReferences = new ListValue<StringValue> {
                            InnerText = string.Join(" ", remapped)
                        };
                    }
                }

                if (!ignoredErrors.Elements<IgnoredError>().Any()) {
                    ignoredErrors.Remove();
                }
            }

            foreach (X14.IgnoredErrors extendedErrors in WorksheetRoot.Descendants<X14.IgnoredErrors>().ToList()) {
                foreach (X14.IgnoredError ignoredError in extendedErrors.Elements<X14.IgnoredError>().ToList()) {
                    if (ignoredError.ReferenceSequence?.Text is not string references
                        || !TryRemapShiftedReferenceListRows(
                            references,
                            firstAffectedRow,
                            rowDelta,
                            lastDeletedRow,
                            out List<string> remapped)) {
                        continue;
                    }

                    if (remapped.Count == 0) {
                        ignoredError.Remove();
                    } else {
                        ignoredError.ReferenceSequence.Text = string.Join(" ", remapped);
                    }
                }

                if (!extendedErrors.Elements<X14.IgnoredError>().Any()) {
                    extendedErrors.Remove();
                }
            }
        }

        private void RemapShiftedScenarios(
            int firstAffectedRow,
            int rowDelta,
            int? lastDeletedRow) {
            Scenarios? scenarios = WorksheetRoot.GetFirstChild<Scenarios>();
            if (scenarios == null) {
                return;
            }

            if (scenarios.SequenceOfReferences?.InnerText is string references
                && TryRemapShiftedReferenceListRows(
                    references,
                    firstAffectedRow,
                    rowDelta,
                    lastDeletedRow,
                    out List<string> remappedResults)) {
                scenarios.SequenceOfReferences = remappedResults.Count == 0
                    ? null
                    : new ListValue<StringValue> { InnerText = string.Join(" ", remappedResults) };
                if (remappedResults.Count == 0) {
                    scenarios.Remove();
                    return;
                }
            }

            List<Scenario> originalScenarios = scenarios.Elements<Scenario>().ToList();
            var removedScenarioIndices = new HashSet<int>();
            for (int scenarioIndex = 0; scenarioIndex < originalScenarios.Count; scenarioIndex++) {
                Scenario scenario = originalScenarios[scenarioIndex];
                bool inputsChanged = false;
                foreach (InputCells input in scenario.Elements<InputCells>().ToList()) {
                    if (input.CellReference?.Value is not string reference
                        || !TryRemapShiftedReferenceListRows(
                            reference,
                            firstAffectedRow,
                            rowDelta,
                            lastDeletedRow,
                            out List<string> remappedInputs)) {
                        continue;
                    }

                    inputsChanged = true;
                    if (remappedInputs.Count == 0) {
                        input.Remove();
                    } else {
                        input.CellReference = remappedInputs[0];
                    }
                }

                uint inputCount = (uint)scenario.Elements<InputCells>().Count();
                if (inputCount == 0U) {
                    scenario.Remove();
                    removedScenarioIndices.Add(scenarioIndex);
                } else if (inputsChanged) {
                    scenario.Count = inputCount;
                }
            }

            int survivingScenarioCount = originalScenarios.Count - removedScenarioIndices.Count;
            if (survivingScenarioCount == 0) {
                scenarios.Remove();
                return;
            }

            if (removedScenarioIndices.Count > 0) {
                if (scenarios.Current?.Value is uint current) {
                    scenarios.Current = RemapScenarioIndex(
                        current,
                        originalScenarios.Count,
                        removedScenarioIndices);
                }
                if (scenarios.Show?.Value is uint shown) {
                    scenarios.Show = RemapScenarioIndex(
                        shown,
                        originalScenarios.Count,
                        removedScenarioIndices);
                }
            }
        }

        private static uint RemapScenarioIndex(
            uint index,
            int originalCount,
            ISet<int> removedIndices) {
            int oldIndex = index >= (uint)originalCount
                ? originalCount - 1
                : (int)index;
            int newIndex = 0;
            int lastSurvivingNewIndex = 0;
            for (int candidate = 0; candidate < originalCount; candidate++) {
                if (removedIndices.Contains(candidate)) {
                    continue;
                }

                lastSurvivingNewIndex = newIndex;
                if (candidate >= oldIndex) {
                    return (uint)newIndex;
                }
                newIndex++;
            }

            return (uint)lastSurvivingNewIndex;
        }

        private void RemapShiftedCellWatches(
            int firstAffectedRow,
            int rowDelta,
            int? lastDeletedRow) {
            CellWatches? watches = WorksheetRoot.GetFirstChild<CellWatches>();
            if (watches == null) {
                return;
            }

            foreach (CellWatch watch in watches.Elements<CellWatch>().ToList()) {
                if (watch.CellReference?.Value is not string reference
                    || !TryRemapShiftedReferenceListRows(
                        reference,
                        firstAffectedRow,
                        rowDelta,
                        lastDeletedRow,
                        out List<string> remapped)) {
                    continue;
                }

                if (remapped.Count == 0) {
                    watch.Remove();
                } else {
                    watch.CellReference = remapped[0];
                }
            }

            if (!watches.Elements<CellWatch>().Any()) {
                watches.Remove();
            }
        }

        private void RemapShiftedDataConsolidationReferences(
            int firstAffectedRow,
            int rowDelta,
            int? lastDeletedRow) {
            foreach (WorksheetPart worksheetPart in WorkbookPartRoot.WorksheetParts) {
                Worksheet? worksheet = worksheetPart.Worksheet;
                if (worksheet == null) {
                    continue;
                }

                bool changed = false;
                foreach (DataReference source in worksheet.Descendants<DataReference>().ToList()) {
                    if (!string.IsNullOrWhiteSpace(source.Id?.Value)
                        || !string.Equals(source.Sheet?.Value, Name, StringComparison.OrdinalIgnoreCase)
                        || source.Reference?.Value is not string reference
                        || !TryRemapShiftedReferenceListRows(
                            reference,
                            firstAffectedRow,
                            rowDelta,
                            lastDeletedRow,
                            out List<string> remapped)) {
                        continue;
                    }

                    if (remapped.Count == 0) {
                        source.Remove();
                    } else {
                        source.Reference = remapped[0];
                    }
                    changed = true;
                }

                if (!changed) {
                    continue;
                }

                foreach (DataReferences sources in worksheet.Descendants<DataReferences>().ToList()) {
                    uint sourceCount = (uint)sources.Elements<DataReference>().Count();
                    if (sourceCount == 0U) {
                        sources.Remove();
                    } else {
                        sources.Count = sourceCount;
                    }
                }
                worksheet.Save();
            }
        }

        private void RemapShiftedWebPublishItems(
            int firstAffectedRow,
            int rowDelta,
            int? lastDeletedRow) {
            WebPublishItems? items = WorkbookRoot.GetFirstChild<WebPublishItems>();
            if (items == null) {
                return;
            }

            bool changed = false;
            foreach (WebPublishItem item in items.Elements<WebPublishItem>().ToList()) {
                if (item.SourceType?.Value != WebSourceValues.Range
                    || !string.Equals(item.SourceObject?.Value, Name, StringComparison.OrdinalIgnoreCase)
                    || item.SourceRef?.Value is not string reference
                    || !TryRemapShiftedReferenceListRows(
                        reference,
                        firstAffectedRow,
                        rowDelta,
                        lastDeletedRow,
                        out List<string> remapped)) {
                    continue;
                }

                if (remapped.Count == 0) {
                    item.Remove();
                } else {
                    item.SourceRef = remapped[0];
                }
                changed = true;
            }

            if (!changed) {
                return;
            }
            uint count = (uint)items.Elements<WebPublishItem>().Count();
            if (count == 0U) {
                items.Remove();
            } else {
                items.Count = count;
            }
            WorkbookRoot.Save();
        }

        private void RemapShiftedCellSmartTags(
            int firstAffectedRow,
            int rowDelta,
            int? lastDeletedRow) {
            var affectedContainers = new HashSet<OpenXmlElement>();
            foreach (OpenXmlElement tag in WorksheetRoot.Descendants()
                .Where(element => string.Equals(element.LocalName, "cellSmartTag", StringComparison.OrdinalIgnoreCase))
                .ToList()) {
                OpenXmlAttribute referenceAttribute = tag.GetAttributes()
                    .FirstOrDefault(attribute => string.Equals(attribute.LocalName, "r", StringComparison.OrdinalIgnoreCase));
                string? referenceValue = referenceAttribute.Value;
                if (string.IsNullOrWhiteSpace(referenceValue)
                    || !TryRemapShiftedReferenceListRows(
                        referenceValue!,
                        firstAffectedRow,
                        rowDelta,
                        lastDeletedRow,
                        out List<string> remapped)) {
                    continue;
                }

                if (tag.Parent is OpenXmlElement container
                    && string.Equals(
                        container.LocalName,
                        "cellSmartTags",
                        StringComparison.OrdinalIgnoreCase)) {
                    affectedContainers.Add(container);
                }
                if (remapped.Count == 0) {
                    tag.Remove();
                } else {
                    tag.SetAttribute(new OpenXmlAttribute(
                        referenceAttribute.Prefix,
                        referenceAttribute.LocalName,
                        referenceAttribute.NamespaceUri,
                        remapped[0]));
                }
            }

            foreach (OpenXmlElement container in affectedContainers) {
                uint count = (uint)container.ChildElements.Count(child =>
                    string.Equals(child.LocalName, "cellSmartTag", StringComparison.OrdinalIgnoreCase));
                if (count == 0U) {
                    container.Remove();
                } else {
                    OpenXmlAttribute countAttribute = container.GetAttributes()
                        .FirstOrDefault(attribute => string.Equals(
                            attribute.LocalName,
                            "count",
                            StringComparison.OrdinalIgnoreCase));
                    container.SetAttribute(new OpenXmlAttribute(
                        countAttribute.Prefix,
                        "count",
                        countAttribute.NamespaceUri,
                        count.ToString(System.Globalization.CultureInfo.InvariantCulture)));
                }
            }
        }

        private static bool RemapShiftedSortStateReferences(
            OpenXmlElement root,
            int firstAffectedRow,
            int rowDelta,
            int? lastDeletedRow) {
            bool changed = false;
            foreach (SortState sortState in root.Descendants<SortState>().ToList()) {
                if (sortState.Reference?.Value is string sortReference
                    && TryRemapShiftedReferenceListRows(
                        sortReference,
                        firstAffectedRow,
                        rowDelta,
                        lastDeletedRow,
                        out List<string> remappedSortReferences)) {
                    if (remappedSortReferences.Count == 0) {
                        sortState.Remove();
                        changed = true;
                        continue;
                    }

                    string updatedReference = remappedSortReferences[0];
                    if (!string.Equals(sortReference, updatedReference, StringComparison.OrdinalIgnoreCase)) {
                        sortState.Reference = updatedReference;
                        changed = true;
                    }
                }

                foreach (SortCondition condition in sortState.Elements<SortCondition>().ToList()) {
                    if (condition.Reference?.Value is not string conditionReference
                        || !TryRemapShiftedReferenceListRows(
                            conditionReference,
                            firstAffectedRow,
                            rowDelta,
                            lastDeletedRow,
                            out List<string> remappedConditionReferences)) {
                        continue;
                    }

                    if (remappedConditionReferences.Count == 0) {
                        condition.Remove();
                        changed = true;
                        continue;
                    }

                    string updatedReference = remappedConditionReferences[0];
                    if (!string.Equals(conditionReference, updatedReference, StringComparison.OrdinalIgnoreCase)) {
                        condition.Reference = updatedReference;
                        changed = true;
                    }
                }

                foreach (X14.SortCondition condition in sortState.Elements<X14.SortCondition>().ToList()) {
                    if (condition.Reference?.Value is not string conditionReference
                        || !TryRemapShiftedReferenceListRows(
                            conditionReference,
                            firstAffectedRow,
                            rowDelta,
                            lastDeletedRow,
                            out List<string> remappedConditionReferences)) {
                        continue;
                    }

                    if (remappedConditionReferences.Count == 0) {
                        condition.Remove();
                        changed = true;
                        continue;
                    }

                    string updatedReference = remappedConditionReferences[0];
                    if (!string.Equals(conditionReference, updatedReference, StringComparison.OrdinalIgnoreCase)) {
                        condition.Reference = updatedReference;
                        changed = true;
                    }
                }
            }

            return changed;
        }

        private void RemapShiftedQueryTableSortStates(
            int firstAffectedRow,
            int rowDelta,
            int? lastDeletedRow) {
            foreach (QueryTablePart part in ExcelPackageQueryTableParts.Enumerate(_worksheetPart)) {
                QueryTable? queryTable = part.QueryTable;
                if (queryTable != null
                    && RemapShiftedSortStateReferences(
                        queryTable,
                        firstAffectedRow,
                        rowDelta,
                        lastDeletedRow)) {
                    queryTable.Save();
                }
            }
        }

        private void RemapShiftedSelections(
            int firstAffectedRow,
            int rowDelta,
            int? lastDeletedRow) {
            foreach (SheetView view in WorksheetRoot.Descendants<SheetView>()) {
                string? current = view.TopLeftCell?.Value;
                string? remapped = RemapShiftedViewTopLeftCell(
                    current,
                    firstAffectedRow,
                    rowDelta,
                    lastDeletedRow);
                if (!string.Equals(current, remapped, StringComparison.OrdinalIgnoreCase)) {
                    view.TopLeftCell = remapped;
                }
            }
            foreach (CustomSheetView view in WorksheetRoot.Descendants<CustomSheetView>()) {
                string? current = view.TopLeftCell?.Value;
                string? remapped = RemapShiftedViewTopLeftCell(
                    current,
                    firstAffectedRow,
                    rowDelta,
                    lastDeletedRow);
                if (!string.Equals(current, remapped, StringComparison.OrdinalIgnoreCase)) {
                    view.TopLeftCell = remapped;
                }
            }
            foreach (Pane pane in WorksheetRoot.Descendants<Pane>()) {
                string? current = pane.TopLeftCell?.Value;
                string? remapped = RemapShiftedViewTopLeftCell(
                    current,
                    firstAffectedRow,
                    rowDelta,
                    lastDeletedRow);
                if (!string.Equals(current, remapped, StringComparison.OrdinalIgnoreCase)) {
                    pane.TopLeftCell = remapped;
                }
            }

            foreach (Selection selection in WorksheetRoot.Descendants<Selection>()) {
                string? activeCell = selection.ActiveCell?.Value;
                if (!string.IsNullOrWhiteSpace(activeCell)
                    && TryRemapShiftedReferenceListRows(
                        activeCell!,
                        firstAffectedRow,
                        rowDelta,
                        lastDeletedRow,
                        out List<string> remappedActiveCell)) {
                    selection.ActiveCell = remappedActiveCell.Count > 0
                        ? remappedActiveCell[0]
                        : lastDeletedRow.HasValue
                            ? ClampDeletedSelectionReference(activeCell!, firstAffectedRow)
                            : activeCell;
                }

                string? references = selection.SequenceOfReferences?.InnerText;
                if (string.IsNullOrWhiteSpace(references)
                    || !TryRemapShiftedReferenceListRows(
                        references!,
                        firstAffectedRow,
                        rowDelta,
                        lastDeletedRow,
                        out List<string> remappedReferences)) {
                    continue;
                }

                string fallback = selection.ActiveCell?.Value ?? "A1";
                List<string> finalReferences = remappedReferences.Count > 0
                    ? remappedReferences
                    : new List<string> { fallback };
                selection.SequenceOfReferences = new ListValue<StringValue> {
                    InnerText = string.Join(" ", finalReferences)
                };
                RemapSelectionActiveCellId(selection, finalReferences);
            }
        }

        private static void RemapSelectionActiveCellId(
            Selection selection,
            IReadOnlyList<string> references) {
            if (selection.ActiveCellId == null || references.Count == 0) {
                return;
            }

            string? activeCell = selection.ActiveCell?.Value;
            if (!string.IsNullOrWhiteSpace(activeCell)
                && TryParseCellOrRange(
                    activeCell!.Replace("$", string.Empty),
                    out int activeRow,
                    out int activeColumn,
                    out _,
                    out _)) {
                for (int index = 0; index < references.Count; index++) {
                    if (TryParseCellOrRange(
                            references[index].Replace("$", string.Empty),
                            out int firstRow,
                            out int firstColumn,
                            out int lastRow,
                            out int lastColumn)
                        && activeRow >= firstRow
                        && activeRow <= lastRow
                        && activeColumn >= firstColumn
                        && activeColumn <= lastColumn) {
                        selection.ActiveCellId = (uint)index;
                        return;
                    }
                }
            }

            selection.ActiveCellId = 0U;
        }

        private static bool TryParseCellOrRange(
            string reference,
            out int firstRow,
            out int firstColumn,
            out int lastRow,
            out int lastColumn) {
            if (A1.TryParseRange(
                    reference,
                    out firstRow,
                    out firstColumn,
                    out lastRow,
                    out lastColumn)) {
                return true;
            }

            if (A1.TryParseCellReferenceFast(reference, out firstRow, out firstColumn)) {
                lastRow = firstRow;
                lastColumn = firstColumn;
                return true;
            }

            firstRow = firstColumn = lastRow = lastColumn = 0;
            return false;
        }

        private void RemapShiftedNamedSheetViewFilters(
            int firstAffectedRow,
            int rowDelta,
            int? lastDeletedRow) {
            foreach (NamedSheetViewsPart part in _worksheetPart.NamedSheetViewsParts) {
                Xnsv.NamedSheetViews? views = part.NamedSheetViews;
                if (views == null) {
                    continue;
                }

                bool changed = false;
                foreach (Xnsv.NsvFilter filter in views.Descendants<Xnsv.NsvFilter>().ToList()) {
                    if (filter.Ref?.Value is not string reference
                        || !TryRemapShiftedReferenceListRows(
                            reference,
                            firstAffectedRow,
                            rowDelta,
                            lastDeletedRow,
                            out List<string> remapped)) {
                        continue;
                    }

                    if (remapped.Count == 0) {
                        filter.Remove();
                    } else {
                        filter.Ref = remapped[0];
                    }
                    changed = true;
                }

                if (changed) {
                    views.Save();
                }
            }
        }

        private static string? RemapShiftedViewTopLeftCell(
            string? reference,
            int firstAffectedRow,
            int rowDelta,
            int? lastDeletedRow) {
            if (string.IsNullOrWhiteSpace(reference)
                || !TryRemapShiftedReferenceListRows(
                    reference!,
                    firstAffectedRow,
                    rowDelta,
                    lastDeletedRow,
                    out List<string> remapped)) {
                return reference;
            }

            return remapped.Count > 0
                ? remapped[0]
                : lastDeletedRow.HasValue
                    ? ClampDeletedSelectionReference(reference!, firstAffectedRow)
                    : reference;
        }

        private static string ClampDeletedSelectionReference(
            string reference,
            int firstDeletedRow) {
            return TryParseReference(reference, out var bounds)
                ? A1.CellReference(firstDeletedRow, bounds.c1)
                : "A1";
        }
    }
}
