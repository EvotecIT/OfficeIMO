using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

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
            }

            foreach (Scenario scenario in scenarios.Elements<Scenario>().ToList()) {
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

                    if (remappedInputs.Count == 0) {
                        input.Remove();
                    } else {
                        input.CellReference = remappedInputs[0];
                    }
                }

                uint inputCount = (uint)scenario.Elements<InputCells>().Count();
                if (inputCount == 0U) {
                    scenario.Remove();
                } else {
                    scenario.Count = inputCount;
                }
            }

            if (!scenarios.Elements<Scenario>().Any()) {
                scenarios.Remove();
            }
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

            foreach (OpenXmlElement container in WorksheetRoot.Descendants()
                .Where(element => string.Equals(element.LocalName, "smartTags", StringComparison.OrdinalIgnoreCase))
                .ToList()) {
                if (!container.ChildElements.Any(child =>
                    string.Equals(child.LocalName, "cellSmartTag", StringComparison.OrdinalIgnoreCase))) {
                    container.Remove();
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
            }

            return changed;
        }
    }
}
