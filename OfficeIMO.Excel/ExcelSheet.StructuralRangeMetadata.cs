using DocumentFormat.OpenXml;
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
