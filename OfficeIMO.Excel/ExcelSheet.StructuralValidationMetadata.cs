using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private void RemapShiftedDataValidations(int firstAffectedRow, int rowDelta, int? lastDeletedRow) {
            DataValidations? validations = WorksheetRoot.GetFirstChild<DataValidations>();
            if (validations != null) {
                uint count = 0;
                foreach (DataValidation validation in validations.Elements<DataValidation>().ToList()) {
                    if (validation.SequenceOfReferences?.InnerText is not string references
                        || !TryGetReferenceListAnchorRow(references, out int oldAnchorRow)) {
                        count++;
                        continue;
                    }

                    string updatedReferences = references;
                    if (TryRemapShiftedReferenceListRows(
                        references,
                        firstAffectedRow,
                        rowDelta,
                        lastDeletedRow,
                        out List<string> remapped)) {
                        if (remapped.Count == 0) {
                            validation.Remove();
                            continue;
                        }

                        updatedReferences = string.Join(" ", remapped);
                        validation.SequenceOfReferences = new ListValue<StringValue> { InnerText = updatedReferences };
                    }

                    if (TryGetReferenceListAnchorRow(updatedReferences, out int newAnchorRow)) {
                        int anchorRowDelta = newAnchorRow - oldAnchorRow;
                        int relativeFormulaSourceRowDelta = GetRelativeFormulaSourceRowDelta(
                            oldAnchorRow,
                            newAnchorRow,
                            firstAffectedRow,
                            rowDelta,
                            lastDeletedRow);
                        RewriteAnchoredFormulaText(
                            validation.Formula1,
                            firstAffectedRow,
                            rowDelta,
                            lastDeletedRow,
                            anchorRowDelta,
                            relativeFormulaSourceRowDelta: relativeFormulaSourceRowDelta);
                        RewriteAnchoredFormulaText(
                            validation.Formula2,
                            firstAffectedRow,
                            rowDelta,
                            lastDeletedRow,
                            anchorRowDelta,
                            relativeFormulaSourceRowDelta: relativeFormulaSourceRowDelta);
                    }

                    count++;
                }

                if (count == 0U) {
                    validations.Remove();
                } else {
                    validations.Count = count;
                }
            }

            RemapShiftedOffice2010DataValidations(firstAffectedRow, rowDelta, lastDeletedRow);
        }

        private void RemapShiftedOffice2010DataValidations(
            int firstAffectedRow,
            int rowDelta,
            int? lastDeletedRow) {
            foreach (X14.DataValidations validations in WorksheetRoot.Descendants<X14.DataValidations>().ToList()) {
                uint count = 0;
                foreach (X14.DataValidation validation in validations.Elements<X14.DataValidation>().ToList()) {
                    if (validation.ReferenceSequence?.Text is not string references
                        || !TryGetReferenceListAnchorRow(references, out int oldAnchorRow)) {
                        count++;
                        continue;
                    }

                    string updatedReferences = references;
                    if (TryRemapShiftedReferenceListRows(
                        references,
                        firstAffectedRow,
                        rowDelta,
                        lastDeletedRow,
                        out List<string> remapped)) {
                        if (remapped.Count == 0) {
                            validation.Remove();
                            continue;
                        }

                        updatedReferences = string.Join(" ", remapped);
                        validation.ReferenceSequence.Text = updatedReferences;
                    }

                    if (TryGetReferenceListAnchorRow(updatedReferences, out int newAnchorRow)) {
                        int anchorRowDelta = newAnchorRow - oldAnchorRow;
                        int relativeFormulaSourceRowDelta = GetRelativeFormulaSourceRowDelta(
                            oldAnchorRow,
                            newAnchorRow,
                            firstAffectedRow,
                            rowDelta,
                            lastDeletedRow);
                        RewriteAnchoredFormulaText(
                            validation.DataValidationForumla1?.Formula,
                            firstAffectedRow,
                            rowDelta,
                            lastDeletedRow,
                            anchorRowDelta,
                            relativeFormulaSourceRowDelta: relativeFormulaSourceRowDelta);
                        RewriteAnchoredFormulaText(
                            validation.DataValidationForumla2?.Formula,
                            firstAffectedRow,
                            rowDelta,
                            lastDeletedRow,
                            anchorRowDelta,
                            relativeFormulaSourceRowDelta: relativeFormulaSourceRowDelta);
                    }

                    count++;
                }

                validations.Count = count;
                if (count == 0U) {
                    validations.Remove();
                }
            }
        }
    }
}
