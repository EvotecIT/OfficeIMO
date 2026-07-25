using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using Xm = DocumentFormat.OpenXml.Office.Excel;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private void RemapShiftedOffice2010ConditionalFormatting(
            int firstAffectedRow,
            int rowDelta,
            int? lastDeletedRow) {
            foreach (X14.ConditionalFormattings formattings in
                WorksheetRoot.Descendants<X14.ConditionalFormattings>().ToList()) {
                foreach (X14.ConditionalFormatting formatting in
                    formattings.Elements<X14.ConditionalFormatting>().ToList()) {
                    Xm.ReferenceSequence? target = formatting.GetFirstChild<Xm.ReferenceSequence>();
                    if (target?.Text is not string references
                        || !TryGetReferenceListAnchorRow(references, out int oldAnchorRow)) {
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
                            formatting.Remove();
                            continue;
                        }

                        updatedReferences = string.Join(" ", remapped);
                        target.Text = updatedReferences;
                    }

                    if (!TryGetReferenceListAnchorRow(updatedReferences, out int newAnchorRow)) {
                        continue;
                    }

                    int anchorRowDelta = newAnchorRow - oldAnchorRow;
                    foreach (Xm.Formula formula in formatting.Descendants<Xm.Formula>()) {
                        RewriteAnchoredFormulaText(
                            formula,
                            firstAffectedRow,
                            rowDelta,
                            lastDeletedRow,
                            anchorRowDelta);
                    }
                }

                if (!formattings.Elements<X14.ConditionalFormatting>().Any()) {
                    formattings.Remove();
                }
            }
        }
    }
}
