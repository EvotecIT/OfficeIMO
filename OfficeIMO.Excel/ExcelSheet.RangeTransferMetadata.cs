using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using Xm = DocumentFormat.OpenXml.Office.Excel;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private void RemoveRangeTransferDestinationMetadata(
            ExcelReference source,
            int destinationFirstRow,
            int destinationFirstColumn,
            int destinationLastRow,
            int destinationLastColumn) {
            source.GetBounds(out int sourceFirstRow, out int sourceFirstColumn, out int sourceLastRow, out int sourceLastColumn);
            var destinationOnly = new List<string>(4);
            AppendRangeDifference(
                destinationOnly,
                (destinationFirstRow, destinationFirstColumn, destinationLastRow, destinationLastColumn),
                (sourceFirstRow, sourceFirstColumn, sourceLastRow, sourceLastColumn));
            foreach (string range in destinationOnly) {
                RemoveDataValidationsCore(range);
                ClearConditionalFormattingCore(range);
                (int r1, int c1, int r2, int c2) bounds = ParseReferenceArgument(range);
                RemoveProtectedRangeOverlap(bounds);
                RemoveIgnoredErrorOverlap(bounds);
                RemoveOffice2010ValidationOverlap(bounds);
                RemoveOffice2010ConditionalFormattingOverlap(bounds);
            }
            CleanupEmptyMetadataExtensions();
        }

        private void RemoveProtectedRangeOverlap((int r1, int c1, int r2, int c2) bounds) {
            ProtectedRanges? ranges = WorksheetRoot.GetFirstChild<ProtectedRanges>();
            if (ranges == null) return;
            foreach (ProtectedRange range in ranges.Elements<ProtectedRange>().ToList()) {
                string references = range.SequenceOfReferences?.InnerText ?? string.Empty;
                if (!TryRemoveReferenceOverlap(references, bounds, out List<string> remaining)) continue;
                if (remaining.Count == 0) range.Remove();
                else {
                    range.SequenceOfReferences = new ListValue<StringValue> {
                        InnerText = string.Join(" ", remaining)
                    };
                }
            }
            if (!ranges.Elements<ProtectedRange>().Any()) ranges.Remove();
        }

        private void RemoveIgnoredErrorOverlap((int r1, int c1, int r2, int c2) bounds) {
            IgnoredErrors? errors = WorksheetRoot.GetFirstChild<IgnoredErrors>();
            if (errors != null) {
                foreach (IgnoredError error in errors.Elements<IgnoredError>().ToList()) {
                    string references = error.SequenceOfReferences?.InnerText ?? string.Empty;
                    if (!TryRemoveReferenceOverlap(references, bounds, out List<string> remaining)) continue;
                    if (remaining.Count == 0) error.Remove();
                    else {
                        error.SequenceOfReferences = new ListValue<StringValue> {
                            InnerText = string.Join(" ", remaining)
                        };
                    }
                }
                if (!errors.Elements<IgnoredError>().Any()) errors.Remove();
            }

            foreach (X14.IgnoredErrors extendedErrors in WorksheetRoot.Descendants<X14.IgnoredErrors>().ToList()) {
                foreach (X14.IgnoredError error in extendedErrors.Elements<X14.IgnoredError>().ToList()) {
                    string references = error.ReferenceSequence?.Text ?? string.Empty;
                    if (!TryRemoveReferenceOverlap(references, bounds, out List<string> remaining)) continue;
                    if (remaining.Count == 0) error.Remove();
                    else error.ReferenceSequence!.Text = string.Join(" ", remaining);
                }
                if (!extendedErrors.Elements<X14.IgnoredError>().Any()) extendedErrors.Remove();
            }
        }

        private void RemoveOffice2010ValidationOverlap((int r1, int c1, int r2, int c2) bounds) {
            foreach (X14.DataValidations validations in WorksheetRoot.Descendants<X14.DataValidations>().ToList()) {
                foreach (X14.DataValidation validation in validations.Elements<X14.DataValidation>().ToList()) {
                    string references = validation.ReferenceSequence?.Text ?? string.Empty;
                    if (!TryRemoveReferenceOverlap(references, bounds, out List<string> remaining)) continue;
                    if (remaining.Count == 0) validation.Remove();
                    else validation.ReferenceSequence!.Text = string.Join(" ", remaining);
                }
                validations.Count = (uint)validations.Elements<X14.DataValidation>().Count();
                if (validations.Count?.Value == 0U) validations.Remove();
            }
        }

        private void RemoveOffice2010ConditionalFormattingOverlap(
            (int r1, int c1, int r2, int c2) bounds) {
            foreach (X14.ConditionalFormattings formattings in
                WorksheetRoot.Descendants<X14.ConditionalFormattings>().ToList()) {
                foreach (X14.ConditionalFormatting formatting in
                    formattings.Elements<X14.ConditionalFormatting>().ToList()) {
                    Xm.ReferenceSequence? target = formatting.GetFirstChild<Xm.ReferenceSequence>();
                    string references = target?.Text ?? string.Empty;
                    if (!TryRemoveReferenceOverlap(references, bounds, out List<string> remaining)) continue;
                    if (remaining.Count == 0) formatting.Remove();
                    else target!.Text = string.Join(" ", remaining);
                }
                if (!formattings.Elements<X14.ConditionalFormatting>().Any()) formattings.Remove();
            }
        }

        private void CleanupEmptyMetadataExtensions() {
            foreach (Extension extension in WorksheetRoot.Descendants<Extension>().ToList()) {
                if (!extension.ChildElements.Any()) extension.Remove();
            }
            foreach (ExtensionList extensions in WorksheetRoot.Elements<ExtensionList>().ToList()) {
                if (!extensions.Elements<Extension>().Any()) extensions.Remove();
            }
        }
    }
}
