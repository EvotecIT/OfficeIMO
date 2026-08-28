using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Changes stored text casing in a cell while preserving cell or rich-run formatting.
        /// Formulas and non-text values are left unchanged.
        /// </summary>
        /// <returns><see langword="true"/> when text was transformed; otherwise <see langword="false"/>.</returns>
        public bool TransformCellTextCase(int row, int column, OfficeIMO.Drawing.OfficeTextCase textCase, CultureInfo? culture = null) {
            if (!TryGetCellValueSnapshot(row, column, out ExcelCellValueSnapshot? snapshot) ||
                snapshot == null ||
                snapshot.Kind != ExcelCellValueKind.Text) {
                return false;
            }

            bool transformedRichText = false;
            WriteLockConditional(() => transformedRichText = TransformRichTextCaseCore(row, column, textCase, culture));
            if (transformedRichText) {
                return true;
            }

            CellValue(row, column, OfficeIMO.Drawing.OfficeTextCaseTransformer.Apply(snapshot.Text, textCase, culture));
            return true;
        }

        private bool TransformRichTextCaseCore(
            int row,
            int column,
            OfficeIMO.Drawing.OfficeTextCase textCase,
            CultureInfo? culture) {
            Cell? cell = TryGetExistingCell(row, column);
            OpenXmlCompositeElement? source = ResolveRichTextOwner(cell);
            Run[] sourceRuns = source?.Elements<Run>().ToArray() ?? Array.Empty<Run>();
            if (cell == null || source == null || sourceRuns.Length == 0) {
                return false;
            }

            string[] original = sourceRuns.Select(run => run.Text?.Text ?? string.Empty).ToArray();
            IReadOnlyList<string> transformed = OfficeIMO.Drawing.OfficeTextCaseTransformer.ApplySegments(original, textCase, culture);
            if (original.SequenceEqual(transformed, StringComparer.Ordinal)) {
                return true;
            }

            InlineString inline;
            if (source is InlineString existingInline) {
                inline = existingInline;
            } else {
                inline = new InlineString();
                foreach (OpenXmlElement child in source.ChildElements) {
                    inline.Append(child.CloneNode(true));
                }
            }

            Run[] targetRuns = inline.Elements<Run>().ToArray();
            if (targetRuns.Length != transformed.Count) {
                throw new InvalidOperationException("The Excel rich-text run collection changed while its text case was being transformed.");
            }

            for (int index = 0; index < targetRuns.Length; index++) {
                Text? text = targetRuns[index].Text;
                if (text != null) {
                    text.Text = transformed[index];
                } else if (transformed[index].Length > 0) {
                    targetRuns[index].Append(new Text(transformed[index]) { Space = SpaceProcessingModeValues.Preserve });
                }
            }

            cell.CellFormula = null;
            cell.CellValue = null;
            cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.InlineString;
            cell.InlineString = inline;
            _excelDocument.MarkFormulaInputMutation();
            ClearHeaderCache();
            return true;
        }

        private OpenXmlCompositeElement? ResolveRichTextOwner(Cell? cell) {
            if (cell?.InlineString != null) {
                return cell.InlineString;
            }

            if (cell?.DataType?.Value != DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString
                || !int.TryParse(cell.CellValue?.InnerText, NumberStyles.None, CultureInfo.InvariantCulture, out int sharedStringIndex)
                || sharedStringIndex < 0) {
                return null;
            }

            return _spreadSheetDocument.WorkbookPart?
                .SharedStringTablePart?
                .SharedStringTable?
                .Elements<SharedStringItem>()
                .ElementAtOrDefault(sharedStringIndex);
        }
    }
}
