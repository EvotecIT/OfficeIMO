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
            Text[] sourceTextNodes = source == null ? Array.Empty<Text>() : EnumerateRichTextNodes(source).ToArray();
            if (cell == null || source == null || sourceTextNodes.Length == 0 || !source.Elements<Run>().Any()) {
                return false;
            }

            string[] original = sourceTextNodes.Select(text => text.Text ?? string.Empty).ToArray();
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

            Text[] targetTextNodes = EnumerateRichTextNodes(inline).ToArray();
            if (targetTextNodes.Length != transformed.Count) {
                throw new InvalidOperationException("The Excel rich-text text-node collection changed while its text case was being transformed.");
            }

            for (int index = 0; index < targetTextNodes.Length; index++) {
                targetTextNodes[index].Text = transformed[index];
            }

            cell.CellFormula = null;
            cell.CellValue = null;
            cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.InlineString;
            cell.InlineString = inline;
            _excelDocument.MarkFormulaInputMutation();
            ClearHeaderCache();
            return true;
        }

        private static IEnumerable<Text> EnumerateRichTextNodes(OpenXmlCompositeElement owner) {
            foreach (OpenXmlElement child in owner.ChildElements) {
                if (child is Text text) {
                    yield return text;
                } else if (child is Run run && run.Text != null) {
                    yield return run.Text;
                }
            }
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
