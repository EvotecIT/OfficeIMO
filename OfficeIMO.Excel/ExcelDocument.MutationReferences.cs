using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        internal void RewriteColumnMutationReferences(ExcelSheet editedSheet, int firstColumn, int count, bool deleting) {
            int lastDeletedColumn = deleting ? firstColumn + count - 1 : 0;
            ExcelReference? Transform(ExcelReference reference) =>
                TransformColumnReference(reference, firstColumn, lastDeletedColumn, count, deleting);

            editedSheet.RemapWorksheetAutoFilterColumns(firstColumn, count, deleting);
            int editedSheetIndex = RewriteMutationReferencesAcrossPackage(editedSheet, Transform);
            RewriteCalculationChainColumns(editedSheetIndex, firstColumn, count, deleting);
            RewriteDrawingColumns(editedSheet.WorksheetPart.DrawingsPart?.WorksheetDrawing, firstColumn, count, deleting);
        }

        internal void RewriteMovedRangeReferences(
            ExcelSheet editedSheet,
            ExcelReference source,
            int destinationRow,
            int destinationColumn,
            bool transpose) {
            ExcelReference? Transform(ExcelReference reference) =>
                TransformMovedRangeReference(reference, source, destinationRow, destinationColumn, transpose);
            editedSheet.RemapMovedConnectionParameters(source, destinationRow, destinationColumn, transpose);
            editedSheet.RemapMutationCommentVml(Transform);
            RewriteMutationReferencesAcrossPackage(editedSheet, Transform);
            editedSheet.CleanupCommentArtifacts();
        }

        /// <summary>Maps one parsed A1 reference through a complete range move.</summary>
        internal static ExcelReference TransformMovedRangeReference(
            ExcelReference reference,
            ExcelReference source,
            int destinationRow,
            int destinationColumn,
            bool transpose) {
            source.GetBounds(out int sr1, out int sc1, out int sr2, out int sc2);
            reference.GetBounds(out int rr1, out int rc1, out int rr2, out int rc2);
            if (rr1 < sr1 || rr2 > sr2 || rc1 < sc1 || rc2 > sc2) return reference;
            int MapRow(int row, int column) => transpose ? destinationRow + column - sc1 : destinationRow + row - sr1;
            int MapColumn(int row, int column) => transpose ? destinationColumn + row - sr1 : destinationColumn + column - sc1;
            ExcelReferenceKind kind = transpose
                ? reference.Kind == ExcelReferenceKind.WholeRow ? ExcelReferenceKind.WholeColumn
                    : reference.Kind == ExcelReferenceKind.WholeColumn ? ExcelReferenceKind.WholeRow
                    : reference.Kind
                : reference.Kind;
            return reference.WithCoordinates(
                kind,
                kind == ExcelReferenceKind.WholeColumn ? 0 : MapRow(reference.Start.Row, reference.Start.Column),
                kind == ExcelReferenceKind.WholeRow ? 0 : MapColumn(reference.Start.Row, reference.Start.Column),
                kind == ExcelReferenceKind.WholeColumn ? 0 : MapRow(reference.End.Row, reference.End.Column),
                kind == ExcelReferenceKind.WholeRow ? 0 : MapColumn(reference.End.Row, reference.End.Column),
                transpose ? reference.Start.ColumnAbsolute : null,
                transpose ? reference.Start.RowAbsolute : null,
                transpose ? reference.End.ColumnAbsolute : null,
                transpose ? reference.End.RowAbsolute : null);
        }

        internal void RewriteCellShiftReferences(
            ExcelSheet editedSheet,
            ExcelReference affected,
            ExcelCellShiftDirection direction,
            bool inserting) {
            ExcelReference? Transform(ExcelReference reference) =>
                TransformCellShiftReference(reference, affected, direction, inserting);
            editedSheet.RemapCellShiftConnectionParameters(affected, direction, inserting);
            editedSheet.RemapMutationCommentVml(Transform);
            RewriteMutationReferencesAcrossPackage(editedSheet, Transform);
            editedSheet.CleanupCommentArtifacts();
        }

        /// <summary>Maps one parsed A1 reference through a rectangular cell shift.</summary>
        internal static ExcelReference? TransformCellShiftReference(
            ExcelReference reference,
            ExcelReference affected,
            ExcelCellShiftDirection direction,
            bool inserting) {
            affected.GetBounds(out int ar1, out int ac1, out int ar2, out int ac2);
            int rowCount = ar2 - ar1 + 1;
            int columnCount = ac2 - ac1 + 1;
            int TransformRow(int row, int column) {
                if (column < ac1 || column > ac2) return row;
                if (direction == ExcelCellShiftDirection.Down && row >= ar1) return checked(row + rowCount);
                if (direction == ExcelCellShiftDirection.Up && row > ar2) return row - rowCount;
                return row;
            }
            int TransformColumn(int row, int column) {
                if (row < ar1 || row > ar2) return column;
                if (direction == ExcelCellShiftDirection.Right && column >= ac1) return checked(column + columnCount);
                if (direction == ExcelCellShiftDirection.Left && column > ac2) return column - columnCount;
                return column;
            }
            bool deletingPoint(int row, int column) => !inserting
                && ((direction == ExcelCellShiftDirection.Left && row >= ar1 && row <= ar2 && column >= ac1 && column <= ac2)
                    || (direction == ExcelCellShiftDirection.Up && column >= ac1 && column <= ac2 && row >= ar1 && row <= ar2));
            bool startDeleted = deletingPoint(reference.Start.Row, reference.Start.Column);
            bool endDeleted = deletingPoint(reference.End.Row, reference.End.Column);
            if (startDeleted && endDeleted) return null;
            if (startDeleted || endDeleted) {
                reference.GetBounds(out int rr1, out int rc1, out int rr2, out int rc2);
                if (direction == ExcelCellShiftDirection.Left) {
                    bool keepsLeft = rc1 < ac1;
                    bool keepsRight = rc2 > ac2;
                    if (!keepsLeft && !keepsRight) return null;
                    int newMinimum = keepsLeft ? rc1 : Math.Max(rc1, ac2 + 1) - columnCount;
                    int newMaximum = keepsRight ? rc2 - columnCount : Math.Min(rc2, ac1 - 1);
                    bool reversed = reference.Start.Column > reference.End.Column;
                    return reference.WithCoordinates(
                        reference.Kind,
                        reference.Start.Row,
                        reversed ? newMaximum : newMinimum,
                        reference.End.Row,
                        reversed ? newMinimum : newMaximum);
                }
                bool keepsAbove = rr1 < ar1;
                bool keepsBelow = rr2 > ar2;
                if (!keepsAbove && !keepsBelow) return null;
                int newTop = keepsAbove ? rr1 : Math.Max(rr1, ar2 + 1) - rowCount;
                int newBottom = keepsBelow ? rr2 - rowCount : Math.Min(rr2, ar1 - 1);
                bool rowsReversed = reference.Start.Row > reference.End.Row;
                return reference.WithCoordinates(
                    reference.Kind,
                    rowsReversed ? newBottom : newTop,
                    reference.Start.Column,
                    rowsReversed ? newTop : newBottom,
                    reference.End.Column);
            }
            return reference.WithCoordinates(
                reference.Kind,
                TransformRow(reference.Start.Row, reference.Start.Column),
                TransformColumn(reference.Start.Row, reference.Start.Column),
                TransformRow(reference.End.Row, reference.End.Column),
                TransformColumn(reference.End.Row, reference.End.Column));
        }

        private int RewriteMutationReferencesAcrossPackage(
            ExcelSheet editedSheet,
            Func<ExcelReference, ExcelReference?> transform) {
            Workbook workbook = WorkbookPartRoot.Workbook ?? throw new InvalidOperationException("Workbook root is missing.");
            List<Sheet> sheets = workbook.Sheets?.Elements<Sheet>().ToList() ?? new List<Sheet>();
            int editedSheetIndex = sheets.FindIndex(sheet => string.Equals(sheet.Name?.Value, editedSheet.Name, StringComparison.OrdinalIgnoreCase));
            foreach (DefinedName name in workbook.DefinedNames?.Elements<DefinedName>() ?? Enumerable.Empty<DefinedName>()) {
                bool local = name.LocalSheetId?.Value is uint localIndex && localIndex == (uint)editedSheetIndex;
                name.Text = RewriteMutationFormula(name.Text, editedSheet.Name, local, transform);
            }
            workbook.Save();
            foreach (Sheet sheetElement in sheets) {
                if (sheetElement.Id?.Value is not string relationshipId || WorkbookPartRoot.GetPartById(relationshipId) is not WorksheetPart worksheetPart) continue;
                bool ownerIsEdited = string.Equals(sheetElement.Name?.Value, editedSheet.Name, StringComparison.OrdinalIgnoreCase);
                RewriteMutationFormulaLeaves(worksheetPart.Worksheet, editedSheet.Name, ownerIsEdited, transform);
                RewriteMutationHyperlinks(worksheetPart.Worksheet, editedSheet.Name, ownerIsEdited, transform);
                if (ownerIsEdited) {
                    RewriteMutationAddressAttributes(worksheetPart.Worksheet, transform);
                    WorksheetCommentsPart? commentsPart = worksheetPart.WorksheetCommentsPart;
                    if (commentsPart?.Comments != null) {
                        RewriteMutationAddressAttributes(commentsPart.Comments, transform);
                        commentsPart.Comments.Save();
                    }
                    foreach (WorksheetThreadedCommentsPart threadedPart in worksheetPart.WorksheetThreadedCommentsParts) {
                        if (threadedPart.ThreadedComments == null) continue;
                        RewriteMutationAddressAttributes(threadedPart.ThreadedComments, transform);
                        threadedPart.ThreadedComments.Save();
                    }
                    foreach (NamedSheetViewsPart namedViewsPart in worksheetPart.NamedSheetViewsParts) {
                        RewriteMutationAddressAttributes(namedViewsPart.NamedSheetViews, transform);
                        namedViewsPart.NamedSheetViews?.Save();
                    }
                }
                foreach (TableDefinitionPart tablePart in worksheetPart.TableDefinitionParts) {
                    RewriteMutationFormulaLeaves(tablePart.Table, editedSheet.Name, ownerIsEdited, transform);
                    if (ownerIsEdited) RewriteMutationAddressAttributes(tablePart.Table, transform);
                    tablePart.Table?.Save();
                }
                RewriteMutationDrawingReferences(
                    worksheetPart.DrawingsPart,
                    editedSheet.Name,
                    ownerIsEdited,
                    transform);
                foreach (PivotTablePart pivotPart in worksheetPart.PivotTableParts) {
                    RewriteMutationFormulaLeaves(pivotPart.PivotTableDefinition, editedSheet.Name, ownerIsEdited, transform);
                    if (ownerIsEdited) RewriteMutationAddressAttributes(pivotPart.PivotTableDefinition, transform);
                    pivotPart.PivotTableDefinition?.Save();
                }
                worksheetPart.Worksheet?.Save();
            }
            foreach (ChartsheetPart chartsheetPart in WorkbookPartRoot.ChartsheetParts) {
                RewriteMutationDrawingReferences(
                    chartsheetPart.DrawingsPart,
                    editedSheet.Name,
                    unqualifiedTargetsEdited: false,
                    transform);
            }
            foreach (PivotTableCacheDefinitionPart cachePart in WorkbookPartRoot.PivotTableCacheDefinitionParts) {
                foreach (WorksheetSource source in cachePart.PivotCacheDefinition?.Descendants<WorksheetSource>() ?? Enumerable.Empty<WorksheetSource>()) {
                    if (!string.Equals(source.Sheet?.Value, editedSheet.Name, StringComparison.OrdinalIgnoreCase)
                        || !ExcelReference.TryParse(source.Reference?.Value, out ExcelReference? parsed)) continue;
                    ExcelReference? rewritten = transform(parsed!);
                    source.Reference = rewritten?.ToString() ?? "#REF!";
                }
                cachePart.PivotCacheDefinition?.Save();
            }
            return editedSheetIndex;
        }

        /// <summary>Maps one parsed A1 reference through a complete-column insertion or deletion.</summary>
        internal static ExcelReference? TransformColumnReference(
            ExcelReference reference,
            int firstColumn,
            int lastDeletedColumn,
            int count,
            bool deleting) {
            if (reference.Kind == ExcelReferenceKind.WholeRow) return reference;
            int start = Math.Min(reference.Start.Column, reference.End.Column);
            int end = Math.Max(reference.Start.Column, reference.End.Column);
            int newStart;
            int newEnd;
            if (!deleting) {
                newStart = start >= firstColumn ? checked(start + count) : start;
                newEnd = end >= firstColumn ? checked(end + count) : end;
            } else if (end < firstColumn) {
                return reference;
            } else if (start > lastDeletedColumn) {
                newStart = start - count;
                newEnd = end - count;
            } else {
                bool keepsLeft = start < firstColumn;
                bool keepsRight = end > lastDeletedColumn;
                if (!keepsLeft && !keepsRight) return null;
                newStart = keepsLeft ? start : firstColumn;
                newEnd = keepsRight ? end - count : firstColumn - 1;
                if (newStart > newEnd) return null;
            }
            bool reversed = reference.Start.Column > reference.End.Column;
            int first = reversed ? newEnd : newStart;
            int second = reversed ? newStart : newEnd;
            return reference.WithCoordinates(
                reference.Kind,
                reference.Start.Row,
                first,
                reference.End.Row,
                second);
        }

        private static string RewriteMutationFormula(
            string? formula,
            string editedSheetName,
            bool unqualifiedTargetsEdited,
            Func<ExcelReference, ExcelReference?> transform) {
            if (string.IsNullOrEmpty(formula)) return formula ?? string.Empty;
            return ExcelFormulaSyntaxTree.Parse(formula!).Rewrite(reference =>
                ReferenceTargetsSheet(reference, editedSheetName, unqualifiedTargetsEdited)
                    ? transform(reference)
                    : reference);
        }

        private static void RewriteMutationFormulaLeaves(
            OpenXmlPartRootElement? root,
            string editedSheetName,
            bool unqualifiedTargetsEdited,
            Func<ExcelReference, ExcelReference?> transform) {
            if (root == null) return;
            foreach (OpenXmlLeafTextElement leaf in root.Descendants<OpenXmlLeafTextElement>().Where(IsMutationFormulaLeaf)) {
                leaf.Text = RewriteMutationFormula(leaf.Text, editedSheetName, unqualifiedTargetsEdited, transform);
            }
        }

        private static void RewriteMutationHyperlinks(
            Worksheet? worksheet,
            string editedSheetName,
            bool unqualifiedTargetsEdited,
            Func<ExcelReference, ExcelReference?> transform) {
            if (worksheet == null) return;
            foreach (Hyperlink hyperlink in worksheet.Descendants<Hyperlink>()) {
                if (!string.IsNullOrWhiteSpace(hyperlink.Id?.Value)
                    || string.IsNullOrWhiteSpace(hyperlink.Location?.Value)) continue;
                hyperlink.Location = RewriteMutationFormula(
                    hyperlink.Location!.Value,
                    editedSheetName,
                    unqualifiedTargetsEdited,
                    transform);
            }
        }

        private static void RewriteMutationDrawingReferences(
            DrawingsPart? drawingsPart,
            string editedSheetName,
            bool unqualifiedTargetsEdited,
            Func<ExcelReference, ExcelReference?> transform) {
            if (drawingsPart == null) return;

            bool drawingChanged = false;
            foreach (Xdr.Shape shape in drawingsPart.WorksheetDrawing?.Descendants<Xdr.Shape>()
                ?? Enumerable.Empty<Xdr.Shape>()) {
                if (shape.TextLink?.Value is not string formula || formula.Length == 0) continue;
                string rewritten = RewriteMutationFormula(
                    formula,
                    editedSheetName,
                    unqualifiedTargetsEdited,
                    transform);
                if (string.Equals(formula, rewritten, StringComparison.Ordinal)) continue;
                shape.TextLink = rewritten;
                drawingChanged = true;
            }
            if (drawingChanged) drawingsPart.WorksheetDrawing?.Save();

            foreach (ChartPart chartPart in drawingsPart.ChartParts) {
                RewriteMutationChartRoot(
                    chartPart.ChartSpace,
                    editedSheetName,
                    unqualifiedTargetsEdited,
                    transform);
            }
            foreach (ExtendedChartPart chartPart in drawingsPart.ExtendedChartParts) {
                RewriteMutationChartRoot(
                    chartPart.ChartSpace,
                    editedSheetName,
                    unqualifiedTargetsEdited,
                    transform);
            }
        }

        private static void RewriteMutationChartRoot(
            OpenXmlPartRootElement? chartRoot,
            string editedSheetName,
            bool unqualifiedTargetsEdited,
            Func<ExcelReference, ExcelReference?> transform) {
            if (chartRoot == null) return;

            bool changed = false;
            foreach (OpenXmlLeafTextElement formula in chartRoot.Descendants<OpenXmlLeafTextElement>()
                .Where(IsMutationFormulaLeaf)) {
                string original = formula.Text ?? string.Empty;
                string rewritten = RewriteMutationFormula(
                    original,
                    editedSheetName,
                    unqualifiedTargetsEdited,
                    transform);
                if (string.Equals(original, rewritten, StringComparison.Ordinal)) continue;
                formula.Text = rewritten;
                ExcelSheet.InvalidateChartFormulaCache(formula);
                changed = true;
            }
            if (changed) chartRoot.Save();
        }

        private static bool IsMutationFormulaLeaf(OpenXmlLeafTextElement leaf) =>
            string.Equals(leaf.LocalName, "f", StringComparison.OrdinalIgnoreCase)
            || leaf.LocalName.IndexOf("formula", StringComparison.OrdinalIgnoreCase) >= 0;

        private static bool ReferenceTargetsSheet(ExcelReference reference, string editedSheetName, bool unqualifiedTargetsEdited) {
            if (!reference.IsQualified) return unqualifiedTargetsEdited;
            string qualifier = reference.Qualifier!;
            if (qualifier.StartsWith("[", StringComparison.Ordinal)) return false;
            if (qualifier.Length >= 2 && qualifier[0] == '\'' && qualifier[qualifier.Length - 1] == '\'') {
                qualifier = qualifier.Substring(1, qualifier.Length - 2).Replace("''", "'");
            }
            int colon = qualifier.IndexOf(':');
            if (colon >= 0) {
                return string.Equals(qualifier.Substring(0, colon), editedSheetName, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(qualifier.Substring(colon + 1), editedSheetName, StringComparison.OrdinalIgnoreCase);
            }
            return string.Equals(qualifier, editedSheetName, StringComparison.OrdinalIgnoreCase);
        }

        private static void RewriteMutationAddressAttributes(
            OpenXmlPartRootElement? root,
            Func<ExcelReference, ExcelReference?> transform) {
            if (root == null) return;
            foreach (OpenXmlElement element in root.Descendants().Prepend(root).ToList()) {
                bool removeElement = false;
                foreach (OpenXmlAttribute attribute in element.GetAttributes().Where(item =>
                    string.Equals(item.LocalName, "ref", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(item.LocalName, "sqref", StringComparison.OrdinalIgnoreCase)).ToArray()) {
                    string original = attribute.Value ?? string.Empty;
                    string rewritten = RewriteReferenceList(original, transform);
                    if (string.Equals(rewritten, original, StringComparison.Ordinal)) continue;
                    if (rewritten.Length == 0) {
                        if (element.Parent == null) {
                            throw new InvalidOperationException($"Structural edit would remove the required '{attribute.LocalName}' reference from package root '{element.LocalName}'.");
                        }
                        removeElement = true;
                        break;
                    }
                    element.SetAttribute(new OpenXmlAttribute(attribute.Prefix, attribute.LocalName, attribute.NamespaceUri, rewritten));
                }
                if (removeElement) element.Remove();
            }
            foreach (OpenXmlElement container in root.Descendants().Where(element =>
                string.Equals(element.LocalName, "mergeCells", StringComparison.OrdinalIgnoreCase)
                || string.Equals(element.LocalName, "dataValidations", StringComparison.OrdinalIgnoreCase)
                || string.Equals(element.LocalName, "ignoredErrors", StringComparison.OrdinalIgnoreCase)
                || string.Equals(element.LocalName, "protectedRanges", StringComparison.OrdinalIgnoreCase)).ToList()) {
                if (!container.ChildElements.Any()) {
                    container.Remove();
                    continue;
                }
                OpenXmlAttribute? count = container.GetAttributes()
                    .FirstOrDefault(attribute => string.Equals(attribute.LocalName, "count", StringComparison.OrdinalIgnoreCase));
                if (count.HasValue && !string.IsNullOrEmpty(count.Value.LocalName)) {
                    container.SetAttribute(new OpenXmlAttribute(
                        count.Value.Prefix,
                        count.Value.LocalName,
                        count.Value.NamespaceUri,
                        container.ChildElements.Count.ToString(CultureInfo.InvariantCulture)));
                }
            }
        }

        private static string RewriteReferenceList(string value, Func<ExcelReference, ExcelReference?> transform) {
            string[] items = value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
            if (items.Length == 0) return value;
            var rewritten = new List<string>(items.Length);
            foreach (string item in items) {
                if (!ExcelReference.TryParse(item, out ExcelReference? reference)) {
                    rewritten.Add(item);
                    continue;
                }
                ExcelReference? mapped = transform(reference!);
                if (mapped != null) rewritten.Add(mapped.ToString());
            }
            return string.Join(" ", rewritten);
        }

        private void RewriteCalculationChainColumns(int editedSheetIndex, int firstColumn, int count, bool deleting) {
            CalculationChain? chain = WorkbookPartRoot.CalculationChainPart?.CalculationChain;
            if (chain == null) return;
            uint currentSheet = 0;
            foreach (CalculationCell cell in chain.Elements<CalculationCell>().ToList()) {
                if (cell.SheetId?.Value is int sheetId) currentSheet = (uint)sheetId;
                if (currentSheet != (uint)editedSheetIndex || !ExcelReference.TryParse(cell.CellReference?.Value, out ExcelReference? reference)) continue;
                ExcelReference? mapped = TransformColumnReference(reference!, firstColumn, firstColumn + count - 1, count, deleting);
                if (mapped == null) cell.Remove();
                else cell.CellReference = mapped.ToString();
            }
            chain.Save();
        }

        private static void RewriteDrawingColumns(Xdr.WorksheetDrawing? drawing, int firstColumn, int count, bool deleting) {
            if (drawing == null) return;
            int firstZeroBased = firstColumn - 1;
            foreach (Xdr.ColumnId column in drawing.Descendants<Xdr.ColumnId>()) {
                if (!int.TryParse(column.Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int value)) continue;
                int mapped = value;
                if (!deleting && value >= firstZeroBased) mapped = checked(value + count);
                else if (deleting && value >= firstZeroBased + count) mapped = value - count;
                else if (deleting && value >= firstZeroBased) mapped = firstZeroBased;
                column.Text = mapped.ToString(CultureInfo.InvariantCulture);
            }
            drawing.Save();
        }
    }
}
