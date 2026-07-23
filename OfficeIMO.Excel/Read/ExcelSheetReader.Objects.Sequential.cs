using System.Globalization;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using System.Threading;
using System.Text;
using System.ComponentModel;
using System.Linq.Expressions;
using System.Runtime.Serialization;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Object-mapping readers for <see cref="ExcelSheetReader"/>.
    /// </summary>
    public sealed partial class ExcelSheetReader {
        private IEnumerable<T> ReadObjectsStreamIterator<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
            string a1Range,
            int r1,
            int c1,
            int r2,
            int c2,
            int cols,
            CancellationToken ct) where T : new() {
            bool canCancel = ct.CanBeCanceled;
            if (canCancel) {
                ct.ThrowIfCancellationRequested();
            }

            TypedPropertyBinding<T>?[]? bindings = null;
            Dictionary<int, Row>? pendingRows = null;
            int nextDataRow = r1 + 1;
            int convertedCells = 0;

            foreach (var row in EnumerateWorksheetRows(ct)) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                int rowIndex = checked((int)row.RowIndex!.Value);
                if (rowIndex < r1) {
                    continue;
                }

                if (rowIndex > r2) {
                    continue;
                }

                if (bindings == null) {
                    if (rowIndex == r1) {
                        bindings = CreateTypedHeaderBindingsFromRow<T>(row, a1Range, c1, c2, cols);

                        while (pendingRows != null && pendingRows.TryGetValue(nextDataRow, out var pendingRow)) {
                            pendingRows.Remove(nextDataRow);
                            var pendingTarget = new T();
                            FillTypedObjectFromRow(pendingRow, c1, c2, bindings, pendingTarget, ct, ref convertedCells);
                            yield return pendingTarget;
                            nextDataRow++;
                        }

                        continue;
                    }

                    pendingRows ??= new Dictionary<int, Row>();
                    AddPendingTypedRow(pendingRows, rowIndex, row);
                    continue;
                }

                if (rowIndex <= r1) {
                    continue;
                }

                if (rowIndex < nextDataRow) {
                    continue;
                }

                if (rowIndex > nextDataRow) {
                    pendingRows ??= new Dictionary<int, Row>();
                    AddPendingTypedRow(pendingRows, rowIndex, row);
                    continue;
                }

                var currentRow = row;
                while (true) {
                    var target = new T();
                    FillTypedObjectFromRow(currentRow, c1, c2, bindings, target, ct, ref convertedCells);
                    yield return target;
                    nextDataRow++;

                    if (pendingRows == null || !pendingRows.TryGetValue(nextDataRow, out currentRow)) {
                        break;
                    }

                    pendingRows.Remove(nextDataRow);
                }
            }

            bindings ??= CreateTypedHeaderBindingsFromMissingRow<T>(a1Range, cols);
            while (nextDataRow <= r2) {
                if (canCancel && ((nextDataRow - r1) & 1023) == 0) {
                    ct.ThrowIfCancellationRequested();
                }

                if (pendingRows != null && pendingRows.TryGetValue(nextDataRow, out var pendingRow)) {
                    pendingRows.Remove(nextDataRow);
                    var target = new T();
                    FillTypedObjectFromRow(pendingRow, c1, c2, bindings, target, ct, ref convertedCells);
                    yield return target;
                } else {
                    yield return new T();
                }

                nextDataRow++;
            }
        }

        private bool TryReadObjectsSequentialSinglePass<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
            string a1Range,
            int r1,
            int c1,
            int r2,
            int c2,
            int rows,
            int cols,
            CancellationToken ct,
            out List<T> result) where T : new() {
            int dataRowCount = rows - 1;
            result = new List<T>(dataRowCount);
            for (int r = 0; r < dataRowCount; r++) {
                result.Add(new T());
            }

            bool canCancel = ct.CanBeCanceled;
            if (canCancel) {
                ct.ThrowIfCancellationRequested();
            }

            TypedPropertyBinding<T>?[]? bindings = null;
            int convertedCells = 0;

            foreach (var row in EnumerateWorksheetRows(ct)) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                int rowIndex = checked((int)row.RowIndex!.Value);
                if (rowIndex < r1) {
                    continue;
                }

                if (rowIndex > r2) {
                    continue;
                }

                if (rowIndex == r1) {
                    bindings = CreateTypedHeaderBindingsFromRow<T>(row, a1Range, c1, c2, cols);
                    continue;
                }

                if (bindings == null) {
                    return false;
                }

                int resultIndex = rowIndex - r1 - 1;
                if ((uint)resultIndex >= (uint)result.Count) {
                    continue;
                }

                FillTypedObjectFromRow(row, c1, c2, bindings, result[resultIndex], ct, ref convertedCells);
            }

            return bindings != null;
        }

        private IEnumerable<T> ReadObjectsSequential<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
            string a1Range,
            int r1,
            int c1,
            int r2,
            int c2,
            int rows,
            int cols,
            CancellationToken ct) where T : new() {
            bool canCancel = ct.CanBeCanceled;
            if (canCancel) {
                ct.ThrowIfCancellationRequested();
            }

            TypedPropertyBinding<T>?[]? bindings = null;
            foreach (var row in EnumerateWorksheetRows(ct)) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                int rowIndex = checked((int)row.RowIndex!.Value);
                if (rowIndex != r1) {
                    continue;
                }

                bindings = CreateTypedHeaderBindingsFromRow<T>(row, a1Range, c1, c2, cols);
                break;
            }

            bindings ??= CreateTypedHeaderBindingsFromMissingRow<T>(a1Range, cols);

            int dataRowCount = rows - 1;
            var result = new List<T>(dataRowCount);
            for (int r = 0; r < dataRowCount; r++) {
                if (canCancel && (r & 1023) == 0) {
                    ct.ThrowIfCancellationRequested();
                }

                result.Add(new T());
            }

            int convertedCells = 0;
            foreach (var row in EnumerateWorksheetRows(ct)) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                int rowIndex = checked((int)row.RowIndex!.Value);
                if (rowIndex <= r1) {
                    continue;
                }

                if (rowIndex > r2) {
                    continue;
                }

                int resultIndex = rowIndex - r1 - 1;
                if ((uint)resultIndex >= (uint)result.Count) {
                    continue;
                }

                FillTypedObjectFromRow(row, c1, c2, bindings, result[resultIndex], ct, ref convertedCells);
            }

            return result;
        }

        private TypedPropertyBinding<T>?[] CreateTypedHeaderBindingsFromRow<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
            DocumentFormat.OpenXml.Spreadsheet.Row row,
            string a1Range,
            int c1,
            int c2,
            int cols) where T : new() {
            var headerValues = new object?[cols];
            foreach (var cell in row.Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>()) {
                int columnIndex = A1.ParseColumnIndexFromCellReferenceFast(cell.CellReference?.Value);
                if (columnIndex < c1 || columnIndex > c2) {
                    continue;
                }

                if (TryConvertCell(cell, out object? value)) {
                    headerValues[columnIndex - c1] = value;
                }
            }

            var headers = ExcelHeaderNameHelper.BuildUniqueHeaders(cols, c => headerValues[c]?.ToString(), _opt.NormalizeHeaders);
            return GetTypedHeaderBindings<T>(headers, a1Range).Bindings;
        }

        private TypedPropertyBinding<T>?[] CreateTypedHeaderBindingsFromMissingRow<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
            string a1Range,
            int cols) where T : new() {
            var headers = ExcelHeaderNameHelper.BuildUniqueHeaders(cols, static _ => null, _opt.NormalizeHeaders);
            return GetTypedHeaderBindings<T>(headers, a1Range).Bindings;
        }

        private void FillTypedObjectFromRow<T>(
            DocumentFormat.OpenXml.Spreadsheet.Row row,
            int c1,
            int c2,
            TypedPropertyBinding<T>?[] bindings,
            T target,
            CancellationToken ct,
            ref int convertedCells) {
            bool canCancel = ct.CanBeCanceled;
            foreach (var cell in row.Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>()) {
                if (canCancel && (++convertedCells & 1023) == 0) {
                    ct.ThrowIfCancellationRequested();
                }

                int columnIndex = A1.ParseColumnIndexFromCellReferenceFast(cell.CellReference?.Value);
                if (columnIndex < c1 || columnIndex > c2) {
                    continue;
                }

                var binding = bindings[columnIndex - c1];
                if (binding == null) {
                    continue;
                }

                bool convertedSuccessfully = TryConvertCellForBinding(cell, binding, out object? converted);
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                if (convertedSuccessfully) {
                    binding.SetValue(target, converted);
                }
            }
        }

        private TypedHeaderBindingCache<T> GetTypedHeaderBindings<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
            string[] headers,
            string a1Range) where T : new() {
            var propertyMaps = TypedObjectBindingCache<T>.PropertyMaps;
            string typeName = propertyMaps.TypeName;
            foreach (string diagnostic in propertyMaps.Diagnostics) {
                _opt.Execution.ReportInfo(diagnostic);
            }

            var headerBindings = TypedObjectBindingCache<T>.GetHeaderBindings(headers);

            if (_opt.StrictTypedMapping) {
                var strictIssues = new List<string>(propertyMaps.Diagnostics);
                strictIssues.AddRange(headerBindings.UnmappedIssues);
                if (strictIssues.Count > 0) {
                    throw new InvalidOperationException(
                        $"Typed mapping for '{typeName}' is strict and could not resolve all headers in range '{a1Range}'. " +
                        string.Join(" ", strictIssues));
                }
            }

            return headerBindings;
        }
    }
}
