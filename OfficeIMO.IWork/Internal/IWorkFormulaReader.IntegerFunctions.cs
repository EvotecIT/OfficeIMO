using System.Globalization;
using System.Threading;

namespace OfficeIMO.IWork.Internal;

internal static partial class IWorkFormulaReader {
    /// <summary>
    /// Gates integer reconstruction to bounded literals, numbers, proven blanks
    /// and canonical integer text in native single-cell references. All-blank
    /// arguments, broader coercion and formula dependencies remain unqualified.
    /// </summary>
    private static bool HasQualifiedIntegerOperands(string name, IReadOnlyList<Operand> arguments,
        IWorkProjectionBudget? budget, CancellationToken cancellationToken, ref int remainingReferenceCells) {
        ulong productBound = 1;
        bool productExceedsBound = false, hasZero = false, hasNumericOperand = false;
        foreach (Operand argument in arguments) {
            if (argument.IntegerReference is not IntegerReference reference) {
                if (!double.TryParse(argument.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out double number)
                    || !AccumulateInteger(number, name, ref productBound, ref productExceedsBound, ref hasZero)) return false;
                hasNumericOperand = true;
                continue;
            }
            IWorkTable? table = reference.Table;
            if (table == null || budget == null || table.MergedRanges.Count != 0
                || reference.LastColumn >= table.ColumnCount || reference.LastRow >= table.RowCount) return false;
            long count = ((long)reference.LastColumn - reference.FirstColumn + 1)
                * ((long)reference.LastRow - reference.FirstRow + 1);
            if (count > remainingReferenceCells) return false;
            remainingReferenceCells -= (int)count;
            budget.AddFormulaRenderingOperations(count);
            for (int row = reference.FirstRow; row <= reference.LastRow; row++) {
                for (int column = reference.FirstColumn; column <= reference.LastColumn; column++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    IWorkTableCell? cell = table.GetCell(row + 1, column + 1);
                    if (cell != null && (cell.SourceFormulaIsDeclared || !cell.CachedValueIsComplete
                        || cell.HasDecodeError || cell.NumericValueIsApproximate)) return false;
                    if (cell == null || cell.Kind == IWorkCellKind.Empty) {
                        if (!table.BodyMetadataIsComplete || !table.CellStorageIsComplete
                            || cell != null && (cell.ValueKind != IWorkCellKind.Empty || cell.Value != null
                                || cell.RichText != null)) return false;
                        // Native LCM ignores a direct blank but treats a blank
                        // within a range as zero. GCD's zero identity is unchanged.
                        if (!reference.IsSingleCellSyntax && name == "LCM") hasZero = true;
                        continue;
                    }
                    double number;
                    if (cell.Kind == IWorkCellKind.Number && cell.ValueKind == IWorkCellKind.Number
                        && cell.Value is double scalar) number = scalar;
                    else if (reference.IsSingleCellSyntax && cell.Kind == IWorkCellKind.Text
                        && cell.ValueKind == IWorkCellKind.Text && cell.Value is string text
                        && TryCanonicalIntegerText(text, out ulong integer)) number = integer;
                    else return false;
                    if (!AccumulateInteger(number, name, ref productBound, ref productExceedsBound, ref hasZero)) return false;
                    hasNumericOperand = true;
                }
            }
        }
        return hasNumericOperand && (hasZero || !productExceedsBound);
    }

    private const ulong MaximumNativeInteger = 999999999999999UL;

    private static bool TryCanonicalIntegerText(string text, out ulong value) {
        value = 0;
        // Deliberately exclude whitespace, signs, leading zeros, fractions and
        // localized text. A range must not inherit single-cell coercion rules.
        return text.Length is > 0 and <= 15
            && ulong.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out value)
            && value <= MaximumNativeInteger && text == value.ToString(CultureInfo.InvariantCulture);
    }

    private static bool AccumulateInteger(double number, string name, ref ulong productBound,
        ref bool productExceedsBound, ref bool hasZero) {
        if (double.IsNaN(number) || double.IsInfinity(number) || number < 0 || number > MaximumNativeInteger) return false;
        if (name != "LCM") return true;
        ulong integer = (ulong)Math.Truncate(number);
        if (integer == 0) hasZero = true;
        else if (!productExceedsBound) {
            if (integer > MaximumNativeInteger / productBound) productExceedsBound = true;
            else productBound *= integer;
        }
        return true;
    }

    // Coordinates come from the canonical native reference binder, never from
    // parsing display labels. A null table requests the later Numbers binding pass.
    private sealed class IntegerReference {
        internal IntegerReference(IWorkTable? table, int firstColumn, int lastColumn, int firstRow, int lastRow,
            bool singleCellSyntax = false) {
            Table = table; FirstColumn = firstColumn; LastColumn = lastColumn; FirstRow = firstRow; LastRow = lastRow;
            IsSingleCellSyntax = singleCellSyntax && firstColumn == lastColumn && firstRow == lastRow;
        }
        internal IWorkTable? Table { get; }
        internal int FirstColumn { get; }
        internal int LastColumn { get; }
        internal int FirstRow { get; }
        internal int LastRow { get; }
        internal bool IsSingleCellSyntax { get; }
    }

    private static IntegerReference? CombineIntegerReferences(Operand first, Operand last) {
        IntegerReference? a = first.IntegerReference, b = last.IntegerReference;
        return a != null && b != null && a.Table == b.Table
            && a.FirstColumn == a.LastColumn && a.FirstRow == a.LastRow
            && b.FirstColumn == b.LastColumn && b.FirstRow == b.LastRow
            && a.FirstColumn <= b.FirstColumn && a.FirstRow <= b.FirstRow
            ? new IntegerReference(a.Table, a.FirstColumn, b.LastColumn, a.FirstRow, b.LastRow) : null;
    }
}
