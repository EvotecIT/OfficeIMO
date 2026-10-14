using System.Threading;
using OfficeIMO.IWork.Internal;

namespace OfficeIMO.IWork;

public sealed partial class IWorkNativeEditSession {
    /// <summary>Replaces an existing numeric Numbers cell without changing its native type, styles or identities.</summary>
    /// <remarks>Coordinates are one-based. Only complete modern sources without selected formulas are admitted.
    /// Numeric cells must use Decimal128 storage and values must be finite. Assigning the original projected value restores its original scalar bytes, including
    /// a higher-precision source representation. Edited saves require explicit retained visual-cache acceptance.</remarks>
    public void SetNumbersCellNumber(IWorkObjectIdentity tableIdentity, int row, int column, double value,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (double.IsNaN(value) || double.IsInfinity(value)) throw new ArgumentOutOfRangeException(nameof(value));
        UpdateNumbersCell(tableIdentity, row, column, IWorkCellKind.Number, value, cancellationToken);
    }

    /// <summary>Replaces an existing boolean Numbers cell while preserving its native type, styles and identities.</summary>
    /// <remarks>Coordinates are one-based. Complete modern sources without selected formulas are admitted;
    /// insertion, type changes and nested-index editing are unsupported. Edited saves require explicit cache acceptance.</remarks>
    public void SetNumbersCellBoolean(IWorkObjectIdentity tableIdentity, int row, int column, bool value,
        CancellationToken cancellationToken = default) =>
        UpdateNumbersCell(tableIdentity, row, column, IWorkCellKind.Boolean, value, cancellationToken);

    private void UpdateNumbersCell(IWorkObjectIdentity identity, int row, int column, IWorkCellKind kind,
        object value, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        if (identity == null) throw new ArgumentNullException(nameof(identity));
        if (row < 1) throw new ArgumentOutOfRangeException(nameof(row));
        if (column < 1) throw new ArgumentOutOfRangeException(nameof(column));
        if (_source.Kind != IWorkDocumentKind.Numbers) throw new InvalidOperationException("Cell editing requires a Numbers source.");
        if (_source.ContainerEntries.Any(entry => entry.Path == "Index.zip"))
            throw new NotSupportedException("Edits to nested Index.zip sources are not supported.");
        var view = _source.WithCancellation(token); var projection = view.ReadNumbers();
        var tables = projection.Sheets.SelectMany(sheet => sheet.Tables).ToArray();
        if (!projection.HasEditableContent || projection.SourceDeclarationIssues.Count != 0 || projection.SourceReferenceIssues.Count != 0
            || tables.SelectMany(table => table.Cells).Any(cell => cell.SourceFormulaIsDeclared || cell.Formula != null
                || cell.HasDecodeError || !cell.CachedValueIsComplete || cell.UnsupportedFeatures != IWorkCellUnsupportedFeatures.None))
            throw new NotSupportedException("Scalar editing requires complete Numbers content without selected formulas or unsupported cell features.");
        var matches = tables.Where(table => SameIdentity(table.SourceIdentity, identity)).Take(2).ToArray();
        if (matches.Length != 1) throw new ArgumentException("The identity must resolve to exactly one table in this source.", nameof(identity));
        var table = matches[0];
        if (row > table.RowCount) throw new ArgumentOutOfRangeException(nameof(row));
        if (column > table.ColumnCount) throw new ArgumentOutOfRangeException(nameof(column));
        var cell = table.GetCell(row, column);
        if (cell == null || cell.Kind != kind || cell.ValueKind != kind)
            throw new NotSupportedException("Only existing scalar cells of the requested native kind can be replaced.");
        var slot = IWorkTableReader.LocateNativeScalar(view, table, tables, row, column, kind);
        byte[] scalar;
        if (value.Equals(cell.Value)) scalar = slot.Original;
        else if (kind == IWorkCellKind.Number) scalar = IWorkDecimal128Writer.Encode((double)value);
        else {
            scalar = BitConverter.GetBytes((bool)value ? 1d : 0d);
            if (!BitConverter.IsLittleEndian) Array.Reverse(scalar);
        }
        byte[] current = _payloads.TryGetValue(slot.Record, out byte[]? pending) ? pending : slot.Record.Payload;
        byte[] updated = (byte[])current.Clone();
        IWorkPreservingWireEdit.ReplaceBytesRange(updated, new[] { (5, slot.RowOccurrence), (6, -1) },
            slot.ValueOffset, scalar, view.Options, token);
        token.ThrowIfCancellationRequested();
        if (updated.SequenceEqual(slot.Record.Payload)) _payloads.Remove(slot.Record); else _payloads[slot.Record] = updated;
    }
}
