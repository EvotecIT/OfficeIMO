using OfficeIMO.Drawing;
using System.Threading;

namespace OfficeIMO.IWork.Internal;

/// <summary>Resolves one physical shared edge from sparse explicit edits, cell properties and table defaults.</summary>
internal sealed class IWorkTableBorderResolver(IWorkTableBorderStyles defaults, IWorkTableBorderOverrides overrides) {
    internal bool FullyReconstructed => defaults.FullyReconstructed && overrides.FullyReconstructed;
    internal bool HasDeclarations => defaults.HasDeclarations || overrides.HasDeclarations;

    internal OfficeBorderBox GetBorders(int row, int column, IWorkTable table) {
        return new OfficeBorderBox(Vertical(row, column - 1), Horizontal(row - 1, column), Vertical(row, column), Horizontal(row, column));

        OfficeBorderSide? Horizontal(int line, int position) {
            var above = line > 0 ? table.GetCell(line, position)?.BorderStyle?[2] : null;
            var below = line < table.RowCount ? table.GetCell(line + 1, position)?.BorderStyle?[0] : null;
            var role = Choose(line > 0 ? defaults.CellRole(line, position, 2) : null,
                line < table.RowCount ? defaults.CellRole(line + 1, position, 0) : null);
            return Resolve(true, line, position - 1, above, below, role ?? defaults.Horizontal(line, position - 1));
        }

        OfficeBorderSide? Vertical(int position, int line) {
            var left = line > 0 ? table.GetCell(position, line)?.BorderStyle?[1] : null;
            var right = line < table.ColumnCount ? table.GetCell(position, line + 1)?.BorderStyle?[3] : null;
            var role = Choose(line > 0 ? defaults.CellRole(position, line, 1) : null,
                line < table.ColumnCount ? defaults.CellRole(position, line + 1, 3) : null);
            return Resolve(false, line, position - 1, left, right, role ?? defaults.Vertical(position - 1, line));
        }
    }

    private OfficeBorderSide? Resolve(bool horizontal, int line, int position,
        IWorkBorderDeclaration? first, IWorkBorderDeclaration? second, IWorkBorderDeclaration? fallback) =>
        (overrides.Get(horizontal, line, position) ?? Choose(first, second) ?? fallback)?.Side;

    // Cell-style edges have no native edit order. A competing pair must not silently choose one side.
    internal bool CellEdgesAreUnambiguous(IWorkTable table, CancellationToken token) {
        foreach (var cell in table.Cells) {
            token.ThrowIfCancellationRequested();
            if (cell.BorderStyle?.HasDeclarations != true) continue;
            if (cell.Row < table.RowCount && overrides.Get(true, cell.Row, cell.Column - 1) == null
                && Choose(cell.BorderStyle[2], table.GetCell(cell.Row + 1, cell.Column)?.BorderStyle?[0])?.IsSupported == false) return false;
            if (cell.Row > 1 && overrides.Get(true, cell.Row - 1, cell.Column - 1) == null
                && Choose(table.GetCell(cell.Row - 1, cell.Column)?.BorderStyle?[2], cell.BorderStyle[0])?.IsSupported == false) return false;
            if (cell.Column < table.ColumnCount && overrides.Get(false, cell.Column, cell.Row - 1) == null
                && Choose(cell.BorderStyle[1], table.GetCell(cell.Row, cell.Column + 1)?.BorderStyle?[3])?.IsSupported == false) return false;
            if (cell.Column > 1 && overrides.Get(false, cell.Column - 1, cell.Row - 1) == null
                && Choose(table.GetCell(cell.Row, cell.Column - 1)?.BorderStyle?[1], cell.BorderStyle[3])?.IsSupported == false) return false;
        }
        return true;
    }

    internal static IWorkBorderDeclaration? Choose(IWorkBorderDeclaration? first, IWorkBorderDeclaration? second) =>
        first == null ? second : second == null ? first
        : first.IsSupported && second.IsSupported && Nullable.Equals(first.Side, second.Side) ? first : IWorkBorderDeclaration.Invalid;

}
