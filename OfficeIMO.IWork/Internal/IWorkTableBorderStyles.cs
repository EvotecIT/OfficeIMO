namespace OfficeIMO.IWork.Internal;

/// <summary>Region and grid defaults queried without allocating the rectangular source grid.</summary>
internal sealed class IWorkTableBorderStyles(int rows, int columns, int headerRows, int headerColumns, int footerRows) {
    internal Dictionary<int, IWorkBorderDeclaration> Strokes { get; } = new();
    internal Dictionary<int, bool?> Visibility { get; } = new();
    internal Dictionary<int, IWorkCellBorderStyle> CellRoles { get; } = new();
    internal bool StructureIsComplete { get; set; } = true;
    internal bool HasDeclarations => !StructureIsComplete || Strokes.Count > 0 || Visibility.Count > 0 || CellRoles.Values.Any(role => role.HasDeclarations);
    internal bool FullyReconstructed => StructureIsComplete
        && Strokes.Values.All(stroke => stroke.IsSupported) && Visibility.Values.All(value => value.HasValue)
        && CellRoles.Values.All(role => role.FullyReconstructed);

    internal bool UsesStroke(int field) => field switch {
        46 => headerRows > 0 && headerRows < rows,
        47 => headerRows > 0,
        48 => headerRows > 1,
        49 => headerRows > 0 && columns > 1,
        50 => headerColumns > 0 && rows > headerRows + footerRows,
        51 => headerColumns > 0 && headerColumns < columns && rows > headerRows + footerRows,
        52 => headerColumns > 0 && rows - headerRows - footerRows > 1,
        53 => headerColumns > 1 && rows > headerRows + footerRows,
        54 => footerRows > 0 && footerRows < rows,
        55 => footerRows > 0,
        56 => footerRows > 1,
        57 => footerRows > 0 && columns > 1,
        58 => columns > headerColumns && rows > headerRows + footerRows,
        59 => columns > headerColumns && rows > headerRows + footerRows,
        60 => rows - headerRows - footerRows > 1 && columns > headerColumns,
        61 => columns - headerColumns > 1 && rows > headerRows + footerRows,
        _ => false
    };

    internal IWorkBorderDeclaration? CellRole(int row, int column, int edge) {
        int role = row <= headerRows ? 19 : row > rows - footerRows ? 21 : column <= headerColumns ? 20 : 18;
        return StructureIsComplete && CellRoles.TryGetValue(role, out var style) ? style[edge] : null;
    }

    internal bool RoleEdgesAreUnambiguous() {
        // Roles are constant within rectangular regions; their boundaries and one interior
        // representative cover every adjacent role pair without allocating or walking the grid.
        var selectedRows = new[] { 1, headerRows, headerRows + 1, rows - footerRows, rows - footerRows + 1, rows }
            .Where(row => row > 0 && row <= rows).Distinct().ToArray();
        var selectedColumns = new[] { 1, headerColumns, headerColumns + 1, columns }
            .Where(column => column > 0 && column <= columns).Distinct().ToArray();
        foreach (int row in selectedRows)
            foreach (int column in selectedColumns) {
                if (row < rows && IWorkTableBorderResolver.Choose(CellRole(row, column, 2), CellRole(row + 1, column, 0))?.IsSupported == false) return false;
                if (column < columns && IWorkTableBorderResolver.Choose(CellRole(row, column, 1), CellRole(row, column + 1, 3))?.IsSupported == false) return false;
            }
        return true;
    }

    internal IWorkBorderDeclaration? Horizontal(int line, int column) {
        if (!StructureIsComplete) return IWorkBorderDeclaration.Invalid;
        if (line == 0) return Stroke(headerRows > 0 ? 47 : column < headerColumns ? 50 : 58);
        if (line == rows) return Stroke(footerRows > 0 ? 55 : column < headerColumns ? 50 : headerRows == rows ? 47 : 58);
        if (headerRows > 0 && line == headerRows) return Stroke(46);
        if (footerRows > 0 && line == rows - footerRows) return Stroke(54);
        if (line < headerRows) return Stroke(48);
        if (line > rows - footerRows) return Stroke(56);
        return Stroke(column < headerColumns ? 52 : 60);
    }

    internal IWorkBorderDeclaration? Vertical(int row, int line) {
        if (!StructureIsComplete) return IWorkBorderDeclaration.Invalid;
        if (line == 0 || line == columns) return Stroke(row < headerRows ? 47 : row >= rows - footerRows ? 55
            : (line == 0 && headerColumns > 0 || line == columns && headerColumns == columns) ? 50 : 59);
        if (row < headerRows) return Stroke(49);
        if (row >= rows - footerRows) return Stroke(57);
        if (headerColumns > 0 && line == headerColumns) return Stroke(51);
        if (line < headerColumns) return Stroke(53);
        return Stroke(61);
    }

    internal IWorkBorderDeclaration? Stroke(int field) {
        int visibleField = VisibilityField(field);
        if (Visibility.TryGetValue(visibleField, out bool? visible)) {
            if (!visible.HasValue) return IWorkBorderDeclaration.Invalid;
            if (!visible.Value) return IWorkBorderDeclaration.None;
        }
        return Strokes.TryGetValue(field, out var stroke) ? stroke : null;
    }

    internal static int VisibilityField(int field) => field switch {
        46 => 35, 51 => 36, 54 => 37,
        47 or 50 => 39, 55 or 58 or 59 => 38,
        49 => 43, 52 => 42, 57 => 44,
        60 => 34, 61 => 33,
        _ => 0
    };
}
