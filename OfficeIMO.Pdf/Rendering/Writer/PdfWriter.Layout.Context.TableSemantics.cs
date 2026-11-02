namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        /// <summary>Resolves source cell semantics independently of header styling and repeating row geometry.</summary>
        private static string GetTableCellHeaderScope(PdfTableStyle style, int row, int column, bool headerRow) {
            if (style.CellHeaderScopes != null && style.CellHeaderScopes.TryGetValue((row, column), out var scope)) {
                return scope switch {
                    PdfCanvasTableHeaderScope.Row => "Row",
                    PdfCanvasTableHeaderScope.Column => "Column",
                    PdfCanvasTableHeaderScope.Both => "Both",
                    _ => throw new System.ArgumentOutOfRangeException(nameof(style), "Unsupported table header scope.")
                };
            }
            return headerRow ? "Column" : string.Empty;
        }
    }
}
