namespace OfficeIMO.Pdf;

public partial class PdfTableStyle {
    private System.Collections.Generic.Dictionary<(int Row, int Column), PdfCanvasTableHeaderScope>? _cellHeaderScopes;

    /// <summary>Additional tagged header cells keyed by zero-based table row and column. These scopes affect
    /// PDF structure only; visual header styling and repetition still follow HeaderRowCount and RepeatHeaderRowCount.</summary>
    public System.Collections.Generic.Dictionary<(int Row, int Column), PdfCanvasTableHeaderScope>? CellHeaderScopes {
        get => _cellHeaderScopes;
        set {
            if (value != null) {
                foreach (var item in value) {
                    if (item.Key.Row < 0 || item.Key.Column < 0)
                        throw new System.ArgumentException("Table header cell coordinates cannot be negative.", nameof(CellHeaderScopes));
                    if ((int)item.Value < (int)PdfCanvasTableHeaderScope.Row || (int)item.Value > (int)PdfCanvasTableHeaderScope.Both)
                        throw new System.ArgumentOutOfRangeException(nameof(CellHeaderScopes), "Unsupported table header scope.");
                }
            }
            _cellHeaderScopes = value == null ? null :
                new System.Collections.Generic.Dictionary<(int Row, int Column), PdfCanvasTableHeaderScope>(value);
        }
    }
}
