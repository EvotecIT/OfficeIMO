namespace OfficeIMO.Pdf;

/// <summary>
/// Reusable typed PDF content that composes through the canonical <see cref="PdfItemCompose"/> surface.
/// Components own structure and data binding while the document engine continues to own layout and rendering.
/// </summary>
public interface IPdfComponent {
    /// <summary>Composes this component into the supplied document content.</summary>
    void Compose(PdfItemCompose content);
}
