namespace OfficeIMO.Pdf;

/// <summary>
/// Reusable typed PDF content that composes through the canonical <see cref="PdfItemCompose"/> surface.
/// Components own structure and data binding while the document engine continues to own layout and rendering.
/// </summary>
public interface IPdfComponent {
    /// <summary>Composes this component into the supplied document content.</summary>
    void Compose(PdfItemCompose content);
}

/// <summary>
/// Reusable typed PDF content that can react to the live pagination context while still composing
/// through the canonical flow engine.
/// </summary>
/// <remarks>
/// Contextual components can be replayed while pagination stabilizes. Implementations must therefore
/// be deterministic and must not consume one-shot enumerators or mutate external state.
/// </remarks>
public interface IPdfContextComponent {
    /// <summary>Composes this component for the supplied live page context.</summary>
    void Compose(PdfItemCompose content, PdfFlowContext context);
}
