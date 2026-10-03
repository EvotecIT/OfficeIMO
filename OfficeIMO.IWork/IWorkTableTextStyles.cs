namespace OfficeIMO.IWork;

/// <summary>Qualified native paragraph defaults for the four table regions. These do not materialize a dense cell grid.</summary>
public sealed class IWorkTableTextStyles {
    internal IWorkTableTextStyles(IWorkParagraphStyle? body, IWorkParagraphStyle? headerRow,
        IWorkParagraphStyle? headerColumn, IWorkParagraphStyle? footerRow) {
        Body = body; HeaderRow = headerRow; HeaderColumn = headerColumn; FooterRow = footerRow;
    }
    /// <summary>Gets body-cell defaults, or null when absent or unresolved.</summary>
    public IWorkParagraphStyle? Body { get; }
    /// <summary>Gets leading header-row defaults, or null when absent or unresolved.</summary>
    public IWorkParagraphStyle? HeaderRow { get; }
    /// <summary>Gets leading header-column defaults, or null when absent or unresolved.</summary>
    public IWorkParagraphStyle? HeaderColumn { get; }
    /// <summary>Gets trailing footer-row defaults, or null when absent or unresolved.</summary>
    public IWorkParagraphStyle? FooterRow { get; }
}
