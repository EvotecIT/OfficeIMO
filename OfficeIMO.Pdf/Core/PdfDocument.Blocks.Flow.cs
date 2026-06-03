using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    /// <summary>
    /// Fluent paragraph builder in the style of OfficeIMO.Word.
    /// Example: Paragraph(p => p.Color(PdfColor.Gray).Text("You can ").Bold("mix ").Italic("styles"));
    /// </summary>
    public PdfDocument Paragraph(System.Action<PdfParagraphBuilder> compose, PdfAlign align = PdfAlign.Left, PdfColor? defaultColor = null, PdfParagraphStyle? style = null) {
        Guard.NotNull(compose, nameof(compose));
        var builder = new PdfParagraphBuilder(align, defaultColor);
        compose(builder);
        AddBlock(builder.Build(style));
        return this;
    }

    /// <summary>
    /// Higher-level composition model (page size/margins/footer + content), similar to other document DSLs.
    /// Sugar only; composes into the same PdfDocument blocks and options.
    /// </summary>
    public PdfDocument Compose(System.Action<PdfCompose> compose) {
        Guard.NotNull(compose, nameof(compose));
        var c = new PdfCompose(this);
        compose(c);
        return this;
    }

    /// <summary>Adds a horizontal rule (line) spanning the content width.</summary>
    public PdfDocument HR(double? thickness = null, PdfColor? color = null, double? spacingBefore = null, double? spacingAfter = null, PdfHorizontalRuleStyle? style = null) {
        AddBlock(new HorizontalRuleBlock(CreateHorizontalRuleStyle(thickness, color, spacingBefore, spacingAfter, style)));
        return this;
    }

    /// <summary>Adds a named bookmark at the current flow position.</summary>
    public PdfDocument Bookmark(string name) {
        AddBlock(new BookmarkBlock(name));
        return this;
    }
}
