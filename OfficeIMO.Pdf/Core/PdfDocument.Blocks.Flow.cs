#pragma warning disable RS0026 // Text and callback paragraph forms share AddParagraph and its optional layout parameters.
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    /// <summary>Appends a paragraph to the current flow and returns it.</summary>
    /// <param name="text">Optional first text run, formatted from the paragraph run defaults.</param>
    /// <param name="align">Paragraph alignment.</param>
    /// <param name="color">Optional text colour; it becomes <see cref="PdfTextRunDefaults.Color"/> and the colour of runs that carry none.</param>
    /// <param name="style">Optional paragraph layout style.</param>
    /// <returns>The paragraph. It stays editable until the document is laid out, so <c>AddParagraph().AddText("x").SetBold()</c> renders bold text.</returns>
    public PdfParagraph AddParagraph(string? text = null, OfficeAlignment align = OfficeAlignment.Left, global::OfficeIMO.OfficeColor? color = null, PdfParagraphStyle? style = null) {
        var paragraph = new PdfParagraph(align, color, this);
        if (text != null) paragraph.AddText(text);
        AddBlock(paragraph.CreateBlock(style));
        return paragraph;
    }

    /// <summary>Appends a paragraph built by a callback and returns it.</summary>
    /// <param name="build">Adds and formats the paragraph runs.</param>
    /// <param name="align">Paragraph alignment.</param>
    /// <param name="color">Optional text colour; it becomes <see cref="PdfTextRunDefaults.Color"/> and the colour of runs that carry none.</param>
    /// <param name="style">Optional paragraph layout style.</param>
    public PdfParagraph AddParagraph(System.Action<PdfParagraph> build, OfficeAlignment align = OfficeAlignment.Left, global::OfficeIMO.OfficeColor? color = null, PdfParagraphStyle? style = null) {
        Guard.NotNull(build, nameof(build));
        var paragraph = new PdfParagraph(align, color, this);
        build(paragraph);
        AddBlock(paragraph.CreateBlock(style));
        return paragraph;
    }

    // Chaining form used by adapters and tests that compose a document in one expression.
    internal PdfDocument Paragraph(System.Action<PdfParagraph> build, OfficeAlignment align = OfficeAlignment.Left, global::OfficeIMO.OfficeColor? color = null, PdfParagraphStyle? style = null) {
        AddParagraph(build, align, color, style);
        return this;
    }

    /// <summary>
    /// Higher-level composition model (page size/margins/footer + content), similar to other document DSLs.
    /// Sugar only; composes into the same PdfDocument blocks and options.
    /// </summary>
    public PdfDocument Compose(System.Action<PdfDocumentBuilder> compose) {
        Guard.NotNull(compose, nameof(compose));
        var c = new PdfDocumentBuilder(this);
        compose(c);
        return this;
    }

    /// <summary>Adds a horizontal rule (line) spanning the content width.</summary>
    internal PdfDocument HR(double? thickness = null, global::OfficeIMO.OfficeColor? color = null, double? spacingBefore = null, double? spacingAfter = null, PdfHorizontalRuleStyle? style = null) {
        AddBlock(new HorizontalRuleBlock(CreateHorizontalRuleStyle(thickness, color, spacingBefore, spacingAfter, style)));
        return this;
    }

    /// <summary>Adds a named bookmark at the current flow position.</summary>
    internal PdfDocument Bookmark(string name) {
        AddBlock(new BookmarkBlock(name));
        return this;
    }
}
