namespace OfficeIMO.Pdf;

public sealed partial class PdfDoc {
    /// <summary>Adds a level-1 heading.</summary>
    public PdfDoc H1(string text, PdfAlign align = PdfAlign.Left, PdfColor? color = null, string? linkUri = null) {
        Guard.NotNull(text, nameof(text));
        AddBlock(new HeadingBlock(1, text, align, color, linkUri)); return this; }
    /// <summary>Adds a level-2 heading.</summary>
    public PdfDoc H2(string text, PdfAlign align = PdfAlign.Left, PdfColor? color = null, string? linkUri = null) {
        Guard.NotNull(text, nameof(text));
        AddBlock(new HeadingBlock(2, text, align, color, linkUri)); return this; }
    /// <summary>Adds a level-3 heading.</summary>
    public PdfDoc H3(string text, PdfAlign align = PdfAlign.Left, PdfColor? color = null, string? linkUri = null) {
        Guard.NotNull(text, nameof(text));
        AddBlock(new HeadingBlock(3, text, align, color, linkUri)); return this; }

    /// <summary>Inserts a page break.</summary>
    public PdfDoc PageBreak() { AddBlock(new PageBreakBlock()); return this; }

    /// <summary>Adds a simple bullet list.</summary>
    public PdfDoc Bullets(System.Collections.Generic.IEnumerable<string> items, PdfAlign align = PdfAlign.Left, PdfColor? color = null) {
        AddBlock(new BulletListBlock(items, align, color));
        return this;
    }

    /// <summary>
    /// Fluent paragraph builder in the style of OfficeIMO.Word.
    /// Example: Paragraph(p => p.Color(PdfColor.Gray).Text("You can ").Bold("mix ").Italic("styles"));
    /// </summary>
    public PdfDoc Paragraph(System.Action<PdfParagraphBuilder> compose, PdfAlign align = PdfAlign.Left, PdfColor? defaultColor = null) {
        Guard.NotNull(compose, nameof(compose));
        var builder = new PdfParagraphBuilder(align, defaultColor);
        compose(builder);
        AddBlock(builder.Build());
        return this;
    }

    /// <summary>
    /// Higher-level composition model (page size/margins/footer + content), similar to other document DSLs.
    /// Sugar only; composes into the same PdfDoc blocks and options.
    /// </summary>
    public PdfDoc Compose(System.Action<PdfCompose> compose) {
        Guard.NotNull(compose, nameof(compose));
        var c = new PdfCompose(this);
        compose(c);
        return this;
    }

    /// <summary>Adds a horizontal rule (line) spanning the content width.</summary>
    public PdfDoc HR(double thickness = 0.5, PdfColor? color = null, double spacingBefore = 6, double spacingAfter = 6) {
        AddBlock(new HorizontalRuleBlock(thickness, color ?? PdfColor.Gray, spacingBefore, spacingAfter));
        return this;
    }

    /// <summary>Adds a paragraph inside a simple panel (background + optional border).</summary>
    public PdfDoc PanelParagraph(System.Action<PdfParagraphBuilder> compose, PanelStyle style, PdfAlign align = PdfAlign.Left, PdfColor? defaultColor = null) {
        Guard.NotNull(compose, nameof(compose));
        Guard.NotNull(style, nameof(style));
        var builder = new PdfParagraphBuilder(align, defaultColor);
        compose(builder);
        AddBlock(new PanelParagraphBlock(builder.Build().Runs, align, defaultColor, style));
        return this;
    }

    /// <summary>Adds a JPEG image at the current flow position.</summary>
    public PdfDoc Image(byte[] jpegBytes, double width, double height, PdfAlign align = PdfAlign.Left) {
        AddBlock(new ImageBlock(jpegBytes, width, height, align));
        return this;
    }

    // Internal for Compose Row
    internal void AddRow(RowBlock row) { AddBlock(row); }

    /// <summary>Adds a simple table from rows of string arrays.</summary>
    public PdfDoc Table(System.Collections.Generic.IEnumerable<string[]> rows, PdfAlign align = PdfAlign.Left, PdfTableStyle? style = null) {
        AddBlock(new TableBlock(rows, align, style));
        return this;
    }

    /// <summary>
    /// Adds a table and attaches link URIs to specific cells.
    /// </summary>
    public PdfDoc TableWithLinks(System.Collections.Generic.IEnumerable<string[]> rows, System.Collections.Generic.Dictionary<(int Row, int Col), string> links, PdfAlign align = PdfAlign.Left, PdfTableStyle? style = null) {
        var tb = new TableBlock(rows, align, style);
        if (links != null) foreach (var kv in links) tb.Links[kv.Key] = kv.Value;
        AddBlock(tb);
        return this;
    }
}

