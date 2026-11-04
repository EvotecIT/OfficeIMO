namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    /// <summary>Adds one paragraph inside a decorated panel.</summary>
    internal PdfDocument PanelParagraph(System.Action<PdfParagraph> compose, PdfPanelStyle? style = null, OfficeAlignment align = OfficeAlignment.Left, global::OfficeIMO.OfficeColor? color = null) {
        Guard.NotNull(compose, nameof(compose));
        return Panel(content => content.AddParagraph(compose, align, color), style);
    }

    /// <summary>Adds a block-preserving decorated flow element.</summary>
    internal PdfDocument Panel(System.Action<PdfContentBuilder> compose, PdfPanelStyle? style = null) {
        Guard.NotNull(compose, nameof(compose));
        AddBlock(new ContainerBlock(CollapseTrailingPanelSpacing(BuildFlowBlocks(compose)), style, useDefaultPanelStyle: true));
        return this;
    }

    // The panel padding is the space under its last paragraph; adding that paragraph's own
    // spacing after makes the bottom inset visibly larger than the top one.
    private System.Collections.Generic.List<IPdfBlock> CollapseTrailingPanelSpacing(System.Collections.Generic.IEnumerable<IPdfBlock> blocks) {
        var list = new System.Collections.Generic.List<IPdfBlock>(blocks);
        if (list.Count > 0 && list[list.Count - 1] is RichParagraphBlock last && last.Style?.SpacingAfter == null) {
            PdfParagraphStyle collapsed = last.Style?.Clone() ?? _options.DefaultParagraphStyle ?? new PdfParagraphStyle();
            collapsed.SpacingAfter = OfficeLength.Zero;
            list[list.Count - 1] = last.WithStyle(collapsed);
        }
        return list;
    }
}
