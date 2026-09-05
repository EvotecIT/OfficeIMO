namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    /// <summary>Adds one paragraph inside a decorated panel.</summary>
    internal PdfDocument PanelParagraph(System.Action<PdfParagraphBuilder> compose, PdfPanelStyle? style = null, PdfAlign align = PdfAlign.Left, PdfColor? defaultColor = null) {
        Guard.NotNull(compose, nameof(compose));
        return Panel(content => content.Paragraph(compose, align, defaultColor), style);
    }

    /// <summary>Adds a block-preserving decorated flow element.</summary>
    internal PdfDocument Panel(System.Action<PdfContentBuilder> compose, PdfPanelStyle? style = null) {
        Guard.NotNull(compose, nameof(compose));
        AddBlock(new ContainerBlock(BuildFlowBlocks(compose), style, useDefaultPanelStyle: true));
        return this;
    }
}
