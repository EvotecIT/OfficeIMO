namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    internal void AddColumnBreak() => AddBlock(new ColumnBreakBlock());

    /// <summary>Flows common block content across equal-width columns in reading order.</summary>
    internal PdfDocument Columns(Action<PdfContentBuilder> compose, PdfMultiColumnOptions? options = null) {
        Guard.NotNull(compose, nameof(compose));
        AddBlock(new MultiColumnBlock(BuildFlowBlocks(compose), options));
        return this;
    }

    /// <summary>Groups common flow blocks inside a padded, styled, one-page container.</summary>
    internal PdfDocument Container(Action<PdfContentBuilder> compose, PdfPanelStyle? style = null) {
        Guard.NotNull(compose, nameof(compose));
        AddBlock(new ContainerBlock(BuildFlowBlocks(compose), style));
        return this;
    }

    internal void AddElement(
        Action<PdfContentBuilder> compose,
        PdfPanelStyle? style,
        PdfSemanticRole? semanticRole,
        string? alternativeText) {
        Guard.NotNull(compose, nameof(compose));
        System.Collections.Generic.IReadOnlyList<IPdfBlock> blocks = BuildFlowBlocks(compose);

        if (style == null && !semanticRole.HasValue) {
            foreach (IPdfBlock block in blocks) {
                AddBlock(block);
            }

            return;
        }

        IPdfBlock element = style == null
            ? new FlowBlock(blocks, options: null, capture: null)
            : new ContainerBlock(blocks, style);
        if (semanticRole.HasValue) {
            element = new SemanticBlock(semanticRole.Value, new[] { element }, alternativeText);
        }

        AddBlock(element);
    }
}
