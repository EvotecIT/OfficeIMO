namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    internal void AddColumnBreak() => AddBlock(new ColumnBreakBlock());

    /// <summary>Flows common block content across equal-width columns in reading order.</summary>
    public PdfDocument Columns(Action<PdfItemCompose> compose, PdfMultiColumnOptions? options = null) {
        Guard.NotNull(compose, nameof(compose));
        AddBlock(new MultiColumnBlock(BuildFlowBlocks(compose), options));
        return this;
    }

    /// <summary>Groups common flow blocks inside a padded, styled, one-page container.</summary>
    public PdfDocument Container(Action<PdfItemCompose> compose, PanelStyle? style = null) {
        Guard.NotNull(compose, nameof(compose));
        AddBlock(new ContainerBlock(BuildFlowBlocks(compose), style));
        return this;
    }
}
