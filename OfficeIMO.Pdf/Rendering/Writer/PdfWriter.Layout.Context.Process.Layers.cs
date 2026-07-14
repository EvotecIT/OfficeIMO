namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private void RenderLayerBlock(LayerBlock layer) {
            EnsurePage();
            activeLayers.Add(layer.Definition);
            BeginLayerContent(layer.Definition);
            try {
                ProcessBlocks(layer.Blocks);
            } finally {
                if (currentPage != null) sb.Append("EMC\n");
                activeLayers.RemoveAt(activeLayers.Count - 1);
            }
        }
    }
}
