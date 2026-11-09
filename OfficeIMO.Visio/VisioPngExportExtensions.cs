using OfficeIMO.Drawing;

namespace OfficeIMO.Visio {
    /// <summary>Adapts the native semantic-preview PNG options to the shared image renderer.</summary>
    internal static class VisioPngExportExtensions {
        internal static OfficeImageExportResult CreateResult(
            VisioPage page,
            VisioPngSaveOptions? options,
            CancellationToken cancellationToken = default) {
            ArgumentNullException.ThrowIfNull(page);
            VisioPngSaveOptions resolved = options ?? new VisioPngSaveOptions();
            using CancellationTokenSource linkedCancellation =
                CancellationTokenSource.CreateLinkedTokenSource(
                    resolved.CancellationToken,
                    cancellationToken);
            var canonical = new VisioImageExportOptions {
                TargetDpi = resolved.PixelsPerInch,
                BackgroundColor = resolved.BackgroundColor ?? OfficeColor.Transparent,
                RenderText = resolved.RenderText,
                LayerMode = resolved.LayerMode,
                FontFilePath = resolved.FontFilePath,
                FontFaceName = resolved.FontFaceName,
                FontCollectionIndex = resolved.FontCollectionIndex,
                Fonts = resolved.Fonts.Clone(),
                TextShapingProvider = resolved.TextShapingProvider,
                TextShapingLanguage = resolved.TextShapingLanguage,
                RenderStencilArtwork = resolved.RenderStencilArtwork,
                RenderConnectorLabels = resolved.RenderConnectorLabels,
                ResolveConnectorLabelOverlaps = resolved.ResolveConnectorLabelOverlaps,
                Supersampling = resolved.Supersampling,
                MaximumRasterPixels = resolved.MaximumRasterPixels,
                MaximumTotalEncodedBytes = resolved.MaximumTotalEncodedBytes,
                RasterOverflowBehavior = resolved.RasterOverflowBehavior,
                ImageCodec = resolved.ImageCodec
            };
            return VisioImageExportEngine.Render(
                page,
                OfficeImageExportFormat.Png,
                canonical,
                page.Name,
                "Visio page",
                linkedCancellation.Token);
        }
    }
}
