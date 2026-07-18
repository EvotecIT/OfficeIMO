using OfficeIMO.Drawing;
using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Native dependency-free PNG export helpers for OfficeIMO Visio documents and pages.
    /// </summary>
    public static class VisioPngExportExtensions {
        /// <summary>
        /// Renders the selected document page to PNG bytes without requiring Microsoft Visio desktop automation.
        /// </summary>
        public static byte[] ToPng(this VisioDocument document, VisioPngSaveOptions? options = null) {
            return CreateResult(document, options).Bytes;
        }

        /// <summary>
        /// Renders a page to PNG bytes without requiring Microsoft Visio desktop automation.
        /// </summary>
        public static byte[] ToPng(this VisioPage page, VisioPngSaveOptions? options = null) {
            return CreateResult(page, options).Bytes;
        }

        /// <summary>
        /// Saves the selected document page as PNG without requiring Microsoft Visio desktop automation.
        /// </summary>
        public static OfficeImageExportResult SaveAsPng(this VisioDocument document, string path, VisioPngSaveOptions? options = null) {
            return CreateResult(document, options).Save(path);
        }

        /// <summary>
        /// Saves a page as PNG without requiring Microsoft Visio desktop automation.
        /// </summary>
        public static OfficeImageExportResult SaveAsPng(this VisioPage page, string path, VisioPngSaveOptions? options = null) {
            return CreateResult(page, options).Save(path);
        }

        /// <summary>
        /// Writes the selected document page as PNG to a stream without requiring Microsoft Visio desktop automation.
        /// </summary>
        public static OfficeImageExportResult SaveAsPng(this VisioDocument document, Stream stream, VisioPngSaveOptions? options = null) {
            return CreateResult(document, options).Save(stream);
        }

        /// <summary>
        /// Writes a page as PNG to a stream without requiring Microsoft Visio desktop automation.
        /// </summary>
        public static OfficeImageExportResult SaveAsPng(this VisioPage page, Stream stream, VisioPngSaveOptions? options = null) {
            return CreateResult(page, options).Save(stream);
        }

        /// <summary>
        /// Asynchronously saves the selected document page as PNG without requiring Microsoft Visio desktop automation.
        /// </summary>
        public static async Task<OfficeImageExportResult> SaveAsPngAsync(
            this VisioDocument document,
            string path,
            VisioPngSaveOptions? options = null,
            CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            OfficeImageExportResult result = CreateResult(document, options, cancellationToken);
            return await result.SaveAsync(path, cancellationToken: cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// Asynchronously saves a page as PNG without requiring Microsoft Visio desktop automation.
        /// </summary>
        public static async Task<OfficeImageExportResult> SaveAsPngAsync(
            this VisioPage page,
            string path,
            VisioPngSaveOptions? options = null,
            CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            OfficeImageExportResult result = CreateResult(page, options, cancellationToken);
            return await result.SaveAsync(path, cancellationToken: cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// Asynchronously writes the selected document page as PNG to a stream without closing it.
        /// </summary>
        public static async Task<OfficeImageExportResult> SaveAsPngAsync(
            this VisioDocument document,
            Stream stream,
            VisioPngSaveOptions? options = null,
            CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            OfficeImageExportResult result = CreateResult(document, options, cancellationToken);
            return await result.SaveAsync(stream, cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// Asynchronously writes a page as PNG to a stream without closing it.
        /// </summary>
        public static async Task<OfficeImageExportResult> SaveAsPngAsync(
            this VisioPage page,
            Stream stream,
            VisioPngSaveOptions? options = null,
            CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            OfficeImageExportResult result = CreateResult(page, options, cancellationToken);
            return await result.SaveAsync(stream, cancellationToken).ConfigureAwait(false);
        }

        private static OfficeImageExportResult CreateResult(
            VisioDocument document,
            VisioPngSaveOptions? options,
            CancellationToken cancellationToken = default) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            VisioPngSaveOptions resolved = options ?? new VisioPngSaveOptions();
            if (document.Pages.Count == 0) {
                throw new InvalidOperationException("The document does not contain any pages to export.");
            }

            if (resolved.PageIndex < 0 || resolved.PageIndex >= document.Pages.Count) {
                throw new ArgumentOutOfRangeException(nameof(options), "PageIndex is outside the document page collection.");
            }

            return CreateResult(document.Pages[resolved.PageIndex], resolved, cancellationToken);
        }

        private static OfficeImageExportResult CreateResult(
            VisioPage page,
            VisioPngSaveOptions? options,
            CancellationToken cancellationToken = default) {
            if (page == null) throw new ArgumentNullException(nameof(page));
            VisioPngSaveOptions resolved = options ?? new VisioPngSaveOptions();
            using CancellationTokenSource linkedCancellation =
                CancellationTokenSource.CreateLinkedTokenSource(
                    resolved.CancellationToken,
                    cancellationToken);
            var canonical = new VisioImageExportOptions {
                TargetDpi = resolved.PixelsPerInch,
                BackgroundColor = resolved.BackgroundColor ?? OfficeColor.Transparent,
                RenderText = resolved.RenderText,
                FontFilePath = resolved.FontFilePath,
                FontFaceName = resolved.FontFaceName,
                FontCollectionIndex = resolved.FontCollectionIndex,
                Fonts = resolved.Fonts.Clone(),
                RenderStencilArtwork = resolved.RenderStencilArtwork,
                RenderConnectorLabels = resolved.RenderConnectorLabels,
                ResolveConnectorLabelOverlaps = resolved.ResolveConnectorLabelOverlaps,
                Supersampling = resolved.Supersampling,
                MaximumRasterPixels = resolved.MaximumRasterPixels,
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
