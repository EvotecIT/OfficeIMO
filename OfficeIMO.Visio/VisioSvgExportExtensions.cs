using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Headless SVG export helpers for OfficeIMO Visio documents and pages.
    /// </summary>
    public static class VisioSvgExportExtensions {
        /// <summary>
        /// Renders the selected document page to SVG without requiring Microsoft Visio desktop automation.
        /// </summary>
        public static string ToSvg(this VisioDocument document, VisioSvgSaveOptions? options = null) {
            return Encoding.UTF8.GetString(CreateResult(document, options).Bytes);
        }

        /// <summary>
        /// Renders a page to SVG without requiring Microsoft Visio desktop automation.
        /// </summary>
        public static string ToSvg(this VisioPage page, VisioSvgSaveOptions? options = null) {
            return Encoding.UTF8.GetString(CreateResult(page, options).Bytes);
        }

        /// <summary>
        /// Saves the selected document page as SVG without requiring Microsoft Visio desktop automation.
        /// </summary>
        public static OfficeImageExportResult SaveAsSvg(this VisioDocument document, string path, VisioSvgSaveOptions? options = null) {
            return CreateResult(document, options).Save(path);
        }

        /// <summary>
        /// Saves a page as SVG without requiring Microsoft Visio desktop automation.
        /// </summary>
        public static OfficeImageExportResult SaveAsSvg(this VisioPage page, string path, VisioSvgSaveOptions? options = null) {
            return CreateResult(page, options).Save(path);
        }

        /// <summary>
        /// Writes the selected document page as SVG to a stream without requiring Microsoft Visio desktop automation.
        /// </summary>
        public static OfficeImageExportResult SaveAsSvg(this VisioDocument document, Stream stream, VisioSvgSaveOptions? options = null) {
            return CreateResult(document, options).Save(stream);
        }

        /// <summary>
        /// Writes a page as SVG to a stream without requiring Microsoft Visio desktop automation.
        /// </summary>
        public static OfficeImageExportResult SaveAsSvg(this VisioPage page, Stream stream, VisioSvgSaveOptions? options = null) {
            return CreateResult(page, options).Save(stream);
        }

        /// <summary>
        /// Asynchronously saves the selected document page as SVG without requiring Microsoft Visio desktop automation.
        /// </summary>
        public static async Task<OfficeImageExportResult> SaveAsSvgAsync(
            this VisioDocument document,
            string path,
            VisioSvgSaveOptions? options = null,
            CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            OfficeImageExportResult result = CreateResult(document, options, cancellationToken);
            return await result.SaveAsync(path, cancellationToken: cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// Asynchronously saves a page as SVG without requiring Microsoft Visio desktop automation.
        /// </summary>
        public static async Task<OfficeImageExportResult> SaveAsSvgAsync(
            this VisioPage page,
            string path,
            VisioSvgSaveOptions? options = null,
            CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            OfficeImageExportResult result = CreateResult(page, options, cancellationToken);
            return await result.SaveAsync(path, cancellationToken: cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// Asynchronously writes the selected document page as SVG to a stream without closing it.
        /// </summary>
        public static async Task<OfficeImageExportResult> SaveAsSvgAsync(
            this VisioDocument document,
            Stream stream,
            VisioSvgSaveOptions? options = null,
            CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            OfficeImageExportResult result = CreateResult(document, options, cancellationToken);
            return await result.SaveAsync(stream, cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// Asynchronously writes a page as SVG to a stream without closing it.
        /// </summary>
        public static async Task<OfficeImageExportResult> SaveAsSvgAsync(
            this VisioPage page,
            Stream stream,
            VisioSvgSaveOptions? options = null,
            CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            OfficeImageExportResult result = CreateResult(page, options, cancellationToken);
            return await result.SaveAsync(stream, cancellationToken).ConfigureAwait(false);
        }

        private static OfficeImageExportResult CreateResult(
            VisioDocument document,
            VisioSvgSaveOptions? options,
            CancellationToken cancellationToken = default) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            VisioSvgSaveOptions resolved = options?.Clone() ?? new VisioSvgSaveOptions();
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
            VisioSvgSaveOptions? options,
            CancellationToken cancellationToken = default) {
            if (page == null) throw new ArgumentNullException(nameof(page));
            VisioSvgSaveOptions resolved = options?.Clone() ?? new VisioSvgSaveOptions();
            var canonical = new VisioImageExportOptions {
                TargetDpi = resolved.PixelsPerInch,
                BackgroundColor = resolved.BackgroundColor ?? OfficeColor.Transparent,
                RenderText = resolved.RenderText,
                Fonts = resolved.Fonts.Clone(),
                RenderStencilArtwork = resolved.RenderStencilArtwork,
                RenderConnectorLabels = resolved.RenderConnectorLabels,
                ResolveConnectorLabelOverlaps = resolved.ResolveConnectorLabelOverlaps,
                IncludeSvgXmlDeclaration = resolved.IncludeXmlDeclaration,
                ImageCodec = resolved.ImageCodec
            };
            return VisioImageExportEngine.Render(
                page,
                OfficeImageExportFormat.Svg,
                canonical,
                page.Name,
                "Visio page",
                cancellationToken);
        }
    }
}
