using OfficeIMO.Drawing.Internal;
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
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("PNG path cannot be null or whitespace.", nameof(path));
            OfficeImageExportResult result = CreateResult(document, options);
            OfficeFileCommit.WriteAllBytes(path, result.Bytes);
            return result;
        }

        /// <summary>
        /// Saves a page as PNG without requiring Microsoft Visio desktop automation.
        /// </summary>
        public static OfficeImageExportResult SaveAsPng(this VisioPage page, string path, VisioPngSaveOptions? options = null) {
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("PNG path cannot be null or whitespace.", nameof(path));
            OfficeImageExportResult result = CreateResult(page, options);
            OfficeFileCommit.WriteAllBytes(path, result.Bytes);
            return result;
        }

        /// <summary>
        /// Writes the selected document page as PNG to a stream without requiring Microsoft Visio desktop automation.
        /// </summary>
        public static OfficeImageExportResult SaveAsPng(this VisioDocument document, Stream stream, VisioPngSaveOptions? options = null) {
            OfficeImageExportResult result = CreateResult(document, options);
            OfficeStreamWriter.WriteAllBytes(stream, result.Bytes);
            return result;
        }

        /// <summary>
        /// Writes a page as PNG to a stream without requiring Microsoft Visio desktop automation.
        /// </summary>
        public static OfficeImageExportResult SaveAsPng(this VisioPage page, Stream stream, VisioPngSaveOptions? options = null) {
            OfficeImageExportResult result = CreateResult(page, options);
            OfficeStreamWriter.WriteAllBytes(stream, result.Bytes);
            return result;
        }

        /// <summary>
        /// Asynchronously saves the selected document page as PNG without requiring Microsoft Visio desktop automation.
        /// </summary>
        public static async Task<OfficeImageExportResult> SaveAsPngAsync(
            this VisioDocument document,
            string path,
            VisioPngSaveOptions? options = null,
            CancellationToken cancellationToken = default) {
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("PNG path cannot be null or whitespace.", nameof(path));
            cancellationToken.ThrowIfCancellationRequested();
            OfficeImageExportResult result = CreateResult(document, options);
            await OfficeFileCommit.WriteAllBytesAsync(path, result.Bytes, cancellationToken: cancellationToken).ConfigureAwait(false);
            return result;
        }

        /// <summary>
        /// Asynchronously saves a page as PNG without requiring Microsoft Visio desktop automation.
        /// </summary>
        public static async Task<OfficeImageExportResult> SaveAsPngAsync(
            this VisioPage page,
            string path,
            VisioPngSaveOptions? options = null,
            CancellationToken cancellationToken = default) {
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("PNG path cannot be null or whitespace.", nameof(path));
            cancellationToken.ThrowIfCancellationRequested();
            OfficeImageExportResult result = CreateResult(page, options);
            await OfficeFileCommit.WriteAllBytesAsync(path, result.Bytes, cancellationToken: cancellationToken).ConfigureAwait(false);
            return result;
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
            OfficeImageExportResult result = CreateResult(document, options);
            await OfficeStreamWriter.WriteAllBytesAsync(stream, result.Bytes, cancellationToken).ConfigureAwait(false);
            return result;
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
            OfficeImageExportResult result = CreateResult(page, options);
            await OfficeStreamWriter.WriteAllBytesAsync(stream, result.Bytes, cancellationToken).ConfigureAwait(false);
            return result;
        }

        private static OfficeImageExportResult CreateResult(VisioDocument document, VisioPngSaveOptions? options) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            VisioPngSaveOptions resolved = options ?? new VisioPngSaveOptions();
            if (document.Pages.Count == 0) {
                throw new InvalidOperationException("The document does not contain any pages to export.");
            }

            if (resolved.PageIndex < 0 || resolved.PageIndex >= document.Pages.Count) {
                throw new ArgumentOutOfRangeException(nameof(options), "PageIndex is outside the document page collection.");
            }

            return CreateResult(document.Pages[resolved.PageIndex], resolved);
        }

        private static OfficeImageExportResult CreateResult(VisioPage page, VisioPngSaveOptions? options) {
            if (page == null) throw new ArgumentNullException(nameof(page));
            VisioPngSaveOptions resolved = options ?? new VisioPngSaveOptions();
            byte[] bytes = VisioPngRenderer.Render(page, resolved);
            return new OfficeImageExportResult(
                OfficeImageExportFormat.Png,
                Math.Max(1, (int)Math.Ceiling(Math.Max(page.Width, 0.01D) * resolved.PixelsPerInch)),
                Math.Max(1, (int)Math.Ceiling(Math.Max(page.Height, 0.01D) * resolved.PixelsPerInch)),
                bytes,
                page.Name);
        }
    }
}
