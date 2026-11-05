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
        /// <summary>Renders the selected page as decoded SVG text.</summary>
        public static string ToSvg(this VisioDocument document, VisioSvgSaveOptions? options = null,
            CancellationToken cancellationToken = default) => document.ToSvg(options, out _, cancellationToken);

        /// <summary>Renders SVG text while retaining structured image fidelity diagnostics.</summary>
        public static string ToSvg(this VisioDocument document, VisioSvgSaveOptions? options,
            out OfficeConversionReport report, CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            OfficeImageExportResult result = CreateResult(document, options, cancellationToken);
            report = OfficeImageExportEvidence.CreateReport("Visio", result);
            return Encoding.UTF8.GetString(result.Bytes);
        }

        /// <summary>Renders SVG text asynchronously and returns the completed report.</summary>
        public static Task<OfficeConversionResult<string>> ToSvgAsync(this VisioDocument document,
            VisioSvgSaveOptions? options = null, CancellationToken cancellationToken = default) =>
            Task.Run(() => {
                string value = document.ToSvg(options, out OfficeConversionReport report, cancellationToken);
                return new OfficeConversionResult<string>(value, report);
            }, cancellationToken);

        /// <summary>Renders the selected page as decoded SVG text.</summary>
        public static string ToSvg(this VisioPage document, VisioSvgSaveOptions? options = null,
            CancellationToken cancellationToken = default) => document.ToSvg(options, out _, cancellationToken);

        /// <summary>Renders SVG text while retaining structured image fidelity diagnostics.</summary>
        public static string ToSvg(this VisioPage document, VisioSvgSaveOptions? options,
            out OfficeConversionReport report, CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            OfficeImageExportResult result = CreateResult(document, options, cancellationToken);
            report = OfficeImageExportEvidence.CreateReport("Visio", result);
            return Encoding.UTF8.GetString(result.Bytes);
        }

        /// <summary>Renders SVG text asynchronously and returns the completed report.</summary>
        public static Task<OfficeConversionResult<string>> ToSvgAsync(this VisioPage document,
            VisioSvgSaveOptions? options = null, CancellationToken cancellationToken = default) =>
            Task.Run(() => {
                string value = document.ToSvg(options, out OfficeConversionReport report, cancellationToken);
                return new OfficeConversionResult<string>(value, report);
            }, cancellationToken);

        private static OfficeImageExportResult CreateResult(
            VisioDocument document,
            VisioSvgSaveOptions? options,
            CancellationToken cancellationToken = default) {
            ArgumentNullException.ThrowIfNull(document);
            VisioSvgSaveOptions resolved = options?.Clone() ?? new VisioSvgSaveOptions();
            if (document.Pages.Count == 0) {
                throw new InvalidOperationException("The document does not contain any pages to export.");
            }

            if (resolved.PageIndex < 0 || resolved.PageIndex >= document.Pages.Count) {
                throw new ArgumentOutOfRangeException(nameof(options), "PageIndex is outside the document page collection.");
            }

            return CreateResult(document.Pages[resolved.PageIndex], resolved, cancellationToken);
        }

        internal static OfficeImageExportResult CreateResult(
            VisioPage page,
            VisioSvgSaveOptions? options,
            CancellationToken cancellationToken = default) {
            ArgumentNullException.ThrowIfNull(page);
            VisioSvgSaveOptions resolved = options?.Clone() ?? new VisioSvgSaveOptions();
            var canonical = new VisioImageExportOptions {
                Scale = resolved.PixelsPerInch / 96D,
                BackgroundColor = resolved.BackgroundColor ?? OfficeColor.Transparent,
                RenderText = resolved.RenderText,
                LayerMode = resolved.LayerMode,
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
