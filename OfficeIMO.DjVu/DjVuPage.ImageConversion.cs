using System.Globalization;
using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.DjVu;

public sealed partial class DjVuPage {
    /// <summary>Encodes the native page raster through shared image limits and diagnostics. SVG embeds the raster; it does not trace vectors.</summary>
    public OfficeImageExportResult ToImage(OfficeImageExportOptions? options = null,
        DjVuRenderOptions? renderOptions = null, CancellationToken cancellationToken = default) {
        OfficeImageExportOptions image = (options ?? new OfficeImageExportOptions()).SnapshotForUnits(Dpi);
        DjVuRenderOptions native = renderOptions?.Clone() ?? new DjVuRenderOptions();
        if (!image.TargetDpi.HasValue && native.Dpi.HasValue) image.Scale *= (double)native.Dpi.Value / Dpi;
        return OfficeImageExportExecutionScope.Run(image, cancellationToken, token => {
            string source = "DjVu page " + Number;
            DjVuRectangle region = native.Region ?? new DjVuRectangle(0, 0, Width, Height);
            bool rotated = native.ApplyRotation && Rotation % 180 != 0;
            int width = rotated ? region.Height : region.Width, height = rotated ? region.Width : region.Height;
            OfficeImageExportFormat rasterFormat = image.Format == OfficeImageExportFormat.Svg ? OfficeImageExportFormat.Png : image.Format;
            var plan = OfficeRasterExportPlanner.Resolve(width, height, rasterFormat, image, source);
            double requestedDpi = Math.Max(1D, Math.Floor(Dpi * plan.Limit.Scale));
            if (requestedDpi > 6000D && requestedDpi != Dpi)
                throw new ArgumentOutOfRangeException(nameof(options), "The DjVu renderer accepts an explicit output density of at most 6000 DPI.");
            native.Dpi = requestedDpi == Dpi ? null : (int)requestedDpi;
            native.MaxPixels = Math.Min(native.MaxPixels, image.MaximumRasterPixels);
            native.Background = image.BackgroundColor;
            DjVuRenderResult rendered = Render(native, token);
            OfficeRasterEncodingOptions encoding = plan.CreateEncodingOptions();
            encoding.Resolution = new OfficeImageResolution(rendered.Dpi, rendered.Dpi);
            byte[] bytes = OfficeRasterImageEncoder.Encode(rendered.Image, rasterFormat, encoding, image.MaximumTotalEncodedBytes, token);
            if (image.Format == OfficeImageExportFormat.Svg) {
                string svg = "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"" + rendered.Image.Width.ToString(CultureInfo.InvariantCulture)
                    + "\" height=\"" + rendered.Image.Height.ToString(CultureInfo.InvariantCulture) + "\"><image width=\"100%\" height=\"100%\" href=\"data:image/png;base64,"
                    + Convert.ToBase64String(bytes) + "\"/></svg>";
                if (Encoding.UTF8.GetByteCount(svg) > image.MaximumTotalEncodedBytes) throw new InvalidOperationException("SVG output exceeds the configured encoded byte limit.");
                bytes = Encoding.UTF8.GetBytes(svg);
            }
            var diagnostics = plan.Diagnostic == null ? Array.Empty<OfficeImageExportDiagnostic>() : new[] { plan.Diagnostic };
            var result = new OfficeImageExportResult(image.Format, rendered.Image.Width, rendered.Image.Height, bytes,
                "Page " + Number, source, diagnostics);
            return image.EnsureAccepted(OfficeImageExportEvidence.AttachSourceReport(result, rendered));
        });
    }
}
