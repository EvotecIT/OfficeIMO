using System;
using System.Threading;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    // A measurement-only canvas shares the positioned paint path without loading installed
    // faces or allocating a raster surface. Font selection remains owned by the caller.
    private OfficeRasterCanvas(OfficeFontFaceCollection fonts, CancellationToken cancellationToken) {
        _fonts = fonts;
        _scopedFontResolutionOnly = true;
        _cancellationToken = cancellationToken;
    }

    /// <summary>Pairs scoped shaping advances and paint bounds, sharing cancellation and
    /// outline budgets across all runs in the returned operation. Installed fonts are not resolved.</summary>
    internal static Func<string, OfficeFontInfo, (double Advance, double Left, double Top, double Right, double Bottom, bool HasInk)>
        CreateScopedPositionedTextMeasurement(OfficeFontFaceCollection fonts, CancellationToken cancellationToken) {
        var canvas = new OfficeRasterCanvas(fonts, cancellationToken);
        return (text, font) => {
            // Mathematical script levels can be below one unit. The general public text
            // measurement floor is unsuitable here; use positioned paint's 0.1-unit floor.
            double size = Math.Max(0.1D, font.Size);
            using var faceScope = canvas.PushTextFace(font.Face);
            double advance = canvas.MeasurePositionedText(text, size, font.FamilyName, font.Style,
                OfficeTextFeatureSettings.Default, OfficeTextDirection.Auto);
            var bounds = canvas.MeasurePositionedTextBounds(text, 0D, 0D, advance, 0D, size,
                font.WithSize(size), advance, OfficeTextAlignment.Left, OfficeTextFeatureSettings.Default, "normal", 0D,
                OfficeTextDecorationStyle.None, OfficeTextDecorationStyle.None);
            return (advance, bounds.Left, bounds.Top, bounds.Right, bounds.Bottom, bounds.HasInk);
        };
    }
}
