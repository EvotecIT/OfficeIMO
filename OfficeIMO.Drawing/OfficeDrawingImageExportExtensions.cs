using System;

namespace OfficeIMO.Drawing;

/// <summary>Applies shared image-export configuration to Drawing scenes.</summary>
public static class OfficeDrawingImageExportExtensions {
    /// <summary>Adds caller-supplied font faces from shared export options to a drawing.</summary>
    public static OfficeDrawing ApplyImageExportOptions(
        this OfficeDrawing drawing,
        OfficeImageExportOptions options) {
        if (drawing == null) throw new ArgumentNullException(nameof(drawing));
        if (options == null) throw new ArgumentNullException(nameof(options));
        drawing.Fonts.AddRange(options.Fonts);
        return drawing;
    }
}
