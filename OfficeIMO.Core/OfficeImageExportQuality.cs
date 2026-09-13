using System;

namespace OfficeIMO.Drawing;

/// <summary>Output-density presets. Vector geometry remains resolution independent; raster detail is rendered at the selected density.</summary>
public enum OfficeImageExportQuality {
    /// <summary>96 DPI for small previews.</summary>
    Preview,
    /// <summary>192 DPI for high-density screens.</summary>
    Screen,
    /// <summary>300 DPI for detailed reports and printing.</summary>
    Print
}

/// <summary>Consistent density presets for all document image-export options.</summary>
public static class OfficeImageExportQualityExtensions {
    /// <summary>Sets target density without changing fonts, format, layout, or safety limits. Clear TargetDpi to use Scale directly, or use the fluent WithScale method.</summary>
    public static T UseQuality<T>(this T options, OfficeImageExportQuality quality) where T : OfficeImageExportOptions {
        if (options == null) throw new ArgumentNullException(nameof(options));
        options.TargetDpi = quality switch {
            OfficeImageExportQuality.Preview => 96D,
            OfficeImageExportQuality.Screen => 192D,
            OfficeImageExportQuality.Print => 300D,
            _ => throw new ArgumentOutOfRangeException(nameof(quality))
        };
        return options;
    }
}
