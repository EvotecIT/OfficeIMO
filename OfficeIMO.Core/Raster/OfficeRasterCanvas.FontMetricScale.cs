using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    internal double FontMetricScale { get; set; } = 1D;
    private Dictionary<OfficeTrueTypeFont, OfficeTrueTypeFont>? _scaledMetricFonts;

    private IOfficeFontProgram? ResolveMetricScale(IOfficeFontProgram? font) {
        if (FontMetricScale == 1D || font is not OfficeTrueTypeFont trueType) return font;
        var cache = _scaledMetricFonts ??= new Dictionary<OfficeTrueTypeFont, OfficeTrueTypeFont>();
        if (cache.TryGetValue(trueType, out OfficeTrueTypeFont? instance)) return instance;
        if (cache.Count >= 1024) cache.Clear();
        instance = trueType.ForOutputScale(FontMetricScale);
        cache[trueType] = instance;
        return instance;
    }
}
