using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeFontFace {
    private readonly object _opticalSizeSync = new();
    private Dictionary<IOfficeFontProgram, OfficeFontFace>? _opticalSizeFaces;

    internal OfficeFontFace ForOpticalSize(double size) {
        if (!_automaticOpticalSizing || ParsedFont is not OfficeTrueTypeFont trueType) return this;
        OfficeTrueTypeFont selected = trueType.ForOpticalSize(size);
        if (ReferenceEquals(selected, trueType)) return this;
        lock (_opticalSizeSync) {
            var cache = _opticalSizeFaces ??= new Dictionary<IOfficeFontProgram, OfficeFontFace>();
            if (cache.TryGetValue(selected, out OfficeFontFace? face)) return face;
            face = new OfficeFontFace(FamilyName, ResourceFamilyName, _data, Style, Descriptor,
                UnicodeRanges, selected, ContainerFormat, canEmbedAsStaticPdfFont: false, useDataSnapshot: true);
            if (cache.Count >= 16) cache.Clear();
            cache.Add(selected, face);
            return face;
        }
    }
}
