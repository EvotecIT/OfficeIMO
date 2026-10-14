using System;
using System.Collections.Generic;
using System.IO;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeTrueTypeFont {
    private readonly object _opticalSizeSync = new();
    private Dictionary<double, OfficeTrueTypeFont>? _opticalSizeFonts;

    internal OfficeTrueTypeFont ForOpticalSize(double authoredSize) {
        if (!_variationModel.TryGetOpticalSize(authoredSize, out double size) ||
            _variationModel.IsSelectedOpticalSize(size)) return this;
        lock (_opticalSizeSync) {
            var cache = _opticalSizeFonts ??= new Dictionary<double, OfficeTrueTypeFont>();
            if (cache.TryGetValue(size, out OfficeTrueTypeFont? instance)) return instance;
            OfficeOpenTypeReader reader = OfficeOpenTypeReader.TryCreate(_data)
                ?? throw new InvalidDataException("The variable font directory is invalid.");
            OfficeFontVariationModel model = _variationModel.ForOpticalSize(reader, authoredSize);
            instance = TryLoad(_data, model, out string? error)
                ?? throw new InvalidDataException(error ?? "The optical font instance could not be loaded.");
            instance._trackingEvaluationScale = _trackingEvaluationScale;
            if (cache.Count >= 16) cache.Clear();
            cache.Add(size, instance);
            return instance;
        }
    }
}
