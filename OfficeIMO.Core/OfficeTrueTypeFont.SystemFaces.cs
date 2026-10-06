using System;
using System.IO;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeTrueTypeFont {
    internal int SourceByteCount => _data.Length;
    internal byte[] CopyStandaloneData(int maximumBytes) =>
        OfficeTrueTypeCollection.ExtractFace(_data, _collectionIndex ?? 0, maximumBytes);

    private OfficeTrueTypeFont? _automaticOpticalFont;
    private OfficeTrueTypeFont? _automaticBoldOpticalFont;
    internal bool HasSelectedBoldWeight => _variationModel.DesignCoordinates.TryGetValue("wght", out float weight) && weight >= 600F;

    // Installed fallback faces have no host-selected coordinates. Keep the public direct
    // font loader's nominal-instance contract, and select the authored instance at paint time.
    internal OfficeTrueTypeFont ForInstalledOpticalSize(double size, bool bold = false) {
        if (_collectionIndex.HasValue) return this;
        OfficeTrueTypeFont automatic;
        lock (_opticalSizeSync) {
            OfficeTrueTypeFont? cached = bold ? _automaticBoldOpticalFont : _automaticOpticalFont;
            if (cached == null) {
                try {
                    OfficeOpenTypeReader? reader = OfficeOpenTypeReader.TryCreate(_data);
                    OfficeFontVariationModel model = reader == null ? OfficeFontVariationModel.None
                        : OfficeFontVariationModel.Create(reader, null);
                    if (bold && reader != null && model.DesignCoordinates.ContainsKey("wght"))
                        model = OfficeFontVariationModel.Create(reader,
                            new System.Collections.Generic.Dictionary<string, float> { ["wght"] = 700F });
                    cached = model.IsVariable
                        ? TryLoad(_data, model, out _) ?? this : this;
                } catch (Exception exception) when (exception is InvalidDataException || exception is NotSupportedException
                    || exception is ArgumentException || exception is OverflowException) {
                    // A nominally usable installed face may have unsupported variation metadata.
                    cached = this;
                }
                if (bold) _automaticBoldOpticalFont = cached;
                else _automaticOpticalFont = cached;
            }
            automatic = cached;
        }
        return automatic.ForOpticalSize(size);
    }
}
