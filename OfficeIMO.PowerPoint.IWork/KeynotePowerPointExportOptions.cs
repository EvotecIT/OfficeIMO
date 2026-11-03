using OfficeIMO.IWork;

namespace OfficeIMO.PowerPoint.IWork;

/// <summary>Controls Keynote-to-PowerPoint conversion.</summary>
public sealed class KeynotePowerPointExportOptions {
    /// <summary>Shared iWork projection settings. Null uses the defaults.</summary>
    public IWorkConversionOptions? ConversionOptions { get; set; }

    /// <summary>Throws <see cref="OfficeConversionException"/> when the conversion reports loss.</summary>
    public bool FailOnLoss { get; set; }

    /// <summary>Creates a copy of these settings; the projection settings are shared.</summary>
    public KeynotePowerPointExportOptions Clone() => new KeynotePowerPointExportOptions { ConversionOptions = ConversionOptions, FailOnLoss = FailOnLoss };
}
