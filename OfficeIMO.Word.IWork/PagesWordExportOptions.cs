using OfficeIMO.IWork;

namespace OfficeIMO.Word.IWork;

/// <summary>Controls Pages-to-Word conversion.</summary>
public sealed class PagesWordExportOptions {
    /// <summary>Shared iWork projection settings. Null uses the defaults.</summary>
    public IWorkConversionOptions? ConversionOptions { get; set; }

    /// <summary>Throws <see cref="OfficeConversionException"/> when the conversion reports loss.</summary>
    public bool FailOnLoss { get; set; }

    /// <summary>Creates a copy of these settings; the projection settings are shared.</summary>
    public PagesWordExportOptions Clone() => new PagesWordExportOptions { ConversionOptions = ConversionOptions, FailOnLoss = FailOnLoss };
}
