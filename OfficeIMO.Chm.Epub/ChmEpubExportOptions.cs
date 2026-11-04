using OfficeIMO.Epub;

namespace OfficeIMO.Chm;

/// <summary>Controls CHM-to-EPUB conversion: topic selection and limits, manuscript import and package writing.</summary>
public sealed class ChmEpubExportOptions {
    private ChmConversionOptions _conversionOptions = new ChmConversionOptions();

    /// <summary>Topic selection, HTML and output limits shared with the other CHM converters.</summary>
    public ChmConversionOptions ConversionOptions {
        get => _conversionOptions;
        set => _conversionOptions = value ?? throw new ArgumentNullException(nameof(value));
    }

    /// <summary>Manuscript import settings such as title and language. Null uses the defaults.</summary>
    public EpubManuscriptOptions? EpubOptions { get; set; }

    /// <summary>Package-write settings used by <c>SaveAsEpub</c>. Output limits are intersected with <see cref="ConversionOptions"/>.</summary>
    public EpubSaveOptions? SaveOptions { get; set; }

    /// <summary>Throws <see cref="OfficeConversionException"/> when the conversion reports loss.</summary>
    public bool FailOnLoss { get; set; }

    /// <summary>Creates an independent copy of these settings.</summary>
    public ChmEpubExportOptions Clone() => new ChmEpubExportOptions {
        ConversionOptions = ConversionOptions.Clone(),
        EpubOptions = EpubOptions?.Clone(),
        SaveOptions = SaveOptions,
        FailOnLoss = FailOnLoss
    };
}
