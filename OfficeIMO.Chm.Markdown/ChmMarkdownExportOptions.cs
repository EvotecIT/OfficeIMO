using OfficeIMO.Markdown.Html;

namespace OfficeIMO.Chm;

/// <summary>Controls CHM-to-Markdown conversion: topic selection and limits, and the HTML-to-Markdown projection.</summary>
public sealed class ChmMarkdownExportOptions {
    private ChmConversionOptions _conversionOptions = new ChmConversionOptions();

    /// <summary>Topic selection, HTML and output limits shared with the other CHM converters.</summary>
    public ChmConversionOptions ConversionOptions {
        get => _conversionOptions;
        set => _conversionOptions = value ?? throw new ArgumentNullException(nameof(value));
    }

    /// <summary>HTML-to-Markdown settings. Null uses the portable profile.</summary>
    public HtmlMarkdownExportOptions? MarkdownOptions { get; set; }

    /// <summary>Throws <see cref="OfficeConversionException"/> when the conversion reports loss.</summary>
    public bool FailOnLoss { get; set; }

    /// <summary>Creates an independent copy of these settings.</summary>
    public ChmMarkdownExportOptions Clone() => new ChmMarkdownExportOptions {
        ConversionOptions = ConversionOptions.Clone(),
        MarkdownOptions = MarkdownOptions?.Clone(),
        FailOnLoss = FailOnLoss
    };
}
