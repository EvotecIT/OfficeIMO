using OfficeIMO.Chm;
using OfficeIMO.Reader.Html;

namespace OfficeIMO.Reader.Chm;

/// <summary>Bounds archive ingestion and selected-topic projection for a CHM Reader registration.</summary>
public sealed class ReaderChmOptions {
    /// <summary>Archive, metadata and per-topic HTML limits.</summary>
    public ChmReadOptions ReadOptions { get; set; } = new ChmReadOptions();
    /// <summary>Topic selection and aggregate HTML budgets.</summary>
    public ChmConversionOptions ConversionOptions { get; set; } = new ChmConversionOptions();
    /// <summary>Options for the canonical rich HTML reader used on each topic.</summary>
    public ReaderHtmlOptions HtmlOptions { get; set; } = ReaderHtmlOptions.CreateOfficeIMOProfile();
    /// <summary>Creates a detached registration snapshot.</summary>
    public ReaderChmOptions Clone() => new ReaderChmOptions {
        ReadOptions = (ReadOptions ?? throw new InvalidOperationException("CHM read options are required.")).Clone(),
        ConversionOptions = (ConversionOptions ?? throw new InvalidOperationException("CHM conversion options are required.")).Clone(),
        HtmlOptions = (HtmlOptions ?? throw new InvalidOperationException("CHM HTML reader options are required.")).Clone()
    };
}
