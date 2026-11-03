using OfficeIMO.Markdown;
using OfficeIMO.Rtf;

namespace OfficeIMO.Rtf.Markdown;

/// <summary>Controls semantic conversion from an OfficeIMO RTF document into Markdown.</summary>
public sealed class RtfMarkdownExportOptions {
    /// <summary>Markdown writer options used when rendering the resulting Markdown document.</summary>
    public MarkdownSaveOptions? MarkdownWriteOptions { get; set; }

    /// <summary>Markdown reader options used when building inline Markdown sequences.</summary>
    public MarkdownLoadOptions? InlineReaderOptions { get; set; }

    /// <summary>Includes RTF hidden text in generated Markdown.</summary>
    public bool IncludeHiddenText { get; set; }

    /// <summary>Emits HTML comments for unsupported RTF block features.</summary>
    public bool EmitUnsupportedHtmlComments { get; set; } = true;

    /// <summary>Creates an image path for each RTF image encountered during conversion.</summary>
    public Func<RtfImage, int, string>? ImagePathFactory { get; set; }

    /// <summary>Exports each RTF image payload to the path selected by <see cref="ImagePathFactory"/>.</summary>
    public Action<RtfImage, int, string>? ImageExporter { get; set; }

    /// <summary>Encoding used when the Markdown is saved. Null writes UTF-8 without a byte-order mark.</summary>
    public System.Text.Encoding? Encoding { get; set; }

    /// <summary>Throws <see cref="OfficeConversionException"/> when the conversion reports loss.</summary>
    public bool FailOnLoss { get; set; }

    /// <summary>Creates a copy of these settings; callbacks and nested writer settings are shared.</summary>
    public RtfMarkdownExportOptions Clone() => (RtfMarkdownExportOptions)MemberwiseClone();
}
