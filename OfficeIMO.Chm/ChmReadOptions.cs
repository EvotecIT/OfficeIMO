using OfficeIMO.Html.Dom;

namespace OfficeIMO.Chm;

/// <summary>Finite budgets and inert HTML services for loading a compiled help book.</summary>
public sealed class ChmReadOptions {
    /// <summary>Maximum source archive bytes. Seekable input is read from the beginning and its position restored.</summary>
    public long MaxInputBytes { get; set; } = 128L * 1024 * 1024;
    /// <summary>Maximum expanded compressed section bytes, before allocation or decoding.</summary>
    public long MaxExpandedBytes { get; set; } = 256L * 1024 * 1024;
    /// <summary>Maximum directory entries, including internal streams.</summary>
    public int MaxEntries { get; set; } = 100_000;
    /// <summary>Maximum bytes in any individual entry.</summary>
    public int MaxEntryBytes { get; set; } = 32 * 1024 * 1024;
    /// <summary>Maximum entry-name bytes, reference characters and compiled navigation string bytes.</summary>
    public int MaxPathLength { get; set; } = 4096;
    /// <summary>Maximum contents and index items combined, and maximum compiled topic-table records.</summary>
    public int MaxNavigationItems { get; set; } = 100_000;
    /// <summary>Maximum contents or index nesting depth, from 1 through 512.</summary>
    public int MaxNavigationDepth { get; set; } = 128;
    /// <summary>Maximum nodes in one parsed sitemap or topic.</summary>
    public int MaxHtmlNodes { get; set; } = 200_000;
    /// <summary>Maximum element nesting depth in one parsed HTML document, from 1 through 512.</summary>
    public int MaxHtmlDepth { get; set; } = 256;
    /// <summary>Optional text encoding override. Otherwise BOM/meta declarations precede the help-book locale.</summary>
    public Encoding? TextEncoding { get; set; }
    /// <summary>Inert parser used for sitemap and topic analysis. No scripts or resources are executed.</summary>
    public IHtmlParserProvider ParserProvider { get; set; } = HtmlDocumentEngine.Default.ParserProvider;
    /// <summary>Charset provider used to resolve legacy HTML and help-book encodings.</summary>
    public IHtmlEncodingProvider EncodingProvider { get; set; } = new HtmlConversionDocumentOptions().InputEncodingProvider;

    /// <summary>Returns a detached settings snapshot.</summary>
    public ChmReadOptions Clone() {
        var copy = (ChmReadOptions)MemberwiseClone();
        copy.TextEncoding = TextEncoding == null ? null : (Encoding)TextEncoding.Clone();
        return copy;
    }

    internal void Validate() {
        if (MaxInputBytes < 1 || MaxInputBytes > int.MaxValue) throw new ArgumentOutOfRangeException(nameof(MaxInputBytes));
        if (MaxExpandedBytes < 1 || MaxExpandedBytes > int.MaxValue) throw new ArgumentOutOfRangeException(nameof(MaxExpandedBytes));
        if (MaxEntries < 1) throw new ArgumentOutOfRangeException(nameof(MaxEntries));
        if (MaxEntryBytes < 1) throw new ArgumentOutOfRangeException(nameof(MaxEntryBytes));
        if (MaxPathLength < 1) throw new ArgumentOutOfRangeException(nameof(MaxPathLength));
        if (MaxNavigationItems < 1) throw new ArgumentOutOfRangeException(nameof(MaxNavigationItems));
        if (MaxNavigationDepth < 1 || MaxNavigationDepth > 512) throw new ArgumentOutOfRangeException(nameof(MaxNavigationDepth));
        if (MaxHtmlNodes < 1) throw new ArgumentOutOfRangeException(nameof(MaxHtmlNodes));
        if (MaxHtmlDepth < 1 || MaxHtmlDepth > 512) throw new ArgumentOutOfRangeException(nameof(MaxHtmlDepth));
        if (ParserProvider == null) throw new ArgumentNullException(nameof(ParserProvider));
        if (EncodingProvider == null) throw new ArgumentNullException(nameof(EncodingProvider));
    }
}
