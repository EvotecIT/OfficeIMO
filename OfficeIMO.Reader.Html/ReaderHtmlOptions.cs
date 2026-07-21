using OfficeIMO.Html;
using OfficeIMO.Markdown.Html;
using System;

namespace OfficeIMO.Reader.Html;

/// <summary>
/// Options for HTML ingestion pipeline (HTML -> Markdown).
/// </summary>
public sealed class ReaderHtmlOptions {
    /// <summary>Starts a new chunk at Markdown headings when possible.</summary>
    public bool ChunkByHeadings { get; set; } = true;

    /// <summary>
    /// Creates the default OfficeIMO HTML reader profile.
    /// </summary>
    /// <returns>A new <see cref="ReaderHtmlOptions"/> instance using the default HTML-to-Markdown profile.</returns>
    public static ReaderHtmlOptions CreateOfficeIMOProfile() => new ReaderHtmlOptions {
        HtmlToMarkdownOptions = HtmlToMarkdownOptions.CreateOfficeIMOProfile()
    };

    /// <summary>
    /// Creates a portable HTML reader profile that favors portable Markdown output.
    /// </summary>
    /// <returns>A new <see cref="ReaderHtmlOptions"/> instance using the portable HTML-to-Markdown profile.</returns>
    public static ReaderHtmlOptions CreatePortableProfile() => new ReaderHtmlOptions {
        HtmlToMarkdownOptions = HtmlToMarkdownOptions.CreatePortableProfile()
    };

    /// <summary>
    /// Creates a bounded HTML reader profile for untrusted or size-sensitive HTML ingestion.
    /// </summary>
    /// <param name="maxInputCharacters">Maximum HTML input length accepted by the HTML-to-Markdown stage.</param>
    /// <returns>A new <see cref="ReaderHtmlOptions"/> instance with a configured HTML input character limit.</returns>
    public static ReaderHtmlOptions CreateUntrustedHtmlProfile(int maxInputCharacters) {
        if (maxInputCharacters <= 0) {
            throw new ArgumentOutOfRangeException(nameof(maxInputCharacters), "Maximum input characters must be greater than zero.");
        }

        var options = HtmlToMarkdownOptions.CreatePortableProfile();
        options.MaxInputCharacters = maxInputCharacters;
        HtmlConversionDocumentOptions conversionOptions = HtmlConversionDocumentOptions.CreateUntrustedProfile();
        conversionOptions.Limits.MaxInputCharacters = maxInputCharacters;

        return new ReaderHtmlOptions {
            ConversionOptions = conversionOptions,
            HtmlToMarkdownOptions = options
        };
    }

    /// <summary>
    /// Shared parsing, trust, URL-policy, and complexity options applied before the reader projects HTML.
    /// </summary>
    public HtmlConversionDocumentOptions? ConversionOptions { get; set; }

    /// <summary>
    /// Options passed to HTML-to-Markdown conversion stage.
    /// </summary>
    public HtmlToMarkdownOptions? HtmlToMarkdownOptions { get; set; }

    /// <summary>
    /// Creates a copy of the current options instance so reader registrations can reuse templates safely.
    /// </summary>
    /// <returns>A new <see cref="ReaderHtmlOptions"/> with cloned nested HTML-to-Markdown options.</returns>
    public ReaderHtmlOptions Clone() => new ReaderHtmlOptions {
        ConversionOptions = ConversionOptions?.Clone(),
        HtmlToMarkdownOptions = HtmlToMarkdownOptions?.Clone(),
        ChunkByHeadings = ChunkByHeadings
    };
}
