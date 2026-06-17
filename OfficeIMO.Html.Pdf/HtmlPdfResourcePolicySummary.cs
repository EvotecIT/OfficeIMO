using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using WordHtml = OfficeIMO.Word.Html;

namespace OfficeIMO.Html.Pdf;

/// <summary>
/// Source-neutral snapshot of the resource policy used by an HTML to PDF conversion profile.
/// </summary>
public sealed class HtmlPdfResourcePolicySummary {
    private HtmlPdfResourcePolicySummary() {
    }

    /// <summary>Selected HTML to PDF profile.</summary>
    public HtmlPdfProfile Profile { get; private set; }

    /// <summary>True when the profile uses the Word HTML importer and exposes its resource policy.</summary>
    public bool UsesWordHtmlPolicy { get; private set; }

    /// <summary>True when stylesheet links declared inside the HTML document may be loaded.</summary>
    public bool AllowDocumentStylesheetLinks { get; private set; }

    /// <summary>Allowed stylesheet URI schemes.</summary>
    public IReadOnlyList<string> AllowedStylesheetUriSchemes { get; private set; } = Array.Empty<string>();

    /// <summary>Allowed stylesheet hosts. Empty means all hosts are allowed after scheme checks.</summary>
    public IReadOnlyList<string> AllowedStylesheetHosts { get; private set; } = Array.Empty<string>();

    /// <summary>True when remote stylesheet content types are validated.</summary>
    public bool ValidateStylesheetContentTypes { get; private set; }

    /// <summary>Allowed stylesheet content types when validation is enabled.</summary>
    public IReadOnlyList<string> AllowedStylesheetContentTypes { get; private set; } = Array.Empty<string>();

    /// <summary>Maximum bytes allowed for one stylesheet, when configured.</summary>
    public long? MaxCssBytes { get; private set; }

    /// <summary>Maximum total stylesheet bytes allowed for one conversion, when configured.</summary>
    public long? MaxTotalCssBytes { get; private set; }

    /// <summary>Image processing mode exposed by the underlying HTML importer, when available.</summary>
    public string? ImageProcessing { get; private set; }

    /// <summary>Allowed image URI schemes.</summary>
    public IReadOnlyList<string> AllowedImageUriSchemes { get; private set; } = Array.Empty<string>();

    /// <summary>Allowed image hosts. Empty means all hosts are allowed after scheme checks.</summary>
    public IReadOnlyList<string> AllowedImageHosts { get; private set; } = Array.Empty<string>();

    /// <summary>True when declared image content types are validated.</summary>
    public bool ValidateImageContentTypes { get; private set; }

    /// <summary>Allowed image content types when validation is enabled.</summary>
    public IReadOnlyList<string> AllowedImageContentTypes { get; private set; } = Array.Empty<string>();

    /// <summary>Maximum bytes allowed for one image, when configured.</summary>
    public long? MaxImageBytes { get; private set; }

    /// <summary>Maximum total image bytes allowed for one conversion, when configured.</summary>
    public long? MaxTotalImageBytes { get; private set; }

    /// <summary>Builds a resource policy summary for the supplied options.</summary>
    public static HtmlPdfResourcePolicySummary From(HtmlPdfSaveOptions options) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        var summary = new HtmlPdfResourcePolicySummary {
            Profile = options.Profile
        };

        if (options.Profile != HtmlPdfProfile.Document) {
            return summary;
        }

        WordHtml.HtmlToWordOptions wordOptions = options.WordHtmlOptions ?? new WordHtml.HtmlToWordOptions();
        summary.UsesWordHtmlPolicy = true;
        summary.AllowDocumentStylesheetLinks = wordOptions.AllowDocumentStylesheetLinks;
        summary.AllowedStylesheetUriSchemes = CopySorted(wordOptions.AllowedStylesheetUriSchemes);
        summary.AllowedStylesheetHosts = CopySorted(wordOptions.AllowedStylesheetHosts);
        summary.ValidateStylesheetContentTypes = wordOptions.ValidateStylesheetContentTypes;
        summary.AllowedStylesheetContentTypes = CopySorted(wordOptions.AllowedStylesheetContentTypes);
        summary.MaxCssBytes = wordOptions.MaxCssBytes;
        summary.MaxTotalCssBytes = wordOptions.MaxTotalCssBytes;
        summary.ImageProcessing = wordOptions.ImageProcessing.ToString();
        summary.AllowedImageUriSchemes = CopySorted(wordOptions.AllowedImageUriSchemes);
        summary.AllowedImageHosts = CopySorted(wordOptions.AllowedImageHosts);
        summary.ValidateImageContentTypes = wordOptions.ValidateImageContentTypes;
        summary.AllowedImageContentTypes = CopySorted(wordOptions.AllowedImageContentTypes);
        summary.MaxImageBytes = wordOptions.MaxImageBytes;
        summary.MaxTotalImageBytes = wordOptions.MaxTotalImageBytes;

        return summary;
    }

    private static IReadOnlyList<string> CopySorted(IEnumerable<string> values) {
        var copy = new List<string>();
        foreach (string value in values) {
            if (!string.IsNullOrWhiteSpace(value)) {
                copy.Add(value);
            }
        }

        copy.Sort(StringComparer.OrdinalIgnoreCase);
        return new ReadOnlyCollection<string>(copy);
    }
}
