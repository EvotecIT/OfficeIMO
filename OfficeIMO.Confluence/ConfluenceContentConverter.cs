using OfficeIMO.Adf;
using OfficeIMO.Markdown;

namespace OfficeIMO.Confluence;

/// <summary>A converted Confluence body and its ADF fidelity evidence.</summary>
public sealed class ConfluenceContentConversionResult<T> {
    internal ConfluenceContentConversionResult(T value, AdfConversionReport report) {
        Value = value;
        Report = report;
    }
    public T Value { get; }
    public AdfConversionReport Report { get; }
}

/// <summary>Creates and projects Confluence page bodies using OfficeIMO's ADF, Markdown, and HTML engines.</summary>
public static class ConfluenceContentConverter {
    /// <summary>Creates a Confluence ADF body from Markdown.</summary>
    public static ConfluenceContentConversionResult<ConfluencePageBody> FromMarkdown(string markdown, ConfluenceBodyFormat format = ConfluenceBodyFormat.AtlasDocFormat) {
        if (markdown == null) throw new ArgumentNullException(nameof(markdown));
        AdfConversionResult<AdfDocument> adf = AdfConverter.FromMarkdown(markdown);
        string value = format == ConfluenceBodyFormat.AtlasDocFormat ? adf.Value.ToJson() : MarkdownReader.Parse(markdown).ToHtmlFragment();
        AdfConversionReport report = format == ConfluenceBodyFormat.AtlasDocFormat
            ? adf.Report
            : new AdfConversionReport(new[] {
                new AdfConversionDiagnostic(
                    "CONFLUENCE_MARKDOWN_STORAGE_PIPELINE",
                    "$",
                    "Markdown is rendered directly to Confluence storage HTML; no ADF conversion is performed.",
                    AdfConversionSeverity.Information),
            });
        return new ConfluenceContentConversionResult<ConfluencePageBody>(new ConfluencePageBody {
            Representation = format == ConfluenceBodyFormat.AtlasDocFormat ? "atlas_doc_format" : "storage",
            Value = value,
        }, report);
    }

    /// <summary>Creates a Confluence body from HTML. ADF output passes through OfficeIMO.Html and Markdown.</summary>
    public static ConfluenceContentConversionResult<ConfluencePageBody> FromHtml(string html, ConfluenceBodyFormat format = ConfluenceBodyFormat.Storage) {
        if (html == null) throw new ArgumentNullException(nameof(html));
        AdfConversionResult<AdfDocument>? adf = format == ConfluenceBodyFormat.AtlasDocFormat ? AdfConverter.FromHtml(html) : null;
        return new ConfluenceContentConversionResult<ConfluencePageBody>(new ConfluencePageBody {
            Representation = format == ConfluenceBodyFormat.AtlasDocFormat ? "atlas_doc_format" : "storage",
            Value = format == ConfluenceBodyFormat.AtlasDocFormat ? adf!.Value.ToJson() : html,
        }, adf?.Report ?? AdfConversionReport.Empty);
    }

    /// <summary>Projects a page body to Markdown with an explicit fidelity report.</summary>
    public static ConfluenceContentConversionResult<string> ToMarkdown(ConfluencePage page) {
        if (page == null) throw new ArgumentNullException(nameof(page));
        if (!string.IsNullOrWhiteSpace(page.Body.AtlasDocFormat?.Value)) {
            AdfConversionResult<string> result = AdfConverter.ToMarkdown(AdfDocument.Parse(page.Body.AtlasDocFormat!.Value!));
            return new ConfluenceContentConversionResult<string>(result.Value, result.Report);
        }
        if (!string.IsNullOrWhiteSpace(page.Body.Storage?.Value)) {
            AdfConversionResult<AdfDocument> adf = AdfConverter.FromHtml(page.Body.Storage!.Value!);
            AdfConversionResult<string> result = AdfConverter.ToMarkdown(adf.Value);
            return new ConfluenceContentConversionResult<string>(result.Value, Combine(adf.Report, result.Report));
        }
        throw new InvalidOperationException("The Confluence page does not contain an ADF or storage body.");
    }

    /// <summary>Projects a page body to HTML with an explicit fidelity report.</summary>
    public static ConfluenceContentConversionResult<string> ToHtml(ConfluencePage page) {
        if (page == null) throw new ArgumentNullException(nameof(page));
        string? storage = page.Body?.Storage?.Value;
        if (!string.IsNullOrWhiteSpace(storage)) {
            return new ConfluenceContentConversionResult<string>(storage!, AdfConversionReport.Empty);
        }
        string? atlasDocFormat = page.Body?.AtlasDocFormat?.Value;
        if (!string.IsNullOrWhiteSpace(atlasDocFormat)) {
            AdfConversionResult<string> result = AdfConverter.ToHtml(AdfDocument.Parse(atlasDocFormat!));
            return new ConfluenceContentConversionResult<string>(result.Value, result.Report);
        }
        throw new InvalidOperationException("The Confluence page does not contain an ADF or storage body.");
    }

    private static AdfConversionReport Combine(AdfConversionReport first, AdfConversionReport second) {
        return new AdfConversionReport(first.Diagnostics.Concat(second.Diagnostics));
    }
}
