using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeIMO.Reader;

/// <summary>Controls page-aware Markdown projection from a normalized read result.</summary>
public sealed class OfficeDocumentPageMarkdownOptions {
    /// <summary>When true, prefixes each page with a portable HTML page marker.</summary>
    public bool IncludePageMarkers { get; set; } = true;
}

/// <summary>Markdown content and citation metadata for one page-like container.</summary>
public sealed class OfficeDocumentPageMarkdown {
    internal OfficeDocumentPageMarkdown(
        OfficeDocumentPage page,
        int totalPageCount,
        OfficeDocumentPageProvenance provenance,
        string markdown) {
        Page = page;
        TotalPageCount = totalPageCount;
        Provenance = provenance;
        Markdown = markdown;
    }

    /// <summary>Source page-like container.</summary>
    public OfficeDocumentPage Page { get; }

    /// <summary>Total physical page count known for the read operation.</summary>
    public int TotalPageCount { get; }

    /// <summary>How this page boundary was obtained.</summary>
    public OfficeDocumentPageProvenance Provenance { get; }

    /// <summary>Portable page-scoped Markdown.</summary>
    public string Markdown { get; }
}

public static partial class OfficeDocumentReadResultExtensions {
    /// <summary>
    /// Projects every page-like container into a separate Markdown value suitable for citations and RAG.
    /// </summary>
    public static IReadOnlyList<OfficeDocumentPageMarkdown> GetPageMarkdown(
        this OfficeDocumentReadResult document,
        OfficeDocumentPageMarkdownOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        OfficeDocumentPageMarkdownOptions effective = options ?? new OfficeDocumentPageMarkdownOptions();
        int total = document.GetTotalPageCount();
        OfficeDocumentPageProvenance provenance = document.GetPageProvenance();
        var pages = new List<OfficeDocumentPageMarkdown>();

        foreach (OfficeDocumentPage page in document.Pages ?? Array.Empty<OfficeDocumentPage>()) {
            var markdown = new StringBuilder();
            if (effective.IncludePageMarkers) {
                markdown.Append("<!-- page: ");
                markdown.Append(page.Number?.ToString(CultureInfo.InvariantCulture) ?? "?");
                if (total > 0) {
                    markdown.Append('/');
                    markdown.Append(total.ToString(CultureInfo.InvariantCulture));
                }
                markdown.Append("; provenance: ");
                markdown.Append(provenance);
                markdown.AppendLine(" -->");
                markdown.AppendLine();
            }

            var emitted = new HashSet<string>(StringComparer.Ordinal);
            foreach (OfficeDocumentBlock block in page.Blocks ?? Array.Empty<OfficeDocumentBlock>()) {
                if (!string.IsNullOrEmpty(block.Id) && !emitted.Add(block.Id)) {
                    continue;
                }
                AppendBlockMarkdown(markdown, block);
            }

            pages.Add(new OfficeDocumentPageMarkdown(
                page,
                total,
                provenance,
                markdown.ToString().TrimEnd()));
        }

        return pages.AsReadOnly();
    }

    /// <summary>
    /// Projects all page-like containers into one Markdown string separated by portable page markers.
    /// </summary>
    public static string ToPageMarkedMarkdown(this OfficeDocumentReadResult document) =>
        string.Join(
            Environment.NewLine + Environment.NewLine,
            document.GetPageMarkdown(new OfficeDocumentPageMarkdownOptions { IncludePageMarkers = true })
                .Select(static page => page.Markdown));

    private static void AppendBlockMarkdown(StringBuilder markdown, OfficeDocumentBlock block) {
        string text = block.Text ?? string.Empty;
        switch (block.Kind) {
            case "heading":
                int level = Math.Max(1, Math.Min(6, block.Level ?? 1));
                markdown.Append(new string('#', level));
                markdown.Append(' ');
                markdown.AppendLine(text);
                break;
            case "list-item":
                markdown.Append(string.IsNullOrWhiteSpace(block.Marker) ? "- " : block.Marker + " ");
                markdown.AppendLine(text);
                break;
            default:
                markdown.AppendLine(text);
                break;
        }
        markdown.AppendLine();
    }
}
