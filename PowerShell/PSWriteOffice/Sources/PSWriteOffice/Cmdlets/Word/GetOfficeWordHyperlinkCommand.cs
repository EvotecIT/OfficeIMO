using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Gets hyperlinks from a Word document.</summary>
/// <example>
///   <summary>Audit external hyperlinks in a report.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$links = Get-OfficeWordHyperlink -Path .\Report.docx -Url 'https://*'
/// $links |
///     Select-Object -Property Text, Uri |
///     Export-Csv -Path .\ReportLinks.csv -NoTypeInformation</code>
///   <para>Filters external hyperlinks and exports them for link validation or review.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeWordHyperlink", DefaultParameterSetName = ParameterSetPath)]
[OutputType(typeof(WordHyperLink))]
public sealed class GetOfficeWordHyperlinkCommand : PSCmdlet {
    private const string ParameterSetPath = "Path";
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetSection = "Section";
    private const string ParameterSetParagraph = "Paragraph";

    /// <summary>Path to the document.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Document to inspect.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public WordDocument Document { get; set; } = null!;

    /// <summary>Section to inspect.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetSection)]
    public WordSection Section { get; set; } = null!;

    /// <summary>Paragraph to inspect.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetParagraph)]
    public WordParagraph Paragraph { get; set; } = null!;

    /// <summary>Filter by hyperlink text (wildcards supported).</summary>
    [Parameter]
    [SupportsWildcards]
    public string[]? Text { get; set; }

    /// <summary>Filter by external URL (wildcards supported).</summary>
    [Parameter]
    [SupportsWildcards]
    [Alias("Uri")]
    public string[]? Url { get; set; }

    /// <summary>Filter by bookmark anchor (wildcards supported).</summary>
    [Parameter]
    [SupportsWildcards]
    [Alias("Bookmark")]
    public string[]? Anchor { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        WordDocument? document = null;
        var dispose = false;

        try {
            IEnumerable<WordHyperLink> hyperlinks = ResolveHyperlinks(ref document, ref dispose);

            hyperlinks = FilterByPatterns(hyperlinks, Text, hyperlink => hyperlink.Text);
            hyperlinks = FilterByPatterns(hyperlinks, Url, hyperlink => hyperlink.Uri?.ToString());
            hyperlinks = FilterByPatterns(hyperlinks, Anchor, hyperlink => hyperlink.Anchor);

            WriteObject(hyperlinks, enumerateCollection: true);
        } finally {
            if (dispose) {
                document?.Dispose();
            }
        }
    }

    private IEnumerable<WordHyperLink> ResolveHyperlinks(ref WordDocument? document, ref bool dispose) {
        switch (ParameterSetName) {
            case ParameterSetParagraph:
                return Paragraph.IsHyperLink && Paragraph.Hyperlink != null
                    ? new[] { Paragraph.Hyperlink }
                    : Array.Empty<WordHyperLink>();
            case ParameterSetSection:
                return Section != null ? Section.HyperLinks : Array.Empty<WordHyperLink>();
            case ParameterSetDocument:
                return Document != null ? Document.HyperLinks : Array.Empty<WordHyperLink>();
            default:
                var resolvedPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
                document = WordDocumentService.LoadDocument(resolvedPath, readOnly: true, autoSave: false);
                dispose = true;
                return document.HyperLinks;
        }
    }

    private static IEnumerable<WordHyperLink> FilterByPatterns(
        IEnumerable<WordHyperLink> hyperlinks,
        string[]? patterns,
        Func<WordHyperLink, string?> valueSelector) {
        var compiledPatterns = BuildPatterns(patterns);
        if (compiledPatterns.Count == 0) {
            return hyperlinks;
        }

        return hyperlinks.Where(hyperlink => {
            var value = valueSelector(hyperlink);
            return value != null && compiledPatterns.Any(pattern => pattern.IsMatch(value));
        });
    }

    private static List<WildcardPattern> BuildPatterns(string[]? patterns) {
        var compiled = new List<WildcardPattern>();
        foreach (var pattern in patterns ?? Array.Empty<string>()) {
            if (!string.IsNullOrWhiteSpace(pattern)) {
                compiled.Add(new WildcardPattern(pattern, WildcardOptions.IgnoreCase));
            }
        }

        return compiled;
    }
}