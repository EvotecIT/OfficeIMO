using System;
using System.Collections.Generic;
using System.IO;
using System.Management.Automation;
using System.Text.RegularExpressions;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Finds text matches inside a Word document.</summary>
/// <para>Returns matching paragraphs or a WordSearchResult when using regex with <c>-AsResult</c>.</para>
/// <example>
///   <summary>Find matching paragraphs and export a review list.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$matches = Find-OfficeWord -Path .\Report.docx -Text 'Quarter'
/// $matches |
///     Select-Object -Property Text |
///     Export-Csv -Path .\ReviewMatches.csv -NoTypeInformation</code>
///   <para>Finds paragraphs containing the search text and writes a lightweight review CSV.</para>
/// </example>
/// <example>
///   <summary>Use regex and keep the grouped WordSearchResult.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$result = Find-OfficeWord -Path .\Report.docx -Pattern 'FY\d{4}' -AsResult
/// $result.Paragraphs |
///     Select-Object -Property Text</code>
///   <para>Returns the full regex result so paragraphs, table matches, headers, and footers can be inspected separately.</para>
/// </example>
[Cmdlet(VerbsCommon.Find, "OfficeWord", DefaultParameterSetName = ParameterSetPathText)]
[OutputType(typeof(WordParagraph), typeof(WordSearchResult))]
public sealed class FindOfficeWordCommand : PSCmdlet {
    private const string ParameterSetPathText = "PathText";
    private const string ParameterSetPathRegex = "PathRegex";
    private const string ParameterSetDocumentText = "DocumentText";
    private const string ParameterSetDocumentRegex = "DocumentRegex";

    /// <summary>Path to the .docx file.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPathText)]
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPathRegex)]
    [Alias("InputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Word document to search.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocumentText)]
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocumentRegex)]
    public WordDocument Document { get; set; } = null!;

    /// <summary>Text to find.</summary>
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = ParameterSetPathText)]
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = ParameterSetDocumentText)]
    public string Text { get; set; } = string.Empty;

    /// <summary>Regular expression pattern to find.</summary>
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = ParameterSetPathRegex)]
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = ParameterSetDocumentRegex)]
    public string Pattern { get; set; } = string.Empty;

    /// <summary>Use case-sensitive matching.</summary>
    [Parameter]
    public SwitchParameter CaseSensitive { get; set; }

    /// <summary>Emit the full <see cref="WordSearchResult"/> result for regex searches.</summary>
    [Parameter(ParameterSetName = ParameterSetPathRegex)]
    [Parameter(ParameterSetName = ParameterSetDocumentRegex)]
    public SwitchParameter AsResult { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        WordDocument? document = null;
        var dispose = false;

        try {
            if (ParameterSetName == ParameterSetPathText || ParameterSetName == ParameterSetPathRegex) {
                var resolvedPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
                document = WordDocumentService.LoadDocument(resolvedPath, readOnly: true, autoSave: false);
                dispose = true;
            } else {
                document = Document;
            }

            if (document == null) {
                throw new InvalidOperationException("Word document was not provided.");
            }

            if (ParameterSetName == ParameterSetPathText || ParameterSetName == ParameterSetDocumentText) {
                var comparison = CaseSensitive.IsPresent ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
                var results = document.Find(Text, comparison);
                WriteObject(results, enumerateCollection: true);
                return;
            }

            var regexOptions = CaseSensitive.IsPresent ? RegexOptions.None : RegexOptions.IgnoreCase;
            var regex = new Regex(Pattern, regexOptions);
            var result = document.Find(regex);
            if (AsResult.IsPresent) {
                WriteObject(result);
            } else {
                WriteObject(Flatten(result), enumerateCollection: true);
            }
        } finally {
            if (dispose) {
                document?.Dispose();
            }
        }
    }

    private static IEnumerable<WordParagraph> Flatten(WordSearchResult result) {
        foreach (var paragraph in result.BodyParagraphs) yield return paragraph;
        foreach (var paragraph in result.TableParagraphs) yield return paragraph;
        foreach (var paragraph in result.DefaultHeaderParagraphs) yield return paragraph;
        foreach (var paragraph in result.EvenPageHeaderParagraphs) yield return paragraph;
        foreach (var paragraph in result.FirstPageHeaderParagraphs) yield return paragraph;
        foreach (var paragraph in result.DefaultFooterParagraphs) yield return paragraph;
        foreach (var paragraph in result.EvenPageFooterParagraphs) yield return paragraph;
        foreach (var paragraph in result.FirstPageFooterParagraphs) yield return paragraph;
    }
}