using System;
using System.Collections.Generic;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Finds Word tables containing matching cell text.</summary>
/// <para>
/// Searches table cell paragraphs in a Word document and returns the matching <see cref="WordTable"/>
/// objects. Use this when a document came from a template or another system and the script needs to
/// locate a table by visible marker text before appending rows, changing cells, or applying table-cell
/// formatting.
/// </para>
/// <para>
/// By default only top-level tables are searched. Use <c>-IncludeNested</c> to include tables inside
/// table cells. Use <c>-Text</c> for literal contains matching or <c>-Pattern</c> for regular
/// expressions.
/// </para>
/// <example>
///   <summary>Find the table that contains a marker and append a row.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$doc = Get-OfficeWord -Path .\Report.docx
/// $table = Find-OfficeWordTable -Document $doc -Text 'Risk register' | Select-Object -First 1
/// $table | Add-OfficeWordTableRow -Values 'Contoso', 'Open', 'High'</code>
///   <para>Searches table cell paragraphs and returns the matching OfficeIMO table object.</para>
/// </example>
/// <example>
///   <summary>Find a risk table, append a row, and update a status cell.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$doc = Get-OfficeWord -Path .\Handover.docx
/// $table = Find-OfficeWordTable -Document $doc -Text 'Risk marker' | Select-Object -First 1
/// $table | Add-OfficeWordTableRow -Values 'Mitigation plan', 'Service Desk', 'Ready'
/// $table | Get-OfficeWordTableCell -Row 2 -Column 2 |
///     Set-OfficeWordTableCell -Text 'Investigating' -ShadingFillColor '#fff2cc' -ShadingPattern Clear
/// $doc | Close-OfficeWord -Save</code>
///   <para>Shows the common existing-document workflow: locate the table, mutate it, then save the open document.</para>
/// </example>
[Cmdlet(VerbsCommon.Find, "OfficeWordTable", DefaultParameterSetName = ParameterSetPathText)]
[OutputType(typeof(WordTable))]
public sealed class FindOfficeWordTableCommand : PSCmdlet {
    private const string ParameterSetPathText = "PathText";
    private const string ParameterSetPathRegex = "PathRegex";
    private const string ParameterSetDocumentText = "DocumentText";
    private const string ParameterSetDocumentRegex = "DocumentRegex";

    /// <summary>Path to the document to open read-only for searching.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPathText)]
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPathRegex)]
    [Alias("InputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Open document to inspect. The caller controls the document lifetime.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocumentText)]
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocumentRegex)]
    public WordDocument Document { get; set; } = null!;

    /// <summary>Literal text to find in table cells.</summary>
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = ParameterSetPathText)]
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = ParameterSetDocumentText)]
    public string Text { get; set; } = string.Empty;

    /// <summary>Regular expression pattern to find in table cells.</summary>
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = ParameterSetPathRegex)]
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = ParameterSetDocumentRegex)]
    public string Pattern { get; set; } = string.Empty;

    /// <summary>Use case-sensitive matching.</summary>
    [Parameter]
    public SwitchParameter CaseSensitive { get; set; }

    /// <summary>Include nested tables inside table cells.</summary>
    [Parameter]
    public SwitchParameter IncludeNested { get; set; }

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

            var matcher = ParameterSetName == ParameterSetPathRegex || ParameterSetName == ParameterSetDocumentRegex
                ? WordObjectSearch.CreateRegexMatcher(Pattern, CaseSensitive.IsPresent)
                : WordObjectSearch.CreateTextMatcher(Text, CaseSensitive.IsPresent);

            IEnumerable<WordTable> tables = IncludeNested.IsPresent
                ? document.TablesIncludingNestedTables
                : document.Tables;

            foreach (var table in tables) {
                if (WordObjectSearch.MatchesTable(table, matcher)) {
                    WriteObject(table);
                }
            }
        } finally {
            if (dispose) {
                document?.Dispose();
            }
        }
    }
}