using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Finds Word lists containing matching list-item text.</summary>
/// <para>
/// Searches list items in a Word document or section and returns matching <see cref="WordList"/>
/// objects. This is useful when a document has an existing checklist or numbered list and the script
/// needs to append to that list rather than create a new one elsewhere in the document.
/// </para>
/// <para>
/// Use <c>-Text</c> for literal contains matching or <c>-Pattern</c> for regular expressions. The
/// returned list objects can be piped directly to <c>Add-OfficeWordListItem</c>.
/// </para>
/// <example>
///   <summary>Find an existing checklist and append another item.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$doc = Get-OfficeWord -Path .\Report.docx
/// Find-OfficeWordList -Document $doc -Text 'Initial review' |
///     Add-OfficeWordListItem -Text 'Final approval'</code>
///   <para>Searches list-item paragraphs and returns matching OfficeIMO list objects for further editing.</para>
/// </example>
/// <example>
///   <summary>Append several follow-up items to an existing checklist.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$doc = Get-OfficeWord -Path .\Handover.docx
/// $list = Find-OfficeWordList -Document $doc -Text 'Initial review' | Select-Object -First 1
/// $list | Add-OfficeWordListItem -Text 'Business sign-off'
/// $list | Add-OfficeWordListItem -Text 'Go-live approval'
/// $doc | Close-OfficeWord -Save</code>
///   <para>Finds a checklist by an item it already contains and appends new items to the same list.</para>
/// </example>
[Cmdlet(VerbsCommon.Find, "OfficeWordList", DefaultParameterSetName = ParameterSetPathText)]
[OutputType(typeof(WordList))]
public sealed class FindOfficeWordListCommand : PSCmdlet {
    private const string ParameterSetPathText = "PathText";
    private const string ParameterSetPathRegex = "PathRegex";
    private const string ParameterSetDocumentText = "DocumentText";
    private const string ParameterSetDocumentRegex = "DocumentRegex";
    private const string ParameterSetSectionText = "SectionText";
    private const string ParameterSetSectionRegex = "SectionRegex";

    /// <summary>Path to the document to open read-only for searching.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPathText)]
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPathRegex)]
    [Alias("InputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Open document to inspect. The caller controls the document lifetime.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocumentText)]
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocumentRegex)]
    public WordDocument Document { get; set; } = null!;

    /// <summary>Section to inspect when the caller only wants lists in a specific section.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetSectionText)]
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetSectionRegex)]
    public WordSection Section { get; set; } = null!;

    /// <summary>Literal text to find in list items.</summary>
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = ParameterSetPathText)]
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = ParameterSetDocumentText)]
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = ParameterSetSectionText)]
    public string Text { get; set; } = string.Empty;

    /// <summary>Regular expression pattern to find in list items.</summary>
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = ParameterSetPathRegex)]
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = ParameterSetDocumentRegex)]
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = ParameterSetSectionRegex)]
    public string Pattern { get; set; } = string.Empty;

    /// <summary>Use case-sensitive matching.</summary>
    [Parameter]
    public SwitchParameter CaseSensitive { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        WordDocument? document = null;
        var dispose = false;

        try {
            IEnumerable<WordList> lists;
            if (ParameterSetName == ParameterSetSectionText || ParameterSetName == ParameterSetSectionRegex) {
                lists = Section != null
                    ? Section.Lists
                    : Array.Empty<WordList>();
            } else {
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

                lists = document.Lists;
            }

            var matcher = ParameterSetName.EndsWith("Regex", StringComparison.Ordinal)
                ? WordObjectSearch.CreateRegexMatcher(Pattern, CaseSensitive.IsPresent)
                : WordObjectSearch.CreateTextMatcher(Text, CaseSensitive.IsPresent);

            foreach (var list in lists.Where(list => list.ListItems.Count > 0)) {
                if (WordObjectSearch.MatchesList(list, matcher)) {
                    WriteObject(list);
                }
            }
        } finally {
            if (dispose) {
                document?.Dispose();
            }
        }
    }
}