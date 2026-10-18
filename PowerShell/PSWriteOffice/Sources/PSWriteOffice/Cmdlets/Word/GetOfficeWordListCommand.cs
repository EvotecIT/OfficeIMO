using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Gets lists from a Word document or section.</summary>
/// <para>
/// Returns existing <see cref="WordList"/> objects so scripts can inspect list items, report on
/// checklist content, or pipe a selected list to <c>Add-OfficeWordListItem</c>. Use this command when
/// the script needs to work with list objects directly instead of using a text search first.
/// </para>
/// <example>
///   <summary>Inspect existing list items.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$doc = Get-OfficeWord -Path .\Report.docx
/// $list = $doc | Get-OfficeWordList | Select-Object -First 1
/// $list.ListItems | Select-Object -Property Text</code>
///   <para>Returns OfficeIMO list objects so existing list items can be reviewed or appended to.</para>
/// </example>
/// <example>
///   <summary>Select a checklist by existing item text and append an item.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$doc = Get-OfficeWord -Path .\Report.docx
/// Get-OfficeWordList -Document $doc |
///     Where-Object { $_.ListItems.Text -contains 'Initial review' } |
///     Add-OfficeWordListItem -Text 'Final approval'
/// $doc | Close-OfficeWord -Save</code>
///   <para>Inspects list objects directly when the caller wants custom selection logic.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeWordList", DefaultParameterSetName = ParameterSetPath)]
[OutputType(typeof(WordList))]
public sealed class GetOfficeWordListCommand : PSCmdlet {
    private const string ParameterSetPath = "Path";
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetSection = "Section";

    /// <summary>Path to the document to open read-only for list inspection.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Open document to inspect. The caller controls the document lifetime.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public WordDocument Document { get; set; } = null!;

    /// <summary>Section to inspect when the caller only wants lists in a specific section.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetSection)]
    public WordSection Section { get; set; } = null!;

    /// <summary>Include numbering definitions that do not currently have list items.</summary>
    [Parameter]
    public SwitchParameter IncludeEmpty { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        WordDocument? document = null;
        var dispose = false;

        try {
            IEnumerable<WordList> lists;

            if (ParameterSetName == ParameterSetSection) {
                lists = Section != null
                    ? Section.Lists
                    : Array.Empty<WordList>();
            } else {
                if (ParameterSetName == ParameterSetPath) {
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

            if (!IncludeEmpty.IsPresent) {
                lists = lists.Where(list => list.ListItems.Count > 0);
            }

            WriteObject(lists, enumerateCollection: true);
        } finally {
            if (dispose) {
                document?.Dispose();
            }
        }
    }
}