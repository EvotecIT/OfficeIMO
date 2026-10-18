using System;
using System.Collections.Generic;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Gets bookmarks from a Word document.</summary>
/// <para>Returns <see cref="WordBookmark"/> objects, optionally filtered by name.</para>
/// <example>
///   <summary>List report bookmarks by name pattern.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$bookmarks = Get-OfficeWordBookmark -Path .\Report.docx -Name 'Appendix*', 'ExecutiveSummary'
/// $bookmarks |
///     Select-Object -Property Name |
///     Format-Table -AutoSize</code>
///   <para>Loads the document read-only, filters bookmark names with wildcards, and displays the matching navigation targets.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeWordBookmark", DefaultParameterSetName = ParameterSetPath)]
[OutputType(typeof(WordBookmark))]
public sealed class GetOfficeWordBookmarkCommand : PSCmdlet {
    private const string ParameterSetPath = "Path";
    private const string ParameterSetDocument = "Document";

    /// <summary>Path to the .docx file.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Word document to read.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public WordDocument Document { get; set; } = null!;

    /// <summary>Bookmark name filter (wildcards supported).</summary>
    [Parameter]
    [SupportsWildcards]
    public string[]? Name { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        WordDocument? document = null;
        var dispose = false;

        try {
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

            var bookmarks = document.Bookmarks;
            IEnumerable<WordBookmark> results = bookmarks;
            if (Name != null && Name.Length > 0) {
                var patterns = new List<WildcardPattern>();
                foreach (var pattern in Name) {
                    if (!string.IsNullOrWhiteSpace(pattern)) {
                        patterns.Add(new WildcardPattern(pattern, WildcardOptions.IgnoreCase));
                    }
                }

                if (patterns.Count > 0) {
                    results = bookmarks.FindAll(b =>
                        b.Name != null && patterns.Exists(p => p.IsMatch(b.Name)));
                }
            }

            WriteObject(results, enumerateCollection: true);
        } finally {
            if (dispose) {
                document?.Dispose();
            }
        }
    }
}