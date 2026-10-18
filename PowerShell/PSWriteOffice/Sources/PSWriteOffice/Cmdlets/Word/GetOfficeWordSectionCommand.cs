using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Gets sections from a Word document.</summary>
/// <example>
///   <summary>Inspect the first section before section-level reads.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$section = Get-OfficeWordSection -Path .\Report.docx -Index 0
/// $section |
///     Get-OfficeWordParagraph |
///     Select-Object -First 5 -Property Text</code>
///   <para>Gets a specific section by zero-based index and pipes it to another thin Word reader.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeWordSection", DefaultParameterSetName = ParameterSetPath)]
[OutputType(typeof(WordSection))]
public sealed class GetOfficeWordSectionCommand : PSCmdlet {
    private const string ParameterSetPath = "Path";
    private const string ParameterSetDocument = "Document";

    /// <summary>Path to the document.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Document to inspect.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public WordDocument Document { get; set; } = null!;

    /// <summary>Optional 0-based section index filter.</summary>
    [Parameter]
    public int[]? Index { get; set; }

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

            IEnumerable<WordSection> sections = document.Sections;

            if (Index != null && Index.Length > 0) {
                var list = document.Sections;
                var results = new List<WordSection>(Index.Length);
                foreach (var idx in Index) {
                    if (idx < 0 || idx >= list.Count) {
                        throw new ArgumentOutOfRangeException(nameof(Index), $"Section index {idx} is out of range.");
                    }
                    results.Add(list[idx]);
                }
                sections = results;
            }

            WriteObject(sections, enumerateCollection: true);
        } finally {
            if (dispose) {
                document?.Dispose();
            }
        }
    }
}