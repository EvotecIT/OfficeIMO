using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Gets structured content controls from a Word document.</summary>
/// <example>
///   <summary>Inventory content controls by tag.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$controls = Get-OfficeWordContentControl -Path .\Template.docx -Tag 'report-*'
/// $controls |
///     Group-Object -Property Tag |
///     Select-Object -Property Name, Count</code>
///   <para>Returns structured document tags matching the template tag convention and summarizes their counts.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeWordContentControl", DefaultParameterSetName = ParameterSetPath)]
[Alias("WordContentControls")]
[OutputType(typeof(WordStructuredDocumentTag))]
public sealed class GetOfficeWordContentControlCommand : PSCmdlet {
    private const string ParameterSetPath = "Path";
    private const string ParameterSetDocument = "Document";

    /// <summary>Path to the .docx file.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Word document to read.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public WordDocument Document { get; set; } = null!;

    /// <summary>Filter by content control alias (wildcards supported).</summary>
    [Parameter]
    [SupportsWildcards]
    public string[]? Alias { get; set; }

    /// <summary>Filter by content control tag (wildcards supported).</summary>
    [Parameter]
    [SupportsWildcards]
    public string[]? Tag { get; set; }

    /// <summary>Filter by content control text (wildcards supported).</summary>
    [Parameter]
    [SupportsWildcards]
    public string[]? Text { get; set; }

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

            var aliasPatterns = WordFilterHelpers.BuildPatterns(Alias);
            var tagPatterns = WordFilterHelpers.BuildPatterns(Tag);
            var textPatterns = WordFilterHelpers.BuildPatterns(Text);

            IEnumerable<WordStructuredDocumentTag> results = document.StructuredDocumentTags;
            results = results.Where(control =>
                WordFilterHelpers.Matches(control.Alias, aliasPatterns) &&
                WordFilterHelpers.Matches(control.Tag, tagPatterns) &&
                WordFilterHelpers.Matches(control.Text, textPatterns));

            WriteObject(results, enumerateCollection: true);
        } finally {
            if (dispose) {
                document?.Dispose();
            }
        }
    }
}