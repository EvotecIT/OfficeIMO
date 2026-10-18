using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Gets date picker content controls from a Word document.</summary>
/// <example>
///   <summary>Find date picker controls used by a form.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$dates = Get-OfficeWordDatePicker -Path .\IntakeForm.docx -Tag 'deadline-*'
/// $dates |
///     Select-Object -Property Alias, Tag |
///     Format-Table -AutoSize</code>
///   <para>Filters date picker content controls by tag before a form validation pass.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeWordDatePicker", DefaultParameterSetName = ParameterSetPath)]
[Alias("WordDatePickers")]
[OutputType(typeof(WordDatePicker))]
public sealed class GetOfficeWordDatePickerCommand : PSCmdlet {
    private const string ParameterSetPath = "Path";
    private const string ParameterSetDocument = "Document";

    /// <summary>Path to the .docx file.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Word document to read.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public WordDocument Document { get; set; } = null!;

    /// <summary>Filter by date picker alias (wildcards supported).</summary>
    [Parameter]
    [SupportsWildcards]
    public string[]? Alias { get; set; }

    /// <summary>Filter by date picker tag (wildcards supported).</summary>
    [Parameter]
    [SupportsWildcards]
    public string[]? Tag { get; set; }

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

            IEnumerable<WordDatePicker> results = document.DatePickers;
            results = results.Where(control =>
                WordFilterHelpers.Matches(control.Alias, aliasPatterns) &&
                WordFilterHelpers.Matches(control.Tag, tagPatterns));

            WriteObject(results, enumerateCollection: true);
        } finally {
            if (dispose) {
                document?.Dispose();
            }
        }
    }
}