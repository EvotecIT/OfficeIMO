using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Gets checkbox content controls from a Word document.</summary>
/// <example>
///   <summary>Review unchecked approval boxes.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$openItems = Get-OfficeWordCheckBox -Path .\Report.docx -Tag 'approval-*' -Unchecked
/// $openItems |
///     Select-Object -Property Alias, Tag, IsChecked |
///     Format-Table -AutoSize</code>
///   <para>Finds unchecked approval checkboxes by tag so a review workflow can show what is still pending.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeWordCheckBox", DefaultParameterSetName = ParameterSetPath)]
[Alias("WordCheckBoxes")]
[OutputType(typeof(WordCheckBox))]
public sealed class GetOfficeWordCheckBoxCommand : PSCmdlet {
    private const string ParameterSetPath = "Path";
    private const string ParameterSetDocument = "Document";

    /// <summary>Path to the .docx file.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Word document to read.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public WordDocument Document { get; set; } = null!;

    /// <summary>Filter by checkbox alias (wildcards supported).</summary>
    [Parameter]
    [SupportsWildcards]
    public string[]? Alias { get; set; }

    /// <summary>Filter by checkbox tag (wildcards supported).</summary>
    [Parameter]
    [SupportsWildcards]
    public string[]? Tag { get; set; }

    /// <summary>Only return checked checkboxes.</summary>
    [Parameter]
    public SwitchParameter Checked { get; set; }

    /// <summary>Only return unchecked checkboxes.</summary>
    [Parameter]
    public SwitchParameter Unchecked { get; set; }

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
            bool? filterChecked = null;
            if (Checked.IsPresent && !Unchecked.IsPresent) {
                filterChecked = true;
            } else if (Unchecked.IsPresent && !Checked.IsPresent) {
                filterChecked = false;
            }

            IEnumerable<WordCheckBox> results = document.CheckBoxes;
            results = results.Where(control =>
                WordFilterHelpers.Matches(control.Alias, aliasPatterns) &&
                WordFilterHelpers.Matches(control.Tag, tagPatterns));

            if (filterChecked.HasValue) {
                results = results.Where(control => control.IsChecked == filterChecked.Value);
            }

            WriteObject(results, enumerateCollection: true);
        } finally {
            if (dispose) {
                document?.Dispose();
            }
        }
    }
}