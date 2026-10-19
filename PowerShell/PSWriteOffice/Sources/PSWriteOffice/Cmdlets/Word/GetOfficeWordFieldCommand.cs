using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Gets fields from a Word document.</summary>
/// <para>Returns <see cref="WordField"/> objects, optionally filtered by type or field code.</para>
/// <example>
///   <summary>Find TOC and page fields before refresh.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$fields = Get-OfficeWordField -Path .\Report.docx -Contains 'TOC'
/// $fields |
///     Select-Object -Property FieldType, Field |
///     Format-Table -AutoSize</code>
///   <para>Filters Word fields by field-code text before running a field or table-of-contents refresh.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeWordField", DefaultParameterSetName = ParameterSetPath)]
[OutputType(typeof(WordField))]
public sealed class GetOfficeWordFieldCommand : PSCmdlet {
    private const string ParameterSetPath = "Path";
    private const string ParameterSetDocument = "Document";

    /// <summary>Path to the .docx file.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Word document to read.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public WordDocument Document { get; set; } = null!;

    /// <summary>Filter by field type.</summary>
    [Parameter]
    public WordFieldType[]? FieldType { get; set; }

    /// <summary>Filter by field code text.</summary>
    [Parameter]
    public string? Contains { get; set; }

    /// <summary>Use case-sensitive matching for <see cref="Contains"/>.</summary>
    [Parameter]
    public SwitchParameter CaseSensitive { get; set; }

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

            var fields = document.Fields;
            IEnumerable<WordField> results = fields;

            if (FieldType != null && FieldType.Length > 0) {
                var allowed = new HashSet<WordFieldType>(FieldType);
                results = fields.FindAll(f => f.FieldType.HasValue && allowed.Contains(f.FieldType.Value));
            }

            if (!string.IsNullOrWhiteSpace(Contains)) {
                var comparison = CaseSensitive.IsPresent ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
                results = results.Where(f => f.Field.IndexOf(Contains!, comparison) >= 0);
            }

            WriteObject(results, enumerateCollection: true);
        } finally {
            if (dispose) {
                document?.Dispose();
            }
        }
    }
}