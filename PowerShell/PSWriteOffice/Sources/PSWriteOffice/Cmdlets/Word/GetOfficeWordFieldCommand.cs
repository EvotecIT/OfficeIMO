using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Gets fields, filtered by type or field code text.</summary>
[Cmdlet(VerbsCommon.Get, "OfficeWordField", DefaultParameterSetName = "Document")]
[OutputType(typeof(WordField))]
public sealed class GetOfficeWordFieldCommand : PSCmdlet {
    /// <summary>Path for this inspection.</summary>
    [Parameter(ParameterSetName = "Path", Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [SupportsWildcards]
    public string[] Path { get; set; } = Array.Empty<string>();
    /// <summary>FieldType for this inspection.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "Document")]
    public WordFieldType[]? FieldType { get; set; }
    /// <summary>Contains for this inspection.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "Document")]
    public string? Contains { get; set; }
    /// <summary>CaseSensitive for this inspection.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "Document")]
    public SwitchParameter CaseSensitive { get; set; }
    /// <summary>LiteralPath for this inspection.</summary>
    [Parameter(ParameterSetName = "LiteralPath", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Alias("PSPath", "FullName")]
    public string[] LiteralPath { get; set; } = Array.Empty<string>();
    /// <summary>Document for this inspection.</summary>
    [Parameter(ParameterSetName = "Document", ValueFromPipeline = true)]
    public WordDocument? Document { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        WordQuerySource.Visit(this, ParameterSetName, Path, LiteralPath, Document, Inspect);
    }

    private void Inspect(WordDocument document) {
        IEnumerable<WordField> items = document.Fields;
        if (FieldType != null && FieldType.Length > 0) items = items.Where(item => item.FieldType.HasValue && FieldType.Contains(item.FieldType.Value));
        if (Contains != null) {
            var comparison = CaseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
            items = items.Where(item => item.Field.IndexOf(Contains, comparison) >= 0);
        }
        WriteObject(items, true);
    }
}
