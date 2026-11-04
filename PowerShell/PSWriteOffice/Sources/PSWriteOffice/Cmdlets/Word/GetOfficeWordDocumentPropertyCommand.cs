using PSWriteOffice.Models.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Gets built-in and custom document properties.</summary>
[Cmdlet(VerbsCommon.Get, "OfficeWordDocumentProperty", DefaultParameterSetName = "Document")]
[OutputType(typeof(WordDocumentPropertyInfo))]
public sealed class GetOfficeWordDocumentPropertyCommand : PSCmdlet {
    /// <summary>Path for this inspection.</summary>
    [Parameter(ParameterSetName = "Path", Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [SupportsWildcards]
    public string[] Path { get; set; } = Array.Empty<string>();
    /// <summary>Name for this inspection.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "Document")]
    public string[]? Name { get; set; }
    /// <summary>BuiltIn for this inspection.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "Document")]
    public SwitchParameter BuiltIn { get; set; }
    /// <summary>Custom for this inspection.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "Document")]
    public SwitchParameter Custom { get; set; }
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
        var names = WordFilterHelpers.BuildPatterns(Name);
        var items = WordDocumentPropertyService.GetProperties(document, !Custom || BuiltIn, !BuiltIn || Custom);
        WriteObject(items.Where(item => WordFilterHelpers.Matches(item.Name, names)), true);
    }
}
