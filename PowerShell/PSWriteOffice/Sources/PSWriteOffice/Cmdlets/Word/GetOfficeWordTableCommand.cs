using System;
using System.Collections.Generic;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Gets tables from a Word document or section.</summary>
/// <example>
///   <summary>Count top-level and nested tables.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$tables = Get-OfficeWordTable -Path .\Report.docx -IncludeNested
/// [pscustomobject]@{
///     Path = '.\Report.docx'
///     TableCount = @($tables).Count
/// }</code>
///   <para>Reads tables from the document and includes nested tables when validating a composed report.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeWordTable", DefaultParameterSetName = ParameterSetPath)]
[OutputType(typeof(WordTable))]
public sealed class GetOfficeWordTableCommand : PSCmdlet {
    private const string ParameterSetPath = "Path";
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetSection = "Section";

    /// <summary>Path to the document.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Document to inspect.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public WordDocument Document { get; set; } = null!;

    /// <summary>Section to enumerate.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetSection)]
    public WordSection Section { get; set; } = null!;

    /// <summary>Include nested tables.</summary>
    [Parameter]
    public SwitchParameter IncludeNested { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        WordDocument? document = null;
        var dispose = false;

        try {
            IEnumerable<WordTable> tables;

            if (ParameterSetName == ParameterSetSection) {
                if (Section == null) {
                    tables = Array.Empty<WordTable>();
                } else {
                    tables = IncludeNested.IsPresent
                        ? Section.TablesIncludingNestedTables
                        : Section.Tables;
                }
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

                tables = IncludeNested.IsPresent
                    ? document.TablesIncludingNestedTables
                    : document.Tables;
            }

            WriteObject(tables, enumerateCollection: true);
        } finally {
            if (dispose) {
                document?.Dispose();
            }
        }
    }
}