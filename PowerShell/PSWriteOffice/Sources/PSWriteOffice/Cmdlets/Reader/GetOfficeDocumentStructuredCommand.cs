using System.Management.Automation;
using OfficeIMO.Reader;
using PSWriteOffice.Services.Reader;

namespace PSWriteOffice.Cmdlets.Reader;

/// <summary>Extracts a bounded schema-friendly view of a supported document.</summary>
/// <example>
///   <summary>Extract sections, tables, forms, and scalar records.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$result = Get-OfficeDocumentStructured -Path .\report.docx; $result.Records | Group-Object Category</code>
///   <para>Returns deterministic structured records and source diagnostics without format-specific parsing.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeDocumentStructured")]
[OutputType(typeof(OfficeDocumentStructuredExtractionResult))]
public sealed class GetOfficeDocumentStructuredCommand : OfficeDocumentReaderCommandBase
{
    /// <summary>Path to read.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    public string Path { get; set; } = string.Empty;

    /// <summary>Optional source-reading limits and format behavior.</summary>
    [Parameter]
    public ReaderOptions? ReaderOptions { get; set; }

    /// <summary>Optional structured extraction categories and limits.</summary>
    [Parameter]
    public OfficeDocumentStructuredExtractionOptions? ExtractionOptions { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() => WriteObject(EffectiveReader.ReadStructured(
        ReaderCommandUtilities.ResolvePath(this, Path), ReaderOptions, ExtractionOptions));
}
