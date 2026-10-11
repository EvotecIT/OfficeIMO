using System;
using System.Linq;
using System.Management.Automation;
using OfficeIMO;
#if !FRAMEWORK
using OfficeIMO.Workflows;
#endif

namespace PSWriteOffice.Cmdlets.Conversion;

/// <summary>Lists conversion contracts from the installed OfficeIMO catalog and their availability for PSWriteOffice batch PDF export.</summary>
/// <para>Each row retains the owning package, API, input kind, fidelity, support evidence, text-formatting contract and known limitations.
/// PdfBatchSupported describes Export-OfficeDocumentPdf in the current host. Other routes may have format-specific cmdlets or .NET APIs;
/// a catalog entry alone does not establish that a converter, credentials, fonts or optional providers are ready for a particular input.</para>
/// <example>
///   <summary>Inspect a Word-to-PDF contract before converting documents.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficeConversionRoute -SourceExtension docx -TargetExtension pdf | Format-List *</code>
///   <para>Extension filters accept a leading dot and compare without regard to case.</para>
/// </example>
/// <example>
///   <summary>Find routes available to batch PDF export in this host.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficeConversionRoute -PdfBatch | Select-Object Id, SourceExtensions, PackageId, Fidelity, KnownLimitations</code>
///   <para>Batch PDF export requires PowerShell 7.4 or newer. Windows PowerShell 5.1 can inspect the catalog and use supported single-file commands.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeConversionRoute")]
[OutputType(typeof(OfficeConversionCapability))]
public sealed class GetOfficeConversionRouteCommand : PSCmdlet
{
    /// <summary>Exact route identifiers, for example docx-pdf or pdf-docx. Unknown identifiers return no rows.</summary>
    [Parameter(ValueFromPipelineByPropertyName = true)]
    [ValidateNotNullOrEmpty]
    public string[]? Id { get; set; }

    /// <summary>Accepted source extension, with or without a leading dot.</summary>
    [Parameter]
    [ValidateNotNullOrEmpty]
    public string? SourceExtension { get; set; }

    /// <summary>Destination extension, with or without a leading dot.</summary>
    [Parameter]
    [ValidateNotNullOrEmpty]
    public string? TargetExtension { get; set; }

    /// <summary>Return only routes supported by Export-OfficeDocumentPdf batch parameter sets in the current host.</summary>
    [Parameter]
    public SwitchParameter PdfBatch { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        string? source = NormalizeExtension(SourceExtension);
        string? target = NormalizeExtension(TargetExtension);
        foreach (OfficeConversionCapability route in OfficeConversionCapabilityCatalog.All)
        {
            if (Id != null && !Id.Contains(route.Id, StringComparer.OrdinalIgnoreCase)) continue;
            if (source != null && !route.SourceExtensions.Contains(source, StringComparer.OrdinalIgnoreCase)) continue;
            if (target != null && !string.Equals(route.TargetExtension, target, StringComparison.OrdinalIgnoreCase)) continue;
            bool pdfBatchSupported = SupportsPdfBatch(route);
            if (PdfBatch.IsPresent && !pdfBatchSupported) continue;

            // A fresh wrapper keeps host-specific surface information out of the immutable owner catalog.
            var row = new PSObject(route);
            row.Properties.Add(new PSNoteProperty("PdfBatchSupported", pdfBatchSupported));
            row.Properties.Add(new PSNoteProperty("PdfBatchCommand", pdfBatchSupported ? "Export-OfficeDocumentPdf" : null));
            row.Properties.Add(new PSNoteProperty("PdfBatchMinimumPowerShellVersion", new Version(7, 4)));
            WriteObject(row);
        }
    }

    private static string? NormalizeExtension(string? extension)
    {
        if (extension == null) return null;
        string value = extension.Trim();
        return value.StartsWith(".", StringComparison.Ordinal) ? value : "." + value;
    }

    private static bool SupportsPdfBatch(OfficeConversionCapability route)
    {
#if FRAMEWORK
        return false;
#else
        return string.Equals(route.TargetExtension, ".pdf", StringComparison.OrdinalIgnoreCase) &&
            (route.InputKind == OfficeConversionInputKind.File || route.InputKind == OfficeConversionInputKind.Text) &&
            OfficeWorkflowCatalog.FindExecutable(route.Id) != null;
#endif
    }
}
