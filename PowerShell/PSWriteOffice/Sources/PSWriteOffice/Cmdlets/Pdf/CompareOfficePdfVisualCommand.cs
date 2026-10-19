using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Compares rendered PDF pages and returns pixel-level review artifacts.</summary>
/// <example>
///   <summary>Compare selected pages with a small pixel tolerance.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$options = New-OfficePdfVisualComparisonOptions -AllowedDifferenceRatio 0.001 -ChannelTolerance 2
/// Compare-OfficePdfVisual -ReferencePath .\Expected.pdf -DifferencePath .\Actual.pdf -PageRange '1-3' -Options $options</code>
///   <para>Returns per-page difference ratios, images, and diagnostics.</para>
/// </example>
[Cmdlet(VerbsData.Compare, "OfficePdfVisual")]
[OutputType(typeof(PdfVisualComparisonReport))]
public sealed class CompareOfficePdfVisualCommand : PSCmdlet
{
    /// <summary>Expected PDF path.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string ReferencePath { get; set; } = string.Empty;

    /// <summary>Actual PDF path.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string DifferencePath { get; set; } = string.Empty;

    /// <summary>Optional one-based ranges such as 1-3,5.</summary>
    [Parameter]
    public string? PageRange { get; set; }

    /// <summary>Optional render, tolerance, alignment, background, and ignored regions.</summary>
    [Parameter]
    public PdfVisualComparisonOptions? Options { get; set; }

    /// <summary>Optional bounded read settings for the expected document.</summary>
    [Parameter]
    public PdfLoadOptions? ReferenceReadOptions { get; set; }

    /// <summary>Optional bounded read settings for the actual document.</summary>
    [Parameter]
    public PdfLoadOptions? DifferenceReadOptions { get; set; }

    /// <summary>Password used to authenticate the expected PDF.</summary>
    [Parameter]
    public string? ReferencePassword { get; set; }

    /// <summary>After authentication, explicitly ignore restrictions on the expected PDF.</summary>
    [Parameter]
    public SwitchParameter IgnoreReferencePermissionRestrictions { get; set; }

    /// <summary>Password used to authenticate the actual PDF.</summary>
    [Parameter]
    public string? DifferencePassword { get; set; }

    /// <summary>After authentication, explicitly ignore restrictions on the actual PDF.</summary>
    [Parameter]
    public SwitchParameter IgnoreDifferencePermissionRestrictions { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var selection = string.IsNullOrWhiteSpace(PageRange) ? null : PdfPageSelection.Parse(PageRange!);
        var expected = PdfCommandUtilities.LoadDocument(
            SessionState.Path.GetUnresolvedProviderPathFromPSPath(ReferencePath),
            PdfCommandUtilities.CreateReadOptions(
                ReferenceReadOptions,
                ReferencePassword,
                IgnoreReferencePermissionRestrictions.IsPresent));
        var actual = PdfCommandUtilities.LoadDocument(
            SessionState.Path.GetUnresolvedProviderPathFromPSPath(DifferencePath),
            PdfCommandUtilities.CreateReadOptions(
                DifferenceReadOptions,
                DifferencePassword,
                IgnoreDifferencePermissionRestrictions.IsPresent));
        WriteObject(expected.Proof.CompareVisual(actual, selection, Options));
    }
}
