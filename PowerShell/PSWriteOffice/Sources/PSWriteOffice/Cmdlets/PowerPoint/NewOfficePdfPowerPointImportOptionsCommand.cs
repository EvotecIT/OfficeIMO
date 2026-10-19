using System.Management.Automation;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Creates discoverable PDF-to-PowerPoint reconstruction settings.</summary>
/// <example>
///   <summary>Import selected PDF pages as bounded slide content.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$options = New-OfficePdfToPowerPointOptions -PageRange '1-5' -MaxPages 5 -IncludeSourceTitles
/// ConvertTo-OfficePdfPowerPoint -Path .\Source.pdf -OutputPath .\Slides.pptx -Options $options</code>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficePdfToPowerPointOptions")]
[Alias("New-OfficePdfPowerPointImportOptions")]
[OutputType(typeof(PdfToPowerPointOptions))]
public sealed class NewOfficePdfToPowerPointOptionsCommand : PSCmdlet {
    /// <summary>Visual, editable-table, hybrid, editable-content, or automatic import mode.</summary>
    [Parameter] public PdfPowerPointImportMode? Mode { get; set; }
    /// <summary>Optional one-based page ranges such as 1-3,5.</summary>
    [Parameter] public string? PageRange { get; set; }
    /// <summary>Raster resolution used by visual import.</summary>
    [Parameter] [ValidateRange(double.Epsilon, double.MaxValue)] public double? Dpi { get; set; }
    /// <summary>Maximum pages imported.</summary>
    [Parameter] [ValidateRange(1, int.MaxValue)] public int? MaxPages { get; set; }
    /// <summary>Maximum pixels per rendered page.</summary>
    [Parameter] [ValidateRange(1, long.MaxValue)] public long? MaxPixelsPerPage { get; set; }
    /// <summary>Maximum encoded bytes per rendered page.</summary>
    [Parameter] [ValidateRange(1, long.MaxValue)] public long? MaxOutputBytesPerPage { get; set; }
    /// <summary>Maximum aggregate encoded output bytes.</summary>
    [Parameter] [ValidateRange(1, long.MaxValue)] public long? MaxTotalOutputBytes { get; set; }
    /// <summary>Maximum editable objects reconstructed per page.</summary>
    [Parameter] [ValidateRange(1, int.MaxValue)] public int? MaxEditableObjectsPerPage { get; set; }
    /// <summary>Maximum body rows imported per table; zero means unlimited.</summary>
    [Parameter] [ValidateRange(0, int.MaxValue)] public int? MaxRows { get; set; }
    /// <summary>Merge compatible table segments across pages.</summary>
    [Parameter] public SwitchParameter MergePageContinuations { get; set; }
    /// <summary>Suppress repeated body header rows.</summary>
    [Parameter] public SwitchParameter SuppressRepeatedBodyHeaderRows { get; set; }
    /// <summary>Maximum rows written to one slide; zero means unlimited.</summary>
    [Parameter] [ValidateRange(0, int.MaxValue)] public int? MaxRowsPerSlide { get; set; }
    /// <summary>Maximum columns written to one slide; zero means unlimited.</summary>
    [Parameter] [ValidateRange(0, int.MaxValue)] public int? MaxColumnsPerSlide { get; set; }
    /// <summary>PowerPoint table style.</summary>
    [Parameter] public PowerPointTableStylePreset? TableStyle { get; set; }
    /// <summary>Add source-page titles.</summary>
    [Parameter] public SwitchParameter IncludeSourceTitles { get; set; }
    /// <summary>Add inferred column headers.</summary>
    [Parameter] public SwitchParameter IncludeColumnHeaderRows { get; set; }
    /// <summary>Enable banded-row styling.</summary>
    [Parameter] public SwitchParameter BandedRows { get; set; }
    /// <summary>Right-align inferred numeric columns.</summary>
    [Parameter] public SwitchParameter AlignNumericColumns { get; set; }
    /// <summary>Title used when no supported content is detected.</summary>
    [Parameter] public string? EmptyPresentationTitle { get; set; }
    /// <summary>Message used when no supported content is detected.</summary>
    [Parameter] public string? EmptyPresentationMessage { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new PdfToPowerPointOptions();
        if (Mode.HasValue) options.Mode = Mode.Value;
        if (!string.IsNullOrWhiteSpace(PageRange))
        {
            options.ReadOptions = new PdfReadOptions { PageSelection = PdfPageSelection.Parse(PageRange!) };
        }
        if (Dpi.HasValue) options.Dpi = Dpi.Value;
        if (MaxPages.HasValue) options.MaxPages = MaxPages.Value;
        if (MaxPixelsPerPage.HasValue) options.MaxPixelsPerPage = MaxPixelsPerPage.Value;
        if (MaxOutputBytesPerPage.HasValue) options.MaxOutputBytesPerPage = MaxOutputBytesPerPage.Value;
        if (MaxTotalOutputBytes.HasValue) options.MaxTotalOutputBytes = MaxTotalOutputBytes.Value;
        if (MaxEditableObjectsPerPage.HasValue) options.MaxEditableObjectsPerPage = MaxEditableObjectsPerPage.Value;
        if (MaxRows.HasValue) options.MaxRows = MaxRows.Value;
        if (MaxRowsPerSlide.HasValue) options.MaxRowsPerSlide = MaxRowsPerSlide.Value;
        if (MaxColumnsPerSlide.HasValue) options.MaxColumnsPerSlide = MaxColumnsPerSlide.Value;
        if (TableStyle.HasValue) options.TableStyle = TableStyle.Value;
        Apply(nameof(MergePageContinuations), value => options.MergePageContinuations = value);
        Apply(nameof(SuppressRepeatedBodyHeaderRows), value => options.SuppressRepeatedBodyHeaderRows = value);
        Apply(nameof(IncludeSourceTitles), value => options.IncludeSourceTitles = value);
        Apply(nameof(IncludeColumnHeaderRows), value => options.IncludeColumnHeaderRows = value);
        Apply(nameof(BandedRows), value => options.BandedRows = value);
        Apply(nameof(AlignNumericColumns), value => options.AlignNumericColumns = value);
        if (EmptyPresentationTitle != null) options.EmptyPresentationTitle = EmptyPresentationTitle;
        if (EmptyPresentationMessage != null) options.EmptyPresentationMessage = EmptyPresentationMessage;
        WriteObject(options);
    }

    private void Apply(string name, System.Action<bool> setter) {
        if (!MyInvocation.BoundParameters.ContainsKey(name)) return;
        setter(((SwitchParameter)MyInvocation.BoundParameters[name]).IsPresent);
    }
}
