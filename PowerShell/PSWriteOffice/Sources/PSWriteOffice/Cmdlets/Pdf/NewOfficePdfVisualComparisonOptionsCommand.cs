using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Pdf;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Comparison;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Creates native PDF comparison settings and independent bounded policies for both inputs.</summary>
[Cmdlet(VerbsCommon.New, "OfficePdfVisualComparisonOptions", DefaultParameterSetName = "Default")]
[OutputType(typeof(OfficePdfVisualComparisonOptions))]
public sealed class NewOfficePdfVisualComparisonOptionsCommand : OfficeCmdlet {
    /// <summary>Native render scale.</summary>
    [Parameter(ParameterSetName = "Default"), ValidateRange(double.Epsilon, double.MaxValue)] public double? Scale { get; set; }
    /// <summary>Maximum per-channel difference treated as equal.</summary>
    [Parameter, ValidateRange(0, 255)] public int? ChannelTolerance { get; set; }
    /// <summary>Maximum differing-pixel ratio accepted as equal.</summary>
    [Parameter, ValidateRange(0d, 1d)] public double? AllowedDifferenceRatio { get; set; }
    /// <summary>Native page-image alignment.</summary>
    [Parameter] public PdfVisualPageAlignment? Alignment { get; set; }
    /// <summary>Raster background.</summary>
    [Parameter, OfficeColorTransformation, ArgumentCompleter(typeof(OfficeColorCompleter))] public OfficeColor? BackgroundColor { get; set; }
    /// <summary>Maximum pages on either side.</summary>
    [Parameter, ValidateRange(1, int.MaxValue)] public int? MaxPages { get; set; }
    /// <summary>Maximum pixels in one expected, actual or difference image.</summary>
    [Parameter, ValidateRange(1, long.MaxValue)] public long? MaximumPixelCountPerImage { get; set; }
    /// <summary>Maximum aggregate raster pixels.</summary>
    [Parameter, ValidateRange(1, long.MaxValue)] public long? MaximumTotalPixelCount { get; set; }
    /// <summary>Maximum retained PNG bytes.</summary>
    [Parameter, ValidateRange(1, long.MaxValue)] public long? MaxTotalOutputBytes { get; set; }
    /// <summary>Explicitly ignore authenticated owner-imposed restrictions on the reference PDF.</summary>
    [Parameter] public SwitchParameter IgnorePermissionRestrictions { get; set; }
    /// <summary>Explicitly ignore authenticated owner-imposed restrictions on the difference PDF.</summary>
    [Parameter] public SwitchParameter IgnoreDifferencePermissionRestrictions { get; set; }
    /// <summary>Bounded policy for the reference input.</summary>
    [Parameter, OfficeOptionsTransformation(typeof(PdfLoadOptions))] public PdfLoadOptions? ReferenceLoadOptions { get; set; }
    /// <summary>Bounded policy for the difference input.</summary>
    [Parameter, OfficeOptionsTransformation(typeof(PdfLoadOptions))] public PdfLoadOptions? DifferenceLoadOptions { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var native = new PdfVisualComparisonOptions();
        if (Scale.HasValue) native.Scale = Scale.Value;
        if (ChannelTolerance.HasValue) native.ChannelTolerance = (byte)ChannelTolerance.Value;
        if (AllowedDifferenceRatio.HasValue) native.AllowedDifferenceRatio = AllowedDifferenceRatio.Value;
        if (Alignment.HasValue) native.Alignment = Alignment.Value;
        if (BackgroundColor.HasValue) native.Background = BackgroundColor.Value;
        if (MaxPages.HasValue) native.MaxPages = MaxPages.Value;
        if (MaximumPixelCountPerImage.HasValue) native.MaxPixelsPerImage = MaximumPixelCountPerImage.Value;
        if (MaximumTotalPixelCount.HasValue) native.MaxTotalPixels = MaximumTotalPixelCount.Value;
        if (MaxTotalOutputBytes.HasValue) native.MaxTotalOutputBytes = MaxTotalOutputBytes.Value;
        var reference = ReferenceLoadOptions?.Clone();
        var difference = DifferenceLoadOptions?.Clone();
        if (BoundSwitch(nameof(IgnorePermissionRestrictions)) is bool ignore)
            reference = (reference ?? new PdfLoadOptions()).Snapshot(null, null, ignore ? PdfPermissionPolicy.IgnoreRestrictions : PdfPermissionPolicy.Enforce);
        if (BoundSwitch(nameof(IgnoreDifferencePermissionRestrictions)) is bool ignoreDifference)
            difference = (difference ?? new PdfLoadOptions()).Snapshot(null, null, ignoreDifference ? PdfPermissionPolicy.IgnoreRestrictions : PdfPermissionPolicy.Enforce);
        WriteObject(new OfficePdfVisualComparisonOptions { Comparison = native, ReferenceLoadOptions = reference, DifferenceLoadOptions = difference });
    }
}
