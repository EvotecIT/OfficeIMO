using System;
using System.Management.Automation;
using OfficeIMO.Drawing;

namespace PSWriteOffice.Cmdlets.Imaging;

/// <summary>Shared PowerShell-native parameters for OfficeIMO image export option builders.</summary>
public abstract class OfficeImageOptionsCommandBase<TOptions> : PSCmdlet where TOptions : OfficeImageExportOptions {
    /// <summary>Output scale multiplier.</summary>
    [Parameter] [ValidateRange(double.Epsilon, double.MaxValue)] public double? Scale { get; set; }
    /// <summary>Maximum output width in pixels.</summary>
    [Parameter] [ValidateRange(1, int.MaxValue)] public int? MaximumOutputWidth { get; set; }
    /// <summary>Maximum output height in pixels.</summary>
    [Parameter] [ValidateRange(1, int.MaxValue)] public int? MaximumOutputHeight { get; set; }
    /// <summary>Background color name or hex value.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? BackgroundColor { get; set; }
    /// <summary>Target output density in dots per inch.</summary>
    [Parameter] [ValidateRange(double.Epsilon, 65535d)] public double? TargetDpi { get; set; }
    /// <summary>Maximum pixels allocated for one raster image.</summary>
    [Parameter] [ValidateRange(1, long.MaxValue)] public long? MaximumRasterPixels { get; set; }
    /// <summary>Reduce or reject oversized raster output.</summary>
    [Parameter] public OfficeRasterOverflowBehavior? RasterOverflowBehavior { get; set; }
    /// <summary>Maximum images accepted from one batch export.</summary>
    [Parameter] [ValidateRange(1, int.MaxValue)] public int? MaximumOutputCount { get; set; }
    /// <summary>Maximum aggregate raster pixels accepted from one batch.</summary>
    [Parameter] [ValidateRange(1, long.MaxValue)] public long? MaximumTotalRasterPixels { get; set; }
    /// <summary>Maximum aggregate encoded bytes accepted from one batch.</summary>
    [Parameter] [ValidateRange(1, long.MaxValue)] public long? MaximumTotalEncodedBytes { get; set; }
    /// <summary>Maximum seconds allowed for one render.</summary>
    [Parameter] [ValidateRange(double.Epsilon, 2147483d)] public double? RenderTimeoutSeconds { get; set; }
    /// <summary>Maximum independent renders processed concurrently.</summary>
    [Parameter] [ValidateRange(1, int.MaxValue)] public int? MaximumDegreeOfParallelism { get; set; }
    /// <summary>BCP 47 language hint for text shaping.</summary>
    [Parameter] public string? TextShapingLanguage { get; set; }

    /// <summary>Applies shared settings to a format-specific option object.</summary>
    protected void ApplyCommon(TOptions options) {
        if (Scale.HasValue) options.Scale = Scale.Value;
        if (MaximumOutputWidth.HasValue) options.MaximumOutputWidth = MaximumOutputWidth.Value;
        if (MaximumOutputHeight.HasValue) options.MaximumOutputHeight = MaximumOutputHeight.Value;
        if (!string.IsNullOrWhiteSpace(BackgroundColor)) options.BackgroundColor = OfficeColor.Parse(BackgroundColor!);
        if (TargetDpi.HasValue) options.TargetDpi = TargetDpi.Value;
        if (MaximumRasterPixels.HasValue) options.MaximumRasterPixels = MaximumRasterPixels.Value;
        if (RasterOverflowBehavior.HasValue) options.RasterOverflowBehavior = RasterOverflowBehavior.Value;
        if (MaximumOutputCount.HasValue) options.MaximumOutputCount = MaximumOutputCount.Value;
        if (MaximumTotalRasterPixels.HasValue) options.MaximumTotalRasterPixels = MaximumTotalRasterPixels.Value;
        if (MaximumTotalEncodedBytes.HasValue) options.MaximumTotalEncodedBytes = MaximumTotalEncodedBytes.Value;
        if (RenderTimeoutSeconds.HasValue) options.RenderTimeout = TimeSpan.FromSeconds(RenderTimeoutSeconds.Value);
        if (MaximumDegreeOfParallelism.HasValue) options.MaximumDegreeOfParallelism = MaximumDegreeOfParallelism.Value;
        if (TextShapingLanguage != null) options.TextShapingLanguage = TextShapingLanguage;
    }

    /// <summary>Returns whether PowerShell bound a parameter, including an explicitly false switch.</summary>
    protected bool IsBound(string name) => MyInvocation.BoundParameters.ContainsKey(name);
}
