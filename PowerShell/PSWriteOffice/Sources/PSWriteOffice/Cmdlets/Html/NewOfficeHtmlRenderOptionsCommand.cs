using OfficeIMO;
using PSWriteOffice.Foundation;
using System;
using System.Management.Automation;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using PSWriteOffice.Cmdlets.Imaging;

namespace PSWriteOffice.Cmdlets.Html;

/// <summary>Creates discoverable layout, resource-limit, and rendering settings for HTML image export.</summary>
/// <example>
///   <summary>Render HTML with a bounded viewport and resource budget.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$render = New-OfficeHtmlRenderOptions -ViewportWidth 1280 -ViewportHeight 720 -MaxPageCount 10
/// Export-OfficeHtmlImage -Path .\Report.html -OutputPath .\Report.svg -RenderOptions $render</code>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficeHtmlRenderOptions", DefaultParameterSetName = "Default")]
[OutputType(typeof(HtmlRenderOptions))]
public sealed class NewOfficeHtmlRenderOptionsCommand : OfficeImageOptionsCommandBase<HtmlRenderOptions> {
    /// <summary>HTML render mode.</summary>
    [Parameter(ParameterSetName = "Default")] public HtmlRenderMode? Mode { get; set; }
    /// <summary>Fidelity policy for unsupported content.</summary>
    [Parameter] public HtmlRenderFidelityPolicy? FidelityPolicy { get; set; }
    /// <summary>Viewport width in CSS pixels.</summary>
    [Parameter, OfficeLengthTransformation] public OfficeLength? ViewportWidth { get; set; }
    /// <summary>Optional viewport height in CSS pixels.</summary>
    [Parameter, OfficeLengthTransformation] public OfficeLength? ViewportHeight { get; set; }
    /// <summary>Page size used by paged rendering.</summary>
    [Parameter, OfficePageSizeTransformation(typeof(OfficePageSize))] public OfficePageSize? PageSize { get; set; }
    /// <summary>Honor CSS page rules.</summary>
    [Parameter] public SwitchParameter NoCssPageRules { get; set; }
    /// <summary>Default font family.</summary>
    [Parameter] public string? DefaultFont { get; set; }
    /// <summary>Default font size.</summary>
    [Parameter, OfficeLengthTransformation] public OfficeLength? DefaultSize { get; set; }
    /// <summary>Default line-height multiplier.</summary>
    [Parameter] [ValidateRange(double.Epsilon, double.MaxValue)] public double? DefaultLineHeight { get; set; }
    /// <summary>Base URI for relative resources.</summary>
    [Parameter] public string? BaseUri { get; set; }
    /// <summary>Maximum rendered page count.</summary>
    [Parameter] [ValidateRange(1, int.MaxValue)] public int? MaxPageCount { get; set; }
    /// <summary>Maximum HTML input characters.</summary>
    [Parameter] [ValidateRange(1, int.MaxValue)] public int? MaxInputCharacters { get; set; }
    /// <summary>Maximum HTML nodes.</summary>
    [Parameter] [ValidateRange(1, int.MaxValue)] public int? MaxHtmlNodes { get; set; }
    /// <summary>Maximum resource bytes loaded for the document.</summary>
    [Parameter] [ValidateRange(1, long.MaxValue)] public long? MaxTotalResourceBytes { get; set; }
    /// <summary>Maximum duration allowed for one resource load.</summary>
    [Parameter] [ValidateRange(double.Epsilon, 2147483d)] public double? ResourceTimeoutSeconds { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new HtmlRenderOptions();
        ApplyCommon(options);
        if (Mode.HasValue) options.Mode = Mode.Value;
        if (FidelityPolicy.HasValue) options.FidelityPolicy = FidelityPolicy.Value;
        if (ViewportWidth.HasValue) options.ViewportWidth = ViewportWidth.Value.ToPixels();
        if (ViewportHeight.HasValue) options.ViewportHeight = ViewportHeight.Value.ToPixels();
        if (PageSize.HasValue) options.PageSize = PageSize.Value;
        if (IsBound(nameof(NoCssPageRules))) options.HonorCssPageRules = !NoCssPageRules.IsPresent;
        if (DefaultFont != null) options.DefaultFontFamily = DefaultFont;
        if (DefaultSize.HasValue) options.DefaultFontSize = DefaultSize.Value.ToPixels();
        if (DefaultLineHeight.HasValue) options.DefaultLineHeight = DefaultLineHeight.Value;
        if (!string.IsNullOrWhiteSpace(BaseUri)) options.BaseUri = HtmlOptionsCommandUtilities.NormalizeBaseUri(SessionState, BaseUri!);
        if (MaxPageCount.HasValue) options.MaxPageCount = MaxPageCount.Value;
        if (MaxInputCharacters.HasValue) options.MaxInputCharacters = MaxInputCharacters.Value;
        if (MaxHtmlNodes.HasValue) options.MaxHtmlNodes = MaxHtmlNodes.Value;
        if (MaxTotalResourceBytes.HasValue) options.MaxTotalResourceBytes = MaxTotalResourceBytes.Value;
        if (ResourceTimeoutSeconds.HasValue) options.ResourceTimeout = TimeSpan.FromSeconds(ResourceTimeoutSeconds.Value);
        WriteObject(options);
    }
}
