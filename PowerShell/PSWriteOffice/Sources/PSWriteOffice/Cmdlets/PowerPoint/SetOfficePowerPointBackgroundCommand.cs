using System;
using System.Management.Automation;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.PowerPoint;
namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Sets the slide background color or image.</summary>
/// <example>
///   <summary>Apply a solid background color.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficePowerPointSlide -Presentation $ppt -Index 0 | Set-OfficePowerPointBackground -Color '#F4F7FB'</code>
///   <para>Applies a solid color fill to the slide background.</para>
/// </example>
/// <example>
///   <summary>Apply a background image.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Set-OfficePowerPointBackground -Slide $slide -ImagePath '.\hero.png'</code>
///   <para>Uses the provided image as the slide background.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficePowerPointBackground", DefaultParameterSetName = ParameterSetColor)]
[Alias("PptBackground")]
[OutputType(typeof(PowerPointSlide))]
public sealed class SetOfficePowerPointBackgroundCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    private const string ParameterSetColor = "Color";
    private const string ParameterSetImage = "Image";
    private const string ParameterSetClear = "Clear";

    /// <summary>Slide to update (optional inside a slide DSL scope).</summary>
    [Parameter(ValueFromPipeline = true)]
    public PowerPointSlide? Slide { get; set; }

    /// <summary>Background color (hex or named color).</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetColor, Position = 0)]
    [OfficeColorArgumentTransformation]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string Color { get; set; } = string.Empty;

    /// <summary>Path to a background image file.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetImage, Position = 0)]
    public string ImagePath { get; set; } = string.Empty;

    /// <summary>Clears any explicit background color or image.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetClear)]
    public SwitchParameter Clear { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        try {
            var slide = Slide ?? PowerPointDslContext.Require(this).RequireSlide();

            switch (ParameterSetName) {
                case ParameterSetImage:
                    slide.SetBackgroundImage(ResolvePath(ImagePath));
                    break;
                case ParameterSetClear:
                    slide.BackgroundColor = null;
                    slide.ClearBackgroundImage();
                    break;
                default:
                    slide.BackgroundColor = NormalizeColor(Color);
                    break;
            }

            WritePassThru(slide);
        } catch (Exception ex) {
            WriteError(new ErrorRecord(ex, "PowerPointSetBackgroundFailed", ErrorCategory.InvalidOperation, Slide));
        }
    }

    private string ResolvePath(string path) {
        var providerPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(path);
        return System.IO.Path.IsPathRooted(providerPath)
            ? providerPath
            : System.IO.Path.Combine(SessionState.Path.CurrentFileSystemLocation.Path, providerPath);
    }

    private static string NormalizeColor(string color) {
        if (string.IsNullOrWhiteSpace(color)) {
            throw new PSArgumentException("Color cannot be empty.", nameof(Color));
        }

        return OfficeColor.Parse(color).ToRgbHex().ToLowerInvariant();
    }
}