using System;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Adds an image to a PowerPoint slide.</summary>
/// <para>Places the picture at the requested coordinates using point measurements.</para>
/// <example>
///   <summary>Insert an image.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$image = '.\Tests\Assets\CellImage.png'
/// New-OfficePowerPoint -Path .\Examples\Documents\PowerPointImage.pptx {
///     $slide = Add-OfficePowerPointSlide -Layout 1 -PassThru
///     Set-OfficePowerPointSlideTitle -Slide $slide -Title 'Evidence'
///     Add-OfficePowerPointImage -Slide $slide -Path $image -X 60 -Y 130 -Width 180 -Height 120
/// }</code>
///   <para>Adds a picture to a generated slide.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficePowerPointImage")]
[Alias("PptImage")]
[OutputType(typeof(PowerPointPicture))]
public sealed class AddOfficePowerPointImageCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    /// <summary>Target slide that will receive the picture (optional inside DSL).</summary>
    [Parameter(ValueFromPipeline = true)]
    public PowerPointSlide? Slide { get; set; }

    /// <summary>Path to the image file.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Path { get; set; } = string.Empty;

    /// <summary>Left offset (in points) from the slide origin.</summary>
    [Parameter]
    public double X { get; set; }

    /// <summary>Top offset (in points) from the slide origin.</summary>
    [Parameter]
    public double Y { get; set; }

    /// <summary>Image width in points.</summary>
    [Parameter]
    public double Width { get; set; } = 200;

    /// <summary>Image height in points.</summary>
    [Parameter]
    public double Height { get; set; } = 150;

    /// <inheritdoc />
    protected override void ProcessRecord() {
        try {
            if (Width <= 0) {
                throw new ArgumentOutOfRangeException(nameof(Width), "Width must be greater than 0.");
            }

            if (Height <= 0) {
                throw new ArgumentOutOfRangeException(nameof(Height), "Height must be greater than 0.");
            }

            var slide = Slide ?? PowerPointDslContext.Require(this).RequireSlide();
            var resolvedPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
            var picture = slide.AddPicturePoints(resolvedPath, X, Y, Width, Height);
            WritePassThru(picture);
        } catch (Exception ex) {
            WriteError(new ErrorRecord(ex, "PowerPointAddImageFailed", ErrorCategory.InvalidOperation, Slide));
        }
    }
}
