using System.IO;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Adds a text or image stamp to an existing PDF.</summary>
/// <remarks>
/// Stamps are existing-PDF operations. Use text stamps for review labels and image stamps for logos or approval marks.
/// Use <c>-Watermark</c> when the stamp should be placed behind existing page content.
/// </remarks>
/// <example>
///   <summary>Add a review stamp to selected pages.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$proof = @(
///     Add-OfficePdfStamp -Path .\Examples\Documents\Report.pdf -OutputPath .\Examples\Documents\Stamped.pdf -Text 'REVIEWED' -Color '#0F766E' -FontSize 24 -Rotation 12 -PageRange '1-2'
///     Get-OfficePdfPreflight -Path .\Examples\Documents\Stamped.pdf
/// )
/// $proof</code>
///   <para>Adds a text stamp to the first two pages and preflights the result.</para>
/// </example>
/// <example>
///   <summary>Add an image watermark.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$logo = '.\Tests\Assets\CellImage.png'
/// Add-OfficePdfStamp -Path .\Examples\Documents\Report.pdf -OutputPath .\Examples\Documents\Watermarked.pdf -Image $logo -Width 160 -Watermark</code>
///   <para>Adds a logo behind existing content as a watermark.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficePdfStamp", DefaultParameterSetName = ParameterSetText, SupportsShouldProcess = true)]
[Alias("PdfStamp")]
[OutputType(typeof(FileInfo))]
public sealed class AddOfficePdfStampCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    private const string ParameterSetText = "Text";
    private const string ParameterSetImage = "Image";

    /// <summary>Input PDF path.</summary>
    [Parameter(Mandatory = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Output PDF path.</summary>
    [Parameter(Mandatory = true)]
    public string OutputPath { get; set; } = string.Empty;

    /// <summary>Password used to authenticate an encrypted input PDF.</summary>
    [Parameter]
    public string? Password { get; set; }

    /// <summary>
    /// After successful password authentication, explicitly ignore owner-imposed usage restrictions.
    /// This does not discover, bypass, or crack a missing password.
    /// </summary>
    [Parameter]
    public SwitchParameter IgnorePermissionRestrictions { get; set; }

    /// <summary>Text to stamp.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetText)]
    public string? Text { get; set; }

    /// <summary>Image path to stamp.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetImage)]
    [Alias("ImagePath")]
    public string? Image { get; set; }

    /// <summary>Stamp selected pages, for example 1-3,5. Omit to stamp every page.</summary>
    [Parameter]
    public string? PageRange { get; set; }

    /// <summary>X coordinate in PDF points.</summary>
    [Parameter]
    public double? X { get; set; }

    /// <summary>Y coordinate in PDF points.</summary>
    [Parameter]
    public double? Y { get; set; }

    /// <summary>Font size for text stamps.</summary>
    [Parameter(ParameterSetName = ParameterSetText)]
    public double FontSize { get; set; } = 24;

    /// <summary>Text color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter(ParameterSetName = ParameterSetText)]
    [OfficeColorArgumentTransformation]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? Color { get; set; }

    /// <summary>Rendered image width in PDF points.</summary>
    [Parameter(ParameterSetName = ParameterSetImage)]
    public double? Width { get; set; }

    /// <summary>Rendered image height in PDF points.</summary>
    [Parameter(ParameterSetName = ParameterSetImage)]
    public double? Height { get; set; }

    /// <summary>Rotation in degrees.</summary>
    [Parameter]
    public double Rotation { get; set; }

    /// <summary>Place the stamp behind existing content as a watermark.</summary>
    [Parameter]
    public SwitchParameter Watermark { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var outputPath = PdfCommandUtilities.ResolvePath(this, OutputPath);
        if (!PdfCommandUtilities.ShouldWrite(this, outputPath, "Write stamped PDF")) {
            return;
        }

        var readOptions = PdfCommandUtilities.CreateReadOptions(Password, IgnorePermissionRestrictions.IsPresent);
        var document = PdfDocument.Load(PdfCommandUtilities.ResolvePath(this, Path), readOptions);
        PdfDocument result = ParameterSetName == ParameterSetImage
            ? StampImage(document)
            : StampText(document);

        PdfCommandUtilities.EnsureDirectory(outputPath);
        result.Save(outputPath).RequireSuccess();
        WritePassThru(new FileInfo(outputPath));
    }

    private PdfDocument StampText(PdfDocument document) {
        var options = new PdfTextStampOptions {
            X = X,
            Y = Y,
            FontSize = FontSize,
            RotationDegrees = Rotation,
            BehindContent = Watermark.IsPresent
        };
        var color = PdfCommandUtilities.ParseColor(Color);
        if (color.HasValue) {
            options.Color = color.Value;
        }

        PdfCommandUtilities.ApplyPageRange(options, PageRange);
        return Watermark.IsPresent
            ? document.Stamp.TextWatermark(Text!, options)
            : document.Stamp.Text(Text!, options);
    }

    private PdfDocument StampImage(PdfDocument document) {
        var options = new PdfImageStampOptions {
            X = X,
            Y = Y,
            Width = Width,
            Height = Height,
            RotationDegrees = Rotation,
            BehindContent = Watermark.IsPresent
        };
        PdfCommandUtilities.ApplyPageRange(options, PageRange);
        using var imageStream = File.OpenRead(PdfCommandUtilities.ResolvePath(this, Image!));
        return Watermark.IsPresent
            ? document.Stamp.ImageWatermark(imageStream, options)
            : document.Stamp.Image(imageStream, options);
    }
}
