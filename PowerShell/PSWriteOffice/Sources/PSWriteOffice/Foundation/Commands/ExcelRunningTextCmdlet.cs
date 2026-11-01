using System;
using System.IO;
using System.Management.Automation;
using System.Threading.Tasks;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Foundation;

/// <summary>Binds independent worksheet header/footer text, shared print flags and native images.</summary>
/// <para>Text edits update odd-page sections. Omitted sections and authored first/even-page content remain intact.
/// Different-page flags are shared by the header and footer; disabling a variant keeps its content for reuse.</para>
public abstract class ExcelRunningTextCmdlet : AsyncPSCmdlet {
    /// <summary>Caller-owned workbook; omission uses the enclosing Excel content block.</summary>
    [Parameter(ValueFromPipeline = true, ParameterSetName = "Text")]
    [Parameter(ValueFromPipeline = true, ParameterSetName = "ImagePath")]
    [Parameter(ValueFromPipeline = true, ParameterSetName = "ImageUrl")]
    public ExcelDocument? Document { get; set; }
    /// <summary>Worksheet name, one-based position or native worksheet.</summary>
    [Parameter, OfficeSheetTransformation]
    public OfficeSheetSelector? Sheet { get; set; }
    /// <summary>Left section; empty text clears it, omission preserves it.</summary>
    [Parameter, AllowEmptyString] public string? Left { get; set; }
    /// <summary>Center section; Excel field codes such as &amp;P and &amp;N remain active.</summary>
    [Parameter, AllowEmptyString] public string? Center { get; set; }
    /// <summary>Right section; empty text clears it, omission preserves it.</summary>
    [Parameter, AllowEmptyString] public string? Right { get; set; }
    /// <summary>Enable a first-page variant; explicitly false disables it without deleting authored variant text.</summary>
    [Parameter] public SwitchParameter DifferentFirstPage { get; set; }
    /// <summary>Enable odd/even variants; omission retains the setting.</summary>
    [Parameter] public SwitchParameter DifferentOddEven { get; set; }
    /// <summary>Disable alignment with page margins; explicitly false restores alignment.</summary>
    [Parameter] public SwitchParameter NoAlignWithMargins { get; set; }
    /// <summary>Disable scaling with the printed worksheet; explicitly false restores scaling.</summary>
    [Parameter] public SwitchParameter NoScaleWithDocument { get; set; }
    /// <summary>Literal local image file.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "ImagePath"), ValidateNotNullOrEmpty]
    public string? ImagePath { get; set; }
    /// <summary>HTTP or HTTPS image URL; existing native remote-image policy applies.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "ImageUrl"), ValidateNotNull]
    public Uri? ImageUrl { get; set; }
    /// <summary>Image section; defaults to Right.</summary>
    [Parameter(ParameterSetName = "ImagePath"), Parameter(ParameterSetName = "ImageUrl")]
    public ExcelHeaderFooterPosition ImagePosition { get; set; } = ExcelHeaderFooterPosition.Right;
    /// <summary>Physical image width; a bare number means points.</summary>
    [Parameter(ParameterSetName = "ImagePath"), Parameter(ParameterSetName = "ImageUrl"), OfficeLengthTransformation]
    public OfficeLength? ImageWidth { get; set; }
    /// <summary>Physical image height; a bare number means points.</summary>
    [Parameter(ParameterSetName = "ImagePath"), Parameter(ParameterSetName = "ImageUrl"), OfficeLengthTransformation]
    public OfficeLength? ImageHeight { get; set; }
    /// <summary>Emit the updated caller-owned worksheet.</summary>
    [Parameter] public SwitchParameter PassThru { get; set; }
    /// <summary>Whether the command edits the header rather than the footer.</summary>
    protected abstract bool IsHeader { get; }

    /// <inheritdoc />
    protected sealed override async Task ProcessRecordAsync() {
        if (ImageWidth.HasValue && ImageWidth.Value <= OfficeLength.Zero || ImageHeight.HasValue && ImageHeight.Value <= OfficeLength.Zero)
            throw new PSArgumentException("Image dimensions must be positive.");
        if (!Enum.IsDefined(typeof(ExcelHeaderFooterPosition), ImagePosition))
            throw new PSArgumentException("ImagePosition must be Left, Center or Right.");
        ExcelSheet sheet = ExcelSheetResolver.ResolveTarget(this, Document, Sheet);
        Uri? url = ImageUrl;
        if (url != null) ExcelImageUrlPolicy.RequireWebUrl(url.ToString());
        byte[]? bytes = null;
        string? contentType = null;
        if (ImagePath != null) {
            string path = OfficePath.ResolveInput(SessionState, ImagePath, literal: true)[0];
            bytes = File.ReadAllBytes(path);
            using var image = new MemoryStream(bytes, writable: false);
            if (!OfficeImageReader.TryValidateContent(image, path, out var info))
                throw new PSArgumentException("ImagePath must contain a supported complete image.");
            contentType = info.MimeType;
        }
        bool? first = OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(DifferentFirstPage));
        bool? oddEven = OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(DifferentOddEven));
        bool? noAlign = OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(NoAlignWithMargins));
        bool? noScale = OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(NoScaleWithDocument));
        if (IsHeader) sheet.SetHeader(Left, Center, Right, first, oddEven, noAlign.HasValue ? !noAlign.Value : null, noScale.HasValue ? !noScale.Value : null);
        else sheet.SetFooter(Left, Center, Right, first, oddEven, noAlign.HasValue ? !noAlign.Value : null, noScale.HasValue ? !noScale.Value : null);
        if (bytes != null) {
            if (IsHeader) sheet.SetHeaderImage(ImagePosition, bytes, contentType!, ImageWidth, ImageHeight);
            else sheet.SetFooterImage(ImagePosition, bytes, contentType!, ImageWidth, ImageHeight);
        } else if (url != null) {
            if (IsHeader) await sheet.SetHeaderImageFromUrlAsync(ImagePosition, url.AbsoluteUri, ImageWidth, ImageHeight, CancelToken).ConfigureAwait(false);
            else await sheet.SetFooterImageFromUrlAsync(ImagePosition, url.AbsoluteUri, ImageWidth, ImageHeight, CancelToken).ConfigureAwait(false);
        }
        if (PassThru) WriteObject(sheet);
    }
}
