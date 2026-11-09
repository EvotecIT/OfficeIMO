using System.Management.Automation;
using System.Security;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.Pdf;

public sealed partial class AddOfficePdfStampCommand {
    /// <summary>Input PDF files with FileSystem wildcard expansion.</summary>
    [Parameter(ParameterSetName = "Text", Mandatory = true, Position = 0)]
    [Parameter(ParameterSetName = "TextDirectory", Mandatory = true, Position = 0)]
    [Parameter(ParameterSetName = "Image", Mandatory = true, Position = 0)]
    [Parameter(ParameterSetName = "ImageDirectory", Mandatory = true, Position = 0)]
    [Parameter(ParameterSetName = "Content", Mandatory = true, Position = 0)]
    [Parameter(ParameterSetName = "ContentDirectory", Mandatory = true, Position = 0)]
    [Parameter(ParameterSetName = "Page", Mandatory = true, Position = 0)]
    [Parameter(ParameterSetName = "PageDirectory", Mandatory = true, Position = 0)]
    [SupportsWildcards, ValidateNotNullOrEmpty]
    public string[] Path { get; set; } = Array.Empty<string>();

    /// <summary>Single output PDF file.</summary>
    [Parameter(ParameterSetName = "Text", Mandatory = true)]
    [Parameter(ParameterSetName = "TextLiteralPath", Mandatory = true)]
    [Parameter(ParameterSetName = "Image", Mandatory = true)]
    [Parameter(ParameterSetName = "ImageLiteralPath", Mandatory = true)]
    [Parameter(ParameterSetName = "Content", Mandatory = true)]
    [Parameter(ParameterSetName = "ContentLiteralPath", Mandatory = true)]
    [Parameter(ParameterSetName = "Page", Mandatory = true)]
    [Parameter(ParameterSetName = "PageLiteralPath", Mandatory = true)]
    public string? OutputPath { get; set; }

    /// <summary>Literal text to stamp.</summary>
    [Parameter(ParameterSetName = "Text", Mandatory = true)]
    [Parameter(ParameterSetName = "TextDirectory", Mandatory = true)]
    [Parameter(ParameterSetName = "TextLiteralPath", Mandatory = true)]
    [Parameter(ParameterSetName = "TextLiteralPathDirectory", Mandatory = true)]
    [Parameter(ParameterSetName = "TextDocument", Mandatory = true)]
    [AllowEmptyString]
    public string? Text { get; set; }

    /// <summary>Text size; omission keeps the native 24-point default.</summary>
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "TextDirectory")]
    [Parameter(ParameterSetName = "TextLiteralPath")]
    [Parameter(ParameterSetName = "TextLiteralPathDirectory")]
    [Parameter(ParameterSetName = "TextDocument")]
    [OfficeLengthTransformation]
    public OfficeLength? Size { get; set; }

    /// <summary>Native Color setting for the stamp.</summary>
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "TextDirectory")]
    [Parameter(ParameterSetName = "TextLiteralPath")]
    [Parameter(ParameterSetName = "TextLiteralPathDirectory")]
    [Parameter(ParameterSetName = "TextDocument")]
    [OfficeColorTransformation, ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? Color { get; set; }

    /// <summary>Native one-based page selector, including last, odd and even.</summary>
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "TextDirectory")]
    [Parameter(ParameterSetName = "TextLiteralPath")]
    [Parameter(ParameterSetName = "TextLiteralPathDirectory")]
    [Parameter(ParameterSetName = "TextDocument")]
    [Parameter(ParameterSetName = "Image")]
    [Parameter(ParameterSetName = "ImageDirectory")]
    [Parameter(ParameterSetName = "ImageLiteralPath")]
    [Parameter(ParameterSetName = "ImageLiteralPathDirectory")]
    [Parameter(ParameterSetName = "ImageDocument")]
    [Parameter(ParameterSetName = "Content")]
    [Parameter(ParameterSetName = "ContentDirectory")]
    [Parameter(ParameterSetName = "ContentLiteralPath")]
    [Parameter(ParameterSetName = "ContentLiteralPathDirectory")]
    [Parameter(ParameterSetName = "ContentDocument")]
    [Parameter(ParameterSetName = "Page")]
    [Parameter(ParameterSetName = "PageDirectory")]
    [Parameter(ParameterSetName = "PageLiteralPath")]
    [Parameter(ParameterSetName = "PageLiteralPathDirectory")]
    [Parameter(ParameterSetName = "PageDocument")]
    [OfficePageSelectorTransformation]
    public OfficePageSelector? Page { get; set; }

    /// <summary>Native X setting for the stamp.</summary>
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "TextDirectory")]
    [Parameter(ParameterSetName = "TextLiteralPath")]
    [Parameter(ParameterSetName = "TextLiteralPathDirectory")]
    [Parameter(ParameterSetName = "TextDocument")]
    [Parameter(ParameterSetName = "Image")]
    [Parameter(ParameterSetName = "ImageDirectory")]
    [Parameter(ParameterSetName = "ImageLiteralPath")]
    [Parameter(ParameterSetName = "ImageLiteralPathDirectory")]
    [Parameter(ParameterSetName = "ImageDocument")]
    [Parameter(ParameterSetName = "Content")]
    [Parameter(ParameterSetName = "ContentDirectory")]
    [Parameter(ParameterSetName = "ContentLiteralPath")]
    [Parameter(ParameterSetName = "ContentLiteralPathDirectory")]
    [Parameter(ParameterSetName = "ContentDocument")]
    [Parameter(ParameterSetName = "Page")]
    [Parameter(ParameterSetName = "PageDirectory")]
    [Parameter(ParameterSetName = "PageLiteralPath")]
    [Parameter(ParameterSetName = "PageLiteralPathDirectory")]
    [Parameter(ParameterSetName = "PageDocument")]
    [OfficeLengthTransformation]
    public OfficeLength? X { get; set; }

    /// <summary>Native Y setting for the stamp.</summary>
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "TextDirectory")]
    [Parameter(ParameterSetName = "TextLiteralPath")]
    [Parameter(ParameterSetName = "TextLiteralPathDirectory")]
    [Parameter(ParameterSetName = "TextDocument")]
    [Parameter(ParameterSetName = "Image")]
    [Parameter(ParameterSetName = "ImageDirectory")]
    [Parameter(ParameterSetName = "ImageLiteralPath")]
    [Parameter(ParameterSetName = "ImageLiteralPathDirectory")]
    [Parameter(ParameterSetName = "ImageDocument")]
    [Parameter(ParameterSetName = "Content")]
    [Parameter(ParameterSetName = "ContentDirectory")]
    [Parameter(ParameterSetName = "ContentLiteralPath")]
    [Parameter(ParameterSetName = "ContentLiteralPathDirectory")]
    [Parameter(ParameterSetName = "ContentDocument")]
    [Parameter(ParameterSetName = "Page")]
    [Parameter(ParameterSetName = "PageDirectory")]
    [Parameter(ParameterSetName = "PageLiteralPath")]
    [Parameter(ParameterSetName = "PageLiteralPathDirectory")]
    [Parameter(ParameterSetName = "PageDocument")]
    [OfficeLengthTransformation]
    public OfficeLength? Y { get; set; }

    /// <summary>Native Rotation setting for the stamp.</summary>
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "TextDirectory")]
    [Parameter(ParameterSetName = "TextLiteralPath")]
    [Parameter(ParameterSetName = "TextLiteralPathDirectory")]
    [Parameter(ParameterSetName = "TextDocument")]
    [Parameter(ParameterSetName = "Image")]
    [Parameter(ParameterSetName = "ImageDirectory")]
    [Parameter(ParameterSetName = "ImageLiteralPath")]
    [Parameter(ParameterSetName = "ImageLiteralPathDirectory")]
    [Parameter(ParameterSetName = "ImageDocument")]
    [Parameter(ParameterSetName = "Content")]
    [Parameter(ParameterSetName = "ContentDirectory")]
    [Parameter(ParameterSetName = "ContentLiteralPath")]
    [Parameter(ParameterSetName = "ContentLiteralPathDirectory")]
    [Parameter(ParameterSetName = "ContentDocument")]
    [Parameter(ParameterSetName = "Page")]
    [Parameter(ParameterSetName = "PageDirectory")]
    [Parameter(ParameterSetName = "PageLiteralPath")]
    [Parameter(ParameterSetName = "PageLiteralPathDirectory")]
    [Parameter(ParameterSetName = "PageDocument")]
    public double Rotation { get; set; }

    /// <summary>Native Opacity setting for the stamp.</summary>
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "TextDirectory")]
    [Parameter(ParameterSetName = "TextLiteralPath")]
    [Parameter(ParameterSetName = "TextLiteralPathDirectory")]
    [Parameter(ParameterSetName = "TextDocument")]
    [Parameter(ParameterSetName = "Image")]
    [Parameter(ParameterSetName = "ImageDirectory")]
    [Parameter(ParameterSetName = "ImageLiteralPath")]
    [Parameter(ParameterSetName = "ImageLiteralPathDirectory")]
    [Parameter(ParameterSetName = "ImageDocument")]
    [Parameter(ParameterSetName = "Content")]
    [Parameter(ParameterSetName = "ContentDirectory")]
    [Parameter(ParameterSetName = "ContentLiteralPath")]
    [Parameter(ParameterSetName = "ContentLiteralPathDirectory")]
    [Parameter(ParameterSetName = "ContentDocument")]
    [Parameter(ParameterSetName = "Page")]
    [Parameter(ParameterSetName = "PageDirectory")]
    [Parameter(ParameterSetName = "PageLiteralPath")]
    [Parameter(ParameterSetName = "PageLiteralPathDirectory")]
    [Parameter(ParameterSetName = "PageDocument")]
    [ValidateRange(0D, 1D)]
    public double Opacity { get; set; } = 1D;

    /// <summary>Place the stamp behind existing page content.</summary>
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "TextDirectory")]
    [Parameter(ParameterSetName = "TextLiteralPath")]
    [Parameter(ParameterSetName = "TextLiteralPathDirectory")]
    [Parameter(ParameterSetName = "TextDocument")]
    [Parameter(ParameterSetName = "Image")]
    [Parameter(ParameterSetName = "ImageDirectory")]
    [Parameter(ParameterSetName = "ImageLiteralPath")]
    [Parameter(ParameterSetName = "ImageLiteralPathDirectory")]
    [Parameter(ParameterSetName = "ImageDocument")]
    [Parameter(ParameterSetName = "Content")]
    [Parameter(ParameterSetName = "ContentDirectory")]
    [Parameter(ParameterSetName = "ContentLiteralPath")]
    [Parameter(ParameterSetName = "ContentLiteralPathDirectory")]
    [Parameter(ParameterSetName = "ContentDocument")]
    [Parameter(ParameterSetName = "Page")]
    [Parameter(ParameterSetName = "PageDirectory")]
    [Parameter(ParameterSetName = "PageLiteralPath")]
    [Parameter(ParameterSetName = "PageLiteralPathDirectory")]
    [Parameter(ParameterSetName = "PageDocument")]
    public SwitchParameter Behind { get; set; }

    /// <summary>Source PDF password.</summary>
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "TextDirectory")]
    [Parameter(ParameterSetName = "TextLiteralPath")]
    [Parameter(ParameterSetName = "TextLiteralPathDirectory")]
    [Parameter(ParameterSetName = "Image")]
    [Parameter(ParameterSetName = "ImageDirectory")]
    [Parameter(ParameterSetName = "ImageLiteralPath")]
    [Parameter(ParameterSetName = "ImageLiteralPathDirectory")]
    [Parameter(ParameterSetName = "Content")]
    [Parameter(ParameterSetName = "ContentDirectory")]
    [Parameter(ParameterSetName = "ContentLiteralPath")]
    [Parameter(ParameterSetName = "ContentLiteralPathDirectory")]
    [Parameter(ParameterSetName = "Page")]
    [Parameter(ParameterSetName = "PageDirectory")]
    [Parameter(ParameterSetName = "PageLiteralPath")]
    [Parameter(ParameterSetName = "PageLiteralPathDirectory")]
    [OfficePasswordTransformation]
    public SecureString? Password { get; set; }

    /// <summary>After authentication, explicitly ignore owner-imposed restrictions.</summary>
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "TextDirectory")]
    [Parameter(ParameterSetName = "TextLiteralPath")]
    [Parameter(ParameterSetName = "TextLiteralPathDirectory")]
    [Parameter(ParameterSetName = "Image")]
    [Parameter(ParameterSetName = "ImageDirectory")]
    [Parameter(ParameterSetName = "ImageLiteralPath")]
    [Parameter(ParameterSetName = "ImageLiteralPathDirectory")]
    [Parameter(ParameterSetName = "Content")]
    [Parameter(ParameterSetName = "ContentDirectory")]
    [Parameter(ParameterSetName = "ContentLiteralPath")]
    [Parameter(ParameterSetName = "ContentLiteralPathDirectory")]
    [Parameter(ParameterSetName = "Page")]
    [Parameter(ParameterSetName = "PageDirectory")]
    [Parameter(ParameterSetName = "PageLiteralPath")]
    [Parameter(ParameterSetName = "PageLiteralPathDirectory")]
    public SwitchParameter IgnorePermissionRestrictions { get; set; }

    /// <summary>Native input parsing, authentication and resource policy.</summary>
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "TextDirectory")]
    [Parameter(ParameterSetName = "TextLiteralPath")]
    [Parameter(ParameterSetName = "TextLiteralPathDirectory")]
    [Parameter(ParameterSetName = "Image")]
    [Parameter(ParameterSetName = "ImageDirectory")]
    [Parameter(ParameterSetName = "ImageLiteralPath")]
    [Parameter(ParameterSetName = "ImageLiteralPathDirectory")]
    [Parameter(ParameterSetName = "Content")]
    [Parameter(ParameterSetName = "ContentDirectory")]
    [Parameter(ParameterSetName = "ContentLiteralPath")]
    [Parameter(ParameterSetName = "ContentLiteralPathDirectory")]
    [Parameter(ParameterSetName = "Page")]
    [Parameter(ParameterSetName = "PageDirectory")]
    [Parameter(ParameterSetName = "PageLiteralPath")]
    [Parameter(ParameterSetName = "PageLiteralPathDirectory")]
    [OfficeOptionsTransformation(typeof(PdfLoadOptions))]
    public PdfLoadOptions? LoadOptions { get; set; }

    /// <summary>Return the saved FileInfo or the updated live document.</summary>
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "TextDirectory")]
    [Parameter(ParameterSetName = "TextLiteralPath")]
    [Parameter(ParameterSetName = "TextLiteralPathDirectory")]
    [Parameter(ParameterSetName = "TextDocument")]
    [Parameter(ParameterSetName = "Image")]
    [Parameter(ParameterSetName = "ImageDirectory")]
    [Parameter(ParameterSetName = "ImageLiteralPath")]
    [Parameter(ParameterSetName = "ImageLiteralPathDirectory")]
    [Parameter(ParameterSetName = "ImageDocument")]
    [Parameter(ParameterSetName = "Content")]
    [Parameter(ParameterSetName = "ContentDirectory")]
    [Parameter(ParameterSetName = "ContentLiteralPath")]
    [Parameter(ParameterSetName = "ContentLiteralPathDirectory")]
    [Parameter(ParameterSetName = "ContentDocument")]
    [Parameter(ParameterSetName = "Page")]
    [Parameter(ParameterSetName = "PageDirectory")]
    [Parameter(ParameterSetName = "PageLiteralPath")]
    [Parameter(ParameterSetName = "PageLiteralPathDirectory")]
    [Parameter(ParameterSetName = "PageDocument")]
    public SwitchParameter PassThru { get; set; }

    /// <summary>Output folder retaining source file names.</summary>
    [Parameter(ParameterSetName = "TextDirectory", Mandatory = true)]
    [Parameter(ParameterSetName = "TextLiteralPathDirectory", Mandatory = true)]
    [Parameter(ParameterSetName = "ImageDirectory", Mandatory = true)]
    [Parameter(ParameterSetName = "ImageLiteralPathDirectory", Mandatory = true)]
    [Parameter(ParameterSetName = "ContentDirectory", Mandatory = true)]
    [Parameter(ParameterSetName = "ContentLiteralPathDirectory", Mandatory = true)]
    [Parameter(ParameterSetName = "PageDirectory", Mandatory = true)]
    [Parameter(ParameterSetName = "PageLiteralPathDirectory", Mandatory = true)]
    public string? OutputDirectory { get; set; }

    /// <summary>Literal PDF files, including FileInfo.FullName.</summary>
    [Parameter(ParameterSetName = "TextLiteralPath", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Parameter(ParameterSetName = "TextLiteralPathDirectory", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Parameter(ParameterSetName = "ImageLiteralPath", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Parameter(ParameterSetName = "ImageLiteralPathDirectory", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Parameter(ParameterSetName = "ContentLiteralPath", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Parameter(ParameterSetName = "ContentLiteralPathDirectory", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Parameter(ParameterSetName = "PageLiteralPath", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Parameter(ParameterSetName = "PageLiteralPathDirectory", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Alias("FullName", "PSPath"), ValidateNotNullOrEmpty]
    public string[] LiteralPath { get; set; } = Array.Empty<string>();

    /// <summary>Caller-owned PDF; omission uses the enclosing editing block.</summary>
    [Parameter(ParameterSetName = "TextDocument", ValueFromPipeline = true)]
    [Parameter(ParameterSetName = "ImageDocument", ValueFromPipeline = true)]
    [Parameter(ParameterSetName = "ContentDocument", ValueFromPipeline = true)]
    [Parameter(ParameterSetName = "PageDocument", ValueFromPipeline = true)]
    public PdfDocument? Document { get; set; }

    /// <summary>Literal image file to stamp.</summary>
    [Parameter(ParameterSetName = "Image", Mandatory = true)]
    [Parameter(ParameterSetName = "ImageDirectory", Mandatory = true)]
    [Parameter(ParameterSetName = "ImageLiteralPath", Mandatory = true)]
    [Parameter(ParameterSetName = "ImageLiteralPathDirectory", Mandatory = true)]
    [Parameter(ParameterSetName = "ImageDocument", Mandatory = true)]
    public string? ImagePath { get; set; }

    /// <summary>Native Width setting for the stamp.</summary>
    [Parameter(ParameterSetName = "Image")]
    [Parameter(ParameterSetName = "ImageDirectory")]
    [Parameter(ParameterSetName = "ImageLiteralPath")]
    [Parameter(ParameterSetName = "ImageLiteralPathDirectory")]
    [Parameter(ParameterSetName = "ImageDocument")]
    [Parameter(ParameterSetName = "Page")]
    [Parameter(ParameterSetName = "PageDirectory")]
    [Parameter(ParameterSetName = "PageLiteralPath")]
    [Parameter(ParameterSetName = "PageLiteralPathDirectory")]
    [Parameter(ParameterSetName = "PageDocument")]
    [OfficeLengthTransformation]
    public OfficeLength? Width { get; set; }

    /// <summary>Native Height setting for the stamp.</summary>
    [Parameter(ParameterSetName = "Image")]
    [Parameter(ParameterSetName = "ImageDirectory")]
    [Parameter(ParameterSetName = "ImageLiteralPath")]
    [Parameter(ParameterSetName = "ImageLiteralPathDirectory")]
    [Parameter(ParameterSetName = "ImageDocument")]
    [Parameter(ParameterSetName = "Page")]
    [Parameter(ParameterSetName = "PageDirectory")]
    [Parameter(ParameterSetName = "PageLiteralPath")]
    [Parameter(ParameterSetName = "PageLiteralPathDirectory")]
    [Parameter(ParameterSetName = "PageDocument")]
    [OfficeLengthTransformation]
    public OfficeLength? Height { get; set; }

    /// <summary>Native canvas callback receiving the canvas and target page context.</summary>
    [Parameter(ParameterSetName = "Content", Mandatory = true)]
    [Parameter(ParameterSetName = "ContentDirectory", Mandatory = true)]
    [Parameter(ParameterSetName = "ContentLiteralPath", Mandatory = true)]
    [Parameter(ParameterSetName = "ContentLiteralPathDirectory", Mandatory = true)]
    [Parameter(ParameterSetName = "ContentDocument", Mandatory = true, Position = 0)]
    public ScriptBlock? Content { get; set; }

    /// <summary>Configure native rendering options for the canvas callback.</summary>
    [Parameter(ParameterSetName = "Content")]
    [Parameter(ParameterSetName = "ContentDirectory")]
    [Parameter(ParameterSetName = "ContentLiteralPath")]
    [Parameter(ParameterSetName = "ContentLiteralPathDirectory")]
    [Parameter(ParameterSetName = "ContentDocument")]
    public ScriptBlock? ConfigureRendering { get; set; }

    /// <summary>Literal PDF file whose selected page is imported.</summary>
    [Parameter(ParameterSetName = "Page", Mandatory = true)]
    [Parameter(ParameterSetName = "PageDirectory", Mandatory = true)]
    [Parameter(ParameterSetName = "PageLiteralPath", Mandatory = true)]
    [Parameter(ParameterSetName = "PageLiteralPathDirectory", Mandatory = true)]
    [Parameter(ParameterSetName = "PageDocument", Mandatory = true)]
    public string? OverlayPath { get; set; }

    /// <summary>One-based page imported from the overlay PDF.</summary>
    [Parameter(ParameterSetName = "Page")]
    [Parameter(ParameterSetName = "PageDirectory")]
    [Parameter(ParameterSetName = "PageLiteralPath")]
    [Parameter(ParameterSetName = "PageLiteralPathDirectory")]
    [Parameter(ParameterSetName = "PageDocument")]
    [ValidateRange(1, int.MaxValue)]
    public int OverlayPage { get; set; } = 1;

    /// <summary>Native Fit setting for the stamp.</summary>
    [Parameter(ParameterSetName = "Page")]
    [Parameter(ParameterSetName = "PageDirectory")]
    [Parameter(ParameterSetName = "PageLiteralPath")]
    [Parameter(ParameterSetName = "PageLiteralPathDirectory")]
    [Parameter(ParameterSetName = "PageDocument")]
    public PdfPageOverlayFit Fit { get; set; } = PdfPageOverlayFit.Contain;

    /// <summary>Native HorizontalAlignment setting for the stamp.</summary>
    [Parameter(ParameterSetName = "Page")]
    [Parameter(ParameterSetName = "PageDirectory")]
    [Parameter(ParameterSetName = "PageLiteralPath")]
    [Parameter(ParameterSetName = "PageLiteralPathDirectory")]
    [Parameter(ParameterSetName = "PageDocument")]
    public OfficeAlignment HorizontalAlignment { get; set; } = OfficeAlignment.Center;

    /// <summary>Native VerticalAlignment setting for the stamp.</summary>
    [Parameter(ParameterSetName = "Page")]
    [Parameter(ParameterSetName = "PageDirectory")]
    [Parameter(ParameterSetName = "PageLiteralPath")]
    [Parameter(ParameterSetName = "PageLiteralPathDirectory")]
    [Parameter(ParameterSetName = "PageDocument")]
    public OfficeTextVerticalAlignment VerticalAlignment { get; set; } = OfficeTextVerticalAlignment.Center;

    /// <summary>Password for the imported overlay PDF.</summary>
    [Parameter(ParameterSetName = "Page")]
    [Parameter(ParameterSetName = "PageDirectory")]
    [Parameter(ParameterSetName = "PageLiteralPath")]
    [Parameter(ParameterSetName = "PageLiteralPathDirectory")]
    [Parameter(ParameterSetName = "PageDocument")]
    [OfficePasswordTransformation]
    public SecureString? OverlayPassword { get; set; }

    /// <summary>After overlay authentication, explicitly ignore owner-imposed restrictions.</summary>
    [Parameter(ParameterSetName = "Page")]
    [Parameter(ParameterSetName = "PageDirectory")]
    [Parameter(ParameterSetName = "PageLiteralPath")]
    [Parameter(ParameterSetName = "PageLiteralPathDirectory")]
    [Parameter(ParameterSetName = "PageDocument")]
    public SwitchParameter IgnoreOverlayPermissionRestrictions { get; set; }

    /// <summary>Native loading policy for the imported overlay PDF.</summary>
    [Parameter(ParameterSetName = "Page")]
    [Parameter(ParameterSetName = "PageDirectory")]
    [Parameter(ParameterSetName = "PageLiteralPath")]
    [Parameter(ParameterSetName = "PageLiteralPathDirectory")]
    [Parameter(ParameterSetName = "PageDocument")]
    [OfficeOptionsTransformation(typeof(PdfLoadOptions))]
    public PdfLoadOptions? OverlayLoadOptions { get; set; }

}
