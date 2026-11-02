using System.Management.Automation;
using OfficeIMO.Html.Pdf;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Creates native semantic or positioned PDF-to-HTML export settings.</summary>
/// <example><summary>Create positioned review HTML with the source links and form fields.</summary>
/// <code>$options = New-OfficePdfHtmlOptions -Profile PositionedReview
/// Convert-OfficeDocument -Path ./Review.pdf -To Html -OutputPath ./Review.html -Options $options</code></example>
/// <example><summary>Export a fragment without source metadata or page containers.</summary>
/// <code>$options = New-OfficePdfHtmlOptions -Fragment -NoMetadata -NoPageContainers
/// Convert-OfficeDocument -Path ./Report.pdf -To Html -AsString -Options $options</code></example>
[Cmdlet(VerbsCommon.New, "OfficePdfHtmlOptions", DefaultParameterSetName = "Default")]
[OutputType(typeof(PdfToHtmlOptions))]
public sealed class NewOfficePdfHtmlOptionsCommand : OfficeCmdlet {
    /// <summary>Semantic content or positioned, inert review HTML; semantic is the default.</summary>
    [Parameter(ParameterSetName = "Default")]
    public PdfHtmlProfile Profile { get; set; } = PdfHtmlProfile.Semantic;
    /// <summary>Use embedded images or placeholders for detected PDF images.</summary>
    [Parameter]
    public PdfHtmlImageExportMode? ImageExportMode { get; set; }
    /// <summary>Maximum source image bytes embedded in the HTML; zero embeds no nonempty image.</summary>
    [Parameter, ValidateRange(0, long.MaxValue)]
    public long? MaxEmbeddedImageSize { get; set; }
    /// <summary>Exclude source document metadata.</summary>
    [Parameter]
    public SwitchParameter NoMetadata { get; set; }
    /// <summary>Omit semantic page wrappers and page-number metadata.</summary>
    [Parameter]
    public SwitchParameter NoPageContainers { get; set; }
    /// <summary>Exclude readable image projections.</summary>
    [Parameter]
    public SwitchParameter NoImagePlaceholders { get; set; }
    /// <summary>Include inert link annotations; positioned review enables them by default.</summary>
    [Parameter]
    public SwitchParameter IncludeLinkAnnotations { get; set; }
    /// <summary>Include inert form-widget projections; positioned review enables them by default.</summary>
    [Parameter]
    public SwitchParameter IncludeFormWidgets { get; set; }
    /// <summary>Return an HTML fragment without the complete document shell.</summary>
    [Parameter]
    public SwitchParameter Fragment { get; set; }
    /// <summary>Document title used when the source PDF has no title.</summary>
    [Parameter, AllowEmptyString, ValidateNotNull]
    public string? DocumentTitleFallback { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (!System.Enum.IsDefined(typeof(PdfHtmlProfile), Profile))
            throw new PSArgumentException("Unsupported PDF HTML profile.");
        if (ImageExportMode.HasValue && !System.Enum.IsDefined(typeof(PdfHtmlImageExportMode), ImageExportMode.Value))
            throw new PSArgumentException("Unsupported PDF HTML image export mode.");
        var options = Profile == PdfHtmlProfile.PositionedReview
            ? PdfToHtmlOptions.CreatePositionedReviewProfile()
            : PdfToHtmlOptions.CreateSemanticProfile();
        options.IncludeMetadata = !NoMetadata.IsPresent;
        options.IncludePageContainers = !NoPageContainers.IsPresent;
        options.IncludeImagePlaceholders = !NoImagePlaceholders.IsPresent;
        options.EmitDocumentShell = !Fragment.IsPresent;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(IncludeLinkAnnotations)))
            options.IncludeLinkAnnotations = IncludeLinkAnnotations.IsPresent;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(IncludeFormWidgets)))
            options.IncludeFormWidgets = IncludeFormWidgets.IsPresent;
        if (ImageExportMode.HasValue) options.ImageExportMode = ImageExportMode.Value;
        if (MaxEmbeddedImageSize.HasValue) options.MaxEmbeddedImageBytes = MaxEmbeddedImageSize.Value;
        if (DocumentTitleFallback != null) options.DocumentTitleFallback = DocumentTitleFallback;
        WriteObject(options);
    }
}
