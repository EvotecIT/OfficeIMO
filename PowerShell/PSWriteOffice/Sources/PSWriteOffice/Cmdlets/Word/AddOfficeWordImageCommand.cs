using System;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.ChartForgeX;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Visuals;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Embeds an image file, SVG or reusable visual in a Word paragraph.</summary>
/// <para>Numbers mean points; unit suffixes such as 4cm are accepted. A prepared visual is reused without changing its payload or metadata. One dimension preserves proportions; two dimensions set the displayed frame.</para>
/// <example><summary>Insert a local logo.</summary>
/// <code>WordParagraph { WordImage -Path ./logo.png -Width 4cm -Description 'Contoso' }</code></example>
/// <example><summary>Place a reusable visual in an existing paragraph.</summary>
/// <code>$visual | WordImage -Paragraph $paragraph -Width 12cm -Title 'Availability' -PassThru</code></example>
[Cmdlet(VerbsCommon.Add, "OfficeWordImage", DefaultParameterSetName = "Path")]
[Alias("WordImage")]
[OutputType(typeof(WordImage))]
public sealed class AddOfficeWordImageCommand : OfficeVisualCommandBase {
    /// <summary>Literal local image path.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Path"), ValidateNotNullOrEmpty]
    public string? Path { get; set; }

    /// <summary>ChartForgeX artifact, OfficeVisualSource, prepared visual or SVG path.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "Visual"), ValidateNotNull]
    public object? InputObject { get; set; }

    /// <summary>Target paragraph; the current DSL paragraph is used when omitted.</summary>
    [Parameter(ValueFromPipeline = true)]
    public WordParagraph? Paragraph { get; set; }

    /// <summary>Text wrapping around the embedded image.</summary>
    [Parameter]
    public WordImageTextWrapping Wrap { get; set; } = WordImageTextWrapping.InLineWithText;

    /// <summary>Accessible image description; omitted visual descriptions are retained.</summary>
    [Parameter]
    public string? Description { get; set; }

    /// <inheritdoc />
    protected override string? VisualAlternativeText => Description;

    /// <inheritdoc />
    protected override string AlternativeTextParameterName => nameof(Description);

    /// <summary>Return the native editable WordImage.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (Width.HasValue && Width.Value <= OfficeLength.Zero) throw new PSArgumentOutOfRangeException(nameof(Width));
        if (Height.HasValue && Height.Value <= OfficeLength.Zero) throw new PSArgumentOutOfRangeException(nameof(Height));
        string? path = Path == null ? null : OfficePath.ResolveInput(SessionState, Path, literal: true)[0];
        OfficeVisualConversionResult? visual = ParameterSetName == "Visual" ? ResolveVisual(InputObject!, nativePlacement: true) : null;
        if (visual != null && (Width.HasValue || Height.HasValue)) {
            double width = Width?.ToPoints() ?? visual.WidthPoints * Height!.Value.ToPoints() / visual.HeightPoints;
            visual = visual.WithSize(width, Height?.ToPoints(), OfficeImageFit.Stretch);
        }
        WordParagraph paragraph = Paragraph ?? ResolveParagraph();
        WordImage image = visual == null
            ? paragraph.AddImage(path!, Width, Height, new WordImageOptions { Wrapping = Wrap, Description = Description ?? string.Empty, Title = Title })
            : paragraph.AddVisualArtifact(visual, Wrap);
        if (Description != null) image.Description = Description;
        if (Title != null) image.Title = Title;
        if (PassThru) WriteObject(image);
    }

    private WordParagraph ResolveParagraph() {
        var context = WordDslContext.Require(this);
        return context.CurrentParagraph ?? context.RequireParagraphHost().AddParagraph();
    }
}
