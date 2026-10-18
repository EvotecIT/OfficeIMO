using System.Management.Automation;
using OfficeIMO.ChartForgeX;
using OfficeIMO.Word;
using PSWriteOffice.Services.Visuals;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Adds a ChartForgeX artifact, portable SVG, or converted Office visual to Word.</summary>
/// <example>
///   <summary>Pipe an ImagePlayground artifact into a Word paragraph.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>WordParagraph { $artifact | Add-OfficeWordVisual -Width 420 }</code>
///   <para>Embeds the selected SVG or PNG payload with the artifact's accessible description.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeWordVisual")]
[Alias("WordVisual")]
[OutputType(typeof(WordImage))]
public sealed class AddOfficeWordVisualCommand : OfficeVisualCommandBase {
    /// <summary>ChartForgeX VisualArtifact, OfficeVisualSource, OfficeVisualConversionResult, or SVG file path.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    public object InputObject { get; set; } = null!;

    /// <summary>Target paragraph. Inside the Word DSL, the current paragraph is used by default.</summary>
    [Parameter]
    public WordParagraph? Paragraph { get; set; }

    /// <summary>Word text-wrapping behavior.</summary>
    [Parameter]
    public WordImageTextWrapping Wrap { get; set; } = WordImageTextWrapping.InLineWithText;

    /// <summary>Emit the image added to the paragraph.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        WordParagraph paragraph = Paragraph ?? ResolveParagraph();
        var image = paragraph.AddVisualArtifact(ResolveVisual(InputObject), Wrap);
        if (PassThru.IsPresent) {
            WriteObject(image);
        }
    }

    private WordParagraph ResolveParagraph() {
        WordDslContext context = WordDslContext.Require(this);
        return context.CurrentParagraph ?? context.RequireParagraphHost().AddParagraph();
    }
}