using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Adds a picture content control to the current paragraph.</summary>
/// <example>
///   <summary>Add a picture control.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Add-OfficeWordParagraph { Add-OfficeWordPictureControl -Path .\logo.png -Width 96 -Height 32 }</code>
///   <para>Embeds an image inside a picture content control.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeWordPictureControl")]
[Alias("WordPictureControl")]
[OutputType(typeof(WordPictureControl))]
public sealed class AddOfficeWordPictureControlCommand : PSCmdlet
{
    /// <summary>Path to the image file.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Path { get; set; } = string.Empty;

    /// <summary>Width in points.</summary>
    [Parameter]
    public double? Width { get; set; }

    /// <summary>Height in points.</summary>
    [Parameter]
    public double? Height { get; set; }

    /// <summary>Optional alias for the control.</summary>
    [Parameter]
    public string? Alias { get; set; }

    /// <summary>Optional tag for the control.</summary>
    [Parameter]
    public string? Tag { get; set; }

    /// <summary>Explicit paragraph to receive the control.</summary>
    [Parameter(ValueFromPipeline = true)]
    public WordParagraph? Paragraph { get; set; }

    /// <summary>Emit the created control.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var paragraph = ResolveParagraph();
        var resolvedPath = ResolvePath();
        var control = paragraph.AddPictureControl(resolvedPath, Width, Height, Alias, Tag);

        if (PassThru.IsPresent)
        {
            WriteObject(control);
        }
    }

    private string ResolvePath()
    {
        var providerPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
        return System.IO.Path.IsPathRooted(providerPath)
            ? providerPath
            : System.IO.Path.Combine(SessionState.Path.CurrentFileSystemLocation.Path, providerPath);
    }

    private WordParagraph ResolveParagraph()
    {
        if (Paragraph != null)
        {
            return Paragraph;
        }

        var context = WordDslContext.Require(this);
        return context.CurrentParagraph ?? context.RequireParagraphHost().AddParagraph();
    }
}
