using System.IO;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Inserts an image into the current paragraph.</summary>
/// <para>Supports width/height overrides and wrapping options using the alias-friendly DSL.</para>
/// <example>
///   <summary>Add a logo.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Add-OfficeWordParagraph { Add-OfficeWordImage -Path .\logo.png -Width 96 -Height 32 }</code>
///   <para>Embeds <c>logo.png</c> at the specified size.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeWordImage")]
[Alias("WordImage")]
public sealed class AddOfficeWordImageCommand : PSCmdlet
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

    /// <summary>Wrap mode for the image.</summary>
    [Parameter]
    public WordImageTextWrapping Wrap { get; set; } = WordImageTextWrapping.InLineWithText;

    /// <summary>Optional description/alt text.</summary>
    [Parameter]
    public string Description { get; set; } = string.Empty;

    /// <summary>Emit the created <see cref="WordImage"/>.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var context = WordDslContext.Require(this);
        var paragraph = context.CurrentParagraph ?? context.RequireParagraphHost().AddParagraph();
        var fullPath = ResolvePath();
        var image = paragraph.InsertImage(fullPath, Width, Height, Wrap, Description);

        if (PassThru.IsPresent)
        {
            WriteObject(image);
        }
    }

    private string ResolvePath()
    {
        var providerPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
        return System.IO.Path.IsPathRooted(providerPath)
            ? providerPath
            : System.IO.Path.Combine(SessionState.Path.CurrentFileSystemLocation.Path, providerPath);
    }
}
