using System.Management.Automation;
using OfficeIMO.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Updates OfficeIMO Word image sizing, wrapping, crop, and metadata.</summary>
/// <example>
///   <summary>Resize and describe all report images.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$doc = Get-OfficeWord -Path .\Report.docx
/// $doc |
///     Get-OfficeWordImage |
///     Set-OfficeWordImage -Width 320 -Wrap Square -Title 'Report image' -Description 'Image used in the report'
/// $doc | Save-OfficeWord -Path .\Report-Images.docx</code>
///   <para>Gets OfficeIMO image objects from an open document, applies thin image metadata and layout updates, and saves the document.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeWordImage")]
[Alias("WordImageStyle")]
[OutputType(typeof(WordImage))]
public sealed class SetOfficeWordImageCommand : PSCmdlet
{
    /// <summary>Image to update.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0)]
    public WordImage Image { get; set; } = null!;

    /// <summary>Image width in points.</summary>
    [Parameter] public double? Width { get; set; }

    /// <summary>Image height in points.</summary>
    [Parameter] public double? Height { get; set; }

    /// <summary>Text wrapping mode.</summary>
    [Parameter] public WordImageTextWrapping? Wrap { get; set; }

    /// <summary>Top crop value.</summary>
    [Parameter] public int? CropTop { get; set; }

    /// <summary>Bottom crop value.</summary>
    [Parameter] public int? CropBottom { get; set; }

    /// <summary>Left crop value.</summary>
    [Parameter] public int? CropLeft { get; set; }

    /// <summary>Right crop value.</summary>
    [Parameter] public int? CropRight { get; set; }

    /// <summary>Image rotation in degrees.</summary>
    [Parameter] public int? Rotation { get; set; }

    /// <summary>Fixed opacity percentage.</summary>
    [Parameter] public int? Opacity { get; set; }

    /// <summary>Horizontally flip the image.</summary>
    [Parameter] public bool? HorizontalFlip { get; set; }

    /// <summary>Vertically flip the image.</summary>
    [Parameter] public bool? VerticalFlip { get; set; }

    /// <summary>Image title metadata.</summary>
    [Parameter] public string? Title { get; set; }

    /// <summary>Image alternate text metadata.</summary>
    [Parameter] public string? Description { get; set; }

    /// <summary>Whether the image is hidden.</summary>
    [Parameter] public bool? Hidden { get; set; }

    /// <summary>Emit the updated image.</summary>
    [Parameter] public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (Image == null)
        {
            return;
        }

        if (MyInvocation.BoundParameters.ContainsKey(nameof(Width))) Image.Width = Width;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(Height))) Image.Height = Height;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(Wrap))) Image.WrapText = Wrap;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(CropTop))) Image.CropTop = CropTop;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(CropBottom))) Image.CropBottom = CropBottom;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(CropLeft))) Image.CropLeft = CropLeft;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(CropRight))) Image.CropRight = CropRight;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(Rotation))) Image.Rotation = Rotation;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(Opacity))) Image.FixedOpacity = Opacity;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(HorizontalFlip))) Image.HorizontalFlip = HorizontalFlip;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(VerticalFlip))) Image.VerticalFlip = VerticalFlip;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(Title))) Image.Title = Title;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(Description))) Image.Description = Description;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(Hidden))) Image.Hidden = Hidden;

        if (PassThru.IsPresent)
        {
            WriteObject(Image);
        }
    }
}
