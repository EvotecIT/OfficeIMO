using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Adds a bulleted list to a PowerPoint slide.</summary>
/// <para>Creates a textbox and populates it with bullet paragraphs.</para>
/// <example>
///   <summary>Add a bullet list.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficePowerPoint -Path .\Examples\Documents\PowerPointBullets.pptx {
///     $slide = Add-OfficePowerPointSlide -Layout 1 -PassThru
///     Set-OfficePowerPointSlideTitle -Slide $slide -Title 'Delivery update'
///     Add-OfficePowerPointBullets -Slide $slide -Bullets 'Wins','Risks','Next steps' -X 60 -Y 120 -Width 420 -Height 180
/// }</code>
///   <para>Creates a slide with a titled bullet list.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficePowerPointBullets")]
[Alias("PptBullets")]
[OutputType(typeof(PowerPointTextBox))]
public sealed class AddOfficePowerPointBulletsCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    /// <summary>Target slide that will receive the bullet list (optional inside DSL).</summary>
    [Parameter(ValueFromPipeline = true)]
    public PowerPointSlide? Slide { get; set; }

    /// <summary>Bullet items to render.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string[] Bullets { get; set; } = Array.Empty<string>();

    /// <summary>Left offset (in points) from the slide origin.</summary>
    [Parameter]
    public double X { get; set; } = 60;

    /// <summary>Top offset (in points) from the slide origin.</summary>
    [Parameter]
    public double Y { get; set; } = 120;

    /// <summary>Textbox width in points.</summary>
    [Parameter]
    public double Width { get; set; } = 400;

    /// <summary>Textbox height in points.</summary>
    [Parameter]
    public double Height { get; set; } = 200;

    /// <summary>List level (0-8).</summary>
    [Parameter]
    public int Level { get; set; }

    /// <summary>Optional bullet character (defaults to •).</summary>
    [Parameter]
    public string? BulletChar { get; set; }

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
            var items = NormalizeItems(Bullets);
            var textBox = slide.AddTextBoxPoints(string.Empty, X, Y, Width, Height);
            textBox.SetBullets(items, Level, ResolveBulletChar());
            WritePassThru(textBox);
        } catch (Exception ex) {
            WriteError(new ErrorRecord(ex, "PowerPointAddBulletsFailed", ErrorCategory.InvalidOperation, Slide));
        }
    }

    private static List<string> NormalizeItems(string[]? items) {
        if (items == null || items.Length == 0) {
            throw new PSArgumentException("Bullets cannot be empty.", nameof(Bullets));
        }

        var list = items
            .Select(item => item?.Trim())
            .Where(item => !string.IsNullOrWhiteSpace(item))
            .Cast<string>()
            .ToList();

        if (list.Count == 0) {
            throw new PSArgumentException("Bullets cannot be empty.", nameof(Bullets));
        }

        return list;
    }

    private char ResolveBulletChar() {
        if (string.IsNullOrWhiteSpace(BulletChar)) {
            return '\u2022';
        }

        return BulletChar!.Trim()[0];
    }
}
