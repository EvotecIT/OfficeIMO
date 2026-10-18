using System;
using System.Management.Automation;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services;
using PSWriteOffice.Services.PowerPoint;
using A = DocumentFormat.OpenXml.Drawing;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Sets layout placeholder text style and bullet/numbering settings.</summary>
/// <example>
///   <summary>Apply a title preset style.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Set-OfficePowerPointLayoutPlaceholderTextStyle -Presentation $ppt -Master 0 -Layout 1 -PlaceholderType Title -Style Title</code>
///   <para>Applies the Title preset to the layout placeholder.</para>
/// </example>
/// <example>
///   <summary>Apply a style inside the DSL.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficePowerPoint -Path .\deck.pptx {
///     $layout = Get-OfficePowerPointLayout | Select-Object -First 1
///     Set-OfficePowerPointLayoutPlaceholderTextStyle -Master $layout.MasterIndex -Layout $layout.LayoutIndex -PlaceholderType Title -Style Title -FontSize 36 -Bold $true
///   }</code>
///   <para>Uses the DSL context to resolve the presentation.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficePowerPointLayoutPlaceholderTextStyle")]
[Alias("PptLayoutPlaceholderTextStyle")]
[OutputType(typeof(PowerPointTextBox))]
public sealed class SetOfficePowerPointLayoutPlaceholderTextStyleCommand : PSCmdlet
{
    /// <summary>Presentation to update (optional inside DSL).</summary>
    [Parameter(ValueFromPipeline = true)]
    public PowerPointPresentation? Presentation { get; set; }

    /// <summary>Slide master index.</summary>
    [Parameter]
    [ValidateRange(0, int.MaxValue)]
    public int Master { get; set; } = 0;

    /// <summary>Layout index within the master.</summary>
    [Parameter(Mandatory = true)]
    [ValidateRange(0, int.MaxValue)]
    public int Layout { get; set; }

    /// <summary>Placeholder type to target.</summary>
    [Parameter(Mandatory = true)]
    [Alias("Type")]
    public PowerPointPlaceholderType PlaceholderType { get; set; }

    /// <summary>Optional placeholder index.</summary>
    [Parameter]
    public uint? Index { get; set; }

    /// <summary>Named style preset (Title, Subtitle, Body, Caption, Emphasis).</summary>
    [Parameter]
    [ValidateSet("Title", "Subtitle", "Body", "Caption", "Emphasis")]
    public string? Style { get; set; }

    /// <summary>Font size in points.</summary>
    [Parameter]
    public int? FontSize { get; set; }

    /// <summary>Font name (Latin).</summary>
    [Parameter]
    public string? FontName { get; set; }

    /// <summary>Text color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? Color { get; set; }

    /// <summary>Apply bold formatting.</summary>
    [Parameter]
    public bool? Bold { get; set; }

    /// <summary>Apply italic formatting.</summary>
    [Parameter]
    public bool? Italic { get; set; }

    /// <summary>Apply underline formatting.</summary>
    [Parameter]
    public bool? Underline { get; set; }

    /// <summary>Highlight color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? HighlightColor { get; set; }

    /// <summary>Paragraph level (0-8) to set before applying style.</summary>
    [Parameter]
    public int? Level { get; set; }

    /// <summary>Optional bullet character (ignored when -Numbering is supplied).</summary>
    [Parameter]
    public string? BulletChar { get; set; }

    /// <summary>Optional numbering scheme name (e.g. ArabicPeriod, RomanUpper).</summary>
    [Parameter]
    public PowerPointNumberingScheme? Numbering { get; set; }

    /// <summary>Create the placeholder if it is missing.</summary>
    [Parameter]
    public SwitchParameter CreateIfMissing { get; set; }

    /// <summary>Emit the placeholder textbox after update.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        PowerPointPresentation? presentation = null;
        try
        {
            if (Numbering.HasValue && !string.IsNullOrWhiteSpace(BulletChar))
            {
                throw new PSArgumentException("Specify either Numbering or BulletChar, not both.");
            }

            bool hasRequestedChanges = TryResolveStyle(out var style);
            if (!hasRequestedChanges)
            {
                throw new PSArgumentException("Specify a style name, style properties, or bullet/numbering settings.");
            }

            PowerPointNumberingScheme? numbering = ResolveNumbering();
            char? bulletChar = ResolveBulletChar();

            presentation = Presentation ?? PowerPointDslContext.Require(this).Presentation;
            presentation.SetLayoutPlaceholderTextStyle(
                Master,
                Layout,
                PlaceholderType,
                style,
                Index,
                Level,
                bulletChar,
                numbering,
                CreateIfMissing.IsPresent);

            if (PassThru.IsPresent)
            {
                var textBox = presentation.GetLayoutPlaceholderTextBox(Master, Layout, PlaceholderType, Index);
                if (textBox != null)
                {
                    WriteObject(textBox);
                }
            }
        }
        catch (Exception ex)
        {
            WriteError(new ErrorRecord(ex, "PowerPointSetLayoutPlaceholderTextStyleFailed", ErrorCategory.InvalidOperation, presentation ?? Presentation));
        }
    }

    private bool TryResolveStyle(out PowerPointTextStyle style)
    {
        style = new PowerPointTextStyle();
        bool hasStyleName = !string.IsNullOrWhiteSpace(Style);
        if (hasStyleName)
        {
            if (!OpenXmlValueParser.TryParse(Style, out style))
            {
                throw new PSArgumentException($"Unknown style '{Style}'.", nameof(Style));
            }
        }

        if (FontSize != null)
        {
            style = style.WithFontSize(FontSize);
        }

        if (!string.IsNullOrWhiteSpace(FontName))
        {
            style = style.WithFontName(FontName!.Trim());
        }

        if (!string.IsNullOrWhiteSpace(Color))
        {
            style = style.WithColor(NormalizeColor(Color));
        }

        if (Bold != null)
        {
            style = style.WithBold(Bold);
        }

        if (Italic != null)
        {
            style = style.WithItalic(Italic);
        }

        if (Underline != null)
        {
            style = style.WithUnderline(Underline);
        }

        if (!string.IsNullOrWhiteSpace(HighlightColor))
        {
            style = style.WithHighlightColor(NormalizeColor(HighlightColor));
        }

        bool hasStyleOverrides = FontSize != null
                                 || !string.IsNullOrWhiteSpace(FontName)
                                 || !string.IsNullOrWhiteSpace(Color)
                                 || Bold != null
                                 || Italic != null
                                 || Underline != null
                                 || !string.IsNullOrWhiteSpace(HighlightColor);

        bool hasBulletSettings = Level != null
                                 || !string.IsNullOrWhiteSpace(BulletChar)
                                 || Numbering.HasValue;

        return hasStyleName || hasStyleOverrides || hasBulletSettings;
    }

    private PowerPointNumberingScheme? ResolveNumbering()
    {
        return Numbering;
    }

    private char? ResolveBulletChar()
    {
        if (string.IsNullOrWhiteSpace(BulletChar))
        {
            return null;
        }

        return BulletChar!.Trim()[0];
    }

    private static string NormalizeColor(string? color)
    {
        if (string.IsNullOrWhiteSpace(color))
        {
            return string.Empty;
        }

        return OfficeColor.Parse(color!).ToRgbHex().ToLowerInvariant();
    }
}
