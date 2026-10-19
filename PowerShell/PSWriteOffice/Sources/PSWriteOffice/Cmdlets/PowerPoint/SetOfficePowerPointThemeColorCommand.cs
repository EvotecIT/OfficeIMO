using System;
using System.Collections;
using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services;
using PSWriteOffice.Services.PowerPoint;
using PSWriteOffice.Services.Text;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Sets one or more PowerPoint theme colors.</summary>
/// <example>
///   <summary>Set a single accent color.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Set-OfficePowerPointThemeColor -Presentation $ppt -Color Accent1 -Value '#C00000'</code>
///   <para>Updates Accent1 on the default master.</para>
/// </example>
/// <example>
///   <summary>Set multiple theme colors across all masters.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Set-OfficePowerPointThemeColor -Presentation $ppt -Colors @{ Accent1 = '#C00000'; Accent2 = '#00B0F0' } -AllMasters</code>
///   <para>Applies multiple theme colors to every master in the presentation.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficePowerPointThemeColor", DefaultParameterSetName = ParameterSetSingle)]
[Alias("PptThemeColor")]
[OutputType(typeof(PowerPointPresentation))]
public sealed class SetOfficePowerPointThemeColorCommand : PSCmdlet
{
    private const string ParameterSetSingle = "Single";
    private const string ParameterSetMultiple = "Multiple";

    /// <summary>Presentation to update (optional inside New-OfficePowerPoint).</summary>
    [Parameter(ValueFromPipeline = true)]
    public PowerPointPresentation? Presentation { get; set; }

    /// <summary>Theme color to update.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetSingle)]
    public PowerPointThemeColor Color { get; set; }

    /// <summary>Named or hexadecimal color value.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetSingle)]
    [OfficeColorArgumentTransformation]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string Value { get; set; } = string.Empty;

    /// <summary>Hashtable of theme color names to named or hexadecimal color values.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetMultiple)]
    public Hashtable Colors { get; set; } = null!;

    /// <summary>Slide master index to update when not using <see cref="AllMasters"/>.</summary>
    [Parameter]
    [ValidateRange(0, int.MaxValue)]
    public int Master { get; set; }

    /// <summary>Apply the changes across all slide masters.</summary>
    [Parameter]
    public SwitchParameter AllMasters { get; set; }

    /// <summary>Emit the presentation after update.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        try
        {
            var presentation = Presentation ?? PowerPointDslContext.Require(this).Presentation;

            if (ParameterSetName == ParameterSetSingle)
            {
                var normalized = NormalizeColor(Value);
                if (AllMasters.IsPresent)
                {
                    presentation.SetThemeColorForAllMasters(Color, normalized);
                }
                else
                {
                    presentation.SetThemeColor(Color, normalized, Master);
                }
            }
            else
            {
                var resolvedColors = ResolveColors();
                if (AllMasters.IsPresent)
                {
                    presentation.SetThemeColorsForAllMasters(resolvedColors);
                }
                else
                {
                    presentation.SetThemeColors(resolvedColors, Master);
                }
            }

            if (PassThru.IsPresent)
            {
                WriteObject(presentation);
            }
        }
        catch (Exception ex)
        {
            WriteError(new ErrorRecord(ex, "PowerPointSetThemeColorFailed", ErrorCategory.InvalidOperation, Presentation));
        }
    }

    private Dictionary<PowerPointThemeColor, string> ResolveColors()
    {
        var resolved = new Dictionary<PowerPointThemeColor, string>();
        foreach (DictionaryEntry entry in Colors)
        {
            var key = LanguagePrimitives.ConvertTo<string>(entry.Key);
            if (!OpenXmlValueParser.TryParse<PowerPointThemeColor>(key, out var color))
            {
                throw new PSArgumentException($"Unknown theme color '{key}'.", nameof(Colors));
            }

            var value = LanguagePrimitives.ConvertTo<string>(entry.Value);
            resolved[color] = NormalizeColor(value);
        }

        return resolved;
    }

    private static string NormalizeColor(string color)
    {
        if (string.IsNullOrWhiteSpace(color))
        {
            throw new PSArgumentException("Provide a non-empty color value.");
        }

        return OfficeColorUtilities.ToRgbHex(color, nameof(Value))!.ToUpperInvariant();
    }
}
