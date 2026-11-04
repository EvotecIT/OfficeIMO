using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Text;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Restyles a deck or changes the selected masters' colors, fonts and theme name.</summary>
/// <example><summary>Restyle every master and layout with a built-in theme.</summary>
/// <code>Set-OfficePowerPointTheme -Document $deck -Theme Boardroom -AllMasters</code></example>
/// <example><summary>Fine-tune the first master's theme without restyling its layouts.</summary>
/// <code>Set-OfficePowerPointTheme -Document $deck -Colors @{ Accent1 = '#0F766E' } -HeadingFont 'Segoe UI Semibold' -BodyFont 'Segoe UI' -Name 'Service Brief'</code></example>
[Cmdlet(VerbsCommon.Set, "OfficePowerPointTheme", DefaultParameterSetName = "Document", SupportsShouldProcess = true)]
[OutputType(typeof(PowerPointDocument))]
public sealed class SetOfficePowerPointThemeCommand : OfficeDocumentCmdlet<PowerPointDocument> {
    /// <summary>Gets or sets a built-in theme, OfficeTheme object or document/template path. Applying it restyles masters and layouts before overrides.</summary>
    [Parameter(ParameterSetName = "Document"), OfficeThemeTransformation, ValidateNotNull]
    public OfficeThemeInput? Theme { get; set; }
    /// <summary>Gets or sets theme slots mapped to OfficeColor values, CSS colors or semantic roles.</summary>
    [Parameter, ValidateNotNull]
    public Hashtable? Colors { get; set; }
    /// <summary>Gets or sets the Latin heading font.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string? HeadingFont { get; set; }
    /// <summary>Gets or sets the Latin body font.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string? BodyFont { get; set; }
    /// <summary>Gets or sets the East Asian heading font.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string? HeadingFontEastAsian { get; set; }
    /// <summary>Gets or sets the East Asian body font.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string? BodyFontEastAsian { get; set; }
    /// <summary>Gets or sets the complex script heading font.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string? HeadingFontComplexScript { get; set; }
    /// <summary>Gets or sets the complex script body font.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string? BodyFontComplexScript { get; set; }
    /// <summary>Gets or sets whether unspecified font slots are cleared. Requires at least one font override.</summary>
    [Parameter]
    public SwitchParameter ClearMissing { get; set; }
    /// <summary>Gets or sets the theme name.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string? Name { get; set; }
    /// <summary>Gets or sets the one-based master position; defaults to the first.</summary>
    [Parameter, ValidateRange(1, int.MaxValue)]
    public int Master { get; set; } = 1;
    /// <summary>Gets or sets whether every master is updated.</summary>
    [Parameter]
    public SwitchParameter AllMasters { get; set; }
    /// <summary>Gets or sets whether the updated native document is returned.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        PowerPointDocument document = GetDocument();
        try {
            var fonts = new[] { HeadingFont, BodyFont, HeadingFontEastAsian, BodyFontEastAsian,
                HeadingFontComplexScript, BodyFontComplexScript };
            bool hasFonts = fonts.Any(value => value != null);
            if (fonts.Any(value => value != null && string.IsNullOrWhiteSpace(value)) ||
                Name != null && string.IsNullOrWhiteSpace(Name))
                throw new ArgumentException("Font and theme names cannot be blank.");
            if (ClearMissing && !hasFonts) throw new ArgumentException("ClearMissing requires a font override.");
            if (Theme == null && Colors?.Count is not > 0 && !hasFonts && Name == null)
                throw new ArgumentException("Supply Theme, Colors, a font override or Name.");

            OfficeTheme? theme = Theme?.Resolve(OfficeTheme.FromDocument);
            int[] masters = AllMasters
                ? PowerPointTemplate.Inspect(document).Masters.Select(master => master.MasterIndex).ToArray()
                : new[] { Master - 1 };
            // Resolve every override before changing any master, including under WhatIf.
            var overrides = new Dictionary<int, Dictionary<PowerPointThemeColor, string>>();
            foreach (int master in masters) {
                OfficeTheme palette = theme ?? document.GetTheme(master);
                if (theme != null) document.GetThemeColors(master); // Validate the selected master before restyling.
                var colors = new Dictionary<PowerPointThemeColor, string>();
                if (Colors != null) foreach (DictionaryEntry entry in Colors) {
                    string key = LanguagePrimitives.ConvertTo<string>(entry.Key);
                    if (!OpenXmlValueParser.TryParse<PowerPointThemeColor>(key, out var slot) ||
                        !Enum.IsDefined(typeof(PowerPointThemeColor), slot))
                        throw new ArgumentException($"Unknown theme color slot '{key}'.", nameof(Colors));
                    OfficeColor color = OfficeColorTransformationAttribute.Bind(entry.Value).Resolve(palette);
                    if (color.A != byte.MaxValue) throw new ArgumentException("Theme colors must be opaque.", nameof(Colors));
                    colors[slot] = color.ToRgbHex().ToUpperInvariant();
                }
                overrides.Add(master, colors);
            }
            if (!ShouldProcess(AllMasters ? "All presentation masters" : $"Presentation master {Master}", "Update PowerPoint theme")) return;
            if (theme != null) document.ApplyTheme(theme, AllMasters ? null : Master - 1);
            foreach (int master in masters) {
                document.SetThemeColors(overrides[master], master);
                if (hasFonts) document.SetThemeFonts(new PowerPointThemeFontSet(HeadingFont, BodyFont,
                    HeadingFontEastAsian, BodyFontEastAsian, HeadingFontComplexScript, BodyFontComplexScript),
                    master, keepExistingWhenNull: !ClearMissing);
                if (Name != null) document.SetThemeName(Name, master);
            }
            if (PassThru) WriteObject(document);
        } catch (ArgumentException error) { InvalidArgument(error, document); }
        catch (Exception error) when (error is not (PipelineStoppedException or ActionPreferenceStopException)) {
            ItemError(error, "powerpoint-theme-update-failed", document);
        }
    }
}
