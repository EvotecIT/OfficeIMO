using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.PowerPoint.Html;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Creates native PowerPoint-to-HTML settings for Convert-OfficeDocument.</summary>
/// <example><summary>Publish a deck with the Modern document theme.</summary>
/// <code>$options = New-OfficePowerPointHtmlOptions -Theme Modern -NoNotes
/// Convert-OfficeDocument -Path ./Review.pptx -To Html -OutputPath ./Review.html -Options $options</code></example>
/// <example><summary>Export a visual review without inventories or extraction proof.</summary>
/// <code>$options = New-OfficePowerPointHtmlOptions -Profile VisualReview -NoMasterInventory -NoExtractionProof
/// Convert-OfficeDocument -Path ./Review.pptx -To Html -OutputPath ./Review.visual.html -Options $options</code></example>
[Cmdlet(VerbsCommon.New, "OfficePowerPointHtmlOptions", DefaultParameterSetName = "Default")]
[OutputType(typeof(PowerPointHtmlSaveOptions))]
public sealed class NewOfficePowerPointHtmlOptionsCommand : OfficeCmdlet {
    /// <summary>Semantic slide content or positioned visual review.</summary>
    [Parameter(ParameterSetName = "Default")]
    public PowerPointHtmlExportProfile Profile { get; set; } = PowerPointHtmlExportProfile.SemanticSlides;

    /// <summary>Shared HTML theme: a built-in name, OfficeTheme or Office document template path. Defaults to Modern.</summary>
    [Parameter, OfficeThemeTransformation]
    public OfficeThemeInput? Theme { get; set; }

    /// <summary>Optional HTML document title.</summary>
    [Parameter]
    public string? Title { get; set; }

    /// <summary>Include slides marked hidden.</summary>
    [Parameter]
    public SwitchParameter IncludeHiddenSlides { get; set; }

    /// <summary>Omit speaker notes.</summary>
    [Parameter]
    public SwitchParameter NoNotes { get; set; }

    /// <summary>Omit table content.</summary>
    [Parameter]
    public SwitchParameter NoTables { get; set; }

    /// <summary>Include shapes marked hidden.</summary>
    [Parameter]
    public SwitchParameter IncludeHiddenShapes { get; set; }

    /// <summary>Omit slide-aligned extraction proof.</summary>
    [Parameter]
    public SwitchParameter NoExtractionProof { get; set; }

    /// <summary>Omit the master, layout and theme inventory.</summary>
    [Parameter]
    public SwitchParameter NoMasterInventory { get; set; }

    /// <summary>Omit the visible document heading; the HTML title is retained.</summary>
    [Parameter]
    public SwitchParameter NoHeading { get; set; }

    /// <summary>Omit the heading before each slide.</summary>
    [Parameter]
    public SwitchParameter NoSlideHeadings { get; set; }

    /// <summary>Omit visible visual-review render diagnostics.</summary>
    [Parameter]
    public SwitchParameter NoRenderDiagnostics { get; set; }

    /// <summary>Omit the shared HTML stylesheet.</summary>
    [Parameter]
    public SwitchParameter NoDefaultStyles { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new PowerPointHtmlSaveOptions {
            ExportProfile = Profile,
            IncludeNotes = FeatureEnabled("Notes"),
            IncludeTables = FeatureEnabled("Tables"),
            IncludeExtractionProof = FeatureEnabled("ExtractionProof"),
            IncludeMasterInventory = FeatureEnabled("MasterInventory"),
            IncludeHeading = FeatureEnabled("Heading"),
            IncludeSlideHeadings = FeatureEnabled("SlideHeadings"),
            IncludeRenderDiagnostics = FeatureEnabled("RenderDiagnostics"),
            IncludeDefaultStyles = FeatureEnabled("DefaultStyles")
        };
        options.DocumentOutput.DocumentTheme = Theme?.Resolve(OfficeTheme.FromDocument) ?? OfficeTheme.Modern;
        if (Title != null) options.Title = Title;
        if (BoundSwitch(nameof(IncludeHiddenSlides)) is bool hiddenSlides) options.IncludeHiddenSlides = hiddenSlides;
        if (BoundSwitch(nameof(IncludeHiddenShapes)) is bool hiddenShapes) options.IncludeHiddenShapes = hiddenShapes;
        WriteObject(options);
    }
}
