using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Creates native PDF rewrite preservation checks for Test-OfficePdfRewrite.</summary>
/// <example><summary>Permit a rewrite to update document metadata.</summary>
/// <code>$options = New-OfficePdfRewriteOptions -NoMetadata -NoXmpMetadata
/// Test-OfficePdfRewrite -ReferencePath ./Original.pdf -DifferencePath ./Updated.pdf -Options $options -FailOnLoss</code></example>
[Cmdlet(VerbsCommon.New, "OfficePdfRewriteOptions", DefaultParameterSetName = "Default")]
[OutputType(typeof(PdfRewritePreservationOptions))]
public sealed class NewOfficePdfRewriteOptionsCommand : OfficeCmdlet {
    /// <summary>Exclude page count from preservation checks.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter NoPageCount { get; set; }
    /// <summary>Exclude page geometry from preservation checks.</summary>
    [Parameter]
    public SwitchParameter NoPageGeometry { get; set; }
    /// <summary>Exclude metadata from preservation checks.</summary>
    [Parameter]
    public SwitchParameter NoMetadata { get; set; }
    /// <summary>Exclude outlines from preservation checks.</summary>
    [Parameter]
    public SwitchParameter NoOutlines { get; set; }
    /// <summary>Exclude named destinations from preservation checks.</summary>
    [Parameter]
    public SwitchParameter NoNamedDestinations { get; set; }
    /// <summary>Exclude page labels from preservation checks.</summary>
    [Parameter]
    public SwitchParameter NoPageLabels { get; set; }
    /// <summary>Exclude link annotations from preservation checks.</summary>
    [Parameter]
    public SwitchParameter NoLinkAnnotations { get; set; }
    /// <summary>Exclude annotations from preservation checks.</summary>
    [Parameter]
    public SwitchParameter NoAnnotations { get; set; }
    /// <summary>Exclude forms from preservation checks.</summary>
    [Parameter]
    public SwitchParameter NoForms { get; set; }
    /// <summary>Exclude embedded files from preservation checks.</summary>
    [Parameter]
    public SwitchParameter NoEmbeddedFiles { get; set; }
    /// <summary>Exclude xmp metadata from preservation checks.</summary>
    [Parameter]
    public SwitchParameter NoXmpMetadata { get; set; }
    /// <summary>Exclude output intents from preservation checks.</summary>
    [Parameter]
    public SwitchParameter NoOutputIntents { get; set; }
    /// <summary>Exclude optional content from preservation checks.</summary>
    [Parameter]
    public SwitchParameter NoOptionalContent { get; set; }
    /// <summary>Exclude tagged content from preservation checks.</summary>
    [Parameter]
    public SwitchParameter NoTaggedContent { get; set; }
    /// <summary>Exclude catalog view settings from preservation checks.</summary>
    [Parameter]
    public SwitchParameter NoCatalogViewSettings { get; set; }
    /// <summary>Exclude open action from preservation checks.</summary>
    [Parameter]
    public SwitchParameter NoOpenAction { get; set; }
    /// <summary>Exclude catalog actions from preservation checks.</summary>
    [Parameter]
    public SwitchParameter NoCatalogActions { get; set; }
    /// <summary>Exclude page actions from preservation checks.</summary>
    [Parameter]
    public SwitchParameter NoPageActions { get; set; }
    /// <summary>Exclude viewer preferences from preservation checks.</summary>
    [Parameter]
    public SwitchParameter NoViewerPreferences { get; set; }
    /// <summary>Exclude document version state from preservation checks.</summary>
    [Parameter]
    public SwitchParameter NoDocumentVersionState { get; set; }
    /// <summary>Exclude revision structure from preservation checks.</summary>
    [Parameter]
    public SwitchParameter NoRevisionStructure { get; set; }
    /// <summary>Exclude security state from preservation checks.</summary>
    [Parameter]
    public SwitchParameter NoSecurityState { get; set; }
    /// <summary>Require preservation of form field and widget actions.</summary>
    [Parameter]
    public SwitchParameter FormWidgetActions { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() => WriteObject(new PdfRewritePreservationOptions {
        PreservePageCount = !NoPageCount.IsPresent,
        PreservePageGeometry = !NoPageGeometry.IsPresent,
        PreserveMetadata = !NoMetadata.IsPresent,
        PreserveOutlines = !NoOutlines.IsPresent,
        PreserveNamedDestinations = !NoNamedDestinations.IsPresent,
        PreservePageLabels = !NoPageLabels.IsPresent,
        PreserveLinkAnnotations = !NoLinkAnnotations.IsPresent,
        PreserveAnnotations = !NoAnnotations.IsPresent,
        PreserveForms = !NoForms.IsPresent,
        PreserveEmbeddedFiles = !NoEmbeddedFiles.IsPresent,
        PreserveXmpMetadata = !NoXmpMetadata.IsPresent,
        PreserveOutputIntents = !NoOutputIntents.IsPresent,
        PreserveOptionalContent = !NoOptionalContent.IsPresent,
        PreserveTaggedContent = !NoTaggedContent.IsPresent,
        PreserveCatalogViewSettings = !NoCatalogViewSettings.IsPresent,
        PreserveOpenAction = !NoOpenAction.IsPresent,
        PreserveCatalogActions = !NoCatalogActions.IsPresent,
        PreservePageActions = !NoPageActions.IsPresent,
        PreserveViewerPreferences = !NoViewerPreferences.IsPresent,
        PreserveDocumentVersionState = !NoDocumentVersionState.IsPresent,
        PreserveRevisionStructure = !NoRevisionStructure.IsPresent,
        PreserveSecurityState = !NoSecurityState.IsPresent,
        PreserveFormWidgetActions = FormWidgetActions.IsPresent,
    });
}
