using System.IO;
using System.Net;
using System.Security;
using OfficeIMO;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Creates a live PDF document or atomically saves a completed composition block.</summary>
/// <para>Without Content no file is written and the native document keeps its optional associated path. A completed block saves and closes; Force is required to replace an existing file.</para>
/// <example><summary>Create a themed PDF report.</summary>
/// <code>New-OfficePdf -Path ./Report.pdf -Theme Modern -PassThru {
///     PdfHeading 'Service review'
///     PdfParagraph 'All services are available.'
/// }</code></example>
/// <example><summary>Create a live document and save it explicitly.</summary>
/// <code>$pdf = New-OfficePdf -Theme Classic -Size 11pt
/// Add-OfficePdfParagraph -Document $pdf -Text 'Reviewed and approved'
/// Save-OfficePdf -Document $pdf -Path ./Review.pdf</code></example>
[Cmdlet(VerbsCommon.New, "OfficePdf", SupportsShouldProcess = true)]
[OutputType(typeof(PdfDocument), typeof(FileInfo))]
public sealed class NewOfficePdfCommand : OfficeLifecycleCmdlet<PdfDocument> {
    /// <summary>Optional associated destination; required with a content block.</summary>
    [Parameter(Position = 0), ValidateNotNullOrEmpty]
    public string? Path { get; set; }

    /// <summary>Allow a completed content block to replace an existing destination.</summary>
    [Parameter]
    public SwitchParameter Force { get; set; }

    /// <summary>Built-in theme, theme object or document/template whose visual theme is read.</summary>
    [Parameter, OfficeThemeTransformation]
    public OfficeThemeInput? Theme { get; set; }

    /// <summary>Body font family; inherited heading defaults follow the document font.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string? Font { get; set; }

    /// <summary>Default body font size. A bare number means points.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength? Size { get; set; }

    /// <summary>Regular embedded font file; requires the Font family name.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string? FontPath { get; set; }

    /// <summary>Optional bold face used with FontPath.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string? BoldFontPath { get; set; }

    /// <summary>Optional italic face used with FontPath.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string? ItalicFontPath { get; set; }

    /// <summary>Optional bold italic face used with FontPath.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string? BoldItalicFontPath { get; set; }

    /// <summary>PDF file header version emitted by OfficeIMO.Pdf.</summary>
    [Parameter]
    public PdfFileVersion? FileVersion { get; set; }

    /// <summary>Create PDF outline/bookmark entries from heading elements.</summary>
    [Parameter]
    public SwitchParameter CreateOutlineFromHeadings { get; set; }

    /// <summary>Initial outline expansion level when heading outlines are created.</summary>
    [Parameter]
    public int? OutlineExpansionLevel { get; set; }

    /// <summary>Catalog page mode hint emitted for generated PDFs.</summary>
    [Parameter]
    public PdfCatalogPageMode? PageMode { get; set; }

    /// <summary>Catalog page layout hint emitted for generated PDFs.</summary>
    [Parameter]
    public PdfCatalogPageLayout? PageLayout { get; set; }

    /// <summary>Emit generated catalog page labels.</summary>
    [Parameter]
    public SwitchParameter IncludePageLabels { get; set; }

    /// <summary>Optional generated page-label prefix.</summary>
    [Parameter]
    public string? PageLabelPrefix { get; set; }

    /// <summary>Initial one-based page shown by PDF viewers that honor open actions.</summary>
    [Parameter]
    public int? OpenActionPage { get; set; }

    /// <summary>Open-action destination mode.</summary>
    [Parameter]
    public PdfOpenActionDestinationMode? OpenActionMode { get; set; }

    /// <summary>Optional open-action top coordinate in the page's native UserUnit space.</summary>
    [Parameter]
    public double? OpenActionTop { get; set; }

    /// <summary>Request PDF viewers to display the document title instead of the file name.</summary>
    [Parameter]
    public SwitchParameter DisplayDocTitle { get; set; }

    /// <summary>Request PDF viewers to fit the document window to the first displayed page.</summary>
    [Parameter]
    public SwitchParameter FitWindow { get; set; }

    /// <summary>Request PDF viewers to center the document window on screen.</summary>
    [Parameter]
    public SwitchParameter CenterWindow { get; set; }

    /// <summary>Request PDF viewers to hide the toolbar.</summary>
    [Parameter]
    public SwitchParameter HideToolbar { get; set; }

    /// <summary>Request PDF viewers to hide the menu bar.</summary>
    [Parameter]
    public SwitchParameter HideMenubar { get; set; }

    /// <summary>Request PDF viewers to hide user-interface elements.</summary>
    [Parameter]
    public SwitchParameter HideWindowUI { get; set; }

    /// <summary>Flatten generated FreeText and Highlight annotations into static page content.</summary>
    [Parameter]
    public SwitchParameter FlattenVisualAnnotations { get; set; }

    /// <summary>Password required to open the generated PDF.</summary>
    [Parameter, OfficePasswordTransformation]
    public SecureString? Password { get; set; }

    /// <summary>Optional owner password for the generated encrypted PDF.</summary>
    [Parameter, OfficePasswordTransformation]
    public SecureString? OwnerPassword { get; set; }

    /// <summary>Operations allowed to readers; defaults to all Standard security permissions.</summary>
    [Parameter]
    public PdfStandardPermissions Permission { get; set; } = PdfStandardPermissions.All;

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (Content == null && MyInvocation.BoundParameters.ContainsKey(nameof(AllowLoss))) {
            InvalidArgument(new PSArgumentException("AllowLoss requires a content block. Supply it to Save-OfficePdf when saving a live document."), Path);
            return;
        }
        PdfCreateOptions options = CreateOptions();
        RunDocument(Path, isNew: true, Force.IsPresent,
            path => path == null ? PdfDocument.Create(options) : PdfDocument.Create(path, options),
            (document, path, allowLoss) => { document.Save(path, new PdfSaveOptions { AllowLoss = allowLoss }); return null; });
    }

    /// <inheritdoc />
    protected override void InvokeContent(PdfDocument document) {
        using (var context = PdfDslContext.Enter(document)) {
            Content!.Invoke();
            context.Apply();
        }
    }

    private PdfCreateOptions CreateOptions() {
        var options = new PdfCreateOptions { Theme = Theme?.Resolve(OfficeTheme.FromDocument) ?? OfficeTheme.Default };
        if (Size.HasValue) options.DefaultFontSize = Size.Value;

        if (FileVersion.HasValue) {
            options.FileVersion = FileVersion.Value;
        }

        if (CreateOutlineFromHeadings.IsPresent) {
            options.CreateOutlineFromHeadings = true;
        }

        if (OutlineExpansionLevel.HasValue) {
            options.OutlineExpansionLevel = OutlineExpansionLevel.Value;
        }

        if (PageMode.HasValue || PageLayout.HasValue) {
            options.SetCatalogView(PageMode, PageLayout);
        }

        if (IncludePageLabels.IsPresent || !string.IsNullOrWhiteSpace(PageLabelPrefix)) {
            options.SetPageLabels(true, PageLabelPrefix);
        }

        if (OpenActionPage.HasValue || OpenActionMode.HasValue || OpenActionTop.HasValue) {
            options.SetOpenAction(
                OpenActionPage ?? 1,
                OpenActionTop,
                OpenActionMode ?? PdfOpenActionDestinationMode.Xyz);
        }

        if (FlattenVisualAnnotations.IsPresent) {
            options.SetFlattenVisualAnnotations();
        }

        ConfigureViewerPreferences(options);
        options.Encryption = CreateEncryption();

        if (FontPath != null) {
            if (string.IsNullOrWhiteSpace(Font)) throw new PSArgumentException("FontPath requires the Font family name.");
            options.UseFontFamily(Font!, PdfCommandUtilities.ResolveExistingFilePath(this, FontPath),
                ResolveOptionalFontPath(BoldFontPath), ResolveOptionalFontPath(ItalicFontPath), ResolveOptionalFontPath(BoldItalicFontPath));
        } else {
            if (BoldFontPath != null || ItalicFontPath != null || BoldItalicFontPath != null)
                throw new PSArgumentException("Additional font faces require FontPath.");
            if (Font != null) options.UseOfficeFontFamily(Font);
        }

        return options;
    }

    private PdfStandardEncryptionOptions? CreateEncryption() {
        string? password = Password == null ? null : new NetworkCredential(string.Empty, Password).Password;
        string? owner = OwnerPassword == null ? null : new NetworkCredential(string.Empty, OwnerPassword).Password;
        PdfStandardPermissions permissions = Permission;
        if ((permissions & ~PdfStandardPermissions.All) != 0) throw new PSArgumentException("Permission contains unsupported PDF flags.");
        if (password == null && owner == null) {
            if (MyInvocation.BoundParameters.ContainsKey(nameof(Permission)))
                throw new PSArgumentException("Permission requires Password or OwnerPassword.");
            return null;
        }
        return password == null
            ? PdfStandardEncryptionOptions.ForPermissions(owner!, permissions)
            : new PdfStandardEncryptionOptions(password) { OwnerPassword = owner, AllowedPermissions = permissions };
    }

    private void ConfigureViewerPreferences(PdfCreateOptions options) {
        if (!DisplayDocTitle.IsPresent &&
            !FitWindow.IsPresent &&
            !CenterWindow.IsPresent &&
            !HideToolbar.IsPresent &&
            !HideMenubar.IsPresent &&
            !HideWindowUI.IsPresent) {
            return;
        }

        options.ConfigureViewerPreferences(preferences => {
            if (DisplayDocTitle.IsPresent) {
                preferences.DisplayDocTitle = true;
            }

            if (FitWindow.IsPresent) {
                preferences.FitWindow = true;
            }

            if (CenterWindow.IsPresent) {
                preferences.CenterWindow = true;
            }

            if (HideToolbar.IsPresent) {
                preferences.HideToolbar = true;
            }

            if (HideMenubar.IsPresent) {
                preferences.HideMenubar = true;
            }

            if (HideWindowUI.IsPresent) {
                preferences.HideWindowUI = true;
            }
        });
    }

    private string? ResolveOptionalFontPath(string? path) {
        return string.IsNullOrWhiteSpace(path)
            ? null
            : PdfCommandUtilities.ResolveExistingFilePath(this, path!);
    }
}
