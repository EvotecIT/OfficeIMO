using System;
using System.IO;
using System.Globalization;
using System.Management.Automation;
using System.Security;
using OfficeIMO;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Creates a themed presentation, returning a live document or saving a completed content block.</summary>
/// <para>Without a content block no file is written, including on disposal. With a block, publication is atomic and an existing file requires Force. A block failure leaves the destination unchanged.</para>
/// <example>
///   <summary>Create a deck and return its saved file.</summary>
///   <code>New-OfficePowerPoint -Path ./Status.pptx -Theme Modern -Author 'Operations team' -PassThru {
///     PowerPointSlide -Title 'Status update'
/// }</code>
/// </example>
/// <example>
///   <summary>Create a live document for explicit editing and saving.</summary>
///   <code>$deck = New-OfficePowerPoint -Theme Boardroom
/// PowerPointSlide -Document $deck -Title 'Service review'
/// Save-OfficePowerPoint -Document $deck -Path ./Review.pptx
/// Close-OfficePowerPoint -Document $deck</code>
/// </example>
/// <example>
///   <summary>Create a compact technical review using the dark theme.</summary>
///   <code>New-OfficePowerPoint -Path ./Architecture.pptx -Theme Dark -Purpose 'technical proposal' -LayoutStrategy Compact {
///     PowerPointSlide -Title 'Architecture review'
///     PowerPointSlide -Title 'Decisions' { PowerPointList 'Retain the API boundary', 'Track rollout evidence' }
/// }</code>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficePowerPoint", DefaultParameterSetName = "Live", SupportsShouldProcess = true)]
[OutputType(typeof(PowerPointDocument), typeof(FileInfo))]
public sealed class NewOfficePowerPointCommand : OfficeNewLifecycleCmdlet<PowerPointDocument> {
    /// <summary>Allow the new document to replace an existing destination.</summary>
    [Parameter]
    public SwitchParameter Force { get; set; }

    /// <summary>Built-in theme, theme object or a Word, Excel or PowerPoint document/template whose visual theme is read. Defaults to Modern.</summary>
    [Parameter]
    [OfficeThemeTransformation]
    public OfficeThemeInput? Theme { get; set; }

    /// <summary>Author shown on title slides and written to the package's creator property.</summary>
    [Parameter]
    public string? Author { get; set; }

    /// <summary>Date shown on title slides using the current culture's short date format.</summary>
    [Parameter]
    public DateTime? Date { get; set; }

    /// <summary>Brand accent used to derive the theme's accent palette while retaining its fonts and surfaces.</summary>
    [Parameter, OfficeColorTransformation]
    public OfficeColor Accent { get; set; }

    /// <summary>Composition tuning layered on the theme; Auto uses the theme and purpose-matched recipe.</summary>
    [Parameter]
    public PowerPointDeckStyle Style { get; set; }

    /// <summary>Plain-language purpose matched by the native design brief to a built-in recipe.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string? Purpose { get; set; }

    /// <summary>ContentFirst, DesignFirst, Compact or VisualFirst allocation; overrides the recipe's strategy.</summary>
    [Parameter]
    public PowerPointAutoLayoutStrategy LayoutStrategy { get; set; } = PowerPointAutoLayoutStrategy.ContentFirst;

    /// <summary>Stable deck seed; an individual slide can supply its own seed.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string? Seed { get; set; }

    /// <summary>Slide dimensions; copied templates retain their dimensions unless this parameter is supplied.</summary>
    [Parameter]
    public PowerPointSlideSizePreset SlideSize { get; set; } = PowerPointSlideSizePreset.Screen16x9;

    /// <summary>Suppress footer text on every designed slide.</summary>
    [Parameter]
    public SwitchParameter NoFooter { get; set; }

    /// <summary>Suppress slide numbers on every designed slide.</summary>
    [Parameter]
    public SwitchParameter NoSlideNumber { get; set; }

    /// <summary>Continuation title format, with title, part and total supplied as {0}, {1} and {2}.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string? ContinuationFormat { get; set; }

    /// <summary>Existing presentation/template copied as starting content. Theme applies visual styling separately.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string? Template { get; set; }

    /// <summary>Password for the saved file; plaintext input warns, and credentials contribute their password.</summary>
    [Parameter(ParameterSetName = "Content")]
    [OfficePasswordTransformation]
    public SecureString? Password { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (!Enum.IsDefined(typeof(PowerPointSlideSizePreset), SlideSize)) {
            InvalidArgument(new PSArgumentException("Choose a supported slide size.", nameof(SlideSize)), SlideSize);
            return;
        }
        if (!Enum.IsDefined(typeof(PowerPointDeckStyle), Style) || !Enum.IsDefined(typeof(PowerPointAutoLayoutStrategy), LayoutStrategy)) {
            InvalidArgument(new PSArgumentException("Choose a supported deck style and layout strategy."), Style);
            return;
        }
        if (Purpose != null && string.IsNullOrWhiteSpace(Purpose)) {
            InvalidArgument(new PSArgumentException("Deck purpose cannot be empty.", nameof(Purpose)), Purpose);
            return;
        }
        if (ContinuationFormat != null) {
            try { _ = string.Format(CultureInfo.CurrentCulture, ContinuationFormat, "Title", 2, 3); }
            catch (FormatException error) { InvalidArgument(error, ContinuationFormat); return; }
        }
        var options = new PowerPointCreateOptions {
            PersistenceMode = DocumentPersistenceMode.Explicit,
            Theme = Theme?.Resolve(OfficeTheme.FromDocument),
            Template = Template == null ? null : ResolveOutputPath(Template)
        };
        RunDocument(Path, isNew: true, force: Force.IsPresent,
            path => CreateDocument(path, options),
            (document, path, allowLoss) => {
                document.Save(path, PowerPointDocumentService.NativeSaveOptions(allowLoss, Password,
                    MyInvocation.BoundParameters.ContainsKey(nameof(Password))));
                return null;
            });
    }

    /// <summary>Transfers the configured native document to the lifecycle only after initialization succeeds.</summary>
    private PowerPointDocument CreateDocument(string? path, PowerPointCreateOptions options) {
        var document = path == null ? PowerPointDocument.Create(options) : PowerPointDocument.Create(path, options);
        try {
            if (MyInvocation.BoundParameters.ContainsKey(nameof(SlideSize))) document.SlideSize.SetPreset(SlideSize);
            if (MyInvocation.BoundParameters.ContainsKey(nameof(Accent))) {
                OfficeTheme theme = document.GetTheme();
                OfficeColor color = Accent.Resolve(theme);
                var brand = PowerPointDesignTheme.FromBrand(color.ToRgbHex());
                document.ApplyTheme(theme.WithPalette(theme.Palette.With(accent1: color,
                    accent2: OfficeColor.ParseHex(brand.Accent2Color), accent3: OfficeColor.ParseHex(brand.Accent3Color))));
            }
            if (Author != null) {
                document.BuiltinDocumentProperties.Creator = Author;
                if (!string.IsNullOrWhiteSpace(Author)) document.LayoutDeck.TitleMetadata.Add(Author);
            }
            if (Date.HasValue) document.LayoutDeck.TitleMetadata.Add(Date.Value.ToString("d", CultureInfo.CurrentCulture));
            document.LayoutDeck.Seed = Seed;
            if (Style != PowerPointDeckStyle.Auto || Purpose != null || MyInvocation.BoundParameters.ContainsKey(nameof(LayoutStrategy))) {
                document.LayoutDeck.Design = PowerPointDocumentService.CreateDesign(document, Style, Purpose,
                    MyInvocation.BoundParameters.ContainsKey(nameof(LayoutStrategy)) ? LayoutStrategy : null, Seed);
            }
            document.LayoutDeck.NoFooter = NoFooter;
            document.LayoutDeck.ShowSlideNumber = !NoSlideNumber;
            document.LayoutDeck.ContinuationFormat = ContinuationFormat;
            return document;
        } catch {
            document.Dispose();
            throw;
        }
    }

    /// <inheritdoc />
    protected override void InvokeContent(PowerPointDocument document) =>
        PowerPointDocumentService.InvokeContent(document, Content!);
}
