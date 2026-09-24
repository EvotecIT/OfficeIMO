using OfficeIMO.Drawing;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

internal static partial class HtmlPdfRenderedConverter {
    private const double PointsPerCssPixel = 72D / HtmlRenderOptions.CssPixelsPerInch;
    private const int MaximumSystemFontFamilyCandidates = 512;
    private const int MaximumLoadedSystemFontFamilies = 32;
    private static readonly ConditionalWeakTable<byte[], CachedPdfImageResources> PdfImageResources = new();

    internal static HtmlPdfRenderResult Convert(HtmlConversionDocument document, HtmlToPdfOptions options, CancellationToken cancellationToken = default) {
        HtmlRenderRequest request = HtmlRenderRequest.FromLegacy(
            options, HtmlRenderEncoder.Pdf, HtmlRenderPageSet.All(), forcePrintPaged: true);
        return Convert(document, request, options, cancellationToken);
    }

    internal static async Task<HtmlPdfRenderResult> ConvertAsync(HtmlConversionDocument document, HtmlToPdfOptions options, CancellationToken cancellationToken) {
        HtmlRenderRequest request = HtmlRenderRequest.FromLegacy(
            options, HtmlRenderEncoder.Pdf, HtmlRenderPageSet.All(), forcePrintPaged: true);
        return await ConvertAsync(document, request, options, cancellationToken).ConfigureAwait(false);
    }

    internal static byte[] ConvertToBytes(
        HtmlConversionDocument document,
        HtmlToPdfOptions options,
        CancellationToken cancellationToken = default) {
        HtmlRenderRequest request = HtmlRenderRequest.FromLegacy(
            options, HtmlRenderEncoder.Pdf, HtmlRenderPageSet.All(), forcePrintPaged: true);
        return ConvertToBytes(document, request, options, cancellationToken);
    }

    internal static async Task<byte[]> ConvertToBytesAsync(
        HtmlConversionDocument document,
        HtmlToPdfOptions options,
        CancellationToken cancellationToken) {
        HtmlRenderRequest request = HtmlRenderRequest.FromLegacy(
            options, HtmlRenderEncoder.Pdf, HtmlRenderPageSet.All(), forcePrintPaged: true);
        return await ConvertToBytesAsync(document, request, options, cancellationToken).ConfigureAwait(false);
    }

    internal static HtmlPdfRenderResult Convert(
        HtmlConversionDocument document,
        HtmlRenderRequest request,
        HtmlToPdfOptions options,
        CancellationToken cancellationToken = default) {
        PreparedPdfRequest prepared = PrepareRequest(request, options);
        HtmlRenderOptions resolved = HtmlRenderEngine.PrepareOptions(document, prepared.Request);
        return HtmlRenderEngine.ExecuteWithDeadline(resolved, cancellationToken,
            operationCancellationToken => ConvertCore(
                document, prepared, resolved, operationCancellationToken));
    }

    internal static async Task<HtmlPdfRenderResult> ConvertAsync(
        HtmlConversionDocument document,
        HtmlRenderRequest request,
        HtmlToPdfOptions options,
        CancellationToken cancellationToken) {
        PreparedPdfRequest prepared = PrepareRequest(request, options);
        HtmlRenderOptions resolved = HtmlRenderEngine.PrepareOptions(document, prepared.Request);
        return await HtmlRenderEngine.ExecuteWithDeadlineAsync(resolved, cancellationToken,
            operationCancellationToken => ConvertCoreAsync(
                document, prepared, resolved, operationCancellationToken)).ConfigureAwait(false);
    }

    internal static byte[] ConvertToBytes(
        HtmlConversionDocument document,
        HtmlRenderRequest request,
        HtmlToPdfOptions options,
        CancellationToken cancellationToken = default) {
        PreparedPdfRequest prepared = PrepareRequest(request, options);
        HtmlRenderOptions resolved = HtmlRenderEngine.PrepareOptions(document, prepared.Request);
        return HtmlRenderEngine.ExecuteWithDeadline(resolved, cancellationToken, operationCancellationToken =>
            ConvertCore(document, prepared, resolved, operationCancellationToken)
                .Document.ToBytes(operationCancellationToken));
    }

    internal static async Task<byte[]> ConvertToBytesAsync(
        HtmlConversionDocument document,
        HtmlRenderRequest request,
        HtmlToPdfOptions options,
        CancellationToken cancellationToken) {
        PreparedPdfRequest prepared = PrepareRequest(request, options);
        HtmlRenderOptions resolved = HtmlRenderEngine.PrepareOptions(document, prepared.Request);
        return await HtmlRenderEngine.ExecuteWithDeadlineAsync(resolved, cancellationToken, async operationCancellationToken => {
            HtmlPdfRenderResult rendered = await ConvertCoreAsync(
                document, prepared, resolved, operationCancellationToken).ConfigureAwait(false);
            return rendered.Document.ToBytes(operationCancellationToken);
        }).ConfigureAwait(false);
    }

    internal static HtmlRenderOptions ResolveRenderOptions(HtmlToPdfOptions options) {
        HtmlRenderRequest request = HtmlRenderRequest.FromLegacy(
            options, HtmlRenderEncoder.Pdf, HtmlRenderPageSet.All(), forcePrintPaged: true);
        return PrepareRequest(request, options).Request.ResolveOptions();
    }

    private static PreparedPdfRequest PrepareRequest(HtmlRenderRequest request, HtmlToPdfOptions options) {
        if (request == null) throw new ArgumentNullException(nameof(request));
        if (options == null) throw new ArgumentNullException(nameof(options));
        if (request.Encoder != HtmlRenderEncoder.Pdf) {
            throw new ArgumentException("HTML-to-PDF conversion requires the Pdf render encoder.", nameof(request));
        }
        HtmlToPdfOptions renderOptions = new HtmlToPdfOptions(request.Options);
        CopyAdapterOptions(options, renderOptions);
        ApplyPrintLayoutWidth(request, renderOptions);
        PdfCore.PdfOptions measurementOptions = options.PdfOptions.Clone();
        measurementOptions.SetTextShapingMode(options.TextShapingMode).SetTextShapingProvider(options.TextShapingProvider);
        if (options.FontFamily != null) measurementOptions.RegisterFontFamily(PdfCore.PdfStandardFont.Helvetica, options.FontFamily);
        renderOptions.FallbackTextMeasurement = (text, font) => PdfCore.PdfWriter.MeasurePositionedText(
            new PdfCore.PdfTextRun(text, bold: font.IsBold, italic: font.IsItalic,
                fontSize: font.Size, font: MapStandardFont(font.FamilyName), fontFamily: font.FamilyName), measurementOptions);
        HtmlRenderResourceResolver? embeddedPackageResolver = options.EmbeddedPackageResourceResolver;
        HtmlUrlPolicy hostResourceUrlPolicy = (options.EmbeddedPackageHostResourceUrlPolicy ?? renderOptions.GetResourceUrlPolicy()).Clone();
        ApplyResourceAccessPolicy(
            hostResourceUrlPolicy,
            allowDataUrls: options.ResourcePolicy.AllowDataUris,
            allowFileUrls: options.ResourcePolicy.AllowLocalFileAccess);
        renderOptions.ResourceUrlPolicy = renderOptions.GetResourceUrlPolicy().Clone();
        ApplyResourceAccessPolicy(
            renderOptions.ResourceUrlPolicy,
            allowDataUrls: options.ResourcePolicy.AllowDataUris,
            allowFileUrls: options.ResourcePolicy.AllowLocalFileAccess ||
                embeddedPackageResolver != null && options.ResourcePolicy.AllowEmbeddedPackageResources);
        HtmlRenderResourceResolver? hostResolver = renderOptions.ResourceResolver;
        if (embeddedPackageResolver != null || hostResolver != null) {
            renderOptions.ResourceResolver = async (request, cancellationToken) => {
                if (embeddedPackageResolver != null && options.ResourcePolicy.AllowEmbeddedPackageResources) {
                    HtmlResolvedResource? embedded = await embeddedPackageResolver(request, cancellationToken).ConfigureAwait(false);
                    if (embedded != null) return embedded;
                }

                if (hostResolver == null) return null;
                bool hostResourceAllowed = request.Uri.IsFile
                    ? options.ResourcePolicy.AllowLocalFileAccess
                    : options.ResourcePolicy.AllowRemoteResourceResolution;
                hostResourceAllowed = hostResourceAllowed
                    && HtmlUrlPolicyEvaluator.IsAllowed(request.Uri.AbsoluteUri, hostResourceUrlPolicy);
                return hostResourceAllowed
                    ? await hostResolver(request, cancellationToken).ConfigureAwait(false)
                    : null;
            };
        }
        return new PreparedPdfRequest(request.WithOptions(renderOptions), renderOptions);
    }

    private static HtmlPdfRenderResult ConvertCore(
        HtmlConversionDocument document,
        PreparedPdfRequest prepared,
        HtmlRenderOptions resolved,
        CancellationToken cancellationToken) {
        HtmlRenderResult rendered = HtmlRenderEngine.ExecuteCore(
            document, prepared.Request, resolved, cancellationToken);
        return CreatePdf(rendered.Document, prepared.Options, cancellationToken).WithRenderResult(rendered);
    }

    private static async Task<HtmlPdfRenderResult> ConvertCoreAsync(
        HtmlConversionDocument document,
        PreparedPdfRequest prepared,
        HtmlRenderOptions resolved,
        CancellationToken cancellationToken) {
        HtmlRenderResult rendered = await HtmlRenderEngine.ExecuteCoreAsync(
            document, prepared.Request, resolved, cancellationToken).ConfigureAwait(false);
        return CreatePdf(rendered.Document, prepared.Options, cancellationToken).WithRenderResult(rendered);
    }

    private static void CopyAdapterOptions(HtmlToPdfOptions source, HtmlToPdfOptions target) {
        target.TextFallbacks = source.TextFallbacks;
        target.TextShapingMode = source.TextShapingMode;
        target.FontFamily = source.FontFamily;
        target.InteractiveFormControls = source.InteractiveFormControls;
        target.PrintLayoutWidthCssPixels = source.PrintLayoutWidthCssPixels;
        target.MaxOutlinedTextCharactersPerRun = source.MaxOutlinedTextCharactersPerRun;
        target.MaxOutlinedTextPathCommands = source.MaxOutlinedTextPathCommands;
        target.PdfOptions = source.PdfOptions.Clone();
        target.TextShapingProvider = source.TextShapingProvider;
        target.ResourcePolicy = source.ResourcePolicy.Clone();
        target.EmbeddedPackageResourceResolver = source.EmbeddedPackageResourceResolver;
        target.EmbeddedPackageHostResourceUrlPolicy = source.EmbeddedPackageHostResourceUrlPolicy?.Clone();
    }

    private readonly struct PreparedPdfRequest {
        internal PreparedPdfRequest(HtmlRenderRequest request, HtmlToPdfOptions options) {
            Request = request;
            Options = options;
        }
        internal HtmlRenderRequest Request { get; }
        internal HtmlToPdfOptions Options { get; }
    }

    private static void ApplyResourceAccessPolicy(HtmlUrlPolicy policy, bool allowDataUrls, bool allowFileUrls) {
        policy.AllowDataUrls = allowDataUrls;
        policy.DisallowFileUrls = !allowFileUrls;
        SetAllowedScheme(policy, "data", allowDataUrls);
        SetAllowedScheme(policy, Uri.UriSchemeFile, allowFileUrls);
    }

    private static void SetAllowedScheme(HtmlUrlPolicy policy, string scheme, bool allowed) {
        if (allowed) {
            policy.AllowedUrlSchemes.Add(scheme);
        } else {
            policy.AllowedUrlSchemes.Remove(scheme);
        }
    }

    internal static HtmlPdfRenderResult CreatePdf(HtmlRenderDocument rendered, HtmlToPdfOptions options, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        HtmlDiagnosticReport diagnostics = rendered.DiagnosticReport.Clone();

        var conversionReport = new PdfCore.PdfConversionReport();
        PdfCore.PdfOptions documentOptions = options.PdfOptions.Clone();
        documentOptions.UseContentStreamCompressionByDefault();
        PdfCore.PdfDocument pdf = PdfCore.PdfDocument.Create(documentOptions);
        pdf.Options.ReportDiagnosticsTo(conversionReport, "OfficeIMO.Html.Pdf");
        if (rendered.Metadata.Title != null
            || rendered.Metadata.Author != null
            || rendered.Metadata.Subject != null
            || rendered.Metadata.Keywords != null) {
            pdf.Meta(
                title: rendered.Metadata.Title,
                author: rendered.Metadata.Author,
                subject: rendered.Metadata.Subject,
                keywords: rendered.Metadata.Keywords);
        }
        if ((string.IsNullOrWhiteSpace(documentOptions.Language)
             || string.Equals(documentOptions.Language, "und", StringComparison.OrdinalIgnoreCase))
            && rendered.Metadata.Language != null) {
            pdf.Language(rendered.Metadata.Language);
        }
        if (rendered.Metadata.Title != null || rendered.Metadata.Direction == HtmlRenderTextDirection.RightToLeft) {
            pdf.ViewerPreferences(preferences => {
                if (rendered.Metadata.Title != null) preferences.DisplayDocTitle = true;
                if (rendered.Metadata.Direction == HtmlRenderTextDirection.RightToLeft) preferences.Direction = PdfCore.PdfViewerDirection.RightToLeft;
            });
        }
        if (options.FontFamily != null) {
            pdf.UseFontFamily(options.FontFamily);
        }

        var reservedFontSlots = new HashSet<PdfCore.PdfStandardFont>();
        if (options.FontFamily != null) reservedFontSlots.Add(PdfCore.PdfStandardFont.Helvetica);
        RegisteredWebFonts webFonts = RegisterWebFonts(
            pdf,
            rendered,
            diagnostics,
            options.MaxOutlinedTextCharactersPerRun,
            options.MaxOutlinedTextPathCommands,
            options.TextShapingProvider,
            options.TextShapingLanguage,
            cancellationToken);
        var activeWebFontFamilies = new HashSet<string>(
            webFonts.Slots.Keys,
            StringComparer.OrdinalIgnoreCase);
        PdfCore.PdfTextFallbackFeatures activeTextFallbacks = ResolveTextFallbackFeatures(rendered, options.TextFallbacks);
        if (options.ResourcePolicy.AllowSystemFontEmbedding) {
            if (options.ResourcePolicy.AllowDocumentFontEmbedding) {
                RegisterUsedSystemFontFamilies(pdf, rendered, activeWebFontFamilies, reservedFontSlots, cancellationToken);
            } else if (activeTextFallbacks != PdfCore.PdfTextFallbackFeatures.None && options.FontFamily == null) {
                RegisterLibrarySelectedDefaultSystemFontFamily(
                    pdf,
                    rendered,
                    activeWebFontFamilies,
                    reservedFontSlots,
                    cancellationToken);
            }
        }
        ReserveUsedStandardFontSlots(rendered, activeWebFontFamilies, reservedFontSlots);
        foreach (PdfCore.PdfStandardFont slot in webFonts.Slots.Values) {
            reservedFontSlots.Add(PdfCore.PdfStandardFontMapper.GetFontFamily(slot));
        }
        if (activeTextFallbacks != PdfCore.PdfTextFallbackFeatures.None) {
            pdf.Options.UseTextFallbacks(
                activeTextFallbacks,
                reservedFontSlots,
                options.ResourcePolicy.AllowSystemFontEmbedding,
                preserveConfiguredFontSlots: options.FontFamily != null,
                requiredText: CollectRenderedTextForFontFallbackSelection(rendered));
        }
        pdf.UseTextShaping(options.TextShapingMode, options.TextShapingProvider);
        var headingDocumentOrder = rendered.Headings
            .Select((heading, index) => new { Heading = heading, Index = index })
            .ToDictionary(item => item.Heading, item => item.Index);
        ILookup<int, HtmlRenderHeading> headingsByPage = rendered.Headings.ToLookup(heading => heading.PageNumber);
        double printLayoutScale = ResolvePrintLayoutScale(options);
        foreach (HtmlRenderPage renderedPage in rendered.Pages) {
            cancellationToken.ThrowIfCancellationRequested();
            double pageWidth = renderedPage.Width * PointsPerCssPixel * printLayoutScale;
            double pageHeight = renderedPage.Height * PointsPerCssPixel * printLayoutScale;
            pdf.Page(page => {
                page.Size(pageWidth, pageHeight)
                    .Margin(0D);
                if (renderedPage.PrintProduction != null) {
                    double trimInset = renderedPage.PrintProduction.TrimInset * PointsPerCssPixel * printLayoutScale;
                    double bleedInset = renderedPage.PrintProduction.BleedInset * PointsPerCssPixel * printLayoutScale;
                    page.PrintProductionPageBoxes(new PdfCore.PdfPrintProductionPageBoxes(
                        PdfCore.PageMargins.Uniform(trimInset),
                        PdfCore.PageMargins.Uniform(bleedInset)));
                }
                page.Canvas(canvas => {
                    if (printLayoutScale == 1D) {
                        AddPageVisuals(canvas, renderedPage, webFonts, conversionReport, options.InteractiveFormControls, cancellationToken);
                    } else {
                        canvas.Effect(OfficeTransform.Scale(printLayoutScale, printLayoutScale), 1D,
                            content => AddPageVisuals(content, renderedPage, webFonts, conversionReport, options.InteractiveFormControls, cancellationToken));
                    }
                    AddPageOutlines(canvas, headingsByPage[renderedPage.PageNumber], headingDocumentOrder, printLayoutScale, cancellationToken);
                });
            });
        }

        if (options.FidelityPolicy == HtmlRenderFidelityPolicy.RequireNoLoss
            && diagnostics.Any(static diagnostic =>
                diagnostic.LossKind != OfficeConversionLossKind.None
                || diagnostic.Severity == HtmlDiagnosticSeverity.Warning
                || diagnostic.Severity == HtmlDiagnosticSeverity.Error)) {
            throw new HtmlConversionException(diagnostics);
        }
        if (options.FidelityPolicy == HtmlRenderFidelityPolicy.RequireNoLoss && conversionReport.HasLoss) {
            throw new HtmlConversionException(diagnostics.Concat(conversionReport.Warnings
                .Where(static warning => warning.LossKind != OfficeConversionLossKind.None)
                .Select(static warning => new HtmlDiagnostic(
                    warning.Converter,
                    warning.Code,
                    warning.Message,
                    warning.Severity == PdfCore.PdfConversionWarningSeverity.Error
                        ? HtmlDiagnosticSeverity.Error : HtmlDiagnosticSeverity.Warning,
                    warning.Source,
                    lossKind: warning.LossKind))));
        }

        cancellationToken.ThrowIfCancellationRequested();
        return new HtmlPdfRenderResult(pdf, diagnostics, conversionReport);
    }

    private static void AddPageOutlines(PdfCore.PdfPageCanvas canvas, IEnumerable<HtmlRenderHeading> headings, IReadOnlyDictionary<HtmlRenderHeading, int> headingDocumentOrder, double pageScale, CancellationToken cancellationToken) {
        foreach (HtmlRenderHeading heading in headings) {
            cancellationToken.ThrowIfCancellationRequested();
            canvas.Outline(heading.Text, heading.Level, Math.Max(0D, heading.Y * PointsPerCssPixel * pageScale), heading.BookmarkState switch {
                HtmlRenderBookmarkState.Open => PdfCore.PdfOutlineState.Open,
                HtmlRenderBookmarkState.Closed => PdfCore.PdfOutlineState.Closed,
                _ => PdfCore.PdfOutlineState.Default
            }, headingDocumentOrder[heading]);
        }
    }

    private static void AddPageVisuals(PdfCore.PdfPageCanvas canvas, HtmlRenderPage page, RegisteredWebFonts webFonts, PdfCore.PdfConversionReport conversionReport, bool interactiveFormControls, CancellationToken cancellationToken) {
        foreach (HtmlRenderVisual visual in page.Scene.OrderBy(item => item.PaintOrder)) {
            cancellationToken.ThrowIfCancellationRequested();
            AddVisual(canvas, visual, webFonts, conversionReport, page.Width, page.Height, interactiveFormControls, cancellationToken);
        }
    }

    private static void AddVisual(
        PdfCore.PdfPageCanvas canvas,
        HtmlRenderVisual visual,
        RegisteredWebFonts webFonts,
        PdfCore.PdfConversionReport conversionReport,
        double surfaceWidth,
        double surfaceHeight,
        bool interactiveFormControls,
        CancellationToken cancellationToken,
        bool textAsSpan = false,
        ClipBounds? activeClip = null,
        bool logicalTextOwned = false) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!activeClip.HasValue && !IsVisualContainer(visual) && ExceedsSurface(visual, surfaceWidth, surfaceHeight)) {
            double right = visual.X + visual.Width;
            double bottom = visual.Y + visual.Height;
            if (right <= 0D || bottom <= 0D || visual.X >= surfaceWidth || visual.Y >= surfaceHeight) return;
            var pageClip = new ClipBounds(0D, 0D, surfaceWidth, surfaceHeight);
            canvas.Clip(0D, 0D, surfaceWidth * PointsPerCssPixel, surfaceHeight * PointsPerCssPixel, clipped =>
                AddVisual(clipped, visual, webFonts, conversionReport, surfaceWidth, surfaceHeight,
                    interactiveFormControls, cancellationToken, textAsSpan, pageClip, logicalTextOwned));
            return;
        }
        if (visual is HtmlRenderFormField formField) {
            bool fullyContained = !activeClip.HasValue || activeClip.Value.AllowsInteractiveWidgets && activeClip.Value.Contains(formField);
            AddFormField(canvas, formField, webFonts, conversionReport, surfaceWidth, surfaceHeight, interactiveFormControls && fullyContained, cancellationToken, textAsSpan, activeClip, logicalTextOwned);
        } else if (visual is HtmlRenderShape shape) {
            AddShape(canvas, shape, conversionReport, cancellationToken);
        } else if (visual is HtmlRenderText text) {
            AddText(canvas, text, webFonts, conversionReport, surfaceWidth, textAsSpan, logicalTextOwned, cancellationToken);
        } else if (visual is HtmlRenderNamedDestination destination) {
            canvas.NamedDestination(
                MapNamedDestination(destination.Name),
                destination.X * PointsPerCssPixel,
                destination.Y * PointsPerCssPixel);
        } else if (visual is HtmlRenderImage image) {
            try {
                if (!AddImage(canvas, image)) {
                    ReportImagePayloadOmitted(conversionReport, image,
                        "The image payload could not be prepared as a PDF raster resource.");
                }
            } catch (System.NotSupportedException exception) {
                ReportImagePayloadOmitted(conversionReport, image, exception.Message);
            }
        } else if (visual is HtmlRenderDrawing drawing) {
            AddDrawing(canvas, drawing, webFonts, conversionReport, cancellationToken);
        } else if (visual is HtmlRenderImagePattern imagePattern) {
            AddImagePattern(canvas, imagePattern, cancellationToken);
        } else if (visual is HtmlRenderClipGroup group) {
            AddClipGroup(canvas, group, webFonts, conversionReport, surfaceWidth, surfaceHeight, interactiveFormControls, cancellationToken, textAsSpan, activeClip, logicalTextOwned);
        } else if (visual is HtmlRenderPathClipGroup pathClipGroup) {
            AddPathClipGroup(canvas, pathClipGroup, webFonts, conversionReport, surfaceWidth, surfaceHeight, interactiveFormControls, cancellationToken, textAsSpan, activeClip, logicalTextOwned);
        } else if (visual is HtmlRenderEffectGroup effectGroup) {
            AddEffectGroup(canvas, effectGroup, webFonts, conversionReport, surfaceWidth, surfaceHeight, interactiveFormControls, cancellationToken, textAsSpan, activeClip, logicalTextOwned);
        } else if (visual is HtmlRenderLayoutRegion layoutRegion) {
            foreach (HtmlRenderVisual child in layoutRegion.Visuals.OrderBy(item => item.PaintOrder)) {
                AddVisual(canvas, child, webFonts, conversionReport, surfaceWidth, surfaceHeight, interactiveFormControls, cancellationToken, textAsSpan, activeClip, logicalTextOwned);
            }
        } else if (visual is HtmlRenderSemanticGroup semanticGroup) {
            AddSemanticGroup(canvas, semanticGroup, webFonts, conversionReport, surfaceWidth, surfaceHeight, interactiveFormControls, cancellationToken, textAsSpan, activeClip, logicalTextOwned);
        } else if (visual is HtmlRenderLogicalTextGroup logicalTextGroup) {
            AddLogicalTextGroup(canvas, logicalTextGroup, webFonts, conversionReport, surfaceWidth, surfaceHeight, interactiveFormControls, cancellationToken, textAsSpan, activeClip, logicalTextOwned);
        }
    }

    private static bool ExceedsSurface(HtmlRenderVisual visual, double surfaceWidth, double surfaceHeight) =>
        visual.X < -0.001D || visual.Y < -0.001D
        || visual.X + visual.Width > surfaceWidth + 0.001D
        || visual.Y + visual.Height > surfaceHeight + 0.001D;

    private static bool IsVisualContainer(HtmlRenderVisual visual) =>
        visual is HtmlRenderClipGroup
        or HtmlRenderPathClipGroup
        or HtmlRenderEffectGroup
        or HtmlRenderSemanticGroup
        or HtmlRenderLogicalTextGroup
        or HtmlRenderLayoutRegion;

    private static void AddFormField(PdfCore.PdfPageCanvas canvas, HtmlRenderFormField field, RegisteredWebFonts webFonts, PdfCore.PdfConversionReport conversionReport, double surfaceWidth, double surfaceHeight, bool interactiveFormControls, CancellationToken cancellationToken, bool textAsSpan, ClipBounds? activeClip, bool logicalTextOwned) {
        bool hasInvalidPdfButtonValue = (field.FieldKind == HtmlRenderFormFieldKind.CheckBox || field.FieldKind == HtmlRenderFormFieldKind.RadioButton)
            && (string.IsNullOrWhiteSpace(field.Value) || string.IsNullOrWhiteSpace(field.RadioOption));
        if (!interactiveFormControls
            || string.IsNullOrWhiteSpace(field.Name)
            || hasInvalidPdfButtonValue
            || field.FieldKind == HtmlRenderFormFieldKind.Choice && field.Options.Count == 0) {
            foreach (HtmlRenderVisual child in field.Visuals.OrderBy(item => item.PaintOrder)) {
                AddVisual(canvas, child, webFonts, conversionReport, surfaceWidth, surfaceHeight, interactiveFormControls, cancellationToken, textAsSpan, activeClip, logicalTextOwned);
            }
            return;
        }

        var style = new PdfCore.PdfFormFieldStyle {
            BackgroundColor = field.BackgroundColor.HasValue ? PdfCore.PdfColor.FromOfficeColorOrNull(field.BackgroundColor.Value) : null,
            BorderColor = field.BorderColor.HasValue ? PdfCore.PdfColor.FromOfficeColorOrNull(field.BorderColor.Value) : null,
            BorderWidth = field.BorderWidth * PointsPerCssPixel,
            BorderStyle = field.BorderStyle == "dashed" ? PdfCore.PdfFormFieldBorderStyle.Dashed : PdfCore.PdfFormFieldBorderStyle.Solid,
            CornerRadius = field.CornerRadius * PointsPerCssPixel,
            TextColor = PdfCore.PdfColor.FromOfficeColorOrNull(field.TextColor) ?? PdfCore.PdfColor.Black,
            MarkColor = PdfCore.PdfColor.FromOfficeColorOrNull(field.TextColor) ?? PdfCore.PdfColor.Black,
            IsReadOnly = field.IsReadOnly,
            IsNoExport = field.IsDisabled || field.MappingName.Length == 0,
            IsRequired = field.IsRequired,
            IsMultiline = field.IsMultiline,
            IsPassword = field.IsPassword,
            IsFileSelect = field.IsFileSelect,
            MaxLength = field.MaximumLength,
            AlternateName = field.AlternateName,
            MappingName = field.MappingName.Length == 0 ? null : field.MappingName,
            TextAlignment = MapFormFieldTextAlignment(field.TextAlignment, field.FieldKind)
        };
        double x = field.X * PointsPerCssPixel;
        double y = field.Y * PointsPerCssPixel;
        double width = field.Width * PointsPerCssPixel;
        double height = field.Height * PointsPerCssPixel;
        double fontSize = Math.Max(1D, field.Font.Size * PointsPerCssPixel);
        if (field.FieldKind == HtmlRenderFormFieldKind.Text) {
            if (field.IsPassword && field.Value.Length > 0) {
                string maskedValue = new('*', field.Value.Length);
                canvas.TextFieldWithInitialAppearance(field.Name, string.Empty, maskedValue, x, y, width, height, fontSize, style, style);
            } else if (field.Value.Length == 0 && field.Placeholder.Length > 0) {
                PdfCore.PdfFormFieldStyle appearanceStyle = style.Clone();
                appearanceStyle.TextColor = PdfCore.PdfColor.FromOfficeColorOrNull(field.PlaceholderTextColor) ?? PdfCore.PdfColor.Black;
                canvas.TextFieldWithInitialAppearance(field.Name, field.Value, field.Placeholder, x, y, width, height, fontSize, style, appearanceStyle);
            } else {
                canvas.TextField(field.Name, field.Value, x, y, width, height, fontSize, style);
            }
            if (!field.IsPassword && !string.IsNullOrWhiteSpace(field.Value)) {
                canvas.SearchableText(field.Value, x, y + Math.Min(height, fontSize));
            }
        } else if (field.FieldKind == HtmlRenderFormFieldKind.CheckBox) {
            canvas.CheckBoxWithExportValue(field.Name, field.IsSelected, x, y, width, height, field.RadioOption ?? "Yes", field.Value, style);
        } else if (field.FieldKind == HtmlRenderFormFieldKind.Choice) {
            IReadOnlyList<PdfCore.PdfFormFieldOption> choiceOptions = field.Options
                .Select((label, index) => new PdfCore.PdfFormFieldOption(
                    index < field.OptionValues.Count ? field.OptionValues[index] : label,
                    label))
                .ToList();
            IReadOnlyList<string>? selectedValues = !field.AllowsMultipleSelection && field.Values.Count == 0 ? null : field.Values;
            IReadOnlyList<int> selectedIndices = field.IsComboBox ? Array.Empty<int>() : field.SelectedOptionIndices;
            canvas.ChoiceFieldWithSelectedIndices(field.Name, choiceOptions, selectedValues, selectedIndices, x, y, width, height, fontSize, field.IsComboBox, field.AllowsMultipleSelection, style);
            IEnumerable<string> searchableLabels = selectedIndices.Count > 0
                ? selectedIndices.Where(index => index >= 0 && index < choiceOptions.Count).Select(index => choiceOptions[index].DisplayText)
                : field.Values.Select(value => choiceOptions.FirstOrDefault(option => string.Equals(option.ExportValue, value, StringComparison.Ordinal))?.DisplayText ?? value);
            string searchableValue = string.Join(" ", searchableLabels
                .Where(value => !string.IsNullOrWhiteSpace(value)));
            if (searchableValue.Length > 0) {
                canvas.SearchableText(searchableValue, x, y + Math.Min(height, fontSize));
            }
        } else {
            canvas.RadioButtonWithExportValue(field.Name, field.RadioOption!, field.Value, field.IsSelected, x, y, width, height, style);
        }
    }

    private static PdfCore.PdfFormFieldTextAlignment? MapFormFieldTextAlignment(OfficeTextAlignment alignment, HtmlRenderFormFieldKind fieldKind) =>
        fieldKind == HtmlRenderFormFieldKind.CheckBox || fieldKind == HtmlRenderFormFieldKind.RadioButton
            ? null
            : alignment == OfficeTextAlignment.Center
                ? PdfCore.PdfFormFieldTextAlignment.Center
                : alignment == OfficeTextAlignment.Right
                    ? PdfCore.PdfFormFieldTextAlignment.Right
                    : PdfCore.PdfFormFieldTextAlignment.Left;

    private static void AddLogicalTextGroup(PdfCore.PdfPageCanvas canvas, HtmlRenderLogicalTextGroup group, RegisteredWebFonts webFonts, PdfCore.PdfConversionReport conversionReport, double surfaceWidth, double surfaceHeight, bool interactiveFormControls, CancellationToken cancellationToken, bool textAsSpan, ClipBounds? activeClip, bool logicalTextOwned) {
        void AddChildren(PdfCore.PdfPageCanvas target) {
            foreach (HtmlRenderVisual child in group.Visuals.OrderBy(item => item.PaintOrder)) {
                cancellationToken.ThrowIfCancellationRequested();
                AddVisual(target, child, webFonts, conversionReport, surfaceWidth, surfaceHeight, interactiveFormControls, cancellationToken, textAsSpan, activeClip, logicalTextOwned: true);
            }
        }
        if (!group.Visuals.Any(child => ContainsPdfRenderableVisual(child, webFonts, surfaceWidth, surfaceHeight, activeClip, cancellationToken))) {
            AddChildren(canvas);
            return;
        }
        var content = new PdfCore.PdfPageCanvas(allowOutOfPageCoordinates: true);
        AddChildren(content);
        if (!HasCanvasContent(content.Items)) {
            canvas.AddItems(content.Items);
            return;
        }
        if (group.Text.Length == 0) {
            canvas.Artifact(nested => nested.AddItems(content.Items));
            return;
        }
        string? logicalText = FilterLogicalPrivateUseGlyphs(group.Text, group.Visuals, webFonts, cancellationToken);
        if (logicalText == null) {
            canvas.AddItems(content.Items);
            return;
        }
        if (logicalText.Length == 0) {
            canvas.AddItems(content.Items);
            return;
        }
        if (logicalTextOwned) canvas.AddItems(content.Items);
        else canvas.ActualText(
            logicalText,
            group.X * PointsPerCssPixel,
            (group.Y + Math.Min(group.Height, 12D)) * PointsPerCssPixel,
            nested => nested.AddItems(content.Items));
    }

    private static void AddSemanticGroup(PdfCore.PdfPageCanvas canvas, HtmlRenderSemanticGroup group, RegisteredWebFonts webFonts, PdfCore.PdfConversionReport conversionReport, double surfaceWidth, double surfaceHeight, bool interactiveFormControls, CancellationToken cancellationToken, bool textAsSpan, ClipBounds? activeClip, bool logicalTextOwned) {
        if (!group.Visuals.Any(child => ContainsPdfRenderableVisual(child, webFonts, surfaceWidth, surfaceHeight, activeClip, cancellationToken))) {
            // Navigation-only groups still carry named destinations. They cannot create
            // an empty structure element. The same path handles groups whose paint is
            // entirely outside the page, while non-painting children still reach the
            // page canvas so empty anchors remain valid link targets.
            foreach (HtmlRenderVisual child in group.Visuals.OrderBy(item => item.PaintOrder)) {
                AddVisual(canvas, child, webFonts, conversionReport, surfaceWidth, surfaceHeight, interactiveFormControls, cancellationToken, textAsSpan, activeClip, logicalTextOwned);
            }
            return;
        }
        if (group.Role == HtmlRenderSemanticGroupRole.Artifact) {
            var artifactContent = new PdfCore.PdfPageCanvas(allowOutOfPageCoordinates: true);
            foreach (HtmlRenderVisual child in group.Visuals.OrderBy(item => item.PaintOrder)) {
                cancellationToken.ThrowIfCancellationRequested();
                AddVisual(artifactContent, child, webFonts, conversionReport, surfaceWidth, surfaceHeight, interactiveFormControls: false, cancellationToken, textAsSpan: true, activeClip: activeClip, logicalTextOwned: logicalTextOwned);
            }
            if (HasCanvasContent(artifactContent.Items))
                canvas.Artifact(nested => nested.AddItems(artifactContent.Items));
            else canvas.AddItems(artifactContent.Items);
            return;
        }
        var options = new PdfCore.PdfCanvasStructureOptions {
            ColumnSpan = group.ColumnSpan,
            RowSpan = group.RowSpan,
            HeaderScope = MapTableHeaderScope(group.HeaderScope),
            StructureElementKey = group.StructureElementKey
        };
        bool childTextAsSpan = textAsSpan || IsTextContentGroup(group.Role);
        string logicalText = string.Empty;
        bool hasLogicalText = IsTextContentGroup(group.Role)
            && TryResolveReorderedLogicalText(group.Visuals, out logicalText);
        string? printableText = hasLogicalText
            ? FilterLogicalPrivateUseGlyphs(logicalText, group.Visuals, webFonts, cancellationToken)
            : null;
        var content = new PdfCore.PdfPageCanvas(allowOutOfPageCoordinates: true);
        foreach (HtmlRenderVisual child in group.Visuals.OrderBy(item => item.PaintOrder)) {
            cancellationToken.ThrowIfCancellationRequested();
            AddVisual(content, child, webFonts, conversionReport, surfaceWidth, surfaceHeight, interactiveFormControls,
                cancellationToken, childTextAsSpan, activeClip, logicalTextOwned || !string.IsNullOrEmpty(printableText));
        }
        if (!HasCanvasContent(content.Items)) {
            canvas.AddItems(content.Items);
            return;
        }
        canvas.Structure(MapSemanticGroupRole(group.Role), nested => {
            if (printableText is { Length: > 0 } actualText)
                nested.ActualText(actualText, target => target.AddItems(content.Items));
            else nested.AddItems(content.Items);
        }, options);
    }

    private static bool ContainsRenderableVisual(HtmlRenderVisual visual, double surfaceWidth, double surfaceHeight, ClipBounds? activeClip) {
        if (!ContainsPaintableVisual(visual)) return false;
        if (activeClip.HasValue) return true;
        if (visual is HtmlRenderLayoutRegion layoutRegion)
            return layoutRegion.Visuals.Any(child => ContainsRenderableVisual(child, surfaceWidth, surfaceHeight, activeClip));
        if (visual is HtmlRenderSemanticGroup semanticGroup)
            return semanticGroup.Visuals.Any(child => ContainsRenderableVisual(child, surfaceWidth, surfaceHeight, activeClip));
        if (visual is HtmlRenderLogicalTextGroup logicalTextGroup)
            return logicalTextGroup.Visuals.Any(child => ContainsRenderableVisual(child, surfaceWidth, surfaceHeight, activeClip));
        if (visual is HtmlRenderClipGroup or HtmlRenderPathClipGroup or HtmlRenderEffectGroup)
            return IntersectsSurface(visual, surfaceWidth, surfaceHeight);
        return IntersectsSurface(visual, surfaceWidth, surfaceHeight);
    }

    private static bool IntersectsSurface(HtmlRenderVisual visual, double surfaceWidth, double surfaceHeight) =>
        visual.X + visual.Width > 0D && visual.Y + visual.Height > 0D
        && visual.X < surfaceWidth && visual.Y < surfaceHeight;

    private static bool ContainsPaintableVisual(HtmlRenderVisual visual) {
        if (visual is HtmlRenderBookmarkAnchor || visual is HtmlRenderNamedDestination) return false;
        if (visual is HtmlRenderLayoutRegion layoutRegion) return layoutRegion.Visuals.Any(ContainsPaintableVisual);
        if (visual is HtmlRenderSemanticGroup semanticGroup) return semanticGroup.Visuals.Any(ContainsPaintableVisual);
        if (visual is HtmlRenderLogicalTextGroup logicalTextGroup) return logicalTextGroup.Visuals.Any(ContainsPaintableVisual);
        if (visual is HtmlRenderClipGroup clipGroup) return clipGroup.Visuals.Any(ContainsPaintableVisual);
        if (visual is HtmlRenderPathClipGroup pathClipGroup) return pathClipGroup.Visuals.Any(ContainsPaintableVisual);
        if (visual is HtmlRenderEffectGroup effectGroup) return effectGroup.Visuals.Any(ContainsPaintableVisual);
        return true;
    }

    private static bool IsTextContentGroup(HtmlRenderSemanticGroupRole role) =>
        role == HtmlRenderSemanticGroupRole.Paragraph
        || role == HtmlRenderSemanticGroupRole.Heading1
        || role == HtmlRenderSemanticGroupRole.Heading2
        || role == HtmlRenderSemanticGroupRole.Heading3
        || role == HtmlRenderSemanticGroupRole.Heading4
        || role == HtmlRenderSemanticGroupRole.Heading5
        || role == HtmlRenderSemanticGroupRole.Heading6;

    private static PdfCore.PdfCanvasTableHeaderScope? MapTableHeaderScope(HtmlRenderTableHeaderScope? scope) {
        if (scope == HtmlRenderTableHeaderScope.Row) return PdfCore.PdfCanvasTableHeaderScope.Row;
        if (scope == HtmlRenderTableHeaderScope.Column) return PdfCore.PdfCanvasTableHeaderScope.Column;
        if (scope == HtmlRenderTableHeaderScope.Both) return PdfCore.PdfCanvasTableHeaderScope.Both;
        return null;
    }

    private static PdfCore.PdfCanvasStructureRole MapSemanticGroupRole(HtmlRenderSemanticGroupRole role) {
        if (role == HtmlRenderSemanticGroupRole.Section) return PdfCore.PdfCanvasStructureRole.Section;
        if (role == HtmlRenderSemanticGroupRole.Division) return PdfCore.PdfCanvasStructureRole.Division;
        if (role == HtmlRenderSemanticGroupRole.Paragraph) return PdfCore.PdfCanvasStructureRole.Paragraph;
        if (role == HtmlRenderSemanticGroupRole.Heading1) return PdfCore.PdfCanvasStructureRole.Heading1;
        if (role == HtmlRenderSemanticGroupRole.Heading2) return PdfCore.PdfCanvasStructureRole.Heading2;
        if (role == HtmlRenderSemanticGroupRole.Heading3) return PdfCore.PdfCanvasStructureRole.Heading3;
        if (role == HtmlRenderSemanticGroupRole.Heading4) return PdfCore.PdfCanvasStructureRole.Heading4;
        if (role == HtmlRenderSemanticGroupRole.Heading5) return PdfCore.PdfCanvasStructureRole.Heading5;
        if (role == HtmlRenderSemanticGroupRole.Heading6) return PdfCore.PdfCanvasStructureRole.Heading6;
        if (role == HtmlRenderSemanticGroupRole.List) return PdfCore.PdfCanvasStructureRole.List;
        if (role == HtmlRenderSemanticGroupRole.ListItem) return PdfCore.PdfCanvasStructureRole.ListItem;
        if (role == HtmlRenderSemanticGroupRole.ListLabel) return PdfCore.PdfCanvasStructureRole.ListLabel;
        if (role == HtmlRenderSemanticGroupRole.ListBody) return PdfCore.PdfCanvasStructureRole.ListBody;
        if (role == HtmlRenderSemanticGroupRole.Table) return PdfCore.PdfCanvasStructureRole.Table;
        if (role == HtmlRenderSemanticGroupRole.TableRow) return PdfCore.PdfCanvasStructureRole.TableRow;
        if (role == HtmlRenderSemanticGroupRole.TableHeaderCell) return PdfCore.PdfCanvasStructureRole.TableHeaderCell;
        if (role == HtmlRenderSemanticGroupRole.TableCell) return PdfCore.PdfCanvasStructureRole.TableCell;
        if (role == HtmlRenderSemanticGroupRole.Footnote) return PdfCore.PdfCanvasStructureRole.Note;
        return PdfCore.PdfCanvasStructureRole.Caption;
    }

    private static void AddEffectGroup(
        PdfCore.PdfPageCanvas canvas,
        HtmlRenderEffectGroup group,
        RegisteredWebFonts webFonts,
        PdfCore.PdfConversionReport conversionReport,
        double surfaceWidth,
        double surfaceHeight,
        bool interactiveFormControls,
        CancellationToken cancellationToken,
        bool textAsSpan,
        ClipBounds? activeClip,
        bool logicalTextOwned) {
        OfficeTransform transform = group.Transform;
        var scaled = new OfficeTransform(
            transform.M11,
            transform.M12,
            transform.M21,
            transform.M22,
            transform.OffsetX * PointsPerCssPixel,
            transform.OffsetY * PointsPerCssPixel);
        canvas.Effect(scaled, group.Opacity, nested => {
            foreach (HtmlRenderVisual child in group.Visuals.OrderBy(item => item.PaintOrder)) {
                cancellationToken.ThrowIfCancellationRequested();
                AddVisual(nested, child, webFonts, conversionReport, surfaceWidth, surfaceHeight, interactiveFormControls, cancellationToken, textAsSpan, ClipBounds.TransformedCoordinateSpace, logicalTextOwned);
            }
        });
    }

    private static void AddClipGroup(
        PdfCore.PdfPageCanvas canvas,
        HtmlRenderClipGroup group,
        RegisteredWebFonts webFonts,
        PdfCore.PdfConversionReport conversionReport,
        double surfaceWidth,
        double surfaceHeight,
        bool interactiveFormControls,
        CancellationToken cancellationToken,
        bool textAsSpan,
        ClipBounds? activeClip,
        bool logicalTextOwned) {
        bool constrainToSurface = !activeClip.HasValue || activeClip.Value.ConstrainToSurface;
        double left = group.ClipHorizontal ? (constrainToSurface ? Math.Max(0D, group.ClipX) : group.ClipX) : (constrainToSurface ? 0D : group.X);
        double top = group.ClipVertical ? (constrainToSurface ? Math.Max(0D, group.ClipY) : group.ClipY) : (constrainToSurface ? 0D : group.Y);
        double right = group.ClipHorizontal ? (constrainToSurface ? Math.Min(surfaceWidth, group.ClipX + group.ClipWidth) : group.ClipX + group.ClipWidth) : (constrainToSurface ? surfaceWidth : group.X + group.Width);
        double bottom = group.ClipVertical ? (constrainToSurface ? Math.Min(surfaceHeight, group.ClipY + group.ClipHeight) : group.ClipY + group.ClipHeight) : (constrainToSurface ? surfaceHeight : group.Y + group.Height);
        if (right <= left + 0.0001D || bottom <= top + 0.0001D) return;
        ClipBounds clip = ClipBounds.Intersect(activeClip, new ClipBounds(left, top, right, bottom, constrainToSurface: constrainToSurface));
        canvas.Clip(
            left * PointsPerCssPixel,
            top * PointsPerCssPixel,
            (right - left) * PointsPerCssPixel,
            (bottom - top) * PointsPerCssPixel,
            clipped => {
                foreach (HtmlRenderVisual child in group.Visuals.OrderBy(item => item.PaintOrder)) {
                    cancellationToken.ThrowIfCancellationRequested();
                    AddVisual(clipped, child, webFonts, conversionReport, surfaceWidth, surfaceHeight, interactiveFormControls, cancellationToken, textAsSpan, clip, logicalTextOwned);
                }
            });
    }

    private static void AddPathClipGroup(
        PdfCore.PdfPageCanvas canvas,
        HtmlRenderPathClipGroup group,
        RegisteredWebFonts webFonts,
        PdfCore.PdfConversionReport conversionReport,
        double surfaceWidth,
        double surfaceHeight,
        bool interactiveFormControls,
        CancellationToken cancellationToken,
        bool textAsSpan,
        ClipBounds? activeClip,
        bool logicalTextOwned) {
        ClipBounds clip = ClipBounds.Intersect(activeClip, new ClipBounds(
            group.X,
            group.Y,
            group.X + group.Width,
            group.Y + group.Height,
            allowsInteractiveWidgets: false));
        double clipX = group.ClipX * PointsPerCssPixel;
        double clipY = group.ClipY * PointsPerCssPixel;
        OfficeClipPath clipPath = group.ClipPath.Scale(PointsPerCssPixel, PointsPerCssPixel);
        Action<PdfCore.PdfPageCanvas> addClip = target => target.Clip(
            clipX,
            clipY,
            clipPath,
            clipped => {
                foreach (HtmlRenderVisual child in group.Visuals.OrderBy(item => item.PaintOrder)) {
                    cancellationToken.ThrowIfCancellationRequested();
                    AddVisual(clipped, child, webFonts, conversionReport, surfaceWidth, surfaceHeight, interactiveFormControls, cancellationToken, textAsSpan, clip, logicalTextOwned);
                }
            });
        if (clipX < 0D || clipY < 0D) {
            canvas.Effect(OfficeTransform.Identity, 1D, addClip);
        } else if (!activeClip.HasValue
                   && (clipX + clipPath.Width > surfaceWidth * PointsPerCssPixel
                       || clipY + clipPath.Height > surfaceHeight * PointsPerCssPixel)) {
            canvas.Clip(0D, 0D, surfaceWidth * PointsPerCssPixel, surfaceHeight * PointsPerCssPixel, addClip);
        } else {
            addClip(canvas);
        }
    }

    private static void AddShape(
        PdfCore.PdfPageCanvas canvas,
        HtmlRenderShape visual,
        PdfCore.PdfConversionReport conversionReport,
        CancellationToken cancellationToken) {
        var drawing = new OfficeDrawing(visual.Width, visual.Height);
        drawing.AddShape(visual.Shape.Clone(), 0D, 0D);
        if (TryAddTranslucentGradient(canvas, visual, drawing, conversionReport, cancellationToken)) return;
        double x = visual.X * PointsPerCssPixel;
        double y = visual.Y * PointsPerCssPixel;
        double drawingX = Math.Max(0D, x);
        double drawingY = Math.Max(0D, y);
        OfficeTransform transform = OfficeTransform.Translate(Math.Min(0D, x), Math.Min(0D, y));
        bool fragmentLink = IsFragmentLink(visual.LinkUri);
        Action<PdfCore.PdfPageCanvas> addDrawing = target => target.Drawing(
            drawing,
            drawingX,
            drawingY,
            visual.Width * PointsPerCssPixel,
            visual.Height * PointsPerCssPixel,
            style: new PdfCore.PdfDrawingStyle { Decorative = true },
            linkUri: fragmentLink ? null : visual.LinkUri,
            linkContents: visual.LinkUri == null || fragmentLink ? null : visual.Source);
        if (transform == OfficeTransform.Identity) {
            addDrawing(canvas);
        } else {
            canvas.Effect(transform, 1D, addDrawing);
        }
        if (fragmentLink) {
            canvas.LinkToNamedDestination(
                MapNamedDestination(visual.LinkUri!.Substring(1)),
                visual.X * PointsPerCssPixel,
                visual.Y * PointsPerCssPixel,
                visual.Width * PointsPerCssPixel,
                visual.Height * PointsPerCssPixel,
                visual.Source);
        }
    }

    private static void AddText(
        PdfCore.PdfPageCanvas canvas,
        HtmlRenderText visual,
        RegisteredWebFonts webFonts,
        PdfCore.PdfConversionReport conversionReport,
        double surfaceWidth,
        bool asSpan,
        bool logicalTextOwned,
        CancellationToken cancellationToken,
        double? baselineFontSize = null,
        bool colorOpacityApplied = false) {
        if (visual.Text.Length == 0) return;
        // Canvas and outline writers must anchor a script to the original line's metrics.
        baselineFontSize ??= visual.Font.Size;
        visual = visual.ResolveBaselineForPainting();
        if (visual.Y < 0D) {
            HtmlRenderText shifted = (HtmlRenderText)visual.TranslatePaint(0D, -visual.Y, visual.PaintOrder);
            canvas.Effect(OfficeTransform.Translate(0D, visual.Y * PointsPerCssPixel), 1D,
                nested => AddText(nested, shifted, webFonts, conversionReport, surfaceWidth, asSpan, logicalTextOwned, cancellationToken, baselineFontSize));
            return;
        }
        string? link = string.IsNullOrWhiteSpace(visual.Text) || IsFragmentLink(visual.LinkUri) ? null : visual.LinkUri;
        string? linkDestination = IsFragmentLink(visual.LinkUri)
            ? MapNamedDestination(visual.LinkUri!.Substring(1))
            : null;
        double frameWidth = visual.Width;
        if (visual.TextAdvanceWidth.HasValue) {
            double metricTolerance = Math.Max(
                visual.Font.Size,
                visual.TextAdvanceWidth.Value * 0.25D);
            frameWidth = Math.Max(frameWidth, visual.TextAdvanceWidth.Value + metricTolerance);
        }
        frameWidth = Math.Max(0.01D, Math.Min(frameWidth, Math.Max(0.01D, surfaceWidth - visual.X)));
        if (TryAddOutlinedText(
                canvas,
                visual,
                webFonts,
                conversionReport,
                frameWidth,
                asSpan,
                logicalTextOwned,
                cancellationToken,
                baselineFontSize.Value)) {
            return;
        }
        if (!colorOpacityApplied && visual.Color.A < 255) {
            canvas.Effect(OfficeTransform.Identity, visual.Color.A / 255D,
                nested => AddText(nested, visual, webFonts, conversionReport, surfaceWidth,
                    asSpan, logicalTextOwned, cancellationToken, baselineFontSize, colorOpacityApplied: true));
            return;
        }
        OfficeFontStyle requestedStyle = (visual.Font.IsBold ? OfficeFontStyle.Bold : OfficeFontStyle.Regular)
            | (visual.Font.IsItalic ? OfficeFontStyle.Italic : OfficeFontStyle.Regular);
        string pdfText = OmitUnavailablePrivateUseGlyphs(visual, webFonts, conversionReport, requestedStyle, cancellationToken);
        if (pdfText.Length == 0) return;
        var runs = webFonts.Faces.PlanFallbackRuns(pdfText, visual.Font.FamilyName, requestedStyle)
            .Select(fallbackRun => new PdfCore.PdfTextRun(
            fallbackRun.Text,
            bold: visual.Font.IsBold,
            underline: visual.Font.IsUnderline,
            color: PdfCore.PdfColor.FromOfficeColorOrNull(visual.Color),
            italic: visual.Font.IsItalic,
            strike: visual.Font.IsStrikethrough,
            fontSize: visual.Font.Size * PointsPerCssPixel,
            font: MapFont(
                fallbackRun.FamilyName,
                fallbackRun.Text,
                requestedStyle,
                webFonts),
            linkUri: link,
            linkContents: link == null ? null : pdfText,
            linkDestinationName: linkDestination,
            fontFamily: fallbackRun.FamilyName,
            baseline: MapTextBaseline(visual.Baseline),
            underlineStyle: visual.UnderlineStyle,
            strikeStyle: visual.StrikethroughStyle,
            decorationColor: PdfCore.PdfColor.FromOfficeColorOrNull(visual.DecorationColor))
            .WithFeatureSettings(visual.FeatureSettings)).ToArray();
        canvas.PositionedText(
            runs,
            asSpan ? PdfCore.PdfCanvasTextStructureRole.Span : MapStructureRole(visual.SemanticRole),
            visual.X * PointsPerCssPixel,
            visual.Y * PointsPerCssPixel,
            frameWidth * PointsPerCssPixel,
            visual.Height * PointsPerCssPixel,
            PdfCore.PdfColor.FromOfficeColorOrNull(visual.Color),
            MapAlignment(visual.Alignment),
            baselineFontSize.Value * PointsPerCssPixel,
            visual.LineHeight * PointsPerCssPixel,
            (visual.TextPaintWidth ?? visual.TextAdvanceWidth) * PointsPerCssPixel);
    }

    private static bool IsFragmentLink(string? link) =>
        link != null && link.Length > 1 && link[0] == '#';

    private static string DecodeFragmentName(string name) {
        try {
            return System.Uri.UnescapeDataString(name);
        } catch (System.UriFormatException) {
            return name;
        }
    }

    private static string MapNamedDestination(string name) => "html-fragment:" + name;

    private static PdfCore.PdfCanvasTextStructureRole MapStructureRole(string? semanticRole) {
        if (semanticRole == "heading-1") return PdfCore.PdfCanvasTextStructureRole.Heading1;
        if (semanticRole == "heading-2") return PdfCore.PdfCanvasTextStructureRole.Heading2;
        if (semanticRole == "heading-3") return PdfCore.PdfCanvasTextStructureRole.Heading3;
        if (semanticRole == "heading-4") return PdfCore.PdfCanvasTextStructureRole.Heading4;
        if (semanticRole == "heading-5") return PdfCore.PdfCanvasTextStructureRole.Heading5;
        if (semanticRole == "heading-6") return PdfCore.PdfCanvasTextStructureRole.Heading6;
        return semanticRole == "span" ? PdfCore.PdfCanvasTextStructureRole.Span : PdfCore.PdfCanvasTextStructureRole.Paragraph;
    }

    private static PdfCore.PdfTextBaseline MapTextBaseline(OfficeTextBaseline baseline) => baseline switch {
        OfficeTextBaseline.Superscript => PdfCore.PdfTextBaseline.Superscript,
        OfficeTextBaseline.Subscript => PdfCore.PdfTextBaseline.Subscript,
        _ => PdfCore.PdfTextBaseline.Normal
    };

    private static void ReportImagePayloadOmitted(PdfCore.PdfConversionReport report, HtmlRenderImage image, string reason) {
        report.Add(new PdfCore.PdfConversionWarning(
            "OfficeIMO.Html.Pdf",
            HtmlPdfDiagnosticCodes.ImagePayloadOmitted,
            image.Source ?? "html-image",
            "An image payload could not be embedded in the PDF: " + reason,
            PdfCore.PdfConversionWarningSeverity.Warning,
            OfficeConversionLossKind.Omission));
    }

    private static bool AddImage(PdfCore.PdfPageCanvas canvas, HtmlRenderImage visual) {
        PdfCore.PdfCanvasImageResource? imageResource = GetSharedPdfImageResource(
            visual.EncodedBytes, visual.ContentType);
        if (imageResource == null) return false;
        PdfCore.PdfImageStyle? style = visual.SourceCrop.HasCrop
            ? new PdfCore.PdfImageStyle {
                SourceCrop = new PdfCore.PdfImageSourceCrop(
                    visual.SourceCrop.Left,
                    visual.SourceCrop.Top,
                    visual.SourceCrop.Right,
                    visual.SourceCrop.Bottom)
            }
            : null;
        bool fragmentLink = IsFragmentLink(visual.LinkUri);
        canvas.ImageShared(
            imageResource,
            visual.X * PointsPerCssPixel,
            visual.Y * PointsPerCssPixel,
            visual.Width * PointsPerCssPixel,
            visual.Height * PointsPerCssPixel,
            style,
            linkUri: fragmentLink ? null : visual.LinkUri,
            linkContents: visual.LinkUri == null || fragmentLink ? null : visual.Source,
            alternativeText: string.IsNullOrWhiteSpace(visual.AlternativeText) ? null : visual.AlternativeText);
        if (fragmentLink) {
            canvas.LinkToNamedDestination(
                MapNamedDestination(visual.LinkUri!.Substring(1)),
                visual.X * PointsPerCssPixel,
                visual.Y * PointsPerCssPixel,
                visual.Width * PointsPerCssPixel,
                visual.Height * PointsPerCssPixel,
                visual.Source);
        }
        return true;
    }

    private static void AddDrawing(
        PdfCore.PdfPageCanvas canvas,
        HtmlRenderDrawing visual,
        RegisteredWebFonts webFonts,
        PdfCore.PdfConversionReport conversionReport,
        CancellationToken cancellationToken) {
        OfficeDrawing source = visual.Drawing;
        double scaleX = visual.Width / source.Width;
        double scaleY = visual.Height / source.Height;
        if (TryAddRasterizedDrawingEffect(
                canvas,
                visual,
                source,
                Math.Max(scaleX, scaleY),
                conversionReport,
                cancellationToken)) return;
        double originX = visual.X * PointsPerCssPixel;
        double originY = visual.Y * PointsPerCssPixel;
        bool fragmentLink = IsFragmentLink(visual.LinkUri);
        string? drawingLinkUri = fragmentLink ? null : visual.LinkUri;
        OfficeTransform drawingToPage = OfficeTransform.Scale(scaleX * PointsPerCssPixel, scaleY * PointsPerCssPixel)
            .Then(OfficeTransform.Translate(originX, originY));
        OfficeTransform pageToDrawing = drawingToPage.Invert();

        void AddElements(PdfCore.PdfPageCanvas target, IReadOnlyList<OfficeDrawingElement> elements) {
            var shapeBatch = new OfficeDrawing(source.Width, source.Height);
            void FlushShapes() {
                if (shapeBatch.Elements.Count == 0) return;
                cancellationToken.ThrowIfCancellationRequested();
                target.DrawingForClippedRendering(
                    shapeBatch,
                    originX,
                    originY,
                    visual.Width * PointsPerCssPixel,
                    visual.Height * PointsPerCssPixel,
                    linkUri: drawingLinkUri,
                    linkContents: drawingLinkUri == null ? null : visual.Source);
                shapeBatch = new OfficeDrawing(source.Width, source.Height);
            }

            foreach (OfficeDrawingElement element in elements) {
                cancellationToken.ThrowIfCancellationRequested();
                if (element is OfficeDrawingShape shape) {
                    // Nested SVG clips, filters, and foreign-object viewports may retain
                    // deliberately out-of-bounds paint that is clipped by an ancestor group.
                    shapeBatch.AddShapeForClippedRendering(shape.Shape, shape.X, shape.Y);
                    continue;
                }
                if (element is OfficeDrawingGroup drawingGroup) {
                    FlushShapes();
                    OfficeTransform groupTransform = OfficeTransform.Translate(drawingGroup.X, drawingGroup.Y);
                    if (drawingGroup.FrameTransform.HasValue && drawingGroup.FrameTransform.Value.HasTransform) {
                        groupTransform = groupTransform.Then(drawingGroup.FrameTransform.Value.CreateDestinationTransform());
                    }
                    OfficeTransform pageGroupTransform = pageToDrawing
                        .Then(groupTransform)
                        .Then(drawingToPage);
                    OfficeDrawing nestedDrawing = drawingGroup.Drawing;
                    Action<PdfCore.PdfPageCanvas> addGroupPaint = groupTarget => groupTarget.Effect(pageGroupTransform, 1D, grouped => {
                        grouped.Clip(
                            originX,
                            originY,
                            drawingGroup.ClipPath.Scale(scaleX * PointsPerCssPixel, scaleY * PointsPerCssPixel),
                            clipped => {
                                OfficeTransform contentTransform = OfficeTransform.Translate(
                                    drawingGroup.ContentOffsetX,
                                    drawingGroup.ContentOffsetY);
                                if (contentTransform == OfficeTransform.Identity) {
                                    AddElements(clipped, nestedDrawing.Elements);
                                } else {
                                    OfficeTransform pageContentTransform = pageToDrawing
                                        .Then(contentTransform)
                                        .Then(drawingToPage);
                                    clipped.Effect(pageContentTransform, 1D,
                                        nested => AddElements(nested, nestedDrawing.Elements));
                                }
                            });
                    });
                    if (drawingGroup.ActualText == null) {
                        addGroupPaint(target);
                    } else {
                        target.ActualText(
                            drawingGroup.ActualText,
                            (visual.X + drawingGroup.ActualTextAnchorX * scaleX) * PointsPerCssPixel,
                            (visual.Y + drawingGroup.ActualTextAnchorY * scaleY) * PointsPerCssPixel,
                            addGroupPaint);
                    }
                    continue;
                }
                if (element is OfficeDrawingEffectGroup effectGroup) {
                    FlushShapes();
                    OfficeTransform pageTransform = pageToDrawing
                        .Then(effectGroup.Transform)
                        .Then(drawingToPage);
                    OfficeDrawing nestedDrawing = effectGroup.Drawing;
                    target.Effect(pageTransform, effectGroup.Opacity, nested => AddElements(nested, nestedDrawing.Elements));
                    continue;
                }
                if (element is OfficeDrawingTilingPattern tilingPattern) {
                    FlushShapes();
                    OfficeDrawing tileDrawing = tilingPattern.Tile;
                    OfficeImagePlacement area = tilingPattern.Area;
                    double clipX = (visual.X + (area.X * scaleX)) * PointsPerCssPixel;
                    double clipY = (visual.Y + (area.Y * scaleY)) * PointsPerCssPixel;
                    double clipWidth = area.Width * scaleX * PointsPerCssPixel;
                    double clipHeight = area.Height * scaleY * PointsPerCssPixel;
                    target.Clip(clipX, clipY, clipWidth, clipHeight, clipped => {
                        foreach (OfficeTransform tileTransform in tilingPattern.GetTileTransforms()) {
                            cancellationToken.ThrowIfCancellationRequested();
                            OfficeTransform pageTransform = pageToDrawing
                                .Then(tileTransform)
                                .Then(drawingToPage);
                            clipped.Effect(pageTransform, tilingPattern.Opacity,
                                nested => AddElements(nested, tileDrawing.Elements));
                        }
                    });
                    continue;
                }
                if (element is OfficeDrawingImage image) {
                    FlushShapes();
                    PdfCore.PdfCanvasImageResource? imageResource = GetSharedPdfImageResource(
                        image.EncodedBytes, image.ContentType);
                    if (imageResource == null) continue;
                    PdfCore.PdfImageStyle? imageStyle = image.Projection.HasCrop
                        ? new PdfCore.PdfImageStyle {
                            SourceCrop = new PdfCore.PdfImageSourceCrop(
                                image.Projection.SourceCrop.Left,
                                image.Projection.SourceCrop.Top,
                                image.Projection.SourceCrop.Right,
                                image.Projection.SourceCrop.Bottom)
                        }
                        : null;
                    Action<PdfCore.PdfPageCanvas> addImage = imageTarget => imageTarget.ImageShared(
                        imageResource,
                        (visual.X + image.Projection.X * scaleX) * PointsPerCssPixel,
                        (visual.Y + image.Projection.Y * scaleY) * PointsPerCssPixel,
                        image.Projection.Width * scaleX * PointsPerCssPixel,
                        image.Projection.Height * scaleY * PointsPerCssPixel,
                        imageStyle,
                        alternativeText: string.IsNullOrWhiteSpace(image.AlternativeText) ? null : image.AlternativeText);
                    OfficeTransform imageTransform = image.Projection.CreateFrameTransform().CreateDestinationTransform();
                    if (imageTransform == OfficeTransform.Identity && image.Opacity >= 1D) {
                        addImage(target);
                    } else {
                        OfficeTransform pageImageTransform = pageToDrawing.Then(imageTransform).Then(drawingToPage);
                        target.Effect(pageImageTransform, image.Opacity, addImage);
                    }
                    continue;
                }
                if (element is OfficeDrawingLink link) {
                    FlushShapes();
                    double linkX = (visual.X + link.X * scaleX) * PointsPerCssPixel;
                    double linkY = (visual.Y + link.Y * scaleY) * PointsPerCssPixel;
                    double linkWidth = link.Width * scaleX * PointsPerCssPixel;
                    double linkHeight = link.Height * scaleY * PointsPerCssPixel;
                    if (IsFragmentLink(link.Uri)) {
                        target.LinkToNamedDestination(
                            MapNamedDestination(DecodeFragmentName(link.Uri.Substring(1))),
                            linkX,
                            linkY,
                            linkWidth,
                            linkHeight,
                            link.AlternativeText);
                    } else {
                        OfficeShape linkArea = OfficeShape.Rectangle(linkWidth, linkHeight);
                        linkArea.FillColor = null;
                        linkArea.StrokeColor = null;
                        target.Shape(linkArea, linkX, linkY, linkUri: link.Uri, linkContents: link.AlternativeText);
                    }
                    continue;
                }
                if (element is not OfficeDrawingText text || text.Text.Length == 0) continue;
                FlushShapes();
                double textY = visual.Y + text.Y * scaleY;
                var projectedText = new HtmlRenderText(
                    text.Text,
                    visual.X + text.X * scaleX,
                    textY,
                    text.Width * scaleX,
                    text.Height * scaleY,
                    text.Font.WithSize(text.Font.Size * scaleY),
                    text.Color ?? OfficeColor.Black,
                    text.Alignment,
                    (text.LineHeight ?? text.Font.Size * 1.2D) * scaleY,
                    paintOrder: 0,
                    linkUri: drawingLinkUri,
                    source: visual.Source,
                    semanticRole: "span",
                    layoutY: textY,
                    semanticNodeId: null,
                    textAdvanceWidth: text.TextAdvanceWidth * scaleX,
                    underlineStyle: text.UnderlineStyle,
                    strikethroughStyle: text.StrikethroughStyle,
                    baseline: text.Baseline,
                    baselineLevel: text.BaselineLevel,
                    baselineScale: text.BaselineScale,
                    baselineOffset: text.BaselineOffset * scaleY,
                    decorationColor: text.DecorationColor,
                    featureSettings: text.FeatureSettings,
                    fontPalette: text.FontPalette);
                AddText(target, projectedText, webFonts, conversionReport,
                    visual.X + visual.Width, asSpan: true, logicalTextOwned: false, cancellationToken);
            }
            FlushShapes();
        }

        if (source.Elements.Count == 0) return;
        if (string.IsNullOrWhiteSpace(visual.AlternativeText)) {
            AddElements(canvas, source.Elements);
        } else {
            canvas.Figure(visual.AlternativeText!, figure => AddElements(figure, source.Elements));
        }
        if (fragmentLink) {
            canvas.LinkToNamedDestination(
                MapNamedDestination(visual.LinkUri!.Substring(1)),
                visual.X * PointsPerCssPixel,
                visual.Y * PointsPerCssPixel,
                visual.Width * PointsPerCssPixel,
                visual.Height * PointsPerCssPixel,
                visual.Source);
        }
    }

    private readonly struct ClipBounds {
        internal ClipBounds(double left, double top, double right, double bottom, bool allowsInteractiveWidgets = true, bool constrainToSurface = true) {
            Left = left;
            Top = top;
            Right = right;
            Bottom = bottom;
            AllowsInteractiveWidgets = allowsInteractiveWidgets;
            ConstrainToSurface = constrainToSurface;
        }

        private double Left { get; }
        private double Top { get; }
        private double Right { get; }
        private double Bottom { get; }
        internal bool AllowsInteractiveWidgets { get; }
        internal bool ConstrainToSurface { get; }
        internal static ClipBounds TransformedCoordinateSpace { get; } = new(
            double.NegativeInfinity,
            double.NegativeInfinity,
            double.PositiveInfinity,
            double.PositiveInfinity,
            allowsInteractiveWidgets: false,
            constrainToSurface: false);

        internal bool Contains(HtmlRenderVisual visual) {
            double right = visual.X + visual.Width;
            double bottom = visual.Y + visual.Height;
            return visual.X >= Left - 0.0001D && visual.Y >= Top - 0.0001D
                && right <= Right + 0.0001D && bottom <= Bottom + 0.0001D;
        }

        internal static ClipBounds Intersect(ClipBounds? active, ClipBounds next) => !active.HasValue
            ? next
            : new ClipBounds(
                Math.Max(active.Value.Left, next.Left),
                Math.Max(active.Value.Top, next.Top),
                Math.Min(active.Value.Right, next.Right),
                Math.Min(active.Value.Bottom, next.Bottom),
                active.Value.AllowsInteractiveWidgets && next.AllowsInteractiveWidgets,
                active.Value.ConstrainToSurface && next.ConstrainToSurface);
    }

    private static void AddImagePattern(PdfCore.PdfPageCanvas canvas, HtmlRenderImagePattern visual, CancellationToken cancellationToken) {
        PdfCore.PdfCanvasImageResource? imageResource = GetSharedPdfImageResource(
            visual.EncodedBytes, visual.ContentType);
        if (imageResource == null) return;
        OfficeImagePatternLayout pattern = visual.Pattern.Scale(PointsPerCssPixel);
        OfficeImagePlacement area = pattern.Area;
        canvas.Clip(area.X, area.Y, area.Width, area.Height, clipped => {
            foreach (OfficeImagePlacement tile in pattern.GetTilePlacements(visual.MaximumTileCount)) {
                cancellationToken.ThrowIfCancellationRequested();
                clipped.ImageShared(imageResource, tile.X, tile.Y, tile.Width, tile.Height);
            }
        });
    }

    private static PdfCore.PdfCanvasImageResource? GetSharedPdfImageResource(byte[] encodedBytes,
        string contentType) {
        return PdfImageResources.GetValue(encodedBytes, static _ => new CachedPdfImageResources())
            .GetOrCreate(encodedBytes, contentType);
    }

    private sealed class CachedPdfImageResources {
        private readonly Dictionary<string, PdfCore.PdfCanvasImageResource?> _resources =
            new(StringComparer.OrdinalIgnoreCase);

        internal PdfCore.PdfCanvasImageResource? GetOrCreate(byte[] encodedBytes, string contentType) {
            lock (_resources) {
                if (_resources.TryGetValue(contentType, out PdfCore.PdfCanvasImageResource? resource)) {
                    return resource;
                }
                if (TryPreparePdfImageBytes(encodedBytes, contentType, out byte[] prepared)) {
                    resource = PdfCore.PdfCanvasImageResource.Create(prepared);
                }
                _resources.Add(contentType, resource);
                return resource;
            }
        }
    }

    private static bool TryPreparePdfImageBytes(byte[] bytes, string contentType, out byte[] pdfBytes) {
        OfficeImageFormat format = OfficeImageInfo.FromMimeType(contentType);
        string extension = OfficeImageInfo.GetDefaultExtension(format);
        if (OfficeImageReader.TryIdentify(bytes, extension, out OfficeImageInfo identified)) {
            format = identified.Format;
        }

        if (format == OfficeImageFormat.Png || format == OfficeImageFormat.Jpeg) {
            pdfBytes = bytes;
            return true;
        }

        return OfficeImagePngConverter.TryConvertToPng(bytes, out pdfBytes);
    }

    private static PdfCore.PdfAlign MapAlignment(OfficeTextAlignment alignment) {
        if (alignment == OfficeTextAlignment.Center) return PdfCore.PdfAlign.Center;
        if (alignment == OfficeTextAlignment.Right) return PdfCore.PdfAlign.Right;
        if (alignment == OfficeTextAlignment.Justify) return PdfCore.PdfAlign.Justify;
        return PdfCore.PdfAlign.Left;
    }
}
