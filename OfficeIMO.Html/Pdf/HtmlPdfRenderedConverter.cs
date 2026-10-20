using OfficeIMO.Drawing;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

internal static partial class HtmlPdfRenderedConverter {
    private const double PointsPerCssPixel = 72D / HtmlRenderOptions.CssPixelsPerInch;
    private const int MaximumSystemFontFamilyCandidates = 512;
    private const int MaximumLoadedSystemFontFamilies = 32;
    private const int MaximumCssFontFamilyCandidates = 32;

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

    // Image encoders use exactly the PDF font measurement and print-fitting policy.
    // Keep the original request axes and selection; only the encoder changes here.
    /// <summary>Applies PDF font and resource policy to image and display-list requests.</summary>
    internal static HtmlRenderOptions PrepareNonPdfOptions(HtmlRenderRequest request) {
        HtmlRenderOptions source = request.Options;
        var options = new HtmlToPdfOptions(source);
        bool preserveHtmlResourcePolicy = source is not HtmlToPdfOptions;
        if (preserveHtmlResourcePolicy) {
            // Retained EPUB, email and site-bundle resolvers already enforce their
            // container boundary. Preparing PDF fonts must not replace that policy.
            options.UrlPolicy = source.UrlPolicy.Clone();
            options.ResourceUrlPolicy = source.GetResourceUrlPolicy().Clone();
        }
        return PrepareRequest(request.WithOptions(options), options, allowNonPdfEncoder: true,
            preserveHtmlResourcePolicy: preserveHtmlResourcePolicy).Request.ResolveOptions();
    }

    private static PreparedPdfRequest PrepareRequest(HtmlRenderRequest request, HtmlToPdfOptions options,
        bool allowNonPdfEncoder = false, bool preserveHtmlResourcePolicy = false) {
        if (request == null) throw new ArgumentNullException(nameof(request));
        if (options == null) throw new ArgumentNullException(nameof(options));
        if (request.Encoder != HtmlRenderEncoder.Pdf && !allowNonPdfEncoder) {
            throw new ArgumentException("HTML-to-PDF conversion requires the Pdf render encoder.", nameof(request));
        }
        HtmlToPdfOptions renderOptions = new HtmlToPdfOptions(request.Options);
        CopyAdapterOptions(options, renderOptions);
        ApplyPrintLayoutWidth(request, renderOptions);
        renderOptions.AutoFitWidePrintRoot = renderOptions.AutoFitWidePrintContent
            && renderOptions.PrintLayoutWidthCssPixels == null
            && request.Profile == HtmlRenderIntentProfile.PrintPaged;
        PdfCore.PdfOptions measurementOptions = options.PdfOptions.Clone();
        measurementOptions.TextShapingLanguageOverride = options.TextShapingLanguage;
        measurementOptions.SetTextShapingMode(options.TextShapingMode).SetTextShapingProvider(options.TextShapingProvider);
        if (options.FontFamily != null) measurementOptions.RegisterFontFamily(PdfCore.PdfStandardFont.Helvetica, options.FontFamily);
        if (renderOptions.LibrarySelectedDefaultFontFamily != null) {
            measurementOptions.RegisterFontFamily(PdfCore.PdfStandardFont.Helvetica,
                renderOptions.LibrarySelectedDefaultFontFamily);
        }
        renderOptions.AllowSystemFontFallback = renderOptions.AllowSystemFontFallback
            && options.ResourcePolicy.AllowSystemFontEmbedding;
        renderOptions.MathOnlySystemFontFallback = renderOptions.MathOnlySystemFontFallback
            || !options.ResourcePolicy.AllowDocumentFontEmbedding;
        renderOptions.FallbackTextMeasurement = CreateFallbackTextMeasurement(options, measurementOptions);
        renderOptions.FallbackTextFaceMetrics = CreateFallbackTextFaceMetrics(options, measurementOptions);
        if (preserveHtmlResourcePolicy) return new PreparedPdfRequest(request.WithOptions(renderOptions), renderOptions);
        // Prepared requests retain archive callbacks across option snapshots. Recheck
        // current permissions for synchronous reads just as for asynchronous reads below.
        if (!options.ResourcePolicy.AllowEmbeddedPackageResources)
            renderOptions.SynchronousResourceResolver = null;
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
        if (hostResolver != null && hostResolver == options.EmbeddedPackageConfiguredResourceResolver) {
            hostResolver = options.EmbeddedPackageHostResourceResolver;
        }
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
        return CreatePdf(rendered.Document, (HtmlToPdfOptions)resolved, cancellationToken).WithRenderResult(rendered);
    }

    private static async Task<HtmlPdfRenderResult> ConvertCoreAsync(
        HtmlConversionDocument document,
        PreparedPdfRequest prepared,
        HtmlRenderOptions resolved,
        CancellationToken cancellationToken) {
        HtmlRenderResult rendered = await HtmlRenderEngine.ExecuteCoreAsync(
            document, prepared.Request, resolved, cancellationToken).ConfigureAwait(false);
        return CreatePdf(rendered.Document, (HtmlToPdfOptions)resolved, cancellationToken).WithRenderResult(rendered);
    }

    private static void CopyAdapterOptions(HtmlToPdfOptions source, HtmlToPdfOptions target) {
        target.TextFallbacks = source.TextFallbacks;
        target.TextShapingMode = source.TextShapingMode;
        target.FontFamily = source.FontFamily;
        target.LibrarySelectedDefaultFontFamily = source.LibrarySelectedDefaultFontFamily;
        target.InteractiveFormControls = source.InteractiveFormControls;
        target.MathMlSourceModificationDate = source.MathMlSourceModificationDate;
        target.PrintLayoutWidthCssPixels = source.PrintLayoutWidthCssPixels;
        target.AutoFitWidePrintContent = source.AutoFitWidePrintContent;
        target.MaxOutlinedTextCharactersPerRun = source.MaxOutlinedTextCharactersPerRun;
        target.MaxOutlinedTextPathCommands = source.MaxOutlinedTextPathCommands;
        target.PdfOptions = source.PdfOptions.Clone();
        target.TextShapingProvider = source.TextShapingProvider;
        target.ResourcePolicy = source.ResourcePolicy.Clone();
        target.EmbeddedPackageResourceResolver = source.EmbeddedPackageResourceResolver;
        target.EmbeddedPackageHostResourceUrlPolicy = source.EmbeddedPackageHostResourceUrlPolicy?.Clone();
        target.EmbeddedPackageConfiguredResourceResolver = source.EmbeddedPackageConfiguredResourceResolver;
        target.EmbeddedPackageHostResourceResolver = source.EmbeddedPackageHostResourceResolver;
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
        documentOptions.TextShapingLanguageOverride = options.TextShapingLanguage;
        documentOptions.UseContentStreamCompressionByDefault();
        PdfCore.PdfDocument pdf = PdfCore.PdfDocument.Create(documentOptions);
        var mathMlFiles = new HtmlPdfMathMlFiles(options.MathMlSourceModificationDate);
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

        var imageResources = new PdfImageResourceCache(options, conversionReport, cancellationToken);
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
        webFonts.AllowInstalledFontFaces = options.ResourcePolicy.AllowSystemFontEmbedding
            && options.ResourcePolicy.AllowDocumentFontEmbedding;
        var activeWebFontFamilies = new HashSet<string>(
            webFonts.Slots.Keys,
            StringComparer.OrdinalIgnoreCase);
        PdfCore.PdfTextFallbackFeatures activeTextFallbacks = ResolveTextFallbackFeatures(rendered, options.TextFallbacks);
        if (options.ResourcePolicy.AllowSystemFontEmbedding) {
            if (options.ResourcePolicy.AllowDocumentFontEmbedding) {
                RegisterUsedSystemFontFamilies(pdf, rendered, activeWebFontFamilies,
                    reservedFontSlots, cancellationToken);
            } else if ((activeTextFallbacks != PdfCore.PdfTextFallbackFeatures.None
                    || options.LibrarySelectedDefaultFontFamily != null) && options.FontFamily == null) {
                RegisterLibrarySelectedDefaultSystemFontFamily(
                    pdf,
                    rendered,
                    activeWebFontFamilies,
                    reservedFontSlots,
                    options.LibrarySelectedDefaultFontFamily,
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
                    // The retained HTML surface already clips its paint at the page edge.
                    // Keep that source boundary outside print and element transforms.
                    canvas.SourceClip(0D, 0D, pageWidth, pageHeight, sourcePaint => {
                        if (printLayoutScale == 1D) {
                            AddPageVisuals(sourcePaint, renderedPage, webFonts, imageResources, conversionReport, options.InteractiveFormControls, cancellationToken, mathMlFiles);
                        } else {
                            sourcePaint.Effect(OfficeTransform.Scale(printLayoutScale, printLayoutScale), 1D,
                                content => AddPageVisuals(content, renderedPage, webFonts, imageResources, conversionReport, options.InteractiveFormControls, cancellationToken, mathMlFiles));
                        }
                    });
                    AddPageOutlines(canvas, headingsByPage[renderedPage.PageNumber], headingDocumentOrder, printLayoutScale, cancellationToken);
                });
            });
        }

        IReadOnlyList<PdfCore.PdfTextEncodingDiagnostic> encodingDiagnostics = pdf.AnalyzeTextEncoding(cancellationToken);
        if (encodingDiagnostics.Count > 0) {
            throw new HtmlConversionException(diagnostics.Concat(encodingDiagnostics.Select(
                diagnostic => new HtmlDiagnostic(
                    "OfficeIMO.Html.Pdf",
                    diagnostic.Code,
                    diagnostic.Message,
                    HtmlDiagnosticSeverity.Error,
                    diagnostic.Source,
                    detail: diagnostic.Location))));
        }

        if (options.FidelityPolicy == HtmlRenderFidelityPolicy.RequireNoLoss
            && diagnostics.Any(static diagnostic =>
                diagnostic.LossKind != OfficeConversionLossKind.None
                || diagnostic.Severity == HtmlDiagnosticSeverity.Warning
                || diagnostic.Severity == HtmlDiagnosticSeverity.Error)) {
            throw new HtmlConversionException(diagnostics);
        }
        if (options.FidelityPolicy == HtmlRenderFidelityPolicy.RequireNoLoss) {
            // Font providers can decline only when the PDF writer paints a run.
            // Qualify that stage before returning a strict model or image scene,
            // as well as before exposing bytes through the direct PDF routes.
            _ = pdf.ToBytes(cancellationToken);
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

    private static void AddPageVisuals(PdfCore.PdfPageCanvas canvas, HtmlRenderPage page, RegisteredWebFonts webFonts, PdfImageResourceCache imageResources, PdfCore.PdfConversionReport conversionReport, bool interactiveFormControls, CancellationToken cancellationToken, HtmlPdfMathMlFiles mathMlFiles) {
        HtmlPdfPagePaintContext pagePaint = HtmlPdfPagePaintContext.Create(page, mathMlFiles);
        foreach (HtmlRenderVisual visual in page.Scene.OrderBy(item => item.PaintOrder)) {
            cancellationToken.ThrowIfCancellationRequested();
            AddVisual(canvas, visual, webFonts, imageResources, conversionReport, page.Width, page.Height, interactiveFormControls, cancellationToken, pagePaint: pagePaint);
        }
    }

    private static void AddVisual(
        PdfCore.PdfPageCanvas canvas,
        HtmlRenderVisual visual,
        RegisteredWebFonts webFonts, PdfImageResourceCache imageResources,
        PdfCore.PdfConversionReport conversionReport,
        double surfaceWidth,
        double surfaceHeight,
        bool interactiveFormControls,
        CancellationToken cancellationToken,
        bool textAsSpan = false,
        ClipBounds? activeClip = null,
        bool logicalTextOwned = false,
        HtmlPdfPagePaintContext? pagePaint = null) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!activeClip.HasValue && !IsVisualContainer(visual) && ExceedsSurface(visual, surfaceWidth, surfaceHeight)) {
            double right = visual.X + visual.Width;
            double bottom = visual.Y + visual.Height;
            if (right <= 0D || bottom <= 0D || visual.X >= surfaceWidth || visual.Y >= surfaceHeight) return;
            var pageClip = new ClipBounds(0D, 0D, surfaceWidth, surfaceHeight);
            canvas.SourceClip(0D, 0D, surfaceWidth * PointsPerCssPixel, surfaceHeight * PointsPerCssPixel, clipped =>
                AddVisual(clipped, visual, webFonts, imageResources, conversionReport, surfaceWidth, surfaceHeight,
                    interactiveFormControls, cancellationToken, textAsSpan, pageClip, logicalTextOwned, pagePaint));
            return;
        }
        if (!textAsSpan && !logicalTextOwned && visual.SemanticSourceOrder.HasValue
            && (visual is HtmlRenderText { IsPaintOnly: false } || visual is HtmlRenderLogicalTextGroup)) {
            var anonymousContent = new PdfCore.PdfPageCanvas(allowOutOfPageCoordinates: true);
            AddVisual(anonymousContent, visual, webFonts, imageResources, conversionReport, surfaceWidth, surfaceHeight,
                interactiveFormControls, cancellationToken, textAsSpan: true, activeClip, logicalTextOwned, pagePaint);
            if (HasCanvasContent(anonymousContent.Items))
                canvas.Structure(PdfCore.PdfCanvasStructureRole.Paragraph, nested => nested.AddItems(anonymousContent.Items),
                    new PdfCore.PdfCanvasStructureOptions { LogicalOrder = visual.SemanticSourceOrder });
            else canvas.AddItems(anonymousContent.Items);
            return;
        }
        if (visual is HtmlRenderFormField formField) {
            bool fullyContained = !activeClip.HasValue || activeClip.Value.AllowsInteractiveWidgets && activeClip.Value.Contains(formField);
            AddFormField(canvas, formField, webFonts, imageResources, conversionReport, surfaceWidth, surfaceHeight, interactiveFormControls && fullyContained, cancellationToken, textAsSpan, activeClip, logicalTextOwned, pagePaint);
        } else if (visual is HtmlRenderAnchorFragment anchorFragment) {
            if (pagePaint?.IsActive(anchorFragment) == true) AddAnchorFragment(canvas, anchorFragment);
        } else if (visual is HtmlRenderShape shape) {
            AddShape(canvas, shape, imageResources, conversionReport, cancellationToken, pagePaint?.Covers(shape) == true);
        } else if (visual is HtmlRenderText text) {
            AddText(canvas, text, webFonts, conversionReport, surfaceWidth, textAsSpan, logicalTextOwned, cancellationToken,
                suppressLink: pagePaint?.Covers(text) == true,
                constrainToSurface: !activeClip.HasValue || activeClip.Value.ConstrainToSurface,
                logicalClip: activeClip ?? new ClipBounds(0D, 0D, surfaceWidth, surfaceHeight));
        } else if (visual is HtmlRenderNamedDestination destination) {
            canvas.NamedDestination(
                MapNamedDestination(destination.Name),
                destination.X * PointsPerCssPixel,
                destination.Y * PointsPerCssPixel);
        } else if (visual is HtmlRenderImage image) {
            try {
                if (!AddImage(canvas, image, imageResources, pagePaint?.Covers(image) == true)) {
                    ReportImagePayloadOmitted(conversionReport, image,
                        "The image payload could not be prepared as a PDF raster resource.");
                }
            } catch (System.NotSupportedException exception) {
                ReportImagePayloadOmitted(conversionReport, image, exception.Message);
            }
        } else if (visual is HtmlRenderDrawing drawing) {
            AddDrawing(canvas, drawing, webFonts, imageResources, conversionReport, cancellationToken, pagePaint?.Covers(drawing) == true);
        } else if (visual is HtmlRenderImagePattern imagePattern) {
            AddImagePattern(canvas, imagePattern, imageResources, cancellationToken);
        } else if (visual is HtmlRenderClipGroup group) {
            AddClipGroup(canvas, group, webFonts, imageResources, conversionReport, surfaceWidth, surfaceHeight, interactiveFormControls, cancellationToken, textAsSpan, activeClip, logicalTextOwned, pagePaint);
        } else if (visual is HtmlRenderPathClipGroup pathClipGroup) {
            AddPathClipGroup(canvas, pathClipGroup, webFonts, imageResources, conversionReport, surfaceWidth, surfaceHeight, interactiveFormControls, cancellationToken, textAsSpan, activeClip, logicalTextOwned, pagePaint);
        } else if (visual is HtmlRenderEffectGroup effectGroup) {
            AddEffectGroup(canvas, effectGroup, webFonts, imageResources, conversionReport, surfaceWidth, surfaceHeight, interactiveFormControls, cancellationToken, textAsSpan, activeClip, logicalTextOwned, pagePaint);
        } else if (visual is HtmlRenderLayoutRegion layoutRegion) {
            foreach (HtmlRenderVisual child in layoutRegion.Visuals.OrderBy(item => item.PaintOrder)) {
                AddVisual(canvas, child, webFonts, imageResources, conversionReport, surfaceWidth, surfaceHeight, interactiveFormControls, cancellationToken, textAsSpan, activeClip, logicalTextOwned, pagePaint);
            }
        } else if (visual is HtmlRenderSemanticGroup semanticGroup) {
            AddSemanticGroup(canvas, semanticGroup, webFonts, imageResources, conversionReport, surfaceWidth, surfaceHeight, interactiveFormControls, cancellationToken, textAsSpan, activeClip, logicalTextOwned, pagePaint);
        } else if (visual is HtmlRenderLogicalTextGroup logicalTextGroup) {
            AddLogicalTextGroup(canvas, logicalTextGroup, webFonts, imageResources, conversionReport, surfaceWidth, surfaceHeight, interactiveFormControls, cancellationToken, textAsSpan, activeClip, logicalTextOwned, pagePaint);
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

    private static void AddFormField(PdfCore.PdfPageCanvas canvas, HtmlRenderFormField field, RegisteredWebFonts webFonts, PdfImageResourceCache imageResources, PdfCore.PdfConversionReport conversionReport, double surfaceWidth, double surfaceHeight, bool interactiveFormControls, CancellationToken cancellationToken, bool textAsSpan, ClipBounds? activeClip, bool logicalTextOwned, HtmlPdfPagePaintContext? pagePaint) {
        bool hasInvalidPdfButtonValue = (field.FieldKind == HtmlRenderFormFieldKind.CheckBox || field.FieldKind == HtmlRenderFormFieldKind.RadioButton)
            && (string.IsNullOrWhiteSpace(field.Value) || string.IsNullOrWhiteSpace(field.RadioOption));
        if (!interactiveFormControls
            || string.IsNullOrWhiteSpace(field.Name)
            || hasInvalidPdfButtonValue
            || field.FieldKind == HtmlRenderFormFieldKind.Choice && field.Options.Count == 0) {
            foreach (HtmlRenderVisual child in field.Visuals.OrderBy(item => item.PaintOrder)) {
                AddVisual(canvas, child, webFonts, imageResources, conversionReport, surfaceWidth, surfaceHeight, interactiveFormControls, cancellationToken, textAsSpan, activeClip, logicalTextOwned, pagePaint);
            }
            return;
        }

        var style = new PdfCore.PdfFormFieldStyle {
            LogicalOrder = field.SemanticSourceOrder,
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
                canvas.SearchableText(field.Value, x, y + Math.Min(height, fontSize), logicalOrder: field.SemanticSourceOrder);
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
                canvas.SearchableText(searchableValue, x, y + Math.Min(height, fontSize), logicalOrder: field.SemanticSourceOrder);
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
        HtmlRenderSemanticGroup.IsTextContentRole(role);

    private static PdfCore.PdfCanvasTableHeaderScope? MapTableHeaderScope(HtmlRenderTableHeaderScope? scope) {
        if (scope == HtmlRenderTableHeaderScope.Row) return PdfCore.PdfCanvasTableHeaderScope.Row;
        if (scope == HtmlRenderTableHeaderScope.Column) return PdfCore.PdfCanvasTableHeaderScope.Column;
        if (scope == HtmlRenderTableHeaderScope.Both) return PdfCore.PdfCanvasTableHeaderScope.Both;
        return null;
    }

    private static void AddEffectGroup(
        PdfCore.PdfPageCanvas canvas,
        HtmlRenderEffectGroup group,
        RegisteredWebFonts webFonts, PdfImageResourceCache imageResources,
        PdfCore.PdfConversionReport conversionReport,
        double surfaceWidth,
        double surfaceHeight,
        bool interactiveFormControls,
        CancellationToken cancellationToken,
        bool textAsSpan,
        ClipBounds? activeClip,
        bool logicalTextOwned,
        HtmlPdfPagePaintContext? pagePaint) {
        OfficeTransform transform = group.Transform;
        var scaled = new OfficeTransform(
            transform.M11,
            transform.M12,
            transform.M21,
            transform.M22,
            transform.OffsetX * PointsPerCssPixel,
            transform.OffsetY * PointsPerCssPixel);
        ClipBounds currentWindow = activeClip.HasValue && !activeClip.Value.ConstrainToSurface
            ? activeClip.Value
            : ClipBounds.Intersect(activeClip, new ClipBounds(0D, 0D, surfaceWidth, surfaceHeight));
        ClipBounds logicalWindow = currentWindow.InEffectCoordinateSpace(transform);
        canvas.Effect(scaled, group.Opacity, nested => {
            foreach (HtmlRenderVisual child in group.Visuals.OrderBy(item => item.PaintOrder)) {
                cancellationToken.ThrowIfCancellationRequested();
                AddVisual(nested, child, webFonts, imageResources, conversionReport, surfaceWidth, surfaceHeight, interactiveFormControls, cancellationToken, textAsSpan, logicalWindow, logicalTextOwned, pagePaint);
            }
        });
    }

    private static void AddClipGroup(
        PdfCore.PdfPageCanvas canvas,
        HtmlRenderClipGroup group,
        RegisteredWebFonts webFonts, PdfImageResourceCache imageResources,
        PdfCore.PdfConversionReport conversionReport,
        double surfaceWidth,
        double surfaceHeight,
        bool interactiveFormControls,
        CancellationToken cancellationToken,
        bool textAsSpan,
        ClipBounds? activeClip,
        bool logicalTextOwned,
        HtmlPdfPagePaintContext? pagePaint) {
        bool constrainToSurface = !activeClip.HasValue || activeClip.Value.ConstrainToSurface;
        double left = group.ClipHorizontal ? (constrainToSurface ? Math.Max(0D, group.ClipX) : group.ClipX) : (constrainToSurface ? 0D : group.X);
        double top = group.ClipVertical ? (constrainToSurface ? Math.Max(0D, group.ClipY) : group.ClipY) : (constrainToSurface ? 0D : group.Y);
        double right = group.ClipHorizontal ? (constrainToSurface ? Math.Min(surfaceWidth, group.ClipX + group.ClipWidth) : group.ClipX + group.ClipWidth) : (constrainToSurface ? surfaceWidth : group.X + group.Width);
        double bottom = group.ClipVertical ? (constrainToSurface ? Math.Min(surfaceHeight, group.ClipY + group.ClipHeight) : group.ClipY + group.ClipHeight) : (constrainToSurface ? surfaceHeight : group.Y + group.Height);
        if (right <= left + 0.0001D || bottom <= top + 0.0001D) {
            // A paint clip does not remove the source's navigation targets.
            AddClippedNamedDestinations(canvas, group.Visuals, OfficeTransform.Identity, cancellationToken);
            return;
        }
        ClipBounds clip = ClipBounds.Intersect(activeClip, new ClipBounds(left, top, right, bottom, constrainToSurface: constrainToSurface));
        canvas.SourceClip(
            left * PointsPerCssPixel,
            top * PointsPerCssPixel,
            (right - left) * PointsPerCssPixel,
            (bottom - top) * PointsPerCssPixel,
            clipped => {
                foreach (HtmlRenderVisual child in group.Visuals.OrderBy(item => item.PaintOrder)) {
                    cancellationToken.ThrowIfCancellationRequested();
                    AddVisual(clipped, child, webFonts, imageResources, conversionReport, surfaceWidth, surfaceHeight, interactiveFormControls, cancellationToken, textAsSpan, clip, logicalTextOwned, pagePaint);
                }
            });
    }

    private static void AddPathClipGroup(
        PdfCore.PdfPageCanvas canvas,
        HtmlRenderPathClipGroup group,
        RegisteredWebFonts webFonts, PdfImageResourceCache imageResources,
        PdfCore.PdfConversionReport conversionReport,
        double surfaceWidth,
        double surfaceHeight,
        bool interactiveFormControls,
        CancellationToken cancellationToken,
        bool textAsSpan,
        ClipBounds? activeClip,
        bool logicalTextOwned,
        HtmlPdfPagePaintContext? pagePaint) {
        ClipBounds clip = ClipBounds.Intersect(activeClip, new ClipBounds(
            group.X,
            group.Y,
            group.X + group.Width,
            group.Y + group.Height,
            allowsInteractiveWidgets: false));
        double clipX = group.ClipX * PointsPerCssPixel;
        double clipY = group.ClipY * PointsPerCssPixel;
        OfficeClipPath clipPath = group.ClipPath.Scale(PointsPerCssPixel, PointsPerCssPixel);
        Action<PdfCore.PdfPageCanvas> addClip = target => target.SourceClip(
            clipX,
            clipY,
            clipPath,
            clipped => {
                foreach (HtmlRenderVisual child in group.Visuals.OrderBy(item => item.PaintOrder)) {
                    cancellationToken.ThrowIfCancellationRequested();
                    AddVisual(clipped, child, webFonts, imageResources, conversionReport, surfaceWidth, surfaceHeight, interactiveFormControls, cancellationToken, textAsSpan, clip, logicalTextOwned, pagePaint);
                }
            });
        if (clipX < 0D || clipY < 0D) {
            canvas.Effect(OfficeTransform.Identity, 1D, addClip);
        } else if (!activeClip.HasValue
                   && (clipX + clipPath.Width > surfaceWidth * PointsPerCssPixel
                       || clipY + clipPath.Height > surfaceHeight * PointsPerCssPixel)) {
            canvas.SourceClip(0D, 0D, surfaceWidth * PointsPerCssPixel, surfaceHeight * PointsPerCssPixel, addClip);
        } else {
            addClip(canvas);
        }
    }

    private static void AddAnchorFragment(PdfCore.PdfPageCanvas canvas, HtmlRenderAnchorFragment fragment) {
        double x = fragment.X * PointsPerCssPixel;
        double y = fragment.Y * PointsPerCssPixel;
        double width = fragment.Width * PointsPerCssPixel;
        double height = fragment.Height * PointsPerCssPixel;
        if (IsFragmentLink(fragment.LinkUri)) {
            canvas.LinkToNamedDestination(MapNamedDestination(fragment.LinkUri!.Substring(1)),
                x, y, width, height, fragment.LinkContents);
            return;
        }
        OfficeShape area = OfficeShape.Rectangle(width, height);
        area.FillColor = null;
        area.StrokeColor = null;
        canvas.Shape(area, x, y, style: new PdfCore.PdfDrawingStyle { Decorative = true }, linkUri: fragment.LinkUri, linkContents: fragment.LinkContents);
    }

    private static void AddShape(
        PdfCore.PdfPageCanvas canvas,
        HtmlRenderShape visual,
        PdfImageResourceCache imageResources,
        PdfCore.PdfConversionReport conversionReport,
        CancellationToken cancellationToken,
        bool suppressLink) {
        var drawing = new OfficeDrawing(visual.Width, visual.Height);
        drawing.AddShape(visual.Shape.Clone(), 0D, 0D);
        if (TryAddTranslucentGradient(canvas, visual, drawing, imageResources, conversionReport, cancellationToken, suppressLink)) return;
        double x = visual.X * PointsPerCssPixel;
        double y = visual.Y * PointsPerCssPixel;
        double drawingX = Math.Max(0D, x);
        double drawingY = Math.Max(0D, y);
        OfficeTransform transform = OfficeTransform.Translate(Math.Min(0D, x), Math.Min(0D, y));
        bool fragmentLink = !suppressLink && IsFragmentLink(visual.LinkUri);
        Action<PdfCore.PdfPageCanvas> addDrawing = target => target.Drawing(
            drawing,
            drawingX,
            drawingY,
            visual.Width * PointsPerCssPixel,
            visual.Height * PointsPerCssPixel,
            style: new PdfCore.PdfDrawingStyle { Decorative = true },
            linkUri: suppressLink || fragmentLink ? null : visual.LinkUri,
            linkContents: suppressLink || visual.LinkUri == null || fragmentLink ? null : visual.Source);
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

    private static PdfCore.PdfAlign MapAlignment(OfficeTextAlignment alignment) {
        if (alignment == OfficeTextAlignment.Center) return PdfCore.PdfAlign.Center;
        if (alignment == OfficeTextAlignment.Right) return PdfCore.PdfAlign.Right;
        if (alignment == OfficeTextAlignment.Justify) return PdfCore.PdfAlign.Justify;
        return PdfCore.PdfAlign.Left;
    }
}
