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

    internal static HtmlPdfRenderResult Convert(HtmlConversionDocument document, HtmlPdfSaveOptions options) {
        HtmlRenderOptions renderOptions = ResolveRenderOptions(options);
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(document, renderOptions);
        return CreatePdf(rendered, options, CancellationToken.None);
    }

    internal static async Task<HtmlPdfRenderResult> ConvertAsync(HtmlConversionDocument document, HtmlPdfSaveOptions options, CancellationToken cancellationToken) {
        HtmlRenderOptions renderOptions = ResolveRenderOptions(options);
        HtmlRenderDocument rendered = await HtmlRenderEngine.RenderAsync(document, renderOptions, cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        return CreatePdf(rendered, options, cancellationToken);
    }

    private static HtmlRenderOptions ResolveRenderOptions(HtmlPdfSaveOptions options) {
        HtmlRenderOptions renderOptions = options.ClonePdf();
        renderOptions.Mode = HtmlRenderMode.Paged;
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
        return renderOptions;
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

    private static HtmlPdfRenderResult CreatePdf(HtmlRenderDocument rendered, HtmlPdfSaveOptions options, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        HtmlDiagnosticReport diagnostics = rendered.DiagnosticReport.Clone();

        var conversionReport = new PdfCore.PdfConversionReport();
        PdfCore.PdfOptions documentOptions = options.PdfOptions.Clone();
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
            cancellationToken);
        var activeWebFontFamilies = new HashSet<string>(
            webFonts.Slots.Keys,
            StringComparer.OrdinalIgnoreCase);
        PdfCore.PdfTextFallbackFeatures activeTextFallbacks = ResolveTextFallbackFeatures(rendered, options.TextFallbacks);
        if (activeTextFallbacks != PdfCore.PdfTextFallbackFeatures.None &&
            options.ResourcePolicy.AllowSystemFontEmbedding) {
            if (options.ResourcePolicy.AllowDocumentFontEmbedding) {
                RegisterUsedSystemFontFamilies(pdf, rendered, activeWebFontFamilies, reservedFontSlots, cancellationToken);
            } else if (options.FontFamily == null) {
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
                preserveConfiguredFontSlots: options.FontFamily != null);
        }
        pdf.UseTextShaping(options.TextShapingMode, options.TextShapingProvider);
        var headingDocumentOrder = rendered.Headings
            .Select((heading, index) => new { Heading = heading, Index = index })
            .ToDictionary(item => item.Heading, item => item.Index);
        ILookup<int, HtmlRenderHeading> headingsByPage = rendered.Headings.ToLookup(heading => heading.PageNumber);
        foreach (HtmlRenderPage renderedPage in rendered.Pages) {
            cancellationToken.ThrowIfCancellationRequested();
            double pageWidth = renderedPage.Width * PointsPerCssPixel;
            double pageHeight = renderedPage.Height * PointsPerCssPixel;
            pdf.Page(page => page
                .Size(pageWidth, pageHeight)
                .Margin(0D)
                .Canvas(canvas => {
                    AddPageVisuals(canvas, renderedPage, webFonts, conversionReport, options.InteractiveFormControls, cancellationToken);
                    AddPageOutlines(canvas, headingsByPage[renderedPage.PageNumber], headingDocumentOrder, cancellationToken);
                }));
        }

        cancellationToken.ThrowIfCancellationRequested();
        return new HtmlPdfRenderResult(pdf, diagnostics, conversionReport);
    }

    private static void AddPageOutlines(PdfCore.PdfPageCanvas canvas, IEnumerable<HtmlRenderHeading> headings, IReadOnlyDictionary<HtmlRenderHeading, int> headingDocumentOrder, CancellationToken cancellationToken) {
        foreach (HtmlRenderHeading heading in headings) {
            cancellationToken.ThrowIfCancellationRequested();
            canvas.Outline(heading.Text, heading.Level, heading.Y * PointsPerCssPixel, heading.BookmarkState switch {
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
        ClipBounds? activeClip = null) {
        cancellationToken.ThrowIfCancellationRequested();
        if (visual is HtmlRenderFormField formField) {
            bool fullyContained = !activeClip.HasValue || activeClip.Value.AllowsInteractiveWidgets && activeClip.Value.Contains(formField);
            AddFormField(canvas, formField, webFonts, conversionReport, surfaceWidth, surfaceHeight, interactiveFormControls && fullyContained, cancellationToken, textAsSpan, activeClip);
        } else if (visual is HtmlRenderShape shape) {
            AddShape(canvas, shape, conversionReport, cancellationToken);
        } else if (visual is HtmlRenderText text) {
            AddText(canvas, text, webFonts, surfaceWidth, textAsSpan);
        } else if (visual is HtmlRenderImage image) {
            AddImage(canvas, image);
        } else if (visual is HtmlRenderDrawing drawing) {
            AddDrawing(canvas, drawing, webFonts, conversionReport, cancellationToken);
        } else if (visual is HtmlRenderImagePattern imagePattern) {
            AddImagePattern(canvas, imagePattern, cancellationToken);
        } else if (visual is HtmlRenderClipGroup group) {
            AddClipGroup(canvas, group, webFonts, conversionReport, surfaceWidth, surfaceHeight, interactiveFormControls, cancellationToken, textAsSpan, activeClip);
        } else if (visual is HtmlRenderPathClipGroup pathClipGroup) {
            AddPathClipGroup(canvas, pathClipGroup, webFonts, conversionReport, surfaceWidth, surfaceHeight, interactiveFormControls, cancellationToken, textAsSpan, activeClip);
        } else if (visual is HtmlRenderEffectGroup effectGroup) {
            AddEffectGroup(canvas, effectGroup, webFonts, conversionReport, surfaceWidth, surfaceHeight, interactiveFormControls, cancellationToken, textAsSpan, activeClip);
        } else if (visual is HtmlRenderLayoutRegion layoutRegion) {
            foreach (HtmlRenderVisual child in layoutRegion.Visuals.OrderBy(item => item.PaintOrder)) {
                AddVisual(canvas, child, webFonts, conversionReport, surfaceWidth, surfaceHeight, interactiveFormControls, cancellationToken, textAsSpan, activeClip);
            }
        } else if (visual is HtmlRenderSemanticGroup semanticGroup) {
            AddSemanticGroup(canvas, semanticGroup, webFonts, conversionReport, surfaceWidth, surfaceHeight, interactiveFormControls, cancellationToken, textAsSpan, activeClip);
        } else if (visual is HtmlRenderLogicalTextGroup logicalTextGroup) {
            AddLogicalTextGroup(canvas, logicalTextGroup, webFonts, conversionReport, surfaceWidth, surfaceHeight, interactiveFormControls, cancellationToken, textAsSpan, activeClip);
        }
    }

    private static void AddFormField(PdfCore.PdfPageCanvas canvas, HtmlRenderFormField field, RegisteredWebFonts webFonts, PdfCore.PdfConversionReport conversionReport, double surfaceWidth, double surfaceHeight, bool interactiveFormControls, CancellationToken cancellationToken, bool textAsSpan, ClipBounds? activeClip) {
        bool hasInvalidPdfButtonValue = (field.FieldKind == HtmlRenderFormFieldKind.CheckBox || field.FieldKind == HtmlRenderFormFieldKind.RadioButton)
            && (string.IsNullOrWhiteSpace(field.Value) || string.IsNullOrWhiteSpace(field.RadioOption));
        if (!interactiveFormControls
            || string.IsNullOrWhiteSpace(field.Name)
            || hasInvalidPdfButtonValue
            || field.FieldKind == HtmlRenderFormFieldKind.Choice && field.Options.Count == 0) {
            foreach (HtmlRenderVisual child in field.Visuals.OrderBy(item => item.PaintOrder)) {
                AddVisual(canvas, child, webFonts, conversionReport, surfaceWidth, surfaceHeight, interactiveFormControls, cancellationToken, textAsSpan, activeClip);
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

    private static void AddLogicalTextGroup(PdfCore.PdfPageCanvas canvas, HtmlRenderLogicalTextGroup group, RegisteredWebFonts webFonts, PdfCore.PdfConversionReport conversionReport, double surfaceWidth, double surfaceHeight, bool interactiveFormControls, CancellationToken cancellationToken, bool textAsSpan, ClipBounds? activeClip) {
        canvas.ActualText(group.Text, nested => {
            foreach (HtmlRenderVisual child in group.Visuals.OrderBy(item => item.PaintOrder)) {
                cancellationToken.ThrowIfCancellationRequested();
                AddVisual(nested, child, webFonts, conversionReport, surfaceWidth, surfaceHeight, interactiveFormControls, cancellationToken, textAsSpan, activeClip);
            }
        });
    }

    private static void AddSemanticGroup(PdfCore.PdfPageCanvas canvas, HtmlRenderSemanticGroup group, RegisteredWebFonts webFonts, PdfCore.PdfConversionReport conversionReport, double surfaceWidth, double surfaceHeight, bool interactiveFormControls, CancellationToken cancellationToken, bool textAsSpan, ClipBounds? activeClip) {
        if (!group.Visuals.Any(ContainsPaintableVisual)) return;
        if (group.Role == HtmlRenderSemanticGroupRole.Artifact) {
            canvas.Artifact(nested => {
                foreach (HtmlRenderVisual child in group.Visuals.OrderBy(item => item.PaintOrder)) {
                    cancellationToken.ThrowIfCancellationRequested();
                    AddVisual(nested, child, webFonts, conversionReport, surfaceWidth, surfaceHeight, interactiveFormControls: false, cancellationToken, textAsSpan: true, activeClip: activeClip);
                }
            });
            return;
        }
        var options = new PdfCore.PdfCanvasStructureOptions {
            ColumnSpan = group.ColumnSpan,
            RowSpan = group.RowSpan,
            HeaderScope = MapTableHeaderScope(group.HeaderScope),
            StructureElementKey = group.StructureElementKey
        };
        bool childTextAsSpan = textAsSpan || IsTextContentGroup(group.Role);
        canvas.Structure(MapSemanticGroupRole(group.Role), nested => {
            void AddChildren(PdfCore.PdfPageCanvas target) {
                foreach (HtmlRenderVisual child in group.Visuals.OrderBy(item => item.PaintOrder)) {
                    cancellationToken.ThrowIfCancellationRequested();
                    AddVisual(target, child, webFonts, conversionReport, surfaceWidth, surfaceHeight, interactiveFormControls, cancellationToken, childTextAsSpan, activeClip);
                }
            }

            if (IsTextContentGroup(group.Role)
                && TryResolveReorderedLogicalText(group.Visuals, out string logicalText)) {
                nested.ActualText(logicalText, AddChildren);
            } else {
                AddChildren(nested);
            }
        }, options);
    }

    private static bool ContainsPaintableVisual(HtmlRenderVisual visual) {
        if (visual is HtmlRenderBookmarkAnchor) return false;
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
        ClipBounds? activeClip) {
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
                AddVisual(nested, child, webFonts, conversionReport, surfaceWidth, surfaceHeight, interactiveFormControls, cancellationToken, textAsSpan, activeClip);
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
        ClipBounds? activeClip) {
        double left = group.ClipHorizontal ? Math.Max(0D, group.ClipX) : 0D;
        double top = group.ClipVertical ? Math.Max(0D, group.ClipY) : 0D;
        double right = group.ClipHorizontal ? Math.Min(surfaceWidth, group.ClipX + group.ClipWidth) : surfaceWidth;
        double bottom = group.ClipVertical ? Math.Min(surfaceHeight, group.ClipY + group.ClipHeight) : surfaceHeight;
        if (right <= left + 0.0001D || bottom <= top + 0.0001D) return;
        ClipBounds clip = ClipBounds.Intersect(activeClip, new ClipBounds(left, top, right, bottom));
        canvas.Clip(
            left * PointsPerCssPixel,
            top * PointsPerCssPixel,
            (right - left) * PointsPerCssPixel,
            (bottom - top) * PointsPerCssPixel,
            clipped => {
                foreach (HtmlRenderVisual child in group.Visuals.OrderBy(item => item.PaintOrder)) {
                    cancellationToken.ThrowIfCancellationRequested();
                    AddVisual(clipped, child, webFonts, conversionReport, surfaceWidth, surfaceHeight, interactiveFormControls, cancellationToken, textAsSpan, clip);
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
        ClipBounds? activeClip) {
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
                    AddVisual(clipped, child, webFonts, conversionReport, surfaceWidth, surfaceHeight, interactiveFormControls, cancellationToken, textAsSpan, clip);
                }
            });
        if (clipX < 0D || clipY < 0D) {
            canvas.Effect(OfficeTransform.Identity, 1D, addClip);
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
        canvas.Drawing(
            drawing,
            visual.X * PointsPerCssPixel,
            visual.Y * PointsPerCssPixel,
            visual.Width * PointsPerCssPixel,
            visual.Height * PointsPerCssPixel,
            style: new PdfCore.PdfDrawingStyle { Decorative = true },
            linkUri: visual.LinkUri,
            linkContents: visual.LinkUri == null ? null : visual.Source);
    }

    private static void AddText(
        PdfCore.PdfPageCanvas canvas,
        HtmlRenderText visual,
        RegisteredWebFonts webFonts,
        double surfaceWidth,
        bool asSpan) {
        if (visual.Text.Length == 0) return;
        string? link = string.IsNullOrWhiteSpace(visual.Text) ? null : visual.LinkUri;
        double frameWidth = visual.Width;
        if (visual.TextAdvanceWidth.HasValue) {
            double metricTolerance = Math.Max(
                visual.Font.Size,
                visual.TextAdvanceWidth.Value * 0.25D);
            frameWidth = Math.Max(frameWidth, visual.TextAdvanceWidth.Value + metricTolerance);
        }
        frameWidth = Math.Max(0.01D, Math.Min(frameWidth, Math.Max(0.01D, surfaceWidth - visual.X)));
        var run = new PdfCore.PdfTextRun(
            visual.Text,
            bold: visual.Font.IsBold,
            underline: visual.Font.IsUnderline,
            color: PdfCore.PdfColor.FromOfficeColorOrNull(visual.Color),
            italic: visual.Font.IsItalic,
            strike: visual.Font.IsStrikethrough,
            fontSize: visual.Font.Size * PointsPerCssPixel,
            font: MapFont(
                visual.Font.FamilyName,
                visual.Text,
                (visual.Font.IsBold ? OfficeFontStyle.Bold : OfficeFontStyle.Regular)
                | (visual.Font.IsItalic ? OfficeFontStyle.Italic : OfficeFontStyle.Regular),
                webFonts),
            linkUri: link,
            linkContents: link == null ? null : visual.Text,
            fontFamily: visual.Font.FamilyName);
        canvas.Text(
            new[] { run },
            asSpan ? PdfCore.PdfCanvasTextStructureRole.Span : MapStructureRole(visual.SemanticRole),
            visual.X * PointsPerCssPixel,
            visual.Y * PointsPerCssPixel,
            frameWidth * PointsPerCssPixel,
            visual.Height * PointsPerCssPixel,
            PdfCore.PdfColor.FromOfficeColorOrNull(visual.Color),
            MapAlignment(visual.Alignment),
            visual.Font.Size * PointsPerCssPixel,
            visual.LineHeight * PointsPerCssPixel);
    }

    private static PdfCore.PdfCanvasTextStructureRole MapStructureRole(string? semanticRole) {
        if (semanticRole == "heading-1") return PdfCore.PdfCanvasTextStructureRole.Heading1;
        if (semanticRole == "heading-2") return PdfCore.PdfCanvasTextStructureRole.Heading2;
        if (semanticRole == "heading-3") return PdfCore.PdfCanvasTextStructureRole.Heading3;
        if (semanticRole == "heading-4") return PdfCore.PdfCanvasTextStructureRole.Heading4;
        if (semanticRole == "heading-5") return PdfCore.PdfCanvasTextStructureRole.Heading5;
        if (semanticRole == "heading-6") return PdfCore.PdfCanvasTextStructureRole.Heading6;
        return semanticRole == "span" ? PdfCore.PdfCanvasTextStructureRole.Span : PdfCore.PdfCanvasTextStructureRole.Paragraph;
    }

    private static void AddImage(PdfCore.PdfPageCanvas canvas, HtmlRenderImage visual) {
        PdfCore.PdfCanvasImageResource? imageResource = GetSharedPdfImageResource(
            visual.EncodedBytes, visual.ContentType);
        if (imageResource == null) return;
        PdfCore.PdfImageStyle? style = visual.SourceCrop.HasCrop
            ? new PdfCore.PdfImageStyle {
                SourceCrop = new PdfCore.PdfImageSourceCrop(
                    visual.SourceCrop.Left,
                    visual.SourceCrop.Top,
                    visual.SourceCrop.Right,
                    visual.SourceCrop.Bottom)
            }
            : null;
        canvas.ImageShared(
            imageResource,
            visual.X * PointsPerCssPixel,
            visual.Y * PointsPerCssPixel,
            visual.Width * PointsPerCssPixel,
            visual.Height * PointsPerCssPixel,
            style,
            linkUri: visual.LinkUri,
            linkContents: visual.LinkUri == null ? null : visual.Source,
            alternativeText: visual.AlternativeText);
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
        OfficeTransform drawingToPage = OfficeTransform.Scale(scaleX * PointsPerCssPixel, scaleY * PointsPerCssPixel)
            .Then(OfficeTransform.Translate(originX, originY));
        OfficeTransform pageToDrawing = drawingToPage.Invert();

        void AddElements(PdfCore.PdfPageCanvas target, IReadOnlyList<OfficeDrawingElement> elements) {
            var shapeBatch = new OfficeDrawing(source.Width, source.Height);
            void FlushShapes() {
                if (shapeBatch.Elements.Count == 0) return;
                cancellationToken.ThrowIfCancellationRequested();
                target.Drawing(
                    shapeBatch,
                    originX,
                    originY,
                    visual.Width * PointsPerCssPixel,
                    visual.Height * PointsPerCssPixel,
                    linkUri: visual.LinkUri,
                    linkContents: visual.LinkUri == null ? null : visual.Source);
                shapeBatch = new OfficeDrawing(source.Width, source.Height);
            }

            foreach (OfficeDrawingElement element in elements) {
                cancellationToken.ThrowIfCancellationRequested();
                if (element is OfficeDrawingShape shape) {
                    shapeBatch.AddShape(shape.Shape, shape.X, shape.Y);
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
                    target.Effect(pageGroupTransform, 1D, grouped => {
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
                if (element is not OfficeDrawingText text || string.IsNullOrWhiteSpace(text.Text)) continue;
                FlushShapes();
                double fontSize = text.Font.Size * scaleY * PointsPerCssPixel;
                double lineHeight = (text.LineHeight ?? text.Font.Size * 1.2D) * scaleY * PointsPerCssPixel;
                PdfCore.PdfColor? color = text.Color.HasValue ? PdfCore.PdfColor.FromOfficeColorOrNull(text.Color.Value) : null;
                IReadOnlyList<OfficeFontFallbackRun> plannedRuns = webFonts.Faces.PlanFallbackRuns(
                    text.Text,
                    text.Font.FamilyName,
                    text.Font.Style);
                IReadOnlyList<PdfCore.PdfTextRun> runs = plannedRuns.Select(run =>
                    new PdfCore.PdfTextRun(
                        run.Text,
                        bold: text.Font.IsBold,
                        underline: text.Font.IsUnderline,
                        color: color,
                        italic: text.Font.IsItalic,
                        strike: text.Font.IsStrikethrough,
                        fontSize: fontSize,
                        font: MapFont(run.FamilyName, run.Text, text.Font.Style, webFonts),
                        linkUri: visual.LinkUri,
                        linkContents: visual.LinkUri == null ? null : run.Text,
                        fontFamily: run.FamilyName))
                    .ToList();
                target.Text(
                    runs,
                    (visual.X + text.X * scaleX) * PointsPerCssPixel,
                    (visual.Y + text.Y * scaleY) * PointsPerCssPixel,
                    text.Width * scaleX * PointsPerCssPixel,
                    text.Height * scaleY * PointsPerCssPixel,
                    color,
                    MapAlignment(text.Alignment),
                    fontSize,
                    lineHeight);
            }
            FlushShapes();
        }

        if (source.Elements.Count == 0) return;
        if (string.IsNullOrWhiteSpace(visual.AlternativeText)) {
            AddElements(canvas, source.Elements);
        } else {
            canvas.Figure(visual.AlternativeText!, figure => AddElements(figure, source.Elements));
        }
    }

    private readonly struct ClipBounds {
        internal ClipBounds(double left, double top, double right, double bottom, bool allowsInteractiveWidgets = true) {
            Left = left;
            Top = top;
            Right = right;
            Bottom = bottom;
            AllowsInteractiveWidgets = allowsInteractiveWidgets;
        }

        private double Left { get; }
        private double Top { get; }
        private double Right { get; }
        private double Bottom { get; }
        internal bool AllowsInteractiveWidgets { get; }

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
                active.Value.AllowsInteractiveWidgets && next.AllowsInteractiveWidgets);
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
