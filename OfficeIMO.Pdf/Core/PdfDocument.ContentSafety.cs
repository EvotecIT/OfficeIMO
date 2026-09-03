using System.Globalization;
using OfficeIMO.ContentSafety;
using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    /// <summary>Inspects decoded PDF text for exact paint, clipping, geometry, contrast, and Unicode evidence.</summary>
    public static OfficeContentSafetyReport InspectContentSafety(
        string filePath,
        OfficeContentSafetyOptions? options = null,
        PdfLoadOptions? readOptions = null) {
        if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("A file path is required.", nameof(filePath));
        OfficeContentSafetyOptions effective = options ?? new OfficeContentSafetyOptions();
        return InspectContentSafety(OfficeContentSafetyInputGuard.ReadAllBytes(filePath, effective), effective, readOptions);
    }

    /// <summary>Inspects encoded PDF bytes without treating concealment as evidence of AI authorship.</summary>
    public static OfficeContentSafetyReport InspectContentSafety(
        byte[] pdf,
        OfficeContentSafetyOptions? options = null,
        PdfLoadOptions? readOptions = null) {
#if NET6_0_OR_GREATER
        ArgumentNullException.ThrowIfNull(pdf);
#else
        if (pdf == null) throw new ArgumentNullException(nameof(pdf));
#endif
        OfficeContentSafetyOptions effective = options ?? new OfficeContentSafetyOptions();
        OfficeContentSafetyInputGuard.ValidateBytes(pdf, effective);
        return InspectPdfContentSafety(pdf, effective, readOptions, targets: null);
    }

    /// <summary>Physically removes exact selected PDF text spans and verifies the rewritten artifact.</summary>
    public static OfficeContentCleanupResult RemoveSelectedContent(
        byte[] pdf,
        OfficeContentCleanupSelection selection,
        OfficeContentCleanupOptions? options = null,
        PdfLoadOptions? readOptions = null) {
#if NET6_0_OR_GREATER
        ArgumentNullException.ThrowIfNull(pdf);
        ArgumentNullException.ThrowIfNull(selection);
#else
        if (pdf == null) throw new ArgumentNullException(nameof(pdf));
        if (selection == null) throw new ArgumentNullException(nameof(selection));
#endif
        options ??= new OfficeContentCleanupOptions();
        options.Validate();
        OfficeContentSafetyReport before = InspectPdfContentSafety(pdf, options.Inspection, readOptions, targets: null);
        IReadOnlyList<OfficeContentSafetyFinding> selected = OfficeContentSafetyBuilder.ResolveSelection(before, selection);
        if (selected.Count == 0) return new OfficeContentCleanupResult((byte[])pdf.Clone(), before, before, Array.Empty<OfficeContentCleanupChange>());

        PdfDocumentSecurityInfo security = PdfSyntax.ReadDocumentSecurityInfo(pdf, readOptions);
        if (security.HasEncryption) throw new InvalidOperationException("PDF content cleanup does not remove or replace encryption. Decrypt the PDF through an explicit security workflow first.");
        if (security.HasSignatures) throw new InvalidOperationException("PDF content cleanup would invalidate existing signatures. OfficeIMO does not silently delete PDF signature revisions or fields; work from an explicitly unsigned copy.");

        var targets = new Dictionary<string, PdfContentSafetyTarget>(StringComparer.Ordinal);
        OfficeContentSafetyReport current = InspectPdfContentSafety(pdf, options.Inspection, readOptions, targets);
        IReadOnlyList<OfficeContentSafetyFinding> currentSelection = OfficeContentSafetyBuilder.ResolveSelection(current, selection);
        var exactTargets = currentSelection.Where(item => !item.SourceTextOffset.HasValue).Select(item => {
            PdfContentSafetyTarget target = targets[item.Id];
            return (target.PageNumber, target.Span);
        }).ToArray();
        var textEdits = currentSelection.Where(item => item.SourceTextOffset.HasValue).Select(item => {
            PdfContentSafetyTarget target = targets[item.Id];
            return (target.PageNumber, target.Span, item.SourceTextOffset!.Value, item.SourceTextLength!.Value);
        }).Where(item => !exactTargets.Any(exact => exact.PageNumber == item.PageNumber && ReferenceEquals(exact.Span, item.Span))).ToArray();
        byte[] output = PdfTextEditor.MutateExactContentSafetySpans(pdf, exactTargets, textEdits, readOptions);
        PdfLoadOptions outputOptions = PdfLoadOptions.WithMinimumInputBytes(readOptions, output.LongLength);
        OfficeContentSafetyReport after = InspectPdfContentSafety(output, options.Inspection, outputOptions, targets: null);
        OfficeContentCleanupChange[] changes = selected.Select(item => new OfficeContentCleanupChange(item.Id, item.Location, item.CleanupCapability)).ToArray();
        return new OfficeContentCleanupResult(output, before, after, changes);
    }

    /// <summary>Atomically writes an explicitly cleaned PDF artifact.</summary>
    public static OfficeContentCleanupResult RemoveSelectedContent(
        string inputPath,
        string outputPath,
        OfficeContentCleanupSelection selection,
        OfficeContentCleanupOptions? options = null,
        PdfLoadOptions? readOptions = null) {
        if (string.IsNullOrWhiteSpace(inputPath)) throw new ArgumentException("An input path is required.", nameof(inputPath));
        if (string.IsNullOrWhiteSpace(outputPath)) throw new ArgumentException("An output path is required.", nameof(outputPath));
        options ??= new OfficeContentCleanupOptions();
        options.Validate();
        OfficeContentCleanupResult result = RemoveSelectedContent(OfficeContentSafetyInputGuard.ReadAllBytes(inputPath, options.Inspection), selection, options, readOptions);
        OfficeFileCommit.WriteAllBytes(outputPath, result.Output);
        return result;
    }

    private static OfficeContentSafetyReport InspectPdfContentSafety(
        byte[] pdf,
        OfficeContentSafetyOptions? options,
        PdfLoadOptions? readOptions,
        IDictionary<string, PdfContentSafetyTarget>? targets) {
        PdfReadDocument document = PdfReadDocument.Open(pdf, readOptions);
        var builder = new OfficeContentSafetyBuilder("PDF", options);
        var concealedAnnotationObjectNumbers = new HashSet<int>();
        bool optionalContentInspectionInconclusive = false;
        for (int pageIndex = 0; pageIndex < document.Pages.Count; pageIndex++) {
            PdfReadPage page = document.Pages[pageIndex];
            (double pageWidth, double pageHeight) = page.GetPageSize();
            IReadOnlyList<PdfTextSpan> spans = page.GetTextSpans(includeArtifactText: true);
            for (int spanIndex = 0; spanIndex < spans.Count; spanIndex++) {
                PdfTextSpan span = spans[spanIndex];
                if (string.IsNullOrWhiteSpace(span.Text)) continue;
                string location = "Page[" + (pageIndex + 1).ToString(CultureInfo.InvariantCulture) + "]/TextSpan[" + (spanIndex + 1).ToString(CultureInfo.InvariantCulture) + "]";
                OfficeContentConcealmentKind? kind = null;
                string? evidence = null;
                if (!span.IsVisible) {
                    kind = OfficeContentConcealmentKind.InvisibleRenderingMode;
                    evidence = "The PDF text rendering mode " + span.TextRenderingMode.ToString(CultureInfo.InvariantCulture) + " does not paint visible glyphs.";
                } else if (span.Color?.A <= 3) {
                    kind = OfficeContentConcealmentKind.TransparentText;
                    evidence = "The effective PDF text paint alpha is fully or nearly transparent.";
                } else if (span.ClipPath.HasValue && !span.CanProjectCompleteText(pageHeight)) {
                    kind = OfficeContentConcealmentKind.ClippedContent;
                    evidence = "The active PDF clipping path does not contain the complete painted text span.";
                } else if (span.Advance <= 0.01D || span.FontSize <= 0D) {
                    kind = OfficeContentConcealmentKind.ZeroDimension;
                    evidence = "The decoded PDF text span has zero or near-zero painted geometry.";
                } else if (span.FontSize <= builder.Options.MaximumTinyFontSizePoints) {
                    kind = OfficeContentConcealmentKind.TinyText;
                    evidence = "The decoded PDF font size is " + span.FontSize.ToString("0.###", CultureInfo.InvariantCulture) + "pt.";
                } else if (IsPdfSpanOffCanvas(span, pageWidth, pageHeight)) {
                    kind = OfficeContentConcealmentKind.OffCanvas;
                    evidence = "The decoded PDF text span is positioned entirely outside the page boundary.";
                } else if (span.Color.HasValue && page.TryGetContentSafetyBackground(span, out OfficeColor background, out string backgroundEvidence)) {
                    double ratio = OfficeColorContrast.ContrastRatio(span.Color.Value, background);
                    if (ratio < builder.Options.MinimumVisibleContrastRatio) {
                        kind = OfficeContentConcealmentKind.LowContrastText;
                        evidence = "The decoded PDF text color against " + backgroundEvidence + " has contrast ratio " + ratio.ToString("0.###", CultureInfo.InvariantCulture) + ".";
                    }
                }

                if (kind.HasValue) {
                    OfficeContentSafetyFinding finding = builder.Add(
                        kind.Value,
                        OfficeContentSafetyRisk.ContextDependent,
                        location,
                        evidence!,
                        span.Text,
                        OfficeContentCleanupCapability.RemoveText,
                        inspectTextIntegrityEvidence: false);
                    if (targets != null) targets[finding.Id] = new PdfContentSafetyTarget(pageIndex + 1, span);
                    OfficeContentCleanupCapability unicodeCapability = kind.Value == OfficeContentConcealmentKind.LowContrastText
                        ? OfficeContentCleanupCapability.RemoveText
                        : OfficeContentCleanupCapability.ReportOnly;
                    IReadOnlyList<OfficeContentSafetyFinding> unicode = builder.InspectChargedTextIntegrity(location + "/Text", span.Text, unicodeCapability);
                    if (targets != null && unicodeCapability == OfficeContentCleanupCapability.RemoveText) {
                        foreach (OfficeContentSafetyFinding item in unicode) targets[item.Id] = new PdfContentSafetyTarget(pageIndex + 1, span);
                    }
                } else {
                    IReadOnlyList<OfficeContentSafetyFinding> unicode = builder.InspectVisibleText(location, span.Text, OfficeContentCleanupCapability.RemoveText);
                    if (targets != null) foreach (OfficeContentSafetyFinding item in unicode) targets[item.Id] = new PdfContentSafetyTarget(pageIndex + 1, span);
                }
            }

            bool hasOptionalContent = document.OptionalContent != null;
            bool unsupportedOptionalContentViewUsage = hasOptionalContent && page.HasUnsupportedOptionalContentViewUsageApplications();
            optionalContentInspectionInconclusive |= unsupportedOptionalContentViewUsage;
            IReadOnlyList<PdfTextSpan> hiddenOptionalContent = !hasOptionalContent || unsupportedOptionalContentViewUsage
                ? Array.Empty<PdfTextSpan>()
                : page.GetHiddenOptionalContentTextSpans(includeArtifactText: true);
            for (int hiddenIndex = 0; hiddenIndex < hiddenOptionalContent.Count; hiddenIndex++) {
                PdfTextSpan span = hiddenOptionalContent[hiddenIndex];
                if (string.IsNullOrWhiteSpace(span.Text)) continue;
                string location = "Page[" + (pageIndex + 1).ToString(CultureInfo.InvariantCulture) + "]/HiddenOptionalContentTextSpan[" + (hiddenIndex + 1).ToString(CultureInfo.InvariantCulture) + "]";
                builder.Add(
                    OfficeContentConcealmentKind.HiddenContainer,
                    OfficeContentSafetyRisk.ContextDependent,
                    location,
                    "The decoded PDF text is inside optional content hidden by the document's default layer configuration.",
                    span.Text,
                    OfficeContentCleanupCapability.ReportOnly,
                    inspectTextIntegrityEvidence: true);
            }

            if (!builder.Options.IncludeNonPrimaryContent) continue;

            IReadOnlyList<PdfAnnotation> annotations = page.GetAnnotationsForContentSafety();
            for (int annotationIndex = 0; annotationIndex < annotations.Count; annotationIndex++) {
                PdfAnnotation annotation = annotations[annotationIndex];
                bool hiddenByFlags = annotation.IsHidden || annotation.IsInvisible || annotation.IsNoView;
                bool hiddenByOptionalContent = page.IsHiddenOptionalContent(annotation.SourceDictionary);
                bool unreadableRectangle = !annotation.HasReadableRectangle;
                bool degenerateRectangle = annotation.HasReadableRectangle && HasDegenerateAnnotationRectangle(annotation);
                bool outsidePage = annotation.HasReadableRectangle && IsAnnotationOutsidePage(page, annotation);
                if (!hiddenByFlags && !hiddenByOptionalContent && !unreadableRectangle && !degenerateRectangle && !outsidePage) continue;
                if (annotation.ObjectNumber.HasValue) concealedAnnotationObjectNumbers.Add(annotation.ObjectNumber.Value);
                string location = "Page[" + (pageIndex + 1).ToString(CultureInfo.InvariantCulture) + "]/HiddenAnnotation[" + (annotationIndex + 1).ToString(CultureInfo.InvariantCulture) + "]";
                string evidence = unreadableRectangle
                    ? "The PDF annotation has no readable rectangle and cannot provide a visible presentation for its stored content."
                    : degenerateRectangle
                    ? "The PDF annotation rectangle has zero area and cannot present its stored content."
                    : outsidePage
                    ? "The PDF annotation rectangle is outside the page boundary and has no visible presentation."
                    : hiddenByFlags && hiddenByOptionalContent
                    ? "The PDF annotation is concealed by its flags and optional-content configuration."
                    : hiddenByOptionalContent
                        ? "The PDF annotation is concealed by the document's default optional-content configuration."
                        : "The PDF annotation is concealed by its Invisible, Hidden, or NoView flag.";
                string? richText = !string.IsNullOrWhiteSpace(annotation.RichContentsPlainText)
                    ? annotation.RichContentsPlainText
                    : annotation.RichContents;
                bool addedText = false;
                if (!string.IsNullOrWhiteSpace(richText)) {
                    builder.Add(
                        OfficeContentConcealmentKind.HiddenByProperty,
                        OfficeContentSafetyRisk.ContextDependent,
                        location + "/RichContents",
                        evidence,
                        richText,
                        OfficeContentCleanupCapability.ReportOnly);
                    addedText = true;
                }
                if (!string.IsNullOrWhiteSpace(annotation.Contents) &&
                    !string.Equals(annotation.Contents, richText, StringComparison.Ordinal)) {
                    builder.Add(
                        OfficeContentConcealmentKind.HiddenByProperty,
                        OfficeContentSafetyRisk.ContextDependent,
                        location + "/Contents",
                        evidence,
                        annotation.Contents,
                        OfficeContentCleanupCapability.ReportOnly);
                    addedText = true;
                }
                if (!addedText) {
                    builder.Add(
                        OfficeContentConcealmentKind.HiddenByProperty,
                        OfficeContentSafetyRisk.ContextDependent,
                        location,
                        evidence,
                        text: null,
                        cleanupCapability: OfficeContentCleanupCapability.ReportOnly);
                }
            }
        }

        IReadOnlyList<PdfFormField> formFields = document.FormFields;
        var reportedValueOwners = new HashSet<int>();
        var reportedDefaultValueOwners = new HashSet<int>();
        var reportedRichValueOwners = new HashSet<int>();
        for (int fieldIndex = 0; builder.Options.IncludeNonPrimaryContent && fieldIndex < formFields.Count; fieldIndex++) {
            PdfFormField field = formFields[fieldIndex];
            string location = "FormField[" + (fieldIndex + 1).ToString(CultureInfo.InvariantCulture) + "]";
            string? currentValues = field.HasValueEntry ? string.Join(" ", field.Values) : null;
            int valueOwnerKey = field.ValueOwnerKey ?? -(fieldIndex + 1);
            bool hiddenChoiceExportValue = HasChoiceExportValueHiddenByDisplayText(field, defaultValue: false);
            if (currentValues is not null &&
                reportedValueOwners.Add(valueOwnerKey) &&
                (hiddenChoiceExportValue ||
                 !HasVisibleWidgetForValueOwner(document, formFields, fieldIndex, defaultValue: false, concealedAnnotationObjectNumbers))) {
                builder.Add(
                    OfficeContentConcealmentKind.HiddenByProperty,
                    OfficeContentSafetyRisk.ContextDependent,
                    location + (hiddenChoiceExportValue ? "/HiddenChoiceExportValue" : "/HiddenWidgetValue"),
                    hiddenChoiceExportValue
                        ? "The PDF choice field stores an export value that differs from the display text presented by its widget."
                        : "The PDF form value has no visible presentation because its widgets are absent, masked by the field type, outside the page boundary, or concealed by annotation flags or optional-content configuration.",
                    currentValues,
                    OfficeContentCleanupCapability.ReportOnly);
            }
            string? defaultValues = field.HasDefaultValueEntry ? string.Join(" ", field.DefaultValues) : null;
            int defaultValueOwnerKey = field.DefaultValueOwnerKey ?? -(fieldIndex + 1);
            bool distinctStoredDefault = currentValues is not null &&
                !string.Equals(defaultValues, currentValues, StringComparison.Ordinal);
            bool hiddenChoiceDefaultExportValue = HasChoiceExportValueHiddenByDisplayText(field, defaultValue: true);
            if (defaultValues is not null &&
                !string.Equals(defaultValues, currentValues, StringComparison.Ordinal) &&
                (distinctStoredDefault ||
                 hiddenChoiceDefaultExportValue ||
                 !HasVisibleWidgetForValueOwner(document, formFields, fieldIndex, defaultValue: true, concealedAnnotationObjectNumbers)) &&
                reportedDefaultValueOwners.Add(defaultValueOwnerKey)) {
                builder.Add(
                    OfficeContentConcealmentKind.HiddenByProperty,
                    OfficeContentSafetyRisk.ContextDependent,
                    location + (hiddenChoiceDefaultExportValue ? "/HiddenChoiceDefaultExportValue" : "/HiddenWidgetDefaultValue"),
                    distinctStoredDefault
                        ? "The PDF default form value differs from the current value presented by its widgets and remains stored for reset or other viewer behavior."
                        : hiddenChoiceDefaultExportValue
                            ? "The PDF choice field stores a default export value that differs from the display text presented by its widget."
                        : "The PDF default form value has no visible presentation because its widgets are absent, masked by the field type, outside the page boundary, or concealed by annotation flags or optional-content configuration.",
                    defaultValues,
                    OfficeContentCleanupCapability.ReportOnly);
            }
            string? richValue = field.RichValue;
            string? richValueText = field.RichValuePlainText;
            int richValueOwnerKey = field.RichValueOwnerKey ?? -(fieldIndex + 1);
            if (field.HasRichValueEntry &&
                !string.IsNullOrWhiteSpace(richValue) &&
                !string.Equals(richValueText ?? richValue, currentValues, StringComparison.Ordinal) &&
                reportedRichValueOwners.Add(richValueOwnerKey)) {
                builder.Add(
                    OfficeContentConcealmentKind.HiddenByProperty,
                    OfficeContentSafetyRisk.ContextDependent,
                    location + "/HiddenRichValue",
                    "The PDF form field stores an independent rich-text value that differs from its simple value and may not match the widget presentation.",
                    richValue,
                    OfficeContentCleanupCapability.ReportOnly);
            }
        }

        if (optionalContentInspectionInconclusive) {
            builder.AddDiagnostic("Optional-content hidden-state inspection was inconclusive because the document uses unsupported default view intent or usage applications.");
        } else if (document.OptionalContent != null) {
            builder.AddDiagnostic("Optional-content metadata and hidden text were inspected using the document's default layer configuration. Hidden optional-content findings are report-only.");
        }
        return builder.Build();
    }

    private static bool HasVisibleWidgetForValueOwner(
        PdfReadDocument document,
        IReadOnlyList<PdfFormField> fields,
        int fieldIndex,
        bool defaultValue,
        HashSet<int> concealedAnnotationObjectNumbers) {
        PdfFormField field = fields[fieldIndex];
        int? ownerKey = defaultValue
            ? field.DefaultValueOwnerKey
            : field.ValueOwnerKey;
        for (int candidateIndex = 0; candidateIndex < fields.Count; candidateIndex++) {
            PdfFormField candidate = fields[candidateIndex];
            int? candidateOwnerKey = defaultValue
                ? candidate.DefaultValueOwnerKey
                : candidate.ValueOwnerKey;
            if (ownerKey.HasValue
                    ? candidateOwnerKey != ownerKey
                    : candidateIndex != fieldIndex) continue;
            if (candidate.IsPassword) continue;
            if (defaultValue &&
                candidate.HasValueEntry &&
                !candidate.DefaultValues.SequenceEqual(candidate.Values, StringComparer.Ordinal)) {
                continue;
            }
            for (int widgetIndex = 0; widgetIndex < candidate.Widgets.Count; widgetIndex++) {
                PdfFormWidget widget = candidate.Widgets[widgetIndex];
                bool concealed = !widget.PageNumber.HasValue ||
                    IsWidgetOutsidePage(document, widget) ||
                    widget.IsHidden ||
                    widget.IsInvisible ||
                    widget.IsNoView ||
                    widget.ObjectNumber.HasValue && concealedAnnotationObjectNumbers.Contains(widget.ObjectNumber.Value);
                if (concealed) continue;
                IReadOnlyList<string> presentedValues = defaultValue
                    ? candidate.DefaultValues
                    : candidate.Values;
                if (!candidate.IsButtonField ||
                    widget.AppearanceState is not null && presentedValues.Contains(widget.AppearanceState, StringComparer.Ordinal)) {
                    return true;
                }
            }
        }
        return false;
    }

    private static bool HasChoiceExportValueHiddenByDisplayText(PdfFormField field, bool defaultValue) {
        if (!field.IsChoiceField) return false;
        IReadOnlyList<PdfFormFieldOption> options = defaultValue
            ? field.DefaultSelectedOptions
            : field.SelectedOptions;
        for (int i = 0; i < options.Count; i++) {
            if (options[i].HasSeparateDisplayText) return true;
        }
        return false;
    }

    private static bool IsWidgetOutsidePage(PdfReadDocument document, PdfFormWidget widget) {
        if (!widget.PageNumber.HasValue || widget.PageNumber.Value < 1 || widget.PageNumber.Value > document.Pages.Count) return true;
        PdfReadPage page = document.Pages[widget.PageNumber.Value - 1];
        (double originX, double originY) = page.GetPageBoundaryOrigin();
        (double width, double height) = page.GetPageSize();
        double left = Math.Min(widget.X1, widget.X2);
        double right = Math.Max(widget.X1, widget.X2);
        double bottom = Math.Min(widget.Y1, widget.Y2);
        double top = Math.Max(widget.Y1, widget.Y2);
        if (right <= left || top <= bottom) return true;
        return right <= originX || top <= originY || left >= originX + width || bottom >= originY + height;
    }

    private static bool IsAnnotationOutsidePage(PdfReadPage page, PdfAnnotation annotation) {
        (double originX, double originY) = page.GetPageBoundaryOrigin();
        (double width, double height) = page.GetPageSize();
        double left = Math.Min(annotation.X1, annotation.X2);
        double right = Math.Max(annotation.X1, annotation.X2);
        double bottom = Math.Min(annotation.Y1, annotation.Y2);
        double top = Math.Max(annotation.Y1, annotation.Y2);
        return right <= originX || top <= originY || left >= originX + width || bottom >= originY + height;
    }

    private static bool HasDegenerateAnnotationRectangle(PdfAnnotation annotation) {
        double left = Math.Min(annotation.X1, annotation.X2);
        double right = Math.Max(annotation.X1, annotation.X2);
        double bottom = Math.Min(annotation.Y1, annotation.Y2);
        double top = Math.Max(annotation.Y1, annotation.Y2);
        return right <= left || top <= bottom;
    }

    private static bool IsPdfSpanOffCanvas(PdfTextSpan span, double pageWidth, double pageHeight) {
        double advance = Math.Max(0.1D, Math.Abs(span.Advance));
        double height = Math.Max(0.1D, Math.Abs(span.FontSize));
        double radians = span.RotationDegrees * Math.PI / 180D;
        double endX = span.X + Math.Cos(radians) * advance;
        double endY = span.Y + Math.Sin(radians) * advance;
        double left = Math.Min(span.X, endX) - height;
        double right = Math.Max(span.X, endX) + height;
        double bottom = Math.Min(span.Y, endY) - height;
        double top = Math.Max(span.Y, endY) + height;
        return right <= 0D || top <= 0D || left >= pageWidth || bottom >= pageHeight;
    }

    private sealed class PdfContentSafetyTarget {
        internal PdfContentSafetyTarget(int pageNumber, PdfTextSpan span) { PageNumber = pageNumber; Span = span; }
        internal int PageNumber { get; }
        internal PdfTextSpan Span { get; }
    }
}
