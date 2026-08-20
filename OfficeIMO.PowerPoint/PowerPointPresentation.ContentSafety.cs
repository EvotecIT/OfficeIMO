using System.Globalization;
using System.Threading;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.ContentSafety;
using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.Provenance;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using P188 = DocumentFormat.OpenXml.Office2021.PowerPoint.Comment;

namespace OfficeIMO.PowerPoint;

public sealed partial class PowerPointPresentation {
    /// <summary>Inspects PPTX, PPTM, and legacy PPT-family presentations through the normal first-party loader.</summary>
    public static OfficeContentSafetyReport InspectContentSafety(
        string filePath,
        OfficeContentSafetyOptions? options = null) {
        if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("A file path is required.", nameof(filePath));
        OfficeContentSafetyOptions effective = options ?? new OfficeContentSafetyOptions();
        return InspectContentSafety(OfficeContentSafetyInputGuard.ReadAllBytes(filePath, effective, inspectZipPackage: true), Path.GetFileName(filePath), effective);
    }

    /// <summary>Inspects encoded presentation bytes. The file name preserves macro and legacy-family routing.</summary>
    public static OfficeContentSafetyReport InspectContentSafety(
        byte[] presentationBytes,
        string fileName = "presentation.pptx",
        OfficeContentSafetyOptions? options = null) {
        if (presentationBytes == null) throw new ArgumentNullException(nameof(presentationBytes));
        OfficeContentSafetyOptions effective = options ?? new OfficeContentSafetyOptions();
        OfficeContentSafetyInputGuard.ValidateBytes(presentationBytes, effective, inspectZipPackage: true);
        using PowerPointPresentation presentation = LoadContentSafetyPresentation(presentationBytes, fileName, readOnly: true);
        return InspectContentSafetyPresentation(presentation, effective, targets: null);
    }

    /// <summary>Removes exact selected concealed-content findings and emits the same physical PowerPoint format.</summary>
    public static OfficeContentCleanupResult RemoveSelectedContent(
        byte[] presentationBytes,
        OfficeContentCleanupSelection selection,
        string fileName = "presentation.pptx",
        OfficeContentCleanupOptions? options = null) {
        if (presentationBytes == null) throw new ArgumentNullException(nameof(presentationBytes));
        if (selection == null) throw new ArgumentNullException(nameof(selection));
        options ??= new OfficeContentCleanupOptions();
        options.Validate();
        OfficeContentSafetyReport before = InspectContentSafety(presentationBytes, fileName, options.Inspection);
        IReadOnlyList<OfficeContentSafetyFinding> selected = OfficeContentSafetyBuilder.ResolveSelection(before, selection);
        if (selected.Count == 0) return new OfficeContentCleanupResult((byte[])presentationBytes.Clone(), before, before, Array.Empty<OfficeContentCleanupChange>());

        using PowerPointPresentation presentation = LoadContentSafetyPresentation(presentationBytes, fileName, readOnly: false);
        presentation.SignatureMutationPolicy = options.SignatureMutationPolicy;
        var targets = new Dictionary<string, PowerPointCleanupTarget>(StringComparer.Ordinal);
        OfficeContentSafetyReport current = InspectContentSafetyPresentation(presentation, options.Inspection, targets);
        IReadOnlyList<OfficeContentSafetyFinding> currentSelection = OfficeContentSafetyBuilder.ResolveSelection(current, selection);
        foreach (IGrouping<PowerPointCleanupTarget, OfficeContentSafetyFinding> group in currentSelection
            .OrderByDescending(item => item.SourceTextOffset ?? -1)
            .GroupBy(item => targets[item.Id])) group.Key.Remove();

        byte[] output = presentation.ToBytes(presentation.SourceFormat, new PowerPointSaveOptions());
        OfficeContentSafetyReport after = InspectContentSafety(output, fileName, options.Inspection);
        OfficeContentCleanupChange[] changes = selected.Select(item => new OfficeContentCleanupChange(item.Id, item.Location, item.CleanupCapability)).ToArray();
        return new OfficeContentCleanupResult(output, before, after, changes);
    }

    /// <summary>Atomically writes an explicitly cleaned PowerPoint artifact.</summary>
    public static OfficeContentCleanupResult RemoveSelectedContent(
        string inputPath,
        string outputPath,
        OfficeContentCleanupSelection selection,
        OfficeContentCleanupOptions? options = null) {
        if (string.IsNullOrWhiteSpace(inputPath)) throw new ArgumentException("An input path is required.", nameof(inputPath));
        if (string.IsNullOrWhiteSpace(outputPath)) throw new ArgumentException("An output path is required.", nameof(outputPath));
        options ??= new OfficeContentCleanupOptions();
        options.Validate();
        OfficeContentCleanupResult result = RemoveSelectedContent(OfficeContentSafetyInputGuard.ReadAllBytes(inputPath, options.Inspection, inspectZipPackage: true), selection, Path.GetFileName(inputPath), options);
        OfficeFileCommit.WriteAllBytes(outputPath, result.Output);
        return result;
    }

    private static PowerPointPresentation LoadContentSafetyPresentation(byte[] bytes, string fileName, bool readOnly) =>
        LoadDocument(bytes, fileName, sourceStream: null, new PowerPointLoadOptions {
            AccessMode = readOnly ? DocumentAccessMode.ReadOnly : DocumentAccessMode.ReadWrite,
            PersistenceMode = DocumentPersistenceMode.Explicit
        }, CancellationToken.None);

    private static OfficeContentSafetyReport InspectContentSafetyPresentation(
        PowerPointPresentation presentation,
        OfficeContentSafetyOptions? options,
        IDictionary<string, PowerPointCleanupTarget>? targets) {
        var builder = new OfficeContentSafetyBuilder("PowerPoint " + presentation.SourceFormat, options);
        long slideWidth = presentation.PresentationRoot.SlideSize?.Cx?.Value ?? 0L;
        long slideHeight = presentation.PresentationRoot.SlideSize?.Cy?.Value ?? 0L;
        for (int slideIndex = 0; slideIndex < presentation.Slides.Count; slideIndex++) {
            PowerPointSlide slide = presentation.Slides[slideIndex];
            InspectPowerPointSlide(slide, slideIndex + 1, slideWidth, slideHeight, builder, targets);
        }
        if (presentation.SourceFormat is PowerPointFileFormat.Ppt or PowerPointFileFormat.Pot or PowerPointFileFormat.Pps) {
            foreach (string diagnostic in presentation.LegacyPptImportDiagnostics.Select(item => item.ToString())) builder.AddDiagnostic(diagnostic);
        }
        builder.AddDiagnostic("PowerPoint color and font findings evaluate explicit run, shape, and resolved slide-background values. Conditional animation states and every master-list-style precedence combination are retained but not simulated as a timed slide show.");
        return builder.Build();
    }

    private static void InspectPowerPointSlide(
        PowerPointSlide slide,
        int slideNumber,
        long slideWidth,
        long slideHeight,
        OfficeContentSafetyBuilder builder,
        IDictionary<string, PowerPointCleanupTarget>? targets) {
        P.Slide root = slide.SlidePart.Slide ?? throw new InvalidDataException("A slide part has no slide root.");
        string slideLocation = "Slide[" + slideNumber.ToString(CultureInfo.InvariantCulture) + "]";
        string slideText = string.Concat(root.Descendants<A.Text>().Select(item => item.Text));
        if (slide.Hidden && !string.IsNullOrWhiteSpace(slideText)) {
            int shapeIndex = 0;
            foreach (PowerPointShape shape in slide.Shapes) {
                string text = string.Concat(shape.Element.Descendants<A.Text>().Select(item => item.Text));
                if (string.IsNullOrWhiteSpace(text)) continue;
                OfficeContentCleanupCapability capability = shape is PowerPointTextBox
                    ? OfficeContentCleanupCapability.RemoveText
                    : OfficeContentCleanupCapability.RemoveElement;
                OfficeContentSafetyFinding finding = builder.Add(
                    OfficeContentConcealmentKind.HiddenContainer,
                    OfficeContentSafetyRisk.ContextDependent,
                    slideLocation + "/Shape[" + (++shapeIndex).ToString(CultureInfo.InvariantCulture) + "](" + (shape.Name ?? shape.ShapeContentType.ToString()) + ")",
                    "The owning slide is hidden from the ordinary slide show.",
                    text,
                    capability,
                    inspectTextIntegrityEvidence: false);
                if (targets != null) targets[finding.Id] = PowerPointCleanupTarget.ForShapeContent(shape);
                InspectPowerPointChargedText(shape.Element.Descendants<A.Text>(), finding.Location, builder, targets);
            }
        } else {
            InspectPowerPointSlideRuns(slide, slideLocation, slideWidth, slideHeight, builder, targets);
        }

        if (builder.Options.IncludeNonPrimaryContent) {
            InspectPowerPointAlternativeText(root, slideLocation, builder, targets);
            InspectPowerPointNotes(slide, slideLocation, builder, targets);
            InspectPowerPointComments(slide, slideLocation, builder, targets);
        }
    }

    private static void InspectPowerPointSlideRuns(
        PowerPointSlide slide,
        string slideLocation,
        long slideWidth,
        long slideHeight,
        OfficeContentSafetyBuilder builder,
        IDictionary<string, PowerPointCleanupTarget>? targets) {
        P.Slide root = slide.SlidePart.Slide!;
        A.ColorScheme? scheme = GetContentSafetyColorScheme(slide);
        var concealedOwners = new HashSet<OpenXmlElement>();
        int runIndex = 0;
        foreach (A.Run run in root.Descendants<A.Run>()) {
            string text = run.Text?.Text ?? string.Empty;
            if (string.IsNullOrWhiteSpace(text)) continue;
            OpenXmlElement? owner = FindPowerPointDrawingOwner(run);
            string location = slideLocation + "/Run[" + (++runIndex).ToString(CultureInfo.InvariantCulture) + "]";
            if (owner != null && TryGetPowerPointOwnerConcealment(owner, slideWidth, slideHeight, out OfficeContentConcealmentKind ownerKind, out string ownerEvidence)) {
                if (!concealedOwners.Add(owner)) continue;
                string ownerText = string.Concat(owner.Descendants<A.Text>().Select(item => item.Text));
                OfficeContentSafetyFinding ownerFinding = builder.Add(
                    ownerKind,
                    OfficeContentSafetyRisk.ContextDependent,
                    slideLocation + "/Shape(" + GetPowerPointShapeName(owner) + ")",
                    ownerEvidence,
                    ownerText,
                    OfficeContentCleanupCapability.RemoveElement,
                    inspectTextIntegrityEvidence: false);
                if (targets != null) targets[ownerFinding.Id] = PowerPointCleanupTarget.ForElement(owner);
                InspectPowerPointChargedText(owner.Descendants<A.Text>(), ownerFinding.Location, builder, targets);
                continue;
            }

            A.RunProperties? properties = run.RunProperties;
            A.DefaultRunProperties? defaults = run.Ancestors<A.Paragraph>().FirstOrDefault()?.ParagraphProperties?.GetFirstChild<A.DefaultRunProperties>();
            int? fontSize = properties?.FontSize?.Value ?? defaults?.FontSize?.Value;
            A.SolidFill? foregroundFill = properties?.GetFirstChild<A.SolidFill>() ?? defaults?.GetFirstChild<A.SolidFill>();
            bool transparent = IsPowerPointTransparent(foregroundFill);
            OfficeContentConcealmentKind? kind = null;
            string? evidence = null;
            if (fontSize.HasValue && fontSize.Value / 100D <= builder.Options.MaximumTinyFontSizePoints) {
                kind = OfficeContentConcealmentKind.TinyText;
                evidence = "The effective explicit PowerPoint run font size is " + (fontSize.Value / 100D).ToString("0.###", CultureInfo.InvariantCulture) + "pt.";
            } else if (transparent) {
                kind = OfficeContentConcealmentKind.TransparentText;
                evidence = "The explicit PowerPoint run color is fully or nearly transparent.";
            } else if (foregroundFill != null && TryResolvePowerPointBackground(owner, slide, scheme, out OfficeColor background) &&
                       OfficeOpenXmlThemeColorResolver.ResolveColor(foregroundFill, scheme) is OfficeColor foreground &&
                       OfficeColorContrast.ContrastRatio(foreground, background) < builder.Options.MinimumVisibleContrastRatio) {
                double ratio = OfficeColorContrast.ContrastRatio(foreground, background);
                kind = OfficeContentConcealmentKind.LowContrastText;
                evidence = "The explicit PowerPoint run/background contrast ratio is " + ratio.ToString("0.###", CultureInfo.InvariantCulture) + ".";
            }

            if (kind.HasValue) {
                OfficeContentSafetyFinding finding = builder.Add(kind.Value, OfficeContentSafetyRisk.ContextDependent, location, evidence!, text, OfficeContentCleanupCapability.RemoveElement, inspectTextIntegrityEvidence: false);
                if (targets != null) targets[finding.Id] = PowerPointCleanupTarget.ForElement(run);
            }
            A.Text? textNode = run.Text;
            if (textNode != null) {
                IReadOnlyList<OfficeContentSafetyFinding> unicode = kind.HasValue
                    ? builder.InspectChargedTextIntegrity(location + "/Text", text, OfficeContentCleanupCapability.RemoveText)
                    : builder.InspectVisibleText(location + "/Text", text, OfficeContentCleanupCapability.RemoveText);
                if (targets != null) foreach (OfficeContentSafetyFinding item in unicode) targets[item.Id] = PowerPointCleanupTarget.ForTextRange(textNode, item);
            }
        }

        foreach (A.Field field in root.Descendants<A.Field>()) {
            string text = field.Text?.Text ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(text) && field.Text != null) {
                string location = slideLocation + "/Field(" + (field.Id?.Value ?? "unknown") + ")";
                IReadOnlyList<OfficeContentSafetyFinding> unicode = builder.InspectVisibleText(location, text, OfficeContentCleanupCapability.RemoveText);
                if (targets != null) foreach (OfficeContentSafetyFinding item in unicode) targets[item.Id] = PowerPointCleanupTarget.ForTextRange(field.Text, item);
            }
        }
    }

    private static void InspectPowerPointChargedText(
        IEnumerable<A.Text> textNodes,
        string location,
        OfficeContentSafetyBuilder builder,
        IDictionary<string, PowerPointCleanupTarget>? targets) {
        int index = 0;
        foreach (A.Text textNode in textNodes) {
            string text = textNode.Text ?? string.Empty;
            if (text.Length == 0) continue;
            IReadOnlyList<OfficeContentSafetyFinding> unicode = builder.InspectChargedTextIntegrity(
                location + "/Text[" + (++index).ToString(CultureInfo.InvariantCulture) + "]",
                text,
                OfficeContentCleanupCapability.RemoveText);
            if (targets != null) foreach (OfficeContentSafetyFinding item in unicode) targets[item.Id] = PowerPointCleanupTarget.ForTextRange(textNode, item);
        }
    }

    private static bool TryGetPowerPointOwnerConcealment(
        OpenXmlElement owner,
        long slideWidth,
        long slideHeight,
        out OfficeContentConcealmentKind kind,
        out string evidence) {
        foreach (OpenXmlElement candidate in owner.Ancestors().Prepend(owner)) {
            P.NonVisualDrawingProperties? nonVisual = candidate.Descendants<P.NonVisualDrawingProperties>().FirstOrDefault();
            if (nonVisual?.Hidden?.Value == true) {
                kind = OfficeContentConcealmentKind.HiddenByProperty;
                evidence = "The owning PowerPoint shape or group has its native hidden flag enabled.";
                return true;
            }
        }
        A.Offset? offset = owner.Descendants<A.Offset>().FirstOrDefault();
        A.Extents? extents = owner.Descendants<A.Extents>().FirstOrDefault();
        long x = offset?.X?.Value ?? 0L;
        long y = offset?.Y?.Value ?? 0L;
        long width = extents?.Cx?.Value ?? 0L;
        long height = extents?.Cy?.Value ?? 0L;
        if (extents != null && (width <= 0L || height <= 0L)) {
            kind = OfficeContentConcealmentKind.ZeroDimension;
            evidence = "The owning PowerPoint shape has zero visible width or height.";
            return true;
        }
        if (offset != null && extents != null && slideWidth > 0L && slideHeight > 0L &&
            (x + width <= 0L || y + height <= 0L || x >= slideWidth || y >= slideHeight)) {
            kind = OfficeContentConcealmentKind.OffCanvas;
            evidence = "The owning PowerPoint shape is positioned entirely outside the slide canvas.";
            return true;
        }
        kind = default;
        evidence = string.Empty;
        return false;
    }

    private static OpenXmlElement? FindPowerPointDrawingOwner(OpenXmlElement element) => element.Ancestors().FirstOrDefault(item =>
        item is P.Shape or P.GraphicFrame or P.Picture or P.ConnectionShape or P.GroupShape);

    private static string GetPowerPointShapeName(OpenXmlElement owner) =>
        owner.Descendants<P.NonVisualDrawingProperties>().FirstOrDefault()?.Name?.Value
        ?? owner.LocalName;

    private static A.ColorScheme? GetContentSafetyColorScheme(PowerPointSlide slide) =>
        slide.SlidePart.ThemeOverridePart?.ThemeOverride?.ColorScheme
        ?? slide.SlidePart.SlideLayoutPart?.ThemeOverridePart?.ThemeOverride?.ColorScheme
        ?? slide.SlidePart.SlideLayoutPart?.SlideMasterPart?.ThemePart?.Theme?.ThemeElements?.ColorScheme;

    private static bool IsPowerPointTransparent(A.SolidFill? fill) {
        if (fill == null) return false;
        int? alpha = fill.Descendants<A.Alpha>().FirstOrDefault()?.Val?.Value;
        return alpha.HasValue && alpha.Value <= 1000;
    }

    private static bool TryResolvePowerPointBackground(OpenXmlElement? owner, PowerPointSlide slide, A.ColorScheme? scheme, out OfficeColor color) {
        A.SolidFill? ownerFill = owner?.Descendants<A.SolidFill>().FirstOrDefault(fill => !fill.Ancestors<A.RunProperties>().Any());
        OfficeColor? resolved = OfficeOpenXmlThemeColorResolver.ResolveColor(ownerFill, scheme);
        if (resolved.HasValue) { color = resolved.Value; return true; }
        PowerPointSlideBackground background = slide.GetBackground();
        if (background.Kind == PowerPointSlideBackgroundKind.SolidColor && OfficeColor.TryParseHex(background.Color ?? string.Empty, out color)) return true;
        color = default;
        return false;
    }

    private static void InspectPowerPointAlternativeText(
        P.Slide root,
        string slideLocation,
        OfficeContentSafetyBuilder builder,
        IDictionary<string, PowerPointCleanupTarget>? targets) {
        int index = 0;
        foreach (P.NonVisualDrawingProperties properties in root.Descendants<P.NonVisualDrawingProperties>()) {
            foreach ((string Attribute, string Text) item in new[] {
                ("descr", properties.Description?.Value ?? string.Empty),
                ("title", properties.Title?.Value ?? string.Empty)
            }) {
                if (string.IsNullOrWhiteSpace(item.Text)) continue;
                OfficeContentSafetyFinding finding = builder.Add(
                    OfficeContentConcealmentKind.NonPrimaryContent,
                    OfficeContentSafetyRisk.Informational,
                    slideLocation + "/AlternativeText[" + (++index).ToString(CultureInfo.InvariantCulture) + "]/@" + item.Attribute,
                    "The text is stored as PowerPoint shape alternative text rather than painted slide text.",
                    item.Text,
                    OfficeContentCleanupCapability.RemoveText);
                if (targets != null) targets[finding.Id] = PowerPointCleanupTarget.ForAttribute(properties, item.Attribute);
            }
        }
    }

    private static void InspectPowerPointNotes(
        PowerPointSlide slide,
        string slideLocation,
        OfficeContentSafetyBuilder builder,
        IDictionary<string, PowerPointCleanupTarget>? targets) {
        P.NotesSlide? notes = slide.SlidePart.NotesSlidePart?.NotesSlide;
        string text = notes == null ? string.Empty : string.Concat(notes.Descendants<A.Text>().Select(item => item.Text));
        if (string.IsNullOrWhiteSpace(text)) return;
        OfficeContentSafetyFinding finding = builder.Add(
            OfficeContentConcealmentKind.NonPrimaryContent,
            OfficeContentSafetyRisk.Informational,
            slideLocation + "/Notes",
            "The text is stored in speaker notes rather than the painted slide canvas.",
            text,
            OfficeContentCleanupCapability.RemoveText);
        if (targets != null) targets[finding.Id] = PowerPointCleanupTarget.ForDescendantText(notes!);
    }

    private static void InspectPowerPointComments(
        PowerPointSlide slide,
        string slideLocation,
        OfficeContentSafetyBuilder builder,
        IDictionary<string, PowerPointCleanupTarget>? targets) {
        int index = 0;
        foreach (P.Comment comment in slide.SlidePart.SlideCommentsPart?.CommentList?.Elements<P.Comment>() ?? Enumerable.Empty<P.Comment>()) {
            string text = comment.Text?.Text ?? string.Empty;
            if (string.IsNullOrWhiteSpace(text)) continue;
            OfficeContentSafetyFinding finding = builder.Add(
                OfficeContentConcealmentKind.NonPrimaryContent,
                OfficeContentSafetyRisk.Informational,
                slideLocation + "/ClassicComment[" + (++index).ToString(CultureInfo.InvariantCulture) + "]",
                "The text is stored in a classic review comment rather than the slide canvas.",
                text,
                OfficeContentCleanupCapability.RemoveElement);
            if (targets != null) targets[finding.Id] = PowerPointCleanupTarget.ForElement(comment);
        }
        foreach (PowerPointCommentPart part in slide.SlidePart.Parts.Select(item => item.OpenXmlPart).OfType<PowerPointCommentPart>()) {
            foreach (P188.Comment comment in part.CommentList?.Elements<P188.Comment>() ?? Enumerable.Empty<P188.Comment>()) {
                string text = GetModernCommentText(comment);
                if (!string.IsNullOrWhiteSpace(text)) {
                    OfficeContentSafetyFinding finding = builder.Add(
                        OfficeContentConcealmentKind.NonPrimaryContent,
                        OfficeContentSafetyRisk.Informational,
                        slideLocation + "/ModernComment[" + (++index).ToString(CultureInfo.InvariantCulture) + "]",
                        "The text is stored in a modern review comment rather than the slide canvas.",
                        text,
                        OfficeContentCleanupCapability.RemoveElement);
                    if (targets != null) targets[finding.Id] = PowerPointCleanupTarget.ForElement(comment);
                }
                foreach (P188.CommentReply reply in comment.GetFirstChild<P188.CommentReplyList>()?.Elements<P188.CommentReply>() ?? Enumerable.Empty<P188.CommentReply>()) {
                    string replyText = GetModernCommentText(reply);
                    if (string.IsNullOrWhiteSpace(replyText)) continue;
                    OfficeContentSafetyFinding finding = builder.Add(
                        OfficeContentConcealmentKind.NonPrimaryContent,
                        OfficeContentSafetyRisk.Informational,
                        slideLocation + "/ModernCommentReply[" + (++index).ToString(CultureInfo.InvariantCulture) + "]",
                        "The text is stored in a modern review-comment reply rather than the slide canvas.",
                        replyText,
                        OfficeContentCleanupCapability.RemoveElement);
                    if (targets != null) targets[finding.Id] = PowerPointCleanupTarget.ForElement(reply);
                }
            }
        }
    }

    private sealed class PowerPointCleanupTarget : IEquatable<PowerPointCleanupTarget> {
        private readonly OpenXmlElement _element;
        private readonly PowerPointCleanupOperation _operation;
        private readonly string? _attribute;
        private readonly PowerPointShape? _shape;
        private readonly int? _offset;
        private readonly int? _length;
        private readonly string? _expected;
        private PowerPointCleanupTarget(OpenXmlElement element, PowerPointCleanupOperation operation, string? attribute = null, PowerPointShape? shape = null, int? offset = null, int? length = null, string? expected = null) {
            _element = element; _operation = operation; _attribute = attribute; _shape = shape; _offset = offset; _length = length; _expected = expected;
        }
        internal static PowerPointCleanupTarget ForElement(OpenXmlElement element) => new PowerPointCleanupTarget(element, PowerPointCleanupOperation.Element);
        internal static PowerPointCleanupTarget ForDescendantText(OpenXmlElement element) => new PowerPointCleanupTarget(element, PowerPointCleanupOperation.DescendantText);
        internal static PowerPointCleanupTarget ForAttribute(OpenXmlElement element, string attribute) => new PowerPointCleanupTarget(element, PowerPointCleanupOperation.Attribute, attribute);
        internal static PowerPointCleanupTarget ForShapeContent(PowerPointShape shape) => new PowerPointCleanupTarget(shape.Element, PowerPointCleanupOperation.ShapeContent, shape: shape);
        internal static PowerPointCleanupTarget ForTextRange(A.Text text, OfficeContentSafetyFinding finding) => new PowerPointCleanupTarget(
            text, PowerPointCleanupOperation.TextRange, offset: finding.SourceTextOffset, length: finding.SourceTextLength,
            expected: text.Text.Substring(finding.SourceTextOffset!.Value, finding.SourceTextLength!.Value));
        internal void Remove() {
            if (_operation == PowerPointCleanupOperation.TextRange && _element is A.Text text && _offset.HasValue && _length.HasValue) {
                string current = text.Text ?? string.Empty;
                if (_offset.Value > current.Length - _length.Value || !string.Equals(current.Substring(_offset.Value, _length.Value), _expected, StringComparison.Ordinal)) {
                    throw new InvalidOperationException("The selected Unicode text range no longer matches the inspected PowerPoint text node.");
                }
                text.Text = current.Remove(_offset.Value, _length.Value);
                return;
            }
            if (_operation == PowerPointCleanupOperation.Element) {
                if (_element is A.Run run && run.Text != null) run.Text.Text = string.Empty;
                else _element.Remove();
                return;
            }
            if (_operation == PowerPointCleanupOperation.ShapeContent && _shape != null) {
                if (_shape is PowerPointTextBox textBox) textBox.Text = string.Empty;
                else _shape.Remove();
                return;
            }
            if (_operation == PowerPointCleanupOperation.DescendantText) {
                foreach (A.Text textNode in _element.Descendants<A.Text>().ToArray()) textNode.Remove();
                return;
            }
            if (_element is P.NonVisualDrawingProperties properties) {
                if (string.Equals(_attribute, "descr", StringComparison.Ordinal)) properties.Description = null;
                else if (string.Equals(_attribute, "title", StringComparison.Ordinal)) properties.Title = null;
            }
        }
        public bool Equals(PowerPointCleanupTarget? other) => other != null && ReferenceEquals(_element, other._element) && _operation == other._operation && string.Equals(_attribute, other._attribute, StringComparison.Ordinal) && _offset == other._offset && _length == other._length;
        public override bool Equals(object? obj) => Equals(obj as PowerPointCleanupTarget);
        public override int GetHashCode() { unchecked { return (_element.GetHashCode() * 397) ^ ((int)_operation * 31) ^ (_attribute?.GetHashCode() ?? 0) ^ (_offset ?? 0); } }
        private enum PowerPointCleanupOperation { Element, DescendantText, Attribute, ShapeContent, TextRange }
    }
}
