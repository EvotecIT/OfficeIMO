using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.ContentSafety;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgDrawingReader {
    private const long MaximumSvgContentSafetyDocumentWorkBytes = 64L * 1024L * 1024L;

    private static OfficeContentSafetyReport InspectSvgContentSafetyDocument(
        SvgContentSafetyDocument document,
        byte[] sourceBytes,
        OfficeContentSafetyOptions options,
        OfficeSvgDrawingReaderOptions? readerOptions,
        IDictionary<string, SvgContentSafetyTarget>? targets) {
        var builder = new OfficeContentSafetyBuilder("SVG", options);
        XElement computedRoot = new XElement(document.Root);
        int unsupported = 0;
        if (!ApplySvgStylesheets(computedRoot, ref unsupported, requireCompleteCss: true)) {
            throw new InvalidDataException("The SVG stylesheet or inline-style cascade uses unsupported constructs or exceeds the bounded CSS inspection limits.");
        }

        IReadOnlyList<SvgContentSafetyCandidate> candidates = CreateSvgContentSafetyCandidates(
            document.Root,
            computedRoot,
            document.MaximumElements);
        var byComputedText = candidates.ToDictionary(item => item.ComputedText, item => item);
        var observer = new SvgContentSafetyTextObserver(byComputedText);
        CollectSvgContentSafetyTextRuns(computedRoot, document, readerOptions, observer, ref unsupported);
        ISet<string> reusableTextIds = CollectSvgReusableTextReferencedIds(document.Root);
        bool hasDynamicRendering = HasSvgDynamicRendering(document.Root);

        int requestedComparisons = document.MaximumVisualComparisons;
        int pixelFundedComparisons = (int)Math.Min(
            requestedComparisons,
            Math.Max(0L, Math.Min(int.MaxValue, document.MaximumVisualPixels - 1L)));
        long estimatedDocumentPassBytes = Math.Max(
            1L,
            Math.Min(long.MaxValue / 2L, sourceBytes.LongLength) * 2L +
            (long)document.Root.DescendantsAndSelf().Count() * 64L);
        int workFundedComparisons = (int)Math.Max(
            0L,
            Math.Min(int.MaxValue, MaximumSvgContentSafetyDocumentWorkBytes / estimatedDocumentPassBytes - 1L));
        int maximumComparisons = Math.Min(pixelFundedComparisons, workFundedComparisons);
        long maximumRasterPixels = maximumComparisons == 0
            ? 0L
            : Math.Max(1L, document.MaximumVisualPixels / (maximumComparisons + 1L));
        OfficeRasterImage? baseline = null;
        int baselineUnsupported = 0;
        bool baselineRendered = maximumComparisons > 0 && TryRenderSvgContentSafety(
            sourceBytes,
            readerOptions,
            Math.Min(document.MaximumViewportPixels, maximumRasterPixels),
            out baseline,
            out baselineUnsupported);
        if (requestedComparisons == 0) builder.AddDiagnostic("SVG visual comparison was disabled by the caller; structural inspection still covered every bounded native text candidate.");
        else if (maximumComparisons == 0) builder.AddDiagnostic("SVG visual comparison was disabled because the cumulative pixel or document-transformation work budget cannot fund a baseline and comparison render.");
        else if (!baselineRendered) builder.AddDiagnostic("SVG visual comparison was unavailable because the bounded SVG drawing surface could not be rendered.");
        else if (baselineUnsupported > 0) builder.AddDiagnostic(
            "SVG visual comparison was limited because the drawing reader reported " +
            baselineUnsupported.ToString(CultureInfo.InvariantCulture) + " unsupported declarations or elements.");
        if (unsupported > 0) builder.AddDiagnostic(
            "SVG cascade or text-layout inspection encountered " + unsupported.ToString(CultureInfo.InvariantCulture) +
            " declarations or constructs outside the bounded native subset.");
        if (maximumComparisons < requestedComparisons && maximumComparisons > 0) builder.AddDiagnostic(
            "SVG visual comparisons were reduced to " + maximumComparisons.ToString(CultureInfo.InvariantCulture) +
            " by the cumulative rendered-pixel and document-transformation work budgets.");

        int comparisons = 0;
        bool comparisonLimitReached = false;
        foreach (SvgContentSafetyCandidate candidate in candidates) {
            string text = candidate.SourceText.Value;
            if (text.Length == 0 || string.IsNullOrWhiteSpace(text)) continue;
            bool contextDependent = TryDescribeSvgContextDependentText(
                candidate.ComputedElement,
                reusableTextIds,
                hasDynamicRendering,
                out string contextEvidence);

            if (TryClassifySvgNonPrimary(
                    candidate,
                    out SvgContentSafetyConcealment nonPrimary,
                    out OfficeContentCleanupCapability nonPrimaryCleanup)) {
                if (options.IncludeNonPrimaryContent) {
                    if (contextDependent) {
                        nonPrimary = new SvgContentSafetyConcealment(
                            nonPrimary.Kind,
                            nonPrimary.Evidence + " " + contextEvidence,
                            nonPrimary.Risk);
                    }
                    AddSvgContentSafetyFinding(
                        builder,
                        targets,
                        candidate,
                        nonPrimary,
                        contextDependent ? OfficeContentCleanupCapability.ReportOnly : nonPrimaryCleanup);
                }
                continue;
            }

            SvgContentSafetyConcealment? concealment = ClassifySvgStructuralConcealment(candidate, document, options);
            bool visualConcealment = false;
            bool hostBackdropDependent = false;
            if (!concealment.HasValue && baselineRendered) {
                if (comparisons < maximumComparisons) {
                    comparisons++;
                    if (TryCompareSvgCandidate(
                            document,
                            candidate,
                            readerOptions,
                            maximumRasterPixels,
                            baseline!,
                            baselineUnsupported,
                            out SvgVisualComparison comparison)) {
                        concealment = ClassifySvgVisualConcealment(candidate, document, baseline!, comparison, options);
                        visualConcealment = concealment.HasValue;
                        hostBackdropDependent = concealment.HasValue &&
                            comparison.HasTransparentBackdrop &&
                            concealment.Value.Kind is OfficeContentConcealmentKind.LowContrastText or OfficeContentConcealmentKind.Other;
                        if (hostBackdropDependent) {
                            SvgContentSafetyConcealment visualFinding = concealment!.Value;
                            concealment = new SvgContentSafetyConcealment(
                                visualFinding.Kind,
                                visualFinding.Evidence +
                                " The candidate backdrop contains transparent pixels, so visibility depends on the host background and cleanup is report-only.",
                                visualFinding.Risk);
                        }
                    }
                } else {
                    comparisonLimitReached = true;
                }
            }

            if (concealment.HasValue) {
                if (contextDependent) {
                    concealment = new SvgContentSafetyConcealment(
                        concealment.Value.Kind,
                        concealment.Value.Evidence + " " + contextEvidence,
                        concealment.Value.Risk);
                }
                AddSvgContentSafetyFinding(
                    builder,
                    targets,
                    candidate,
                    concealment.Value,
                    contextDependent || visualConcealment && (baselineUnsupported > 0 || hostBackdropDependent)
                        ? OfficeContentCleanupCapability.ReportOnly
                        : OfficeContentCleanupCapability.RemoveText);
            } else if (contextDependent) {
                if (options.IncludeNonPrimaryContent) {
                    AddSvgContentSafetyFinding(
                        builder,
                        targets,
                        candidate,
                        new SvgContentSafetyConcealment(
                            OfficeContentConcealmentKind.NonPrimaryContent,
                            contextEvidence,
                            OfficeContentSafetyRisk.Informational),
                        OfficeContentCleanupCapability.ReportOnly);
                } else {
                    IReadOnlyList<OfficeContentSafetyFinding> unicode = builder.InspectVisibleText(
                        candidate.Location,
                        text,
                        OfficeContentCleanupCapability.ReportOnly);
                    RegisterSvgTargets(targets, candidate, unicode);
                }
            } else {
                IReadOnlyList<OfficeContentSafetyFinding> unicode = builder.InspectVisibleText(
                    candidate.Location,
                    text,
                    OfficeContentCleanupCapability.RemoveText);
                RegisterSvgTargets(targets, candidate, unicode);
            }
        }

        if (comparisonLimitReached) {
            builder.AddDiagnostic(
                "SVG paint-order and background comparisons were capped at " +
                maximumComparisons.ToString(CultureInfo.InvariantCulture) +
                " text nodes; property, opacity, font, and geometry inspection still covered every bounded candidate.");
        }
        return builder.Build();
    }

    private static void AddSvgContentSafetyFinding(
        OfficeContentSafetyBuilder builder,
        IDictionary<string, SvgContentSafetyTarget>? targets,
        SvgContentSafetyCandidate candidate,
        SvgContentSafetyConcealment concealment,
        OfficeContentCleanupCapability cleanupCapability) {
        OfficeContentSafetyFinding finding = builder.Add(
            concealment.Kind,
            concealment.Risk,
            candidate.Location,
            concealment.Evidence,
            candidate.SourceText.Value,
            cleanupCapability,
            inspectTextIntegrityEvidence: false);
        if (targets != null && cleanupCapability != OfficeContentCleanupCapability.ReportOnly) {
            targets[finding.Id] = candidate.Target;
        }
        IReadOnlyList<OfficeContentSafetyFinding> unicode = builder.InspectChargedTextIntegrity(
            candidate.Location,
            candidate.SourceText.Value,
            cleanupCapability);
        RegisterSvgTargets(targets, candidate, unicode);
    }

    private static void RegisterSvgTargets(
        IDictionary<string, SvgContentSafetyTarget>? targets,
        SvgContentSafetyCandidate candidate,
        IEnumerable<OfficeContentSafetyFinding> findings) {
        if (targets == null) return;
        foreach (OfficeContentSafetyFinding finding in findings) {
            if (finding.CleanupCapability != OfficeContentCleanupCapability.ReportOnly) {
                targets[finding.Id] = candidate.Target;
            }
        }
    }

    private static IReadOnlyList<SvgContentSafetyCandidate> CreateSvgContentSafetyCandidates(
        XElement sourceRoot,
        XElement computedRoot,
        int maximumCandidates) {
        XElement[] sourceElements = sourceRoot.DescendantsAndSelf().ToArray();
        XElement[] computedElements = computedRoot.DescendantsAndSelf().ToArray();
        if (sourceElements.Length != computedElements.Length) {
            throw new InvalidDataException("The computed SVG tree no longer matches the source tree.");
        }

        var candidates = new List<SvgContentSafetyCandidate>();
        for (int elementIndex = 0; elementIndex < sourceElements.Length; elementIndex++) {
            XElement sourceElement = sourceElements[elementIndex];
            XElement computedElement = computedElements[elementIndex];
            bool isNativeSvg = IsNativeSvgElement(sourceElement, sourceRoot.Name.Namespace);
            if (!IsSvgContentSafetyTextContainer(sourceElement, isNativeSvg)) continue;
            using IEnumerator<XText> sourceText = sourceElement.Nodes().OfType<XText>().GetEnumerator();
            using IEnumerator<XText> computedText = computedElement.Nodes().OfType<XText>().GetEnumerator();
            int textIndex = 0;
            while (true) {
                bool hasSource = sourceText.MoveNext();
                bool hasComputed = computedText.MoveNext();
                if (hasSource != hasComputed) {
                    throw new InvalidDataException("The computed SVG text tree no longer matches the source tree.");
                }
                if (!hasSource) break;
                if (candidates.Count >= maximumCandidates) {
                    throw new InvalidDataException("The SVG exceeds the bounded content-safety text-node limit.");
                }
                candidates.Add(new SvgContentSafetyCandidate(
                    sourceText.Current,
                    computedText.Current,
                    computedElement,
                    elementIndex,
                    textIndex,
                    BuildSvgContentSafetyLocation(sourceElement, textIndex),
                    isNativeSvg));
                textIndex++;
            }
        }
        return candidates.AsReadOnly();
    }

    private static bool IsSvgContentSafetyTextContainer(XElement element, bool isNativeSvg) {
        if (!isNativeSvg) return element.Nodes().OfType<XText>().Any();
        return element.Nodes().OfType<XText>().Any(text => !string.IsNullOrWhiteSpace(text.Value));
    }

    private static string BuildSvgContentSafetyLocation(XElement element, int textIndex) {
        var segments = new Stack<string>();
        for (XElement? current = element; current != null; current = current.Parent) {
            int ordinal = current.ElementsBeforeSelf()
                .Count(sibling => sibling.Name.LocalName.Equals(current.Name.LocalName, StringComparison.OrdinalIgnoreCase)) + 1;
            segments.Push(current.Name.LocalName + "[" + ordinal.ToString(CultureInfo.InvariantCulture) + "]");
        }
        return "SVG/" + string.Join("/", segments) + "/text()[" + (textIndex + 1).ToString(CultureInfo.InvariantCulture) + "]";
    }

    private static void CollectSvgContentSafetyTextRuns(
        XElement root,
        SvgContentSafetyDocument document,
        OfficeSvgDrawingReaderOptions? readerOptions,
        SvgContentSafetyTextObserver observer,
        ref int unsupported) {
        SvgDefinitionRegistry definitions = SvgDefinitionRegistry.Create(root);
        var paintServers = new SvgPaintServerRegistry(definitions);
        var references = new SvgElementReferenceRegistry(definitions, readerOptions?.ForeignObjectRenderer);
        SvgPaintContext defaults = SvgPaintContext.Default;
        defaults.DashPercentageReference = NormalizedSvgDiagonal(document.ViewWidth, document.ViewHeight);
        SvgPaintContext rootStyle = ResolvePaintContext(root, defaults, paintServers, ref unsupported);
        OfficeTransform rootTransform = ResolveTransform(root, OfficeTransform.Identity, document.ViewX, document.ViewY, ref unsupported);
        var fonts = new OfficeFontFaceCollection();
        fonts.AddRange(readerOptions?.Fonts);
        CollectSvgContentSafetyTextElements(
            root,
            rootStyle,
            rootTransform,
            paintServers,
            references,
            fonts,
            document,
            observer,
            document.ViewX,
            document.ViewY,
            document.ViewWidth,
            document.ViewHeight,
            depth: 0,
            ref unsupported);
    }

    private static void CollectSvgContentSafetyTextElements(
        XElement parent,
        SvgPaintContext inheritedStyle,
        OfficeTransform inheritedTransform,
        SvgPaintServerRegistry paintServers,
        SvgElementReferenceRegistry references,
        OfficeFontFaceCollection fonts,
        SvgContentSafetyDocument document,
        SvgContentSafetyTextObserver observer,
        double viewX,
        double viewY,
        double viewWidth,
        double viewHeight,
        int depth,
        ref int unsupported) {
        if (depth > MaximumSvgNestingDepth) {
            throw new InvalidDataException("The SVG exceeds the bounded element-nesting limit.");
        }
        foreach (XElement child in parent.Elements()) {
            if (!IsNativeSvgElement(child, document.Root.Name.Namespace)) continue;
            SvgPaintContext style = ResolvePaintContext(child, inheritedStyle, paintServers, ref unsupported);
            OfficeTransform transform = ResolveTransform(child, inheritedTransform, viewX, viewY, ref unsupported);
            double childViewX = viewX;
            double childViewY = viewY;
            double childViewWidth = viewWidth;
            double childViewHeight = viewHeight;
            if (child.Name.LocalName.Equals("svg", StringComparison.OrdinalIgnoreCase) &&
                !TryResolveSvgContentSafetyNestedViewport(
                    child,
                    transform,
                    viewX,
                    viewY,
                    viewWidth,
                    viewHeight,
                    document,
                    out transform,
                    out childViewX,
                    out childViewY,
                    out childViewWidth,
                    out childViewHeight)) {
                throw new InvalidDataException("The SVG contains nested viewport geometry outside the bounded native subset.");
            }
            if (child.Name.LocalName.Equals("text", StringComparison.OrdinalIgnoreCase)) {
                var runs = new List<SvgTextRun>();
                var paths = new List<SvgTextPathLayout>();
                var cursor = new SvgTextCursor { Chunk = -1 };
                bool preserve = string.Equals(child.Attribute(XNamespace.Xml + "space")?.Value, "preserve", StringComparison.OrdinalIgnoreCase);
                AddTextElementRuns(
                    child,
                    style,
                    paintServers,
                    references,
                    fonts,
                    transform,
                    preserve,
                    resolveElement: false,
                    childViewX,
                    childViewY,
                    childViewWidth,
                    childViewHeight,
                    runs,
                    paths,
                    observer,
                    0D,
                    0D,
                    inheritedPositioning: null,
                    depth: 0,
                    ref cursor,
                    ref unsupported);
                if (cursor.LimitReported) {
                    throw new InvalidDataException("The SVG exceeds the bounded text-run inspection limit.");
                }
                ApplyTextAnchors(runs);
                ApplyTextPaths(runs, paths, references, childViewX, childViewY, ref unsupported);
                observer.FinalizeRuns(runs);
                continue;
            }
            CollectSvgContentSafetyTextElements(
                child,
                style,
                transform,
                paintServers,
                references,
                fonts,
                document,
                observer,
                childViewX,
                childViewY,
                childViewWidth,
                childViewHeight,
                depth + 1,
                ref unsupported);
        }
    }

    private static SvgContentSafetyConcealment? ClassifySvgStructuralConcealment(
        SvgContentSafetyCandidate candidate,
        SvgContentSafetyDocument document,
        OfficeContentSafetyOptions options) {
        if (TryFindSvgHiddenProperty(candidate.ComputedElement, out string hiddenEvidence)) {
            return new SvgContentSafetyConcealment(OfficeContentConcealmentKind.HiddenByProperty, hiddenEvidence);
        }
        if (TryFindEmptySvgClip(candidate.ComputedElement, document, out string clipEvidence)) {
            return new SvgContentSafetyConcealment(OfficeContentConcealmentKind.ClippedContent, clipEvidence);
        }
        if (candidate.HasStyle) {
            SvgPaintContext style = candidate.Style;
            if (style.Opacity <= 0.01D || !HasVisibleSvgTextPaint(style)) {
                return new SvgContentSafetyConcealment(
                    OfficeContentConcealmentKind.TransparentText,
                    "Computed SVG opacity and text paint leave no visible fill or stroke contribution.");
            }
            double size = style.FontSize;
            bool tiny = size <= options.MaximumTinyFontSizePoints;
            if (!tiny && TryFindSvgTinyFont(candidate.ComputedElement, options, out double authoredSize)) {
                size = authoredSize;
                tiny = true;
            }
            if (tiny) {
                return new SvgContentSafetyConcealment(
                    OfficeContentConcealmentKind.TinyText,
                    "Computed SVG font size is " + size.ToString("0.###", CultureInfo.InvariantCulture) + " user units.");
            }
        }
        if (candidate.HasBounds) {
            double width = candidate.Right - candidate.Left;
            double height = candidate.Bottom - candidate.Top;
            if (width <= 0.01D || height <= 0.01D) {
                return new SvgContentSafetyConcealment(
                    OfficeContentConcealmentKind.ZeroDimension,
                    "Resolved SVG text geometry has zero or near-zero painted bounds.");
            }
            if (candidate.MaximumEffectiveFontSize > 0D &&
                candidate.MaximumEffectiveFontSize <= options.MaximumTinyFontSizePoints) {
                return new SvgContentSafetyConcealment(
                    OfficeContentConcealmentKind.TinyText,
                    "Resolved SVG transforms reduce the effective font size to " +
                    candidate.MaximumEffectiveFontSize.ToString("0.###", CultureInfo.InvariantCulture) + " user units.");
            }
            if (candidate.Right <= 0D || candidate.Bottom <= 0D ||
                candidate.Left >= document.ViewWidth || candidate.Top >= document.ViewHeight) {
                return new SvgContentSafetyConcealment(
                    OfficeContentConcealmentKind.OffCanvas,
                    "Resolved SVG text bounds fall completely outside the ordinary view box.");
            }
        }
        return null;
    }

    private static bool TryFindEmptySvgClip(
        XElement element,
        SvgContentSafetyDocument document,
        out string evidence) {
        foreach (XElement current in element.AncestorsAndSelf()) {
            string? value = ReadPresentationProperty(current, "clip-path")?.Trim();
            if (string.IsNullOrWhiteSpace(value) || string.Equals(value, "none", StringComparison.OrdinalIgnoreCase)) continue;
            string normalized = value!;
            if (!normalized.StartsWith("url(", StringComparison.OrdinalIgnoreCase) || !normalized.EndsWith(")", StringComparison.Ordinal)) continue;
            string reference = normalized.Substring(4, normalized.Length - 5).Trim().Trim('\'', '"');
            if (!reference.StartsWith("#", StringComparison.Ordinal) || reference.Length == 1) continue;
            string id = reference.Substring(1);
            XElement? clip = document.Root.DescendantsAndSelf().FirstOrDefault(candidate =>
                IsNativeSvgElement(candidate, document.Root.Name.Namespace) &&
                candidate.Name.LocalName.Equals("clipPath", StringComparison.OrdinalIgnoreCase) &&
                string.Equals(candidate.Attribute("id")?.Value, id, StringComparison.Ordinal));
            if (clip == null) continue;
            XElement[] geometry = clip.Elements().Where(child =>
                IsNativeSvgElement(child, document.Root.Name.Namespace) &&
                child.Name.LocalName is not "title" and not "desc" and not "metadata").ToArray();
            if (geometry.Length == 0 || geometry.All(child => IsEmptySvgClipGeometry(child, document.ViewWidth, document.ViewHeight))) {
                evidence = "The computed SVG clip references empty or zero-area geometry.";
                return true;
            }
        }
        evidence = string.Empty;
        return false;
    }

    private static bool TryFindSvgHiddenProperty(XElement element, out string evidence) {
        string visibility = "visible";
        foreach (XElement current in element.AncestorsAndSelf().Reverse()) {
            string? display = ReadPresentationProperty(current, "display")?.Trim();
            if (string.Equals(display, "none", StringComparison.OrdinalIgnoreCase)) {
                evidence = "Computed SVG display is none on " + current.Name.LocalName + ".";
                return true;
            }
            string? candidate = ReadPresentationProperty(current, "visibility")?.Trim();
            if (string.IsNullOrWhiteSpace(candidate) || string.Equals(candidate, "inherit", StringComparison.OrdinalIgnoreCase)) continue;
            if (string.Equals(candidate, "initial", StringComparison.OrdinalIgnoreCase)) visibility = "visible";
            else if (!string.Equals(candidate, "unset", StringComparison.OrdinalIgnoreCase)) visibility = candidate!;
        }
        if (string.Equals(visibility, "hidden", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(visibility, "collapse", StringComparison.OrdinalIgnoreCase)) {
            evidence = "Computed SVG visibility is " + visibility + ".";
            return true;
        }
        evidence = string.Empty;
        return false;
    }

    private static bool TryFindSvgTinyFont(
        XElement element,
        OfficeContentSafetyOptions options,
        out double fontSize) {
        fontSize = double.PositiveInfinity;
        foreach (XElement current in element.AncestorsAndSelf()) {
            string? value = ReadPresentationProperty(current, "font-size");
            if (string.IsNullOrWhiteSpace(value) || string.Equals(value!.Trim(), "inherit", StringComparison.OrdinalIgnoreCase)) continue;
            if (TrySvgLength(value, out double parsed)) {
                fontSize = parsed;
                return parsed <= options.MaximumTinyFontSizePoints;
            }
            return false;
        }
        return false;
    }

    private static bool HasVisibleSvgTextPaint(SvgPaintContext style) {
        double opacity = Math.Max(0D, Math.Min(1D, style.Opacity));
        bool hasFillPaint = style.Fill.HasValue || style.FillGradient != null || style.FillRadialGradient != null ||
            style.FillDeferredGradient != null || style.FillPattern != null;
        bool fill = hasFillPaint && style.FillOpacity * opacity > 0.01D &&
            (!style.Fill.HasValue || style.Fill.Value.A * style.FillOpacity * opacity > 3D);
        bool hasStrokePaint = style.Stroke.HasValue || style.StrokeGradient != null || style.StrokeRadialGradient != null ||
            style.StrokeDeferredGradient != null || style.StrokePattern != null;
        bool stroke = style.StrokeWidth > 0D && hasStrokePaint && style.StrokeOpacity * opacity > 0.01D &&
            (!style.Stroke.HasValue || style.Stroke.Value.A * style.StrokeOpacity * opacity > 3D);
        return fill || stroke;
    }

    private static SvgContentSafetyConcealment? ClassifySvgVisualConcealment(
        SvgContentSafetyCandidate candidate,
        SvgContentSafetyDocument document,
        OfficeRasterImage baseline,
        SvgVisualComparison comparison,
        OfficeContentSafetyOptions options) {
        if (comparison.ChangedPixels > 0 && comparison.MaximumContrastRatio + 0.000001D < options.MinimumVisibleContrastRatio) {
            return new SvgContentSafetyConcealment(
                OfficeContentConcealmentKind.LowContrastText,
                "Paint-order-resolved SVG text pixels have maximum contrast ratio " +
                comparison.MaximumContrastRatio.ToString("0.###", CultureInfo.InvariantCulture) + ".");
        }
        if (comparison.ChangedPixels != 0) return null;

        if (HasActiveSvgClip(candidate.ComputedElement)) {
            return new SvgContentSafetyConcealment(
                OfficeContentConcealmentKind.ClippedContent,
                "The active computed SVG clip prevents the text from contributing any painted pixels.");
        }
        if (TryResolveSvgPaintContrast(candidate, document, baseline, out double contrast) &&
            contrast + 0.000001D < options.MinimumVisibleContrastRatio) {
            return new SvgContentSafetyConcealment(
                OfficeContentConcealmentKind.LowContrastText,
                "Resolved SVG text paint against the paint-order-composited background has maximum contrast ratio " +
                contrast.ToString("0.###", CultureInfo.InvariantCulture) + ".");
        }
        return new SvgContentSafetyConcealment(
            OfficeContentConcealmentKind.Other,
            "The bounded native SVG renderer found no painted pixel contribution after clipping, compositing, and document paint order.");
    }

    private static bool HasActiveSvgClip(XElement element) => element.AncestorsAndSelf().Any(current => {
        string? clip = ReadPresentationProperty(current, "clip-path")?.Trim();
        return !string.IsNullOrWhiteSpace(clip) && !string.Equals(clip, "none", StringComparison.OrdinalIgnoreCase);
    });

    private static bool TryCompareSvgCandidate(
        SvgContentSafetyDocument document,
        SvgContentSafetyCandidate candidate,
        OfficeSvgDrawingReaderOptions? readerOptions,
        double maximumRasterPixels,
        OfficeRasterImage baseline,
        int baselineUnsupported,
        out SvgVisualComparison comparison) {
        XDocument variant = new XDocument(document.Document);
        XElement[] elements = variant.Root!.DescendantsAndSelf().ToArray();
        if (candidate.ElementIndex < 0 || candidate.ElementIndex >= elements.Length) {
            comparison = default;
            return false;
        }
        XText[] textNodes = elements[candidate.ElementIndex].Nodes().OfType<XText>().ToArray();
        if (candidate.TextNodeIndex < 0 || candidate.TextNodeIndex >= textNodes.Length) {
            comparison = default;
            return false;
        }
        XText target = textNodes[candidate.TextNodeIndex];
        var suppressed = new XElement(
            elements[candidate.ElementIndex].Name.Namespace + "tspan",
            new XAttribute(
                "style",
                "display:inline!important;opacity:0!important;font-family:inherit!important;font-size:inherit!important;" +
                "font-style:inherit!important;font-weight:inherit!important;line-height:inherit!important;" +
                "text-anchor:inherit!important;writing-mode:inherit!important"),
            target.Value);
        target.ReplaceWith(suppressed);
        byte[] bytes = SerializeSvgContentSafetyDocument(variant, MaximumInputBytes);
        if (!TryRenderSvgContentSafety(
                bytes,
                readerOptions,
                Math.Min(document.MaximumViewportPixels, maximumRasterPixels),
                out OfficeRasterImage? without,
                out int unsupported) ||
            unsupported != baselineUnsupported || without!.Width != baseline.Width || without.Height != baseline.Height) {
            throw new InvalidDataException(
                "The bounded SVG comparison variant could not be rendered consistently with the inspected source (baseline unsupported " +
                baselineUnsupported.ToString(CultureInfo.InvariantCulture) + ", variant unsupported " +
                unsupported.ToString(CultureInfo.InvariantCulture) + ").");
        }
        comparison = CompareSvgRasters(baseline, without);
        if (!comparison.HasTransparentBackdrop &&
            HasTransparentSvgBackdrop(candidate, document, without)) {
            comparison = new SvgVisualComparison(
                comparison.ChangedPixels,
                comparison.MaximumContrastRatio,
                hasTransparentBackdrop: true);
        }
        return true;
    }

    private static bool TryRenderSvgContentSafety(
        byte[] svgBytes,
        OfficeSvgDrawingReaderOptions? readerOptions,
        double maximumRasterPixels,
        out OfficeRasterImage? image,
        out int unsupported) {
        image = null;
        if (!TryRead(svgBytes, readerOptions, out OfficeDrawing? drawing, out unsupported) || drawing == null) return false;
        const long visualPixelLimit = 4_000_000L;
        long maximumPixels = (long)Math.Max(1D, Math.Min(visualPixelLimit, maximumRasterPixels));
        double sourcePixels = drawing.Width * drawing.Height;
        double scale = sourcePixels > maximumPixels ? Math.Sqrt(maximumPixels / sourcePixels) : 1D;
        if ((long)Math.Ceiling(drawing.Width * scale) * (long)Math.Ceiling(drawing.Height * scale) > maximumPixels) {
            double lower = 0D;
            double upper = scale;
            for (int iteration = 0; iteration < 64; iteration++) {
                double midpoint = (lower + upper) / 2D;
                long pixels = (long)Math.Ceiling(drawing.Width * midpoint) *
                    (long)Math.Ceiling(drawing.Height * midpoint);
                if (pixels <= maximumPixels) lower = midpoint;
                else upper = midpoint;
            }
            scale = lower;
        }
        if (scale <= 0D) return false;
        image = OfficeDrawingRasterRenderer.Render(drawing, new OfficeDrawingRasterRenderOptions {
            Scale = scale,
            Background = OfficeColor.Transparent,
            MaximumRasterPixels = maximumPixels
        });
        return true;
    }

    private static SvgVisualComparison CompareSvgRasters(OfficeRasterImage withText, OfficeRasterImage withoutText) {
        byte[] withPixels = withText.PixelBuffer;
        byte[] withoutPixels = withoutText.PixelBuffer;
        int changed = 0;
        double maximumContrast = 1D;
        bool hasTransparentBackdrop = false;
        for (int offset = 0; offset < withPixels.Length; offset += 4) {
            int delta = Math.Abs(withPixels[offset] - withoutPixels[offset]) +
                Math.Abs(withPixels[offset + 1] - withoutPixels[offset + 1]) +
                Math.Abs(withPixels[offset + 2] - withoutPixels[offset + 2]) +
                Math.Abs(withPixels[offset + 3] - withoutPixels[offset + 3]);
            if (delta <= 3) continue;
            changed++;
            OfficeColor foreground = OfficeColor.FromRgba(withPixels[offset], withPixels[offset + 1], withPixels[offset + 2], withPixels[offset + 3]);
            OfficeColor background = OfficeColor.FromRgba(withoutPixels[offset], withoutPixels[offset + 1], withoutPixels[offset + 2], withoutPixels[offset + 3]);
            hasTransparentBackdrop |= background.A < byte.MaxValue;
            OfficeColor foregroundOnBlack = CompositeSvgPaint(foreground, OfficeColor.Black);
            OfficeColor backgroundOnBlack = CompositeSvgPaint(background, OfficeColor.Black);
            OfficeColor foregroundOnWhite = CompositeSvgPaint(foreground, OfficeColor.White);
            OfficeColor backgroundOnWhite = CompositeSvgPaint(background, OfficeColor.White);
            maximumContrast = Math.Max(maximumContrast, Math.Max(
                OfficeColorContrast.ContrastRatio(foregroundOnBlack, backgroundOnBlack),
                OfficeColorContrast.ContrastRatio(foregroundOnWhite, backgroundOnWhite)));
        }
        return new SvgVisualComparison(changed, maximumContrast, hasTransparentBackdrop);
    }

    private static bool HasTransparentSvgBackdrop(
        SvgContentSafetyCandidate candidate,
        SvgContentSafetyDocument document,
        OfficeRasterImage background) {
        if (!candidate.HasBounds) return false;
        double scaleX = background.Width / document.ViewWidth;
        double scaleY = background.Height / document.ViewHeight;
        int left = Math.Max(0, (int)Math.Floor(candidate.Left * scaleX));
        int top = Math.Max(0, (int)Math.Floor(candidate.Top * scaleY));
        int right = Math.Min(background.Width, (int)Math.Ceiling(candidate.Right * scaleX));
        int bottom = Math.Min(background.Height, (int)Math.Ceiling(candidate.Bottom * scaleY));
        if (left >= right || top >= bottom) return false;
        for (int y = top; y < bottom; y++) {
            for (int x = left; x < right; x++) {
                if (background.GetPixel(x, y).A < byte.MaxValue) return true;
            }
        }
        return false;
    }

    private static bool TryResolveSvgPaintContrast(
        SvgContentSafetyCandidate candidate,
        SvgContentSafetyDocument document,
        OfficeRasterImage background,
        out double maximumContrast) {
        maximumContrast = 1D;
        if (!candidate.HasStyle || !candidate.HasBounds) return false;
        SvgPaintContext style = candidate.Style;
        OfficeColor? paint = style.Fill ?? style.Stroke;
        if (!paint.HasValue) return false;
        double paintOpacity = style.Fill.HasValue ? style.FillOpacity * style.Opacity : style.StrokeOpacity * style.Opacity;
        byte alpha = (byte)Math.Max(0D, Math.Min(255D, Math.Round(paint.Value.A * paintOpacity)));
        OfficeColor authored = OfficeColor.FromRgba(paint.Value.R, paint.Value.G, paint.Value.B, alpha);
        double scaleX = background.Width / document.ViewWidth;
        double scaleY = background.Height / document.ViewHeight;
        int left = Math.Max(0, (int)Math.Floor(candidate.Left * scaleX));
        int top = Math.Max(0, (int)Math.Floor(candidate.Top * scaleY));
        int right = Math.Min(background.Width, (int)Math.Ceiling(candidate.Right * scaleX));
        int bottom = Math.Min(background.Height, (int)Math.Ceiling(candidate.Bottom * scaleY));
        if (left >= right || top >= bottom) return false;
        bool sampled = false;
        int stepX = Math.Max(1, (right - left) / 32);
        int stepY = Math.Max(1, (bottom - top) / 32);
        for (int y = top; y < bottom; y += stepY) {
            for (int x = left; x < right; x += stepX) {
                OfficeColor backdrop = background.GetPixel(x, y);
                OfficeColor composited = CompositeSvgPaint(authored, backdrop);
                maximumContrast = Math.Max(maximumContrast, OfficeColorContrast.ContrastRatio(composited, backdrop));
                sampled = true;
            }
        }
        return sampled;
    }

    private static OfficeColor CompositeSvgPaint(OfficeColor foreground, OfficeColor background) {
        double alpha = foreground.A / 255D;
        double inverse = 1D - alpha;
        return OfficeColor.FromRgba(
            (byte)Math.Round(foreground.R * alpha + background.R * inverse),
            (byte)Math.Round(foreground.G * alpha + background.G * inverse),
            (byte)Math.Round(foreground.B * alpha + background.B * inverse),
            255);
    }
}
