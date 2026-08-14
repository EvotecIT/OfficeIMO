using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Drawing;

/// <summary>
/// Reads a bounded subset of SVG into the shared dependency-free drawing scene.
/// </summary>
public static partial class OfficeSvgDrawingReader {
    private const int MaximumInputBytes = 8 * 1024 * 1024;
    private const int MaximumSvgNestingDepth = 128;
    private const int MaximumSvgPathCommands = 20000;
    private const double MaximumSvgTransformCoefficient = 1024D;
    private const double MaximumSvgTransformOffset = 1000000D;

    /// <summary>Attempts to interpret supported SVG vector primitives as a shared drawing.</summary>
    public static bool TryRead(byte[]? bytes, out OfficeDrawing? drawing) =>
        TryRead(bytes, options: null, out drawing, out _);

    /// <summary>
    /// Attempts to interpret supported SVG vector primitives as a shared drawing and reports the
    /// number of elements or declarations that required omission or fallback.
    /// </summary>
    public static bool TryRead(byte[]? bytes, out OfficeDrawing? drawing, out int unsupportedFeatureCount) =>
        TryRead(bytes, options: null, out drawing, out unsupportedFeatureCount);

    /// <summary>Attempts to interpret supported SVG vector primitives using explicit bounded import options.</summary>
    public static bool TryRead(byte[]? bytes, OfficeSvgDrawingReaderOptions? options, out OfficeDrawing? drawing) =>
        TryRead(bytes, options, out drawing, out _);

    /// <summary>
    /// Attempts to interpret supported SVG vector primitives using explicit bounded import options and reports the
    /// number of elements or declarations that required omission or fallback.
    /// </summary>
    public static bool TryRead(
        byte[]? bytes,
        OfficeSvgDrawingReaderOptions? options,
        out OfficeDrawing? drawing,
        out int unsupportedFeatureCount) =>
        TryReadCore(bytes, options, allowUnresolvedViewport: false, out drawing, out unsupportedFeatureCount);

    private static bool TryReadCore(
        byte[]? bytes,
        OfficeSvgDrawingReaderOptions? options,
        bool allowUnresolvedViewport,
        out OfficeDrawing? drawing,
        out int unsupportedFeatureCount) {
        drawing = null;
        unsupportedFeatureCount = 0;
        if (!TryReadBoundedDocument(
                bytes,
                options,
                allowUnresolvedViewport,
                out XElement root,
                out int maximumElements,
                out double maximumViewportDimension,
                out double maximumViewportPixels,
                out double viewX,
                out double viewY,
                out double viewWidth,
                out double viewHeight,
                out double viewportWidth,
                out double viewportHeight)) return false;

        try {
            var scene = new OfficeDrawing(viewWidth, viewHeight);
            int visited = 0;
            int pathCommands = 0;
            bool pathCommandLimitExceeded = false;
            SvgDefinitionRegistry definitions = SvgDefinitionRegistry.Create(root);
            var paintServers = new SvgPaintServerRegistry(definitions);
            var references = new SvgElementReferenceRegistry(definitions);
            var context = ResolvePaintContext(root, SvgPaintContext.Default, paintServers, ref unsupportedFeatureCount);
            OfficeTransform rootTransform = ResolveTransform(root, OfficeTransform.Identity, viewX, viewY, ref unsupportedFeatureCount);
            AddChildren(root, scene, context, paintServers, references, rootTransform, viewX, viewY,
                maximumElements, maximumViewportDimension, maximumViewportPixels, 0,
                ref visited, ref pathCommands, ref pathCommandLimitExceeded, ref unsupportedFeatureCount);
            if (visited > maximumElements) return false;
            if (Math.Abs(viewportWidth - viewWidth) < 0.000001D && Math.Abs(viewportHeight - viewHeight) < 0.000001D) {
                drawing = scene;
            } else {
                if (!TryParsePreserveAspectRatio(root.Attribute("preserveAspectRatio")?.Value, out SvgAspectAlignment alignment, out bool slice)) {
                    alignment = SvgAspectAlignment.XMidYMid;
                    slice = false;
                    unsupportedFeatureCount++;
                }
                var viewport = new OfficeDrawing(viewportWidth, viewportHeight)
                    .AddEffectDrawing(scene, ResolveViewportTransform(viewWidth, viewHeight, viewportWidth, viewportHeight, alignment, slice));
                drawing = new OfficeDrawing(viewportWidth, viewportHeight)
                    .AddClippedDrawing(viewport, 0D, 0D, OfficeClipPath.Rectangle(viewportWidth, viewportHeight));
            }
            return IsSupportedSvgViewport(viewportWidth, viewportHeight, maximumViewportDimension, maximumViewportPixels);
        } catch (XmlException) {
            return false;
        } catch (InvalidOperationException) {
            return false;
        } catch (ArgumentException) {
            return false;
        }
    }

    /// <summary>
    /// Returns whether an SVG payload is well formed and stays within the supplied parser and viewport safety limits,
    /// regardless of whether every valid SVG shape can be imported into an <see cref="OfficeDrawing"/>.
    /// </summary>
    public static bool IsWithinSafetyLimits(byte[]? bytes, OfficeSvgDrawingReaderOptions? options = null) {
        if (!TryReadBoundedDocument(
                bytes,
                options,
                allowUnresolvedViewport: true,
                out XElement root,
                out int maximumElements,
                out _,
                out double maximumViewportPixels,
                out double viewX,
                out double viewY,
                out double viewWidth,
                out double viewHeight,
                out double viewportWidth,
                out double viewportHeight)) return false;

        return !ExceedsSvgElementNestingLimit(root) &&
               !ExceedsSvgDocumentPathCommandLimit(root) &&
               !ExceedsSvgRenderedExpansionLimits(root, maximumElements, maximumViewportPixels,
                   viewX, viewY, viewWidth, viewHeight, viewportWidth, viewportHeight);
    }

    private static bool TryReadBoundedDocument(
        byte[]? bytes,
        OfficeSvgDrawingReaderOptions? options,
        bool allowUnresolvedViewport,
        out XElement root,
        out int maximumElements,
        out double maximumViewportDimension,
        out double maximumViewportPixels,
        out double viewX,
        out double viewY,
        out double viewWidth,
        out double viewHeight,
        out double viewportWidth,
        out double viewportHeight) {
        root = null!;
        maximumElements = options?.MaximumElements ?? OfficeSvgDrawingReaderOptions.DefaultMaximumElements;
        maximumViewportDimension = options?.MaximumViewportDimension ?? OfficeSvgDrawingReaderOptions.DefaultMaximumViewportDimension;
        maximumViewportPixels = options?.MaximumViewportPixels ?? OfficeSvgDrawingReaderOptions.DefaultMaximumViewportPixels;
        viewX = viewY = viewWidth = viewHeight = viewportWidth = viewportHeight = 0D;
        if (bytes == null || bytes.Length == 0 || bytes.Length > MaximumInputBytes) return false;
        if (maximumElements <= 0 || maximumElements > OfficeSvgDrawingReaderOptions.MaximumAllowedElements) return false;
        if (maximumViewportDimension <= 0D || maximumViewportDimension > OfficeSvgDrawingReaderOptions.MaximumAllowedViewportDimension ||
            maximumViewportPixels <= 0D || maximumViewportPixels > OfficeSvgDrawingReaderOptions.MaximumAllowedViewportPixels) return false;

        try {
            var settings = new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                MaxCharactersInDocument = MaximumInputBytes
            };
            XDocument document;
            using (var stream = new MemoryStream(bytes, writable: false))
            using (XmlReader reader = XmlReader.Create(stream, settings)) {
                document = XDocument.Load(reader, LoadOptions.None);
            }

            XElement? documentRoot = document.Root;
            if (documentRoot == null || !string.Equals(documentRoot.Name.LocalName, "svg", StringComparison.OrdinalIgnoreCase)) return false;
            root = documentRoot;
            if (root.Descendants().Take(maximumElements + 1).Count() > maximumElements) return false;
            return TryResolveViewport(bytes, root, maximumViewportDimension, maximumViewportPixels, allowUnresolvedViewport,
                out viewX, out viewY, out viewWidth, out viewHeight,
                out viewportWidth, out viewportHeight);
        } catch (XmlException) {
            return false;
        } catch (InvalidOperationException) {
            return false;
        } catch (ArgumentException) {
            return false;
        }
    }

    private static bool TryResolveViewport(
        byte[] bytes,
        XElement root,
        double maximumViewportDimension,
        double maximumViewportPixels,
        bool allowUnresolvedViewport,
        out double viewX,
        out double viewY,
        out double viewWidth,
        out double viewHeight,
        out double viewportWidth,
        out double viewportHeight) {
        viewX = viewY = 0D;
        viewWidth = viewHeight = viewportWidth = viewportHeight = 0D;
        // ChartForgeX selects the first XML attribute by case-insensitive local name for
        // raster output dimensions; inline CSS does not override this allocation sink.
        string? widthText = ReadRasterViewportAttribute(root, "width");
        string? heightText = ReadRasterViewportAttribute(root, "height");
        bool hasDeclaredWidth = OfficeImageReader.TryParseSvgLength(widthText, out double declaredWidth);
        bool hasDeclaredHeight = OfficeImageReader.TryParseSvgLength(heightText, out double declaredHeight);
        if ((!string.IsNullOrWhiteSpace(widthText) && !hasDeclaredWidth)
            || (!string.IsNullOrWhiteSpace(heightText) && !hasDeclaredHeight)) return false;
        if (TryParseNumberList(ReadRasterProjectedAttribute(root, "viewBox"), out IReadOnlyList<double> viewBox)
            && viewBox.Count == 4
            && viewBox[2] > 0D
            && viewBox[3] > 0D) {
            viewX = viewBox[0];
            viewY = viewBox[1];
            viewWidth = viewBox[2];
            viewHeight = viewBox[3];
            viewportWidth = viewWidth;
            viewportHeight = viewHeight;
            if (hasDeclaredWidth) viewportWidth = declaredWidth;
            if (hasDeclaredHeight) viewportHeight = declaredHeight;
            if (hasDeclaredWidth && !hasDeclaredHeight) viewportHeight = declaredWidth * viewHeight / viewWidth;
            if (!hasDeclaredWidth && hasDeclaredHeight) viewportWidth = declaredHeight * viewWidth / viewHeight;
            return IsSupportedSvgViewport(viewWidth, viewHeight, maximumViewportDimension, maximumViewportPixels)
                && IsSupportedSvgViewport(viewportWidth, viewportHeight, maximumViewportDimension, maximumViewportPixels);
        }

        bool hasIntrinsicWidth = hasDeclaredWidth;
        bool hasIntrinsicHeight = hasDeclaredHeight;
        double intrinsicWidth = declaredWidth;
        double intrinsicHeight = declaredHeight;
        if ((hasIntrinsicWidth && intrinsicWidth > maximumViewportDimension) ||
            (hasIntrinsicHeight && intrinsicHeight > maximumViewportDimension)) {
            return false;
        }
        if (hasIntrinsicWidth && hasIntrinsicHeight) {
            viewWidth = viewportWidth = intrinsicWidth;
            viewHeight = viewportHeight = intrinsicHeight;
            return IsSupportedSvgViewport(viewWidth, viewHeight, maximumViewportDimension, maximumViewportPixels);
        }

        if (!OfficeImageReader.TryIdentify(bytes, ".svg", out OfficeImageInfo info) || info.Width <= 0 || info.Height <= 0) {
            if (!allowUnresolvedViewport) return false;
            viewWidth = viewportWidth = hasIntrinsicWidth ? intrinsicWidth : 300D;
            viewHeight = viewportHeight = hasIntrinsicHeight ? intrinsicHeight : 150D;
            return IsSupportedSvgViewport(viewWidth, viewHeight, maximumViewportDimension, maximumViewportPixels);
        }
        viewWidth = viewportWidth = info.Width * 96D / Math.Max(1D, info.DpiX);
        viewHeight = viewportHeight = info.Height * 96D / Math.Max(1D, info.DpiY);
        return IsSupportedSvgViewport(viewWidth, viewHeight, maximumViewportDimension, maximumViewportPixels);
    }

    private static string? ReadRasterViewportAttribute(XElement root, string name) =>
        root.Attributes()
            .FirstOrDefault(attribute => attribute.Name.LocalName.Equals(name, StringComparison.OrdinalIgnoreCase))
            ?.Value;

    private static bool IsSupportedSvgViewport(
        double width,
        double height,
        double maximumViewportDimension,
        double maximumViewportPixels) =>
        width > 0D && height > 0D &&
        width <= maximumViewportDimension && height <= maximumViewportDimension &&
        width * height <= maximumViewportPixels;

    private static bool ExceedsSvgElementNestingLimit(XElement root) {
        var pending = new Stack<(XElement Element, int Depth)>();
        foreach (XElement child in root.Elements()) pending.Push((child, 0));
        while (pending.Count > 0) {
            (XElement element, int depth) = pending.Pop();
            if (depth > MaximumSvgNestingDepth) return true;
            foreach (XElement child in element.Elements()) pending.Push((child, depth + 1));
        }
        return false;
    }

    private static bool ExceedsSvgDocumentPathCommandLimit(XElement root) {
        int commandCount = 0;
        foreach (XElement element in root.DescendantsAndSelf()) {
            if (!TryAddSvgGeometryCommands(element, ref commandCount)) return true;
        }
        return false;
    }

    private static bool ExceedsSvgRenderedExpansionLimits(
        XElement root,
        int maximumElements,
        double maximumViewportPixels,
        double viewX,
        double viewY,
        double viewWidth,
        double viewHeight,
        double viewportWidth,
        double viewportHeight) {
        if (HasPotentialStylesheetRenderedDefinitionReference(root)
            || HasStylesheetRasterGeometryDeclaration(root)
            || HasUnsupportedInlineRasterGeometryDeclaration(root)) return true;
        if (!TryResolveSupportedRasterTransform(
                root,
                OfficeTransform.Identity,
                viewX,
                viewY,
                out OfficeTransform rootTransform)) return true;
        if (!TryResolveRasterPixelScales(
                root,
                viewWidth,
                viewHeight,
                viewportWidth,
                viewportHeight,
                out double pixelScaleX,
                out double pixelScaleY)) return true;
        int commandCount = 0;
        int elementCount = 0;
        var rasterWork = new SvgRasterWorkBudget(maximumViewportPixels, viewX, viewY,
            viewWidth, viewHeight, viewportWidth, viewportHeight, pixelScaleX, pixelScaleY,
            HasStylesheetNonScalingStrokeDeclaration(root));
        var references = new SvgElementReferenceRegistry(SvgDefinitionRegistry.Create(root));
        string? fill = ResolveInheritedSvgPaint(root, "fill", inherited: null);
        string? stroke = ResolveInheritedSvgPaint(root, "stroke", inherited: null);
        if (!TryResolveRasterStrokeStyle(root, SvgRasterStrokeStyle.Default, out SvgRasterStrokeStyle strokeStyle)) return true;
        if (!TryResolveRasterTextStyle(root, SvgRasterTextStyle.Default, out SvgRasterTextStyle textStyle)) return true;
        string? marker = ResolveInheritedSvgPaint(root, "marker", inherited: null);
        string? markerStart = ResolveInheritedSvgPaint(root, "marker-start", marker);
        string? markerMid = ResolveInheritedSvgPaint(root, "marker-mid", marker);
        string? markerEnd = ResolveInheritedSvgPaint(root, "marker-end", marker);
        foreach (string propertyName in RenderedSvgLocalReferenceProperties) {
            if (!TryAddRenderedSvgLocalReference(
                    ReadRasterPresentationProperty(root, propertyName),
                    propertyName,
                    references,
                    maximumElements,
                    ref elementCount,
                    ref commandCount,
                    rootTransform,
                    viewX,
                    viewY,
                    rasterWork)) return true;
        }
        foreach (XElement child in root.Elements()) {
            if (!TryAddRenderedSvgExpansion(
                    child,
                    references,
                    maximumElements,
                    ref elementCount,
                    ref commandCount,
                    rootTransform,
                    viewX,
                    viewY,
                    rasterWork,
                    fill,
                    stroke,
                    markerStart,
                    markerMid,
                    markerEnd,
                    strokeStyle,
                    textStyle)) return true;
        }
        return false;
    }

    private static bool HasPotentialStylesheetRenderedDefinitionReference(XElement root) {
        var expandedDefinitionIds = new HashSet<string>(StringComparer.Ordinal);
        foreach (XElement element in root.Descendants()) {
            string name = element.Name.LocalName;
            bool isExpandedDefinition = name.Equals("pattern", StringComparison.OrdinalIgnoreCase)
                || name.Equals("mask", StringComparison.OrdinalIgnoreCase)
                || name.Equals("clipPath", StringComparison.OrdinalIgnoreCase)
                || name.Equals("filter", StringComparison.OrdinalIgnoreCase)
                || name.Equals("marker", StringComparison.OrdinalIgnoreCase)
                || name.Equals("linearGradient", StringComparison.OrdinalIgnoreCase)
                || name.Equals("radialGradient", StringComparison.OrdinalIgnoreCase);
            if (!isExpandedDefinition) continue;
            string? id = ReadRasterElementId(element);
            if (!string.IsNullOrEmpty(id)) expandedDefinitionIds.Add(id!);
        }
        return expandedDefinitionIds.Count > 0
            && (root.Descendants().Any(element =>
                    element.Name.LocalName.Equals("style", StringComparison.OrdinalIgnoreCase)
                    && ContainsLocalCssUrlReference(element.Value, expandedDefinitionIds))
                || root.DescendantsAndSelf().Any(element =>
                    ContainsLocalCssCustomPropertyUrlReference(ReadRasterInlineStyleAttribute(element), expandedDefinitionIds)));
    }

    // ChartForgeX projects non-XML attributes by local name into a dictionary, so a later
    // namespace-qualified attribute replaces an earlier value with the same local name.
    private static string? ReadRasterInlineStyleAttribute(XElement element) =>
        ReadRasterProjectedAttribute(element, "style");

    private static bool TryAddRenderedSvgExpansion(
        XElement element,
        SvgElementReferenceRegistry references,
        int maximumElements,
        ref int elementCount,
        ref int commandCount,
        OfficeTransform inheritedTransform,
        double viewX,
        double viewY,
        SvgRasterWorkBudget rasterWork,
        string? inheritedFill = null,
        string? inheritedStroke = null,
        string? inheritedMarkerStart = null,
        string? inheritedMarkerMid = null,
        string? inheritedMarkerEnd = null,
        SvgRasterStrokeStyle inheritedStrokeStyle = default,
        SvgRasterTextStyle inheritedTextStyle = default) {
        elementCount++;
        if (elementCount > maximumElements) return false;

        string name = element.Name.LocalName.ToLowerInvariant();
        if (name is "defs" or "title" or "desc" or "metadata" or "lineargradient" or "radialgradient" or "stop") return true;
        if (!TryResolveSupportedRasterTransform(
                element,
                inheritedTransform,
                viewX,
                viewY,
                out OfficeTransform transform)) return false;
        if (name == "use") {
            if (!TryReadRasterUsePlacement(element, out double useX, out double useY)) return false;
            transform = OfficeTransform.Translate(useX, useY).Then(transform);
            if (!IsSupportedSvgTransform(transform)) return false;
        }

        if (name == "svg"
            && !TryResolveNestedRasterViewportTransform(element, transform, out transform)) return false;

        if (!TryAddRenderedSvgPayloadComplexity(element, maximumElements, ref elementCount)) return false;
        string? fill = ResolveInheritedSvgPaint(element, "fill", inheritedFill);
        string? stroke = ResolveInheritedSvgPaint(element, "stroke", inheritedStroke);
        if (!TryResolveRasterStrokeStyle(element, inheritedStrokeStyle, out SvgRasterStrokeStyle strokeStyle)
            || !TryResolveRasterTextStyle(element, inheritedTextStyle, out SvgRasterTextStyle textStyle)
            || !rasterWork.TryChargeRenderedElement(element, transform, stroke, strokeStyle, textStyle)) return false;
        if (name == "textpath"
            && !TryAddRenderedSvgElementReference(
                element,
                "path",
                references,
                maximumElements,
                ref elementCount,
                ref commandCount,
                transform,
                viewX,
                viewY,
                rasterWork)) return false;

        string? marker = ResolveInheritedSvgPaint(element, "marker", inherited: null);
        string? markerStart = ResolveInheritedSvgPaint(element, "marker-start", marker ?? inheritedMarkerStart);
        string? markerMid = ResolveInheritedSvgPaint(element, "marker-mid", marker ?? inheritedMarkerMid);
        string? markerEnd = ResolveInheritedSvgPaint(element, "marker-end", marker ?? inheritedMarkerEnd);
        if (IsRenderedSvgPaintConsumer(name)) {
            if (!TryAddRenderedSvgPatternReference(
                    fill,
                    references,
                    maximumElements,
                    ref elementCount,
                    ref commandCount,
                    transform,
                    viewX,
                    viewY,
                    rasterWork)
                || !TryAddRenderedSvgPatternReference(
                    stroke,
                    references,
                    maximumElements,
                    ref elementCount,
                    ref commandCount,
                    transform,
                    viewX,
                    viewY,
                    rasterWork)) return false;
        }

        SvgMarkerPlacementCounts markerPlacements = CountSvgMarkerPlacements(element);
        if (markerPlacements.HasAny
            && (!TryAddRenderedSvgLocalReferenceApplications(
                    markerStart,
                    "marker",
                    markerPlacements.Start,
                    references,
                    maximumElements,
                    ref elementCount,
                    ref commandCount,
                    transform,
                    viewX,
                    viewY,
                    rasterWork)
                || !TryAddRenderedSvgLocalReferenceApplications(
                    markerMid,
                    "marker",
                    markerPlacements.Mid,
                    references,
                    maximumElements,
                    ref elementCount,
                    ref commandCount,
                    transform,
                    viewX,
                    viewY,
                    rasterWork)
                || !TryAddRenderedSvgLocalReferenceApplications(
                    markerEnd,
                    "marker",
                    markerPlacements.End,
                    references,
                    maximumElements,
                    ref elementCount,
                    ref commandCount,
                    transform,
                    viewX,
                    viewY,
                    rasterWork))) return false;

        foreach (string propertyName in RenderedSvgLocalReferenceProperties) {
            if (!TryAddRenderedSvgLocalReference(
                    ReadRasterPresentationProperty(element, propertyName),
                    propertyName,
                    references,
                    maximumElements,
                    ref elementCount,
                    ref commandCount,
                    transform,
                    viewX,
                    viewY,
                    rasterWork)) return false;
        }

        if (name == "use") {
            SvgElementReferenceEntryResult useResult = references.TryEnterDetailed(
                element,
                out string referenceId,
                out XElement? target);
            if (useResult is SvgElementReferenceEntryResult.DepthExceeded or SvgElementReferenceEntryResult.Cycle) return false;
            if (useResult != SvgElementReferenceEntryResult.Entered) return !HasLocalSvgElementReference(element);
            try {
                if (!TryResolveRenderedUseTargetTransform(
                        element,
                        target!,
                        transform,
                        out OfficeTransform targetTransform)) return false;
                return TryAddRenderedSvgExpansion(
                    target!,
                    references,
                    maximumElements,
                    ref elementCount,
                    ref commandCount,
                    targetTransform,
                    viewX,
                    viewY,
                    rasterWork,
                    fill,
                    stroke,
                    markerStart,
                    markerMid,
                    markerEnd,
                    strokeStyle,
                    textStyle);
            } finally {
                references.Exit(referenceId);
            }
        }

        if (name is "pattern" or "filter") {
            SvgElementReferenceEntryResult inheritedDefinitionResult = references.TryEnterDetailed(
                element,
                name,
                out string inheritedDefinitionId,
                out XElement? inheritedDefinition);
            if (inheritedDefinitionResult is SvgElementReferenceEntryResult.DepthExceeded or SvgElementReferenceEntryResult.Cycle) return false;
            if (inheritedDefinitionResult != SvgElementReferenceEntryResult.Entered
                && HasLocalSvgElementReference(element)) return false;
            if (inheritedDefinitionResult == SvgElementReferenceEntryResult.Entered) {
                try {
                    if (!TryAddRenderedSvgDefinitionExpansion(
                            inheritedDefinition!,
                            references,
                            maximumElements,
                            ref elementCount,
                            ref commandCount,
                            transform,
                            viewX,
                            viewY,
                            rasterWork)) return false;
                } finally {
                    references.Exit(inheritedDefinitionId);
                }
            }
        }

        if (!TryAddSvgGeometryCommands(element, ref commandCount)) return false;
        foreach (XElement child in element.Elements()) {
            if (!TryAddRenderedSvgExpansion(
                    child,
                    references,
                    maximumElements,
                    ref elementCount,
                    ref commandCount,
                    transform,
                    viewX,
                    viewY,
                    rasterWork,
                    fill,
                    stroke,
                    markerStart,
                    markerMid,
                    markerEnd,
                    strokeStyle,
                    textStyle)) return false;
        }
        return true;
    }

    private static readonly string[] RenderedSvgLocalReferenceProperties = {
        "mask",
        "clip-path",
        "filter"
    };

    private static bool TryAddRenderedSvgLocalReference(
        string? value,
        string propertyName,
        SvgElementReferenceRegistry references,
        int maximumElements,
        ref int elementCount,
        ref int commandCount,
        OfficeTransform transform,
        double viewX,
        double viewY,
        SvgRasterWorkBudget rasterWork) {
        SvgElementReferenceEntryResult result = references.TryEnterLocalDetailed(
            value,
            out string referenceId,
            out XElement? target);
        if (result is SvgElementReferenceEntryResult.DepthExceeded or SvgElementReferenceEntryResult.Cycle) return false;
        if (result != SvgElementReferenceEntryResult.Entered) return !HasPotentialSvgUrlFunction(value);
        try {
            string targetName = target!.Name.LocalName;
            bool conservativeReferencePlacement =
                (propertyName.Equals("mask", StringComparison.OrdinalIgnoreCase)
                    && targetName.Equals("mask", StringComparison.OrdinalIgnoreCase))
                || (propertyName.Equals("clip-path", StringComparison.OrdinalIgnoreCase)
                    && targetName.Equals("clipPath", StringComparison.OrdinalIgnoreCase)
                    && string.Equals(
                        ReadRasterProjectedAttribute(target, "clipPathUnits")?.Trim(),
                        "objectBoundingBox",
                        StringComparison.OrdinalIgnoreCase));
            if (conservativeReferencePlacement) rasterWork.EnterConservativePlacement();
            try {
                return TryAddRenderedSvgDefinitionExpansion(
                    target!,
                    references,
                    maximumElements,
                    ref elementCount,
                    ref commandCount,
                    transform,
                    viewX,
                    viewY,
                    rasterWork);
            } finally {
                if (conservativeReferencePlacement) rasterWork.ExitConservativePlacement();
            }
        } finally {
            references.Exit(referenceId);
        }
    }

    private static bool TryAddRenderedSvgPatternReference(
        string? value,
        SvgElementReferenceRegistry references,
        int maximumElements,
        ref int elementCount,
        ref int commandCount,
        OfficeTransform transform,
        double viewX,
        double viewY,
        SvgRasterWorkBudget rasterWork) {
        SvgElementReferenceEntryResult result = references.TryEnterLocalDetailed(
            value,
            out string referenceId,
            out XElement? target);
        if (result is SvgElementReferenceEntryResult.DepthExceeded or SvgElementReferenceEntryResult.Cycle) return false;
        if (result != SvgElementReferenceEntryResult.Entered) return !HasPotentialSvgUrlFunction(value);
        try {
            string targetName = target!.Name.LocalName;
            if (targetName.Equals("linearGradient", StringComparison.OrdinalIgnoreCase)
                || targetName.Equals("radialGradient", StringComparison.OrdinalIgnoreCase)) {
                return TryValidateInheritedGradientReference(target, references);
            }
            if (!targetName.Equals("pattern", StringComparison.OrdinalIgnoreCase)) return false;
            // Pattern units, content units, viewBox, and preserveAspectRatio can remap every
            // descendant across the consumer bounds. Keep safe patterns compatible while
            // conservatively charging each expanded paint operation as a viewport repaint.
            rasterWork.EnterConservativePlacement();
            try {
                return TryAddRenderedSvgDefinitionExpansion(
                    target,
                    references,
                    maximumElements,
                    ref elementCount,
                    ref commandCount,
                    transform,
                    viewX,
                    viewY,
                    rasterWork);
            } finally {
                rasterWork.ExitConservativePlacement();
            }
        } finally {
            references.Exit(referenceId);
        }
    }

    private static bool TryValidateInheritedGradientReference(
        XElement gradient,
        SvgElementReferenceRegistry references) {
        SvgElementReferenceEntryResult result = references.TryEnterDetailed(
            gradient,
            gradient.Name.LocalName,
            out string referenceId,
            out XElement? inheritedGradient);
        if (result is SvgElementReferenceEntryResult.DepthExceeded or SvgElementReferenceEntryResult.Cycle) return false;
        if (result != SvgElementReferenceEntryResult.Entered) return !HasLocalSvgElementReference(gradient);
        try {
            return TryValidateInheritedGradientReference(inheritedGradient!, references);
        } finally {
            references.Exit(referenceId);
        }
    }

    private static bool HasPotentialSvgUrlFunction(string? value) =>
        ContainsPotentialCssIdentifier(value, "url");

    private static string? ResolveInheritedSvgPaint(XElement element, string propertyName, string? inherited) {
        string? value = ReadRasterPresentationProperty(element, propertyName);
        if (string.IsNullOrWhiteSpace(value)) return inherited;
        return value!.Trim().Equals("inherit", StringComparison.OrdinalIgnoreCase) ? inherited : value;
    }

    private static bool IsRenderedSvgPaintConsumer(string name) => name is not (
        "svg" or "g" or "a" or "switch" or "symbol" or "pattern" or "mask" or "clippath" or
        "filter" or "marker" or "style");

    private static bool TryAddSvgGeometryCommands(XElement element, ref int commandCount) {
        string name = element.Name.LocalName;
        int remaining = MaximumSvgPathCommands - commandCount;
        if (name.Equals("path", StringComparison.OrdinalIgnoreCase)) {
            _ = OfficeSvgPathDataParser.TryParse(
                ReadRasterProjectedAttribute(element, "d"),
                remaining + 1,
                out IReadOnlyList<OfficePathCommand> commands,
                out bool commandLimitExceeded);
            if (commandLimitExceeded || commands.Count > remaining) return false;
            commandCount += commands.Count;
            return true;
        }

        bool close = name.Equals("polygon", StringComparison.OrdinalIgnoreCase);
        if (!close && !name.Equals("polyline", StringComparison.OrdinalIgnoreCase)) return true;
        int maximumValues = (remaining + 1) * 2;
        bool parsed = TryParseNumberList(
            ReadRasterProjectedAttribute(element, "points"),
            maximumValues,
            out IReadOnlyList<double> values,
            out bool valueLimitExceeded);
        if (valueLimitExceeded) return false;
        int elementCommands = values.Count / 2;
        if (parsed && close && values.Count >= 6 && values.Count % 2 == 0) elementCommands++;
        if (elementCommands > remaining) return false;
        commandCount += elementCommands;
        return true;
    }

    private static void AddChildren(
        XElement parent,
        OfficeDrawing drawing,
        SvgPaintContext inherited,
        SvgPaintServerRegistry paintServers,
        SvgElementReferenceRegistry references,
        OfficeTransform inheritedTransform,
        double viewX,
        double viewY,
        int maximumElements,
        double maximumViewportDimension,
        double maximumViewportPixels,
        int depth,
        ref int visited,
        ref int pathCommands,
        ref bool pathCommandLimitExceeded,
        ref int unsupported) {
        if (depth > MaximumSvgNestingDepth) {
            unsupported++;
            return;
        }
        foreach (XElement element in parent.Elements()) {
            AddElement(element, drawing, inherited, paintServers, references, inheritedTransform, viewX, viewY,
                maximumElements, maximumViewportDimension, maximumViewportPixels, depth,
                ref visited, ref pathCommands, ref pathCommandLimitExceeded, ref unsupported);
            if (visited > maximumElements) return;
        }
    }

    private static void AddElement(
        XElement element,
        OfficeDrawing drawing,
        SvgPaintContext inherited,
        SvgPaintServerRegistry paintServers,
        SvgElementReferenceRegistry references,
        OfficeTransform inheritedTransform,
        double viewX,
        double viewY,
        int maximumElements,
        double maximumViewportDimension,
        double maximumViewportPixels,
        int depth,
        ref int visited,
        ref int pathCommands,
        ref bool pathCommandLimitExceeded,
        ref int unsupported) {
        visited++;
        if (visited > maximumElements) return;
        string name = element.Name.LocalName.ToLowerInvariant();
        if (name is "title" or "desc" or "metadata" or "lineargradient" or "radialgradient" or "stop") return;
        if (name == "defs") return;

        SvgPaintContext style = ResolvePaintContext(element, inherited, paintServers, ref unsupported);
        if (!style.Visible) return;
        OfficeTransform transform = ResolveTransform(element, inheritedTransform, viewX, viewY, ref unsupported);
        if (name is "g" or "svg" or "a" or "switch") {
            bool hasEffects = TryResolveSvgEffects(
                element,
                drawing.Width,
                drawing.Height,
                style,
                paintServers,
                references,
                transform,
                viewX,
                viewY,
                maximumElements,
                maximumViewportDimension,
                maximumViewportPixels,
                depth,
                ref visited,
                ref pathCommands,
                ref pathCommandLimitExceeded,
                ref unsupported,
                out OfficeBlendMode blendMode,
                out OfficeDrawingSoftMask? softMask);
            OfficeDrawing target = hasEffects ? new OfficeDrawing(drawing.Width, drawing.Height) : drawing;
            AddChildren(element, target, style, paintServers, references, transform, viewX, viewY,
                maximumElements, maximumViewportDimension, maximumViewportPixels, depth + 1,
                ref visited, ref pathCommands, ref pathCommandLimitExceeded, ref unsupported);
            if (hasEffects) drawing.AddEffectDrawing(target, OfficeTransform.Identity, blendMode, softMask);
            return;
        }
        if (name is "use" or "text") {
            bool hasEffects = TryResolveSvgEffects(
                element,
                drawing.Width,
                drawing.Height,
                style,
                paintServers,
                references,
                transform,
                viewX,
                viewY,
                maximumElements,
                maximumViewportDimension,
                maximumViewportPixels,
                depth,
                ref visited,
                ref pathCommands,
                ref pathCommandLimitExceeded,
                ref unsupported,
                out OfficeBlendMode blendMode,
                out OfficeDrawingSoftMask? softMask);
            OfficeDrawing target = hasEffects ? new OfficeDrawing(drawing.Width, drawing.Height) : drawing;
            if (name == "use") {
                AddReferencedElement(element, target, style, paintServers, references, transform, viewX, viewY,
                    maximumElements, maximumViewportDimension, maximumViewportPixels, depth + 1,
                    ref visited, ref pathCommands, ref pathCommandLimitExceeded, ref unsupported);
            } else {
                AddText(element, target, style, paintServers, transform, viewX, viewY, ref unsupported);
            }
            if (hasEffects) drawing.AddEffectDrawing(target, OfficeTransform.Identity, blendMode, softMask);
            return;
        }

        OfficeDrawingShape? shape = name switch {
            "rect" => CreateRectangle(element, style, viewX, viewY, drawing.Width, drawing.Height, ref unsupported),
            "circle" => CreateCircle(element, style, viewX, viewY, drawing.Width, drawing.Height),
            "ellipse" => CreateEllipse(element, style, viewX, viewY, drawing.Width, drawing.Height),
            "line" => CreateLine(element, style, viewX, viewY, drawing.Width, drawing.Height),
            "polygon" => CreatePolygon(element, style, viewX, viewY, close: true, ref pathCommands, ref pathCommandLimitExceeded),
            "polyline" => CreatePolygon(element, style, viewX, viewY, close: false, ref pathCommands, ref pathCommandLimitExceeded),
            "path" => CreatePath(element, style, viewX, viewY, ref pathCommands, ref pathCommandLimitExceeded),
            _ => null
        };
        if (shape == null) {
            unsupported++;
            return;
        }

        ApplyDeferredPaint(shape.Shape, style, shape.X, shape.Y, drawing.Width, drawing.Height, viewX, viewY, ref unsupported);

        ApplyTransform(shape, transform);

        try {
            bool hasEffects = TryResolveSvgEffects(
                element,
                drawing.Width,
                drawing.Height,
                style,
                paintServers,
                references,
                transform,
                viewX,
                viewY,
                maximumElements,
                maximumViewportDimension,
                maximumViewportPixels,
                depth,
                ref visited,
                ref pathCommands,
                ref pathCommandLimitExceeded,
                ref unsupported,
                out OfficeBlendMode blendMode,
                out OfficeDrawingSoftMask? softMask);
            if (hasEffects) {
                var target = new OfficeDrawing(drawing.Width, drawing.Height);
                target.AddShape(shape.Shape, shape.X, shape.Y);
                drawing.AddEffectDrawing(target, OfficeTransform.Identity, blendMode, softMask);
            } else {
                drawing.AddShape(shape.Shape, shape.X, shape.Y);
            }
        } catch (ArgumentOutOfRangeException) {
            unsupported++;
        }
    }

    private static OfficeTransform ResolveTransform(
        XElement element,
        OfficeTransform inherited,
        double viewX,
        double viewY,
        ref int unsupported) {
        string? value = element.Attribute("transform")?.Value;
        if (string.IsNullOrWhiteSpace(value)) return inherited;
        if (!OfficeSvgTransformParser.TryParse(value, out OfficeTransform parsed)) {
            unsupported++;
            return inherited;
        }
        OfficeTransform normalized = OfficeTransform.Translate(viewX, viewY)
            .Then(parsed)
            .Then(OfficeTransform.Translate(-viewX, -viewY));
        OfficeTransform combined = normalized.Then(inherited);
        if (!IsSupportedSvgTransform(combined)) {
            unsupported++;
            return inherited;
        }
        return combined;
    }

    private static bool IsSupportedSvgTransform(OfficeTransform transform) =>
        Math.Abs(transform.M11) <= MaximumSvgTransformCoefficient &&
        Math.Abs(transform.M12) <= MaximumSvgTransformCoefficient &&
        Math.Abs(transform.M21) <= MaximumSvgTransformCoefficient &&
        Math.Abs(transform.M22) <= MaximumSvgTransformCoefficient &&
        Math.Abs(transform.OffsetX) <= MaximumSvgTransformOffset &&
        Math.Abs(transform.OffsetY) <= MaximumSvgTransformOffset;

    private static void ApplyTransform(OfficeDrawingShape drawingShape, OfficeTransform transform) {
        if (transform == OfficeTransform.Identity) return;
        OfficeTransform local = OfficeTransform.Translate(drawingShape.X, drawingShape.Y)
            .Then(transform)
            .Then(OfficeTransform.Translate(-drawingShape.X, -drawingShape.Y));
        OfficeShape shape = drawingShape.Shape;
        shape.Transform = shape.Transform.HasValue ? shape.Transform.Value.Then(local) : local;
    }

    private static OfficeDrawingShape? CreateRectangle(
        XElement element,
        SvgPaintContext style,
        double viewX,
        double viewY,
        double viewportWidth,
        double viewportHeight,
        ref int unsupported) {
        if (!TryViewportLength(element, "width", viewportWidth, out double width)
            || !TryViewportLength(element, "height", viewportHeight, out double height)
            || width <= 0D
            || height <= 0D) return null;
        double x = ReadViewportCoordinate(element, "x", viewX, viewportWidth);
        double y = ReadViewportCoordinate(element, "y", viewY, viewportHeight);
        double rx = ReadViewportLength(element, "rx", viewportWidth);
        double ry = ReadViewportLength(element, "ry", viewportHeight);
        if (rx <= 0D && ry > 0D) rx = ry;
        if (ry <= 0D && rx > 0D) ry = rx;
        OfficeShape shape;
        if (rx > 0D || ry > 0D) {
            if (Math.Abs(rx - ry) > 0.0001D) unsupported++;
            shape = OfficeShape.RoundedRectangle(width, height, Math.Min(Math.Min(rx, ry), Math.Min(width, height) / 2D));
        } else {
            shape = OfficeShape.Rectangle(width, height);
        }
        ApplyPaint(shape, style);
        return new OfficeDrawingShape(shape, x, y);
    }

    private static OfficeDrawingShape? CreateCircle(
        XElement element,
        SvgPaintContext style,
        double viewX,
        double viewY,
        double viewportWidth,
        double viewportHeight) {
        double normalizedDiagonal = Math.Sqrt((viewportWidth * viewportWidth) + (viewportHeight * viewportHeight)) / Math.Sqrt(2D);
        if (!TryViewportLength(element, "r", normalizedDiagonal, out double radius) || radius <= 0D) return null;
        double x = ReadViewportCoordinate(element, "cx", viewX, viewportWidth) - radius;
        double y = ReadViewportCoordinate(element, "cy", viewY, viewportHeight) - radius;
        OfficeShape shape = OfficeShape.Ellipse(radius * 2D, radius * 2D);
        ApplyPaint(shape, style);
        return new OfficeDrawingShape(shape, x, y);
    }

    private static OfficeDrawingShape? CreateEllipse(
        XElement element,
        SvgPaintContext style,
        double viewX,
        double viewY,
        double viewportWidth,
        double viewportHeight) {
        if (!TryViewportLength(element, "rx", viewportWidth, out double radiusX)
            || !TryViewportLength(element, "ry", viewportHeight, out double radiusY)
            || radiusX <= 0D
            || radiusY <= 0D) return null;
        double x = ReadViewportCoordinate(element, "cx", viewX, viewportWidth) - radiusX;
        double y = ReadViewportCoordinate(element, "cy", viewY, viewportHeight) - radiusY;
        OfficeShape shape = OfficeShape.Ellipse(radiusX * 2D, radiusY * 2D);
        ApplyPaint(shape, style);
        return new OfficeDrawingShape(shape, x, y);
    }

    private static OfficeDrawingShape? CreateLine(
        XElement element,
        SvgPaintContext style,
        double viewX,
        double viewY,
        double viewportWidth,
        double viewportHeight) {
        double x1 = ReadViewportCoordinate(element, "x1", viewX, viewportWidth);
        double y1 = ReadViewportCoordinate(element, "y1", viewY, viewportHeight);
        double x2 = ReadViewportCoordinate(element, "x2", viewX, viewportWidth);
        double y2 = ReadViewportCoordinate(element, "y2", viewY, viewportHeight);
        if (Math.Abs(x1 - x2) <= 0.0001D && Math.Abs(y1 - y2) <= 0.0001D) return null;
        OfficeShape shape = OfficeShape.Line(x1, y1, x2, y2);
        shape.FillColor = null;
        shape.StrokeColor = style.Stroke ?? style.Fill ?? OfficeColor.Black;
        shape.StrokeGradient = style.StrokeGradient;
        shape.StrokeRadialGradient = style.StrokeRadialGradient;
        shape.StrokeWidth = style.StrokeWidth;
        shape.StrokeOpacity = style.StrokeOpacity * style.Opacity;
        shape.StrokeDashStyle = style.DashStyle;
        shape.StrokeLineCap = style.LineCap;
        shape.StrokeLineJoin = style.LineJoin;
        double x = Math.Min(x1, x2);
        double y = Math.Min(y1, y2);
        return new OfficeDrawingShape(shape, x, y);
    }

    private static OfficeDrawingShape? CreatePolygon(XElement element, SvgPaintContext style, double viewX,
        double viewY, bool close, ref int pathCommands, ref bool pathCommandLimitExceeded) {
        int remainingCommands = MaximumSvgPathCommands - pathCommands;
        if (remainingCommands <= 0) {
            int minimumValues = close ? 6 : 4;
            _ = TryParseNumberList(element.Attribute("points")?.Value, minimumValues,
                out IReadOnlyList<double> probeValues, out _);
            pathCommandLimitExceeded |= probeValues.Count >= minimumValues;
            return null;
        }
        bool parsed = TryParseNumberList(element.Attribute("points")?.Value, remainingCommands * 2,
            out IReadOnlyList<double> values, out bool limitExceeded);
        if (!parsed || values.Count < 4 || values.Count % 2 != 0) {
            if (limitExceeded) {
                pathCommands = MaximumSvgPathCommands;
                pathCommandLimitExceeded = true;
            } else if (values.Count > 0) {
                int parsedCommands = Math.Max(1, (values.Count + 1) / 2);
                pathCommands += Math.Min(remainingCommands, parsedCommands);
            }
            return null;
        }
        int commandCount = values.Count / 2;
        if (close) commandCount++;
        if (close && values.Count < 6) {
            return null;
        }
        if (commandCount > remainingCommands) {
            pathCommands = MaximumSvgPathCommands;
            pathCommandLimitExceeded = true;
            return null;
        }
        var points = new List<OfficePoint>(values.Count / 2);
        for (int index = 0; index < values.Count; index += 2) points.Add(new OfficePoint(values[index] - viewX, values[index + 1] - viewY));
        double minX = points.Min(point => point.X);
        double minY = points.Min(point => point.Y);
        OfficeShape shape;
        if (close) {
            shape = OfficeShape.Polygon(points);
        } else {
            var commands = new List<OfficePathCommand> { OfficePathCommand.MoveTo(points[0]) };
            for (int index = 1; index < points.Count; index++) commands.Add(OfficePathCommand.LineTo(points[index]));
            try {
                shape = OfficeShape.Path(commands);
            } catch (ArgumentException) {
                return null;
            }
            shape.FillColor = null;
        }
        pathCommands += commandCount;
        ApplyPaint(shape, style);
        return new OfficeDrawingShape(shape, minX, minY);
    }

    private static OfficeDrawingShape? CreatePath(XElement element, SvgPaintContext style, double viewX,
        double viewY, ref int pathCommands, ref bool pathCommandLimitExceeded) {
        int remaining = MaximumSvgPathCommands - pathCommands;
        if (remaining <= 0) {
            _ = OfficeSvgPathDataParser.TryParse(element.Attribute("d")?.Value, 1,
                out IReadOnlyList<OfficePathCommand> probeCommands, out _);
            pathCommandLimitExceeded |= probeCommands.Count > 0;
            return null;
        }
        if (!OfficeSvgPathDataParser.TryParse(element.Attribute("d")?.Value, remaining,
                out IReadOnlyList<OfficePathCommand> parsed, out bool commandLimitExceeded)) {
            pathCommands += Math.Min(remaining, parsed.Count);
            if (commandLimitExceeded) {
                pathCommands = MaximumSvgPathCommands;
                pathCommandLimitExceeded = true;
            }
            return null;
        }
        pathCommands += parsed.Count;
        var commands = new List<OfficePathCommand>(parsed.Count + 1);
        double minX = double.PositiveInfinity;
        double minY = double.PositiveInfinity;
        double maxX = double.NegativeInfinity;
        double maxY = double.NegativeInfinity;
        foreach (OfficePathCommand source in parsed) {
            OfficePathCommand command = source.Translate(viewX, viewY);
            commands.Add(command);
            IncludeCommandBounds(command, ref minX, ref minY, ref maxX, ref maxY);
        }
        if (double.IsInfinity(minX) || double.IsInfinity(minY)) return null;
        if (maxX - minX <= 0.0001D) commands.Add(OfficePathCommand.MoveTo(maxX + 0.0001D, maxY));
        if (maxY - minY <= 0.0001D) commands.Add(OfficePathCommand.MoveTo(maxX, maxY + 0.0001D));
        OfficeShape shape;
        try {
            shape = OfficeShape.Path(commands);
        } catch (ArgumentException) {
            return null;
        }
        ApplyPaint(shape, style);
        return new OfficeDrawingShape(shape, minX, minY);
    }

    private static void IncludeCommandBounds(
        OfficePathCommand command,
        ref double minX,
        ref double minY,
        ref double maxX,
        ref double maxY) {
        switch (command.Kind) {
            case OfficePathCommandKind.MoveTo:
            case OfficePathCommandKind.LineTo:
                IncludePoint(command.Point, ref minX, ref minY, ref maxX, ref maxY);
                break;
            case OfficePathCommandKind.QuadraticBezierTo:
                IncludePoint(command.ControlPoint1, ref minX, ref minY, ref maxX, ref maxY);
                IncludePoint(command.Point, ref minX, ref minY, ref maxX, ref maxY);
                break;
            case OfficePathCommandKind.CubicBezierTo:
                IncludePoint(command.ControlPoint1, ref minX, ref minY, ref maxX, ref maxY);
                IncludePoint(command.ControlPoint2, ref minX, ref minY, ref maxX, ref maxY);
                IncludePoint(command.Point, ref minX, ref minY, ref maxX, ref maxY);
                break;
        }
    }

    private static void IncludePoint(
        OfficePoint point,
        ref double minX,
        ref double minY,
        ref double maxX,
        ref double maxY) {
        minX = Math.Min(minX, point.X);
        minY = Math.Min(minY, point.Y);
        maxX = Math.Max(maxX, point.X);
        maxY = Math.Max(maxY, point.Y);
    }

    private static SvgPaintContext ResolvePaintContext(XElement element, SvgPaintContext inherited, SvgPaintServerRegistry paintServers, ref int unsupported) {
        SvgPaintContext result = inherited;
        ApplyProperty("color", element.Attribute("color")?.Value, paintServers, ref result, ref unsupported);
        string? styleText = element.Attribute("style")?.Value;
        string[] declarations = string.IsNullOrWhiteSpace(styleText) ? Array.Empty<string>() : styleText!.Split(';');
        foreach (string declaration in declarations) {
            int colon = declaration.IndexOf(':');
            if (colon <= 0 || !declaration.Substring(0, colon).Trim().Equals("color", StringComparison.OrdinalIgnoreCase)) continue;
            ApplyProperty("color", declaration.Substring(colon + 1).Trim(), paintServers, ref result, ref unsupported);
        }
        ApplyProperty("fill", element.Attribute("fill")?.Value, paintServers, ref result, ref unsupported);
        ApplyProperty("stroke", element.Attribute("stroke")?.Value, paintServers, ref result, ref unsupported);
        ApplyProperty("stroke-width", element.Attribute("stroke-width")?.Value, paintServers, ref result, ref unsupported);
        ApplyProperty("opacity", element.Attribute("opacity")?.Value, paintServers, ref result, ref unsupported);
        ApplyProperty("fill-opacity", element.Attribute("fill-opacity")?.Value, paintServers, ref result, ref unsupported);
        ApplyProperty("stroke-opacity", element.Attribute("stroke-opacity")?.Value, paintServers, ref result, ref unsupported);
        ApplyProperty("stroke-dasharray", element.Attribute("stroke-dasharray")?.Value, paintServers, ref result, ref unsupported);
        ApplyProperty("stroke-linecap", element.Attribute("stroke-linecap")?.Value, paintServers, ref result, ref unsupported);
        ApplyProperty("stroke-linejoin", element.Attribute("stroke-linejoin")?.Value, paintServers, ref result, ref unsupported);
        ApplyProperty("fill-rule", element.Attribute("fill-rule")?.Value, paintServers, ref result, ref unsupported);
        ApplyProperty("font-family", element.Attribute("font-family")?.Value, paintServers, ref result, ref unsupported);
        ApplyProperty("font-size", element.Attribute("font-size")?.Value, paintServers, ref result, ref unsupported);
        ApplyProperty("font-style", element.Attribute("font-style")?.Value, paintServers, ref result, ref unsupported);
        ApplyProperty("font-weight", element.Attribute("font-weight")?.Value, paintServers, ref result, ref unsupported);
        ApplyProperty("text-anchor", element.Attribute("text-anchor")?.Value, paintServers, ref result, ref unsupported);
        ApplyProperty("dominant-baseline", element.Attribute("dominant-baseline")?.Value, paintServers, ref result, ref unsupported);
        ApplyProperty("display", element.Attribute("display")?.Value, paintServers, ref result, ref unsupported);
        ApplyProperty("visibility", element.Attribute("visibility")?.Value, paintServers, ref result, ref unsupported);
        ApplyProperty("filter", element.Attribute("filter")?.Value, paintServers, ref result, ref unsupported);
        ApplyProperty("clip-path", element.Attribute("clip-path")?.Value, paintServers, ref result, ref unsupported);
        ApplyProperty("marker-start", element.Attribute("marker-start")?.Value, paintServers, ref result, ref unsupported);
        ApplyProperty("marker-mid", element.Attribute("marker-mid")?.Value, paintServers, ref result, ref unsupported);
        ApplyProperty("marker-end", element.Attribute("marker-end")?.Value, paintServers, ref result, ref unsupported);
        foreach (string declaration in declarations) {
            int colon = declaration.IndexOf(':');
            if (colon <= 0) continue;
            string name = declaration.Substring(0, colon).Trim();
            if (name.Equals("color", StringComparison.OrdinalIgnoreCase)) continue;
            ApplyProperty(name, declaration.Substring(colon + 1).Trim(), paintServers, ref result, ref unsupported);
        }
        return result;
    }

    private static void ApplyProperty(string name, string? value, SvgPaintServerRegistry paintServers, ref SvgPaintContext style, ref int unsupported) {
        if (string.IsNullOrWhiteSpace(value)) return;
        string normalized = value!.Trim();
        switch (name.Trim().ToLowerInvariant()) {
            case "color":
                if (normalized.Equals("currentcolor", StringComparison.OrdinalIgnoreCase)) break;
                if (!OfficeColor.TryParseCss(normalized, out OfficeColor currentColor)) unsupported++;
                else style.Color = currentColor;
                break;
            case "fill":
                if (!TryPaint(normalized, paintServers, style.Color, out SvgResolvedPaint fill)) {
                    unsupported++;
                    if (normalized.StartsWith("url(", StringComparison.OrdinalIgnoreCase)) style.SetFill(default);
                }
                else style.SetFill(fill);
                break;
            case "stroke":
                if (!TryPaint(normalized, paintServers, style.Color, out SvgResolvedPaint stroke)) {
                    unsupported++;
                    if (normalized.StartsWith("url(", StringComparison.OrdinalIgnoreCase)) style.SetStroke(default);
                }
                else style.SetStroke(stroke);
                break;
            case "stroke-width":
                if (!TrySvgLength(normalized, out double strokeWidth) || strokeWidth < 0D) unsupported++;
                else style.StrokeWidth = strokeWidth;
                break;
            case "opacity":
                if (!TryUnit(normalized, out double opacity)) unsupported++;
                else style.Opacity *= opacity;
                break;
            case "fill-opacity":
                if (!TryUnit(normalized, out double fillOpacity)) unsupported++;
                else style.FillOpacity = fillOpacity;
                break;
            case "stroke-opacity":
                if (!TryUnit(normalized, out double strokeOpacity)) unsupported++;
                else style.StrokeOpacity = strokeOpacity;
                break;
            case "stroke-dasharray":
                if (normalized.Equals("none", StringComparison.OrdinalIgnoreCase)) style.DashStyle = OfficeStrokeDashStyle.Solid;
                else if (TryParseNumberList(normalized, out IReadOnlyList<double> dash) && dash.Count >= 2) style.DashStyle = OfficeStrokeDashStyle.Dash;
                else unsupported++;
                break;
            case "stroke-linecap":
                if (normalized.Equals("butt", StringComparison.OrdinalIgnoreCase)) style.LineCap = OfficeStrokeLineCap.Butt;
                else if (normalized.Equals("round", StringComparison.OrdinalIgnoreCase)) style.LineCap = OfficeStrokeLineCap.Round;
                else if (normalized.Equals("square", StringComparison.OrdinalIgnoreCase)) style.LineCap = OfficeStrokeLineCap.Square;
                else unsupported++;
                break;
            case "stroke-linejoin":
                if (normalized.Equals("miter", StringComparison.OrdinalIgnoreCase)) style.LineJoin = OfficeStrokeLineJoin.Miter;
                else if (normalized.Equals("round", StringComparison.OrdinalIgnoreCase)) style.LineJoin = OfficeStrokeLineJoin.Round;
                else if (normalized.Equals("bevel", StringComparison.OrdinalIgnoreCase)) style.LineJoin = OfficeStrokeLineJoin.Bevel;
                else unsupported++;
                break;
            case "fill-rule":
                if (normalized.Equals("nonzero", StringComparison.OrdinalIgnoreCase)) style.FillRule = OfficeFillRule.NonZero;
                else if (normalized.Equals("evenodd", StringComparison.OrdinalIgnoreCase)) style.FillRule = OfficeFillRule.EvenOdd;
                else unsupported++;
                break;
            case "font-family":
                string family = normalized.Split(',')[0].Trim().Trim('\'', '"');
                if (family.Length == 0) unsupported++;
                else style.FontFamily = family;
                break;
            case "font-size":
                if (!TrySvgLength(normalized, out double fontSize) || fontSize <= 0D) unsupported++;
                else style.FontSize = fontSize;
                break;
            case "font-style":
                if (normalized.Equals("normal", StringComparison.OrdinalIgnoreCase)) style.FontStyle &= ~OfficeFontStyle.Italic;
                else if (normalized.Equals("italic", StringComparison.OrdinalIgnoreCase) || normalized.Equals("oblique", StringComparison.OrdinalIgnoreCase)) style.FontStyle |= OfficeFontStyle.Italic;
                else unsupported++;
                break;
            case "font-weight":
                if (normalized.Equals("normal", StringComparison.OrdinalIgnoreCase) || normalized == "400") style.FontStyle &= ~OfficeFontStyle.Bold;
                else if (normalized.Equals("bold", StringComparison.OrdinalIgnoreCase) || normalized.Equals("bolder", StringComparison.OrdinalIgnoreCase)) style.FontStyle |= OfficeFontStyle.Bold;
                else if (int.TryParse(normalized, NumberStyles.Integer, CultureInfo.InvariantCulture, out int weight) && weight >= 1 && weight <= 1000) {
                    if (weight >= 600) style.FontStyle |= OfficeFontStyle.Bold;
                    else style.FontStyle &= ~OfficeFontStyle.Bold;
                }
                else unsupported++;
                break;
            case "text-anchor":
                string anchor = normalized.ToLowerInvariant();
                if (anchor is "start" or "middle" or "end") style.TextAnchor = anchor;
                else unsupported++;
                break;
            case "dominant-baseline":
                string baseline = normalized.ToLowerInvariant();
                if (baseline is "auto" or "alphabetic") style.DominantBaseline = SvgDominantBaseline.Alphabetic;
                else if (baseline is "hanging" or "text-before-edge") style.DominantBaseline = SvgDominantBaseline.Hanging;
                else if (baseline is "middle" or "central") style.DominantBaseline = SvgDominantBaseline.Middle;
                else if (baseline is "text-after-edge" or "ideographic") style.DominantBaseline = SvgDominantBaseline.TextAfterEdge;
                else unsupported++;
                break;
            case "display":
                if (normalized.Equals("none", StringComparison.OrdinalIgnoreCase)) style.Visible = false;
                break;
            case "visibility":
                if (normalized.Equals("hidden", StringComparison.OrdinalIgnoreCase) || normalized.Equals("collapse", StringComparison.OrdinalIgnoreCase)) style.Visible = false;
                break;
            case "transform":
            case "filter":
            case "clip-path":
            case "marker-start":
            case "marker-mid":
            case "marker-end":
                unsupported++;
                break;
        }
    }

    private static void ApplyPaint(OfficeShape shape, SvgPaintContext style) {
        shape.FillColor = style.Fill;
        shape.FillGradient = style.FillGradient;
        shape.FillRadialGradient = style.FillRadialGradient;
        shape.StrokeColor = style.Stroke;
        shape.StrokeGradient = style.StrokeGradient;
        shape.StrokeRadialGradient = style.StrokeRadialGradient;
        shape.StrokeWidth = style.StrokeWidth;
        shape.FillOpacity = style.FillOpacity * style.Opacity;
        shape.StrokeOpacity = style.StrokeOpacity * style.Opacity;
        shape.StrokeDashStyle = style.DashStyle;
        shape.StrokeLineCap = style.LineCap;
        shape.StrokeLineJoin = style.LineJoin;
        shape.FillRule = style.FillRule;
    }

    private static void ApplyDeferredPaint(
        OfficeShape shape,
        SvgPaintContext style,
        double shapeX,
        double shapeY,
        double viewportWidth,
        double viewportHeight,
        double viewX,
        double viewY,
        ref int unsupported) {
        if (style.FillDeferredGradient != null) {
            if (style.FillDeferredGradient.TryCreateForShape(shape, shapeX, shapeY, viewportWidth, viewportHeight, viewX, viewY, out OfficeLinearGradient? linear, out OfficeRadialGradient? radial)) {
                shape.FillGradient = linear;
                shape.FillRadialGradient = radial;
            } else {
                unsupported++;
            }
        }
        if (style.StrokeDeferredGradient != null) {
            if (style.StrokeDeferredGradient.TryCreateForShape(shape, shapeX, shapeY, viewportWidth, viewportHeight, viewX, viewY, out OfficeLinearGradient? linear, out OfficeRadialGradient? radial)) {
                shape.StrokeGradient = linear;
                shape.StrokeRadialGradient = radial;
            } else {
                unsupported++;
            }
        }
    }

    private static double ReadLength(XElement element, string name) => TryLength(element, name, out double value) ? value : 0D;
    private static bool TryLength(XElement element, string name, out double value) => TrySvgLength(element.Attribute(name)?.Value, out value);

    private static double ReadViewportCoordinate(XElement element, string name, double origin, double extent) {
        string? text = element.Attribute(name)?.Value;
        if (!TryViewportLength(text, extent, out double value, out _)) return -origin;
        return value - origin;
    }

    private static double ReadViewportLength(XElement element, string name, double extent) =>
        TryViewportLength(element, name, extent, out double value) ? value : 0D;

    private static bool TryViewportLength(XElement element, string name, double extent, out double value) =>
        TryViewportLength(element.Attribute(name)?.Value, extent, out value, out _);

    private static bool TryViewportLength(string? text, double extent, out double value, out bool percentage) {
        value = 0D;
        percentage = false;
        if (string.IsNullOrWhiteSpace(text)) return false;
        string normalized = text!.Trim();
        percentage = normalized.EndsWith("%", StringComparison.Ordinal);
        if (percentage) normalized = normalized.Substring(0, normalized.Length - 1).Trim();
        else if (normalized.EndsWith("px", StringComparison.OrdinalIgnoreCase)) normalized = normalized.Substring(0, normalized.Length - 2).Trim();
        if (!double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)
            || double.IsNaN(parsed)
            || double.IsInfinity(parsed)) return false;
        value = percentage ? parsed * extent / 100D : parsed;
        return !double.IsNaN(value) && !double.IsInfinity(value);
    }

    private static double ReadFirstLength(XElement element, string name) {
        string? value = element.Attribute(name)?.Value;
        if (string.IsNullOrWhiteSpace(value)) return 0D;
        int separator = value!.IndexOfAny(new[] { ' ', '\t', '\r', '\n', ',' });
        return TrySvgLength(separator < 0 ? value : value.Substring(0, separator), out double parsed) ? parsed : 0D;
    }

    private static bool TrySvgLength(string? value, out double result) {
        result = 0D;
        if (string.IsNullOrWhiteSpace(value)) return false;
        string text = value!.Trim();
        if (text.EndsWith("px", StringComparison.OrdinalIgnoreCase)) text = text.Substring(0, text.Length - 2).Trim();
        return double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out result)
            && !double.IsNaN(result)
            && !double.IsInfinity(result);
    }

    private static bool TryUnit(string value, out double result) =>
        double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out result)
        && !double.IsNaN(result)
        && !double.IsInfinity(result)
        && result >= 0D
        && result <= 1D;

    private static bool TryParseNumberList(string? value, out IReadOnlyList<double> values) =>
        TryParseNumberList(value, int.MaxValue, out values);

    private static bool TryParseNumberList(string? value, int maximumValues,
        out IReadOnlyList<double> values) =>
        TryParseNumberList(value, maximumValues, out values, out _);

    private static bool TryParseNumberList(string? value, int maximumValues,
        out IReadOnlyList<double> values, out bool limitExceeded) {
        var result = new List<double>(Math.Min(maximumValues, 16));
        values = result;
        limitExceeded = false;
        if (maximumValues <= 0 || string.IsNullOrWhiteSpace(value)) return false;
        int index = 0;
        while (index < value!.Length) {
            int separatorStart = index;
            while (index < value.Length && (char.IsWhiteSpace(value[index]) || value[index] == ',')) {
                index++;
                if (index - separatorStart > 128) {
                    limitExceeded = true;
                    return false;
                }
            }
            if (index >= value.Length) break;
            if (result.Count >= maximumValues) {
                limitExceeded = true;
                return false;
            }
            int start = index;
            while (index < value.Length && !char.IsWhiteSpace(value[index]) && value[index] != ',') {
                index++;
                if (index - start > 128) {
                    limitExceeded = true;
                    return false;
                }
            }
            int length = index - start;
            if (length <= 0
                || !double.TryParse(value.Substring(start, length), NumberStyles.Float,
                    CultureInfo.InvariantCulture, out double number)
                || double.IsNaN(number)
                || double.IsInfinity(number)) return false;
            result.Add(number);
        }
        return result.Count > 0;
    }

    private struct SvgPaintContext {
        internal OfficeColor Color;
        internal OfficeColor? Fill;
        internal OfficeLinearGradient? FillGradient;
        internal OfficeRadialGradient? FillRadialGradient;
        internal SvgGradientDefinition? FillDeferredGradient;
        internal OfficeColor? Stroke;
        internal OfficeLinearGradient? StrokeGradient;
        internal OfficeRadialGradient? StrokeRadialGradient;
        internal SvgGradientDefinition? StrokeDeferredGradient;
        internal double StrokeWidth;
        internal double Opacity;
        internal double FillOpacity;
        internal double StrokeOpacity;
        internal OfficeStrokeDashStyle DashStyle;
        internal OfficeStrokeLineCap LineCap;
        internal OfficeStrokeLineJoin LineJoin;
        internal OfficeFillRule FillRule;
        internal string FontFamily;
        internal double FontSize;
        internal OfficeFontStyle FontStyle;
        internal string TextAnchor;
        internal SvgDominantBaseline DominantBaseline;
        internal bool Visible;

        internal void SetFill(SvgResolvedPaint paint) {
            Fill = paint.Color;
            FillGradient = paint.LinearGradient;
            FillRadialGradient = paint.RadialGradient;
            FillDeferredGradient = paint.DeferredGradient;
        }

        internal void SetStroke(SvgResolvedPaint paint) {
            Stroke = paint.Color;
            StrokeGradient = paint.LinearGradient;
            StrokeRadialGradient = paint.RadialGradient;
            StrokeDeferredGradient = paint.DeferredGradient;
        }

        internal static SvgPaintContext Default => new SvgPaintContext {
            Color = OfficeColor.Black,
            Fill = OfficeColor.Black,
            Stroke = null,
            StrokeWidth = 1D,
            Opacity = 1D,
            FillOpacity = 1D,
            StrokeOpacity = 1D,
            DashStyle = OfficeStrokeDashStyle.Solid,
            LineCap = OfficeStrokeLineCap.Butt,
            LineJoin = OfficeStrokeLineJoin.Miter,
            FillRule = OfficeFillRule.NonZero,
            FontFamily = "Arial",
            FontSize = 16D,
            FontStyle = OfficeFontStyle.Regular,
            TextAnchor = "start",
            DominantBaseline = SvgDominantBaseline.Alphabetic,
            Visible = true
        };
    }

    private enum SvgDominantBaseline {
        Alphabetic,
        Hanging,
        Middle,
        TextAfterEdge
    }
}
