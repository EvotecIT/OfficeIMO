using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgDrawingReader {
    private const int MaximumElementReferenceDepth = 16;

    private enum SvgElementReferenceEntryResult {
        Invalid,
        Cycle,
        DepthExceeded,
        Entered
    }

    private static void AddReferencedElement(
        XElement use,
        OfficeDrawing drawing,
        SvgPaintContext style,
        SvgPaintServerRegistry paintServers,
        SvgElementReferenceRegistry references,
        OfficeTransform transform,
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
        if (!references.TryEnter(use, out string referenceId, out XElement? target)) {
            unsupported++;
            return;
        }

        try {
            string targetName = target!.Name.LocalName.ToLowerInvariant();
            if (targetName is "defs" or "lineargradient" or "radialgradient" or "stop") {
                unsupported++;
                return;
            }
            if (targetName == "symbol") {
                AddReferencedSymbol(use, target, drawing, style, paintServers, references, transform,
                    maximumElements, maximumViewportDimension, maximumViewportPixels, depth,
                    ref visited, ref pathCommands, ref pathCommandLimitExceeded, ref unsupported);
                return;
            }
            if (!TryOptionalUseLength(use, "x", out double x) || !TryOptionalUseLength(use, "y", out double y)) {
                unsupported++;
                return;
            }

            OfficeTransform placement = OfficeTransform.Translate(x, y).Then(transform);
            AddElement(target, drawing, style, paintServers, references, placement, viewX, viewY,
                maximumElements, maximumViewportDimension, maximumViewportPixels, depth,
                ref visited, ref pathCommands, ref pathCommandLimitExceeded, ref unsupported);
        } finally {
            references.Exit(referenceId);
        }
    }

    private static bool TryOptionalUseLength(XElement use, string name, out double value) {
        string? text = use.Attribute(name)?.Value;
        if (string.IsNullOrWhiteSpace(text)) {
            value = 0D;
            return true;
        }
        return TrySvgLength(text, out value);
    }

    private sealed class SvgElementReferenceRegistry {
        private readonly SvgDefinitionRegistry _definitions;
        private readonly ISet<string> _activeIds = new HashSet<string>(StringComparer.Ordinal);
        private const int MaximumExpandedTextCharacters = 131_072;
        private int _expandedTextCharacters;
        private const double MaximumIntermediateSurfacePixels = 64_000_000D;
        private double _intermediateSurfacePixels;
        private const long MaximumEmbeddedRasterBytes = 64L * 1024L * 1024L;
        private long _embeddedRasterBytes;
        private readonly Dictionary<XAttribute, (byte[] Bytes, string ContentType, OfficeImageInfo Info)> _embeddedRasters =
            new Dictionary<XAttribute, (byte[] Bytes, string ContentType, OfficeImageInfo Info)>();
        private const int MaximumForeignObjectRenderCalls = 128;
        private const long MaximumForeignObjectSourceCharacters = 8L * 1024L * 1024L;
        private const double MaximumForeignObjectPixels = 64_000_000D;
        private int _foreignObjectRenderCalls;
        private long _foreignObjectSourceCharacters;
        private long _foreignObjectSerializedCharacters;
        private double _foreignObjectPixels;
        private const double MaximumMarkerScenePixels = 64_000_000D;
        private double _markerScenePixels;
        private readonly Dictionary<(XElement Element, double Width, double Height), (OfficeDrawing Drawing, int Elements)> _foreignObjects =
            new Dictionary<(XElement Element, double Width, double Height), (OfficeDrawing Drawing, int Elements)>();
        private readonly Dictionary<XElement, bool> _foreignObjectHasContent = new Dictionary<XElement, bool>();
        internal OfficeCffOperationBudget CffOperationBudget { get; } = new OfficeCffOperationBudget();

        internal SvgElementReferenceRegistry(
            SvgDefinitionRegistry definitions,
            OfficeSvgForeignObjectRenderer? foreignObjectRenderer = null) {
            _definitions = definitions;
            ForeignObjectRenderer = foreignObjectRenderer;
        }

        internal OfficeSvgForeignObjectRenderer? ForeignObjectRenderer { get; }

        internal XNamespace NativeNamespace => _definitions.NativeNamespace;

        internal bool CanChargeTextCharacters(int count) =>
            count >= 0 && count <= MaximumExpandedTextCharacters - _expandedTextCharacters;

        internal bool TryChargeTextCharacters(int count) {
            if (!CanChargeTextCharacters(count)) return false;
            _expandedTextCharacters += count;
            return true;
        }

        internal bool TryChargeNestedViewport(double width, double height, double sceneWidth, double sceneHeight) {
            // The viewBox scene can be much larger than its displayed viewport.
            // Raster rendering retains that full scene before fitting. The clipped
            // group draws into its parent without another viewport-sized surface.
            return TryChargeIntermediatePixels(width * height + sceneWidth * sceneHeight);
        }

        internal bool TryChargeNestedViewportExpansion(double extraPixels) {
            return !double.IsNaN(extraPixels) && !double.IsInfinity(extraPixels)
                && (extraPixels <= 0D || TryChargeIntermediatePixels(extraPixels));
        }

        internal bool TryChargeIntermediateSurface(double width, double height, int surfaces = 1) =>
            TryChargeIntermediatePixels(width * height * surfaces);

        internal bool TryChargeEffectSurfaces(double width, double height, bool hasSoftMask) =>
            TryChargeIntermediateSurface(width, height, hasSoftMask ? 4 : 1);

        internal bool TryChargeEmbeddedImage(double canvasWidth, double canvasHeight,
            int imageWidth, int imageHeight, double opacity) =>
            TryChargeIntermediatePixels(canvasWidth * canvasHeight +
                (double)imageWidth * imageHeight * (opacity < 1D ? 2D : 1D));

        private bool TryChargeIntermediatePixels(double pixels) {
            // Full-size effect, image, viewport, and symbol layers all reach the
            // raster renderer; count their retained surfaces in one document budget.
            if (double.IsNaN(pixels) || double.IsInfinity(pixels) || pixels < 0D
                || pixels > MaximumIntermediateSurfacePixels - _intermediateSurfacePixels) return false;
            _intermediateSurfacePixels += pixels;
            return true;
        }

        internal bool TryGetEmbeddedRaster(XAttribute source, out byte[] bytes, out string contentType, out OfficeImageInfo info) {
            if (_embeddedRasters.TryGetValue(source, out var cached)) {
                bytes = cached.Bytes;
                contentType = cached.ContentType;
                info = cached.Info;
                return true;
            }
            bytes = Array.Empty<byte>();
            contentType = string.Empty;
            info = null!;
            return false;
        }

        internal bool TryCacheEmbeddedRaster(XAttribute source, byte[] bytes, string contentType, OfficeImageInfo info) {
            if (bytes.Length > MaximumEmbeddedRasterBytes - _embeddedRasterBytes) return false;
            _embeddedRasterBytes += bytes.Length;
            _embeddedRasters.Add(source, (bytes, contentType, info));
            return true;
        }

        internal bool TryGetForeignObject(XElement element, double width, double height, out OfficeDrawing drawing, out int elements) {
            if (_foreignObjects.TryGetValue((element, width, height), out var cached)) {
                drawing = cached.Drawing;
                elements = cached.Elements;
                return true;
            }
            drawing = null!;
            elements = 0;
            return false;
        }

        internal bool HasForeignObjectContent(XElement element) {
            if (_foreignObjectHasContent.TryGetValue(element, out bool hasContent)) return hasContent;
            hasContent = element.Nodes().Any(node => node is XElement || node is XText text && !string.IsNullOrWhiteSpace(text.Value));
            _foreignObjectHasContent[element] = hasContent;
            return hasContent;
        }

        internal bool TryReserveForeignObject(XElement element, double width, double height) {
            double pixels = width * height;
            if (_foreignObjectRenderCalls >= MaximumForeignObjectRenderCalls
                || double.IsNaN(pixels) || double.IsInfinity(pixels) || pixels < 0D
                || pixels > MaximumForeignObjectPixels - _foreignObjectPixels) return false;
            long characters = 0;
            foreach (XElement node in element.DescendantsAndSelf()) {
                characters += node.Name.LocalName.Length + 4;
                foreach (XAttribute attribute in node.Attributes()) characters += attribute.Name.LocalName.Length + attribute.Value.Length + 4;
                foreach (XText text in node.Nodes().OfType<XText>()) characters += text.Value.Length;
                if (characters > MaximumForeignObjectSourceCharacters - _foreignObjectSourceCharacters) return false;
            }
            _foreignObjectRenderCalls++;
            _foreignObjectSourceCharacters += characters;
            _foreignObjectPixels += pixels;
            return true;
        }

        internal void CacheForeignObject(XElement element, double width, double height, OfficeDrawing drawing, int elements) =>
            _foreignObjects[(element, width, height)] = (drawing, elements);

        internal bool TryChargeForeignObjectPlacement(double drawingWidth, double drawingHeight) {
            // Every placement creates a full-size effect surface, including cache hits.
            return TryChargeIntermediateSurface(drawingWidth, drawingHeight);
        }

        internal bool TryChargeSerializedForeignObject(int characters) {
            if (characters < 0 || characters > MaximumForeignObjectSourceCharacters - _foreignObjectSerializedCharacters) return false;
            _foreignObjectSerializedCharacters += characters;
            return true;
        }

        internal bool TryChargeMarkerScene(double width, double height) {
            double pixels = width * height;
            if (double.IsNaN(pixels) || double.IsInfinity(pixels) || pixels <= 0D
                || pixels > MaximumMarkerScenePixels - _markerScenePixels) return false;
            _markerScenePixels += pixels;
            return true;
        }

        internal bool TryEnter(XElement use, out string id, out XElement? target) {
            return TryEnterDetailed(use, out id, out target) == SvgElementReferenceEntryResult.Entered;
        }

        internal SvgElementReferenceEntryResult TryEnterDetailed(XElement use, out string id, out XElement? target) {
            return TryEnterDetailed(use, expectedTargetName: null, out id, out target);
        }

        internal SvgElementReferenceEntryResult TryEnterDetailed(
            XElement use,
            string? expectedTargetName,
            out string id,
            out XElement? target) {
            id = string.Empty;
            target = null;
            XAttribute[] hrefAttributes = use.Attributes()
                .Where(attribute => attribute.Name.LocalName.Equals("href", StringComparison.Ordinal) &&
                    (attribute.Name.NamespaceName.Length == 0 ||
                     attribute.Name.NamespaceName.Equals("http://www.w3.org/1999/xlink", StringComparison.Ordinal)))
                .Take(2)
                .ToArray();
            if (hrefAttributes.Length != 1
                || !TryReadLocalElementReference(hrefAttributes[0].Value, out id)
                || !_definitions.TryGetUnique(id, out target)
                || (expectedTargetName != null
                    && !target!.Name.LocalName.Equals(expectedTargetName, StringComparison.OrdinalIgnoreCase))) {
                return SvgElementReferenceEntryResult.Invalid;
            }
            if (_activeIds.Contains(id)) return SvgElementReferenceEntryResult.Cycle;
            if (_activeIds.Count >= MaximumElementReferenceDepth) return SvgElementReferenceEntryResult.DepthExceeded;
            _activeIds.Add(id);
            return SvgElementReferenceEntryResult.Entered;
        }

        internal void Exit(string id) {
            if (id.Length > 0) _activeIds.Remove(id);
        }

        internal bool TryEnterLocal(string? value, out string id, out XElement? target) {
            return TryEnterLocalDetailed(value, out id, out target) == SvgElementReferenceEntryResult.Entered;
        }

        internal SvgElementReferenceEntryResult TryEnterLocalDetailed(string? value, out string id, out XElement? target) {
            return TryEnterLocalDetailed(value, expectedTargetName: null, out id, out target);
        }

        internal SvgElementReferenceEntryResult TryEnterLocalDetailed(
            string? value,
            string? expectedTargetName,
            out string id,
            out XElement? target) {
            id = string.Empty;
            target = null;
            if (!TryReadLocalUrlReference(value, out id)
                || !_definitions.TryGetUnique(id, out target)
                || (expectedTargetName != null
                    && !target!.Name.LocalName.Equals(expectedTargetName, StringComparison.OrdinalIgnoreCase))) {
                return SvgElementReferenceEntryResult.Invalid;
            }
            if (_activeIds.Contains(id)) return SvgElementReferenceEntryResult.Cycle;
            if (_activeIds.Count >= MaximumElementReferenceDepth) return SvgElementReferenceEntryResult.DepthExceeded;
            _activeIds.Add(id);
            return SvgElementReferenceEntryResult.Entered;
        }

        private static bool TryReadLocalElementReference(string text, out string id) {
            id = string.Empty;
            string normalized = TrimSvgCssWhitespace(text);
            if (normalized.Length < 2 || normalized[0] != '#') return false;
            string encodedId = normalized.Substring(1);
            if (encodedId.Length == 0 || encodedId.IndexOfAny(new[] { ' ', '\t', '\r', '\n', '#', '(', ')' }) >= 0) return false;
            try {
                id = Uri.UnescapeDataString(encodedId);
            } catch (UriFormatException) {
                return false;
            }
            return id.Length > 0;
        }

        private static bool TryReadLocalUrlReference(string? text, out string id) {
            id = string.Empty;
            if (string.IsNullOrWhiteSpace(text)) return false;
            if (!TryReadSvgLocalUrlReference(text!, out string reference)) return false;
            return TryReadLocalElementReference(reference, out id);
        }
    }
}
