using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using OfficeIMO.ContentSafety;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgDrawingReader {
    /// <summary>Inspects bounded SVG text for native, computed, geometric, clipping, compositing, and contrast concealment.</summary>
    public static OfficeContentSafetyReport InspectContentSafety(
        byte[] svgBytes,
        OfficeContentSafetyOptions? options = null,
        OfficeSvgDrawingReaderOptions? readerOptions = null) {
        if (svgBytes == null) throw new ArgumentNullException(nameof(svgBytes));
        OfficeContentSafetyOptions effective = options ?? new OfficeContentSafetyOptions();
        SvgContentSafetyDocument document = LoadSvgContentSafetyDocument(svgBytes, effective, readerOptions);
        return InspectSvgContentSafetyDocument(document, svgBytes, effective, readerOptions, targets: null);
    }

    /// <summary>Inspects a bounded SVG file without treating concealment as evidence of AI authorship.</summary>
    public static OfficeContentSafetyReport InspectContentSafety(
        string filePath,
        OfficeContentSafetyOptions? options = null,
        OfficeSvgDrawingReaderOptions? readerOptions = null) {
        if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("A file path is required.", nameof(filePath));
        OfficeContentSafetyOptions effective = options ?? new OfficeContentSafetyOptions();
        return InspectContentSafety(OfficeContentSafetyInputGuard.ReadAllBytes(filePath, effective), effective, readerOptions);
    }

    /// <summary>Removes only exact current SVG findings selected by the caller and then reopens and reinspects the result.</summary>
    public static OfficeContentCleanupResult RemoveSelectedContent(
        byte[] svgBytes,
        OfficeContentCleanupSelection selection,
        OfficeContentCleanupOptions? options = null,
        OfficeSvgDrawingReaderOptions? readerOptions = null) {
        if (svgBytes == null) throw new ArgumentNullException(nameof(svgBytes));
        if (selection == null) throw new ArgumentNullException(nameof(selection));
        options ??= new OfficeContentCleanupOptions();
        options.Validate();

        OfficeContentSafetyReport before = InspectContentSafety(svgBytes, options.Inspection, readerOptions);
        IReadOnlyList<OfficeContentSafetyFinding> selected = OfficeContentSafetyBuilder.ResolveSelection(before, selection);
        if (selected.Count == 0) {
            return new OfficeContentCleanupResult(
                (byte[])svgBytes.Clone(),
                before,
                before,
                Array.Empty<OfficeContentCleanupChange>());
        }

        SvgContentSafetyDocument document = LoadSvgContentSafetyDocument(svgBytes, options.Inspection, readerOptions);
        var targets = new Dictionary<string, SvgContentSafetyTarget>(StringComparer.Ordinal);
        OfficeContentSafetyReport current = InspectSvgContentSafetyDocument(document, svgBytes, options.Inspection, readerOptions, targets);
        IReadOnlyList<OfficeContentSafetyFinding> currentSelection = OfficeContentSafetyBuilder.ResolveSelection(current, selection);
        ApplySvgSignatureMutationPolicy(document.Document, options.SignatureMutationPolicy);
        foreach (IGrouping<SvgContentSafetyTarget, OfficeContentSafetyFinding> group in currentSelection
            .OrderByDescending(item => item.SourceTextOffset ?? -1)
            .GroupBy(item => targets[item.Id])) {
            group.Key.Remove(group);
        }

        byte[] output = SerializeSvgContentSafetyDocument(
            document.Document,
            Math.Min(options.Inspection.MaxInputBytes, MaximumInputBytes));
        OfficeContentSafetyReport after = InspectContentSafety(output, options.Inspection, readerOptions);
        OfficeContentCleanupChange[] changes = selected
            .Select(item => new OfficeContentCleanupChange(item.Id, item.Location, item.CleanupCapability))
            .ToArray();
        return new OfficeContentCleanupResult(output, before, after, changes);
    }

    private static void ApplySvgSignatureMutationPolicy(
        XDocument document,
        OfficeSignatureMutationPolicy policy) {
        XNamespace xmlDsig = "http://www.w3.org/2000/09/xmldsig#";
        XElement[] signatures = document.Descendants(xmlDsig + "Signature").ToArray();
        if (signatures.Length == 0) return;
        if (policy == OfficeSignatureMutationPolicy.BlockSave) {
            throw new InvalidOperationException(
                "Cleaning SVG content would invalidate an XML digital signature. Choose an explicit signature mutation policy.");
        }
        if (policy == OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures) {
            foreach (XElement signature in signatures) signature.Remove();
        }
    }

    /// <summary>Atomically writes an explicitly cleaned SVG artifact.</summary>
    public static OfficeContentCleanupResult RemoveSelectedContent(
        string inputPath,
        string outputPath,
        OfficeContentCleanupSelection selection,
        OfficeContentCleanupOptions? options = null,
        OfficeSvgDrawingReaderOptions? readerOptions = null) {
        if (string.IsNullOrWhiteSpace(inputPath)) throw new ArgumentException("An input path is required.", nameof(inputPath));
        if (string.IsNullOrWhiteSpace(outputPath)) throw new ArgumentException("An output path is required.", nameof(outputPath));
        options ??= new OfficeContentCleanupOptions();
        options.Validate();
        byte[] input = OfficeContentSafetyInputGuard.ReadAllBytes(inputPath, options.Inspection);
        OfficeContentCleanupResult result = RemoveSelectedContent(input, selection, options, readerOptions);
        OfficeFileCommit.WriteAllBytes(outputPath, result.Output);
        return result;
    }

    private static SvgContentSafetyDocument LoadSvgContentSafetyDocument(
        byte[] svgBytes,
        OfficeContentSafetyOptions options,
        OfficeSvgDrawingReaderOptions? readerOptions) {
        OfficeContentSafetyInputGuard.ValidateBytes(svgBytes, options);
        if (!TryReadBoundedDocument(
                svgBytes,
                readerOptions,
                allowUnresolvedViewport: false,
                maximumCharactersInDocument: options.MaxCharacters,
                out _,
                out int maximumElements,
                out double maximumViewportDimension,
                out double maximumViewportPixels,
                out double viewX,
                out double viewY,
                out double viewWidth,
                out double viewHeight,
                out double viewportWidth,
                out double viewportHeight)) {
            throw new InvalidDataException("The SVG is malformed or exceeds the bounded SVG parser or viewport limits.");
        }

        var settings = new XmlReaderSettings {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            MaxCharactersInDocument = Math.Min(options.MaxCharacters, Math.Min(options.MaxInputBytes, MaximumInputBytes)),
            MaxCharactersFromEntities = 0,
            IgnoreWhitespace = false
        };
        XDocument parsed;
        using (var stream = new MemoryStream(svgBytes, writable: false))
        using (XmlReader reader = XmlReader.Create(stream, settings)) {
            parsed = XDocument.Load(reader, LoadOptions.PreserveWhitespace | LoadOptions.SetLineInfo);
        }
        if (parsed.DescendantNodes().OfType<XProcessingInstruction>().Any(instruction =>
                instruction.Target.Equals("xml-stylesheet", StringComparison.OrdinalIgnoreCase))) {
            throw new InvalidDataException("The SVG uses an XML stylesheet processing instruction outside the bounded native subset.");
        }
        XElement root = parsed.Root ?? throw new InvalidDataException("The SVG has no root element.");
        if (!root.Name.LocalName.Equals("svg", StringComparison.Ordinal)) {
            throw new InvalidDataException("The SVG root element name must match XML SVG casing.");
        }
        if (root.Name.NamespaceName.Length > 0 &&
            !root.Name.NamespaceName.Equals("http://www.w3.org/2000/svg", StringComparison.Ordinal)) {
            throw new InvalidDataException("The SVG root must use the standard SVG namespace or no namespace.");
        }
        if (root.Attributes().Any(IsInvalidSvgRootViewportAttribute)) {
            throw new InvalidDataException("The SVG uses a namespaced or case-mismatched root viewport attribute outside XML SVG geometry.");
        }
        XNamespace svgNamespace = root.Name.Namespace;
        if (root.DescendantsAndSelf()
                .Where(element => IsNativeSvgElement(element, svgNamespace))
                .Any(HasUnsupportedSvgViewportSyntax)) {
            throw new InvalidDataException("The SVG uses viewport syntax outside the bounded case-sensitive SVG grammar.");
        }
        if (root.DescendantsAndSelf()
                .Where(element => IsNativeSvgElement(element, svgNamespace))
                .Any(HasUnsupportedSvgTransformSyntax)) {
            throw new InvalidDataException("The SVG uses transform syntax outside the bounded SVG grammar.");
        }
        if (root.DescendantsAndSelf().Where(element => IsNativeSvgElement(element, svgNamespace)).Any(element =>
                element.Attributes().Any(attribute =>
                    attribute.Name.NamespaceName.Length == 0 &&
                    attribute.Name.LocalName.Equals(attribute.Name.LocalName.ToLowerInvariant(), StringComparison.Ordinal) &&
                    IsSvgPresentationPropertyName(attribute.Name.LocalName) &&
                    IsUnsupportedSvgCssWideKeyword(attribute.Value)))) {
            throw new InvalidDataException("The SVG uses unsupported revert cascade semantics in a presentation attribute.");
        }
        if (root.DescendantsAndSelf().Where(element => IsNativeSvgElement(element, svgNamespace)).Any(element =>
                element.Attributes().Any(attribute =>
                    attribute.Name.NamespaceName.Length == 0 &&
                    attribute.Name.LocalName.Equals(attribute.Name.LocalName.ToLowerInvariant(), StringComparison.Ordinal) &&
                    IsSvgPresentationPropertyName(attribute.Name.LocalName) &&
                    attribute.Value.IndexOf('\\') >= 0))) {
            throw new InvalidDataException("The SVG uses escaped presentation-attribute syntax outside the bounded native CSS subset.");
        }
        if (root.DescendantsAndSelf().Where(element => IsNativeSvgElement(element, svgNamespace)).Any(element =>
                element.Attributes().Any(attribute =>
                    attribute.Name.NamespaceName.Length == 0 &&
                    attribute.Name.LocalName.Equals(attribute.Name.LocalName.ToLowerInvariant(), StringComparison.Ordinal) &&
                    IsSvgPresentationPropertyName(attribute.Name.LocalName) &&
                    ContainsNonSvgCssWhitespace(attribute.Value)))) {
            throw new InvalidDataException("The SVG uses non-CSS whitespace in a presentation attribute outside the bounded native CSS subset.");
        }
        if (root.DescendantsAndSelf().Where(element => IsNativeSvgElement(element, svgNamespace)).Any(element =>
                element.Attributes().Any(attribute =>
                    attribute.Name.NamespaceName.Length == 0 &&
                    attribute.Name.LocalName.Equals(attribute.Name.LocalName.ToLowerInvariant(), StringComparison.Ordinal) &&
                    IsSvgPresentationPropertyName(attribute.Name.LocalName) &&
                    ContainsUnsupportedSvgCssMathFunction(attribute.Value)))) {
            throw new InvalidDataException("The SVG uses CSS math functions outside the bounded native presentation-attribute subset.");
        }
        if (root.DescendantsAndSelf().Where(element => IsNativeSvgElement(element, svgNamespace))
            .Any(HasUnsupportedSvgTextPositioningLength)) {
            throw new InvalidDataException("The SVG uses relative or unsupported text-positioning lengths outside the bounded native layout subset.");
        }
        if (ExceedsSvgElementNestingLimit(root)) {
            throw new InvalidDataException("The SVG exceeds the bounded element-nesting limit.");
        }
        int maximumVisualComparisons = readerOptions?.MaximumContentSafetyVisualComparisons ??
            OfficeSvgDrawingReaderOptions.DefaultMaximumContentSafetyVisualComparisons;
        long maximumVisualPixels = readerOptions?.MaximumContentSafetyVisualPixels ??
            OfficeSvgDrawingReaderOptions.DefaultMaximumContentSafetyVisualPixels;
        if (maximumVisualComparisons < 0 ||
            maximumVisualComparisons > OfficeSvgDrawingReaderOptions.MaximumAllowedContentSafetyVisualComparisons) {
            throw new ArgumentOutOfRangeException(nameof(readerOptions), "The SVG content-safety comparison count is outside the supported range.");
        }
        if (maximumVisualPixels <= 0 ||
            maximumVisualPixels > OfficeSvgDrawingReaderOptions.MaximumAllowedContentSafetyVisualPixels) {
            throw new ArgumentOutOfRangeException(nameof(readerOptions), "The SVG content-safety pixel-work budget is outside the supported range.");
        }
        return new SvgContentSafetyDocument(
            parsed,
            root,
            maximumElements,
            maximumViewportDimension,
            maximumViewportPixels,
            viewX,
            viewY,
            viewWidth,
            viewHeight,
            viewportWidth,
            viewportHeight,
            maximumVisualComparisons,
            maximumVisualPixels);
    }

    private static bool HasUnsupportedSvgTextPositioningLength(XElement element) {
        if (!element.AncestorsAndSelf().Any(ancestor => ancestor.Name.LocalName.Equals("text", StringComparison.Ordinal))) {
            return false;
        }
        foreach (string name in new[] { "x", "y", "dx", "dy" }) {
            string? value = element.Attribute(name)?.Value;
            if (string.IsNullOrWhiteSpace(value)) continue;
            string[] tokens = value!.Split(new[] { ' ', '\t', '\r', '\n', ',' }, StringSplitOptions.RemoveEmptyEntries);
            if (tokens.Length == 0 || tokens.Length > MaximumTextRuns) return true;
            foreach (string token in tokens) {
                if (!TryViewportLength(token, 1D, out _, out _)) return true;
            }
        }
        return false;
    }

    private static bool IsInvalidSvgRootViewportAttribute(XAttribute attribute) {
        if (attribute.IsNamespaceDeclaration) return false;
        string? expected = attribute.Name.LocalName.ToLowerInvariant() switch {
            "width" => "width",
            "height" => "height",
            "viewbox" => "viewBox",
            "preserveaspectratio" => "preserveAspectRatio",
            _ => null
        };
        return expected != null &&
            (attribute.Name.NamespaceName.Length != 0 || !attribute.Name.LocalName.Equals(expected, StringComparison.Ordinal));
    }

    private static bool HasUnsupportedSvgViewportSyntax(XElement element) {
        string name = element.Name.LocalName;
        bool acceptsViewBox = name is "svg" or "symbol" or "marker" or "pattern" or "view";
        bool acceptsPreserveAspectRatio = acceptsViewBox || name.Equals("image", StringComparison.Ordinal);
        XAttribute? viewBox = acceptsViewBox ? element.Attribute("viewBox") : null;
        if (viewBox != null &&
            (ContainsNonSvgCssWhitespace(viewBox.Value) ||
             !TryParseNumberList(viewBox.Value, out IReadOnlyList<double> values) ||
             values.Count != 4)) return true;
        XAttribute? preserveAspectRatio = acceptsPreserveAspectRatio ? element.Attribute("preserveAspectRatio") : null;
        return preserveAspectRatio != null &&
            !TryParsePreserveAspectRatio(preserveAspectRatio.Value, out _, out _);
    }

    private static bool HasUnsupportedSvgTransformSyntax(XElement element) {
        foreach (string name in new[] { "transform", "gradientTransform", "patternTransform" }) {
            XAttribute? attribute = element.Attribute(name);
            if (attribute != null && !OfficeSvgTransformParser.TryParse(attribute.Value, out _)) return true;
        }
        return false;
    }

    private static bool ContainsUnsupportedSvgCssMathFunction(string value) {
        foreach (string name in new[] {
            "calc", "min", "max", "clamp", "round", "mod", "rem", "sin", "cos", "tan",
            "asin", "acos", "atan", "atan2", "pow", "sqrt", "hypot", "log", "exp", "abs", "sign"
        }) {
            if (ContainsPotentialCssIdentifier(value, name)) return true;
        }
        return false;
    }

    private static byte[] SerializeSvgContentSafetyDocument(XDocument document, long maximumBytes) {
        using var output = new MemoryStream();
        var settings = new XmlWriterSettings {
            Encoding = new UTF8Encoding(false),
            Indent = false,
            OmitXmlDeclaration = document.Declaration == null,
            NewLineHandling = NewLineHandling.None
        };
        using (XmlWriter writer = XmlWriter.Create(output, settings)) document.Save(writer);
        if (output.Length > maximumBytes) {
            throw new InvalidDataException("The serialized SVG exceeds the configured output-byte limit.");
        }
        return output.ToArray();
    }
}
