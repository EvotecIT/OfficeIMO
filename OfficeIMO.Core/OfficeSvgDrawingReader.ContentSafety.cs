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
            MaxCharactersInDocument = Math.Min(options.MaxInputBytes, MaximumInputBytes),
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
        if (root.Name.NamespaceName.Length > 0 &&
            !root.Name.NamespaceName.Equals("http://www.w3.org/2000/svg", StringComparison.Ordinal)) {
            throw new InvalidDataException("The SVG root must use the standard SVG namespace or no namespace.");
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
            throw new InvalidDataException("The cleaned SVG exceeds the configured output-byte limit.");
        }
        return output.ToArray();
    }
}
