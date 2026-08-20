using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.ContentSafety;
using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using OfficeIMO.Provenance;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

public partial class WordDocument {
    /// <summary>Inspects DOCX-family content that remains machine-readable while hidden or outside the primary document story.</summary>
    public static OfficeContentSafetyReport InspectContentSafety(
        string filePath,
        OfficeContentSafetyOptions? options = null) {
        if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("A file path is required.", nameof(filePath));
        OfficeContentSafetyOptions effective = options ?? new OfficeContentSafetyOptions();
        return InspectContentSafety(OfficeContentSafetyInputGuard.ReadAllBytes(filePath, effective, inspectZipPackage: true), effective);
    }

    /// <summary>Inspects encoded DOCX, DOCM, DOTX, or DOTM bytes.</summary>
    public static OfficeContentSafetyReport InspectContentSafety(
        byte[] documentBytes,
        OfficeContentSafetyOptions? options = null) {
        if (documentBytes == null) throw new ArgumentNullException(nameof(documentBytes));
        OfficeContentSafetyOptions effective = options ?? new OfficeContentSafetyOptions();
        OfficeContentSafetyInputGuard.ValidateBytes(documentBytes, effective, inspectZipPackage: true);
        using var stream = new MemoryStream(documentBytes, writable: false);
        using WordprocessingDocument document = WordprocessingDocument.Open(stream, false);
        return InspectContentSafetyDocument(document, effective, targets: null);
    }

    /// <summary>Removes exact selected concealed-content findings and reinspects the encoded document.</summary>
    public static OfficeContentCleanupResult RemoveSelectedContent(
        byte[] documentBytes,
        OfficeContentCleanupSelection selection,
        OfficeContentCleanupOptions? options = null) {
        if (documentBytes == null) throw new ArgumentNullException(nameof(documentBytes));
        if (selection == null) throw new ArgumentNullException(nameof(selection));
        options ??= new OfficeContentCleanupOptions();
        options.Validate();
        OfficeContentSafetyReport before = InspectContentSafety(documentBytes, options.Inspection);
        IReadOnlyList<OfficeContentSafetyFinding> selected = OfficeContentSafetyBuilder.ResolveSelection(before, selection);
        if (selected.Count == 0) return new OfficeContentCleanupResult((byte[])documentBytes.Clone(), before, before, Array.Empty<OfficeContentCleanupChange>());

        byte[] mutableBytes = PrepareContentSafetyMutation(documentBytes, options.SignatureMutationPolicy);
        using var stream = new MemoryStream(mutableBytes.Length + 4096);
        stream.Write(mutableBytes, 0, mutableBytes.Length);
        stream.Position = 0;
        using (WordprocessingDocument document = WordprocessingDocument.Open(stream, true)) {
            var targets = new Dictionary<string, WordCleanupTarget>(StringComparer.Ordinal);
            OfficeContentSafetyReport current = InspectContentSafetyDocument(document, options.Inspection, targets);
            IReadOnlyList<OfficeContentSafetyFinding> currentSelection = OfficeContentSafetyBuilder.ResolveSelection(current, selection);
            foreach (IGrouping<WordCleanupTarget, OfficeContentSafetyFinding> group in currentSelection
                .OrderByDescending(item => item.SourceTextOffset ?? -1)
                .GroupBy(item => targets[item.Id])) {
                group.Key.Remove();
            }
        }
        byte[] output = stream.ToArray();
        OfficeContentSafetyReport after = InspectContentSafety(output, options.Inspection);
        OfficeContentCleanupChange[] changes = selected.Select(item => new OfficeContentCleanupChange(item.Id, item.Location, item.CleanupCapability)).ToArray();
        return new OfficeContentCleanupResult(output, before, after, changes);
    }

    /// <summary>Atomically writes an explicitly cleaned DOCX-family artifact.</summary>
    public static OfficeContentCleanupResult RemoveSelectedContent(
        string inputPath,
        string outputPath,
        OfficeContentCleanupSelection selection,
        OfficeContentCleanupOptions? options = null) {
        if (string.IsNullOrWhiteSpace(inputPath)) throw new ArgumentException("An input path is required.", nameof(inputPath));
        if (string.IsNullOrWhiteSpace(outputPath)) throw new ArgumentException("An output path is required.", nameof(outputPath));
        options ??= new OfficeContentCleanupOptions();
        options.Validate();
        OfficeContentCleanupResult result = RemoveSelectedContent(OfficeContentSafetyInputGuard.ReadAllBytes(inputPath, options.Inspection, inspectZipPackage: true), selection, options);
        OfficeFileCommit.WriteAllBytes(outputPath, result.Output);
        return result;
    }

    private static OfficeContentSafetyReport InspectContentSafetyDocument(
        WordprocessingDocument document,
        OfficeContentSafetyOptions? options,
        IDictionary<string, WordCleanupTarget>? targets) {
        MainDocumentPart main = document.MainDocumentPart ?? throw new InvalidDataException("The package has no Word main document part.");
        var builder = new OfficeContentSafetyBuilder("Word Open XML", options);
        WordStyleResolver styleResolver = new WordStyleResolver(main);
        W.Document mainDocument = main.Document ?? throw new InvalidDataException("The Word main document part has no document root.");
        InspectWordRoot(mainDocument, "Document", false, styleResolver, builder, targets);
        int index = 0;
        foreach (HeaderPart part in main.HeaderParts) if (part.Header != null) InspectWordRoot(part.Header, "Header[" + (++index).ToString(CultureInfo.InvariantCulture) + "]", false, styleResolver, builder, targets);
        index = 0;
        foreach (FooterPart part in main.FooterParts) if (part.Footer != null) InspectWordRoot(part.Footer, "Footer[" + (++index).ToString(CultureInfo.InvariantCulture) + "]", false, styleResolver, builder, targets);
        if (main.FootnotesPart?.Footnotes != null) InspectWordRoot(main.FootnotesPart.Footnotes, "Footnotes", true, styleResolver, builder, targets);
        if (main.EndnotesPart?.Endnotes != null) InspectWordRoot(main.EndnotesPart.Endnotes, "Endnotes", true, styleResolver, builder, targets);
        if (main.WordprocessingCommentsPart?.Comments != null) InspectWordRoot(main.WordprocessingCommentsPart.Comments, "Comments", true, styleResolver, builder, targets);
        InspectWordAlternativeText(main, builder, targets);
        return builder.Build();
    }

    private static void InspectWordRoot(
        OpenXmlPartRootElement root,
        string rootLocation,
        bool nonPrimary,
        WordStyleResolver styleResolver,
        OfficeContentSafetyBuilder builder,
        IDictionary<string, WordCleanupTarget>? targets) {
        int runIndex = 0;
        foreach (W.Run run in root.Descendants<W.Run>()) {
            string text = run.InnerText;
            if (string.IsNullOrWhiteSpace(text)) continue;
            string location = rootLocation + "/Run[" + (++runIndex).ToString(CultureInfo.InvariantCulture) + "]";
            EffectiveWordRunStyle style = styleResolver.Resolve(run);
            OfficeContentConcealmentKind? kind = null;
            string? evidence = null;
            if (style.Hidden) {
                kind = OfficeContentConcealmentKind.HiddenByProperty;
                evidence = "The effective Word run formatting enables vanish, webHidden, or specVanish.";
            } else if (run.Ancestors<W.DeletedRun>().Any()) {
                kind = OfficeContentConcealmentKind.HiddenByProperty;
                evidence = "The run is retained as deleted revision content and is not ordinary current text.";
            } else if (style.FontSizePoints.HasValue && style.FontSizePoints.Value <= builder.Options.MaximumTinyFontSizePoints) {
                kind = OfficeContentConcealmentKind.TinyText;
                evidence = "The effective Word font size is " + style.FontSizePoints.Value.ToString("0.###", CultureInfo.InvariantCulture) + "pt.";
            } else if (HasZeroDrawingExtent(run)) {
                kind = OfficeContentConcealmentKind.ZeroDimension;
                evidence = "The owning Word drawing or VML shape has zero visible geometry.";
            } else if (TryGetWordContrast(style, out double contrast, out string colors) && contrast < builder.Options.MinimumVisibleContrastRatio) {
                kind = OfficeContentConcealmentKind.LowContrastText;
                evidence = colors + " has contrast ratio " + contrast.ToString("0.###", CultureInfo.InvariantCulture) + ".";
            } else if (nonPrimary && builder.Options.IncludeNonPrimaryContent) {
                kind = OfficeContentConcealmentKind.NonPrimaryContent;
                evidence = "The text is stored outside the primary body story in " + rootLocation + ".";
            }

            if (kind.HasValue) {
                OfficeContentSafetyFinding finding = builder.Add(
                    kind.Value,
                    OfficeContentSafetyRisk.ContextDependent,
                    location,
                    evidence!,
                    text,
                    OfficeContentCleanupCapability.RemoveElement,
                    inspectTextIntegrityEvidence: false);
                if (targets != null) targets[finding.Id] = WordCleanupTarget.ForElement(run);
            }
            int textIndex = 0;
            foreach (W.Text textNode in run.Descendants<W.Text>()) {
                string nodeText = textNode.Text ?? string.Empty;
                if (nodeText.Length == 0) continue;
                string nodeLocation = location + "/Text[" + (++textIndex).ToString(CultureInfo.InvariantCulture) + "]";
                IReadOnlyList<OfficeContentSafetyFinding> unicode = kind.HasValue
                    ? builder.InspectChargedTextIntegrity(nodeLocation, nodeText, OfficeContentCleanupCapability.RemoveText)
                    : builder.InspectVisibleText(nodeLocation, nodeText, OfficeContentCleanupCapability.RemoveText);
                if (targets != null) foreach (OfficeContentSafetyFinding item in unicode) targets[item.Id] = WordCleanupTarget.ForTextRange(textNode, item);
            }
        }
    }

    private static void InspectWordAlternativeText(
        MainDocumentPart main,
        OfficeContentSafetyBuilder builder,
        IDictionary<string, WordCleanupTarget>? targets) {
        if (!builder.Options.IncludeNonPrimaryContent) return;
        IEnumerable<OpenXmlPartRootElement> roots = new OpenXmlPartRootElement?[] { main.Document }
            .Concat(main.HeaderParts.Select(part => (OpenXmlPartRootElement?)part.Header))
            .Concat(main.FooterParts.Select(part => (OpenXmlPartRootElement?)part.Footer))
            .Where(root => root != null)
            .Cast<OpenXmlPartRootElement>();
        int index = 0;
        foreach (DW.DocProperties properties in roots.SelectMany(root => root.Descendants<DW.DocProperties>())) {
            index++;
            AddAlternativeText(properties, properties.Description?.Value, "description", "Description", index, builder, targets);
            AddAlternativeText(properties, properties.Title?.Value, "title", "Title", index, builder, targets);
        }
    }

    private static void AddAlternativeText(
        DW.DocProperties properties,
        string? value,
        string attributeName,
        string displayName,
        int index,
        OfficeContentSafetyBuilder builder,
        IDictionary<string, WordCleanupTarget>? targets) {
        if (string.IsNullOrWhiteSpace(value)) return;
        string location = "DrawingProperties[" + index.ToString(CultureInfo.InvariantCulture) + "]/@" + attributeName;
        OfficeContentSafetyFinding finding = builder.Add(
            OfficeContentConcealmentKind.NonPrimaryContent,
            OfficeContentSafetyRisk.Informational,
            location,
            "Drawing " + displayName + " alternative text is machine-readable but not ordinary body text.",
            value,
            OfficeContentCleanupCapability.RemoveText);
        if (targets != null) targets[finding.Id] = WordCleanupTarget.ForAttribute(properties, attributeName);
    }

    private static bool HasZeroDrawingExtent(W.Run run) {
        foreach (W.Drawing drawing in run.Descendants<W.Drawing>()) {
            DW.Extent? extent = drawing.Descendants<DW.Extent>().FirstOrDefault();
            if (extent != null && (extent.Cx?.Value <= 0 || extent.Cy?.Value <= 0)) return true;
        }
        foreach (OpenXmlElement ancestor in run.Ancestors()) {
            if (!string.Equals(ancestor.LocalName, "shape", StringComparison.OrdinalIgnoreCase)) continue;
            string style = ancestor.GetAttribute("style", string.Empty).Value ?? string.Empty;
            string compact = style.Replace(" ", string.Empty).ToLowerInvariant();
            if (compact.Contains("width:0") || compact.Contains("height:0")) return true;
        }
        return false;
    }

    private static bool TryGetWordContrast(EffectiveWordRunStyle style, out double ratio, out string evidence) {
        ratio = 0;
        evidence = string.Empty;
        if (!TryParseWordColor(style.Foreground, out OfficeColor foreground)) return false;
        OfficeColor background = TryParseWordColor(style.Background, out OfficeColor parsedBackground) ? parsedBackground : OfficeColor.White;
        ratio = OfficeColorContrast.ContrastRatio(foreground, background);
        evidence = "Effective Word foreground #" + foreground.ToRgbHex() + " against background #" + background.ToRgbHex();
        return true;
    }

    private static bool TryParseWordColor(string? value, out OfficeColor color) {
        string normalized = value?.Trim() ?? string.Empty;
        if (normalized.Length == 6 && OfficeColor.TryParseHex(normalized, out color)) return true;
        color = default;
        return false;
    }

    private static byte[] PrepareContentSafetyMutation(byte[] data, OfficeSignatureMutationPolicy signaturePolicy) {
        var provenanceOptions = new OfficeProvenanceRemovalOptions { SignatureMutationPolicy = signaturePolicy };
        bool hasSignatures = HasPackageSignatures(data, provenanceOptions);
        if (!hasSignatures) return (byte[])data.Clone();
        if (signaturePolicy == OfficeSignatureMutationPolicy.BlockSave) {
            throw new InvalidOperationException("Content cleanup would invalidate existing Word package signatures. Select RemoveInvalidatedSignatures or PreserveSignatureMarkup explicitly.");
        }
        return signaturePolicy == OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures
            ? StripPackageSignatures(data, provenanceOptions.Limits).Data
            : (byte[])data.Clone();
    }

    private sealed class WordStyleResolver {
        private readonly Dictionary<string, W.Style> _styles;
        private readonly W.RunPropertiesBaseStyle? _defaults;
        internal WordStyleResolver(MainDocumentPart main) {
            _styles = main.StyleDefinitionsPart?.Styles?.Elements<W.Style>()
                .Where(style => style.StyleId?.Value != null)
                .ToDictionary(style => style.StyleId!.Value!, StringComparer.Ordinal) ?? new Dictionary<string, W.Style>(StringComparer.Ordinal);
            _defaults = main.StyleDefinitionsPart?.Styles?.DocDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle;
        }
        internal EffectiveWordRunStyle Resolve(W.Run run) {
            var result = new EffectiveWordRunStyle();
            Apply(result, _defaults);
            W.Paragraph? paragraph = run.Ancestors<W.Paragraph>().FirstOrDefault();
            ApplyStyleChain(result, paragraph?.ParagraphProperties?.ParagraphStyleId?.Val?.Value);
            ApplyStyleChain(result, run.RunProperties?.RunStyle?.Val?.Value);
            Apply(result, paragraph?.ParagraphProperties?.ParagraphMarkRunProperties);
            Apply(result, run.RunProperties);
            W.Shading? paragraphShading = paragraph?.ParagraphProperties?.Shading;
            if (!string.IsNullOrWhiteSpace(paragraphShading?.Fill?.Value) && string.IsNullOrWhiteSpace(result.Background)) result.Background = paragraphShading!.Fill!.Value;
            W.TableCell? cell = run.Ancestors<W.TableCell>().FirstOrDefault();
            W.Shading? cellShading = cell?.TableCellProperties?.Shading;
            if (!string.IsNullOrWhiteSpace(cellShading?.Fill?.Value) && string.IsNullOrWhiteSpace(result.Background)) result.Background = cellShading!.Fill!.Value;
            return result;
        }
        private void ApplyStyleChain(EffectiveWordRunStyle target, string? styleId) {
            if (string.IsNullOrWhiteSpace(styleId)) return;
            var chain = new Stack<W.Style>();
            var visited = new HashSet<string>(StringComparer.Ordinal);
            string? current = styleId;
            while (!string.IsNullOrWhiteSpace(current) && visited.Add(current!) && _styles.TryGetValue(current!, out W.Style? style)) {
                chain.Push(style);
                current = style.BasedOn?.Val?.Value;
            }
            while (chain.Count > 0) Apply(target, chain.Pop().StyleRunProperties);
        }
        private static void Apply(EffectiveWordRunStyle target, OpenXmlElement? properties) {
            if (properties == null) return;
            ApplyHidden<W.Vanish>(properties, value => target.Vanish = value);
            ApplyHidden<W.WebHidden>(properties, value => target.WebHidden = value);
            ApplyHidden<W.SpecVanish>(properties, value => target.SpecVanish = value);
            W.FontSize? size = properties.GetFirstChild<W.FontSize>();
            if (size?.Val?.Value != null && double.TryParse(size.Val.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double halfPoints)) target.FontSizePoints = halfPoints / 2D;
            W.Color? color = properties.GetFirstChild<W.Color>();
            string? foreground = color?.Val?.Value;
            if (!string.IsNullOrWhiteSpace(foreground) && !string.Equals(foreground, "auto", StringComparison.OrdinalIgnoreCase)) target.Foreground = foreground;
            W.Shading? shading = properties.GetFirstChild<W.Shading>();
            string? background = shading?.Fill?.Value;
            if (!string.IsNullOrWhiteSpace(background) && !string.Equals(background, "auto", StringComparison.OrdinalIgnoreCase)) target.Background = background;
        }
        private static void ApplyHidden<T>(OpenXmlElement properties, Action<bool> apply) where T : W.OnOffType {
            T? value = properties.GetFirstChild<T>();
            if (value == null) return;
            apply(value.Val?.Value != false);
        }
    }

    private sealed class EffectiveWordRunStyle {
        internal bool Vanish { get; set; }
        internal bool WebHidden { get; set; }
        internal bool SpecVanish { get; set; }
        internal bool Hidden => Vanish || WebHidden || SpecVanish;
        internal double? FontSizePoints { get; set; }
        internal string? Foreground { get; set; }
        internal string? Background { get; set; }
    }

    private sealed class WordCleanupTarget : IEquatable<WordCleanupTarget> {
        private readonly OpenXmlElement _element;
        private readonly string? _attribute;
        private readonly int? _offset;
        private readonly int? _length;
        private readonly string? _expected;
        private WordCleanupTarget(OpenXmlElement element, string? attribute, int? offset = null, int? length = null, string? expected = null) { _element = element; _attribute = attribute; _offset = offset; _length = length; _expected = expected; }
        internal static WordCleanupTarget ForElement(OpenXmlElement element) => new WordCleanupTarget(element, null);
        internal static WordCleanupTarget ForAttribute(OpenXmlElement element, string attribute) => new WordCleanupTarget(element, attribute);
        internal static WordCleanupTarget ForTextRange(W.Text text, OfficeContentSafetyFinding finding) => new WordCleanupTarget(
            text, null, finding.SourceTextOffset, finding.SourceTextLength,
            text.Text.Substring(finding.SourceTextOffset!.Value, finding.SourceTextLength!.Value));
        internal void Remove() {
            if (_element is W.Text text && _offset.HasValue && _length.HasValue) {
                string current = text.Text ?? string.Empty;
                if (_offset.Value > current.Length - _length.Value || !string.Equals(current.Substring(_offset.Value, _length.Value), _expected, StringComparison.Ordinal)) {
                    throw new InvalidOperationException("The selected Unicode text range no longer matches the inspected Word text node.");
                }
                text.Text = current.Remove(_offset.Value, _length.Value);
                return;
            }
            if (_attribute == null) { _element.Remove(); return; }
            if (_element is DW.DocProperties properties) {
                if (string.Equals(_attribute, "description", StringComparison.Ordinal)) properties.Description = null;
                else if (string.Equals(_attribute, "title", StringComparison.Ordinal)) properties.Title = null;
            }
        }
        public bool Equals(WordCleanupTarget? other) => other != null && ReferenceEquals(_element, other._element) && string.Equals(_attribute, other._attribute, StringComparison.Ordinal) && _offset == other._offset && _length == other._length;
        public override bool Equals(object? obj) => Equals(obj as WordCleanupTarget);
        public override int GetHashCode() { unchecked { return (_element.GetHashCode() * 397) ^ (_attribute?.GetHashCode() ?? 0) ^ (_offset ?? 0); } }
    }
}
