using OfficeIMO.ContentSafety;
using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;

namespace OfficeIMO.OpenDocument;

public abstract partial class OdfDocument {
    /// <summary>Inspects ODT, ODS, or ODP native visibility, geometry, contrast, notes, comments, and Unicode evidence.</summary>
    public static OfficeContentSafetyReport InspectContentSafety(
        byte[] documentBytes,
        OfficeContentSafetyOptions? options = null,
        OdfLoadOptions? loadOptions = null) {
#if NET6_0_OR_GREATER
        ArgumentNullException.ThrowIfNull(documentBytes);
#else
        if (documentBytes == null) throw new ArgumentNullException(nameof(documentBytes));
#endif
        OfficeContentSafetyOptions effective = options ?? new OfficeContentSafetyOptions();
        OfficeContentSafetyInputGuard.ValidateBytes(documentBytes, effective, inspectZipPackage: true);
        using var stream = new MemoryStream(documentBytes, writable: false);
        OdfDocument document = Load(stream, loadOptions);
        return InspectOdfContentSafety(document, effective, targets: null);
    }

    /// <summary>Inspects an OpenDocument package without treating concealment as evidence of AI authorship.</summary>
    public static OfficeContentSafetyReport InspectContentSafety(
        string filePath,
        OfficeContentSafetyOptions? options = null,
        OdfLoadOptions? loadOptions = null) {
        if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("A file path is required.", nameof(filePath));
        OfficeContentSafetyOptions effective = options ?? new OfficeContentSafetyOptions();
        return InspectContentSafety(OfficeContentSafetyInputGuard.ReadAllBytes(filePath, effective, inspectZipPackage: true), effective, loadOptions);
    }

    /// <summary>Removes exact selected OpenDocument text nodes and verifies the preservation-aware rewritten package.</summary>
    public static OfficeContentCleanupResult RemoveSelectedContent(
        byte[] documentBytes,
        OfficeContentCleanupSelection selection,
        OfficeContentCleanupOptions? options = null,
        OdfLoadOptions? loadOptions = null) {
#if NET6_0_OR_GREATER
        ArgumentNullException.ThrowIfNull(documentBytes);
        ArgumentNullException.ThrowIfNull(selection);
#else
        if (documentBytes == null) throw new ArgumentNullException(nameof(documentBytes));
        if (selection == null) throw new ArgumentNullException(nameof(selection));
#endif
        options ??= new OfficeContentCleanupOptions();
        options.Validate();
        OfficeContentSafetyReport before = InspectContentSafety(documentBytes, options.Inspection, loadOptions);
        IReadOnlyList<OfficeContentSafetyFinding> selected = OfficeContentSafetyBuilder.ResolveSelection(before, selection);
        if (selected.Count == 0) return new OfficeContentCleanupResult((byte[])documentBytes.Clone(), before, before, Array.Empty<OfficeContentCleanupChange>());

        using var stream = new MemoryStream(documentBytes, writable: false);
        OdfDocument document = Load(stream, loadOptions);
        if (document.Security.SourceIsEncrypted) {
            throw new InvalidOperationException("OpenDocument content cleanup does not silently remove or replace package encryption. Save an explicitly decrypted copy first.");
        }
        var targets = new Dictionary<string, OdfContentSafetyTarget>(StringComparer.Ordinal);
        OfficeContentSafetyReport current = InspectOdfContentSafety(document, options.Inspection, targets);
        IReadOnlyList<OfficeContentSafetyFinding> currentSelection = OfficeContentSafetyBuilder.ResolveSelection(current, selection);
        foreach (OfficeContentSafetyFinding finding in currentSelection.OrderByDescending(item => item.SourceTextOffset ?? -1)) targets[finding.Id].Remove();
        document.MarkPartDirty("content.xml");

        OdfSignatureHandling signatureHandling = options.SignatureMutationPolicy == OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures
            ? OdfSignatureHandling.RemoveInvalidated
            : OdfSignatureHandling.RejectInvalidation;
        if (document.Package.IsSigned && options.SignatureMutationPolicy == OfficeSignatureMutationPolicy.PreserveSignatureMarkup) {
            throw new InvalidOperationException("OpenDocument cleanup cannot preserve signature markup as valid evidence after mutation. Select BlockSave or RemoveInvalidatedSignatures explicitly.");
        }
        byte[] output = document.Serialize(new OdfSaveOptions { SignatureHandling = signatureHandling }).RequireValue();
        OfficeContentSafetyReport after = InspectContentSafety(output, options.Inspection, loadOptions);
        OfficeContentCleanupChange[] changes = selected
            .Select(item => new OfficeContentCleanupChange(item.Id, item.Location, item.CleanupCapability))
            .ToArray();
        return new OfficeContentCleanupResult(output, before, after, changes);
    }

    /// <summary>Atomically writes an explicitly cleaned OpenDocument artifact.</summary>
    public static OfficeContentCleanupResult RemoveSelectedContent(
        string inputPath,
        string outputPath,
        OfficeContentCleanupSelection selection,
        OfficeContentCleanupOptions? options = null,
        OdfLoadOptions? loadOptions = null) {
        if (string.IsNullOrWhiteSpace(inputPath)) throw new ArgumentException("An input path is required.", nameof(inputPath));
        if (string.IsNullOrWhiteSpace(outputPath)) throw new ArgumentException("An output path is required.", nameof(outputPath));
        options ??= new OfficeContentCleanupOptions();
        options.Validate();
        OfficeContentCleanupResult result = RemoveSelectedContent(OfficeContentSafetyInputGuard.ReadAllBytes(inputPath, options.Inspection, inspectZipPackage: true), selection, options, loadOptions);
        OfficeFileCommit.WriteAllBytes(outputPath, result.Output);
        return result;
    }

    private static OfficeContentSafetyReport InspectOdfContentSafety(
        OdfDocument document,
        OfficeContentSafetyOptions? options,
        IDictionary<string, OdfContentSafetyTarget>? targets) {
        string format = document.Kind switch {
            OdfDocumentKind.Text => "OpenDocument Text",
            OdfDocumentKind.Spreadsheet => "OpenDocument Spreadsheet",
            OdfDocumentKind.Presentation => "OpenDocument Presentation",
            _ => "OpenDocument"
        };
        var builder = new OfficeContentSafetyBuilder(format, options);
        XDocument content = document.GetXml("content.xml");
        XElement body = content.Root?.Element(OdfNamespaces.Office + "body")
            ?? throw new InvalidDataException("OpenDocument content has no office:body.");
        int segmentIndex = 0;
        foreach (OdfTextSegment segment in EnumerateOdfTextSegments(body)) {
            string text = OdfTextCodec.ReadNodes(segment.Nodes);
            if (string.IsNullOrWhiteSpace(text)) continue;
            segmentIndex++;
            XElement owner = segment.Owner;
            string location = BuildOdfLocation(owner, segmentIndex);
            OdfContentSafetyState state = ResolveOdfState(document, owner);
            OfficeContentConcealmentKind? kind = null;
            string? evidence = null;
            if (state.HiddenEvidence != null) {
                kind = state.HiddenContainer ? OfficeContentConcealmentKind.HiddenContainer : OfficeContentConcealmentKind.HiddenByProperty;
                evidence = state.HiddenEvidence;
            } else if (state.ZeroGeometryEvidence != null) {
                kind = OfficeContentConcealmentKind.ZeroDimension;
                evidence = state.ZeroGeometryEvidence;
            } else if (state.Opacity.HasValue && state.Opacity.Value <= 0.01D) {
                kind = OfficeContentConcealmentKind.TransparentText;
                evidence = "The effective OpenDocument drawing opacity is zero or nearly zero.";
            } else if (state.FontSizePoints.HasValue && state.FontSizePoints.Value <= builder.Options.MaximumTinyFontSizePoints) {
                kind = OfficeContentConcealmentKind.TinyText;
                evidence = "The effective OpenDocument font size is " + state.FontSizePoints.Value.ToString("0.###", CultureInfo.InvariantCulture) + "pt.";
            } else if (TryGetOdfContrast(state, out double ratio, out string colors) && ratio < builder.Options.MinimumVisibleContrastRatio) {
                kind = OfficeContentConcealmentKind.LowContrastText;
                evidence = colors + " has contrast ratio " + ratio.ToString("0.###", CultureInfo.InvariantCulture) + ".";
            } else if (state.NonPrimary && builder.Options.IncludeNonPrimaryContent) {
                kind = OfficeContentConcealmentKind.NonPrimaryContent;
                evidence = "The text is stored in an OpenDocument note, annotation, alternative description, or presentation-notes story.";
            }

            if (kind.HasValue) {
                OfficeContentSafetyFinding finding = builder.Add(kind.Value, OfficeContentSafetyRisk.ContextDependent, location, evidence!, text, OfficeContentCleanupCapability.RemoveText, inspectTextIntegrityEvidence: false);
                if (targets != null) targets[finding.Id] = new OdfContentSafetyTarget(segment.Nodes);
            }
            int textNodeIndex = 0;
            foreach (XText textNode in segment.Nodes.OfType<XText>()) {
                if (textNode.Value.Length == 0) continue;
                string nodeLocation = location + "/TextNode[" + (++textNodeIndex).ToString(CultureInfo.InvariantCulture) + "]";
                IReadOnlyList<OfficeContentSafetyFinding> unicode = kind.HasValue
                    ? builder.InspectChargedTextIntegrity(nodeLocation, textNode.Value, OfficeContentCleanupCapability.RemoveText)
                    : builder.InspectVisibleText(nodeLocation, textNode.Value, OfficeContentCleanupCapability.RemoveText);
                if (targets != null) foreach (OfficeContentSafetyFinding item in unicode) targets[item.Id] = new OdfContentSafetyTarget(textNode, item);
            }
        }
        InspectOdfMachineReadableAttributes(document, body, builder, targets);
        builder.AddDiagnostic("OpenDocument conditional formatting and formula-driven visibility are not render-evaluated; exact native hidden properties and resolved style chains are inspected.");
        return builder.Build();
    }

    private static void InspectOdfMachineReadableAttributes(
        OdfDocument document,
        XElement body,
        OfficeContentSafetyBuilder builder,
        IDictionary<string, OdfContentSafetyTarget>? targets) {
        int index = 0;
        foreach (XElement element in body.DescendantsAndSelf()) {
            OdfContentSafetyState state = ResolveOdfState(document, element);
            bool nativeHiddenField = element.Name == OdfNamespaces.Text + "hidden-text" || element.Name == OdfNamespaces.Text + "hidden-paragraph";
            foreach (XAttribute attribute in element.Attributes()) {
                bool hiddenString = nativeHiddenField && attribute.Name == OdfNamespaces.Text + "string-value";
                bool storedValue = attribute.Name == OdfNamespaces.Office + "string-value" ||
                    attribute.Name == OdfNamespaces.Office + "value" ||
                    attribute.Name == OdfNamespaces.Office + "boolean-value" ||
                    attribute.Name == OdfNamespaces.Office + "date-value" ||
                    attribute.Name == OdfNamespaces.Office + "time-value";
                bool formula = attribute.Name == OdfNamespaces.Table + "formula";
                if (!hiddenString && !storedValue && !formula) continue;
                string value = attribute.Value;
                if (string.IsNullOrWhiteSpace(value)) continue;
                bool concealedOwner = state.HiddenEvidence != null || state.ZeroGeometryEvidence != null || state.NonPrimary;
                bool instructionLike = OfficeContentInstructionDetector.Detect(value).Count > 0;
                if (!hiddenString && !concealedOwner && !instructionLike) continue;

                OfficeContentConcealmentKind kind = hiddenString
                    ? OfficeContentConcealmentKind.HiddenByProperty
                    : concealedOwner
                        ? (state.HiddenContainer ? OfficeContentConcealmentKind.HiddenContainer : OfficeContentConcealmentKind.NonPrimaryContent)
                        : OfficeContentConcealmentKind.NonPrimaryContent;
                string location = BuildOdfAttributeLocation(element, attribute, ++index);
                OfficeContentSafetyFinding finding = builder.Add(
                    kind,
                    OfficeContentSafetyRisk.ContextDependent,
                    location,
                    hiddenString
                        ? "The canonical OpenDocument hidden-text payload is stored in text:string-value rather than a text node."
                        : "The OpenDocument stored value or formula is machine-readable outside ordinary displayed text.",
                    value,
                    OfficeContentCleanupCapability.RemoveText,
                    inspectTextIntegrityEvidence: false);
                if (targets != null) targets[finding.Id] = new OdfContentSafetyTarget(attribute);
                IReadOnlyList<OfficeContentSafetyFinding> unicode = builder.InspectChargedTextIntegrity(
                    location + "/Value",
                    value,
                    OfficeContentCleanupCapability.RemoveText);
                if (targets != null) foreach (OfficeContentSafetyFinding item in unicode) targets[item.Id] = new OdfContentSafetyTarget(attribute, item);
            }
        }
    }

    private static IEnumerable<OdfTextSegment> EnumerateOdfTextSegments(XElement root) {
        var pending = new List<XNode>();
        foreach (XNode node in root.Nodes()) {
            if (IsOdfTextPrimitive(node)) {
                pending.Add(node);
                continue;
            }
            if (pending.Count > 0) {
                yield return new OdfTextSegment(root, pending.ToArray());
                pending.Clear();
            }
            if (node is XElement child) {
                foreach (OdfTextSegment nested in EnumerateOdfTextSegments(child)) yield return nested;
            }
        }
        if (pending.Count > 0) yield return new OdfTextSegment(root, pending.ToArray());
    }

    private static bool IsOdfTextPrimitive(XNode node) =>
        node is XText || node is XElement element &&
        (element.Name == OdfNamespaces.Text + "s" || element.Name == OdfNamespaces.Text + "tab" || element.Name == OdfNamespaces.Text + "line-break");

    private static OdfContentSafetyState ResolveOdfState(OdfDocument document, XElement owner) {
        var state = new OdfContentSafetyState {
            CanUseDefaultWhiteBackground = document.Kind != OdfDocumentKind.Presentation &&
                !owner.AncestorsAndSelf().Any(element => element.Name.Namespace == OdfNamespaces.Draw)
        };
        XElement[] ancestry = owner.AncestorsAndSelf().Reverse().ToArray();
        foreach (XElement element in ancestry) {
            ApplyOdfElementState(document, element, state);
            foreach (OdfStyle style in ResolveOdfElementStyles(document, element)) {
                foreach (OdfStyle candidate in document.Styles.Resolve(style).Reverse()) ApplyOdfStyleState(candidate, state);
            }
        }
        if (document.Kind == OdfDocumentKind.Spreadsheet && TryGetOdsColumnElement(owner, out XElement? column)) {
            ApplyOdfElementState(document, column!, state);
            foreach (OdfStyle style in ResolveOdfElementStyles(document, column!)) {
                foreach (OdfStyle candidate in document.Styles.Resolve(style).Reverse()) ApplyOdfStyleState(candidate, state);
            }
        }
        return state;
    }

    private static IEnumerable<OdfStyle> ResolveOdfElementStyles(OdfDocument document, XElement element) {
        string? name = (string?)element.Attribute(OdfNamespaces.Text + "style-name");
        OdfStyleFamily family = element.Name == OdfNamespaces.Text + "p" || element.Name == OdfNamespaces.Text + "h"
            ? OdfStyleFamily.Paragraph
            : OdfStyleFamily.Text;
        if (!string.IsNullOrWhiteSpace(name)) {
            OdfStyle? style = document.Styles.FindInPart(family, name!, "content.xml");
            if (style != null) yield return style;
        }
        name = (string?)element.Attribute(OdfNamespaces.Draw + "style-name");
        if (!string.IsNullOrWhiteSpace(name)) {
            OdfStyle? style = document.Styles.FindInPart(OdfStyleFamily.Graphic, name!, "content.xml");
            if (style != null) yield return style;
        }
        name = (string?)element.Attribute(OdfNamespaces.Table + "style-name");
        if (!string.IsNullOrWhiteSpace(name)) {
            OdfStyleFamily tableFamily = element.Name.LocalName.IndexOf("row", StringComparison.Ordinal) >= 0
                ? OdfStyleFamily.TableRow
                : element.Name.LocalName.IndexOf("column", StringComparison.Ordinal) >= 0
                    ? OdfStyleFamily.TableColumn
                    : element.Name.LocalName.IndexOf("cell", StringComparison.Ordinal) >= 0
                        ? OdfStyleFamily.TableCell
                        : OdfStyleFamily.Table;
            OdfStyle? style = document.Styles.FindInPart(tableFamily, name!, "content.xml");
            if (style != null) yield return style;
        }
    }

    private static void ApplyOdfElementState(OdfDocument document, XElement element, OdfContentSafetyState state) {
        string? presentationVisibility = (string?)element.Attribute(OdfNamespaces.Presentation + "visibility");
        if (string.Equals(presentationVisibility, "hidden", StringComparison.OrdinalIgnoreCase)) {
            state.HiddenEvidence = "An owning OpenDocument presentation page or shape has presentation:visibility='hidden'.";
            state.HiddenContainer = true;
        }
        string? tableVisibility = (string?)element.Attribute(OdfNamespaces.Table + "visibility");
        if (string.Equals(tableVisibility, "collapse", StringComparison.OrdinalIgnoreCase) || string.Equals(tableVisibility, "filter", StringComparison.OrdinalIgnoreCase)) {
            state.HiddenEvidence = "An owning OpenDocument sheet, row, or column has table:visibility='" + tableVisibility + "'.";
            state.HiddenContainer = true;
        }
        if (string.Equals((string?)element.Attribute(OdfNamespaces.Table + "display"), "false", StringComparison.OrdinalIgnoreCase)) {
            state.HiddenEvidence = "An owning OpenDocument table group has table:display='false'.";
            state.HiddenContainer = true;
        }
        if ((element.Name == OdfNamespaces.Text + "hidden-text" || element.Name == OdfNamespaces.Text + "hidden-paragraph") &&
            OdfBoolean.ReadCompatible((string?)element.Attribute(OdfNamespaces.Text + "is-hidden"), true)) {
            state.HiddenEvidence = "The native OpenDocument hidden-text or hidden-paragraph field evaluates as hidden.";
        }
        if (string.Equals((string?)element.Attribute(OdfNamespaces.Text + "display"), "none", StringComparison.OrdinalIgnoreCase)) {
            state.HiddenEvidence = "The OpenDocument element has text:display='none'.";
        }
        if (element.Name == OdfNamespaces.Office + "annotation" || element.Name == OdfNamespaces.Presentation + "notes" ||
            element.Name == OdfNamespaces.Text + "note" || element.Name == OdfNamespaces.Svg + "title" || element.Name == OdfNamespaces.Svg + "desc") {
            state.NonPrimary = true;
        }
        if (element.Name.Namespace == OdfNamespaces.Draw) {
            string? width = (string?)element.Attribute(OdfNamespaces.Svg + "width");
            string? height = (string?)element.Attribute(OdfNamespaces.Svg + "height");
            if (IsZeroOdfLength(width) || IsZeroOdfLength(height)) state.ZeroGeometryEvidence = "An owning OpenDocument drawing has zero width or height.";
        }
    }

    private static void ApplyOdfStyleState(OdfStyle style, OdfContentSafetyState state) {
        XElement? text = style.Element.Element(OdfNamespaces.Style + "text-properties");
        if (text != null) {
            if (string.Equals((string?)text.Attribute(OdfNamespaces.Text + "display"), "none", StringComparison.OrdinalIgnoreCase)) {
                state.HiddenEvidence = "The resolved OpenDocument text style has text:display='none'.";
            }
            string? fontSize = (string?)text.Attribute(OdfNamespaces.Fo + "font-size");
            if (!string.IsNullOrWhiteSpace(fontSize)) {
                OdfLength size = OdfLength.Parse(fontSize!);
                state.FontSizePoints = size.TryToPoints(out double points) ? points : null;
            }
            string? foreground = (string?)text.Attribute(OdfNamespaces.Fo + "color");
            if (!string.IsNullOrWhiteSpace(foreground)) state.Foreground = foreground;
            string? background = (string?)text.Attribute(OdfNamespaces.Fo + "background-color");
            if (!string.IsNullOrWhiteSpace(background) && !string.Equals(background, "transparent", StringComparison.OrdinalIgnoreCase)) state.Background = background;
        }
        XElement? graphic = style.Element.Element(OdfNamespaces.Style + "graphic-properties");
        if (graphic != null) {
            string? opacity = (string?)graphic.Attribute(OdfNamespaces.Draw + "opacity");
            if (TryParseOdfPercent(opacity, out double parsedOpacity)) state.Opacity = parsedOpacity;
            if (string.Equals((string?)graphic.Attribute(OdfNamespaces.Draw + "fill"), "solid", StringComparison.OrdinalIgnoreCase)) {
                string? fill = (string?)graphic.Attribute(OdfNamespaces.Draw + "fill-color");
                if (!string.IsNullOrWhiteSpace(fill)) state.Background = fill;
            }
        }
        foreach ((string propertiesName, string lengthName) in new[] {
            ("table-row-properties", "row-height"),
            ("table-column-properties", "column-width")
        }) {
            string? length = (string?)style.Element.Element(OdfNamespaces.Style + propertiesName)?.Attribute(OdfNamespaces.Style + lengthName);
            if (IsZeroOdfLength(length)) state.ZeroGeometryEvidence = "The resolved OpenDocument row or column style has zero visible geometry.";
        }
        string? cellBackground = (string?)style.Element.Element(OdfNamespaces.Style + "table-cell-properties")?.Attribute(OdfNamespaces.Fo + "background-color");
        if (!string.IsNullOrWhiteSpace(cellBackground) && !string.Equals(cellBackground, "transparent", StringComparison.OrdinalIgnoreCase)) state.Background = cellBackground;
        string? paragraphBackground = (string?)style.Element.Element(OdfNamespaces.Style + "paragraph-properties")?.Attribute(OdfNamespaces.Fo + "background-color");
        if (!string.IsNullOrWhiteSpace(paragraphBackground) && !string.Equals(paragraphBackground, "transparent", StringComparison.OrdinalIgnoreCase)) state.Background = paragraphBackground;
    }

    private static bool TryGetOdsColumnElement(XElement owner, out XElement? column) {
        column = null;
        XElement? cell = owner.AncestorsAndSelf().FirstOrDefault(item => item.Name == OdfNamespaces.Table + "table-cell" || item.Name == OdfNamespaces.Table + "covered-table-cell");
        XElement? table = cell?.Ancestors(OdfNamespaces.Table + "table").FirstOrDefault();
        if (cell == null || table == null) return false;
        long columnIndex = 0;
        foreach (XElement sibling in cell.ElementsBeforeSelf().Where(item => item.Name == OdfNamespaces.Table + "table-cell" || item.Name == OdfNamespaces.Table + "covered-table-cell")) {
            columnIndex = checked(columnIndex + ReadOdfRepeat(sibling, OdfNamespaces.Table + "number-columns-repeated"));
        }
        long cursor = 0;
        foreach (XElement candidate in table.Elements(OdfNamespaces.Table + "table-column")) {
            long repeat = ReadOdfRepeat(candidate, OdfNamespaces.Table + "number-columns-repeated");
            if (columnIndex >= cursor && columnIndex < checked(cursor + repeat)) { column = candidate; return true; }
            cursor = checked(cursor + repeat);
        }
        return false;
    }

    private static long ReadOdfRepeat(XElement element, XName name) {
        string? lexical = (string?)element.Attribute(name);
        return long.TryParse(lexical, NumberStyles.None, CultureInfo.InvariantCulture, out long value) && value > 0 ? value : 1L;
    }

    private static bool TryGetOdfContrast(OdfContentSafetyState state, out double ratio, out string evidence) {
        ratio = 0D;
        evidence = string.Empty;
        if (!OdfColor.TryParse(state.Foreground, out OdfColor foreground)) return false;
        OdfColor background;
        if (OdfColor.TryParse(state.Background, out OdfColor parsedBackground)) {
            background = parsedBackground;
        } else if (state.CanUseDefaultWhiteBackground) {
            background = new OdfColor(255, 255, 255);
        } else {
            return false;
        }
        OfficeColor foregroundColor = OfficeColor.ParseHex(foreground.ToString());
        OfficeColor backgroundColor = OfficeColor.ParseHex(background.ToString());
        ratio = OfficeColorContrast.ContrastRatio(foregroundColor, backgroundColor);
        evidence = "Effective OpenDocument foreground " + foreground + " against background " + background;
        return true;
    }

    private static bool IsZeroOdfLength(string? value) {
        if (string.IsNullOrWhiteSpace(value)) return false;
        OdfLength length = OdfLength.Parse(value!);
        return length.TryToPoints(out double points) && Math.Abs(points) <= 0.01D;
    }

    private static bool TryParseOdfPercent(string? value, out double ratio) {
        ratio = 0D;
        string lexical = value?.Trim() ?? string.Empty;
        if (!lexical.EndsWith("%", StringComparison.Ordinal) || !double.TryParse(lexical.Substring(0, lexical.Length - 1), NumberStyles.Float, CultureInfo.InvariantCulture, out double percent)) return false;
        ratio = Math.Max(0D, Math.Min(1D, percent / 100D));
        return true;
    }

    private static string BuildOdfLocation(XElement owner, int textIndex) {
        string label = owner.Name.LocalName;
        string? name = (string?)owner.Attribute(OdfNamespaces.Table + "name") ?? (string?)owner.Attribute(OdfNamespaces.Draw + "name");
        if (!string.IsNullOrWhiteSpace(name)) label += "('" + name!.Replace("'", "''") + "')";
        return "content.xml/" + label + "/Text[" + textIndex.ToString(CultureInfo.InvariantCulture) + "]";
    }

    private static string BuildOdfAttributeLocation(XElement owner, XAttribute attribute, int index) =>
        "content.xml/" + owner.Name.LocalName + "[" + index.ToString(CultureInfo.InvariantCulture) + "]/@" + attribute.Name.LocalName;

    private sealed class OdfContentSafetyState {
        internal string? HiddenEvidence { get; set; }
        internal bool HiddenContainer { get; set; }
        internal string? ZeroGeometryEvidence { get; set; }
        internal double? Opacity { get; set; }
        internal double? FontSizePoints { get; set; }
        internal string? Foreground { get; set; }
        internal string? Background { get; set; }
        internal bool CanUseDefaultWhiteBackground { get; set; }
        internal bool NonPrimary { get; set; }
    }

    private sealed class OdfTextSegment {
        internal OdfTextSegment(XElement owner, IReadOnlyList<XNode> nodes) { Owner = owner; Nodes = nodes; }
        internal XElement Owner { get; }
        internal IReadOnlyList<XNode> Nodes { get; }
    }

    private sealed class OdfContentSafetyTarget {
        private readonly IReadOnlyList<XNode>? _nodes;
        private readonly XAttribute? _attribute;
        private readonly XText? _text;
        private readonly int? _offset;
        private readonly int? _length;
        private readonly string? _expected;
        internal OdfContentSafetyTarget(IReadOnlyList<XNode> nodes) { _nodes = nodes; }
        internal OdfContentSafetyTarget(XAttribute attribute) { _attribute = attribute; }
        internal OdfContentSafetyTarget(XAttribute attribute, OfficeContentSafetyFinding finding) {
            _attribute = attribute;
            _offset = finding.SourceTextOffset;
            _length = finding.SourceTextLength;
            _expected = attribute.Value.Substring(_offset!.Value, _length!.Value);
        }
        internal OdfContentSafetyTarget(XText text, OfficeContentSafetyFinding finding) {
            _text = text;
            _offset = finding.SourceTextOffset;
            _length = finding.SourceTextLength;
            _expected = text.Value.Substring(_offset!.Value, _length!.Value);
        }
        internal void Remove() {
            if (_text != null && _offset.HasValue && _length.HasValue) {
                string current = _text.Value;
                if (_offset.Value > current.Length - _length.Value || !string.Equals(current.Substring(_offset.Value, _length.Value), _expected, StringComparison.Ordinal)) {
                    throw new InvalidOperationException("The selected Unicode text range no longer matches the inspected OpenDocument text node.");
                }
                _text.Value = current.Remove(_offset.Value, _length.Value);
                return;
            }
            if (_attribute != null && _offset.HasValue && _length.HasValue) {
                string current = _attribute.Value;
                if (_offset.Value > current.Length - _length.Value || !string.Equals(current.Substring(_offset.Value, _length.Value), _expected, StringComparison.Ordinal)) {
                    throw new InvalidOperationException("The selected Unicode text range no longer matches the inspected OpenDocument attribute.");
                }
                _attribute.Value = current.Remove(_offset.Value, _length.Value);
                return;
            }
            if (_attribute != null) { _attribute.Remove(); return; }
            foreach (XNode node in _nodes ?? Array.Empty<XNode>()) node.Remove();
        }
    }
}
