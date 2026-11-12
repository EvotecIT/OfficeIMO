using System.Globalization;

namespace OfficeIMO.Pdf;

internal sealed partial class PdfAConverter {
    private const int MaximumGraphicsStateNesting = 28;
    private const int MaximumFormDepth = 64;
    private const int MaximumStringBytes = 32767;
    private const int MaximumNameBytes = 127;

    private static readonly HashSet<string> ContentOperators = new(StringComparer.Ordinal) {
        "b", "B", "b*", "B*", "BDC", "BI", "BMC", "BT", "BX", "c", "cm", "CS", "cs", "d", "d0", "d1", "Do", "DP",
        "EI", "EMC", "ET", "EX", "f", "F", "f*", "G", "g", "gs", "h", "i", "ID", "j", "J", "K", "k", "l", "m", "M",
        "MP", "n", "q", "Q", "re", "RG", "rg", "ri", "s", "S", "SC", "sc", "SCN", "scn", "sh", "T*", "Tc", "Td",
        "TD", "Tf", "Tj", "TJ", "TL", "Tm", "Tr", "Ts", "Tw", "Tz", "v", "w", "W", "W*", "y", "'", "\""
    };

    private static readonly HashSet<string> RenderingIntents = new(StringComparer.Ordinal) {
        "RelativeColorimetric", "AbsoluteColorimetric", "Perceptual", "Saturation"
    };

    private static readonly HashSet<string> BlendModes = new(StringComparer.Ordinal) {
        "Normal", "Compatible", "Multiply", "Screen", "Overlay", "Darken", "Lighten", "ColorDodge", "ColorBurn",
        "HardLight", "SoftLight", "Difference", "Exclusion", "Hue", "Saturation", "Color", "Luminosity"
    };

    /// <summary>One font as the pages use it: which codes are drawn and whether any are visible.</summary>
    private sealed class FontUse {
        internal FontUse(PdfDictionary font, string resourceName, string location) {
            Font = font;
            ResourceName = resourceName;
            Location = location;
        }

        internal PdfDictionary Font { get; }
        internal string ResourceName { get; }
        internal string Location { get; }
        internal SortedSet<int> Codes { get; } = new();
        internal SortedSet<int> VisibleCodes { get; } = new();
        internal List<byte[]> CompositeStrings { get; } = new();
        /// <summary>A text-showing operator drew with this font in a visible mode, even if its string was empty.</summary>
        internal bool Visible { get; set; }
    }

    /// <summary>A place that paints with a device colour space, with the resources whose Default spaces apply.</summary>
    private sealed class DeviceColorUse {
        internal DeviceColorUse(string family, string location, PdfObject? owner, PdfDictionary? resources, int pageIndex) {
            Family = family;
            Location = location;
            Owner = owner;
            Resources = resources;
            PageIndex = pageIndex;
        }

        /// <summary>The page whose content painted first with this colour; PDF/A-4 lets a page carry its own output intent.</summary>
        internal int PageIndex { get; }

        internal string Family { get; }
        internal string Location { get; }
        internal PdfObject? Owner { get; }
        internal PdfDictionary? Resources { get; }
    }

    private readonly Dictionary<PdfObject, FontUse> _fontUses = new(ReferenceComparer.Instance);
    private readonly List<DeviceColorUse> _deviceColors = new();
    private readonly HashSet<(PdfObject Content, PdfObject? Resources)> _visitedContent = new(new ContentKeyComparer());
    private readonly HashSet<PdfObject> _checkedResources = new(ReferenceComparer.Instance);
    private readonly HashSet<PdfObject> _checkedColorSpaces = new(ReferenceComparer.Instance);
    private readonly HashSet<PdfObject> _checkedType3 = new(ReferenceComparer.Instance);

    private int _currentPageIndex = -1;

    // True for PDF/A-4 files where a page carries its own output intent, so device colour is checked per page.
    private bool _pageOutputIntents;
    private readonly List<DeviceColorCapture> _deviceColorCaptures = new();
    private readonly Dictionary<(PdfObject Content, PdfObject? Resources), DeviceColorCapture> _contentDeviceColors = new(new ContentKeyComparer());

    private void AnalyzeContent() {
        IReadOnlyList<PdfReadPage> pages = _document.Pages;
        _pageOutputIntents = _part == 4 && pages.Any(page => page.PageDictionary.Items.ContainsKey("OutputIntents"));
        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++) {
            _cancellationToken.ThrowIfCancellationRequested();
            _currentPageIndex = pageIndex;
            PdfReadPage page = pages[pageIndex];
            PdfDictionary pageDictionary = page.PageDictionary;
            string location = "page " + (pageIndex + 1).ToString(CultureInfo.InvariantCulture);
            PdfDictionary? resources = page.GetFontInspectionResources();
            CheckTransparencyGroup(pageDictionary, location, pageDictionary, resources);
            var streams = new List<PdfStream>();
            PdfObject? contents = pageDictionary.Items.TryGetValue("Contents", out PdfObject? value) ? Resolve(value) : null;
            if (contents is PdfStream single) {
                streams.Add(single);
            } else if (contents is PdfArray array) {
                foreach (PdfObject item in array.Items) {
                    if (Resolve(item) is PdfStream part) streams.Add(part);
                }
            }
            if (streams.Count > 0) WalkContent(streams, resources, location, pageDictionary, depth: 0);

            if (Resolve(pageDictionary.Items.TryGetValue("Annots", out PdfObject? annots) ? annots : null) is PdfArray annotations) {
                for (int index = 0; index < annotations.Items.Count; index++) {
                    if (ResolveDictionary(annotations.Items[index]) is not PdfDictionary annotation) continue;
                    string annotationLocation = location + ", annotation " + (index + 1).ToString(CultureInfo.InvariantCulture);
                    if (ResolveDictionary(annotation.Items.TryGetValue("AP", out PdfObject? ap) ? ap : null) is not PdfDictionary appearances) continue;
                    foreach (KeyValuePair<string, PdfObject> state in appearances.Items) {
                        PdfObject? resolved = Resolve(state.Value);
                        if (resolved is PdfStream stream) {
                            WalkForm(stream, resources, annotationLocation + " appearance", depth: 0);
                        } else if (resolved is PdfDictionary states) {
                            foreach (PdfObject stateValue in states.Items.Values) {
                                if (Resolve(stateValue) is PdfStream stateStream) WalkForm(stateStream, resources, annotationLocation + " appearance", depth: 0);
                            }
                        }
                    }
                }
            }
        }
    }

    private void WalkForm(PdfStream form, PdfDictionary? inheritedResources, string location, int depth) {
        if (depth > MaximumFormDepth) return;
        PdfDictionary? resources = ResolveDictionary(form.Dictionary.Items.TryGetValue("Resources", out PdfObject? value) ? value : null) ?? inheritedResources;
        CheckTransparencyGroup(form.Dictionary, location, form, resources);
        WalkContent(new[] { form }, resources, location, form, depth);
    }

    private void WalkContent(IReadOnlyList<PdfStream> streams, PdfDictionary? resources, string location, PdfObject owner, int depth) {
        if (depth > MaximumFormDepth) return;
        PdfObject key = streams.Count == 1 ? streams[0] : owner;
        if (!_visitedContent.Add((key, resources))) {
            // Content shared by pages is checked once, but under PDF/A-4 page output intents each page that paints it needs its
            // own device colour record, including the colour of forms nested in it.
            if (_pageOutputIntents && _contentDeviceColors.TryGetValue((key, resources), out DeviceColorCapture? painted)) {
                string page = "page " + (_currentPageIndex + 1).ToString(CultureInfo.InvariantCulture);
                foreach (DeviceColorUse use in painted.Uses) RecordDeviceColor(use.Family, page + ", shared with " + use.Location, use.Owner, use.Resources);
            }
            return;
        }
        var captured = new DeviceColorCapture();
        _deviceColorCaptures.Add(captured);
        try {
            WalkContentOnce(streams, resources, location, owner, depth);
        } finally {
            _deviceColorCaptures.RemoveAt(_deviceColorCaptures.Count - 1);
        }
        if (_pageOutputIntents) _contentDeviceColors[(key, resources)] = captured;
    }

    private void WalkContentOnce(IReadOnlyList<PdfStream> streams, PdfDictionary? resources, string location, PdfObject owner, int depth) {
        _cancellationToken.ThrowIfCancellationRequested();
        if (resources is not null) CheckResources(resources, location, depth);

        PdfReadLimits limits = _document.ReadOptions.Limits;
        if (!PdfContentStreamSequenceDecoder.TryDecode(streams, _objects, limits, enforcePageContentLimit: false, out string content)) {
            AddViolation(PdfAConversionRefusalReason.ContentNotCompliant, location, "The content stream cannot be decoded, so its fonts and colours cannot be checked.", owner);
            return;
        }

        var stack = new Stack<(PdfDictionary? Font, string? FontName, int RenderMode)>();
        PdfDictionary? font = null;
        string? fontName = null;
        int renderMode = 0;
        int nesting = 0;
        bool reportedNesting = false;
        bool reportedLimit = false;
        PdfAConverter self = this;
        try {
            PdfContentStreamInterpreter.Interpret(
                content,
                limits.MaxContentOperations,
                operation => {
                    IReadOnlyList<object> operands = operation.Operands;
                    if (_part != 4 && !reportedLimit && !CheckOperandLimits(operands)) {
                        reportedLimit = true;
                        AddViolation(PdfAConversionRefusalReason.ImplementationLimitExceeded, location,
                            "The content stream has a number, string or name beyond the PDF/A implementation limits.", owner);
                    }
                    switch (operation.Name) {
                        case "q":
                            stack.Push((font, fontName, renderMode));
                            nesting++;
                            if (_part != 4 && nesting > MaximumGraphicsStateNesting && !reportedNesting) {
                                reportedNesting = true;
                                AddViolation(PdfAConversionRefusalReason.ImplementationLimitExceeded, location,
                                    "Graphics states are nested deeper than the PDF/A limit of 28.", owner);
                            }
                            break;
                        case "Q":
                            if (stack.Count > 0) (font, fontName, renderMode) = stack.Pop();
                            if (nesting > 0) nesting--;
                            break;
                        case "Tf":
                            if (operands.Count >= 1 && operands[0] is string selected) {
                                fontName = selected;
                                font = ResolveResource(resources, "Font", selected);
                                if (font is null) ReportUndefinedResource(location, owner, "font", selected);
                                if (font is not null && GetName(font, "Subtype") == "Type3") WalkType3(font, resources, location, depth);
                            }
                            break;
                        case "Tr":
                            if (operands.Count >= 1 && operands[0] is double mode) renderMode = (int)mode;
                            break;
                        case "Tj":
                        case "'":
                            if (operands.Count >= 1 && operands[operands.Count - 1] is byte[] shown) RecordText(font, fontName, shown, renderMode, location);
                            break;
                        case "\"":
                            if (operands.Count >= 3 && operands[2] is byte[] spaced) RecordText(font, fontName, spaced, renderMode, location);
                            break;
                        case "TJ":
                            if (operands.Count >= 1 && operands[0] is List<object> items) {
                                foreach (object item in items) {
                                    if (item is byte[] part) RecordText(font, fontName, part, renderMode, location);
                                }
                            }
                            break;
                        case "rg":
                        case "RG":
                            RecordDeviceColor("DeviceRGB", location, owner, resources);
                            break;
                        case "k":
                        case "K":
                            RecordDeviceColor("DeviceCMYK", location, owner, resources);
                            break;
                        case "g":
                        case "G":
                            RecordDeviceColor("DeviceGray", location, owner, resources);
                            break;
                        case "cs":
                        case "CS":
                            if (operands.Count >= 1 && operands[0] is string space) {
                                CheckColorSpaceName(space, resources, location, owner);
                                if (!IsBuiltInColorSpaceName(space) && ResolveResourceObject(resources, "ColorSpace", space) is null) {
                                    ReportUndefinedResource(location, owner, "colour space", space);
                                }
                            }
                            break;
                        case "sh":
                            if (operands.Count >= 1 && operands[0] is string shadingName) {
                                PdfObject? shading = ResolveResourceObject(resources, "Shading", shadingName);
                                if (shading is not null) CheckShading(shading, resources, location);
                                else ReportUndefinedResource(location, owner, "shading", shadingName);
                            }
                            break;
                        case "gs":
                            if (operands.Count >= 1 && operands[0] is string stateName && ResolveResourceObject(resources, "ExtGState", stateName) is null) {
                                ReportUndefinedResource(location, owner, "graphics state", stateName);
                            }
                            break;
                        case "Do":
                            if (operands.Count >= 1 && operands[0] is string xobjectName) {
                                if (ResolveResourceObject(resources, "XObject", xobjectName) is PdfStream xobject) {
                                    self.CheckXObject(xobject, resources, location + ", XObject " + xobjectName, depth);
                                } else {
                                    ReportUndefinedResource(location, owner, "XObject", xobjectName);
                                }
                            }
                            break;
                        case "BDC":
                        case "DP":
                            if (operands.Count >= 2 && operands[1] is string propertiesName && ResolveResourceObject(resources, "Properties", propertiesName) is null) {
                                ReportUndefinedResource(location, owner, "property list", propertiesName);
                            } else if (_part == 4 && operands.Count >= 2 && operands[1] is PdfContentDictionary inline &&
                                inline.Items.TryGetValue("ActualText", out object? replacement) && replacement is byte[] replacementBytes &&
                                ContainsPrivateUse(DecodeTextString(replacementBytes))) {
                                AddViolation(PdfAConversionRefusalReason.ContentNotCompliant, location,
                                    "Marked content carries replacement text (ActualText) made of private-use characters, which PDF/A-4 forbids; removing it would require rewriting the page content.", owner);
                            }
                            break;
                        case "BI":
                            if (operation.InlineImage is not null) CheckInlineImage(operation.InlineImage.Dictionary, resources, location, owner);
                            break;
                        case "ri":
                            if (_part != 4 && operands.Count >= 1 && operands[0] is string intent && !RenderingIntents.Contains(intent)) {
                                AddViolation(PdfAConversionRefusalReason.ContentNotCompliant, location,
                                    "The rendering intent /" + intent + " is not one PDF/A allows.", owner);
                            }
                            break;
                        default:
                            if (!ContentOperators.Contains(operation.Name) && operation.Name.Length > 0) {
                                AddViolation(PdfAConversionRefusalReason.ContentNotCompliant, location,
                                    "The content stream uses the operator '" + operation.Name + "', which PDF/A forbids even inside BX/EX.", owner);
                            }
                            break;
                    }
                },
                inlineImageComponentCount: name => GetComponentCount(ResolveColorSpaceName(name, resources)),
                inlineImageArrayComponentCount: array => GetComponentCount(array));
        } catch (PdfReadLimitException exception) {
            AddViolation(PdfAConversionRefusalReason.ContentNotCompliant, location, "The content stream exceeds a reading limit: " + exception.Message, owner);
        }
    }

    private static bool IsBuiltInColorSpaceName(string name) =>
        name is "DeviceGray" or "DeviceRGB" or "DeviceCMYK" or "Pattern" or "G" or "RGB" or "CMYK" or "I" or "Indexed";

    /// <summary>PDF/A-4 (6.2.2) requires every resource a content stream names to be defined in its resources.</summary>
    private void ReportUndefinedResource(string location, PdfObject owner, string kind, string name) {
        if (_part != 4) return;
        AddViolation(PdfAConversionRefusalReason.ContentNotCompliant, location,
            "The content stream uses the " + kind + " /" + name + ", which its resources do not define; PDF/A-4 requires every named resource to be defined.", owner);
    }

    private static string DecodeTextString(byte[] bytes) {
        if (bytes.Length >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF) return Encoding.BigEndianUnicode.GetString(bytes, 2, bytes.Length - 2);
        if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF) return Encoding.UTF8.GetString(bytes, 3, bytes.Length - 3);
        return string.Empty;
    }

    private void WalkType3(PdfDictionary font, PdfDictionary? resources, string location, int depth) {
        if (!_checkedType3.Add(font)) return;
        PdfDictionary? fontResources = ResolveDictionary(font.Items.TryGetValue("Resources", out PdfObject? value) ? value : null) ?? resources;
        if (ResolveDictionary(font.Items.TryGetValue("CharProcs", out PdfObject? procs) ? procs : null) is not PdfDictionary charProcs) return;
        foreach (KeyValuePair<string, PdfObject> glyph in charProcs.Items) {
            if (Resolve(glyph.Value) is PdfStream glyphStream) {
                WalkContent(new[] { glyphStream }, fontResources, location + ", Type 3 glyph " + glyph.Key, glyphStream, depth + 1);
            }
        }
    }

    private void RecordText(PdfDictionary? font, string? fontName, byte[] text, int renderMode, string location) {
        if (font is null) return;
        if (!_fontUses.TryGetValue(font, out FontUse? use)) {
            use = new FontUse(font, fontName ?? string.Empty, location);
            _fontUses.Add(font, use);
        }
        // Only mode 3 draws nothing; clipping modes still use the glyph outlines.
        bool visible = renderMode != 3;
        use.Visible |= visible;
        if (GetName(font, "Subtype") == "Type0") {
            string? encoding = GetName(font, "Encoding");
            if (encoding is "Identity-H" or "Identity-V") {
                for (int index = 0; index + 1 < text.Length; index += 2) {
                    int code = (text[index] << 8) | text[index + 1];
                    use.Codes.Add(code);
                    if (visible) use.VisibleCodes.Add(code);
                }
            } else if (use.CompositeStrings.Count < 4096) {
                use.CompositeStrings.Add(text);
                if (visible) use.VisibleCodes.Add(-1);
            }
            return;
        }
        foreach (byte code in text) {
            use.Codes.Add(code);
            if (visible) use.VisibleCodes.Add(code);
        }
    }

    private void RecordDeviceColor(string family, string location, PdfObject? owner, PdfDictionary? resources) {
        var use = new DeviceColorUse(family, location, owner, resources, _currentPageIndex);
        foreach (DeviceColorCapture capture in _deviceColorCaptures) capture.Add(use);
        foreach (DeviceColorUse existing in _deviceColors) {
            if (existing.Family == family && ReferenceEquals(existing.Owner, owner) && ReferenceEquals(existing.Resources, resources) &&
                (!_pageOutputIntents || existing.PageIndex == _currentPageIndex)) return;
        }
        _deviceColors.Add(use);
    }

    private PdfDictionary? ResolveResource(PdfDictionary? resources, string category, string name) =>
        ResolveDictionary(ResolveResourceObject(resources, category, name));

    private PdfObject? ResolveResourceObject(PdfDictionary? resources, string category, string name) {
        if (resources is null) return null;
        if (ResolveDictionary(resources.Items.TryGetValue(category, out PdfObject? group) ? group : null) is not PdfDictionary entries) return null;
        return entries.Items.TryGetValue(name, out PdfObject? value) ? Resolve(value) : null;
    }

    private PdfObject? ResolveColorSpaceName(string name, PdfDictionary? resources) {
        string normalized = name switch {
            "G" => "DeviceGray",
            "RGB" => "DeviceRGB",
            "CMYK" => "DeviceCMYK",
            "I" => "Indexed",
            _ => name
        };
        if (normalized is "DeviceGray" or "DeviceRGB" or "DeviceCMYK" or "Pattern") return new PdfName(normalized);
        return ResolveResourceObject(resources, "ColorSpace", name) ?? new PdfName(normalized);
    }

    private int GetComponentCount(PdfObject? space) {
        space = Resolve(space);
        if (space is PdfName name) {
            return name.Name switch {
                "DeviceGray" or "G" or "CalGray" => 1,
                "DeviceCMYK" or "CMYK" => 4,
                _ => 3
            };
        }
        if (space is PdfArray array && array.Items.Count > 0 && Resolve(array.Items[0]) is PdfName family) {
            switch (family.Name) {
                case "Indexed":
                case "I":
                case "Separation":
                    return 1;
                case "ICCBased":
                    return array.Items.Count > 1 && ResolveDictionary(array.Items[1]) is PdfDictionary icc &&
                        Resolve(icc.Items.TryGetValue("N", out PdfObject? n) ? n : null) is PdfNumber count ? (int)count.Value : 3;
                case "DeviceN":
                    return array.Items.Count > 1 && Resolve(array.Items[1]) is PdfArray names ? names.Items.Count : 1;
                case "CalGray":
                    return 1;
                case "Lab":
                case "CalRGB":
                    return 3;
            }
        }
        return 3;
    }

    private static bool CheckOperandLimits(IReadOnlyList<object> operands) {
        foreach (object operand in operands) {
            if (!CheckOperandLimit(operand)) return false;
        }
        return true;
    }

    private static bool CheckOperandLimit(object operand) {
        switch (operand) {
            case double number:
                return !IsBeyondNumberLimit(number);
            case byte[] bytes:
                return bytes.Length <= MaximumStringBytes;
            case string name:
                return Encoding.UTF8.GetByteCount(name) <= MaximumNameBytes;
            case List<object> items:
                foreach (object item in items) {
                    if (!CheckOperandLimit(item)) return false;
                }
                return true;
            case double[] numbers:
                foreach (double item in numbers) {
                    if (IsBeyondNumberLimit(item)) return false;
                }
                return true;
            default:
                return true;
        }
    }

    private static bool IsBeyondNumberLimit(double number) {
        if (double.IsNaN(number) || double.IsInfinity(number)) return true;
        if (Math.Abs(number) > 3.403e38) return true;
        return number == Math.Floor(number) && (number > int.MaxValue || number < int.MinValue);
    }

    /// <summary>
    /// The distinct device colour uses one content stream paints, including those of nested forms. Keeping each use once
    /// bounds the record when forms draw the same nested form many times.
    /// </summary>
    private sealed class DeviceColorCapture {
        private readonly HashSet<(string Family, PdfObject? Owner, PdfDictionary? Resources)> _seen = new(new UseKeyComparer());

        internal List<DeviceColorUse> Uses { get; } = new();

        internal void Add(DeviceColorUse use) {
            if (_seen.Add((use.Family, use.Owner, use.Resources))) Uses.Add(use);
        }

        private sealed class UseKeyComparer : IEqualityComparer<(string Family, PdfObject? Owner, PdfDictionary? Resources)> {
            public bool Equals((string Family, PdfObject? Owner, PdfDictionary? Resources) x, (string Family, PdfObject? Owner, PdfDictionary? Resources) y) =>
                x.Family == y.Family && ReferenceEquals(x.Owner, y.Owner) && ReferenceEquals(x.Resources, y.Resources);

            public int GetHashCode((string Family, PdfObject? Owner, PdfDictionary? Resources) obj) =>
                (obj.Family.GetHashCode() * 31 + (obj.Owner is null ? 0 : System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(obj.Owner))) * 31 +
                (obj.Resources is null ? 0 : System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(obj.Resources));
        }
    }

    private sealed class ContentKeyComparer : IEqualityComparer<(PdfObject Content, PdfObject? Resources)> {
        public bool Equals((PdfObject Content, PdfObject? Resources) x, (PdfObject Content, PdfObject? Resources) y) =>
            ReferenceEquals(x.Content, y.Content) && ReferenceEquals(x.Resources, y.Resources);

        public int GetHashCode((PdfObject Content, PdfObject? Resources) obj) =>
            System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(obj.Content) * 31 +
            (obj.Resources is null ? 0 : System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(obj.Resources));
    }
}
