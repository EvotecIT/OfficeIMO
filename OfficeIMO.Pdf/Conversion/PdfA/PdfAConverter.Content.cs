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
        internal DeviceColorUse(string family, string location, PdfObject? owner, PdfDictionary? resources) {
            Family = family;
            Location = location;
            Owner = owner;
            Resources = resources;
        }

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

    private void AnalyzeContent() {
        IReadOnlyList<PdfReadPage> pages = _document.Pages;
        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++) {
            _cancellationToken.ThrowIfCancellationRequested();
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
        if (!_visitedContent.Add((key, resources))) return;
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
                    if (!reportedLimit && !CheckOperandLimits(operands)) {
                        reportedLimit = true;
                        AddViolation(PdfAConversionRefusalReason.ImplementationLimitExceeded, location,
                            "The content stream has a number, string or name beyond the PDF/A implementation limits.", owner);
                    }
                    switch (operation.Name) {
                        case "q":
                            stack.Push((font, fontName, renderMode));
                            nesting++;
                            if (nesting > MaximumGraphicsStateNesting && !reportedNesting) {
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
                            if (operands.Count >= 1 && operands[0] is string space) CheckColorSpaceName(space, resources, location, owner);
                            break;
                        case "sh":
                            if (operands.Count >= 1 && operands[0] is string shadingName) {
                                PdfObject? shading = ResolveResourceObject(resources, "Shading", shadingName);
                                if (shading is not null) CheckShading(shading, resources, location);
                            }
                            break;
                        case "Do":
                            if (operands.Count >= 1 && operands[0] is string xobjectName &&
                                ResolveResourceObject(resources, "XObject", xobjectName) is PdfStream xobject) {
                                self.CheckXObject(xobject, resources, location + ", XObject " + xobjectName, depth);
                            }
                            break;
                        case "BI":
                            if (operation.InlineImage is not null) CheckInlineImage(operation.InlineImage.Dictionary, resources, location, owner);
                            break;
                        case "ri":
                            if (operands.Count >= 1 && operands[0] is string intent && !RenderingIntents.Contains(intent)) {
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
        foreach (DeviceColorUse existing in _deviceColors) {
            if (existing.Family == family && ReferenceEquals(existing.Owner, owner) && ReferenceEquals(existing.Resources, resources)) return;
        }
        _deviceColors.Add(new DeviceColorUse(family, location, owner, resources));
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

    private sealed class ContentKeyComparer : IEqualityComparer<(PdfObject Content, PdfObject? Resources)> {
        public bool Equals((PdfObject Content, PdfObject? Resources) x, (PdfObject Content, PdfObject? Resources) y) =>
            ReferenceEquals(x.Content, y.Content) && ReferenceEquals(x.Resources, y.Resources);

        public int GetHashCode((PdfObject Content, PdfObject? Resources) obj) =>
            System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(obj.Content) * 31 +
            (obj.Resources is null ? 0 : System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(obj.Resources));
    }
}
