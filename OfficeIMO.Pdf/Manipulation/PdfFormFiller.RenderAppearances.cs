namespace OfficeIMO.Pdf;

internal static partial class PdfFormFiller {
    private const int MaximumFieldInheritanceDepth = 32;

    /// <summary>
    /// Builds a transient normal appearance for a text, choice or button widget from its inherited field value,
    /// default appearance and /MK colours, as viewers do for widgets without an appearance or when the AcroForm
    /// sets NeedAppearances. Nothing is added to <paramref name="objects"/>; fonts come from the field or AcroForm
    /// /DR, falling back to Helvetica.
    /// </summary>
    internal static bool TryCreateTransientWidgetAppearance(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary widget,
        PdfDictionary? acroForm,
        out PdfStream stream) {
        stream = null!;
        if (!IsWidget(widget) || !TryReadRect(widget, out double width, out double height) || !(width > 0D) || !(height > 0D)) return false;

        string? fieldType = null, defaultAppearance = null;
        PdfObject? value = null;
        PdfDictionary? defaultResources = null;
        int? flags = null, quadding = null, maxLength = null;
        var visited = new HashSet<PdfDictionary>();
        PdfDictionary? field = widget;
        for (int depth = 0; field != null && depth < MaximumFieldInheritanceDepth && visited.Add(field); depth++) {
            fieldType ??= TryReadName(objects, field, "FT");
            defaultAppearance ??= TryReadText(objects, field, "DA");
            defaultResources ??= TryReadDefaultResources(objects, field);
            if (value == null && field.Items.TryGetValue("V", out PdfObject? valueObject)) value = ResolveObject(objects, valueObject);
            if (flags == null && ResolveObject(objects, field.Items.TryGetValue("Ff", out PdfObject? flagsObject) ? flagsObject : null) is PdfNumber flagNumber) {
                flags = (int)flagNumber.Value;
            }
            if (quadding == null && TryReadQuadding(objects, field, out int fieldQuadding)) quadding = fieldQuadding;
            if (maxLength == null && TryReadMaxLength(objects, field, out int fieldMaxLength)) maxLength = fieldMaxLength;
            field = ResolveDictionary(objects, field.Items.TryGetValue("Parent", out PdfObject? parentObject) ? parentObject : null);
        }
        if (acroForm != null) {
            defaultAppearance ??= TryReadText(objects, acroForm, "DA");
            defaultResources ??= TryReadDefaultResources(objects, acroForm);
            if (quadding == null && TryReadQuadding(objects, acroForm, out int formQuadding)) quadding = formQuadding;
        }

        int fieldFlags = flags ?? 0;
        PdfFormFieldStyle style = ReadWidgetAppearanceStyle(objects, widget, fieldFlags, quadding, maxLength, defaultAppearance);
        // Authoring defaults do not apply: without /MK colours a viewer paints no background or border.
        PdfDictionary? characteristics = ResolveDictionary(objects, widget.Items.TryGetValue("MK", out PdfObject? characteristicsObject) ? characteristicsObject : null);
        if (characteristics == null || !TryReadColor(objects, characteristics, "BG", out _)) style.BackgroundColor = null;
        if (characteristics == null || !TryReadColor(objects, characteristics, "BC", out _)) style.BorderColor = null;
        string content;
        PdfDictionary? resources = null;
        try {
            switch (fieldType) {
                case "Tx":
                case "Ch":
                    string text = ReadDisplayedFieldText(objects, value);
                    if (!TryCreateDefaultAppearanceSimpleFontResources(objects, defaultAppearance, defaultResources, null, null,
                            out string? fontResourceName, out resources)) {
                        resources = CreateAppearanceResources();
                        fontResourceName = "Helv";
                    }
                    double fontSize = ReadWidgetAppearanceFontSize(defaultAppearance, height);
                    // Single-line text is centred the way viewers regenerate it: baseline at half height less 0.4 em.
                    content = PdfAcroFormDictionaryBuilder.BuildTextFieldAppearanceContent(width, height, text, fontSize, style,
                        fontResourceName: fontResourceName, singleLineBaseline: Math.Max(2D, height / 2D - fontSize * 0.4D));
                    break;
                case "Btn":
                    if ((fieldFlags & PushButtonFlag) != 0) {
                        content = PdfAcroFormDictionaryBuilder.BuildCheckBoxAppearanceContent(width, height, selected: false, style);
                        break;
                    }
                    string? state = TryReadName(objects, widget, "AS") ?? (value as PdfName)?.Name;
                    bool selected = state != null && !string.Equals(state, "Off", StringComparison.Ordinal);
                    content = (fieldFlags & RadioButtonFlag) != 0
                        ? PdfAcroFormDictionaryBuilder.BuildRadioButtonAppearanceContent(width, height, selected, style)
                        : PdfAcroFormDictionaryBuilder.BuildCheckBoxAppearanceContent(width, height, selected, style);
                    break;
                default:
                    return false;
            }
        } catch (ArgumentException) {
            return false;
        }

        var dictionary = new PdfDictionary();
        dictionary.Items["Type"] = new PdfName("XObject");
        dictionary.Items["Subtype"] = new PdfName("Form");
        dictionary.Items["BBox"] = CreateNumberArray(0D, 0D, width, height);
        if (resources != null) dictionary.Items["Resources"] = resources;
        stream = new PdfStream(dictionary, PdfEncoding.Latin1GetBytes(content));
        return true;
    }

    /// <summary>True for widgets of text and choice fields, whose appearance NeedAppearances asks viewers to regenerate.</summary>
    internal static bool IsVariableTextWidget(Dictionary<int, PdfIndirectObject> objects, PdfDictionary widget) {
        if (!IsWidget(widget)) return false;
        var visited = new HashSet<PdfDictionary>();
        PdfDictionary? field = widget;
        for (int depth = 0; field != null && depth < MaximumFieldInheritanceDepth && visited.Add(field); depth++) {
            string? fieldType = TryReadName(objects, field, "FT");
            if (fieldType != null) return fieldType is "Tx" or "Ch";
            field = ResolveDictionary(objects, field.Items.TryGetValue("Parent", out PdfObject? parentObject) ? parentObject : null);
        }
        return false;
    }

    // Choice fields show their first selected value on one line.
    private static string ReadDisplayedFieldText(Dictionary<int, PdfIndirectObject> objects, PdfObject? value) => value switch {
        PdfStringObj text => text.Value ?? string.Empty,
        PdfName name => name.Name,
        PdfArray array when array.Items.Count > 0 => ResolveObject(objects, array.Items[0]) is PdfStringObj first ? first.Value ?? string.Empty : string.Empty,
        _ => string.Empty
    };
}
