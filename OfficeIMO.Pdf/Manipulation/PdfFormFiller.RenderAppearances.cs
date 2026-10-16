using System.Threading;

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
        out PdfStream stream, int maximumDecodedBytes = 256 * 1024 * 1024, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        stream = null!;
        if (!IsWidget(widget) || !TryReadRect(widget, out double width, out double height) || !(width > 0D) || !(height > 0D)) return false;

        string? fieldType = null, defaultAppearance = null;
        PdfObject? value = null;
        PdfArray? choiceOptions = null;
        PdfDictionary? defaultResources = null;
        int? flags = null, quadding = null, maxLength = null;
        var visited = new HashSet<PdfDictionary>();
        PdfDictionary? field = widget;
        for (int depth = 0; field != null && depth < MaximumFieldInheritanceDepth && visited.Add(field); depth++) {
            fieldType ??= TryReadName(objects, field, "FT");
            choiceOptions ??= TryReadChoiceOptions(objects, field);
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
                    if (fieldType == "Ch" && choiceOptions != null) {
                        var lookup = ChoiceOptionLookup.Create(objects, choiceOptions);
                        text = lookup.ResolveDisplayValue(text) ?? text;
                    }
                    if (!TryBuildTransientTextAppearance(objects, defaultAppearance, defaultResources, width, height,
                            text, style, out content, out resources, maximumDecodedBytes, cancellationToken)) return false;
                    break;
                case "Btn":
                    if ((fieldFlags & PushButtonFlag) != 0) {
                        string caption = characteristics == null ? string.Empty : TryReadText(objects, characteristics, "CA") ?? string.Empty;
                        style.TextAlignment = PdfFormFieldTextAlignment.Center;
                        if (!TryBuildTransientTextAppearance(objects, defaultAppearance, defaultResources, width, height,
                                caption, style, out content, out resources, maximumDecodedBytes, cancellationToken)) return false;
                        break;
                    }
                    string? state = TryReadName(objects, widget, "AS") ?? (value as PdfName)?.Name;
                    bool selected = state != null && !string.Equals(state, "Off", StringComparison.Ordinal);
                    if ((fieldFlags & RadioButtonFlag) != 0 &&
                        !TryResolveTransientRadioSelection(objects, widget, value, out selected)) return false;
                    content = (fieldFlags & RadioButtonFlag) != 0
                        ? PdfAcroFormDictionaryBuilder.BuildRadioButtonAppearanceContent(width, height, selected, style)
                        : PdfAcroFormDictionaryBuilder.BuildCheckBoxAppearanceContent(width, height, selected, style);
                    break;
                default:
                    return false;
            }
        } catch (ArgumentException) {
            return false;
        } catch (Exception exception) when (exception is NotSupportedException || exception is IOException && exception is not PdfReadLimitException) {
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
