using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfFormFiller {
    private static bool TryBuildTransientTextAppearance(Dictionary<int, PdfIndirectObject> objects,
        string? defaultAppearance, PdfDictionary? defaultResources, double width, double height,
        string text, PdfFormFieldStyle style, out string content, out PdfDictionary? resources,
        int maximumDecodedBytes, CancellationToken cancellationToken) {
        content = string.Empty;
        resources = null;
        string displayedText = PdfAcroFormDictionaryBuilder.GetTextFieldAppearanceDisplayValue(text, style);
        if (!TryPrepareTransientTextFont(objects, defaultAppearance, defaultResources, displayedText,
                out string fontName, out TextAppearanceFontPlan? plan, out resources, maximumDecodedBytes, cancellationToken)) return false;
        double fontSize = ReadWidgetAppearanceFontSize(defaultAppearance, height);
        content = PdfAcroFormDictionaryBuilder.BuildTextFieldAppearanceContent(width, height, text, fontSize, style,
            encodedTextHex: plan?.EncodedTextHex, fontResourceName: fontName,
            encodeTextSegmentHex: plan?.EncodeTextSegmentHex, measureTextSegmentWidth: plan?.MeasureTextSegmentWidth,
            singleLineBaseline: Math.Max(2D, height / 2D - fontSize * 0.4D));
        return true;
    }

    private static bool TryPrepareTransientTextFont(Dictionary<int, PdfIndirectObject> objects,
        string? defaultAppearance, PdfDictionary? defaultResources, string displayedText,
        out string fontName, out TextAppearanceFontPlan? plan, out PdfDictionary? resources,
        int maximumDecodedBytes, CancellationToken cancellationToken) {
        resources = null;
        fontName = "Helv";
        plan = null;
        if (PdfDefaultAppearanceParser.TryReadFontResourceName(defaultAppearance, out string selectedName) &&
            defaultResources != null && ResolveDictionary(objects,
                defaultResources.Items.TryGetValue("Font", out PdfObject? fontsObject) ? fontsObject : null) is PdfDictionary fonts &&
            fonts.Items.TryGetValue(selectedName, out PdfObject? fontObject) && ResolveDictionary(objects, fontObject) is PdfDictionary font) {
            if (font.Get<PdfName>("Subtype")?.Name == "Type0") {
                if (!TryCreateInheritedTextAppearanceFontPlan(objects, selectedName, fontObject, displayedText, out plan,
                        maximumDecodedBytes, cancellationToken)) return false;
                resources = plan!.Resources;
                fontName = selectedName;
            } else if (IsTransientWinAnsiFont(objects, font)) {
                TryCreateDefaultAppearanceSimpleFontResources(objects, defaultAppearance, defaultResources, null, null,
                    out _, out resources);
                fontName = selectedName;
            }
        }
        if (plan == null && !PdfWinAnsiEncoding.CanEncode(displayedText, out _, cancellationToken)) return false;
        resources ??= CreateAppearanceResources();
        return true;
    }

    private static bool IsTransientWinAnsiFont(Dictionary<int, PdfIndirectObject> objects, PdfDictionary font) {
        if (font.Get<PdfName>("Subtype")?.Name is not ("Type1" or "TrueType")) return false;
        PdfObject? encoding = ResolveObject(objects, font.Items.TryGetValue("Encoding", out PdfObject? item) ? item : null);
        return encoding is PdfName name ? name.Name == "WinAnsiEncoding" :
            encoding is PdfDictionary dictionary && TryReadName(objects, dictionary, "BaseEncoding") == "WinAnsiEncoding" &&
            !dictionary.Items.ContainsKey("Differences");
    }

    private static bool TryResolveTransientRadioSelection(Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary widget, PdfObject? value, out bool selected) {
        string? state = TryReadName(objects, widget, "AS");
        if (state != null) { selected = state != "Off"; return true; }
        string selectedValue = ReadDisplayedFieldText(objects, value);
        selected = false;
        if (selectedValue.Length == 0 || selectedValue == "Off") return true;
        if (TryGetNormalAppearanceObject(objects, widget, out PdfObject? normal) && normal is PdfDictionary states &&
            states.Items.Keys.Any(key => key != "Off")) {
            selected = states.Items.ContainsKey(selectedValue);
            return true;
        }
        PdfDictionary? parent = ResolveDictionary(objects, widget.Items.TryGetValue("Parent", out PdfObject? item) ? item : null);
        if (parent == null || !parent.Items.TryGetValue("Kids", out PdfObject? kidsObject) ||
            ResolveObject(objects, kidsObject) is not PdfArray kids || TryReadChoiceOptions(objects, parent) is not PdfArray options ||
            options.Items.Count != kids.Items.Count) return false;
        for (int index = 0; index < kids.Items.Count; index++) {
            if (!ReferenceEquals(ResolveDictionary(objects, kids.Items[index]), widget)) continue;
            if (!TryReadOptionText(objects, options.Items[index], out string? exportValue)) return false;
            selected = string.Equals(selectedValue, exportValue, StringComparison.Ordinal);
            return true;
        }
        return false;
    }
}
