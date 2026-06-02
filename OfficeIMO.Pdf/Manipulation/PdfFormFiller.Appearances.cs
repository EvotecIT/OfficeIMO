namespace OfficeIMO.Pdf;

public static partial class PdfFormFiller {
    private static void SetWidgetAppearanceStates(Dictionary<int, PdfIndirectObject> objects, PdfDictionary field, string name, bool isRadioButtonGroup, HashSet<int> visited, ref int nextObjectNumber) {
        if (IsWidget(field)) {
            string appearanceState = isRadioButtonGroup && !HasButtonNormalAppearanceState(objects, field, name) ? "Off" : name;
            field.Items["AS"] = new PdfName(appearanceState);
            EnsureButtonWidgetAppearances(objects, field, appearanceState, ref nextObjectNumber);
        }

        if (!field.Items.TryGetValue("Kids", out var kidsObject) ||
            ResolveObject(objects, kidsObject) is not PdfArray kids) {
            return;
        }

        for (int i = 0; i < kids.Items.Count; i++) {
            PdfObject kidObject = kids.Items[i];
            if (kidObject is PdfReference reference && !visited.Add(reference.ObjectNumber)) {
                continue;
            }

            if (ResolveObject(objects, kidObject) is PdfDictionary kid) {
                SetWidgetAppearanceStates(objects, kid, name, isRadioButtonGroup, visited, ref nextObjectNumber);
            }
        }
    }

    private static HashSet<string> CollectButtonNormalAppearanceStates(Dictionary<int, PdfIndirectObject> objects, PdfDictionary field, HashSet<int> visited) {
        var states = new HashSet<string>(StringComparer.Ordinal);
        CollectButtonNormalAppearanceStates(objects, field, states, visited);
        states.Remove("Off");
        return states;
    }

    private static void CollectButtonNormalAppearanceStates(Dictionary<int, PdfIndirectObject> objects, PdfDictionary field, HashSet<string> states, HashSet<int> visited) {
        if (IsWidget(field) &&
            TryGetNormalAppearanceObject(objects, field, out PdfObject? normalAppearance) &&
            normalAppearance is PdfDictionary appearanceStates) {
            foreach (string stateName in appearanceStates.Items.Keys) {
                states.Add(stateName);
            }
        }

        if (!field.Items.TryGetValue("Kids", out var kidsObject) ||
            ResolveObject(objects, kidsObject) is not PdfArray kids) {
            return;
        }

        for (int i = 0; i < kids.Items.Count; i++) {
            PdfObject kidObject = kids.Items[i];
            if (kidObject is PdfReference reference && !visited.Add(reference.ObjectNumber)) {
                continue;
            }

            if (ResolveObject(objects, kidObject) is PdfDictionary kid) {
                CollectButtonNormalAppearanceStates(objects, kid, states, visited);
            }
        }
    }

    private static bool HasButtonNormalAppearanceState(Dictionary<int, PdfIndirectObject> objects, PdfDictionary widget, string stateName) {
        if (string.IsNullOrEmpty(stateName)) {
            return false;
        }

        return TryGetNormalAppearanceObject(objects, widget, out PdfObject? normalAppearance) &&
            normalAppearance is PdfDictionary appearanceStates &&
            appearanceStates.Items.ContainsKey(stateName);
    }

    private static void EnsureButtonWidgetAppearances(Dictionary<int, PdfIndirectObject> objects, PdfDictionary widget, string selectedName, ref int nextObjectNumber) {
        if (!TryReadRect(widget, out double width, out double height)) {
            return;
        }

        PdfDictionary normalAppearances = GetOrCreateButtonNormalAppearanceDictionary(objects, widget);
        if (!normalAppearances.Items.ContainsKey("Off")) {
            int offAppearanceObjectNumber = nextObjectNumber++;
            objects[offAppearanceObjectNumber] = new PdfIndirectObject(offAppearanceObjectNumber, 0, CreateButtonAppearanceStream(width, height, selected: false, ReadWidgetAppearanceStyle(objects, widget)));
            normalAppearances.Items["Off"] = new PdfReference(offAppearanceObjectNumber, 0);
        }

        if (!string.Equals(selectedName, "Off", StringComparison.Ordinal) && !normalAppearances.Items.ContainsKey(selectedName)) {
            int selectedAppearanceObjectNumber = nextObjectNumber++;
            objects[selectedAppearanceObjectNumber] = new PdfIndirectObject(selectedAppearanceObjectNumber, 0, CreateButtonAppearanceStream(width, height, selected: true, ReadWidgetAppearanceStyle(objects, widget)));
            normalAppearances.Items[selectedName] = new PdfReference(selectedAppearanceObjectNumber, 0);
        }
    }

    private static PdfDictionary GetOrCreateButtonNormalAppearanceDictionary(Dictionary<int, PdfIndirectObject> objects, PdfDictionary widget) {
        PdfDictionary appearance;
        if (widget.Items.TryGetValue("AP", out var appearanceObject) &&
            ResolveDictionary(objects, appearanceObject) is PdfDictionary existingAppearance) {
            appearance = existingAppearance;
        } else {
            appearance = new PdfDictionary();
            widget.Items["AP"] = appearance;
        }

        if (appearance.Items.TryGetValue("N", out var normalAppearanceObject) &&
            ResolveDictionary(objects, normalAppearanceObject) is PdfDictionary existingNormalAppearance) {
            return existingNormalAppearance;
        }

        var normalAppearances = new PdfDictionary();
        appearance.Items["N"] = normalAppearances;
        return normalAppearances;
    }

    private static void SetTextWidgetAppearances(Dictionary<int, PdfIndirectObject> objects, PdfDictionary field, string value, HashSet<int> visited, ref int nextObjectNumber) {
        if (IsWidget(field) && TryReadRect(field, out double width, out double height)) {
            int appearanceObjectNumber = nextObjectNumber++;
            objects[appearanceObjectNumber] = new PdfIndirectObject(appearanceObjectNumber, 0, CreateTextAppearanceStream(value, width, height, ReadWidgetAppearanceStyle(objects, field)));

            var appearance = new PdfDictionary();
            appearance.Items["N"] = new PdfReference(appearanceObjectNumber, 0);
            field.Items["AP"] = appearance;
        }

        if (!field.Items.TryGetValue("Kids", out var kidsObject) ||
            ResolveObject(objects, kidsObject) is not PdfArray kids) {
            return;
        }

        for (int i = 0; i < kids.Items.Count; i++) {
            PdfObject kidObject = kids.Items[i];
            if (kidObject is PdfReference reference && !visited.Add(reference.ObjectNumber)) {
                continue;
            }

            if (ResolveObject(objects, kidObject) is PdfDictionary kid) {
                SetTextWidgetAppearances(objects, kid, value, visited, ref nextObjectNumber);
            }
        }
    }

    private static PdfStream CreateTextAppearanceStream(string value, double width, double height, PdfFormFieldStyle? style = null) {
        double fontSize = Math.Max(6D, Math.Min(12D, height - 4D));
        string content = PdfAcroFormDictionaryBuilder.BuildTextFieldAppearanceContent(width, height, value, fontSize, style);

        var dictionary = new PdfDictionary();
        dictionary.Items["Type"] = new PdfName("XObject");
        dictionary.Items["Subtype"] = new PdfName("Form");
        dictionary.Items["BBox"] = CreateNumberArray(0D, 0D, width, height);
        dictionary.Items["Resources"] = CreateAppearanceResources();
        return new PdfStream(dictionary, PdfEncoding.Latin1GetBytes(content));
    }

    private static PdfStream CreateButtonAppearanceStream(double width, double height, bool selected, PdfFormFieldStyle? style = null) {
        string content = PdfAcroFormDictionaryBuilder.BuildCheckBoxAppearanceContent(width, height, selected, style);
        var dictionary = new PdfDictionary();
        dictionary.Items["Type"] = new PdfName("XObject");
        dictionary.Items["Subtype"] = new PdfName("Form");
        dictionary.Items["BBox"] = CreateNumberArray(0D, 0D, width, height);
        return new PdfStream(dictionary, PdfEncoding.Latin1GetBytes(content));
    }

    private static PdfFormFieldStyle ReadWidgetAppearanceStyle(Dictionary<int, PdfIndirectObject> objects, PdfDictionary widget) {
        var style = new PdfFormFieldStyle();
        if (ResolveDictionary(objects, widget.Items.TryGetValue("MK", out var mkObject) ? mkObject : null) is PdfDictionary mk) {
            if (TryReadColor(objects, mk, "BG", out PdfColor backgroundColor)) {
                style.BackgroundColor = backgroundColor;
            }

            if (TryReadColor(objects, mk, "BC", out PdfColor borderColor)) {
                style.BorderColor = borderColor;
            }
        }

        if (TryReadDefaultAppearanceTextColor(objects, widget, out PdfColor textColor)) {
            style.TextColor = textColor;
        }

        return style;
    }

    private static bool TryReadColor(Dictionary<int, PdfIndirectObject> objects, PdfDictionary dictionary, string key, out PdfColor color) {
        color = default;
        if (!dictionary.Items.TryGetValue(key, out var colorObject) ||
            ResolveObject(objects, colorObject) is not PdfArray colorArray ||
            colorArray.Items.Count < 3 ||
            ResolveObject(objects, colorArray.Items[0]) is not PdfNumber red ||
            ResolveObject(objects, colorArray.Items[1]) is not PdfNumber green ||
            ResolveObject(objects, colorArray.Items[2]) is not PdfNumber blue ||
            red.Value < 0 || red.Value > 1 ||
            green.Value < 0 || green.Value > 1 ||
            blue.Value < 0 || blue.Value > 1) {
            return false;
        }

        color = new PdfColor(red.Value, green.Value, blue.Value);
        return true;
    }

    private static bool TryReadDefaultAppearanceTextColor(Dictionary<int, PdfIndirectObject> objects, PdfDictionary widget, out PdfColor color) {
        color = default;
        string? defaultAppearance = TryReadText(objects, widget, "DA");
        if (string.IsNullOrWhiteSpace(defaultAppearance)) {
            return false;
        }

        string[] parts = defaultAppearance!.Split(DefaultAppearanceSeparators, StringSplitOptions.RemoveEmptyEntries);
        for (int i = 3; i < parts.Length; i++) {
            if (!string.Equals(parts[i], "rg", StringComparison.Ordinal)) {
                continue;
            }

            if (double.TryParse(parts[i - 3], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double red) &&
                double.TryParse(parts[i - 2], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double green) &&
                double.TryParse(parts[i - 1], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double blue) &&
                red >= 0 && red <= 1 &&
                green >= 0 && green <= 1 &&
                blue >= 0 && blue <= 1) {
                color = new PdfColor(red, green, blue);
                return true;
            }
        }

        return false;
    }

    private static PdfDictionary CreateAppearanceResources() {
        var font = new PdfDictionary();
        font.Items["Type"] = new PdfName("Font");
        font.Items["Subtype"] = new PdfName("Type1");
        font.Items["BaseFont"] = new PdfName("Helvetica");

        var fonts = new PdfDictionary();
        fonts.Items["Helv"] = font;

        var resources = new PdfDictionary();
        resources.Items["Font"] = fonts;
        return resources;
    }

    private static PdfArray CreateNumberArray(params double[] values) {
        var array = new PdfArray();
        foreach (double value in values) {
            array.Items.Add(new PdfNumber(value));
        }

        return array;
    }

    private static PdfArray CreateStringArray(IEnumerable<string> values) {
        var array = new PdfArray();
        foreach (string value in values) {
            array.Items.Add(new PdfStringObj(value, useTextStringEncoding: true));
        }

        return array;
    }

    private static bool IsWidget(PdfDictionary dictionary) {
        return dictionary.Items.TryGetValue("Subtype", out var subtype) &&
            subtype is PdfName name &&
            string.Equals(name.Name, "Widget", StringComparison.Ordinal);
    }

    private static bool TryReadRect(PdfDictionary dictionary, out double width, out double height) {
        if (TryReadRectCoordinates(dictionary, out _, out _, out width, out height)) {
            return true;
        }

        width = 0D;
        height = 0D;
        return false;
    }

    private static bool TryReadRectCoordinates(PdfDictionary dictionary, out double x, out double y, out double width, out double height) {
        x = 0D;
        y = 0D;
        width = 0D;
        height = 0D;
        if (!dictionary.Items.TryGetValue("Rect", out var rectObject) ||
            rectObject is not PdfArray rect ||
            rect.Items.Count < 4 ||
            rect.Items[0] is not PdfNumber x1 ||
            rect.Items[1] is not PdfNumber y1 ||
            rect.Items[2] is not PdfNumber x2 ||
            rect.Items[3] is not PdfNumber y2) {
            return false;
        }

        x = Math.Min(x1.Value, x2.Value);
        y = Math.Min(y1.Value, y2.Value);
        width = Math.Abs(x2.Value - x1.Value);
        height = Math.Abs(y2.Value - y1.Value);
        return width > 0D && height > 0D;
    }

    private static string GetButtonWidgetFlattenAppearanceState(Dictionary<int, PdfIndirectObject> objects, PdfDictionary widget, string? inheritedValue) {
        string? widgetState = TryReadName(objects, widget, "AS");
        if (!string.IsNullOrEmpty(widgetState)) {
            return widgetState!;
        }

        if (!string.IsNullOrEmpty(inheritedValue) &&
            HasButtonNormalAppearanceState(objects, widget, inheritedValue!)) {
            return inheritedValue!;
        }

        return "Off";
    }

    private static bool TryGetNormalAppearanceReference(Dictionary<int, PdfIndirectObject> objects, PdfDictionary widget, out PdfReference? reference) {
        reference = null;
        if (!TryGetNormalAppearanceObject(objects, widget, out PdfObject? normalAppearance) ||
            normalAppearance is not PdfReference normalAppearanceReference) {
            return false;
        }

        reference = normalAppearanceReference;
        return true;
    }

    private static bool TryGetButtonAppearanceReference(Dictionary<int, PdfIndirectObject> objects, PdfDictionary widget, string inheritedValue, out PdfReference? reference) {
        reference = null;
        if (!TryGetNormalAppearanceObject(objects, widget, out PdfObject? normalAppearance)) {
            return false;
        }

        if (normalAppearance is PdfReference singleAppearanceReference) {
            reference = singleAppearanceReference;
            return true;
        }

        if (normalAppearance is not PdfDictionary appearanceStates) {
            return false;
        }

        if (!string.IsNullOrEmpty(inheritedValue) &&
            TryGetAppearanceStateReference(appearanceStates, inheritedValue, out reference)) {
            return true;
        }

        return TryGetAppearanceStateReference(appearanceStates, "Off", out reference);
    }

    private static bool TryGetNormalAppearanceObject(Dictionary<int, PdfIndirectObject> objects, PdfDictionary widget, out PdfObject? normalAppearance) {
        normalAppearance = null;
        if (!widget.Items.TryGetValue("AP", out var appearanceObject) ||
            ResolveDictionary(objects, appearanceObject) is not PdfDictionary appearance ||
            !appearance.Items.TryGetValue("N", out var normalAppearanceObject)) {
            return false;
        }

        if (normalAppearanceObject is PdfReference normalAppearanceReference) {
            PdfObject? resolved = ResolveObject(objects, normalAppearanceReference);
            normalAppearance = resolved is PdfStream ? normalAppearanceReference : resolved;
            return normalAppearance is not null;
        }

        normalAppearance = normalAppearanceObject;
        return true;
    }

    private static bool TryGetAppearanceStateReference(PdfDictionary appearanceStates, string stateName, out PdfReference? reference) {
        reference = null;
        if (!appearanceStates.Items.TryGetValue(stateName, out var stateAppearance) ||
            stateAppearance is not PdfReference stateAppearanceReference) {
            return false;
        }

        reference = stateAppearanceReference;
        return true;
    }
}
