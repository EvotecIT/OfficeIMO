namespace OfficeIMO.Pdf;

internal static partial class PdfFormFiller {
    private static void SetWidgetAppearanceStates(Dictionary<int, PdfIndirectObject> objects, PdfDictionary field, string name, bool isRadioButtonGroup, HashSet<int> visited, ref int nextObjectNumber) {
        if (IsWidget(field)) {
            string appearanceState = isRadioButtonGroup && !HasButtonNormalAppearanceState(objects, field, name) ? "Off" : name;
            field.Items["AS"] = new PdfName(appearanceState);
            EnsureButtonWidgetAppearances(objects, field, appearanceState, isRadioButtonGroup, ref nextObjectNumber);
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

    private static void EnsureButtonWidgetAppearances(Dictionary<int, PdfIndirectObject> objects, PdfDictionary widget, string selectedName, bool isRadioButton, ref int nextObjectNumber) {
        if (!TryReadRect(widget, out double width, out double height)) {
            return;
        }

        PdfDictionary normalAppearances = GetOrCreateButtonNormalAppearanceDictionary(objects, widget);
        if (!normalAppearances.Items.ContainsKey("Off")) {
            int offAppearanceObjectNumber = nextObjectNumber++;
            objects[offAppearanceObjectNumber] = new PdfIndirectObject(offAppearanceObjectNumber, 0, CreateButtonAppearanceStream(width, height, selected: false, isRadioButton, ReadWidgetAppearanceStyle(objects, widget)));
            normalAppearances.Items["Off"] = new PdfReference(offAppearanceObjectNumber, 0);
        }

        if (!string.Equals(selectedName, "Off", StringComparison.Ordinal) && !normalAppearances.Items.ContainsKey(selectedName)) {
            int selectedAppearanceObjectNumber = nextObjectNumber++;
            objects[selectedAppearanceObjectNumber] = new PdfIndirectObject(selectedAppearanceObjectNumber, 0, CreateButtonAppearanceStream(width, height, selected: true, isRadioButton, ReadWidgetAppearanceStyle(objects, widget)));
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

    private static void SetTextWidgetAppearances(Dictionary<int, PdfIndirectObject> objects, PdfDictionary field, string value, string? fieldName, int inheritedFlags, int? inheritedQuadding, int? inheritedMaxLength, PdfDictionary? inheritedDefaultResources, string? inheritedDefaultAppearance, bool forceMultilineAppearance, PdfFormFillerOptions? options, HashSet<int> visited, ref int nextObjectNumber) {
        int fieldFlags = ReadFieldFlags(objects, field, inheritedFlags);
        int? fieldQuadding = ReadFieldQuadding(objects, field, inheritedQuadding);
        int? fieldMaxLength = ReadFieldMaxLength(objects, field, inheritedMaxLength);
        PdfDictionary? defaultResources = TryReadDefaultResources(objects, field) ?? inheritedDefaultResources;
        string? defaultAppearance = TryReadText(objects, field, "DA") ?? inheritedDefaultAppearance;
        if (IsWidget(field) && TryReadRect(field, out double width, out double height)) {
            PdfDictionary? widgetAppearanceResources = TryReadNormalAppearanceResources(objects, field);
            PdfDictionary? widgetPageResources = TryReadWidgetPageResources(objects, field);
            PdfFormFieldStyle widgetStyle = ReadWidgetAppearanceStyle(objects, field, fieldFlags, fieldQuadding, fieldMaxLength, defaultAppearance);
            if (forceMultilineAppearance) {
                widgetStyle.IsMultiline = true;
            }

            int appearanceObjectNumber = nextObjectNumber++;
            objects[appearanceObjectNumber] = new PdfIndirectObject(appearanceObjectNumber, 0, CreateTextAppearanceStream(objects, defaultResources, widgetAppearanceResources, widgetPageResources, value, width, height, widgetStyle, defaultAppearance, ReadWidgetAppearanceFontSize(defaultAppearance, height), options, fieldName, ref nextObjectNumber));

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
                SetTextWidgetAppearances(objects, kid, value, fieldName, fieldFlags, fieldQuadding, fieldMaxLength, defaultResources, defaultAppearance, forceMultilineAppearance, options, visited, ref nextObjectNumber);
            }
        }
    }

    private static PdfStream CreateTextAppearanceStream(Dictionary<int, PdfIndirectObject> objects, PdfDictionary? inheritedDefaultResources, PdfDictionary? widgetAppearanceResources, PdfDictionary? widgetPageResources, string value, double width, double height, PdfFormFieldStyle? style, string? defaultAppearance, double fontSize, PdfFormFillerOptions? options, string? fieldName, ref int nextObjectNumber, IReadOnlyList<PdfFreeTextRichTextRun>? richAppearanceRuns = null) {
        PdfFormFieldStyle effectiveStyle = style ?? new PdfFormFieldStyle();
        if (richAppearanceRuns != null && !effectiveStyle.IsPassword && !effectiveStyle.IsComb) {
            return CreateRichTextAppearanceStream(richAppearanceRuns, width, height, effectiveStyle, fontSize);
        }

        string displayValue = PdfAcroFormDictionaryBuilder.GetTextFieldAppearanceDisplayValue(value, effectiveStyle);
        string diagnosticSource = CreateTextAppearanceDiagnosticSource(fieldName);
        bool hasEmbeddedAppearanceFont = TryCreateInheritedTextAppearanceFontPlan(objects, inheritedDefaultResources, widgetAppearanceResources, widgetPageResources, displayValue, out TextAppearanceFontPlan? fontPlan);
        bool hasDefaultAppearanceSimpleFont = false;
        string? defaultAppearanceFontResourceName = null;
        PdfDictionary? defaultAppearanceFontResources = null;
        if (!hasEmbeddedAppearanceFont) {
            hasEmbeddedAppearanceFont = TryCreateEmbeddedTextAppearanceFontPlan(options, displayValue, diagnosticSource, ref nextObjectNumber, out fontPlan, out string? configuredFontFailure);
            if (!hasEmbeddedAppearanceFont) {
                hasEmbeddedAppearanceFont = TryCreateFallbackTextAppearanceFontPlan(options, displayValue, diagnosticSource, ref nextObjectNumber, out fontPlan, out string? fallbackFontFailure);
                if (!hasEmbeddedAppearanceFont &&
                    (options?.HasAppearanceFontFamily == true || options?.HasAppearanceFontFallbacks == true) &&
                    !string.IsNullOrEmpty(displayValue)) {
                    throw new InvalidOperationException(fallbackFontFailure ?? configuredFontFailure ?? "The configured appearance font could not be used for the form field appearance.");
                }
            }

            if (!hasEmbeddedAppearanceFont) {
                hasDefaultAppearanceSimpleFont = TryCreateDefaultAppearanceSimpleFontResources(objects, defaultAppearance, inheritedDefaultResources, widgetAppearanceResources, widgetPageResources, out defaultAppearanceFontResourceName, out defaultAppearanceFontResources);
            }
        }

        string content = PdfAcroFormDictionaryBuilder.BuildTextFieldAppearanceContent(
            width,
            height,
            value,
            fontSize,
            effectiveStyle,
            fontPlan?.EncodedTextHex,
            fontResourceName: fontPlan?.FontResourceName ?? defaultAppearanceFontResourceName,
            encodeTextSegmentHex: fontPlan?.EncodeTextSegmentHex,
            measureTextSegmentWidth: fontPlan?.MeasureTextSegmentWidth,
            encodeTextSegments: fontPlan?.EncodeTextSegments);
        fontPlan?.Materialize(objects);

        var dictionary = new PdfDictionary();
        dictionary.Items["Type"] = new PdfName("XObject");
        dictionary.Items["Subtype"] = new PdfName("Form");
        dictionary.Items["BBox"] = CreateNumberArray(0D, 0D, width, height);
        dictionary.Items["Resources"] = hasEmbeddedAppearanceFont
            ? fontPlan!.Resources
            : hasDefaultAppearanceSimpleFont
                ? defaultAppearanceFontResources!
                : CreateAppearanceResources();
        return new PdfStream(dictionary, PdfEncoding.Latin1GetBytes(content));
    }

    private static PdfStream CreateRichTextAppearanceStream(IReadOnlyList<PdfFreeTextRichTextRun> richRuns, double width, double height, PdfFormFieldStyle style, double fontSize) {
        string content = PdfAnnotationDictionaryBuilder.BuildFreeTextRichAppearanceContent(
            width,
            height,
            richRuns,
            out IReadOnlyList<(string Name, PdfStandardFont Font)> fontResources,
            fontSize,
            style.TextColor,
            style.BorderColor,
            style.BorderWidth,
            style.BackgroundColor,
            MapFormTextAlignment(style.TextAlignment),
            padding: 3D,
            borderDashPattern: style.BorderDashPattern,
            borderStyle: style.BorderStyle);

        var dictionary = new PdfDictionary();
        dictionary.Items["Type"] = new PdfName("XObject");
        dictionary.Items["Subtype"] = new PdfName("Form");
        dictionary.Items["BBox"] = CreateNumberArray(0D, 0D, width, height);
        dictionary.Items["Resources"] = CreateRichTextAppearanceResources(fontResources);
        return new PdfStream(dictionary, PdfEncoding.Latin1GetBytes(content));
    }

    private static PdfAlign MapFormTextAlignment(PdfFormFieldTextAlignment? alignment) {
        switch (alignment) {
            case PdfFormFieldTextAlignment.Center:
                return PdfAlign.Center;
            case PdfFormFieldTextAlignment.Right:
                return PdfAlign.Right;
            default:
                return PdfAlign.Left;
        }
    }

    private static string CreateTextAppearanceDiagnosticSource(string? fieldName) =>
        string.IsNullOrWhiteSpace(fieldName)
            ? "form field appearance"
            : "form field '" + fieldName + "' appearance";

    private static PdfStream CreateButtonAppearanceStream(double width, double height, bool selected, bool isRadioButton, PdfFormFieldStyle? style = null) {
        string content = isRadioButton
            ? PdfAcroFormDictionaryBuilder.BuildRadioButtonAppearanceContent(width, height, selected, style)
            : PdfAcroFormDictionaryBuilder.BuildCheckBoxAppearanceContent(width, height, selected, style);
        var dictionary = new PdfDictionary();
        dictionary.Items["Type"] = new PdfName("XObject");
        dictionary.Items["Subtype"] = new PdfName("Form");
        dictionary.Items["BBox"] = CreateNumberArray(0D, 0D, width, height);
        return new PdfStream(dictionary, PdfEncoding.Latin1GetBytes(content));
    }

    private static PdfFormFieldStyle ReadWidgetAppearanceStyle(Dictionary<int, PdfIndirectObject> objects, PdfDictionary widget, int fieldFlags = 0, int? inheritedQuadding = null, int? inheritedMaxLength = null, string? inheritedDefaultAppearance = null) {
        var style = new PdfFormFieldStyle();
        style.IsMultiline = (fieldFlags & MultilineFlag) != 0;
        style.IsPassword = (fieldFlags & PasswordFlag) != 0;
        style.IsComb = (fieldFlags & CombFlag) != 0;
        if (TryReadMaxLength(objects, widget, out int maxLength)) {
            style.MaxLength = maxLength;
        } else if (inheritedMaxLength.HasValue) {
            style.MaxLength = inheritedMaxLength.Value;
        }

        if (ResolveDictionary(objects, widget.Items.TryGetValue("MK", out var mkObject) ? mkObject : null) is PdfDictionary mk) {
            if (TryReadColor(objects, mk, "BG", out PdfColor backgroundColor)) {
                style.BackgroundColor = backgroundColor;
            }

            if (TryReadColor(objects, mk, "BC", out PdfColor borderColor)) {
                style.BorderColor = borderColor;
            }
        }

        if (TryReadWidgetBorderWidth(objects, widget, out double borderWidth)) {
            style.BorderWidth = borderWidth;
        }

        if (TryReadWidgetBorderStyle(objects, widget, out PdfFormFieldBorderStyle borderStyle)) {
            style.BorderStyle = borderStyle;
        }

        if (TryReadWidgetBorderDashPattern(objects, widget, out IReadOnlyList<double>? borderDashPattern)) {
            style.BorderDashPattern = borderDashPattern;
        }

        if (TryReadDefaultAppearanceTextColor(objects, widget, inheritedDefaultAppearance, out PdfColor textColor)) {
            style.TextColor = textColor;
        }

        if (TryReadWidgetTextAlignment(objects, widget, inheritedQuadding, out PdfFormFieldTextAlignment textAlignment)) {
            style.TextAlignment = textAlignment;
        }

        return style;
    }

    private static bool TryReadMaxLength(Dictionary<int, PdfIndirectObject> objects, PdfDictionary field, out int maxLength) {
        maxLength = 0;
        if (!field.Items.TryGetValue("MaxLen", out PdfObject? maxLengthObject) ||
            ResolveObject(objects, maxLengthObject) is not PdfNumber maxLengthNumber ||
            maxLengthNumber.Value < 1 ||
            maxLengthNumber.Value > int.MaxValue ||
            Math.Truncate(maxLengthNumber.Value) != maxLengthNumber.Value) {
            return false;
        }

        maxLength = (int)maxLengthNumber.Value;
        return true;
    }

    private static int? ReadFieldQuadding(Dictionary<int, PdfIndirectObject> objects, PdfDictionary field, int? inheritedQuadding) {
        return TryReadQuadding(objects, field, out int quadding) ? quadding : inheritedQuadding;
    }

    private static int? ReadFieldMaxLength(Dictionary<int, PdfIndirectObject> objects, PdfDictionary field, int? inheritedMaxLength) {
        return TryReadMaxLength(objects, field, out int maxLength) ? maxLength : inheritedMaxLength;
    }

    private static bool TryReadWidgetTextAlignment(Dictionary<int, PdfIndirectObject> objects, PdfDictionary widget, int? inheritedQuadding, out PdfFormFieldTextAlignment textAlignment) {
        textAlignment = PdfFormFieldTextAlignment.Unknown;
        int effectiveQuadding;
        if (TryReadQuadding(objects, widget, out int quadding)) {
            effectiveQuadding = quadding;
        } else if (inheritedQuadding.HasValue) {
            effectiveQuadding = inheritedQuadding.Value;
        } else {
            return false;
        }

        switch (effectiveQuadding) {
            case 0:
                textAlignment = PdfFormFieldTextAlignment.Left;
                return true;
            case 1:
                textAlignment = PdfFormFieldTextAlignment.Center;
                return true;
            case 2:
                textAlignment = PdfFormFieldTextAlignment.Right;
                return true;
            default:
                return false;
        }
    }

    private static bool TryReadQuadding(Dictionary<int, PdfIndirectObject> objects, PdfDictionary field, out int quadding) {
        quadding = 0;
        if (!field.Items.TryGetValue("Q", out var quaddingObject) ||
            ResolveObject(objects, quaddingObject) is not PdfNumber quaddingNumber ||
            quaddingNumber.Value < int.MinValue ||
            quaddingNumber.Value > int.MaxValue ||
            Math.Truncate(quaddingNumber.Value) != quaddingNumber.Value) {
            return false;
        }

        quadding = (int)quaddingNumber.Value;
        return true;
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

    private static bool TryReadWidgetBorderWidth(Dictionary<int, PdfIndirectObject> objects, PdfDictionary widget, out double borderWidth) {
        borderWidth = 0D;
        if (ResolveDictionary(objects, widget.Items.TryGetValue("BS", out PdfObject? borderStyleObject) ? borderStyleObject : null) is PdfDictionary borderStyle &&
            borderStyle.Items.TryGetValue("W", out PdfObject? borderStyleWidthObject) &&
            TryReadNonNegativeFiniteNumber(objects, borderStyleWidthObject, out borderWidth)) {
            return true;
        }

        if (widget.Items.TryGetValue("Border", out PdfObject? borderObject) &&
            ResolveObject(objects, borderObject) is PdfArray border &&
            border.Items.Count >= 3 &&
            TryReadNonNegativeFiniteNumber(objects, border.Items[2], out borderWidth)) {
            return true;
        }

        return false;
    }

    private static bool TryReadWidgetBorderStyle(Dictionary<int, PdfIndirectObject> objects, PdfDictionary widget, out PdfFormFieldBorderStyle borderStyle) {
        borderStyle = PdfFormFieldBorderStyle.Solid;
        if (ResolveDictionary(objects, widget.Items.TryGetValue("BS", out PdfObject? borderStyleObject) ? borderStyleObject : null) is not PdfDictionary borderStyleDictionary ||
            borderStyleDictionary.Get<PdfName>("S") is not PdfName styleName) {
            return false;
        }

        switch (styleName.Name) {
            case "D":
                borderStyle = PdfFormFieldBorderStyle.Dashed;
                return true;
            case "U":
                borderStyle = PdfFormFieldBorderStyle.Underline;
                return true;
            case "B":
                borderStyle = PdfFormFieldBorderStyle.Beveled;
                return true;
            case "I":
                borderStyle = PdfFormFieldBorderStyle.Inset;
                return true;
            case "S":
                borderStyle = PdfFormFieldBorderStyle.Solid;
                return true;
            default:
                return false;
        }
    }

    private static bool TryReadWidgetBorderDashPattern(Dictionary<int, PdfIndirectObject> objects, PdfDictionary widget, out IReadOnlyList<double>? borderDashPattern) {
        borderDashPattern = null;
        if (ResolveDictionary(objects, widget.Items.TryGetValue("BS", out PdfObject? borderStyleObject) ? borderStyleObject : null) is not PdfDictionary borderStyle ||
            borderStyle.Get<PdfName>("S")?.Name != "D") {
            return false;
        }

        if (!borderStyle.Items.TryGetValue("D", out PdfObject? dashObject)) {
            borderDashPattern = new[] { 3D };
            return true;
        }

        return TryReadDashPattern(objects, dashObject, out borderDashPattern);
    }

    private static bool TryReadDashPattern(Dictionary<int, PdfIndirectObject> objects, PdfObject dashObject, out IReadOnlyList<double>? dashPattern) {
        dashPattern = null;
        if (ResolveObject(objects, dashObject) is not PdfArray dashArray || dashArray.Items.Count == 0) {
            return false;
        }

        var values = new double[dashArray.Items.Count];
        bool hasPositiveSegment = false;
        for (int i = 0; i < dashArray.Items.Count; i++) {
            if (!TryReadNonNegativeFiniteNumber(objects, dashArray.Items[i], out double segment)) {
                return false;
            }

            if (segment > 0D) {
                hasPositiveSegment = true;
            }

            values[i] = segment;
        }

        if (!hasPositiveSegment) {
            return false;
        }

        dashPattern = values;
        return true;
    }

    private static bool TryReadNonNegativeFiniteNumber(Dictionary<int, PdfIndirectObject> objects, PdfObject numberObject, out double value) {
        value = 0D;
        if (ResolveObject(objects, numberObject) is not PdfNumber number ||
            number.Value < 0D ||
            double.IsNaN(number.Value) ||
            double.IsInfinity(number.Value)) {
            return false;
        }

        value = number.Value;
        return true;
    }

    private static bool TryReadDefaultAppearanceTextColor(Dictionary<int, PdfIndirectObject> objects, PdfDictionary widget, string? inheritedDefaultAppearance, out PdfColor color) {
        return PdfDefaultAppearanceParser.TryReadTextColor(TryReadText(objects, widget, "DA") ?? inheritedDefaultAppearance, out color);
    }

    private static double ReadWidgetAppearanceFontSize(string? defaultAppearance, double height) {
        return PdfDefaultAppearanceParser.TryReadFontSize(defaultAppearance, out double fontSize)
            ? fontSize
            : Math.Max(6D, Math.Min(12D, height - 4D));
    }

    private static bool TryCreateDefaultAppearanceSimpleFontResources(Dictionary<int, PdfIndirectObject> objects, string? defaultAppearance, PdfDictionary? inheritedDefaultResources, PdfDictionary? widgetAppearanceResources, PdfDictionary? widgetPageResources, out string? fontResourceName, out PdfDictionary? resources) {
        fontResourceName = null;
        resources = null;
        if (!PdfDefaultAppearanceParser.TryReadFontResourceName(defaultAppearance, out string defaultAppearanceFontName)) {
            return false;
        }

        foreach (PdfDictionary candidateResources in EnumerateCandidateTextAppearanceResources(inheritedDefaultResources, widgetAppearanceResources, widgetPageResources)) {
            if (ResolveDictionary(objects, candidateResources.Items.TryGetValue("Font", out PdfObject? fontsObject) ? fontsObject : null) is not PdfDictionary fonts ||
                !fonts.Items.TryGetValue(defaultAppearanceFontName, out PdfObject? fontObject)) {
                continue;
            }

            var appearanceFonts = new PdfDictionary();
            appearanceFonts.Items[defaultAppearanceFontName] = fontObject;
            resources = new PdfDictionary();
            resources.Items["Font"] = appearanceFonts;
            fontResourceName = defaultAppearanceFontName;
            return true;
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

    private static PdfDictionary CreateRichTextAppearanceResources(IReadOnlyList<(string Name, PdfStandardFont Font)> fontResources) {
        var fonts = new PdfDictionary();
        for (int i = 0; i < fontResources.Count; i++) {
            (string name, PdfStandardFont font) = fontResources[i];
            fonts.Items[name] = PdfStandardFontDictionaryBuilder.BuildStandardType1FontDictionary(font);
        }

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
