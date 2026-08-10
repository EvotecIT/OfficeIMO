namespace OfficeIMO.Pdf;

internal static partial class PdfAcroFormEditor {
    private const int FieldFlagReadOnly = 1;
    private const int FieldFlagRequired = 2;
    private const int FieldFlagNoExport = 4;
    private const int FieldFlagMultiline = 4096;
    private const int FieldFlagPassword = 8192;
    private const int FieldFlagNoToggleToOff = 16384;
    private const int FieldFlagRadio = 32768;
    private const int FieldFlagPushButton = 65536;
    private const int FieldFlagCombo = 131072;
    private const int FieldFlagEdit = 262144;
    private const int FieldFlagSort = 524288;
    private const int FieldFlagFileSelect = 1048576;
    private const int FieldFlagDoNotSpellCheck = 4194304;
    private const int FieldFlagDoNotScroll = 8388608;
    private const int FieldFlagComb = 16777216;
    private const int FieldFlagCommitOnSelectionChange = 67108864;

    private static void EnsureAcroFormAppearanceDefaults(Dictionary<int, PdfIndirectObject> objects, PdfDictionary acroForm) {
        if (!acroForm.Items.ContainsKey("DA")) acroForm.Items["DA"] = new PdfStringObj("/Helv 10 Tf 0 g", true);
        PdfDictionary resources = acroForm.Items.TryGetValue("DR", out PdfObject? resourcesObject) && ResolveDictionary(objects, resourcesObject) is PdfDictionary existingResources
            ? existingResources
            : new PdfDictionary();
        PdfDictionary fonts = resources.Items.TryGetValue("Font", out PdfObject? fontsObject) && ResolveDictionary(objects, fontsObject) is PdfDictionary existingFonts
            ? existingFonts
            : new PdfDictionary();
        if (!fonts.Items.ContainsKey("Helv")) {
            var helvetica = new PdfDictionary();
            helvetica.Items["Type"] = new PdfName("Font");
            helvetica.Items["Subtype"] = new PdfName("Type1");
            helvetica.Items["BaseFont"] = new PdfName("Helvetica");
            fonts.Items["Helv"] = helvetica;
        }
        resources.Items["Font"] = fonts;
        acroForm.Items["DR"] = resources;
    }

    private static void ApplyCreateRadioButtonGroup(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary acroForm,
        PdfArray fields,
        int[] pages,
        PdfFormFieldCreateOptions options,
        Dictionary<string, string> refillValues,
        ref int nextObjectNumber) {
        PdfDictionary page = RequirePage(objects, pages, options.PageNumber);
        int parentObjectNumber = nextObjectNumber++;
        string selectedValue = ResolveInitialValue(options);
        var parent = new PdfDictionary();
        parent.Items["FT"] = new PdfName("Btn");
        parent.Items["T"] = new PdfStringObj(options.Name, true);
        parent.Items["Ff"] = new PdfNumber(GetCreateFieldFlags(options));
        parent.Items["V"] = new PdfName(selectedValue);
        parent.Items["DV"] = new PdfName(options.DefaultValue ?? selectedValue);
        ApplyCreateFieldStyle(parent, options, includeWidgetStyle: false);

        var kids = new PdfArray();
        parent.Items["Kids"] = kids;
        objects[parentObjectNumber] = new PdfIndirectObject(parentObjectNumber, 0, parent);
        fields.Items.Add(new PdfReference(parentObjectNumber, 0));

        PdfArray annotations = EnsureAnnotationArray(objects, page);
        double top = options.Y + options.Height;
        for (int i = 0; i < options.ChoiceOptions.Count; i++) {
            string option = options.ChoiceOptions[i];
            double widgetTop = top - i * (options.RadioButtonSize + options.RadioButtonGap);
            double widgetBottom = widgetTop - options.RadioButtonSize;
            int widgetObjectNumber = nextObjectNumber++;
            var widget = new PdfDictionary();
            widget.Items["Type"] = new PdfName("Annot");
            widget.Items["Subtype"] = new PdfName("Widget");
            widget.Items["Parent"] = new PdfReference(parentObjectNumber, 0);
            widget.Items["Rect"] = CreateRectangle(options.X, widgetBottom, options.X + options.Width, widgetTop);
            widget.Items["P"] = CreateReference(objects, pages[options.PageNumber - 1]);
            widget.Items["F"] = new PdfNumber(options.WidgetFlags);
            widget.Items["AS"] = new PdfName(string.Equals(option, selectedValue, StringComparison.Ordinal) ? option : "Off");
            PdfFormFieldStyle style = options.Style ?? new PdfFormFieldStyle();
            ApplyWidgetVisualStyle(widget, style, option);
            int offAppearanceObjectNumber = nextObjectNumber++;
            objects[offAppearanceObjectNumber] = new PdfIndirectObject(
                offAppearanceObjectNumber,
                0,
                PdfFormFiller.CreateAuthoredLabeledRadioWidgetAppearance(objects, acroForm, page, widget, option, options.Width, options.RadioButtonSize, style, options.FontSize, options.Name, selected: false, ref nextObjectNumber));
            int selectedAppearanceObjectNumber = nextObjectNumber++;
            objects[selectedAppearanceObjectNumber] = new PdfIndirectObject(
                selectedAppearanceObjectNumber,
                0,
                PdfFormFiller.CreateAuthoredLabeledRadioWidgetAppearance(objects, acroForm, page, widget, option, options.Width, options.RadioButtonSize, style, options.FontSize, options.Name, selected: true, ref nextObjectNumber));
            var normalAppearances = new PdfDictionary();
            normalAppearances.Items["Off"] = new PdfReference(offAppearanceObjectNumber, 0);
            normalAppearances.Items[option] = new PdfReference(selectedAppearanceObjectNumber, 0);
            var appearances = new PdfDictionary();
            appearances.Items["N"] = normalAppearances;
            widget.Items["AP"] = appearances;
            ApplyWidgetJavaScript(widget, options.JavaScript, usePrimaryAction: false);
            objects[widgetObjectNumber] = new PdfIndirectObject(widgetObjectNumber, 0, widget);
            var widgetReference = new PdfReference(widgetObjectNumber, 0);
            kids.Items.Add(widgetReference);
            annotations.Items.Add(widgetReference);
        }

        refillValues[options.Name] = selectedValue;
    }

    private static void AddPushButtonAppearance(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary acroForm,
        PdfDictionary page,
        PdfDictionary widget,
        PdfFormFieldCreateOptions options,
        ref int nextObjectNumber) {
        PdfFormFieldStyle style = options.Style?.Clone() ?? new PdfFormFieldStyle();
        style.TextAlignment = PdfFormFieldTextAlignment.Center;
        PdfStream appearance = PdfFormFiller.CreateAuthoredTextWidgetAppearance(
            objects,
            acroForm,
            page,
            widget,
            options.Caption,
            options.Width,
            options.Height,
            style,
            options.FontSize,
            options.Name,
            ref nextObjectNumber);
        int appearanceObjectNumber = nextObjectNumber++;
        objects[appearanceObjectNumber] = new PdfIndirectObject(appearanceObjectNumber, 0, appearance);
        var appearances = new PdfDictionary();
        appearances.Items["N"] = new PdfReference(appearanceObjectNumber, 0);
        widget.Items["AP"] = appearances;
    }

    private static void ApplyCreateFieldStyle(PdfDictionary field, PdfFormFieldCreateOptions options, bool includeWidgetStyle) {
        PdfFormFieldStyle style = options.Style ?? new PdfFormFieldStyle();
        if (!string.IsNullOrWhiteSpace(style.AlternateName)) field.Items["TU"] = new PdfStringObj(style.AlternateName!, true);
        if (!string.IsNullOrWhiteSpace(style.MappingName)) field.Items["TM"] = new PdfStringObj(style.MappingName!, true);
        if (options.Kind == PdfFormFieldCreationKind.Text && style.MaxLength.HasValue) field.Items["MaxLen"] = new PdfNumber(style.MaxLength.Value);
        if ((options.Kind == PdfFormFieldCreationKind.Text || options.Kind == PdfFormFieldCreationKind.Choice) && style.TextAlignment.HasValue) {
            field.Items["Q"] = new PdfNumber(GetQuadding(style.TextAlignment.Value));
        }
        if (options.Kind == PdfFormFieldCreationKind.Text || options.Kind == PdfFormFieldCreationKind.Choice) {
            field.Items["DA"] = new PdfStringObj(BuildDefaultAppearance(options.FontSize, style.TextColor), true);
        }
        if (includeWidgetStyle) ApplyWidgetVisualStyle(field, style, options.Kind == PdfFormFieldCreationKind.PushButton ? options.Caption : null);
    }

    private static void ApplyWidgetVisualStyle(PdfDictionary widget, PdfFormFieldStyle style, string? caption) {
        var characteristics = new PdfDictionary();
        if (style.BackgroundColor.HasValue) characteristics.Items["BG"] = CreateColorArray(style.BackgroundColor.Value);
        if (style.BorderColor.HasValue) characteristics.Items["BC"] = CreateColorArray(style.BorderColor.Value);
        if (!string.IsNullOrEmpty(caption)) characteristics.Items["CA"] = new PdfStringObj(caption!, true);
        if (characteristics.Items.Count > 0) widget.Items["MK"] = characteristics;

        var border = new PdfDictionary();
        border.Items["W"] = new PdfNumber(style.BorderWidth);
        border.Items["S"] = new PdfName(GetBorderStyleName(style.BorderStyle));
        if (style.BorderDashPattern is not null && style.BorderDashPattern.Count > 0) {
            var dash = new PdfArray();
            for (int i = 0; i < style.BorderDashPattern.Count; i++) dash.Items.Add(new PdfNumber(style.BorderDashPattern[i]));
            border.Items["D"] = dash;
        }
        widget.Items["BS"] = border;
    }

    private static void ApplyWidgetJavaScript(PdfDictionary widget, string? javaScript, bool usePrimaryAction) {
        if (javaScript is null) return;
        byte[] encodedSource = PdfJavaScriptStringEncoding.EncodeUnicode(javaScript, nameof(javaScript));
        var action = new PdfDictionary();
        action.Items["S"] = new PdfName("JavaScript");
        action.Items["JS"] = new PdfStringObj(encodedSource, useTextStringEncoding: true);
        if (usePrimaryAction) {
            widget.Items["A"] = action;
            return;
        }

        PdfDictionary additional = widget.Items.TryGetValue("AA", out PdfObject? value) && value is PdfDictionary existing
            ? existing
            : new PdfDictionary();
        additional.Items["U"] = action;
        widget.Items["AA"] = additional;
    }

    private static int GetCreateFieldFlags(PdfFormFieldCreateOptions options) {
        PdfFormFieldStyle style = options.Style ?? new PdfFormFieldStyle();
        int flags = options.FieldFlags;
        if (style.IsReadOnly) flags |= FieldFlagReadOnly;
        if (style.IsRequired) flags |= FieldFlagRequired;
        if (style.IsNoExport) flags |= FieldFlagNoExport;
        if (options.Kind == PdfFormFieldCreationKind.Text) {
            if (style.IsMultiline) flags |= FieldFlagMultiline;
            if (style.IsPassword) flags |= FieldFlagPassword;
            if (style.IsFileSelect) flags |= FieldFlagFileSelect;
            if (style.DoesNotSpellCheck) flags |= FieldFlagDoNotSpellCheck;
            if (style.DoesNotScroll) flags |= FieldFlagDoNotScroll;
            if (style.IsComb) flags |= FieldFlagComb;
        } else if (options.Kind == PdfFormFieldCreationKind.Choice) {
            if (options.IsComboBox) flags |= FieldFlagCombo;
            if (style.IsEditableChoice) flags |= FieldFlagEdit;
            if (style.IsSortedChoice) flags |= FieldFlagSort;
            if (style.DoesNotSpellCheck) flags |= FieldFlagDoNotSpellCheck;
            if (style.CommitsOnSelectionChange) flags |= FieldFlagCommitOnSelectionChange;
        } else if (options.Kind == PdfFormFieldCreationKind.RadioButtonGroup) {
            flags |= FieldFlagNoToggleToOff | FieldFlagRadio;
        } else if (options.Kind == PdfFormFieldCreationKind.PushButton) {
            flags |= FieldFlagPushButton;
        }
        return flags;
    }

    private static string ResolveInitialValue(PdfFormFieldCreateOptions options) {
        if ((options.Kind == PdfFormFieldCreationKind.Choice || options.Kind == PdfFormFieldCreationKind.RadioButtonGroup) && string.IsNullOrEmpty(options.Value)) {
            return options.ChoiceOptions[0];
        }
        return options.Value;
    }

    private static PdfArray CreateColorArray(PdfColor color) => CreateNumberArray(color.R, color.G, color.B);

    private static PdfArray CreateNumberArray(params double[] values) {
        var array = new PdfArray();
        for (int i = 0; i < values.Length; i++) array.Items.Add(new PdfNumber(values[i]));
        return array;
    }

    private static string BuildDefaultAppearance(double fontSize, PdfColor textColor) =>
        "/Helv " + fontSize.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture) + " Tf " +
        textColor.R.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture) + " " +
        textColor.G.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture) + " " +
        textColor.B.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture) + " rg";

    private static string GetBorderStyleName(PdfFormFieldBorderStyle style) => style switch {
        PdfFormFieldBorderStyle.Dashed => "D",
        PdfFormFieldBorderStyle.Beveled => "B",
        PdfFormFieldBorderStyle.Inset => "I",
        PdfFormFieldBorderStyle.Underline => "U",
        _ => "S"
    };

    private static int GetQuadding(PdfFormFieldTextAlignment alignment) => alignment switch {
        PdfFormFieldTextAlignment.Left => 0,
        PdfFormFieldTextAlignment.Center => 1,
        PdfFormFieldTextAlignment.Right => 2,
        _ => throw new ArgumentOutOfRangeException(nameof(alignment), alignment, "Form field alignment must be Left, Center, or Right.")
    };
}
