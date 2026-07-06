using System.Globalization;

namespace OfficeIMO.Pdf;

internal static partial class PdfAnnotationDictionaryBuilder {
    private const int FieldFlagReadOnly = 1;
    private const int FieldFlagRequired = 2;
    private const int FieldFlagNoExport = 4;
    private const int FieldFlagMultiline = 4096;
    private const int FieldFlagPassword = 8192;
    private const int FieldFlagCombo = 131072;
    private const int FieldFlagEdit = 262144;
    private const int FieldFlagSort = 524288;
    private const int FieldFlagFileSelect = 1048576;
    private const int FieldFlagDoNotSpellCheck = 4194304;
    private const int FieldFlagDoNotScroll = 8388608;
    private const int FieldFlagComb = 16777216;
    private const int FieldFlagCommitOnSelectionChange = 67108864;

    internal static string BuildUriLinkAnnotation(double x1, double y1, double x2, double y2, string uri, string? contents = null, int? structParentIndex = null) {
        ValidateRectangle(x1, y1, x2, y2);
        Guard.UriAction(uri, nameof(uri));

        return "<< /Type /Annot /Subtype /Link /Border [0 0 0]" + BuildContentsEntry(contents) + " /Rect [" +
            FormatCoordinate(x1) + " " +
            FormatCoordinate(y1) + " " +
            FormatCoordinate(x2) + " " +
            FormatCoordinate(y2) +
            "] /A << /S /URI /URI " +
            PdfSyntaxEscaper.LiteralString(uri) +
            " >>" +
            BuildStructParentEntry(structParentIndex) +
            " >>\n";
    }

    internal static string BuildGoToNamedDestinationLinkAnnotation(double x1, double y1, double x2, double y2, string destinationName, string? contents = null, int? structParentIndex = null) {
        ValidateRectangle(x1, y1, x2, y2);
        Guard.NotNullOrWhiteSpace(destinationName, nameof(destinationName));

        return "<< /Type /Annot /Subtype /Link /Border [0 0 0]" + BuildContentsEntry(contents) + " /Rect [" +
            FormatCoordinate(x1) + " " +
            FormatCoordinate(y1) + " " +
            FormatCoordinate(x2) + " " +
            FormatCoordinate(y2) +
            "] /A << /S /GoTo /D " +
            PdfSyntaxEscaper.LiteralString(destinationName) +
            " >>" +
            BuildStructParentEntry(structParentIndex) +
            " >>\n";
    }

    internal static string BuildAppearanceStreamDictionary(double width, double height, int contentLength, int helveticaFontId = 0, bool usesHighlightBlendMode = false) {
        IReadOnlyList<(string Name, int Id)> fontResources = helveticaFontId > 0
            ? new[] { ("Helv", helveticaFontId) }
            : Array.Empty<(string Name, int Id)>();

        return BuildAppearanceStreamDictionary(width, height, contentLength, fontResources, usesHighlightBlendMode);
    }

    internal static string BuildAppearanceStreamDictionary(double width, double height, int contentLength, IReadOnlyList<(string Name, int Id)> fontResources, bool usesHighlightBlendMode = false) {
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        Guard.NotNull(fontResources, nameof(fontResources));
        if (contentLength < 0) {
            throw new ArgumentOutOfRangeException(nameof(contentLength), "PDF annotation appearance stream length cannot be negative.");
        }

        string resources = BuildAppearanceStreamResources(fontResources, usesHighlightBlendMode);
        return "<< /Type /XObject /Subtype /Form /BBox [0 0 " +
            FormatCoordinate(width) +
            " " +
            FormatCoordinate(height) +
            "]" +
            resources +
            " /Length " +
            contentLength.ToString(CultureInfo.InvariantCulture) +
            " >>";
    }

    private static string BuildAppearanceStreamResources(IReadOnlyList<(string Name, int Id)> fontResources, bool usesHighlightBlendMode) {
        if (fontResources.Count == 0 && !usesHighlightBlendMode) {
            return string.Empty;
        }

        var sb = new StringBuilder();
        sb.Append(" /Resources <<");
        if (fontResources.Count > 0) {
            sb.Append(" /Font <<");
            for (int i = 0; i < fontResources.Count; i++) {
                (string name, int id) = fontResources[i];
                Guard.NotNullOrWhiteSpace(name, nameof(fontResources));
                if (id <= 0) {
                    throw new ArgumentOutOfRangeException(nameof(fontResources), id, "PDF annotation appearance font resource object id must be positive.");
                }

                string normalizedName = name[0] == '/' ? name.Substring(1) : name;
                Guard.NotNullOrWhiteSpace(normalizedName, nameof(fontResources));
                sb.Append(" /")
                    .Append(PdfSyntaxEscaper.Name(normalizedName))
                    .Append(' ')
                    .Append(PdfSyntaxEscaper.IndirectReference(id));
            }

            sb.Append(" >>");
        }

        if (usesHighlightBlendMode) {
            sb.Append(" /ExtGState << /OfficeIMOHighlightGs << /Type /ExtGState /BM /Multiply /CA 0.35 /ca 0.35 >> >>");
        }

        sb.Append(" >>");
        return sb.ToString();
    }

    internal static string BuildTextFieldWidgetAnnotation(double x1, double y1, double x2, double y2, string name, string value, double fontSize, int normalAppearanceId, PdfFormFieldStyle? style = null, int? structParentIndex = null) {
        ValidateRectangle(x1, y1, x2, y2);
        Guard.NotNullOrWhiteSpace(name, nameof(name));
        Guard.NotNull(value, nameof(value));
        ValidateFinite(fontSize, nameof(fontSize));
        if (fontSize <= 0) {
            throw new ArgumentOutOfRangeException(nameof(fontSize), fontSize, "PDF text field font size must be a positive finite number.");
        }

        return "<< /Type /Annot /Subtype /Widget /FT /Tx /T " +
            PdfSyntaxEscaper.TextString(name) +
            BuildFormFieldMetadataEntries(style) +
            BuildTextFieldFlagsEntry(style) +
            BuildMaxLengthEntry(style) +
            " /V " +
            PdfSyntaxEscaper.TextString(value) +
            " /DV " +
            PdfSyntaxEscaper.TextString(value) +
            " /Rect [" +
            FormatCoordinate(x1) + " " +
            FormatCoordinate(y1) + " " +
            FormatCoordinate(x2) + " " +
            FormatCoordinate(y2) +
            "] /F 4 /DA " +
            PdfSyntaxEscaper.LiteralString("/Helv " + FormatCoordinate(fontSize) + " Tf " + PdfAcroFormDictionaryBuilder.FormatColor((style ?? new PdfFormFieldStyle()).TextColor) + " rg") +
            BuildQuaddingEntry(style) +
            BuildMkEntry(style) +
            BuildBorderStyleEntry(style) +
            " /AP << /N " +
            PdfSyntaxEscaper.IndirectReference(normalAppearanceId) +
            " >>" +
            BuildStructParentEntry(structParentIndex) +
            " >>\n";
    }

    internal static string BuildCheckBoxWidgetAnnotation(double x1, double y1, double x2, double y2, string name, bool isChecked, string checkedValueName, int offAppearanceId, int checkedAppearanceId, PdfFormFieldStyle? style = null, int? structParentIndex = null) {
        ValidateRectangle(x1, y1, x2, y2);
        Guard.NotNullOrWhiteSpace(name, nameof(name));
        Guard.NotNullOrWhiteSpace(checkedValueName, nameof(checkedValueName));
        if (string.Equals(checkedValueName, "Off", StringComparison.Ordinal)) {
            throw new ArgumentException("PDF check box selected value name cannot be Off.", nameof(checkedValueName));
        }

        ValidateAsciiPdfNameValue(checkedValueName, nameof(checkedValueName));

        string selectedName = isChecked ? checkedValueName : "Off";
        return "<< /Type /Annot /Subtype /Widget /FT /Btn /T " +
            PdfSyntaxEscaper.TextString(name) +
            BuildFormFieldMetadataEntries(style) +
            BuildFieldFlagsEntry(style) +
            " /V /" +
            PdfSyntaxEscaper.Name(selectedName) +
            " /DV /" +
            PdfSyntaxEscaper.Name(selectedName) +
            " /Rect [" +
            FormatCoordinate(x1) + " " +
            FormatCoordinate(y1) + " " +
            FormatCoordinate(x2) + " " +
            FormatCoordinate(y2) +
            "] /F 4 /AS /" +
            PdfSyntaxEscaper.Name(selectedName) +
            BuildMkEntry(style) +
            BuildBorderStyleEntry(style) +
            " /AP << /N << /Off " +
            PdfSyntaxEscaper.IndirectReference(offAppearanceId) +
            " /" +
            PdfSyntaxEscaper.Name(checkedValueName) +
            " " +
            PdfSyntaxEscaper.IndirectReference(checkedAppearanceId) +
            " >> >>" +
            BuildStructParentEntry(structParentIndex) +
            " >>\n";
    }

    internal static string BuildChoiceFieldWidgetAnnotation(double x1, double y1, double x2, double y2, string name, IReadOnlyList<string> options, string value, double fontSize, int normalAppearanceId, bool isComboBox, PdfFormFieldStyle? style = null) =>
        BuildChoiceFieldWidgetAnnotation(x1, y1, x2, y2, name, options, new[] { value }, fontSize, normalAppearanceId, isComboBox, allowsMultipleSelection: false, style);

    internal static string BuildChoiceFieldWidgetAnnotation(double x1, double y1, double x2, double y2, string name, IReadOnlyList<string> options, IReadOnlyList<string> values, double fontSize, int normalAppearanceId, bool isComboBox, bool allowsMultipleSelection, PdfFormFieldStyle? style = null, int? structParentIndex = null) {
        ValidateRectangle(x1, y1, x2, y2);
        Guard.NotNullOrWhiteSpace(name, nameof(name));
        Guard.NotNull(options, nameof(options));
        Guard.NotNull(values, nameof(values));
        ValidateFinite(fontSize, nameof(fontSize));
        if (fontSize <= 0) {
            throw new ArgumentOutOfRangeException(nameof(fontSize), fontSize, "PDF choice field font size must be a positive finite number.");
        }

        if (options.Count == 0) {
            throw new ArgumentException("PDF choice field requires at least one option.", nameof(options));
        }

        if (values.Count == 0) {
            throw new ArgumentException("PDF choice field requires at least one selected value.", nameof(values));
        }

        if (!allowsMultipleSelection && values.Count > 1) {
            throw new ArgumentException("PDF scalar choice field cannot contain multiple selected values.", nameof(values));
        }

        if (allowsMultipleSelection && isComboBox) {
            throw new ArgumentException("PDF multi-select choice fields must be list boxes, not combo boxes.", nameof(isComboBox));
        }

        var optionBuilder = new StringBuilder();
        var optionSet = new HashSet<string>(StringComparer.Ordinal);
        for (int i = 0; i < options.Count; i++) {
            string option = options[i];
            Guard.NotNullOrWhiteSpace(option, nameof(options));
            if (!optionSet.Add(option)) {
                throw new ArgumentException("PDF choice field options must be unique.", nameof(options));
            }

            optionBuilder.Append(' ')
                .Append(PdfSyntaxEscaper.TextString(option));
        }

        bool allowsCustomScalarValue = isComboBox && !allowsMultipleSelection && style != null && style.IsEditableChoice;
        var valueSet = new HashSet<string>(StringComparer.Ordinal);
        for (int i = 0; i < values.Count; i++) {
            string value = values[i];
            Guard.NotNullOrWhiteSpace(value, nameof(values));
            if (!allowsCustomScalarValue && !optionSet.Contains(value)) {
                throw new ArgumentException("PDF choice field values must match the provided options.", nameof(values));
            }

            if (!valueSet.Add(value)) {
                throw new ArgumentException("PDF choice field selected values must be unique.", nameof(values));
            }
        }

        int flags = BuildChoiceFieldFlags(style, (isComboBox ? FieldFlagCombo : 0) | (allowsMultipleSelection ? 2097152 : 0), isComboBox);
        return "<< /Type /Annot /Subtype /Widget /FT /Ch /T " +
            PdfSyntaxEscaper.TextString(name) +
            BuildFormFieldMetadataEntries(style) +
            " /V " +
            BuildChoiceValue(values, allowsMultipleSelection) +
            " /DV " +
            BuildChoiceValue(values, allowsMultipleSelection) +
            " /Opt [" +
            optionBuilder +
            " ]" +
            (flags == 0 ? string.Empty : " /Ff " + flags.ToString(CultureInfo.InvariantCulture)) +
            " /Rect [" +
            FormatCoordinate(x1) + " " +
            FormatCoordinate(y1) + " " +
            FormatCoordinate(x2) + " " +
            FormatCoordinate(y2) +
            "] /F 4 /DA " +
            PdfSyntaxEscaper.LiteralString("/Helv " + FormatCoordinate(fontSize) + " Tf " + PdfAcroFormDictionaryBuilder.FormatColor((style ?? new PdfFormFieldStyle()).TextColor) + " rg") +
            BuildQuaddingEntry(style) +
            BuildMkEntry(style) +
            BuildBorderStyleEntry(style) +
            " /AP << /N " +
            PdfSyntaxEscaper.IndirectReference(normalAppearanceId) +
            " >>" +
            BuildStructParentEntry(structParentIndex) +
            " >>\n";
    }

    internal static string BuildRadioButtonFieldDictionary(string name, IReadOnlyList<string> options, string value, IReadOnlyList<int> widgetObjectIds, PdfFormFieldStyle? style = null) {
        Guard.NotNullOrWhiteSpace(name, nameof(name));
        Guard.NotNull(options, nameof(options));
        Guard.NotNullOrWhiteSpace(value, nameof(value));
        Guard.NotNull(widgetObjectIds, nameof(widgetObjectIds));
        ValidateRadioOptions(options, value);
        if (widgetObjectIds.Count != options.Count) {
            throw new ArgumentException("PDF radio button group requires one widget object per option.", nameof(widgetObjectIds));
        }

        var sb = new StringBuilder();
        sb.Append("<< /FT /Btn /T ")
            .Append(PdfSyntaxEscaper.TextString(name))
            .Append(BuildFormFieldMetadataEntries(style))
            .Append(BuildFieldFlagsEntry(style, 49152))
            .Append(" /V /")
            .Append(PdfSyntaxEscaper.Name(value))
            .Append(" /DV /")
            .Append(PdfSyntaxEscaper.Name(value))
            .Append(" /Kids [");
        for (int i = 0; i < widgetObjectIds.Count; i++) {
            sb.Append(' ')
                .Append(PdfSyntaxEscaper.IndirectReference(widgetObjectIds[i]));
        }

        sb.Append(" ] >>\n");
        return sb.ToString();
    }

    internal static string BuildRadioButtonWidgetAnnotation(double x1, double y1, double x2, double y2, int parentObjectId, string option, string value, int offAppearanceId, int selectedAppearanceId, PdfFormFieldStyle? style = null, int? structParentIndex = null) {
        ValidateRectangle(x1, y1, x2, y2);
        if (parentObjectId <= 0) {
            throw new ArgumentOutOfRangeException(nameof(parentObjectId), parentObjectId, "PDF radio button parent object id must be positive.");
        }

        Guard.NotNullOrWhiteSpace(option, nameof(option));
        Guard.NotNullOrWhiteSpace(value, nameof(value));
        if (string.Equals(option, "Off", StringComparison.Ordinal)) {
            throw new ArgumentException("PDF radio button option value cannot be Off.", nameof(option));
        }

        ValidateAsciiPdfNameValue(option, nameof(option), "PDF radio button option values must contain only ASCII PDF name characters.");
        string stateName = string.Equals(option, value, StringComparison.Ordinal) ? option : "Off";
        return "<< /Type /Annot /Subtype /Widget /Parent " +
            PdfSyntaxEscaper.IndirectReference(parentObjectId) +
            " /Rect [" +
            FormatCoordinate(x1) + " " +
            FormatCoordinate(y1) + " " +
            FormatCoordinate(x2) + " " +
            FormatCoordinate(y2) +
            "] /F 4 /AS /" +
            PdfSyntaxEscaper.Name(stateName) +
            BuildMkEntry(style) +
            BuildBorderStyleEntry(style) +
            " /AP << /N << /Off " +
            PdfSyntaxEscaper.IndirectReference(offAppearanceId) +
            " /" +
            PdfSyntaxEscaper.Name(option) +
            " " +
            PdfSyntaxEscaper.IndirectReference(selectedAppearanceId) +
            " >> >>" +
            BuildStructParentEntry(structParentIndex) +
            " >>\n";
    }

    private static string BuildChoiceValue(IReadOnlyList<string> values, bool forceArray) {
        if (values.Count == 1 && !forceArray) {
            return PdfSyntaxEscaper.TextString(values[0]);
        }

        var valueBuilder = new StringBuilder();
        valueBuilder.Append('[');
        for (int i = 0; i < values.Count; i++) {
            if (i > 0) {
                valueBuilder.Append(' ');
            }

            valueBuilder.Append(PdfSyntaxEscaper.TextString(values[i]));
        }

        valueBuilder.Append(']');
        return valueBuilder.ToString();
    }

    private static string BuildContentsEntry(string? contents) =>
        string.IsNullOrWhiteSpace(contents)
            ? string.Empty
            : " /Contents " + PdfSyntaxEscaper.LiteralString(contents!);

    private static string BuildFormFieldMetadataEntries(PdfFormFieldStyle? style) {
        if (style == null) {
            return string.Empty;
        }

        var sb = new StringBuilder();
        if (!string.IsNullOrWhiteSpace(style.AlternateName)) {
            sb.Append(" /TU ")
                .Append(PdfSyntaxEscaper.TextString(style.AlternateName!));
        }

        if (!string.IsNullOrWhiteSpace(style.MappingName)) {
            sb.Append(" /TM ")
                .Append(PdfSyntaxEscaper.TextString(style.MappingName!));
        }

        return sb.ToString();
    }

    private static string BuildQuaddingEntry(PdfFormFieldStyle? style) {
        if (style == null || !style.TextAlignment.HasValue) {
            return string.Empty;
        }

        return " /Q " + PdfAcroFormDictionaryBuilder.ToQuadding(style.TextAlignment.Value).ToString(CultureInfo.InvariantCulture);
    }

    private static string BuildFieldFlagsEntry(PdfFormFieldStyle? style, int baseFlags = 0) {
        int flags = BuildFieldFlags(style, baseFlags);
        return flags == 0 ? string.Empty : " /Ff " + flags.ToString(CultureInfo.InvariantCulture);
    }

    private static string BuildTextFieldFlagsEntry(PdfFormFieldStyle? style) {
        int flags = BuildFieldFlags(style);
        if (style != null) {
            ValidateCombTextFieldStyle(style);

            if (style.IsMultiline) {
                flags |= FieldFlagMultiline;
            }

            if (style.IsPassword) {
                flags |= FieldFlagPassword;
            }

            if (style.IsFileSelect) {
                flags |= FieldFlagFileSelect;
            }

            if (style.DoesNotSpellCheck) {
                flags |= FieldFlagDoNotSpellCheck;
            }

            if (style.DoesNotScroll) {
                flags |= FieldFlagDoNotScroll;
            }

            if (style.IsComb) {
                flags |= FieldFlagComb;
            }
        }

        return flags == 0 ? string.Empty : " /Ff " + flags.ToString(CultureInfo.InvariantCulture);
    }

    private static void ValidateCombTextFieldStyle(PdfFormFieldStyle style) {
        if (!style.IsComb) {
            return;
        }

        if (!style.MaxLength.HasValue || style.IsMultiline || style.IsPassword || style.IsFileSelect) {
            throw new ArgumentException("PDF comb text fields require MaxLength and cannot also be multiline, password, or file-select fields.", nameof(style));
        }
    }

    private static int BuildChoiceFieldFlags(PdfFormFieldStyle? style, int baseFlags, bool isComboBox) {
        int flags = BuildFieldFlags(style, baseFlags);
        if (style != null && style.DoesNotSpellCheck) {
            flags |= FieldFlagDoNotSpellCheck;
        }

        if (style != null && style.IsEditableChoice && isComboBox) {
            flags |= FieldFlagEdit;
        }

        if (style != null && style.IsSortedChoice) {
            flags |= FieldFlagSort;
        }

        if (style != null && style.CommitsOnSelectionChange) {
            flags |= FieldFlagCommitOnSelectionChange;
        }

        return flags;
    }

    private static string BuildMaxLengthEntry(PdfFormFieldStyle? style) {
        if (style == null || !style.MaxLength.HasValue) {
            return string.Empty;
        }

        return " /MaxLen " + style.MaxLength.Value.ToString(CultureInfo.InvariantCulture);
    }

    private static int BuildFieldFlags(PdfFormFieldStyle? style, int baseFlags = 0) {
        int flags = baseFlags;
        if (style == null) {
            return flags;
        }

        if (style.IsReadOnly) {
            flags |= FieldFlagReadOnly;
        }

        if (style.IsRequired) {
            flags |= FieldFlagRequired;
        }

        if (style.IsNoExport) {
            flags |= FieldFlagNoExport;
        }

        return flags;
    }

    private static string BuildStructParentEntry(int? structParentIndex) {
        if (!structParentIndex.HasValue) {
            return string.Empty;
        }

        if (structParentIndex.Value < 0) {
            throw new ArgumentOutOfRangeException(nameof(structParentIndex), structParentIndex.Value, "PDF annotation StructParent index must be non-negative.");
        }

        return " /StructParent " + structParentIndex.Value.ToString(CultureInfo.InvariantCulture);
    }

    private static string BuildMkEntry(PdfFormFieldStyle? style) {
        PdfFormFieldStyle effectiveStyle = style ?? new PdfFormFieldStyle();
        var sb = new StringBuilder();
        if (effectiveStyle.BorderColor.HasValue && effectiveStyle.BorderWidth > 0) {
            sb.Append(" /BC [").Append(PdfAcroFormDictionaryBuilder.FormatColor(effectiveStyle.BorderColor.Value)).Append(']');
        }

        if (effectiveStyle.BackgroundColor.HasValue) {
            sb.Append(" /BG [").Append(PdfAcroFormDictionaryBuilder.FormatColor(effectiveStyle.BackgroundColor.Value)).Append(']');
        }

        return sb.Length == 0 ? string.Empty : " /MK <<" + sb + " >>";
    }

    private static string BuildBorderStyleEntry(PdfFormFieldStyle? style) {
        if (style == null || style.BorderWidth <= 0D) {
            return string.Empty;
        }

        bool hasDashPattern = style.BorderDashPattern != null && style.BorderDashPattern.Count > 0;
        string? styleName = ResolveBorderStyleName(style.BorderStyle, hasDashPattern);
        if (styleName == null) {
            return string.Empty;
        }

        bool isDashed = string.Equals(styleName, "D", StringComparison.Ordinal);
        var sb = new StringBuilder(" /BS << /S /");
        sb.Append(styleName);
        sb.Append(" /W ");
        sb.Append(FormatCoordinate(style.BorderWidth));
        if (isDashed) {
            IReadOnlyList<double> dashPattern = hasDashPattern ? style.BorderDashPattern! : DefaultBorderDashPattern;
            sb.Append(" /D [");
            for (int i = 0; i < dashPattern.Count; i++) {
                if (i > 0) {
                    sb.Append(' ');
                }

                sb.Append(FormatCoordinate(dashPattern[i]));
            }

            sb.Append(']');
        }

        sb.Append(" >>");
        return sb.ToString();
    }

    private static string? ResolveBorderStyleName(PdfFormFieldBorderStyle borderStyle, bool hasDashPattern) {
        switch (borderStyle) {
            case PdfFormFieldBorderStyle.Dashed:
                return "D";
            case PdfFormFieldBorderStyle.Underline:
                return "U";
            case PdfFormFieldBorderStyle.Beveled:
                return "B";
            case PdfFormFieldBorderStyle.Inset:
                return "I";
            case PdfFormFieldBorderStyle.Solid:
                return hasDashPattern ? "D" : null;
            default:
                return null;
        }
    }

    private static void ValidateRadioOptions(IReadOnlyList<string> options, string value) {
        if (options.Count == 0) {
            throw new ArgumentException("PDF radio button group requires at least one option.", nameof(options));
        }

        var optionSet = new HashSet<string>(StringComparer.Ordinal);
        for (int i = 0; i < options.Count; i++) {
            string option = options[i];
            Guard.NotNullOrWhiteSpace(option, nameof(options));
            if (string.Equals(option, "Off", StringComparison.Ordinal)) {
                throw new ArgumentException("PDF radio button option value cannot be Off.", nameof(options));
            }

            ValidateAsciiPdfNameValue(option, nameof(options), "PDF radio button option values must contain only ASCII PDF name characters.");
            if (!optionSet.Add(option)) {
                throw new ArgumentException("PDF radio button options must be unique.", nameof(options));
            }
        }

        if (!optionSet.Contains(value)) {
            throw new ArgumentException("PDF radio button value must match the provided options.", nameof(value));
        }
    }

    private static void ValidateAsciiPdfNameValue(string value, string paramName, string message) {
        for (int i = 0; i < value.Length; i++) {
            if (value[i] > 0x7E) {
                throw new ArgumentException(message, paramName);
            }
        }
    }

    private static void ValidateAsciiPdfNameValue(string value, string paramName) =>
        ValidateAsciiPdfNameValue(value, paramName, "PDF check box selected value name must contain only ASCII PDF name characters.");

    private static void ValidateRectangle(double x1, double y1, double x2, double y2) {
        ValidateFinite(x1, nameof(x1));
        ValidateFinite(y1, nameof(y1));
        ValidateFinite(x2, nameof(x2));
        ValidateFinite(y2, nameof(y2));

        if (x2 <= x1) {
            throw new ArgumentOutOfRangeException(nameof(x2), x2, "PDF link annotation rectangle must have positive width.");
        }

        if (y2 <= y1) {
            throw new ArgumentOutOfRangeException(nameof(y2), y2, "PDF link annotation rectangle must have positive height.");
        }
    }

    private static void ValidateFinite(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value)) {
            throw new ArgumentOutOfRangeException(paramName, value, "PDF annotation coordinates must be finite numbers.");
        }
    }

    private static string FormatCoordinate(double value) =>
        value.ToString("0.###", CultureInfo.InvariantCulture);

}
