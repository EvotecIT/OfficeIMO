using System.Globalization;

namespace OfficeIMO.Pdf;

internal static class PdfAnnotationDictionaryBuilder {
    internal static string BuildUriLinkAnnotation(double x1, double y1, double x2, double y2, string uri, string? contents = null, int? structParentIndex = null) {
        ValidateRectangle(x1, y1, x2, y2);
        Guard.AbsoluteUri(uri, nameof(uri));

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
            " /V " +
            PdfSyntaxEscaper.WinAnsiHexString(value) +
            " /DV " +
            PdfSyntaxEscaper.WinAnsiHexString(value) +
            " /Rect [" +
            FormatCoordinate(x1) + " " +
            FormatCoordinate(y1) + " " +
            FormatCoordinate(x2) + " " +
            FormatCoordinate(y2) +
            "] /F 4 /DA " +
            PdfSyntaxEscaper.LiteralString("/Helv " + FormatCoordinate(fontSize) + " Tf " + PdfAcroFormDictionaryBuilder.FormatColor((style ?? new PdfFormFieldStyle()).TextColor) + " rg") +
            BuildMkEntry(style) +
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
                .Append(PdfSyntaxEscaper.WinAnsiHexString(option));
        }

        var valueSet = new HashSet<string>(StringComparer.Ordinal);
        for (int i = 0; i < values.Count; i++) {
            string value = values[i];
            Guard.NotNullOrWhiteSpace(value, nameof(values));
            if (!optionSet.Contains(value)) {
                throw new ArgumentException("PDF choice field values must match the provided options.", nameof(values));
            }

            if (!valueSet.Add(value)) {
                throw new ArgumentException("PDF choice field selected values must be unique.", nameof(values));
            }
        }

        int flags = (isComboBox ? 131072 : 0) | (allowsMultipleSelection ? 2097152 : 0);
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
            BuildMkEntry(style) +
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
            .Append(" /Ff 49152 /V /")
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
            return PdfSyntaxEscaper.WinAnsiHexString(values[0]);
        }

        var valueBuilder = new StringBuilder();
        valueBuilder.Append('[');
        for (int i = 0; i < values.Count; i++) {
            if (i > 0) {
                valueBuilder.Append(' ');
            }

            valueBuilder.Append(PdfSyntaxEscaper.WinAnsiHexString(values[i]));
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
