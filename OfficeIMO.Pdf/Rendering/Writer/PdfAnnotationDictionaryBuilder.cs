using System.Globalization;

namespace OfficeIMO.Pdf;

internal static class PdfAnnotationDictionaryBuilder {
    internal static string BuildUriLinkAnnotation(double x1, double y1, double x2, double y2, string uri, string? contents = null) {
        ValidateRectangle(x1, y1, x2, y2);
        Guard.AbsoluteUri(uri, nameof(uri));

        return "<< /Type /Annot /Subtype /Link /Border [0 0 0]" + BuildContentsEntry(contents) + " /Rect [" +
            FormatCoordinate(x1) + " " +
            FormatCoordinate(y1) + " " +
            FormatCoordinate(x2) + " " +
            FormatCoordinate(y2) +
            "] /A << /S /URI /URI " +
            PdfSyntaxEscaper.LiteralString(uri) +
            " >> >>\n";
    }

    internal static string BuildGoToNamedDestinationLinkAnnotation(double x1, double y1, double x2, double y2, string destinationName, string? contents = null) {
        ValidateRectangle(x1, y1, x2, y2);
        Guard.NotNullOrWhiteSpace(destinationName, nameof(destinationName));

        return "<< /Type /Annot /Subtype /Link /Border [0 0 0]" + BuildContentsEntry(contents) + " /Rect [" +
            FormatCoordinate(x1) + " " +
            FormatCoordinate(y1) + " " +
            FormatCoordinate(x2) + " " +
            FormatCoordinate(y2) +
            "] /A << /S /GoTo /D " +
            PdfSyntaxEscaper.LiteralString(destinationName) +
            " >> >>\n";
    }

    internal static string BuildTextFieldWidgetAnnotation(double x1, double y1, double x2, double y2, string name, string value, double fontSize, int normalAppearanceId) {
        ValidateRectangle(x1, y1, x2, y2);
        Guard.NotNullOrWhiteSpace(name, nameof(name));
        Guard.NotNull(value, nameof(value));
        ValidateFinite(fontSize, nameof(fontSize));
        if (fontSize <= 0) {
            throw new ArgumentOutOfRangeException(nameof(fontSize), fontSize, "PDF text field font size must be a positive finite number.");
        }

        return "<< /Type /Annot /Subtype /Widget /FT /Tx /T " +
            PdfSyntaxEscaper.TextString(name) +
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
            PdfSyntaxEscaper.LiteralString("/Helv " + FormatCoordinate(fontSize) + " Tf 0 g") +
            " /MK << /BC [0.75 0.75 0.75] /BG [1 1 1] >> /AP << /N " +
            PdfSyntaxEscaper.IndirectReference(normalAppearanceId) +
            " >> >>\n";
    }

    internal static string BuildCheckBoxWidgetAnnotation(double x1, double y1, double x2, double y2, string name, bool isChecked, string checkedValueName, int offAppearanceId, int checkedAppearanceId) {
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
            " /MK << /BC [0.75 0.75 0.75] /BG [1 1 1] >> /AP << /N << /Off " +
            PdfSyntaxEscaper.IndirectReference(offAppearanceId) +
            " /" +
            PdfSyntaxEscaper.Name(checkedValueName) +
            " " +
            PdfSyntaxEscaper.IndirectReference(checkedAppearanceId) +
            " >> >> >>\n";
    }

    internal static string BuildChoiceFieldWidgetAnnotation(double x1, double y1, double x2, double y2, string name, IReadOnlyList<string> options, string value, double fontSize, int normalAppearanceId, bool isComboBox) =>
        BuildChoiceFieldWidgetAnnotation(x1, y1, x2, y2, name, options, new[] { value }, fontSize, normalAppearanceId, isComboBox, allowsMultipleSelection: false);

    internal static string BuildChoiceFieldWidgetAnnotation(double x1, double y1, double x2, double y2, string name, IReadOnlyList<string> options, IReadOnlyList<string> values, double fontSize, int normalAppearanceId, bool isComboBox, bool allowsMultipleSelection) {
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
            PdfSyntaxEscaper.LiteralString("/Helv " + FormatCoordinate(fontSize) + " Tf 0 g") +
            " /MK << /BC [0.75 0.75 0.75] /BG [1 1 1] >> /AP << /N " +
            PdfSyntaxEscaper.IndirectReference(normalAppearanceId) +
            " >> >>\n";
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

    private static void ValidateAsciiPdfNameValue(string value, string paramName) {
        for (int i = 0; i < value.Length; i++) {
            if (value[i] > 0x7E) {
                throw new ArgumentException("PDF check box selected value name must contain only ASCII PDF name characters.", paramName);
            }
        }
    }

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
