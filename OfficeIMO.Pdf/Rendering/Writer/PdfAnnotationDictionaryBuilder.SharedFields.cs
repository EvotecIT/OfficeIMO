namespace OfficeIMO.Pdf;

internal static partial class PdfAnnotationDictionaryBuilder {
    /// <summary>Writes one logical text field whose visible fragments are separate widgets.</summary>
    internal static string BuildTextFieldDictionary(string name, string value, double fontSize, IReadOnlyList<int> widgetObjectIds, PdfFormFieldStyle? style = null) =>
        "<<" + BuildTextFieldEntries(name, value, fontSize, style, isWidget: false) +
        BuildSharedFieldDefaultAppearance(fontSize, style) + BuildQuaddingEntry(style) +
        BuildSharedFieldKids(widgetObjectIds) + " >>\n";

    /// <summary>Writes a shared choice field while retaining scalar-option validation.</summary>
    internal static string BuildChoiceFieldDictionary(string name, IReadOnlyList<string> options, IReadOnlyList<string> values, double fontSize, bool isComboBox, bool allowsMultipleSelection, IReadOnlyList<int> widgetObjectIds, PdfFormFieldStyle? style = null, IReadOnlyList<int>? selectedIndices = null, int? topIndex = null) {
        Guard.NotNull(options, nameof(options));
        return BuildSharedChoiceFieldDictionary(name, options.Select(option => new PdfFormFieldOption(option, option)).ToList(), values, fontSize, isComboBox, allowsMultipleSelection, widgetObjectIds, style, selectedIndices, topIndex, requireUniqueExportValues: true, emitPairedOptions: false);
    }

    /// <summary>Writes a shared choice field with distinct export and display option values.</summary>
    internal static string BuildChoiceFieldDictionary(string name, IReadOnlyList<PdfFormFieldOption> options, IReadOnlyList<string> values, double fontSize, bool isComboBox, bool allowsMultipleSelection, IReadOnlyList<int> widgetObjectIds, PdfFormFieldStyle? style = null, IReadOnlyList<int>? selectedIndices = null, int? topIndex = null) =>
        BuildSharedChoiceFieldDictionary(name, options, values, fontSize, isComboBox, allowsMultipleSelection, widgetObjectIds, style, selectedIndices, topIndex, requireUniqueExportValues: false, emitPairedOptions: true);

    private static string BuildSharedChoiceFieldDictionary(string name, IReadOnlyList<PdfFormFieldOption> options, IReadOnlyList<string> values, double fontSize, bool isComboBox, bool allowsMultipleSelection, IReadOnlyList<int> widgetObjectIds, PdfFormFieldStyle? style, IReadOnlyList<int>? selectedIndices, int? topIndex, bool requireUniqueExportValues, bool emitPairedOptions) =>
        "<<" + BuildChoiceFieldEntries(name, options, values, fontSize, isComboBox, allowsMultipleSelection, style, requireUniqueExportValues, emitPairedOptions, selectedIndices, topIndex, isWidget: false) +
        BuildSharedFieldDefaultAppearance(fontSize, style) + BuildQuaddingEntry(style) +
        BuildSharedFieldKids(widgetObjectIds) + " >>\n";

    private static string BuildTextFieldEntries(string name, string value, double fontSize, PdfFormFieldStyle? style, bool isWidget) {
        Guard.NotNullOrWhiteSpace(name, nameof(name));
        Guard.NotNull(value, nameof(value));
        ValidateFinite(fontSize, nameof(fontSize));
        if (fontSize <= 0) {
            throw new ArgumentOutOfRangeException(nameof(fontSize), fontSize, "PDF text field font size must be a positive finite number.");
        }

        return " /FT /Tx /T " + PdfSyntaxEscaper.TextString(name) +
            BuildFormFieldMetadataEntries(style, isWidget) + BuildTextFieldFlagsEntry(style) +
            BuildMaxLengthEntry(style) + " /V " + PdfSyntaxEscaper.TextString(value) +
            " /DV " + PdfSyntaxEscaper.TextString(value);
    }

    private static string BuildSharedFieldDefaultAppearance(double fontSize, PdfFormFieldStyle? style) =>
        " /DA " + PdfSyntaxEscaper.LiteralString("/Helv " + FormatCoordinate(fontSize) + " Tf " +
            PdfAcroFormDictionaryBuilder.FormatColor((style ?? new PdfFormFieldStyle()).TextColor) + " rg");

    private static string BuildSharedFieldKids(IReadOnlyList<int> widgetObjectIds) {
        Guard.NotNull(widgetObjectIds, nameof(widgetObjectIds));
        if (widgetObjectIds.Count == 0) {
            throw new ArgumentException("PDF shared form fields require at least one widget.", nameof(widgetObjectIds));
        }

        var seen = new HashSet<int>();
        var kids = new StringBuilder(" /Kids [");
        foreach (int objectId in widgetObjectIds) {
            if (objectId <= 0 || !seen.Add(objectId)) {
                throw new ArgumentException("PDF shared form field widget object ids must be positive and unique.", nameof(widgetObjectIds));
            }
            kids.Append(' ').Append(PdfSyntaxEscaper.IndirectReference(objectId));
        }
        return kids.Append(" ]").ToString();
    }

    /// <summary>Writes widget-only data; the parent owns the name, value, flags and options.</summary>
    private static string BuildChildFormWidgetAnnotation(double x1, double y1, double x2, double y2, int parentObjectId, int normalAppearanceId, PdfFormFieldStyle? style, int? structParentIndex) {
        ValidateRectangle(x1, y1, x2, y2);
        if (parentObjectId <= 0) {
            throw new ArgumentOutOfRangeException(nameof(parentObjectId), parentObjectId, "PDF widget parent object id must be positive.");
        }
        if (normalAppearanceId <= 0) {
            throw new ArgumentOutOfRangeException(nameof(normalAppearanceId), normalAppearanceId, "PDF widget appearance object id must be positive.");
        }

        return "<< /Type /Annot /Subtype /Widget /Parent " + PdfSyntaxEscaper.IndirectReference(parentObjectId) +
            " /Rect [" + FormatCoordinate(x1) + " " + FormatCoordinate(y1) + " " +
            FormatCoordinate(x2) + " " + FormatCoordinate(y2) + "] /F 4" +
            BuildFormFieldAlternateNameEntry(style) + BuildContentsEntry(style?.AlternateName) +
            BuildMkEntry(style) + BuildBorderStyleEntry(style) +
            " /AP << /N " + PdfSyntaxEscaper.IndirectReference(normalAppearanceId) + " >>" +
            BuildStructParentEntry(structParentIndex) + " >>\n";
    }
}
