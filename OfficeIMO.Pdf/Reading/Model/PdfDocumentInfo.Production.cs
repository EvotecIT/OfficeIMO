namespace OfficeIMO.Pdf;

public sealed partial class PdfDocumentInfo {
    /// <summary>Number of AcroForm text fields.</summary>
    public int TextFormFieldCount => CountFormFields(PdfFormFieldKind.Text);

    /// <summary>Number of AcroForm button fields.</summary>
    public int ButtonFormFieldCount => CountFormFields(PdfFormFieldKind.Button);

    /// <summary>Number of AcroForm choice fields.</summary>
    public int ChoiceFormFieldCount => CountFormFields(PdfFormFieldKind.Choice);

    /// <summary>Number of AcroForm signature fields.</summary>
    public int SignatureFormFieldCount => CountFormFields(PdfFormFieldKind.Signature);

    /// <summary>Number of fields with the common read-only flag.</summary>
    public int ReadOnlyFormFieldCount => FormFields.Count(static formField => formField.IsReadOnly);

    /// <summary>Number of fields with the common required flag.</summary>
    public int RequiredFormFieldCount => FormFields.Count(static formField => formField.IsRequired);

    /// <summary>Number of fields with the common no-export flag.</summary>
    public int NoExportFormFieldCount => FormFields.Count(static formField => formField.IsNoExport);

    /// <summary>Number of pages with CropBox, BleedBox, TrimBox, or ArtBox information.</summary>
    public int PageProductionBoxCount => Pages.Count(static page => page.Geometry.HasNonDefaultBoundaryBoxes);

    /// <summary>Number of pages with a TrimBox.</summary>
    public int TrimBoxPageCount => Pages.Count(static page => page.TrimBox is not null);

    /// <summary>Number of pages with a BleedBox.</summary>
    public int BleedBoxPageCount => Pages.Count(static page => page.BleedBox is not null);

    /// <summary>Number of pages with an ArtBox.</summary>
    public int ArtBoxPageCount => Pages.Count(static page => page.ArtBox is not null);

    /// <summary>True when any page exposes production-oriented boxes such as TrimBox, BleedBox, or ArtBox.</summary>
    public bool HasProductionPageBoxes => TrimBoxPageCount > 0 || BleedBoxPageCount > 0 || ArtBoxPageCount > 0;

    /// <summary>Number of annotations with a primary, additional, or chained action.</summary>
    public int ActiveAnnotationCount => Annotations.Count(static annotation => annotation.HasAction || annotation.HasAdditionalActions || annotation.HasChainedActions);

    /// <summary>Number of annotations with JavaScript, launch, submit, import-data, or movie/rendition actions.</summary>
    public int RiskyAnnotationActionCount => Annotations.Count(static annotation => HasRiskyAnnotationAction(annotation));

    /// <summary>Readable annotation counts grouped by subtype.</summary>
    public IReadOnlyDictionary<string, int> AnnotationSubtypeCounts => Annotations
        .GroupBy(static annotation => annotation.Subtype, StringComparer.Ordinal)
        .ToDictionary(static group => group.Key, static group => group.Count(), StringComparer.Ordinal);

    /// <summary>Readable annotation counts grouped by action type.</summary>
    public IReadOnlyDictionary<string, int> AnnotationActionTypeCounts => AnnotationsByActionType
        .ToDictionary(static item => item.Key, static item => item.Value.Count, StringComparer.Ordinal);

    private int CountFormFields(PdfFormFieldKind kind) => FormFields.Count(formField => formField.Kind == kind);

    private static bool HasRiskyAnnotationAction(PdfAnnotation annotation) {
        if (IsRiskyAction(annotation.ActionType)) {
            return true;
        }

        for (int i = 0; i < annotation.AdditionalActions.Count; i++) {
            if (IsRiskyAction(annotation.AdditionalActions[i].ActionType)) {
                return true;
            }
        }

        for (int i = 0; i < annotation.ChainedActions.Count; i++) {
            if (IsRiskyAction(annotation.ChainedActions[i].ActionType)) {
                return true;
            }
        }

        return false;
    }

    private static bool IsRiskyAction(string? actionType) =>
        string.Equals(actionType, "JavaScript", StringComparison.Ordinal) ||
        string.Equals(actionType, "Launch", StringComparison.Ordinal) ||
        string.Equals(actionType, "SubmitForm", StringComparison.Ordinal) ||
        string.Equals(actionType, "ImportData", StringComparison.Ordinal) ||
        string.Equals(actionType, "Movie", StringComparison.Ordinal) ||
        string.Equals(actionType, "Rendition", StringComparison.Ordinal);
}
