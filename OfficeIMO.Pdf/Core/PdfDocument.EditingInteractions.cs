namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    /// <summary>Reads visual hit regions for permitted annotation and form editing, without extracting page text or images.</summary>
    /// <remarks>Annotation regions require an executable annotation mutation plan. Form widgets require content-copy, form-filling, or form-authoring authorization. Unavailable regions are omitted.</remarks>
    public PdfPageInteractionMap GetEditingInteractions(int pageNumber) {
        PdfReadDocument source = GetReadDocument(ReadOptions);
        IReadOnlyList<PdfPageLayoutInfo> layouts = GetPageLayouts();
        if (pageNumber < 1 || pageNumber > layouts.Count) throw new ArgumentOutOfRangeException(nameof(pageNumber));
        IReadOnlyList<PdfAnnotation> annotations = PdfPermissionAuthorization.CanMutate(source.Security, ReadOptions.PermissionPolicy,
            PdfMutationOperation.ModifyAnnotations) && PlanMutation(PdfMutationOperation.ModifyAnnotations).CanExecute
            ? Annotations.GetForEditing() : [];
        IReadOnlyList<PdfFormField> fields = PdfPermissionAuthorization.CanReadFormMetadata(source.Security, ReadOptions.PermissionPolicy)
            ? source.UncheckedFormFields : [];
        return PdfPageInteractionMap.CreateAnnotationEditing(layouts[pageNumber - 1], annotations, fields);
    }
}
