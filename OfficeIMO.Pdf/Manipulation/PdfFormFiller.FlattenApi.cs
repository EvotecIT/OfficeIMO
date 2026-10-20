using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfFormFiller {
    /// <summary>
    /// Returns a new PDF with simple text, choice, and button AcroForm widgets painted into page content and removed from the form tree.
    /// </summary>
    public static byte[] FlattenFields(byte[] pdf) {
        return FlattenFields(pdf, options: null, readOptions: null);
    }

    /// <summary>
    /// Returns a new PDF with simple text, choice, and button AcroForm widgets painted into page content and removed from the form tree.
    /// </summary>
    public static byte[] FlattenFields(byte[] pdf, PdfFormFillerOptions? options) {
        return FlattenFields(pdf, options, readOptions: null);
    }

    internal static byte[] FlattenFields(byte[] pdf, PdfFormFillerOptions? options, PdfLoadOptions? readOptions, CancellationToken cancellationToken = default) =>
        FlattenFields(pdf, options, readOptions, out _, cancellationToken);

    internal static byte[] FlattenFields(byte[] pdf, PdfFormFillerOptions? options, PdfLoadOptions? readOptions,
        out PdfLoadOptions outputReadOptions, CancellationToken cancellationToken = default) {
        using var linkedCancellation = LinkCancellation(options?.CancellationToken ?? default, ref cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        Guard.NotNull(pdf, nameof(pdf));
        _ = PdfMutationPlanner.RequireFullRewrite(pdf, PdfMutationOperation.FlattenFormFields, readOptions, cancellationToken: cancellationToken);

        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf, readOptions, out _, out _, cancellationToken);
        int catalogObjectNumber = FindCatalogObjectNumber(objects, trailerRaw);
        if (catalogObjectNumber == 0 ||
            objects[catalogObjectNumber].Value is not PdfDictionary catalog ||
            !catalog.Items.TryGetValue("AcroForm", out var acroFormObject) ||
            ResolveDictionary(objects, acroFormObject) is not PdfDictionary acroForm ||
            !acroForm.Items.TryGetValue("Fields", out var fieldsObject) ||
            ResolveObject(objects, fieldsObject) is not PdfArray fields) {
            throw new ArgumentException("PDF does not contain a readable AcroForm field tree.", nameof(pdf));
        }

        int nextObjectNumber = objects.Keys.Count == 0 ? 1 : objects.Keys.Max() + 1;
        int firstGeneratedObject = nextObjectNumber;
        PdfDictionary? acroFormDefaultResources = TryReadDefaultResources(objects, acroForm);
        string? acroFormDefaultAppearance = TryReadText(objects, acroForm, "DA");
        var widgets = new Dictionary<int, FlattenWidgetState>();
        var removableObjects = new HashSet<int>();
        for (int i = 0; i < fields.Items.Count; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            CollectFlattenWidgets(objects, fields.Items[i], null, 0, ReadFieldQuadding(objects, acroForm, null), null, acroFormDefaultResources, acroFormDefaultAppearance, null, null, null, null, options, widgets, removableObjects, new HashSet<int>(), ref nextObjectNumber, cancellationToken);
        }

        if (widgets.Count == 0) {
            throw new NotSupportedException(UnsupportedFlattenWidgetMessage);
        }

        int flattenedWidgetCount = FlattenPageWidgets(objects, widgets, ref nextObjectNumber, readOptions, firstGeneratedObject, out PdfGeneratedOutputGrowth generatedGrowth, cancellationToken);
        if (flattenedWidgetCount != widgets.Count) {
            throw new NotSupportedException(UnsupportedFlattenAnnotationMessage);
        }

        catalog.Items.Remove("AcroForm");
        if (acroFormObject is PdfReference acroFormReference) {
            removableObjects.Add(acroFormReference.ObjectNumber);
        }

        PdfStructureTreeAnnotationPruner.RemoveAnnotationReferences(objects, widgets.Keys, cancellationToken);
        foreach (int objectNumber in removableObjects) {
            cancellationToken.ThrowIfCancellationRequested();
            objects.Remove(objectNumber);
        }

        PdfObjectGraphPruner.PruneUnreachableObjects(objects, catalogObjectNumber, cancellationToken);
        return RewriteAllObjects(objects, catalogObjectNumber, PdfReadDocument.Open(pdf, readOptions, cancellationToken), pdf, generatedGrowth, out outputReadOptions, cancellationToken);
    }

    /// <summary>
    /// Returns a new PDF with only the named simple text, choice, and button AcroForm fields painted into page content and removed from the form tree.
    /// </summary>
    public static byte[] FlattenFields(byte[] pdf, IReadOnlyCollection<string> fieldNames, PdfFormFillerOptions? options = null) {
        return FlattenFieldsCore(pdf, fieldNames, options, readOptions: null, requireMutationPlan: true, out _);
    }

    internal static byte[] FlattenFields(byte[] pdf, IReadOnlyCollection<string> fieldNames, PdfFormFillerOptions? options, PdfLoadOptions? readOptions) =>
        FlattenFieldsCore(pdf, fieldNames, options, readOptions, requireMutationPlan: true, out _);

    internal static byte[] FlattenFields(byte[] pdf, IReadOnlyCollection<string> fieldNames, PdfFormFillerOptions? options,
        PdfLoadOptions? readOptions, out PdfLoadOptions outputReadOptions) =>
        FlattenFieldsCore(pdf, fieldNames, options, readOptions, requireMutationPlan: true, out outputReadOptions);

    internal static byte[] FlattenFieldsWithinPlannedRewrite(byte[] pdf, IReadOnlyCollection<string> fieldNames, PdfFormFillerOptions? options = null, PdfLoadOptions? readOptions = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        return FlattenFieldsCore(pdf, fieldNames, options, readOptions, requireMutationPlan: false, out _, cancellationToken);
    }

    internal static byte[] FlattenFieldsWithinPlannedRewrite(byte[] pdf, IReadOnlyCollection<string> fieldNames,
        PdfFormFillerOptions? options, PdfLoadOptions? readOptions, out PdfLoadOptions outputReadOptions,
        System.Threading.CancellationToken cancellationToken = default) =>
        FlattenFieldsCore(pdf, fieldNames, options, readOptions, requireMutationPlan: false, out outputReadOptions, cancellationToken);

    private static byte[] FlattenFieldsCore(byte[] pdf, IReadOnlyCollection<string> fieldNames, PdfFormFillerOptions? options, PdfLoadOptions? readOptions, bool requireMutationPlan, out PdfLoadOptions outputReadOptions, System.Threading.CancellationToken cancellationToken = default) {
        using var linkedCancellation = LinkCancellation(options?.CancellationToken ?? default, ref cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        Guard.NotNull(pdf, nameof(pdf));
        ValidateFlattenFieldNames(fieldNames, cancellationToken);
        if (requireMutationPlan) _ = PdfMutationPlanner.RequireFullRewriteDocument(pdf, PdfMutationOperation.FlattenFormFields, readOptions, fieldNames: fieldNames, cancellationToken: cancellationToken);

        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf, readOptions, out _, out _, cancellationToken);
        int catalogObjectNumber = FindCatalogObjectNumber(objects, trailerRaw);
        if (catalogObjectNumber == 0 ||
            objects[catalogObjectNumber].Value is not PdfDictionary catalog ||
            !catalog.Items.TryGetValue("AcroForm", out var acroFormObject) ||
            ResolveDictionary(objects, acroFormObject) is not PdfDictionary acroForm ||
            !acroForm.Items.TryGetValue("Fields", out var fieldsObject) ||
            ResolveObject(objects, fieldsObject) is not PdfArray fields) {
            throw new ArgumentException("PDF does not contain a readable AcroForm field tree.", nameof(pdf));
        }

        var requested = new HashSet<string>(fieldNames, StringComparer.Ordinal);
        var matched = new HashSet<string>(StringComparer.Ordinal);
        var widgets = new Dictionary<int, FlattenWidgetState>();
        var removableObjects = new HashSet<int>();
        int nextObjectNumber = objects.Keys.Count == 0 ? 1 : objects.Keys.Max() + 1;
        int firstGeneratedObject = nextObjectNumber;
        CollectSelectedFlattenFields(
            objects,
            fields,
            requested,
            matched,
            inheritedFieldType: null,
            inheritedFlags: 0,
            inheritedQuadding: ReadFieldQuadding(objects, acroForm, null),
            inheritedMaxLength: null,
            inheritedDefaultResources: TryReadDefaultResources(objects, acroForm),
            inheritedDefaultAppearance: TryReadText(objects, acroForm, "DA"),
            inheritedDisplayValue: null,
            inheritedRichAppearanceRuns: null,
            inheritedName: null,
            inheritedChoiceOptions: null,
            options,
            widgets,
            removableObjects,
            ref nextObjectNumber, cancellationToken);

        if (matched.Count != requested.Count) {
            throw new ArgumentException("PDF form field was not found: " + string.Join(", ", requested.Where(name => !matched.Contains(name))), nameof(fieldNames));
        }
        if (widgets.Count == 0) {
            throw new NotSupportedException(UnsupportedFlattenWidgetMessage);
        }

        int flattenedWidgetCount = FlattenPageWidgets(objects, widgets, ref nextObjectNumber, readOptions, firstGeneratedObject, out PdfGeneratedOutputGrowth generatedGrowth, cancellationToken);
        if (flattenedWidgetCount != widgets.Count) {
            throw new NotSupportedException(UnsupportedFlattenAnnotationMessage);
        }

        FilterCalculationOrder(objects, acroForm, removableObjects, cancellationToken);
        if (fields.Items.Count == 0) {
            catalog.Items.Remove("AcroForm");
            if (acroFormObject is PdfReference acroFormReference) removableObjects.Add(acroFormReference.ObjectNumber);
        }

        PdfStructureTreeAnnotationPruner.RemoveAnnotationReferences(objects, widgets.Keys, cancellationToken);
        foreach (int objectNumber in removableObjects) {
            cancellationToken.ThrowIfCancellationRequested();
            objects.Remove(objectNumber);
        }
        PdfObjectGraphPruner.PruneUnreachableObjects(objects, catalogObjectNumber, cancellationToken);
        return RewriteAllObjects(objects, catalogObjectNumber, PdfReadDocument.Open(pdf, readOptions, cancellationToken), pdf, generatedGrowth, out outputReadOptions, cancellationToken);
    }

}
