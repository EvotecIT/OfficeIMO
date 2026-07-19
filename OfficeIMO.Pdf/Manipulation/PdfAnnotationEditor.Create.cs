namespace OfficeIMO.Pdf;

internal static partial class PdfAnnotationEditor {
    /// <summary>Adds a standard annotation to an existing page and validates readback.</summary>
    public static PdfAnnotationEditResult AddAnnotation(byte[] pdf, PdfAnnotationCreateOptions options) => AddAnnotation(pdf, options, readOptions: null);

    /// <summary>Adds a standard annotation using explicit read limits or credentials and validates readback.</summary>
    public static PdfAnnotationEditResult AddAnnotation(byte[] pdf, PdfAnnotationCreateOptions options, PdfReadOptions? readOptions) {
        Guard.NotNull(pdf, nameof(pdf)); Guard.NotNull(options, nameof(options)); ValidateCreateOptions(options);
        PdfMutationPlan plan = PdfMutationPlanner.Require(pdf, PdfMutationOperation.ModifyAnnotations, readOptions, executionPreference: options.ExecutionPreference);
        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf, readOptions); int catalog = FindCatalogObjectNumber(objects, trailerRaw);
        if (catalog == 0) throw new ArgumentException("PDF does not contain a readable catalog.", nameof(pdf));
        List<int> pages = GetPageObjectNumbersInDocumentOrder(objects);
        if (options.PageNumber > pages.Count) throw new ArgumentOutOfRangeException(nameof(options), "Annotation page number exceeds the PDF page count.");
        int pageObjectNumber = pages[options.PageNumber - 1]; PdfDictionary page = (PdfDictionary)objects[pageObjectNumber].Value;
        int annotationObjectNumber = NextAnnotationObjectNumber(objects);
        var annotation = new PdfDictionary(); annotation.Items["Type"] = new PdfName("Annot"); annotation.Items["Subtype"] = new PdfName(options.Subtype); annotation.Items["P"] = new PdfReference(pageObjectNumber, 0);
        objects[annotationObjectNumber] = new PdfIndirectObject(annotationObjectNumber, 0, annotation);

        int? popupObjectNumber = null;
        if (options.CreatePopup) {
            popupObjectNumber = annotationObjectNumber + 1;
            var popup = new PdfDictionary(); popup.Items["Type"] = new PdfName("Annot"); popup.Items["Subtype"] = new PdfName("Popup"); popup.Items["Parent"] = new PdfReference(annotationObjectNumber, 0); popup.Items["P"] = new PdfReference(pageObjectNumber, 0);
            popup.Items["Rect"] = CreateNumberArray(options.PopupRectangle ?? DefaultPopupRectangle(options.Rectangle)); popup.Items["Open"] = new PdfBoolean(options.PopupOpen);
            objects[popupObjectNumber.Value] = new PdfIndirectObject(popupObjectNumber.Value, 0, popup); annotation.Items["Popup"] = new PdfReference(popupObjectNumber.Value, 0);
        }

        var update = new PdfAnnotationUpdateOptions {
            Contents = options.Contents, Title = options.Title, Name = options.Name, Flags = options.Flags, Color = options.Color,
            Rectangle = options.Rectangle, QuadPoints = options.QuadPoints, Vertices = options.Vertices, Line = options.Line, InkPaths = options.InkPaths,
            LineStartEnding = options.LineStartEnding, LineEndEnding = options.LineEndEnding, InReplyToObjectNumber = options.InReplyToObjectNumber,
            ReplyType = options.ReplyType, RegenerateAppearance = options.GenerateAppearance && IsAppearanceSubtype(options.Subtype)
        };
        if (options.IconName != null) annotation.Items["Name"] = new PdfName(options.IconName);
        IReadOnlyList<int> generatedObjects = ApplyUpdates(objects, annotation, update);
        var references = new List<PdfReference> { new PdfReference(annotationObjectNumber, 0) }; if (popupObjectNumber.HasValue) references.Add(new PdfReference(popupObjectNumber.Value, 0));
        int owner = pageObjectNumber; foreach (PdfReference reference in references) owner = AddAnnotationReference(objects, pageObjectNumber, page, reference);

        byte[] output;
        if (plan.ExecutionMode == PdfMutationExecutionMode.AppendOnly) {
            int[] changed = new[] { owner, annotationObjectNumber }.Concat(popupObjectNumber.HasValue ? new[] { popupObjectNumber.Value } : Array.Empty<int>()).Concat(generatedObjects).Distinct().ToArray();
            output = PdfIncrementalObjectWriter.Append(pdf, objects, plan.Preflight.Probe.Security, trailerRaw, changed, encryptionHandler: GetAppendEncryptionHandler(objects, trailerRaw, readOptions, plan.Preflight.Probe.Security));
            PdfSignatureMutationReport proof = BuildAppendOnlyProof(pdf, output, plan, readOptions); ValidateCreatedAnnotation(output, options, annotationObjectNumber, PdfReadOptions.WithMinimumInputBytes(readOptions, output.LongLength)); return new PdfAnnotationEditResult(output, 1, plan, proof, readOptions: readOptions);
        }
        PdfObjectGraphPruner.PruneUnreachableObjects(objects, catalog); output = RewriteAllObjects(objects, catalog, PdfReadDocument.Open(pdf, readOptions).Metadata, pdf);
        ValidateCreatedAnnotation(output, options, null, PdfReadOptions.WithMinimumInputBytes(readOptions, output.LongLength)); return CreateFullRewriteResult(pdf, output, 1, plan, annotationsChanged: true, readOptions: readOptions);
    }

    private static void ValidateCreateOptions(PdfAnnotationCreateOptions options) {
        if (options.PageNumber <= 0) throw new ArgumentOutOfRangeException(nameof(options), "Annotation page number must be positive.");
        Guard.NotNullOrWhiteSpace(options.Subtype, nameof(options.Subtype));
        if (!IsCreatableSubtype(options.Subtype)) throw new NotSupportedException("This annotation subtype must use a dedicated engine or is not supported for existing-page creation: " + options.Subtype);
        Guard.NonNegative(options.Flags, nameof(options.Flags));
        var update = new PdfAnnotationUpdateOptions { Rectangle = options.Rectangle, Color = options.Color, QuadPoints = options.QuadPoints, Vertices = options.Vertices, Line = options.Line, InkPaths = options.InkPaths, LineStartEnding = options.LineStartEnding, LineEndEnding = options.LineEndEnding, InReplyToObjectNumber = options.InReplyToObjectNumber, ReplyType = options.ReplyType };
        ValidateUpdateOptions(update);
        if (options.CreatePopup) ValidateCoordinateArray(options.PopupRectangle, 4, 4, nameof(options.PopupRectangle));
        ValidatePdfName(options.IconName, nameof(options.IconName));
        if (options.Subtype == "Line" && options.Line is null) throw new ArgumentException("Line annotations require endpoint coordinates.", nameof(options));
        if ((options.Subtype == "Polygon" || options.Subtype == "PolyLine") && options.Vertices is null) throw new ArgumentException("Path annotations require vertices.", nameof(options));
        if (options.Subtype == "Ink" && options.InkPaths is null) throw new ArgumentException("Ink annotations require ink paths.", nameof(options));
    }

    private static bool IsAppearanceSubtype(string subtype) => subtype == "FreeText" || subtype == "Highlight" || subtype == "Underline" || subtype == "Squiggly" || subtype == "StrikeOut" || subtype == "Square" || subtype == "Circle" || subtype == "Line" || subtype == "Ink" || subtype == "Polygon" || subtype == "PolyLine" || subtype == "Stamp" || subtype == "Caret";
    private static bool IsCreatableSubtype(string subtype) => subtype == "Text" || IsAppearanceSubtype(subtype);
    private static double[] DefaultPopupRectangle(IReadOnlyList<double> parent) => new[] { parent[2] + 8D, parent[1], parent[2] + 208D, parent[1] + 120D };
    private static void ValidateCreatedAnnotation(byte[] output, PdfAnnotationCreateOptions options, int? expectedObjectNumber, PdfReadOptions? readOptions) {
        PdfDocumentInfo info = PdfInspector.Inspect(output, readOptions); IReadOnlyList<PdfAnnotation> matches = info.GetAnnotationsBySubtype(options.Subtype); PdfAnnotation? found = matches.Count == 0 ? null : matches[matches.Count - 1];
        if (found == null || found.PageNumber != options.PageNumber || (expectedObjectNumber.HasValue && found.ObjectNumber != expectedObjectNumber)) throw new InvalidOperationException("PDF annotation creation readback failed; the artifact was not returned.");
        if (options.GenerateAppearance && IsAppearanceSubtype(options.Subtype) && !found.HasNormalAppearance) throw new InvalidOperationException("PDF annotation appearance readback failed; the artifact was not returned.");
    }
}
