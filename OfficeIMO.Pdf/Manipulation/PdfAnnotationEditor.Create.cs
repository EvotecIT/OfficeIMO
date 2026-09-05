namespace OfficeIMO.Pdf;

internal static partial class PdfAnnotationEditor {
    private static readonly double[] InvisibleLinkBorder = { 0D, 0D, 0D };
    /// <summary>Adds a standard annotation to an existing page and validates readback.</summary>
    public static PdfAnnotationEditResult AddAnnotation(byte[] pdf, PdfAnnotationCreateOptions options) => AddAnnotation(pdf, options, readOptions: null);

    /// <summary>Adds a standard annotation using explicit read limits or credentials and validates readback.</summary>
    public static PdfAnnotationEditResult AddAnnotation(byte[] pdf, PdfAnnotationCreateOptions options, PdfLoadOptions? readOptions) {
        Guard.NotNull(pdf, nameof(pdf)); Guard.NotNull(options, nameof(options)); ValidateCreateOptions(options);
        PdfMutationPlan plan = PdfMutationPlanner.Require(pdf, PdfMutationOperation.ModifyAnnotations, readOptions, executionPreference: options.ExecutionPreference);
        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf, readOptions); int catalog = FindCatalogObjectNumber(objects, trailerRaw);
        if (catalog == 0) throw new ArgumentException("PDF does not contain a readable catalog.", nameof(pdf));
        ValidateLinkUriAgainstCatalog(options, objects, catalog);
        List<int> pages = GetPageObjectNumbersInDocumentOrder(objects);
        if (options.PageNumber > pages.Count) throw new ArgumentOutOfRangeException(nameof(options), "Annotation page number exceeds the PDF page count.");
        int pageObjectNumber = pages[options.PageNumber - 1]; PdfIndirectObject pageIndirect = objects[pageObjectNumber]; PdfDictionary page = (PdfDictionary)pageIndirect.Value;
        int annotationObjectNumber = NextAnnotationObjectNumber(objects);
        var annotation = new PdfDictionary(); annotation.Items["Type"] = new PdfName("Annot"); annotation.Items["Subtype"] = new PdfName(options.Subtype); annotation.Items["P"] = new PdfReference(pageObjectNumber, pageIndirect.Generation);
        objects[annotationObjectNumber] = new PdfIndirectObject(annotationObjectNumber, 0, annotation);

        int? popupObjectNumber = null;
        if (options.CreatePopup) {
            popupObjectNumber = annotationObjectNumber + 1;
            var popup = new PdfDictionary(); popup.Items["Type"] = new PdfName("Annot"); popup.Items["Subtype"] = new PdfName("Popup"); popup.Items["Parent"] = new PdfReference(annotationObjectNumber, 0); popup.Items["P"] = new PdfReference(pageObjectNumber, pageIndirect.Generation);
            popup.Items["Rect"] = CreateNumberArray(options.PopupRectangle ?? DefaultPopupRectangle(options.Rectangle)); popup.Items["Open"] = new PdfBoolean(options.PopupOpen);
            objects[popupObjectNumber.Value] = new PdfIndirectObject(popupObjectNumber.Value, 0, popup); annotation.Items["Popup"] = new PdfReference(popupObjectNumber.Value, 0);
        }

        var update = new PdfAnnotationUpdateOptions {
            Contents = options.Contents, Title = options.Title, Name = options.Name, Flags = options.Flags, Color = options.Color,
            InteriorColor = options.InteriorColor, Opacity = options.Opacity, BorderWidth = options.BorderWidth, BorderStyle = options.BorderStyle, BorderDashPattern = options.BorderDashPattern,
            Rectangle = options.Rectangle, QuadPoints = options.QuadPoints, Vertices = options.Vertices, Line = options.Line, InkPaths = options.InkPaths,
            LineStartEnding = options.LineStartEnding, LineEndEnding = options.LineEndEnding, InReplyToObjectNumber = options.InReplyToObjectNumber,
            ReplyType = options.ReplyType, ReviewState = options.ReviewState, Subject = options.Subject, Intent = options.Intent,
            RegenerateAppearance = options.GenerateAppearance && IsAppearanceSubtype(options.Subtype)
        };
        if (options.LinkUri != null) {
            var action = new PdfDictionary();
            action.Items["S"] = new PdfName("URI");
            action.Items["URI"] = new PdfStringObj(options.LinkUri, useTextStringEncoding: true);
            annotation.Items["A"] = action;
            annotation.Items["Border"] = CreateNumberArray(InvisibleLinkBorder);
        }
        if (options.IconName != null) annotation.Items["Name"] = new PdfName(options.IconName);
        IReadOnlyList<int> generatedObjects = ApplyUpdates(objects, annotation, update);
        var references = new List<PdfReference> { new PdfReference(annotationObjectNumber, 0) }; if (popupObjectNumber.HasValue) references.Add(new PdfReference(popupObjectNumber.Value, 0));
        int owner = pageObjectNumber; foreach (PdfReference reference in references) owner = AddAnnotationReference(objects, pageObjectNumber, page, reference);
        PdfGeneratedOutputGrowth generatedGrowth = BuildGeneratedOutputGrowth(
            objects,
            generatedObjects
                .Concat(references.Select(static reference => reference.ObjectNumber))
                .Concat(new[] { owner }),
            additionalAnnotationsPerPage: references.Count,
            additionalRevisions: plan.ExecutionMode == PdfMutationExecutionMode.AppendOnly ? 1 : 0);

        byte[] output;
        if (plan.ExecutionMode == PdfMutationExecutionMode.AppendOnly) {
            int[] changed = new[] { owner, annotationObjectNumber }.Concat(popupObjectNumber.HasValue ? new[] { popupObjectNumber.Value } : Array.Empty<int>()).Concat(generatedObjects).Distinct().ToArray();
            output = PdfIncrementalObjectWriter.Append(pdf, objects, plan.Preflight.Probe.Security, trailerRaw, changed, encryptionHandler: GetAppendEncryptionHandler(objects, trailerRaw, readOptions, plan.Preflight.Probe.Security));
            PdfLoadOptions outputReadOptions = PdfLoadOptions.ForGeneratedOutput(readOptions, pdf, output, generatedGrowth);
            PdfSignatureMutationReport proof = BuildAppendOnlyProof(pdf, output, plan, readOptions, outputReadOptions); ValidateCreatedAnnotation(output, options, annotationObjectNumber, options.InReplyToObjectNumber, outputReadOptions); return new PdfAnnotationEditResult(output, 1, plan, proof, readOptions: outputReadOptions);
        }
        PdfObjectGraphPruner.PruneUnreachableObjects(objects, catalog); output = RewriteAllObjects(objects, catalog, PdfReadDocument.Open(pdf, readOptions).UncheckedMetadata, pdf, out IReadOnlyDictionary<int, int> numberMap);
        int? rewrittenParent = options.InReplyToObjectNumber.HasValue ? numberMap[options.InReplyToObjectNumber.Value] : null;
        PdfLoadOptions rewrittenReadOptions = PdfLoadOptions.ForGeneratedOutput(readOptions, pdf, output, generatedGrowth);
        ValidateCreatedAnnotation(output, options, numberMap[annotationObjectNumber], rewrittenParent, rewrittenReadOptions); return CreateFullRewriteResult(pdf, output, 1, plan, annotationsChanged: true, readOptions: readOptions, rewrittenReadOptions: rewrittenReadOptions);
    }

    private static void ValidateCreateOptions(PdfAnnotationCreateOptions options) {
        if (options.PageNumber <= 0) throw new ArgumentOutOfRangeException(nameof(options), "Annotation page number must be positive.");
        Guard.NotNullOrWhiteSpace(options.Subtype, nameof(options.Subtype));
        if (!IsCreatableSubtype(options.Subtype)) throw new NotSupportedException("This annotation subtype must use a dedicated engine or is not supported for existing-page creation: " + options.Subtype);
        Guard.NonNegative(options.Flags, nameof(options.Flags));
        var update = new PdfAnnotationUpdateOptions { Rectangle = options.Rectangle, Color = options.Color, InteriorColor = options.InteriorColor, Opacity = options.Opacity, BorderWidth = options.BorderWidth, BorderStyle = options.BorderStyle, BorderDashPattern = options.BorderDashPattern, QuadPoints = options.QuadPoints, Vertices = options.Vertices, Line = options.Line, InkPaths = options.InkPaths, LineStartEnding = options.LineStartEnding, LineEndEnding = options.LineEndEnding, InReplyToObjectNumber = options.InReplyToObjectNumber, ReplyType = options.ReplyType, ReviewState = options.ReviewState, Subject = options.Subject, Intent = options.Intent };
        ValidateUpdateOptions(update);
        if (options.CreatePopup) ValidateCoordinateArray(options.PopupRectangle, 4, 4, nameof(options.PopupRectangle));
        ValidatePdfName(options.IconName, nameof(options.IconName));
        if (options.Subtype == "Line" && options.Line is null) throw new ArgumentException("Line annotations require endpoint coordinates.", nameof(options));
        if ((options.Subtype == "Polygon" || options.Subtype == "PolyLine") && options.Vertices is null) throw new ArgumentException("Path annotations require vertices.", nameof(options));
        if (options.Subtype == "Ink" && options.InkPaths is null) throw new ArgumentException("Ink annotations require ink paths.", nameof(options));
        if (options.Subtype == "Link") {
            Guard.NotNullOrWhiteSpace(options.LinkUri, nameof(options.LinkUri));
            Guard.UriAction(options.LinkUri!, nameof(options.LinkUri));
            if (options.Title != null ||
                options.IconName != null ||
                options.InReplyToObjectNumber.HasValue ||
                options.ReplyType != null ||
                options.ReviewState.HasValue ||
                options.Subject != null ||
                options.Intent != null ||
                options.CreatePopup ||
                options.PopupRectangle != null ||
                options.PopupOpen) {
                throw new ArgumentException(
                    "Link annotations do not support markup-only author, popup, reply, review, subject, intent, or icon options.",
                    nameof(options));
            }
        } else if (options.LinkUri != null) {
            throw new ArgumentException("LinkUri can be used only with Link annotations.", nameof(options));
        }
    }

    private static bool IsAppearanceSubtype(string subtype) => subtype == "Text" || subtype == "FreeText" || subtype == "Highlight" || subtype == "Underline" || subtype == "Squiggly" || subtype == "StrikeOut" || subtype == "Square" || subtype == "Circle" || subtype == "Line" || subtype == "Ink" || subtype == "Polygon" || subtype == "PolyLine" || subtype == "Stamp" || subtype == "Caret";
    private static bool IsCreatableSubtype(string subtype) => subtype == "Text" || subtype == "Link" || subtype == "Redact" || IsAppearanceSubtype(subtype);

    private static void ValidateLinkUriAgainstCatalog(
        PdfAnnotationCreateOptions options,
        Dictionary<int, PdfIndirectObject> objects,
        int catalogObjectNumber) {
        if (options.LinkUri == null ||
            Uri.TryCreate(options.LinkUri, UriKind.Absolute, out _)) return;

        if (objects.TryGetValue(catalogObjectNumber, out PdfIndirectObject? catalogObject) &&
            catalogObject.Value is PdfDictionary catalog &&
            catalog.Items.TryGetValue("URI", out PdfObject? uriObject) &&
            PdfObjectLookup.TryResolveReferenceChain(objects, uriObject, out PdfObject? resolvedUri) &&
            resolvedUri is PdfDictionary uriDictionary &&
            uriDictionary.Items.TryGetValue("Base", out PdfObject? baseObject) &&
            PdfObjectLookup.TryResolveReferenceChain(objects, baseObject, out PdfObject? resolvedBase) &&
            resolvedBase is PdfStringObj baseString &&
            Uri.TryCreate(baseString.Value, UriKind.Absolute, out _)) return;

        throw new ArgumentException(
            "Relative PDF URI link targets require an existing catalog URI base.",
            nameof(options));
    }

    private static PdfGeneratedOutputGrowth BuildGeneratedOutputGrowth(
        Dictionary<int, PdfIndirectObject> objects,
        IEnumerable<int> generatedObjectNumbers,
        int additionalAnnotationsPerPage = 0,
        int additionalRevisions = 0) {
        return PdfGeneratedOutputGrowth.FromSerializedObjects(
            objects,
            generatedObjectNumbers,
            additionalAnnotationsPerPage,
            additionalRevisions);
    }
    private static double[] DefaultPopupRectangle(IReadOnlyList<double> parent) => new[] { parent[2] + 8D, parent[1], parent[2] + 208D, parent[1] + 120D };
    private static void ValidateCreatedAnnotation(byte[] output, PdfAnnotationCreateOptions options, int expectedObjectNumber, int? expectedParentObjectNumber, PdfLoadOptions? readOptions) {
        PdfDocumentInfo info = ReadAnnotationMetadata(output, readOptions);
        PdfAnnotation? found = info.Annotations.FirstOrDefault(annotation => annotation.ObjectNumber == expectedObjectNumber);
        if (found == null || found.Subtype != options.Subtype || found.PageNumber != options.PageNumber) throw new InvalidOperationException("PDF annotation creation readback failed; the artifact was not returned.");
        if (options.GenerateAppearance && IsAppearanceSubtype(options.Subtype) && !found.HasNormalAppearance) throw new InvalidOperationException("PDF annotation appearance readback failed; the artifact was not returned.");
        if (expectedParentObjectNumber.HasValue && found.Review?.InReplyToObjectNumber != expectedParentObjectNumber) throw new InvalidOperationException("PDF annotation reply relationship readback failed; the artifact was not returned.");
        if (options.ReviewState.HasValue && found.Review?.StandardState != options.ReviewState) throw new InvalidOperationException("PDF annotation review state readback failed; the artifact was not returned.");
        if (options.LinkUri != null && !info.GetLinkAnnotationsByUri(options.LinkUri).Any(link => link.PageNumber == options.PageNumber)) throw new InvalidOperationException("PDF link annotation readback failed; the URI target was not returned.");
        if (options.InteriorColor is not null && !NumbersEqual(found.InteriorColor, options.InteriorColor.Select(ClampColor).ToArray())) throw new InvalidOperationException("PDF annotation interior-color readback failed; the artifact was not returned.");
        if (options.Opacity.HasValue && !NumberEquals(found.Opacity, options.Opacity.Value)) throw new InvalidOperationException("PDF annotation opacity readback failed; the artifact was not returned.");
        if (options.BorderWidth.HasValue && !NumberEquals(found.BorderWidth, options.BorderWidth.Value)) throw new InvalidOperationException("PDF annotation border-width readback failed; the artifact was not returned.");
        if (options.BorderStyle.HasValue && !string.Equals(found.BorderStyle, GetAnnotationBorderStyleDisplayName(options.BorderStyle.Value), StringComparison.Ordinal)) throw new InvalidOperationException("PDF annotation border-style readback failed; the artifact was not returned.");
        if (options.BorderDashPattern is not null && !NumbersEqual(found.BorderDashPattern, options.BorderDashPattern)) throw new InvalidOperationException("PDF annotation dash-pattern readback failed; the artifact was not returned.");
    }

    private static void ValidateUpdatedAnnotation(byte[] output, int expectedObjectNumber, PdfAnnotationUpdateOptions options, PdfLoadOptions? readOptions) {
        PdfAnnotation? found = ReadAnnotationMetadata(output, readOptions).Annotations.FirstOrDefault(annotation => annotation.ObjectNumber == expectedObjectNumber);
        if (found is null) throw new InvalidOperationException("PDF annotation update readback failed; the annotation was not returned.");
        if (options.ReviewState.HasValue && found.Review?.StandardState != options.ReviewState) throw new InvalidOperationException("PDF annotation review state readback failed; the artifact was not returned.");
        if (options.InteriorColor is not null && !NumbersEqual(found.InteriorColor, options.InteriorColor.Select(ClampColor).ToArray())) throw new InvalidOperationException("PDF annotation interior-color readback failed; the artifact was not returned.");
        if (options.Opacity.HasValue && !NumberEquals(found.Opacity, options.Opacity.Value)) throw new InvalidOperationException("PDF annotation opacity readback failed; the artifact was not returned.");
        if (options.BorderWidth.HasValue && !NumberEquals(found.BorderWidth, options.BorderWidth.Value)) throw new InvalidOperationException("PDF annotation border-width readback failed; the artifact was not returned.");
        if (options.BorderStyle.HasValue && !string.Equals(found.BorderStyle, GetAnnotationBorderStyleDisplayName(options.BorderStyle.Value), StringComparison.Ordinal)) throw new InvalidOperationException("PDF annotation border-style readback failed; the artifact was not returned.");
        if (options.BorderDashPattern is not null && !NumbersEqual(found.BorderDashPattern, options.BorderDashPattern)) throw new InvalidOperationException("PDF annotation dash-pattern readback failed; the artifact was not returned.");
    }

    private static bool NumbersEqual(IReadOnlyList<double> left, IReadOnlyList<double> right) => left.Count == right.Count && left.Zip(right, NumberEquals).All(static equal => equal);
    private static bool NumberEquals(double? left, double right) => left.HasValue && Math.Abs(left.Value - right) <= 0.000001D;
    private static bool NumberEquals(double left, double right) => Math.Abs(left - right) <= 0.000001D;
    private static string GetAnnotationBorderStyleDisplayName(PdfAnnotationBorderStyle style) => style.ToString();
}
