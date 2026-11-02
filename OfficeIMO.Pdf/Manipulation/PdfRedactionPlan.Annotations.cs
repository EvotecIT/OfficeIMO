namespace OfficeIMO.Pdf;

public sealed partial class PdfRedactionPlan {
    /// <summary>Projects only authorized metadata and annotation-edge cleanup throughout each fingerprint graph.</summary>
    private static Func<PdfDictionary, string, bool> CreateIdentityEntryFilter(PdfReadDocument document,
        IReadOnlyList<PdfRedactionArea> areas, HashSet<int> removedAnnotations) {
        var metadata = PdfRedactionApplier.SearchMetadataHolders(document.Objects, document.CatalogDictionary, document.TrailerRaw,
            document.Pages.Select(static page => page.ObjectNumber).ToArray(), PdfRedactionApplier.GetSearchCriteria(areas), document.ReadOptions.Limits);
        var metadataOwners = new HashSet<PdfDictionary>(metadata.Owners);
        HashSet<PdfDictionary> annotations = PdfRedactionApplier.GetAnnotationOwners(document.Objects);
        return (dictionary, key) => {
            if (key == "Metadata" && metadataOwners.Contains(dictionary)) return false;
            if (!annotations.Contains(dictionary)) return true;
            if (key is "IRT" or "RT" && dictionary.Items.TryGetValue("IRT", out PdfObject? reply) && reply is PdfReference target && removedAnnotations.Contains(target.ObjectNumber)) return false;
            return key != "Popup" || !dictionary.Items.TryGetValue("Popup", out PdfObject? popup) || popup is not PdfReference popupTarget || !removedAnnotations.Contains(popupTarget.ObjectNumber);
        };
    }

    /// <summary>Identifies annotations and their popups authorized for removal by the reviewed geometry or search.</summary>
    private static HashSet<int> FindRemovedAnnotations(PdfReadDocument document, IReadOnlyList<PdfRedactionArea> areas) {
        var removed = new HashSet<int>();
        var annotations = new List<PdfAnnotation>();
        for (int index = 0; index < document.Pages.Count; index++) {
            PdfReadPage page = document.Pages[index];
            PdfRedactionArea[] pageAreas = areas.Where(area => area.PageNumber == index + 1).ToArray();
            foreach (PdfAnnotation annotation in page.GetAnnotationsForContentSafety()) {
                annotations.Add(annotation);
                if (!annotation.ObjectNumber.HasValue || annotation.Subtype == "Widget") continue;
                if (!(annotation.HasReadableRectangle && IntersectsReviewedUnderlay(pageAreas, annotation.X1, annotation.Y1, annotation.Width, annotation.Height)) &&
                    !PdfRedactionApplier.CarriesSearchedText(document.Objects, page.ObjectNumber, annotation, pageAreas)) continue;
                removed.Add(annotation.ObjectNumber.Value);
                if (annotation.SourceDictionary?.Items.TryGetValue("Popup", out PdfObject? popup) == true && popup is PdfReference reference) removed.Add(reference.ObjectNumber);
            }
        }
        foreach (PdfAnnotation annotation in annotations) {
            if (annotation.Subtype == "Popup" && annotation.ObjectNumber.HasValue &&
                annotation.SourceDictionary?.Items.TryGetValue("Parent", out PdfObject? parent) == true && parent is PdfReference target && removed.Contains(target.ObjectNumber))
                removed.Add(annotation.ObjectNumber.Value);
        }
        return removed;
    }
    private static void AppendUnredactedAnnotationIdentity(
        System.Text.StringBuilder identity,
        PdfReadDocument document,
        PdfReadPage page,
        IReadOnlyList<PdfRedactionArea> pageAreas,
        IReadOnlyDictionary<int, string> stablePageReferences,
        HashSet<string> affectedFields, HashSet<int> removedAnnotations, Func<PdfDictionary, string, bool>? includeEntry = null) {
        IReadOnlyList<PdfAnnotation> annotations = page.GetAnnotationsForContentSafety();
        for (int i = 0; i < annotations.Count; i++) {
            PdfAnnotation annotation = annotations[i];
            if (annotation.ObjectNumber.HasValue && removedAnnotations.Contains(annotation.ObjectNumber.Value)) continue;
            bool detachedReply = annotation.SourceDictionary?.Items.TryGetValue("IRT", out PdfObject? reply) == true &&
                reply is PdfReference target && removedAnnotations.Contains(target.ObjectNumber);
            if (annotation.HasReadableRectangle &&
                IntersectsReviewedUnderlay(pageAreas, annotation.X1, annotation.Y1, annotation.Width, annotation.Height)) continue;
            if (affectedFields.Count > 0 && annotation.SourceDictionary is PdfDictionary widget &&
                affectedFields.Contains(PdfRedactionApplier.GetFieldName(document.Objects, widget))) continue;
            // Carrying the searched text takes it out of the kept set wherever it sits.
            if (PdfRedactionApplier.CarriesSearchedText(document.Objects, page.ObjectNumber, annotation, pageAreas)) continue;

            identity.Append("|A:");
            AppendIdentityString(identity, annotation.Subtype);
            identity.Append(':').Append(FormatIdentityNumber(annotation.X1))
                .Append(',').Append(FormatIdentityNumber(annotation.Y1))
                .Append(',').Append(FormatIdentityNumber(annotation.X2))
                .Append(',').Append(FormatIdentityNumber(annotation.Y2));
            AppendIdentityString(identity, annotation.Contents);
            AppendIdentityString(identity, annotation.Name);
            AppendIdentityString(identity, annotation.Title);
            AppendIdentityString(identity, annotation.ActionType);
            identity.Append(':').Append(annotation.Flags?.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? "null")
                .Append(':').Append(annotation.HasNormalAppearance ? '1' : '0');
            AppendIdentityString(identity, annotation.AppearanceState);
            AppendIdentityString(identity, annotation.DefaultAppearance);
            AppendIdentityString(identity, annotation.DefaultStyle);
            AppendIdentityString(identity, annotation.RichContents);
            AppendIdentityString(identity, annotation.RichContentsPlainText);
            AppendIdentityNullableNumber(identity, annotation.EffectiveFontSize);
            AppendIdentityPdfColor(identity, annotation.EffectiveTextColor);
            identity.Append(':').Append(annotation.EffectiveTextAlign.HasValue
                ? ((int)annotation.EffectiveTextAlign.Value).ToString(System.Globalization.CultureInfo.InvariantCulture)
                : "null");
            AppendIdentityNullableNumber(identity, annotation.Opacity);
            AppendIdentityNullableNumber(identity, annotation.BorderWidth);
            AppendIdentityString(identity, annotation.BorderStyle);
            AppendIdentityNumbers(identity, annotation.BorderDashPattern);
            AppendIdentityString(identity, annotation.BorderEffectStyle);
            AppendIdentityNullableNumber(identity, annotation.BorderEffectIntensity);
            AppendIdentityNumbers(identity, annotation.RectangleDifferences);
            AppendIdentityNumbers(identity, annotation.CalloutLine);
            AppendIdentityString(identity, annotation.CalloutLineEnding);
            AppendIdentityString(identity, annotation.LineStartEnding);
            AppendIdentityString(identity, annotation.LineEndEnding);
            identity.Append(":appearance:");
            PdfRedactionImageIdentity.AppendObjectGraph(identity, annotation.NormalAppearanceObject, document.Objects, includeEntry);
            AppendIdentityNumbers(identity, annotation.Color);
            AppendIdentityNumbers(identity, annotation.InteriorColor);
            AppendIdentityNumbers(identity, annotation.QuadPoints);
            AppendIdentityNumbers(identity, annotation.LineCoordinates);
            AppendIdentityNumbers(identity, annotation.Vertices);
            for (int pathIndex = 0; pathIndex < annotation.InkList.Count; pathIndex++) {
                AppendIdentityNumbers(identity, annotation.InkList[pathIndex]);
            }
            for (int actionIndex = 0; actionIndex < annotation.AdditionalActions.Count; actionIndex++) {
                PdfAnnotationAdditionalAction action = annotation.AdditionalActions[actionIndex];
                AppendIdentityString(identity, action.TriggerName);
                AppendIdentityString(identity, action.ActionType);
            }
            for (int actionIndex = 0; actionIndex < annotation.ChainedActions.Count; actionIndex++) {
                PdfAnnotationChainedAction action = annotation.ChainedActions[actionIndex];
                AppendIdentityString(identity, action.SourceName);
                AppendIdentityString(identity, action.ActionPath);
                AppendIdentityString(identity, action.ActionType);
            }
            if (annotation.Review != null && (!detachedReply || annotation.Review.State is not null || annotation.Review.StateModel is not null || annotation.Review.Subject is not null || annotation.Review.Intent is not null)) {
                AppendIdentityString(identity, detachedReply ? null : annotation.Review.ReplyType);
                AppendIdentityString(identity, annotation.Review.State);
                AppendIdentityString(identity, annotation.Review.StateModel);
                AppendIdentityString(identity, annotation.Review.Subject);
                AppendIdentityString(identity, annotation.Review.Intent);
            }
            if (annotation.SourceDictionary != null) {
                identity.Append(":dictionary:")
                    .Append(PdfRedactionAnnotationIdentity.Compute(
                        annotation.SourceDictionary,
                        document.Objects,
                        stablePageReferences, includeEntry));
            }
        }
    }

}
