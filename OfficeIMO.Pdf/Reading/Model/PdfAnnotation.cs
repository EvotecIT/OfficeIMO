namespace OfficeIMO.Pdf;

/// <summary>Generic page annotation metadata read from a PDF page.</summary>
public sealed class PdfAnnotation {
    internal PdfAnnotation(int? objectNumber, int? pageNumber, string subtype, string? contents, double x1, double y1, double x2, double y2, bool hasNormalAppearance, string? actionType = null, IReadOnlyList<PdfAnnotationAdditionalAction>? additionalActions = null, IReadOnlyList<PdfAnnotationChainedAction>? chainedActions = null) {
        ObjectNumber = objectNumber;
        PageNumber = pageNumber;
        Subtype = subtype;
        Contents = contents;
        X1 = x1;
        Y1 = y1;
        X2 = x2;
        Y2 = y2;
        HasNormalAppearance = hasNormalAppearance;
        ActionType = actionType;
        AdditionalActions = additionalActions ?? Array.Empty<PdfAnnotationAdditionalAction>();
        ChainedActions = chainedActions ?? Array.Empty<PdfAnnotationChainedAction>();
    }

    /// <summary>Indirect annotation object number, when the annotation is referenced indirectly.</summary>
    public int? ObjectNumber { get; }

    /// <summary>One-based page number when known.</summary>
    public int? PageNumber { get; }

    /// <summary>PDF annotation subtype name, for example Link, Text, FreeText, or Highlight.</summary>
    public string Subtype { get; }

    /// <summary>Optional annotation contents metadata from /Contents.</summary>
    public string? Contents { get; }

    /// <summary>Left edge of the annotation rectangle in PDF points.</summary>
    public double X1 { get; }

    /// <summary>Bottom edge of the annotation rectangle in PDF points.</summary>
    public double Y1 { get; }

    /// <summary>Right edge of the annotation rectangle in PDF points.</summary>
    public double X2 { get; }

    /// <summary>Top edge of the annotation rectangle in PDF points.</summary>
    public double Y2 { get; }

    /// <summary>Rectangle width in PDF points.</summary>
    public double Width => X2 - X1;

    /// <summary>Rectangle height in PDF points.</summary>
    public double Height => Y2 - Y1;

    /// <summary>True when the annotation dictionary exposes a normal appearance stream through /AP /N.</summary>
    public bool HasNormalAppearance { get; }

    /// <summary>Primary annotation action type from /A /S, when present.</summary>
    public string? ActionType { get; }

    /// <summary>Additional annotation actions from the /AA dictionary, when present.</summary>
    public IReadOnlyList<PdfAnnotationAdditionalAction> AdditionalActions { get; }

    /// <summary>Chained annotation actions discovered through /Next entries on /A or /AA action dictionaries.</summary>
    public IReadOnlyList<PdfAnnotationChainedAction> ChainedActions { get; }

    /// <summary>True when the annotation has a primary action dictionary.</summary>
    public bool HasAction => !string.IsNullOrEmpty(ActionType);

    /// <summary>True when the annotation has at least one additional action dictionary.</summary>
    public bool HasAdditionalActions => AdditionalActions.Count > 0;

    /// <summary>True when at least one chained /Next action was discovered for this annotation.</summary>
    public bool HasChainedActions => ChainedActions.Count > 0;

    /// <summary>Number of chained /Next actions discovered for this annotation.</summary>
    public int ChainedActionCount => ChainedActions.Count;

    internal PdfAnnotation WithPageNumber(int pageNumber) =>
        PageNumber == pageNumber
            ? this
            : new PdfAnnotation(ObjectNumber, pageNumber, Subtype, Contents, X1, Y1, X2, Y2, HasNormalAppearance, ActionType, AdditionalActions, ChainedActions);
}

/// <summary>Additional annotation action metadata read from an annotation /AA dictionary.</summary>
public sealed class PdfAnnotationAdditionalAction {
    internal PdfAnnotationAdditionalAction(string triggerName, string actionType) {
        TriggerName = triggerName;
        ActionType = actionType;
    }

    /// <summary>PDF additional-action trigger name, for example E, X, D, U, Fo, Bl, PO, PC, PV, or PI.</summary>
    public string TriggerName { get; }

    /// <summary>Action type from the additional action dictionary /S entry.</summary>
    public string ActionType { get; }
}

/// <summary>Chained annotation action metadata read from /Next entries without exposing action payloads.</summary>
public sealed class PdfAnnotationChainedAction {
    internal PdfAnnotationChainedAction(string sourceName, string actionPath, string actionType) {
        SourceName = sourceName;
        ActionPath = actionPath;
        ActionType = actionType;
    }

    /// <summary>Primary action source, A, or page additional-action trigger name such as E, X, D, U, Fo, or Bl.</summary>
    public string SourceName { get; }

    /// <summary>Stable chained-action path, for example A.Next, E.Next, or E.Next.0.</summary>
    public string ActionPath { get; }

    /// <summary>Action type from the chained action dictionary /S entry.</summary>
    public string ActionType { get; }
}
