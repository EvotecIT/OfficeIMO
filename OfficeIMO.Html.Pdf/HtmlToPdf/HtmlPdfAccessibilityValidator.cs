using System.Collections.Generic;
using System.Linq;
using System.Threading;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

/// <summary>Severity of one bounded tagged-PDF accessibility validation issue.</summary>
public enum HtmlPdfAccessibilityIssueSeverity {
    /// <summary>The artifact remains structurally usable but carries an accessibility warning.</summary>
    Warning,
    /// <summary>The artifact violates the validator's required structural contract.</summary>
    Error
}

/// <summary>One stable issue emitted by <see cref="HtmlPdfAccessibilityValidator"/>.</summary>
public sealed class HtmlPdfAccessibilityIssue {
    internal HtmlPdfAccessibilityIssue(string code, string message, HtmlPdfAccessibilityIssueSeverity severity) {
        Code = code;
        Message = message;
        Severity = severity;
    }

    /// <summary>Stable machine-readable issue code.</summary>
    public string Code { get; }
    /// <summary>Human-readable issue explanation.</summary>
    public string Message { get; }
    /// <summary>Issue severity.</summary>
    public HtmlPdfAccessibilityIssueSeverity Severity { get; }
}

/// <summary>Result of bounded structural accessibility validation for an HTML-derived PDF.</summary>
public sealed class HtmlPdfAccessibilityValidationResult {
    internal HtmlPdfAccessibilityValidationResult(IReadOnlyList<HtmlPdfAccessibilityIssue> issues) {
        Issues = Array.AsReadOnly((issues ?? throw new ArgumentNullException(nameof(issues))).ToArray());
    }

    /// <summary>True when the artifact satisfies every required validator check.</summary>
    public bool IsValid => Issues.All(issue => issue.Severity != HtmlPdfAccessibilityIssueSeverity.Error);
    /// <summary>Stable ordered validation issues.</summary>
    public IReadOnlyList<HtmlPdfAccessibilityIssue> Issues { get; }
}

/// <summary>
/// Validates the deterministic tagged-PDF structure emitted by the direct HTML renderer.
/// This is a bounded structural contract, not a claim of full PDF/UA certification.
/// </summary>
public static class HtmlPdfAccessibilityValidator {
    /// <summary>
    /// Validates both the bounded source accessibility contract and the generated PDF structure.
    /// Use this overload when source-only omissions, such as an image without an accessible name,
    /// must remain diagnosable after conversion.
    /// </summary>
    /// <remarks>
    /// Assumes default print CSS without caller or externally resolved stylesheets. For generated
    /// documents, use <see cref="Validate(HtmlPdfRenderRequestResult, CancellationToken)"/> to inspect
    /// the operation's retained source evidence instead of recomputing source styles.
    /// </remarks>
    public static HtmlPdfAccessibilityValidationResult Validate(HtmlConversionDocument sourceDocument, byte[] pdfBytes) {
        if (sourceDocument == null) throw new ArgumentNullException(nameof(sourceDocument));
        HtmlPdfAccessibilityValidationResult artifactResult = Validate(pdfBytes);
        var issues = new List<HtmlPdfAccessibilityIssue>(artifactResult.Issues);
        var document = sourceDocument.CreateNativeDocumentForConversion(HtmlCssMediaContext.Print);
        var styles = HtmlComputedStyleEngine.Compute(document, HtmlCssMediaContext.Print);
        AddMissingImageNames(issues, HtmlAccessibilitySourceInspection.CountImagesMissingNames(document, styles));
        return new HtmlPdfAccessibilityValidationResult(issues.AsReadOnly());
    }

    /// <summary>
    /// Validates the PDF and source image-name evidence retained by its render operation.
    /// Uses the effective media, viewport, caller and resolved stylesheets without fetching resources again.
    /// </summary>
    /// <remarks>
    /// Source checks cover CSS-visible source images before output page selection and clipping.
    /// Validate before modifying the returned PDF. This bounded check does not certify PDF/UA conformance.
    /// </remarks>
    public static HtmlPdfAccessibilityValidationResult Validate(HtmlPdfRenderRequestResult rendered,
        CancellationToken cancellationToken = default) {
        if (rendered == null) throw new ArgumentNullException(nameof(rendered));
        cancellationToken.ThrowIfCancellationRequested();
        int missingNames = rendered.RenderResult.Document.SourceImagesMissingNames
            ?? throw new InvalidOperationException("The render result does not retain source accessibility evidence.");
        HtmlPdfAccessibilityValidationResult artifactResult = Validate(rendered.ToBytes(cancellationToken));
        cancellationToken.ThrowIfCancellationRequested();
        var issues = new List<HtmlPdfAccessibilityIssue>(artifactResult.Issues);
        AddMissingImageNames(issues, missingNames);
        return new HtmlPdfAccessibilityValidationResult(issues.AsReadOnly());
    }

    private static void AddMissingImageNames(ICollection<HtmlPdfAccessibilityIssue> issues, int count) {
        for (int index = 0; index < count; index++)
            AddError(issues, "HtmlPdfAccessibilityImageNameMissing", "A rendered source image does not declare an accessible name or an explicit decorative alternative.");
    }

    /// <summary>Validates language, tag roots, parent-tree evidence, hierarchy links, marked content, tables, lists, links, and figure alternate text.</summary>
    public static HtmlPdfAccessibilityValidationResult Validate(byte[] pdfBytes) {
        if (pdfBytes == null) throw new ArgumentNullException(nameof(pdfBytes));
        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(pdfBytes);
        var issues = new List<HtmlPdfAccessibilityIssue>();
        if (string.IsNullOrWhiteSpace(info.CatalogLanguage)) {
            AddError(issues, "HtmlPdfAccessibilityLanguageMissing", "The PDF catalog does not declare a document language.");
        }

        PdfCore.PdfTaggedContentInfo? tagged = info.TaggedContent;
        if (tagged == null) {
            AddError(issues, "HtmlPdfAccessibilityTagsMissing", "The PDF does not expose a readable structure tree.");
            return new HtmlPdfAccessibilityValidationResult(issues.AsReadOnly());
        }
        if (tagged.Marked != true) AddError(issues, "HtmlPdfAccessibilityMarkedMissing", "The PDF catalog does not declare marked content.");
        if (!tagged.HasDocumentStructureElement) AddError(issues, "HtmlPdfAccessibilityDocumentRootMissing", "The structure tree does not contain a Document element.");
        if (tagged.ParentTreeEntryCount == 0) AddError(issues, "HtmlPdfAccessibilityParentTreeMissing", "The structure tree does not expose parent-tree entries.");

        var elements = new Dictionary<int, PdfCore.PdfStructureElementInfo>();
        foreach (PdfCore.PdfStructureElementInfo element in tagged.StructureElements) {
            if (elements.ContainsKey(element.ObjectNumber)) {
                AddError(issues, "HtmlPdfAccessibilityDuplicateStructureObject", "Multiple structure elements use the same PDF object number.");
            } else {
                elements.Add(element.ObjectNumber, element);
            }
        }
        foreach (int rootObjectNumber in tagged.RootElementObjectNumbers) {
            if (!elements.ContainsKey(rootObjectNumber)) {
                AddError(issues, "HtmlPdfAccessibilityDanglingRoot", "The structure tree references an unreadable root element.");
            }
        }
        foreach (PdfCore.PdfStructureElementInfo element in tagged.StructureElements) {
            foreach (int childObjectNumber in element.ChildElementObjectNumbers) {
                if (!elements.TryGetValue(childObjectNumber, out PdfCore.PdfStructureElementInfo? child)) {
                    AddError(issues, "HtmlPdfAccessibilityDanglingChild", "A structure element references an unreadable child element.");
                } else if (child.ParentObjectNumber != element.ObjectNumber) {
                    AddError(issues, "HtmlPdfAccessibilityParentMismatch", "A structure child does not point back to its declared parent.");
                }
            }
        }

        if (tagged.StructureElements.Any(IsTextBearingStructure) && !tagged.HasMarkedContentReferences) {
            AddError(issues, "HtmlPdfAccessibilityMarkedContentMissing", "Text-bearing structure elements do not reference marked page content.");
        }
        if (!tagged.FiguresHaveAlternateText) {
            AddError(issues, "HtmlPdfAccessibilityFigureAltMissing", "One or more Figure elements do not declare alternate text.");
        }

        ValidateTableHierarchy(tagged.StructureElements, elements, issues);
        ValidateListHierarchy(tagged.StructureElements, elements, issues);
        ValidateLinks(tagged.StructureElements, issues);
        return new HtmlPdfAccessibilityValidationResult(issues.AsReadOnly());
    }

    private static void ValidateTableHierarchy(
        IReadOnlyList<PdfCore.PdfStructureElementInfo> elements,
        IReadOnlyDictionary<int, PdfCore.PdfStructureElementInfo> byObjectNumber,
        ICollection<HtmlPdfAccessibilityIssue> issues) {
        foreach (PdfCore.PdfStructureElementInfo cell in elements.Where(element => element.StructureType is "TH" or "TD")) {
            if (!cell.ParentObjectNumber.HasValue ||
                !byObjectNumber.TryGetValue(cell.ParentObjectNumber.Value, out PdfCore.PdfStructureElementInfo? parent) ||
                parent.StructureType != "TR") {
                AddError(issues, "HtmlPdfAccessibilityTableCellParentInvalid", "A table cell is not parented by a TR structure element.");
            }
        }
    }

    private static void ValidateListHierarchy(
        IReadOnlyList<PdfCore.PdfStructureElementInfo> elements,
        IReadOnlyDictionary<int, PdfCore.PdfStructureElementInfo> byObjectNumber,
        ICollection<HtmlPdfAccessibilityIssue> issues) {
        foreach (PdfCore.PdfStructureElementInfo item in elements.Where(element => element.StructureType == "LI")) {
            string? parentType = item.ParentObjectNumber.HasValue && byObjectNumber.TryGetValue(item.ParentObjectNumber.Value, out PdfCore.PdfStructureElementInfo? parent)
                ? parent.StructureType
                : null;
            var childTypes = item.ChildElementObjectNumbers
                .Select(child => byObjectNumber.TryGetValue(child, out PdfCore.PdfStructureElementInfo? element) ? element.StructureType : null)
                .ToList();
            if (parentType != "L" || !childTypes.Contains("Lbl") || !childTypes.Contains("LBody")) {
                AddError(issues, "HtmlPdfAccessibilityListHierarchyInvalid", "A list item does not expose the required L, Lbl, and LBody hierarchy.");
            }
        }
    }

    private static void ValidateLinks(
        IReadOnlyList<PdfCore.PdfStructureElementInfo> elements,
        ICollection<HtmlPdfAccessibilityIssue> issues) {
        foreach (PdfCore.PdfStructureElementInfo link in elements.Where(element => element.StructureType == "Link")) {
            if (link.MarkedContentReferenceCount == 0 && link.ObjectReferenceCount == 0 && link.ChildElementObjectNumbers.Count == 0) {
                AddError(issues, "HtmlPdfAccessibilityLinkTargetMissing", "A Link structure element has no marked content, object reference, or structured child.");
            }
        }
    }

    private static bool IsTextBearingStructure(PdfCore.PdfStructureElementInfo element) =>
        element.StructureType is "P" or "Span" or "H" or "H1" or "H2" or "H3" or "H4" or "H5" or "H6" or "Lbl" or "LBody" or "TH" or "TD" or "Link";

    private static void AddError(ICollection<HtmlPdfAccessibilityIssue> issues, string code, string message) =>
        issues.Add(new HtmlPdfAccessibilityIssue(code, message, HtmlPdfAccessibilityIssueSeverity.Error));
}
