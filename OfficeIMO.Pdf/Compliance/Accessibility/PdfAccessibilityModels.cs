namespace OfficeIMO.Pdf;

/// <summary>How the accessibility checker decides one condition.</summary>
public enum PdfAccessibilityCoverage {
    /// <summary>The checker evaluates the condition from the file.</summary>
    Automatic,
    /// <summary>The condition needs a person's judgement; it is always reported for review and never as a pass.</summary>
    ManualReview,
    /// <summary>The condition is machine-checkable but this version does not evaluate it yet.</summary>
    NotChecked
}

/// <summary>Outcome of one condition for one document.</summary>
public enum PdfAccessibilityConditionStatus {
    /// <summary>Applicable objects were inspected and none fails the condition.</summary>
    Passed,
    /// <summary>At least one object fails the condition.</summary>
    Failed,
    /// <summary>A person must review the document for this condition.</summary>
    NeedsReview,
    /// <summary>The document contains nothing the condition applies to.</summary>
    NotApplicable,
    /// <summary>Some evidence could not be inspected and no failure was found.</summary>
    Indeterminate,
    /// <summary>The checker does not evaluate this condition yet.</summary>
    NotChecked
}

/// <summary>Severity of one accessibility finding.</summary>
public enum PdfAccessibilityFindingSeverity {
    /// <summary>The object fails a machine-checkable condition.</summary>
    Error,
    /// <summary>A person must review the item; it is not a pass.</summary>
    Review,
    /// <summary>The evidence could not be inspected, so the condition remains open for this object.</summary>
    Indeterminate
}

/// <summary>
/// One PDF/UA-1 condition: a Matterhorn Protocol 1.1 failure condition, or an ISO 14289-1 requirement
/// that veraPDF checks without a Matterhorn number.
/// </summary>
public sealed class PdfAccessibilityCondition {
    internal PdfAccessibilityCondition(string checkpoint, string conditionId, string description, string pdfUaClause,
        PdfAccessibilityCoverage coverage, string[] veraPdfRules, string[] pdfUa2Rules) {
        Checkpoint = checkpoint;
        ConditionId = conditionId;
        Description = description;
        PdfUaClause = pdfUaClause;
        Coverage = coverage;
        VeraPdfRules = Array.AsReadOnly(veraPdfRules);
        PdfUa2Rules = Array.AsReadOnly(pdfUa2Rules);
    }

    /// <summary>Matterhorn checkpoint number, for example <c>28</c> for annotations.</summary>
    public string Checkpoint { get; }

    /// <summary>Matterhorn failure condition, for example <c>28-010</c>, or <c>UA1:7.1-11</c> for an ISO 14289-1 requirement without a Matterhorn number.</summary>
    public string ConditionId { get; }

    /// <summary>True when <see cref="ConditionId"/> is a Matterhorn Protocol failure condition.</summary>
    public bool IsMatterhornCondition => !ConditionId.StartsWith("UA1:", StringComparison.Ordinal);

    /// <summary>What fails the condition.</summary>
    public string Description { get; }

    /// <summary>ISO 14289-1 clause the condition belongs to.</summary>
    public string PdfUaClause { get; }

    /// <summary>Whether the condition is checked automatically, needs review, or is not checked yet.</summary>
    public PdfAccessibilityCoverage Coverage { get; }

    /// <summary>veraPDF PDF/UA-1 profile rules (clause-test) that test the same requirement.</summary>
    public IReadOnlyList<string> VeraPdfRules { get; }

    /// <summary>veraPDF PDF/UA-2 profile rules for the shared requirement; empty when PDF/UA-2 dropped or changed it.</summary>
    public IReadOnlyList<string> PdfUa2Rules { get; }
}

/// <summary>Result of one condition for the inspected document.</summary>
public sealed class PdfAccessibilityConditionResult {
    internal PdfAccessibilityConditionResult(PdfAccessibilityCondition condition, PdfAccessibilityConditionStatus status, int occurrenceCount) {
        Condition = condition;
        Status = status;
        OccurrenceCount = occurrenceCount;
    }

    /// <summary>The condition.</summary>
    public PdfAccessibilityCondition Condition { get; }

    /// <summary>The outcome for this document.</summary>
    public PdfAccessibilityConditionStatus Status { get; }

    /// <summary>Number of failing or indeterminate occurrences, including those beyond the retained findings.</summary>
    public int OccurrenceCount { get; }
}

/// <summary>One located accessibility problem or review item.</summary>
public sealed class PdfAccessibilityFinding {
    internal PdfAccessibilityFinding(PdfAccessibilityCondition condition, string? veraPdfRule, PdfAccessibilityFindingSeverity severity,
        int? pageNumber, int? objectNumber, int? structureElementObjectNumber, string? structureType, string message) {
        Checkpoint = condition.Checkpoint;
        ConditionId = condition.ConditionId;
        VeraPdfRule = veraPdfRule;
        Severity = severity;
        PageNumber = pageNumber;
        ObjectNumber = objectNumber;
        StructureElementObjectNumber = structureElementObjectNumber;
        StructureType = structureType;
        Message = message;
    }

    /// <summary>Matterhorn checkpoint number.</summary>
    public string Checkpoint { get; }

    /// <summary>Matterhorn failure condition or ISO 14289-1 requirement identifier.</summary>
    public string ConditionId { get; }

    /// <summary>The veraPDF PDF/UA-1 rule this finding corresponds to, when one exists.</summary>
    public string? VeraPdfRule { get; }

    /// <summary>Severity of the finding.</summary>
    public PdfAccessibilityFindingSeverity Severity { get; }

    /// <summary>True when a person must decide the item.</summary>
    public bool RequiresHumanReview => Severity == PdfAccessibilityFindingSeverity.Review;

    /// <summary>One-based page number, or null for document-level findings.</summary>
    public int? PageNumber { get; }

    /// <summary>Object number of the annotation, font, XObject, content stream or dictionary concerned, when known.</summary>
    public int? ObjectNumber { get; }

    /// <summary>Object number of the structure element concerned, when known.</summary>
    public int? StructureElementObjectNumber { get; }

    /// <summary>Structure type (/S) of that structure element, as written in the file.</summary>
    public string? StructureType { get; }

    /// <summary>What is wrong and what to change.</summary>
    public string Message { get; }
}

/// <summary>
/// PDF/UA-1 (ISO 14289-1) accessibility check of one exact PDF, organised by the Matterhorn Protocol 1.1
/// failure conditions. It reports what software can decide and lists everything a person must still review;
/// it does not certify conformance.
/// </summary>
public sealed class PdfAccessibilityReport {
    internal PdfAccessibilityReport(string sourceSha256, int pageCount, int? declaredPdfUaPart,
        IReadOnlyList<PdfAccessibilityConditionResult> conditions, IReadOnlyList<PdfAccessibilityFinding> findings) {
        SourceSha256 = sourceSha256;
        PageCount = pageCount;
        DeclaredPdfUaPart = declaredPdfUaPart;
        Conditions = conditions;
        Findings = findings;
    }

    /// <summary>SHA-256 of the exact inspected PDF.</summary>
    public string SourceSha256 { get; }

    /// <summary>Number of pages inspected.</summary>
    public int PageCount { get; }

    /// <summary>The pdfuaid:part value declared in the XMP metadata, or null when none is declared.</summary>
    public int? DeclaredPdfUaPart { get; }

    /// <summary>Every condition in the catalogue, in catalogue order, with its outcome for this document.</summary>
    public IReadOnlyList<PdfAccessibilityConditionResult> Conditions { get; }

    /// <summary>Located findings: failures, indeterminate evidence, and manual review items. At most 100 are kept per condition.</summary>
    public IReadOnlyList<PdfAccessibilityFinding> Findings { get; }

    /// <summary>True when at least one machine-checkable condition fails.</summary>
    public bool HasFailures => Conditions.Any(static result => result.Status == PdfAccessibilityConditionStatus.Failed);

    /// <summary>True when every automatic condition passed or did not apply. Manual review items remain open regardless.</summary>
    public bool PassesAutomaticChecks => Conditions.All(static result =>
        result.Condition.Coverage != PdfAccessibilityCoverage.Automatic ||
        result.Status == PdfAccessibilityConditionStatus.Passed ||
        result.Status == PdfAccessibilityConditionStatus.NotApplicable);

    /// <summary>Conditions that failed.</summary>
    public IReadOnlyList<PdfAccessibilityConditionResult> FailedConditions =>
        Conditions.Where(static result => result.Status == PdfAccessibilityConditionStatus.Failed).ToArray();

    /// <summary>Conditions a person must review.</summary>
    public IReadOnlyList<PdfAccessibilityConditionResult> ManualReviewConditions =>
        Conditions.Where(static result => result.Status == PdfAccessibilityConditionStatus.NeedsReview).ToArray();
}
