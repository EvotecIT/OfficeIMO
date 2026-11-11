using System.Collections.ObjectModel;

namespace OfficeIMO.Word;

/// <summary>Severity assigned to a Word accessibility finding.</summary>
public enum WordAccessibilitySeverity {
    /// <summary>Useful inspection context.</summary>
    Info,
    /// <summary>Potential accessibility risk that deserves review.</summary>
    Warning,
    /// <summary>Policy violation that should fail the selected CI gate.</summary>
    Error
}

/// <summary>Built-in Word accessibility policy profiles.</summary>
public enum WordAccessibilityPolicyProfile {
    /// <summary>Practical policy: structural risks such as skipped headings and floating objects are warnings.</summary>
    Default,
    /// <summary>Stricter policy: structural risks are errors. Heuristic findings stay warnings.</summary>
    Strict
}

/// <summary>Options controlling Word accessibility inspection.</summary>
public sealed class WordAccessibilityOptions {
    private double _minimumTextContrastRatio = 4.5D;
    private double _minimumLargeTextContrastRatio = 3D;

    /// <summary>Selected policy profile.</summary>
    public WordAccessibilityPolicyProfile Profile { get; set; } = WordAccessibilityPolicyProfile.Default;
    /// <summary>Whether the document must define a title. Defaults to true.</summary>
    public bool RequireDocumentTitle { get; set; } = true;
    /// <summary>Whether the document must define a default text language. Defaults to true.</summary>
    public bool RequireLanguage { get; set; } = true;
    /// <summary>Whether images, shapes, charts and SmartArt need alternative text unless marked decorative. Defaults to true.</summary>
    public bool RequireAlternativeText { get; set; } = true;
    /// <summary>Whether data tables need a header row. Defaults to true.</summary>
    public bool RequireTableHeaders { get; set; } = true;
    /// <summary>Whether heading levels are checked for skips, and long documents for missing headings. Defaults to true.</summary>
    public bool CheckHeadings { get; set; } = true;
    /// <summary>Whether hyperlinks need meaningful visible text. Defaults to true.</summary>
    public bool CheckMeaningfulLinks { get; set; } = true;
    /// <summary>Whether text with an explicit colour is checked against its shading or page background. Defaults to true.</summary>
    public bool CheckContrast { get; set; } = true;
    /// <summary>Whether objects that are not inline with text are reported. Defaults to true.</summary>
    public bool CheckFloatingObjects { get; set; } = true;
    /// <summary>Whether consecutive blank paragraphs used for spacing are reported. Defaults to true.</summary>
    public bool CheckBlankParagraphs { get; set; } = true;

    /// <summary>Minimum contrast ratio for normal text. Defaults to the WCAG AA value 4.5.</summary>
    public double MinimumTextContrastRatio {
        get => _minimumTextContrastRatio;
        set => _minimumTextContrastRatio = ValidateContrast(value, nameof(MinimumTextContrastRatio));
    }

    /// <summary>Minimum contrast ratio for large text (18 pt, or 14 pt bold). Defaults to the WCAG AA value 3.</summary>
    public double MinimumLargeTextContrastRatio {
        get => _minimumLargeTextContrastRatio;
        set => _minimumLargeTextContrastRatio = ValidateContrast(value, nameof(MinimumLargeTextContrastRatio));
    }

    private static double ValidateContrast(double value, string name) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value < 1D || value > 21D) {
            throw new ArgumentOutOfRangeException(name, "Contrast ratio must be between 1 and 21.");
        }
        return value;
    }
}

/// <summary>Safe automatic fixes that <see cref="WordDocument.FixAccessibility"/> applies.</summary>
public sealed class WordAccessibilityFixOptions {
    /// <summary>Title to set when the document has none.</summary>
    public string? Title { get; set; }
    /// <summary>BCP 47 language, such as <c>en-US</c>, to set as the document default when none is defined.</summary>
    public string? Language { get; set; }
    /// <summary>Whether to mark the first row of data tables without a header row as the header row. Tables that look like layout tables are left alone.</summary>
    public bool MarkTableHeaderRows { get; set; }
}

/// <summary>One stable, machine-readable Word accessibility finding.</summary>
public sealed class WordAccessibilityFinding {
    internal WordAccessibilityFinding(WordAccessibilitySeverity severity, string code, string message, string location,
        int? paragraphIndex = null, string? objectName = null, bool isHeuristic = false,
        double? measuredValue = null, double? requiredValue = null) {
        Severity = severity;
        Code = code;
        Message = message;
        Location = location;
        ParagraphIndex = paragraphIndex;
        ObjectName = objectName;
        IsHeuristic = isHeuristic;
        MeasuredValue = measuredValue;
        RequiredValue = requiredValue;
    }

    /// <summary>Finding severity.</summary>
    public WordAccessibilitySeverity Severity { get; }
    /// <summary>Stable diagnostic code suitable for CI policies, shared with the PowerPoint checker where the rule is the same.</summary>
    public string Code { get; }
    /// <summary>Human-readable explanation.</summary>
    public string Message { get; }
    /// <summary>Where the finding is, such as <c>Paragraph 12</c> or <c>Table 2</c>.</summary>
    public string Location { get; }
    /// <summary>Zero-based index of the affected paragraph among the body's paragraphs, including those in tables, when applicable.</summary>
    public int? ParagraphIndex { get; }
    /// <summary>Name of the affected image, shape, chart or SmartArt object, when applicable.</summary>
    public string? ObjectName { get; }
    /// <summary>Whether the finding comes from a heuristic, such as guessing that a table is used for layout.</summary>
    public bool IsHeuristic { get; }
    /// <summary>Measured value, such as a contrast ratio.</summary>
    public double? MeasuredValue { get; }
    /// <summary>Value the policy requires.</summary>
    public double? RequiredValue { get; }

    /// <inheritdoc />
    public override string ToString() => Severity + " " + Code + " at " + Location + ": " + Message;
}

/// <summary>Structured Word accessibility inspection result.</summary>
public sealed class WordAccessibilityReport {
    internal WordAccessibilityReport(WordAccessibilityPolicyProfile profile, IList<WordAccessibilityFinding> findings) {
        Profile = profile;
        Findings = new ReadOnlyCollection<WordAccessibilityFinding>(findings);
    }

    /// <summary>Policy profile used for inspection.</summary>
    public WordAccessibilityPolicyProfile Profile { get; }
    /// <summary>Findings in document order, after the document-level findings.</summary>
    public IReadOnlyList<WordAccessibilityFinding> Findings { get; }
    /// <summary>Number of error findings.</summary>
    public int ErrorCount => Findings.Count(finding => finding.Severity == WordAccessibilitySeverity.Error);
    /// <summary>Number of warning findings.</summary>
    public int WarningCount => Findings.Count(finding => finding.Severity == WordAccessibilitySeverity.Warning);
    /// <summary>Whether no policy errors were found.</summary>
    public bool IsSuccessful => ErrorCount == 0;

    /// <summary>Throws when the report contains errors, or warnings when requested.</summary>
    public WordAccessibilityReport EnsureCompliant(bool includeWarnings = false) {
        if (ErrorCount > 0 || includeWarnings && WarningCount > 0) throw new WordAccessibilityException(this, includeWarnings);
        return this;
    }
}

/// <summary>Raised when an accessibility policy rejects a Word document.</summary>
public sealed class WordAccessibilityException : InvalidOperationException {
    internal WordAccessibilityException(WordAccessibilityReport report, bool includeWarnings)
        : base("Word accessibility inspection found " + report.ErrorCount + " error(s) and " + report.WarningCount + " warning(s).") {
        Report = report;
        IncludedWarnings = includeWarnings;
    }

    /// <summary>Report that caused the policy failure.</summary>
    public WordAccessibilityReport Report { get; }
    /// <summary>Whether warnings participated in the failure threshold.</summary>
    public bool IncludedWarnings { get; }
}
