using System.Text.RegularExpressions;

namespace OfficeIMO.Pdf;

/// <summary>Explicit criteria used to derive reviewable redaction rectangles from logical content.</summary>
public sealed class PdfRedactionSearchOptions {
    private readonly List<string> _literalText = new List<string>();
    private readonly List<string> _regularExpressions = new List<string>();
    private readonly List<string> _formFieldNames = new List<string>();
    private readonly HashSet<PdfLogicalElementKind> _logicalElementKinds = new HashSet<PdfLogicalElementKind>();

    /// <summary>Case-sensitive literal matching when true.</summary>
    public bool MatchCase { get; set; }

    /// <summary>Timeout applied independently to every regular-expression match.</summary>
    public TimeSpan RegexTimeout { get; set; } = TimeSpan.FromSeconds(2);

    /// <summary>Regular-expression options. CultureInvariant is recommended for reproducible plans.</summary>
    public RegexOptions RegexOptions { get; set; } = RegexOptions.CultureInvariant;

    /// <summary>Literal text criteria.</summary>
    public IList<string> LiteralText => _literalText;

    /// <summary>Regular-expression criteria.</summary>
    public IList<string> RegularExpressions => _regularExpressions;

    /// <summary>Fully qualified AcroForm field names whose widgets should be selected.</summary>
    public IList<string> FormFieldNames => _formFieldNames;

    /// <summary>Logical element kinds whose text blocks should be selected.</summary>
    public ISet<PdfLogicalElementKind> LogicalElementKinds => _logicalElementKinds;

    /// <summary>Adds literal text criteria.</summary>
    public PdfRedactionSearchOptions AddLiteral(params string[] values) { AddValues(_literalText, values, nameof(values)); return this; }

    /// <summary>Adds bounded regular-expression criteria.</summary>
    public PdfRedactionSearchOptions AddRegex(params string[] patterns) { AddValues(_regularExpressions, patterns, nameof(patterns)); return this; }

    /// <summary>Adds fully qualified AcroForm field names.</summary>
    public PdfRedactionSearchOptions AddFormField(params string[] fieldNames) { AddValues(_formFieldNames, fieldNames, nameof(fieldNames)); return this; }

    /// <summary>Adds logical element kinds such as Heading, ListItem, or TextBlock.</summary>
    public PdfRedactionSearchOptions AddLogicalKind(params PdfLogicalElementKind[] kinds) { Guard.NotNull(kinds, nameof(kinds)); for (int i = 0; i < kinds.Length; i++) _logicalElementKinds.Add(kinds[i]); return this; }

    private static void AddValues(List<string> target, string[] values, string parameterName) { Guard.NotNull(values, parameterName); for (int i = 0; i < values.Length; i++) { Guard.NotNullOrWhiteSpace(values[i], parameterName); target.Add(values[i]); } }
}
