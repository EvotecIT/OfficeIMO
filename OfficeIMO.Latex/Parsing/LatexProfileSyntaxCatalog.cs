namespace OfficeIMO.Latex;

internal enum LatexArgumentGroupKind {
    Optional = 0,
    Required
}

internal sealed class LatexCommandSyntaxSignature {
    internal LatexCommandSyntaxSignature(bool allowsStar, params LatexArgumentGroupKind[] arguments) {
        AllowsStar = allowsStar;
        Arguments = arguments ?? Array.Empty<LatexArgumentGroupKind>();
    }

    internal bool AllowsStar { get; }
    internal IReadOnlyList<LatexArgumentGroupKind> Arguments { get; }
}

/// <summary>
/// Bounded argument shapes for commands and environment headers understood by the OfficeIMO profile.
/// Unknown commands remain lossless but use the structural parser's conservative fallback binding.
/// </summary>
internal static class LatexProfileSyntaxCatalog {
    private static readonly LatexCommandSyntaxSignature Zero = new LatexCommandSyntaxSignature(false);
    private static readonly LatexCommandSyntaxSignature OneRequired = new LatexCommandSyntaxSignature(false, LatexArgumentGroupKind.Required);
    private static readonly LatexCommandSyntaxSignature OneOptional = new LatexCommandSyntaxSignature(false, LatexArgumentGroupKind.Optional);
    private static readonly LatexCommandSyntaxSignature OptionalRequired = new LatexCommandSyntaxSignature(false, LatexArgumentGroupKind.Optional, LatexArgumentGroupKind.Required);
    private static readonly LatexCommandSyntaxSignature StarOptionalRequired = new LatexCommandSyntaxSignature(true, LatexArgumentGroupKind.Optional, LatexArgumentGroupKind.Required);
    private static readonly LatexCommandSyntaxSignature Citation = new LatexCommandSyntaxSignature(false,
        LatexArgumentGroupKind.Optional, LatexArgumentGroupKind.Optional, LatexArgumentGroupKind.Required);
    private static readonly LatexCommandSyntaxSignature Begin = new LatexCommandSyntaxSignature(false, LatexArgumentGroupKind.Required);
    private static readonly LatexCommandSyntaxSignature End = new LatexCommandSyntaxSignature(false, LatexArgumentGroupKind.Required);

    private static readonly IReadOnlyDictionary<string, LatexCommandSyntaxSignature> Commands =
        new Dictionary<string, LatexCommandSyntaxSignature>(StringComparer.Ordinal) {
            ["documentclass"] = OptionalRequired,
            ["usepackage"] = OptionalRequired,
            ["title"] = OneRequired,
            ["author"] = OneRequired,
            ["date"] = OneRequired,
            ["maketitle"] = Zero,
            ["part"] = StarOptionalRequired,
            ["chapter"] = StarOptionalRequired,
            ["section"] = StarOptionalRequired,
            ["subsection"] = StarOptionalRequired,
            ["subsubsection"] = StarOptionalRequired,
            ["paragraph"] = StarOptionalRequired,
            ["subparagraph"] = StarOptionalRequired,
            ["textbf"] = OneRequired,
            ["textit"] = OneRequired,
            ["emph"] = OneRequired,
            ["texttt"] = OneRequired,
            ["underline"] = OneRequired,
            ["textsuperscript"] = OneRequired,
            ["textsubscript"] = OneRequired,
            ["sout"] = OneRequired,
            ["newline"] = Zero,
            ["linebreak"] = OneOptional,
            ["label"] = OneRequired,
            ["ref"] = OneRequired,
            ["pageref"] = OneRequired,
            ["autoref"] = OneRequired,
            ["eqref"] = OneRequired,
            ["cite"] = Citation,
            ["citep"] = Citation,
            ["citet"] = Citation,
            ["nocite"] = OneRequired,
            ["includegraphics"] = StarOptionalRequired,
            ["caption"] = OptionalRequired,
            ["item"] = new LatexCommandSyntaxSignature(false, LatexArgumentGroupKind.Optional),
            ["footnote"] = OptionalRequired,
            ["url"] = OneRequired,
            ["href"] = new LatexCommandSyntaxSignature(false, LatexArgumentGroupKind.Required, LatexArgumentGroupKind.Required),
            ["begin"] = Begin,
            ["end"] = End,
            ["newcommand"] = new LatexCommandSyntaxSignature(true,
                LatexArgumentGroupKind.Required, LatexArgumentGroupKind.Optional, LatexArgumentGroupKind.Optional, LatexArgumentGroupKind.Required),
            ["renewcommand"] = new LatexCommandSyntaxSignature(true,
                LatexArgumentGroupKind.Required, LatexArgumentGroupKind.Optional, LatexArgumentGroupKind.Optional, LatexArgumentGroupKind.Required),
            ["providecommand"] = new LatexCommandSyntaxSignature(true,
                LatexArgumentGroupKind.Required, LatexArgumentGroupKind.Optional, LatexArgumentGroupKind.Optional, LatexArgumentGroupKind.Required),
            ["newtheorem"] = new LatexCommandSyntaxSignature(true,
                LatexArgumentGroupKind.Required, LatexArgumentGroupKind.Optional, LatexArgumentGroupKind.Required, LatexArgumentGroupKind.Optional),
            ["bibliography"] = OneRequired,
            ["bibliographystyle"] = OneRequired,
            ["multicolumn"] = new LatexCommandSyntaxSignature(false,
                LatexArgumentGroupKind.Required, LatexArgumentGroupKind.Required, LatexArgumentGroupKind.Required),
            ["multirow"] = new LatexCommandSyntaxSignature(false,
                LatexArgumentGroupKind.Optional, LatexArgumentGroupKind.Required, LatexArgumentGroupKind.Optional,
                LatexArgumentGroupKind.Required, LatexArgumentGroupKind.Optional, LatexArgumentGroupKind.Required),
            ["hline"] = Zero,
            ["toprule"] = Zero,
            ["midrule"] = Zero,
            ["bottomrule"] = Zero
        };

    internal static LatexCommandSyntaxSignature? GetCommand(string name) =>
        Commands.TryGetValue(name, out LatexCommandSyntaxSignature? signature) ? signature : null;

    internal static LatexCommandSyntaxSignature GetEnvironmentBegin(string environmentName) {
        if (string.Equals(environmentName, "tabular", StringComparison.Ordinal)) {
            return new LatexCommandSyntaxSignature(false,
                LatexArgumentGroupKind.Required, LatexArgumentGroupKind.Optional, LatexArgumentGroupKind.Required);
        }
        if (string.Equals(environmentName, "figure", StringComparison.Ordinal) ||
            string.Equals(environmentName, "table", StringComparison.Ordinal) ||
            IsTheoremEnvironment(environmentName)) {
            return new LatexCommandSyntaxSignature(false, LatexArgumentGroupKind.Required, LatexArgumentGroupKind.Optional);
        }
        return Begin;
    }

    internal static LatexCommandSyntaxSignature EnvironmentEnd => End;

    private static bool IsTheoremEnvironment(string name) =>
        name == "theorem" || name == "lemma" || name == "proposition" || name == "corollary" ||
        name == "definition" || name == "remark" || name == "proof";
}
