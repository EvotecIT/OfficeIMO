namespace OfficeIMO.Latex;

/// <summary>Parsed LaTeX document with lossless syntax and bounded profile semantics.</summary>
public sealed class LatexDocument {
    private readonly IReadOnlyList<LatexToken> _tokens;
    private readonly IReadOnlyList<LatexDiagnostic> _diagnostics;
    private readonly IReadOnlyList<LatexCommand> _commands;
    private readonly IReadOnlyList<LatexEnvironment> _environments;
    private readonly IReadOnlyList<LatexMath> _math;
    private readonly IReadOnlyList<LatexHeading> _headings;
    private readonly IReadOnlyList<LatexParagraph> _paragraphs;
    private readonly IReadOnlyList<LatexList> _lists;
    private readonly IReadOnlyList<LatexFigure> _figures;
    private readonly IReadOnlyList<LatexTable> _tables;
    private readonly IReadOnlyList<LatexCitation> _citations;
    private readonly IReadOnlyList<LatexReference> _references;
    private readonly IReadOnlyList<LatexLabel> _labels;
    private readonly IReadOnlyList<LatexTheorem> _theorems;
    private readonly IReadOnlyList<LatexMacroDefinition> _macroDefinitions;
    private readonly LatexParseOptions _options;

    internal LatexDocument(
        LatexSourceText source,
        LatexSyntaxTree syntaxTree,
        IReadOnlyList<LatexToken> tokens,
        IReadOnlyList<LatexDiagnostic> diagnostics,
        LatexParseOptions options) {
        Source = source;
        SyntaxTree = syntaxTree;
        _tokens = tokens;
        _diagnostics = diagnostics;
        Profile = options.Profile;
        _options = options;

        LatexSemanticModel model = LatexSemanticBuilder.Build(source, syntaxTree, options.Profile);
        _commands = model.Commands;
        _environments = model.Environments;
        _math = model.Math;
        _headings = model.Headings;
        _paragraphs = model.Paragraphs;
        _lists = model.Lists;
        _figures = model.Figures;
        _tables = model.Tables;
        _citations = model.Citations;
        _references = model.References;
        _labels = model.Labels;
        _theorems = model.Theorems;
        _macroDefinitions = model.MacroDefinitions;
        DocumentClassCommand = Commands.FirstOrDefault(static command => string.Equals(command.Name, "documentclass", StringComparison.Ordinal));
        Body = Environments.FirstOrDefault(static environment => string.Equals(environment.Name, "document", StringComparison.Ordinal));
    }

    /// <summary>Original decoded source.</summary>
    public LatexSourceText Source { get; }
    /// <summary>Lossless syntax tree.</summary>
    public LatexSyntaxTree SyntaxTree { get; }
    /// <summary>Selected bounded profile.</summary>
    public LatexDocumentProfile Profile { get; }
    /// <summary>All exact tokens.</summary>
    public IReadOnlyList<LatexToken> Tokens => _tokens;
    /// <summary>Parser and recovery diagnostics.</summary>
    public IReadOnlyList<LatexDiagnostic> Diagnostics => _diagnostics;
    /// <summary>Commands including unknown commands.</summary>
    public IReadOnlyList<LatexCommand> Commands => _commands;
    /// <summary>Nested environments.</summary>
    public IReadOnlyList<LatexEnvironment> Environments => _environments;
    /// <summary>Inline and display math regions.</summary>
    public IReadOnlyList<LatexMath> Math => _math;
    /// <summary>Part/chapter/section headings.</summary>
    public IReadOnlyList<LatexHeading> Headings => _headings;
    /// <summary>Paragraph source regions in the document body.</summary>
    public IReadOnlyList<LatexParagraph> Paragraphs => _paragraphs;
    /// <summary>Itemize, enumerate, and description lists.</summary>
    public IReadOnlyList<LatexList> Lists => _lists;
    /// <summary>Figure environments and included graphics.</summary>
    public IReadOnlyList<LatexFigure> Figures => _figures;
    /// <summary>Tabular environments.</summary>
    public IReadOnlyList<LatexTable> Tables => _tables;
    /// <summary>Citation commands.</summary>
    public IReadOnlyList<LatexCitation> Citations => _citations;
    /// <summary>Cross-reference commands.</summary>
    public IReadOnlyList<LatexReference> References => _references;
    /// <summary>Label declarations.</summary>
    public IReadOnlyList<LatexLabel> Labels => _labels;
    /// <summary>Theorem-like environments.</summary>
    public IReadOnlyList<LatexTheorem> Theorems => _theorems;
    /// <summary>Document-local new/renew/provide command definitions.</summary>
    public IReadOnlyList<LatexMacroDefinition> MacroDefinitions => _macroDefinitions;
    /// <summary>Document class command, when present.</summary>
    public LatexCommand? DocumentClassCommand { get; }
    /// <summary>Main document environment, when present.</summary>
    public LatexEnvironment? Body { get; }
    /// <summary>Document class name.</summary>
    public string? DocumentClassName => DocumentClassCommand?.GetRequiredArgument(0)?.Content;
    /// <summary>True for article, report, or book documents with a document environment.</summary>
    public bool IsRecognizedProfile => Profile == LatexDocumentProfile.OfficeIMO && Body != null &&
        (string.Equals(DocumentClassName, "article", StringComparison.Ordinal) ||
         string.Equals(DocumentClassName, "report", StringComparison.Ordinal) ||
         string.Equals(DocumentClassName, "book", StringComparison.Ordinal));
    /// <summary>True when an editable semantic region changed.</summary>
    public bool IsModified => GetSourceEdits().Any(static edit => edit.IsModified);

    /// <summary>Parses LaTeX source without executing it.</summary>
    public static LatexParseResult Parse(string source, LatexParseOptions? options = null) => LatexParser.Parse(source, options);

    /// <summary>Loads decoded text using runtime UTF-8 BOM detection.</summary>
    public static LatexParseResult Load(string path, LatexParseOptions? options = null) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        return Parse(File.ReadAllText(path), options);
    }

    /// <summary>Writes in preserve mode.</summary>
    public string ToLatex() => LatexWriter.Write(this, null);
    /// <summary>Writes using explicit options.</summary>
    public string ToLatex(LatexWriterOptions? options) => LatexWriter.Write(this, options);
    /// <summary>Saves current text using runtime UTF-8 behavior.</summary>
    public void Save(string path, LatexWriterOptions? options = null) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        File.WriteAllText(path, ToLatex(options));
    }

    /// <summary>Explicitly expands structurally safe simple definitions when enabled by parse options.</summary>
    public LatexMacroExpansionResult ExpandSimpleMacros(string value) {
        if (_options.MacroExpansion != LatexMacroExpansion.SafeSimpleDefinitions) {
            throw new InvalidOperationException("Safe simple macro expansion was not enabled in LatexParseOptions.");
        }
        return LatexSimpleMacroExpander.Expand(value, MacroDefinitions, _options.MaximumExpansionDepth, _options.MaximumExpansionLength);
    }

    internal IEnumerable<ILatexSourceEdit> GetSourceEdits() {
        foreach (LatexCommand command in Commands) {
            for (int index = 0; index < command.Arguments.Count; index++) yield return command.Arguments[index];
        }
        foreach (LatexEnvironment environment in Environments) yield return environment;
        foreach (LatexMath math in Math.Where(static math => math.Environment == null)) yield return math;
        foreach (LatexParagraph paragraph in Paragraphs) yield return paragraph;
        foreach (LatexList list in Lists) {
            foreach (LatexListItem item in list.Items) yield return item;
        }
        foreach (LatexTable table in Tables) {
            foreach (LatexTableCell cell in table.Rows.SelectMany(static row => row.Cells)) yield return cell;
        }
    }
}

/// <summary>Parse result and diagnostics.</summary>
public sealed class LatexParseResult {
    internal LatexParseResult(LatexDocument document, IReadOnlyList<LatexDiagnostic> diagnostics) {
        Document = document;
        Diagnostics = diagnostics;
    }

    /// <summary>Parsed document.</summary>
    public LatexDocument Document { get; }
    /// <summary>Diagnostics.</summary>
    public IReadOnlyList<LatexDiagnostic> Diagnostics { get; }
    /// <summary>True when complete source coverage is retained.</summary>
    public bool IsLossless => Document.SyntaxTree.IsLossless;
    /// <summary>True when an error diagnostic exists.</summary>
    public bool HasErrors => Diagnostics.Any(static diagnostic => diagnostic.Severity == LatexDiagnosticSeverity.Error);
}
