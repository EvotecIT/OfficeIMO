using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace OfficeIMO.Spreadsheet;

/// <summary>Identifies a supported spreadsheet formula grammar.</summary>
public enum SpreadsheetFormulaDialect {
    /// <summary>Excel invariant A1 formula syntax.</summary>
    ExcelA1,

    /// <summary>OASIS OpenFormula syntax used by OpenDocument spreadsheets.</summary>
    OpenFormula
}

/// <summary>Identifies the structural role of a spreadsheet formula syntax node.</summary>
public enum SpreadsheetFormulaSyntaxKind {
    /// <summary>The complete formula.</summary>
    Root,
    /// <summary>A function call and its nested argument syntax.</summary>
    FunctionCall,
    /// <summary>A parenthesized expression.</summary>
    ParenthesizedExpression,
    /// <summary>An inline array constant.</summary>
    InlineArray,
    /// <summary>A leaf token.</summary>
    Token
}

/// <summary>Identifies a lossless formula token's semantic role.</summary>
public enum SpreadsheetFormulaTokenKind {
    /// <summary>The dialect prefix, such as <c>=</c> or <c>of:=</c>.</summary>
    Prefix,
    /// <summary>Whitespace without operator semantics.</summary>
    Whitespace,
    /// <summary>A numeric literal.</summary>
    NumberLiteral,
    /// <summary>A quoted text literal.</summary>
    StringLiteral,
    /// <summary>A spreadsheet error literal.</summary>
    ErrorLiteral,
    /// <summary>A function or named-expression identifier.</summary>
    Identifier,
    /// <summary>A typed cell, row, column, or range reference.</summary>
    Reference,
    /// <summary>An arithmetic, comparison, concatenation, or postfix operator.</summary>
    Operator,
    /// <summary>A function argument separator.</summary>
    ArgumentSeparator,
    /// <summary>An inline-array column separator.</summary>
    ArrayColumnSeparator,
    /// <summary>An inline-array row separator.</summary>
    ArrayRowSeparator,
    /// <summary>A range-union operator.</summary>
    UnionOperator,
    /// <summary>A range-intersection operator.</summary>
    IntersectionOperator,
    /// <summary>An opening delimiter retained in a structural node.</summary>
    OpenDelimiter,
    /// <summary>A closing delimiter retained in a structural node.</summary>
    CloseDelimiter,
    /// <summary>Syntax that was preserved but cannot be translated safely.</summary>
    Unsupported
}

/// <summary>Severity of a spreadsheet formula parse or translation diagnostic.</summary>
public enum SpreadsheetFormulaDiagnosticSeverity {
    /// <summary>The formula remains translatable, but the caller should review the finding.</summary>
    Warning,
    /// <summary>The formula cannot be translated without risking invalid or changed semantics.</summary>
    Error
}

/// <summary>One structured formula parse or translation finding.</summary>
public sealed class SpreadsheetFormulaDiagnostic {
    internal SpreadsheetFormulaDiagnostic(
        string code,
        SpreadsheetFormulaDiagnosticSeverity severity,
        string message,
        int position,
        int length) {
        Code = code;
        Severity = severity;
        Message = message;
        Position = position;
        Length = length;
    }

    /// <summary>Gets the stable diagnostic code.</summary>
    public string Code { get; }

    /// <summary>Gets the finding severity.</summary>
    public SpreadsheetFormulaDiagnosticSeverity Severity { get; }

    /// <summary>Gets the human-readable finding.</summary>
    public string Message { get; }

    /// <summary>Gets the zero-based source position.</summary>
    public int Position { get; }

    /// <summary>Gets the source length associated with the finding.</summary>
    public int Length { get; }
}

/// <summary>One immutable node in a lossless spreadsheet formula syntax tree.</summary>
public sealed class SpreadsheetFormulaSyntaxNode {
    private readonly ReadOnlyCollection<SpreadsheetFormulaSyntaxNode> _children;

    internal SpreadsheetFormulaSyntaxNode(
        SpreadsheetFormulaSyntaxKind kind,
        SpreadsheetFormulaTokenKind? tokenKind,
        string text,
        int position,
        IEnumerable<SpreadsheetFormulaSyntaxNode>? children = null,
        SpreadsheetRangeReference? reference = null,
        string? name = null) {
        Kind = kind;
        TokenKind = tokenKind;
        Text = text;
        Position = position;
        Reference = reference;
        Name = name;
        _children = Array.AsReadOnly((children ?? Array.Empty<SpreadsheetFormulaSyntaxNode>()).ToArray());
    }

    /// <summary>Gets the node's structural role.</summary>
    public SpreadsheetFormulaSyntaxKind Kind { get; }

    /// <summary>Gets the leaf token role, or <see langword="null"/> for structural nodes.</summary>
    public SpreadsheetFormulaTokenKind? TokenKind { get; }

    /// <summary>Gets the exact authored source represented by this node.</summary>
    public string Text { get; }

    /// <summary>Gets the zero-based source position.</summary>
    public int Position { get; }

    /// <summary>Gets the nested syntax in authored order.</summary>
    public IReadOnlyList<SpreadsheetFormulaSyntaxNode> Children => _children;

    /// <summary>Gets the typed reference for a reference token.</summary>
    public SpreadsheetRangeReference? Reference { get; }

    /// <summary>Gets the decoded function name for a function-call node.</summary>
    public string? Name { get; }
}

/// <summary>
/// Lossless, nested syntax tree for Excel A1 or OpenFormula text. Parsing keeps literals and trivia
/// while assigning separator and reference semantics from their structural context.
/// </summary>
public sealed class SpreadsheetFormulaSyntaxTree {
    private readonly ReadOnlyCollection<SpreadsheetFormulaDiagnostic> _diagnostics;

    internal SpreadsheetFormulaSyntaxTree(
        string text,
        SpreadsheetFormulaDialect dialect,
        SpreadsheetFormulaSyntaxNode root,
        IEnumerable<SpreadsheetFormulaDiagnostic> diagnostics) {
        Text = text;
        Dialect = dialect;
        Root = root;
        _diagnostics = Array.AsReadOnly(diagnostics.ToArray());
    }

    /// <summary>Gets the exact authored formula.</summary>
    public string Text { get; }

    /// <summary>Gets the source grammar.</summary>
    public SpreadsheetFormulaDialect Dialect { get; }

    /// <summary>Gets the root syntax node.</summary>
    public SpreadsheetFormulaSyntaxNode Root { get; }

    /// <summary>Gets parse findings in source order.</summary>
    public IReadOnlyList<SpreadsheetFormulaDiagnostic> Diagnostics => _diagnostics;

    /// <summary>Gets whether the authored formula is structurally valid for safe translation.</summary>
    public bool IsValid => !_diagnostics.Any(diagnostic => diagnostic.Severity == SpreadsheetFormulaDiagnosticSeverity.Error);

    /// <summary>Parses a formula using the requested grammar.</summary>
    public static SpreadsheetFormulaSyntaxTree Parse(string text, SpreadsheetFormulaDialect dialect) {
        if (text == null) throw new ArgumentNullException(nameof(text));
        return SpreadsheetFormulaParser.Parse(text, dialect);
    }

    /// <summary>Translates the parsed tree to another supported grammar.</summary>
    public SpreadsheetFormulaTranslationResult TranslateTo(SpreadsheetFormulaDialect targetDialect) =>
        SpreadsheetFormulaTranslator.Translate(this, targetDialect);
}

/// <summary>Result of translating a typed spreadsheet formula syntax tree.</summary>
public sealed class SpreadsheetFormulaTranslationResult {
    private readonly ReadOnlyCollection<SpreadsheetFormulaDiagnostic> _diagnostics;

    internal SpreadsheetFormulaTranslationResult(
        string formula,
        SpreadsheetFormulaDialect sourceDialect,
        SpreadsheetFormulaDialect targetDialect,
        IEnumerable<SpreadsheetFormulaDiagnostic> diagnostics) {
        Formula = formula;
        SourceDialect = sourceDialect;
        TargetDialect = targetDialect;
        _diagnostics = Array.AsReadOnly(diagnostics.ToArray());
    }

    /// <summary>Gets the translated formula candidate.</summary>
    public string Formula { get; }

    /// <summary>Gets the source grammar.</summary>
    public SpreadsheetFormulaDialect SourceDialect { get; }

    /// <summary>Gets the destination grammar.</summary>
    public SpreadsheetFormulaDialect TargetDialect { get; }

    /// <summary>Gets parse and translation findings.</summary>
    public IReadOnlyList<SpreadsheetFormulaDiagnostic> Diagnostics => _diagnostics;

    /// <summary>Gets whether the candidate is safe to emit.</summary>
    public bool IsSuccessful => !_diagnostics.Any(diagnostic => diagnostic.Severity == SpreadsheetFormulaDiagnosticSeverity.Error);

    /// <summary>Returns the translated formula or throws when safe translation was not possible.</summary>
    public string RequireFormula() {
        if (!IsSuccessful) {
            throw new InvalidOperationException("The spreadsheet formula could not be translated safely. Inspect Diagnostics for details.");
        }
        return Formula;
    }
}
