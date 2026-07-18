using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Drawing;

/// <summary>Factory helpers for authoring reusable mathematical expressions.</summary>
public static class OfficeMath {
    /// <summary>Creates a horizontal sequence.</summary>
    public static OfficeMathExpression Row(params OfficeMathExpression[] children) =>
        Create(OfficeMathKind.Row, children: children);

    /// <summary>Creates literal text.</summary>
    public static OfficeMathExpression Text(string text) => Token(OfficeMathKind.Text, text);

    /// <summary>Creates an identifier.</summary>
    public static OfficeMathExpression Identifier(string text) => Token(OfficeMathKind.Identifier, text);

    /// <summary>Creates a number.</summary>
    public static OfficeMathExpression Number(string text) => Token(OfficeMathKind.Number, text);

    /// <summary>Creates an operator.</summary>
    public static OfficeMathExpression Operator(string text) => Token(OfficeMathKind.Operator, text);

    /// <summary>Creates a fraction.</summary>
    public static OfficeMathExpression Fraction(OfficeMathExpression numerator, OfficeMathExpression denominator) =>
        Create(OfficeMathKind.Fraction, children: Required(numerator, denominator));

    /// <summary>Creates a square root.</summary>
    public static OfficeMathExpression Radical(OfficeMathExpression content) =>
        Create(OfficeMathKind.Radical, children: Required(content));

    /// <summary>Creates an indexed root.</summary>
    public static OfficeMathExpression Radical(OfficeMathExpression content, OfficeMathExpression index) =>
        Create(OfficeMathKind.Radical, children: Required(content, index));

    /// <summary>Creates a superscript.</summary>
    public static OfficeMathExpression Superscript(OfficeMathExpression basis, OfficeMathExpression script) =>
        Create(OfficeMathKind.Superscript, children: Required(basis, script));

    /// <summary>Creates a subscript.</summary>
    public static OfficeMathExpression Subscript(OfficeMathExpression basis, OfficeMathExpression script) =>
        Create(OfficeMathKind.Subscript, children: Required(basis, script));

    /// <summary>Creates combined subscript and superscript.</summary>
    public static OfficeMathExpression SubSuperscript(OfficeMathExpression basis, OfficeMathExpression subscript, OfficeMathExpression superscript) =>
        Create(OfficeMathKind.SubSuperscript, children: Required(basis, subscript, superscript));

    /// <summary>Creates left-side combined subscript and superscript.</summary>
    public static OfficeMathExpression LeftSubSuperscript(OfficeMathExpression basis, OfficeMathExpression subscript, OfficeMathExpression superscript) =>
        Create(OfficeMathKind.LeftSubSuperscript, children: Required(basis, subscript, superscript));

    /// <summary>Creates a limit centered below a base.</summary>
    public static OfficeMathExpression LowerLimit(OfficeMathExpression basis, OfficeMathExpression limit) =>
        Create(OfficeMathKind.LowerLimit, children: Required(basis, limit));

    /// <summary>Creates a limit centered above a base.</summary>
    public static OfficeMathExpression UpperLimit(OfficeMathExpression basis, OfficeMathExpression limit) =>
        Create(OfficeMathKind.UpperLimit, children: Required(basis, limit));

    /// <summary>Creates a diagonal or skewed fraction.</summary>
    public static OfficeMathExpression SlashedFraction(OfficeMathExpression numerator, OfficeMathExpression denominator) =>
        Create(OfficeMathKind.SlashedFraction, children: Required(numerator, denominator));

    /// <summary>Creates a large operator with content and optional limits.</summary>
    public static OfficeMathExpression Nary(string character, OfficeMathExpression content, OfficeMathExpression? lower = null, OfficeMathExpression? upper = null) {
        if (string.IsNullOrEmpty(character)) throw new ArgumentException("An n-ary operator character is required.", nameof(character));
        var children = new List<OfficeMathExpression> { content ?? throw new ArgumentNullException(nameof(content)) };
        if (lower != null) children.Add(lower);
        if (upper != null) children.Add(upper);
        return Create(OfficeMathKind.Nary, character: character, children: children, naryUpperOnly: lower == null && upper != null);
    }

    /// <summary>Creates delimited content.</summary>
    public static OfficeMathExpression Delimited(OfficeMathExpression content, string left = "(", string right = ")") =>
        Create(OfficeMathKind.Delimited, character: left, secondaryCharacter: right, children: Required(content));

    /// <summary>Creates multiple separated expressions inside one delimiter pair.</summary>
    public static OfficeMathExpression DelimiterList(string left, string right, string separator, params OfficeMathExpression[] items) {
        if (items == null) throw new ArgumentNullException(nameof(items));
        if (items.Length == 0) throw new ArgumentException("At least one delimiter-list item is required.", nameof(items));
        return Create(OfficeMathKind.DelimiterList, character: left, secondaryCharacter: right,
            separatorCharacter: separator, children: Required(items));
    }

    /// <summary>Creates expressions stacked vertically.</summary>
    public static OfficeMathExpression Stack(params OfficeMathExpression[] rows) =>
        Create(OfficeMathKind.Stack, children: Required(rows));

    /// <summary>Creates expressions stacked vertically with native stretch semantics.</summary>
    public static OfficeMathExpression StretchStack(params OfficeMathExpression[] rows) =>
        Create(OfficeMathKind.StretchStack, children: Required(rows));

    /// <summary>Creates a named function application.</summary>
    public static OfficeMathExpression Function(string name, OfficeMathExpression argument) {
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("A function name is required.", nameof(name));
        return Create(OfficeMathKind.Function, text: name, children: Required(argument));
    }

    /// <summary>Creates a matrix from row-major cells.</summary>
    public static OfficeMathExpression Matrix(int rows, int columns, params OfficeMathExpression[] cells) =>
        Create(OfficeMathKind.Matrix, children: cells, rowCount: rows, columnCount: columns);

    /// <summary>Creates an aligned equation array from row-major cells.</summary>
    public static OfficeMathExpression EquationArray(int rows, int columns, params OfficeMathExpression[] cells) =>
        Create(OfficeMathKind.EquationArray, children: cells, rowCount: rows, columnCount: columns);

    /// <summary>Creates accented content.</summary>
    public static OfficeMathExpression Accent(OfficeMathExpression content, string character) =>
        Create(OfficeMathKind.Accent, character: character, children: Required(content));

    /// <summary>Creates content with an overbar.</summary>
    public static OfficeMathExpression Overbar(OfficeMathExpression content) => Create(OfficeMathKind.Overbar, children: Required(content));

    /// <summary>Creates content with an underbar.</summary>
    public static OfficeMathExpression Underbar(OfficeMathExpression content) => Create(OfficeMathKind.Underbar, children: Required(content));

    /// <summary>Creates boxed content.</summary>
    public static OfficeMathExpression Box(OfficeMathExpression content) => Create(OfficeMathKind.Box, children: Required(content));

    /// <summary>Creates invisible content that retains layout space.</summary>
    public static OfficeMathExpression Phantom(OfficeMathExpression content) => Create(OfficeMathKind.Phantom, children: Required(content));

    private static OfficeMathExpression Token(OfficeMathKind kind, string text) {
        if (text == null) throw new ArgumentNullException(nameof(text));
        return Create(kind, text: text);
    }

    private static OfficeMathExpression[] Required(params OfficeMathExpression?[] children) {
        if (children.Any(child => child == null)) throw new ArgumentNullException(nameof(children));
        return children.Cast<OfficeMathExpression>().ToArray();
    }

    internal static OfficeMathExpression Create(
        OfficeMathKind kind,
        string? text = null,
        IEnumerable<OfficeMathExpression>? children = null,
        string? character = null,
        string? secondaryCharacter = null,
        int rowCount = 0,
        int columnCount = 0,
        string? separatorCharacter = null,
        bool naryUpperOnly = false) =>
        new OfficeMathExpression(kind, text, children, character, secondaryCharacter, rowCount, columnCount, separatorCharacter, naryUpperOnly);
}
