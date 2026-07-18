using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;

namespace OfficeIMO.Drawing;

/// <summary>Format-neutral mathematical structure shared by Office document engines.</summary>
public enum OfficeMathKind {
    /// <summary>An ordered horizontal sequence.</summary>
    Row = 0,
    /// <summary>Literal text.</summary>
    Text = 1,
    /// <summary>A mathematical identifier.</summary>
    Identifier = 2,
    /// <summary>A number.</summary>
    Number = 3,
    /// <summary>An operator or punctuation token.</summary>
    Operator = 4,
    /// <summary>A numerator over a denominator.</summary>
    Fraction = 5,
    /// <summary>A square or indexed root.</summary>
    Radical = 6,
    /// <summary>A base with a superscript.</summary>
    Superscript = 7,
    /// <summary>A base with a subscript.</summary>
    Subscript = 8,
    /// <summary>A base with subscript and superscript.</summary>
    SubSuperscript = 9,
    /// <summary>A large operator with optional limits.</summary>
    Nary = 10,
    /// <summary>An expression surrounded by delimiters.</summary>
    Delimited = 11,
    /// <summary>A named function and its argument.</summary>
    Function = 12,
    /// <summary>A rectangular matrix.</summary>
    Matrix = 13,
    /// <summary>Vertically aligned equations.</summary>
    EquationArray = 14,
    /// <summary>An accent placed above content.</summary>
    Accent = 15,
    /// <summary>A bar placed above content.</summary>
    Overbar = 16,
    /// <summary>A bar placed below content.</summary>
    Underbar = 17,
    /// <summary>Content enclosed by a box.</summary>
    Box = 18,
    /// <summary>Invisible content that retains its layout space.</summary>
    Phantom = 19,
    /// <summary>A base with left-side subscript and superscript.</summary>
    LeftSubSuperscript = 20,
    /// <summary>A limit centered below its base.</summary>
    LowerLimit = 21,
    /// <summary>A limit centered above its base.</summary>
    UpperLimit = 22,
    /// <summary>A numerator and denominator separated by a diagonal slash.</summary>
    SlashedFraction = 23,
    /// <summary>Expressions stacked vertically.</summary>
    Stack = 24,
    /// <summary>Expressions stacked vertically with stretch semantics.</summary>
    StretchStack = 25,
    /// <summary>Multiple expressions separated inside one delimiter pair.</summary>
    DelimiterList = 26
}

/// <summary>An immutable node in a reusable mathematical expression tree.</summary>
public sealed class OfficeMathExpression : IEquatable<OfficeMathExpression> {
    private readonly ReadOnlyCollection<OfficeMathExpression> _children;
    private readonly bool _naryUpperOnly;

    internal OfficeMathExpression(
        OfficeMathKind kind,
        string? text,
        IEnumerable<OfficeMathExpression>? children,
        string? character,
        string? secondaryCharacter,
        int rowCount,
        int columnCount,
        string? separatorCharacter,
        bool naryUpperOnly) {
        Kind = kind;
        Text = text;
        Character = character;
        SecondaryCharacter = secondaryCharacter;
        RowCount = rowCount;
        ColumnCount = columnCount;
        SeparatorCharacter = separatorCharacter;
        _naryUpperOnly = naryUpperOnly;
        _children = new ReadOnlyCollection<OfficeMathExpression>(children?.ToList() ?? new List<OfficeMathExpression>());
        Validate();
    }

    /// <summary>Node kind.</summary>
    public OfficeMathKind Kind { get; }

    /// <summary>Token text for text, identifier, number, operator, or function nodes.</summary>
    public string? Text { get; }

    /// <summary>Primary delimiter, accent, or n-ary operator character.</summary>
    public string? Character { get; }

    /// <summary>Secondary delimiter character.</summary>
    public string? SecondaryCharacter { get; }

    /// <summary>Separator used by a delimiter list.</summary>
    public string? SeparatorCharacter { get; }

    /// <summary>Child expressions in semantic order.</summary>
    public IReadOnlyList<OfficeMathExpression> Children => _children;

    /// <summary>Lower limit for an n-ary expression, or <see langword="null"/> when omitted.</summary>
    public OfficeMathExpression? NaryLowerLimit =>
        Kind == OfficeMathKind.Nary && !_naryUpperOnly && _children.Count > 1 ? _children[1] : null;

    /// <summary>Upper limit for an n-ary expression, or <see langword="null"/> when omitted.</summary>
    public OfficeMathExpression? NaryUpperLimit =>
        Kind != OfficeMathKind.Nary ? null : _naryUpperOnly ? _children[1] : _children.Count > 2 ? _children[2] : null;

    /// <summary>Row count for matrix-like structures.</summary>
    public int RowCount { get; }

    /// <summary>Column count for matrix-like structures.</summary>
    public int ColumnCount { get; }

    /// <summary>Returns a readable plain-text projection.</summary>
    public string ToPlainText() {
        switch (Kind) {
            case OfficeMathKind.Text:
            case OfficeMathKind.Identifier:
            case OfficeMathKind.Number:
            case OfficeMathKind.Operator:
                return Text ?? string.Empty;
            case OfficeMathKind.Row:
                return string.Concat(_children.Select(child => child.ToPlainText()));
            case OfficeMathKind.Fraction:
                return "(" + _children[0].ToPlainText() + ")/(" + _children[1].ToPlainText() + ")";
            case OfficeMathKind.Radical:
                return _children.Count == 2
                    ? "root[" + _children[1].ToPlainText() + "](" + _children[0].ToPlainText() + ")"
                    : "sqrt(" + _children[0].ToPlainText() + ")";
            case OfficeMathKind.Superscript:
                return _children[0].ToPlainText() + "^(" + _children[1].ToPlainText() + ")";
            case OfficeMathKind.Subscript:
                return _children[0].ToPlainText() + "_(" + _children[1].ToPlainText() + ")";
            case OfficeMathKind.SubSuperscript:
                return _children[0].ToPlainText() + "_(" + _children[1].ToPlainText() + ")^(" + _children[2].ToPlainText() + ")";
            case OfficeMathKind.LeftSubSuperscript:
                return "_(" + _children[1].ToPlainText() + ")^(" + _children[2].ToPlainText() + ")" + _children[0].ToPlainText();
            case OfficeMathKind.LowerLimit:
                return _children[0].ToPlainText() + "_(" + _children[1].ToPlainText() + ")";
            case OfficeMathKind.UpperLimit:
                return _children[0].ToPlainText() + "^(" + _children[1].ToPlainText() + ")";
            case OfficeMathKind.SlashedFraction:
                return "(" + _children[0].ToPlainText() + ")/(" + _children[1].ToPlainText() + ")";
            case OfficeMathKind.Nary:
                return (Character ?? "∑") + LimitsPlainText() + _children[0].ToPlainText();
            case OfficeMathKind.Delimited:
                return (Character ?? "(") + _children[0].ToPlainText() + (SecondaryCharacter ?? ")");
            case OfficeMathKind.DelimiterList:
                return (Character ?? "(") + string.Join(SeparatorCharacter ?? ",", _children.Select(child => child.ToPlainText())) + (SecondaryCharacter ?? ")");
            case OfficeMathKind.Function:
                return (Text ?? string.Empty) + "(" + _children[0].ToPlainText() + ")";
            case OfficeMathKind.Matrix:
            case OfficeMathKind.EquationArray:
                return MatrixPlainText();
            case OfficeMathKind.Accent:
                return _children[0].ToPlainText() + (Character ?? "̂");
            case OfficeMathKind.Overbar:
                return "overbar(" + _children[0].ToPlainText() + ")";
            case OfficeMathKind.Underbar:
                return "underbar(" + _children[0].ToPlainText() + ")";
            case OfficeMathKind.Box:
                return "[" + _children[0].ToPlainText() + "]";
            case OfficeMathKind.Phantom:
                return _children[0].ToPlainText();
            case OfficeMathKind.Stack:
            case OfficeMathKind.StretchStack:
                return string.Join("; ", _children.Select(child => child.ToPlainText()));
            default:
                return string.Empty;
        }
    }

    /// <inheritdoc />
    public override string ToString() => ToPlainText();

    /// <inheritdoc />
    public bool Equals(OfficeMathExpression? other) {
        if (other == null || Kind != other.Kind || Text != other.Text || Character != other.Character ||
            SecondaryCharacter != other.SecondaryCharacter || SeparatorCharacter != other.SeparatorCharacter || RowCount != other.RowCount ||
            ColumnCount != other.ColumnCount || _naryUpperOnly != other._naryUpperOnly || _children.Count != other._children.Count) return false;
        for (int index = 0; index < _children.Count; index++) {
            if (!_children[index].Equals(other._children[index])) return false;
        }
        return true;
    }

    /// <inheritdoc />
    public override bool Equals(object? obj) => Equals(obj as OfficeMathExpression);

    /// <inheritdoc />
    public override int GetHashCode() {
        unchecked {
            int hash = (int)Kind;
            hash = (hash * 397) ^ (Text?.GetHashCode() ?? 0);
            hash = (hash * 397) ^ (Character?.GetHashCode() ?? 0);
            hash = (hash * 397) ^ (SecondaryCharacter?.GetHashCode() ?? 0);
            hash = (hash * 397) ^ (SeparatorCharacter?.GetHashCode() ?? 0);
            hash = (hash * 397) ^ RowCount;
            hash = (hash * 397) ^ ColumnCount;
            hash = (hash * 397) ^ _naryUpperOnly.GetHashCode();
            for (int index = 0; index < _children.Count; index++) hash = (hash * 397) ^ _children[index].GetHashCode();
            return hash;
        }
    }

    private string LimitsPlainText() {
        var builder = new StringBuilder();
        if (NaryLowerLimit != null) builder.Append("_(").Append(NaryLowerLimit.ToPlainText()).Append(')');
        if (NaryUpperLimit != null) builder.Append("^(").Append(NaryUpperLimit.ToPlainText()).Append(')');
        return builder.ToString();
    }

    private string MatrixPlainText() {
        var builder = new StringBuilder();
        builder.Append('[');
        for (int row = 0; row < RowCount; row++) {
            if (row > 0) builder.Append("; ");
            for (int column = 0; column < ColumnCount; column++) {
                if (column > 0) builder.Append(", ");
                int index = row * ColumnCount + column;
                if (index < _children.Count) builder.Append(_children[index].ToPlainText());
            }
        }
        return builder.Append(']').ToString();
    }

    private void Validate() {
        if ((Kind == OfficeMathKind.Text || Kind == OfficeMathKind.Identifier || Kind == OfficeMathKind.Number ||
             Kind == OfficeMathKind.Operator) && Text == null) throw new ArgumentNullException(nameof(Text));
        int count = _children.Count;
        switch (Kind) {
            case OfficeMathKind.Fraction:
            case OfficeMathKind.SlashedFraction:
            case OfficeMathKind.Superscript:
            case OfficeMathKind.Subscript:
            case OfficeMathKind.LowerLimit:
            case OfficeMathKind.UpperLimit:
                RequireChildren(count, 2); break;
            case OfficeMathKind.SubSuperscript:
            case OfficeMathKind.LeftSubSuperscript:
                RequireChildren(count, 3); break;
            case OfficeMathKind.Radical:
                if (count < 1 || count > 2) throw new ArgumentException("A radical requires content and an optional index.");
                break;
            case OfficeMathKind.Nary:
                if (count < 1 || count > 3) throw new ArgumentException("An n-ary expression requires content and optional lower and upper limits.");
                if (_naryUpperOnly && count != 2) throw new ArgumentException("An upper-only n-ary expression requires content and one upper limit.");
                break;
            case OfficeMathKind.Delimited:
            case OfficeMathKind.Function:
            case OfficeMathKind.Accent:
            case OfficeMathKind.Overbar:
            case OfficeMathKind.Underbar:
            case OfficeMathKind.Box:
            case OfficeMathKind.Phantom:
                RequireChildren(count, 1); break;
            case OfficeMathKind.DelimiterList:
            case OfficeMathKind.Stack:
            case OfficeMathKind.StretchStack:
                if (count < 1) throw new ArgumentException("The mathematical structure requires at least one child.");
                break;
            case OfficeMathKind.Matrix:
            case OfficeMathKind.EquationArray:
                if (RowCount < 1 || ColumnCount < 1 || count != checked(RowCount * ColumnCount)) {
                    throw new ArgumentException("Matrix dimensions must match the number of cells.");
                }
                break;
        }
        if (_naryUpperOnly && Kind != OfficeMathKind.Nary) {
            throw new ArgumentException("Only n-ary expressions can declare an upper-only limit.");
        }
    }

    private static void RequireChildren(int actual, int expected) {
        if (actual != expected) throw new ArgumentException("The mathematical structure has an invalid number of children.");
    }
}
