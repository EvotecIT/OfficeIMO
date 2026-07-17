using OfficeIMO.Drawing;
using System.Text;

namespace OfficeIMO.OneNote;

internal static class OneNoteMathNativeCodec {
    internal const char ObjectStart = '\uFDD0';
    internal const char ObjectSeparator = '\uFDEE';
    internal const char ObjectEnd = '\uFDEF';
    internal const uint PlainTextType = 0x90000000U;

    internal sealed class EncodedRun {
        internal EncodedRun(string text, OneNoteMathInlineDescriptor descriptor) { Text = text; Descriptor = descriptor; }
        internal string Text { get; }
        internal OneNoteMathInlineDescriptor Descriptor { get; }
    }

    internal static IReadOnlyList<EncodedRun> Encode(OfficeMathExpression expression) {
        if (expression == null) throw new ArgumentNullException(nameof(expression));
        var runs = new List<EncodedRun>();
        AppendBoundaryGroup(expression, runs);
        if (runs.Count == 0) runs.Add(new EncodedRun(string.Empty, new OneNoteMathInlineDescriptor { Type = PlainTextType }));
        return runs;
    }

    private static void AppendBoundaryGroup(OfficeMathExpression expression, IList<EncodedRun> runs) {
        var descriptor = new OneNoteMathInlineDescriptor { Type = 12, Count = 1 };
        runs.Add(new EncodedRun(ObjectStart.ToString(), descriptor));
        if (IsSimpleSequence(expression)) {
            runs.Add(new EncodedRun(Sanitize(expression.ToPlainText()) + ObjectEnd, new OneNoteMathInlineDescriptor { Type = 12 }));
        } else {
            AppendTopLevel(expression, runs);
            runs.Add(new EncodedRun(ObjectEnd.ToString(), new OneNoteMathInlineDescriptor { Type = 12 }));
        }
    }

    internal static OfficeMathExpression Canonicalize(OfficeMathExpression expression) {
        var runs = new List<OneNoteTextRun>();
        foreach (EncodedRun encoded in Encode(expression)) {
            var run = new OneNoteTextRun { Text = encoded.Text, MathDescriptor = encoded.Descriptor };
            run.Style.IsMath = true;
            runs.Add(run);
        }
        return Decode(runs);
    }

    internal static OfficeMathExpression Decode(
        IList<OneNoteTextRun> runs,
        int maximumDepth = OneNoteReaderOptions.DefaultMaxPropertySetDepth) {
        if (runs == null) throw new ArgumentNullException(nameof(runs));
        if (maximumDepth < 1) throw new ArgumentOutOfRangeException(nameof(maximumDepth));
        var text = new StringBuilder();
        var descriptors = new Dictionary<int, OneNoteMathInlineDescriptor>();
        for (int index = 0; index < runs.Count; index++) {
            OneNoteTextRun run = runs[index];
            int start = text.Length;
            string value = (run.Text ?? string.Empty).TrimEnd('\0');
            text.Append(value);
            if (value.Length > 0 && value[0] == ObjectStart && run.MathDescriptor != null) descriptors[start] = run.MathDescriptor;
        }
        if (text.Length == 0) return OfficeMath.Text(string.Empty);
        var parser = new NativeParser(text.ToString(), descriptors, maximumDepth);
        return parser.Parse();
    }

    private static void AppendTopLevel(OfficeMathExpression expression, IList<EncodedRun> runs) {
        if (expression.Kind == OfficeMathKind.Row) {
            for (int index = 0; index < expression.Children.Count; index++) AppendTopLevel(expression.Children[index], runs);
        } else if (IsSimple(expression)) {
            runs.Add(new EncodedRun(Sanitize(expression.ToPlainText()), new OneNoteMathInlineDescriptor { Type = PlainTextType }));
        } else {
            AppendObject(expression, runs);
        }
    }

    private static void AppendObject(OfficeMathExpression expression, IList<EncodedRun> runs) {
        uint type = NativeType(expression.Kind);
        IReadOnlyList<OfficeMathExpression> children = NativeChildren(expression);
        OneNoteMathInlineDescriptor root = Descriptor(expression, type);
        root.Count = (uint)children.Count;
        runs.Add(new EncodedRun(ObjectStart.ToString(), root));
        for (int index = 0; index < children.Count; index++) {
            char terminal = index == children.Count - 1 ? ObjectEnd : ObjectSeparator;
            OfficeMathExpression child = children[index];
            OneNoteMathInlineDescriptor parentDescriptor = Descriptor(expression, type);
            if (index > 0) parentDescriptor.Count = (uint)index;
            if (IsSimpleSequence(child)) {
                runs.Add(new EncodedRun(Sanitize(child.ToPlainText()) + terminal, parentDescriptor));
            } else {
                AppendTopLevel(child, runs);
                runs.Add(new EncodedRun(terminal.ToString(), parentDescriptor));
            }
        }
    }

    private static OneNoteMathInlineDescriptor Descriptor(OfficeMathExpression expression, uint type) {
        var descriptor = new OneNoteMathInlineDescriptor { Type = type };
        string? primary = expression.Character;
        string? secondary = expression.SecondaryCharacter;
        string? tertiary = expression.SeparatorCharacter;
        if (expression.Kind == OfficeMathKind.Fraction) primary = "/";
        else if (expression.Kind == OfficeMathKind.Superscript) primary = "^";
        else if (expression.Kind == OfficeMathKind.Subscript) primary = "_";
        else if (expression.Kind == OfficeMathKind.SubSuperscript) primary = "^";
        else if (expression.Kind == OfficeMathKind.Function) primary = "\u2061";
        descriptor.Character = NativeCharacter(primary, "primary");
        descriptor.Character1 = NativeCharacter(secondary, "secondary");
        descriptor.Character2 = NativeCharacter(tertiary, "separator");
        if (expression.Kind == OfficeMathKind.Matrix || expression.Kind == OfficeMathKind.EquationArray) {
            descriptor.Column = (byte)Math.Min(byte.MaxValue, expression.ColumnCount);
            descriptor.Alignment = 1;
        }
        return descriptor;
    }

    internal static void ValidateNativeCharacters(OfficeMathExpression expression) {
        if (expression == null) throw new ArgumentNullException(nameof(expression));
        NativeCharacter(expression.Character, "primary");
        NativeCharacter(expression.SecondaryCharacter, "secondary");
        NativeCharacter(expression.SeparatorCharacter, "separator");
    }

    private static ushort? NativeCharacter(string? value, string role) {
        if (value == null) return null;
        if (value.Length != 1 || char.IsSurrogate(value[0])) {
            throw new OneNoteFormatException(
                "ONENOTE_WRITE_MATH_CHARACTER",
                "A native OneNote math " + role + " character must be exactly one non-surrogate UTF-16 code unit.");
        }
        return value[0];
    }

    private static IReadOnlyList<OfficeMathExpression> NativeChildren(OfficeMathExpression expression) {
        if (expression.Kind == OfficeMathKind.Function) {
            return new[] { OfficeMath.Identifier(expression.Text ?? string.Empty), expression.Children[0] };
        }
        if (expression.Kind == OfficeMathKind.Radical && expression.Children.Count == 2) {
            return new[] { expression.Children[1], expression.Children[0] };
        }
        if (expression.Kind == OfficeMathKind.Nary) {
            var values = new List<OfficeMathExpression>();
            if (expression.Children.Count > 1) values.Add(expression.Children[1]);
            if (expression.Children.Count > 2) values.Add(expression.Children[2]);
            values.Add(expression.Children[0]);
            return values;
        }
        return expression.Children;
    }

    private static uint NativeType(OfficeMathKind kind) {
        switch (kind) {
            case OfficeMathKind.Accent: return 10;
            case OfficeMathKind.Box: return 11;
            case OfficeMathKind.Delimited: return 13;
            case OfficeMathKind.DelimiterList: return 14;
            case OfficeMathKind.EquationArray: return 15;
            case OfficeMathKind.Fraction: return 16;
            case OfficeMathKind.LeftSubSuperscript: return 18;
            case OfficeMathKind.LowerLimit: return 19;
            case OfficeMathKind.Function: return 17;
            case OfficeMathKind.Matrix: return 20;
            case OfficeMathKind.Nary: return 21;
            case OfficeMathKind.Overbar: return 23;
            case OfficeMathKind.Phantom: return 24;
            case OfficeMathKind.Radical: return 25;
            case OfficeMathKind.SlashedFraction: return 26;
            case OfficeMathKind.Stack: return 27;
            case OfficeMathKind.StretchStack: return 28;
            case OfficeMathKind.Subscript: return 29;
            case OfficeMathKind.SubSuperscript: return 30;
            case OfficeMathKind.Superscript: return 31;
            case OfficeMathKind.Underbar: return 32;
            case OfficeMathKind.UpperLimit: return 33;
            default: return 12;
        }
    }

    private static bool IsSimple(OfficeMathExpression expression) =>
        expression.Kind == OfficeMathKind.Text || expression.Kind == OfficeMathKind.Identifier ||
        expression.Kind == OfficeMathKind.Number || expression.Kind == OfficeMathKind.Operator;

    private static bool IsSimpleSequence(OfficeMathExpression expression) {
        if (IsSimple(expression)) return true;
        if (expression.Kind != OfficeMathKind.Row) return false;
        for (int index = 0; index < expression.Children.Count; index++) if (!IsSimple(expression.Children[index])) return false;
        return true;
    }

    private static string Sanitize(string value) => value.Replace(ObjectStart, ' ').Replace(ObjectSeparator, ' ').Replace(ObjectEnd, ' ');

    private sealed class NativeParser {
        private readonly string _text;
        private readonly IReadOnlyDictionary<int, OneNoteMathInlineDescriptor> _descriptors;
        private readonly int _maximumDepth;
        private int _position;

        internal NativeParser(
            string text,
            IReadOnlyDictionary<int, OneNoteMathInlineDescriptor> descriptors,
            int maximumDepth) {
            _text = text;
            _descriptors = descriptors;
            _maximumDepth = maximumDepth;
        }

        internal OfficeMathExpression Parse() => Collapse(ParseSequence(0));

        private IList<OfficeMathExpression> ParseSequence(int depth) {
            var expressions = new List<OfficeMathExpression>();
            var plain = new StringBuilder();
            while (_position < _text.Length && _text[_position] != ObjectSeparator && _text[_position] != ObjectEnd) {
                if (_text[_position] == ObjectStart) {
                    FlushPlain(expressions, plain);
                    expressions.Add(ParseObject(depth));
                } else {
                    plain.Append(_text[_position++]);
                }
            }
            FlushPlain(expressions, plain);
            return expressions;
        }

        private OfficeMathExpression ParseObject(int depth) {
            if (depth >= _maximumDepth) {
                throw new OneNoteFormatException(
                    "ONENOTE_MATH_DEPTH",
                    "The native OneNote math nesting depth limit was exceeded.");
            }
            int start = _position++;
            _descriptors.TryGetValue(start, out OneNoteMathInlineDescriptor? descriptor);
            descriptor ??= new OneNoteMathInlineDescriptor { Type = 12 };
            var children = new List<OfficeMathExpression>();
            while (_position <= _text.Length) {
                children.Add(Collapse(ParseSequence(depth + 1)));
                if (_position >= _text.Length) break;
                char terminal = _text[_position++];
                if (terminal == ObjectEnd) break;
            }
            return Build(descriptor, children);
        }

        private static OfficeMathExpression Build(OneNoteMathInlineDescriptor descriptor, IReadOnlyList<OfficeMathExpression> children) {
            OfficeMathExpression Empty() => OfficeMath.Text(string.Empty);
            OfficeMathExpression Child(int index) => index < children.Count ? children[index] : Empty();
            string character = descriptor.Character.HasValue ? ((char)descriptor.Character.Value).ToString() : string.Empty;
            switch (descriptor.Type) {
                case 10: return OfficeMath.Accent(Child(0), string.IsNullOrEmpty(character) ? "^" : character);
                case 11: return OfficeMath.Box(Child(0));
                case 12: return Collapse(children);
                case 13:
                    return OfficeMath.Delimited(Collapse(children), string.IsNullOrEmpty(character) ? "(" : character,
                        descriptor.Character1.HasValue ? ((char)descriptor.Character1.Value).ToString() : ")");
                case 14:
                    return OfficeMath.DelimiterList(
                        string.IsNullOrEmpty(character) ? "(" : character,
                        descriptor.Character1.HasValue ? ((char)descriptor.Character1.Value).ToString() : ")",
                        descriptor.Character2.HasValue ? ((char)descriptor.Character2.Value).ToString() : ",",
                        children.ToArray());
                case 15: return Matrix(OfficeMathKind.EquationArray, descriptor, children);
                case 16: return OfficeMath.Fraction(Child(0), Child(1));
                case 26: return OfficeMath.SlashedFraction(Child(0), Child(1));
                case 17: return OfficeMath.Function(Child(0).ToPlainText(), Child(1));
                case 18: return OfficeMath.LeftSubSuperscript(Child(0), Child(1), Child(2));
                case 30: return OfficeMath.SubSuperscript(Child(0), Child(1), Child(2));
                case 19: return OfficeMath.LowerLimit(Child(0), Child(1));
                case 29: return OfficeMath.Subscript(Child(0), Child(1));
                case 20: return Matrix(OfficeMathKind.Matrix, descriptor, children);
                case 21:
                    if (children.Count >= 3) return OfficeMath.Nary(string.IsNullOrEmpty(character) ? "∑" : character, Child(2), Child(0), Child(1));
                    if (children.Count == 2) return OfficeMath.Nary(string.IsNullOrEmpty(character) ? "∑" : character, Child(1), Child(0));
                    return OfficeMath.Nary(string.IsNullOrEmpty(character) ? "∑" : character, Child(0));
                case 22: return OfficeMath.Operator(string.IsNullOrEmpty(character) ? Child(0).ToPlainText() : character);
                case 23: return OfficeMath.Overbar(Child(0));
                case 24: return OfficeMath.Phantom(Child(0));
                case 25: return children.Count > 1 ? OfficeMath.Radical(Child(1), Child(0)) : OfficeMath.Radical(Child(0));
                case 27: return OfficeMath.Stack(children.ToArray());
                case 28: return OfficeMath.StretchStack(children.ToArray());
                case 31: return OfficeMath.Superscript(Child(0), Child(1));
                case 33: return OfficeMath.UpperLimit(Child(0), Child(1));
                case 32: return OfficeMath.Underbar(Child(0));
                default: return Collapse(children);
            }
        }

        private static OfficeMathExpression Matrix(OfficeMathKind kind, OneNoteMathInlineDescriptor descriptor, IReadOnlyList<OfficeMathExpression> children) {
            int columns = Math.Max(1, Math.Min(children.Count, descriptor.Column ?? (byte)1));
            int rows = Math.Max(1, (int)Math.Ceiling(children.Count / (double)columns));
            var cells = new List<OfficeMathExpression>(rows * columns);
            cells.AddRange(children);
            while (cells.Count < rows * columns) cells.Add(OfficeMath.Text(string.Empty));
            return kind == OfficeMathKind.Matrix
                ? OfficeMath.Matrix(rows, columns, cells.ToArray())
                : OfficeMath.EquationArray(rows, columns, cells.ToArray());
        }

        private static OfficeMathExpression Collapse(IEnumerable<OfficeMathExpression> values) {
            OfficeMathExpression[] expressions = values.Where(value => value != null).ToArray();
            if (expressions.Length == 0) return OfficeMath.Text(string.Empty);
            return expressions.Length == 1 ? expressions[0] : OfficeMath.Row(expressions);
        }

        private static void FlushPlain(ICollection<OfficeMathExpression> output, StringBuilder plain) {
            if (plain.Length == 0) return;
            string value = plain.ToString();
            plain.Clear();
            int position = 0;
            while (position < value.Length) {
                char current = value[position];
                if (IsOperator(current)) {
                    output.Add(OfficeMath.Operator(current.ToString()));
                    position++;
                    continue;
                }

                int start = position;
                if (char.IsLetter(current)) {
                    while (position < value.Length && (char.IsLetterOrDigit(value[position]) || value[position] == '_')) position++;
                    output.Add(OfficeMath.Identifier(value.Substring(start, position - start)));
                } else if (char.IsDigit(current) || (current == '.' && position + 1 < value.Length && char.IsDigit(value[position + 1]))) {
                    bool decimalPoint = false;
                    while (position < value.Length && (char.IsDigit(value[position]) || (!decimalPoint && value[position] == '.'))) {
                        if (value[position] == '.') decimalPoint = true;
                        position++;
                    }
                    output.Add(OfficeMath.Number(value.Substring(start, position - start)));
                } else {
                    while (position < value.Length && !char.IsLetterOrDigit(value[position]) && !IsOperator(value[position]) &&
                           !(value[position] == '.' && position + 1 < value.Length && char.IsDigit(value[position + 1]))) position++;
                    output.Add(OfficeMath.Text(value.Substring(start, position - start)));
                }
            }
        }

        private static bool IsOperator(char value) => "+-=×÷±→<>≤≥≠,.;:*/^_()[]{}|∑∏∫".IndexOf(value) >= 0;
    }
}
