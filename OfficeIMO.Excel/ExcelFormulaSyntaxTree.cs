using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;

namespace OfficeIMO.Excel {
    /// <summary>Base node in a parsed Excel formula reference tree.</summary>
    public abstract class ExcelFormulaSyntaxNode {
        private protected ExcelFormulaSyntaxNode(string text) { Text = text; }

        /// <summary>Exact authored text represented by this node.</summary>
        public string Text { get; }
    }

    /// <summary>Formula text that is neither a reference nor a string literal.</summary>
    public sealed class ExcelFormulaTextSyntax : ExcelFormulaSyntaxNode {
        internal ExcelFormulaTextSyntax(string text) : base(text) { }
    }

    /// <summary>Quoted Excel formula string literal, including authored quotes.</summary>
    public sealed class ExcelFormulaStringSyntax : ExcelFormulaSyntaxNode {
        internal ExcelFormulaStringSyntax(string text) : base(text) { }
    }

    /// <summary>Parsed function identifier immediately followed by an opening parenthesis.</summary>
    public sealed class ExcelFormulaFunctionSyntax : ExcelFormulaSyntaxNode {
        internal ExcelFormulaFunctionSyntax(string text, string name) : base(text) { Name = name; }

        /// <summary>Authored function name, including a compatibility prefix such as <c>_xlfn.</c>.</summary>
        public string Name { get; }
    }

    /// <summary>Parsed cell, range, row, or column reference in a formula.</summary>
    public sealed class ExcelFormulaReferenceSyntax : ExcelFormulaSyntaxNode {
        internal ExcelFormulaReferenceSyntax(string text, ExcelReference reference) : base(text) { Reference = reference; }

        /// <summary>Format-neutral parsed reference.</summary>
        public ExcelReference Reference { get; }
    }

    /// <summary>Parsed workbook- or worksheet-defined name in a formula.</summary>
    public sealed class ExcelFormulaNameSyntax : ExcelFormulaSyntaxNode {
        internal ExcelFormulaNameSyntax(string text, string name) : base(text) { Name = name; }

        /// <summary>Authored name, including a worksheet qualifier when present.</summary>
        public string Name { get; }
    }

    /// <summary>Parsed Excel table structured reference.</summary>
    public sealed class ExcelFormulaStructuredReferenceSyntax : ExcelFormulaSyntaxNode {
        internal ExcelFormulaStructuredReferenceSyntax(string text, string? tableName, string selector) : base(text) {
            TableName = tableName;
            Selector = selector;
        }

        /// <summary>Table name, or null for a row-context reference such as <c>[@Amount]</c>.</summary>
        public string? TableName { get; }

        /// <summary>Authored bracket selector, including its outer brackets.</summary>
        public string Selector { get; }
    }

    /// <summary>
    /// Immutable syntax tree for reference-bearing Excel formula text. It preserves all non-reference
    /// tokens verbatim and provides the common rewriter used by public conversion and search APIs.
    /// </summary>
    public sealed class ExcelFormulaSyntaxTree {
        private static readonly string[] ErrorLiterals = {
            "#GETTING_DATA", "#BLOCKED!", "#CONNECT!", "#UNKNOWN!", "#PYTHON!",
            "#DIV/0!", "#VALUE!", "#SPILL!", "#FIELD!", "#CALC!", "#BUSY!",
            "#NULL!", "#NAME?", "#REF!", "#NUM!", "#N/A"
        };
        private readonly IReadOnlyList<ExcelFormulaSyntaxNode> _nodes;

        private ExcelFormulaSyntaxTree(string text, IReadOnlyList<ExcelFormulaSyntaxNode> nodes) {
            Text = text;
            _nodes = nodes;
        }

        /// <summary>Exact authored formula text.</summary>
        public string Text { get; }

        /// <summary>Ordered syntax nodes preserving the authored formula.</summary>
        public IReadOnlyList<ExcelFormulaSyntaxNode> Nodes => _nodes;

        /// <summary>Parses A1 formula references while preserving all other tokens verbatim.</summary>
        public static ExcelFormulaSyntaxTree Parse(string formula) {
            if (formula == null) throw new ArgumentNullException(nameof(formula));
            var nodes = new List<ExcelFormulaSyntaxNode>();
            int cursor = 0;
            int textStart = 0;
            while (cursor < formula.Length) {
                if (formula[cursor] == '"') {
                    AddText(nodes, formula, textStart, cursor - textStart);
                    int literalStart = cursor++;
                    while (cursor < formula.Length) {
                        if (formula[cursor] != '"') { cursor++; continue; }
                        if (cursor + 1 < formula.Length && formula[cursor + 1] == '"') { cursor += 2; continue; }
                        cursor++;
                        break;
                    }
                    nodes.Add(new ExcelFormulaStringSyntax(formula.Substring(literalStart, cursor - literalStart)));
                    textStart = cursor;
                    continue;
                }

                if (TryReadErrorLiteral(formula, cursor, out int errorLength)) {
                    cursor += errorLength;
                    continue;
                }

                ExcelFormulaReferenceRewriter.TryReadReferenceAt(formula, cursor, out ExcelFormulaReferenceCandidate? match);
                if (TryReadRepeatedQualifiedRange(formula, match, out int repeatedRangeLength, out ExcelReference? repeatedRange)) {
                    AddText(nodes, formula, textStart, cursor - textStart);
                    nodes.Add(new ExcelFormulaReferenceSyntax(
                        formula.Substring(cursor, repeatedRangeLength),
                        repeatedRange!));
                    cursor += repeatedRangeLength;
                    textStart = cursor;
                    continue;
                }
                if (match != null
                    && !IsSpacedFunctionCall(formula, match)) {
                    AddText(nodes, formula, textStart, cursor - textStart);
                    nodes.Add(new ExcelFormulaReferenceSyntax(match.Text, match.Reference));
                    cursor += match.Length;
                    textStart = cursor;
                    continue;
                }

                if (TryReadQualifiedName(formula, cursor, out int qualifiedNameLength)) {
                    AddText(nodes, formula, textStart, cursor - textStart);
                    string name = formula.Substring(cursor, qualifiedNameLength);
                    nodes.Add(new ExcelFormulaNameSyntax(name, name));
                    cursor += qualifiedNameLength;
                    textStart = cursor;
                    continue;
                }

                if (TryReadStructuredReference(formula, cursor, out int structuredLength, out string? tableName, out string selector)) {
                    AddText(nodes, formula, textStart, cursor - textStart);
                    string text = formula.Substring(cursor, structuredLength);
                    nodes.Add(new ExcelFormulaStructuredReferenceSyntax(text, tableName, selector));
                    cursor += structuredLength;
                    textStart = cursor;
                    continue;
                }

                if (TryReadFunctionName(formula, cursor, out int functionLength)) {
                    AddText(nodes, formula, textStart, cursor - textStart);
                    string name = formula.Substring(cursor, functionLength);
                    nodes.Add(new ExcelFormulaFunctionSyntax(name, name));
                    cursor += functionLength;
                    textStart = cursor;
                    continue;
                }

                if (TryReadName(formula, cursor, out int nameLength)) {
                    AddText(nodes, formula, textStart, cursor - textStart);
                    string name = formula.Substring(cursor, nameLength);
                    nodes.Add(new ExcelFormulaNameSyntax(name, name));
                    cursor += nameLength;
                    textStart = cursor;
                    continue;
                }
                cursor++;
            }
            AddText(nodes, formula, textStart, formula.Length - textStart);
            return new ExcelFormulaSyntaxTree(formula, new ReadOnlyCollection<ExcelFormulaSyntaxNode>(nodes));
        }

        private static bool TryReadErrorLiteral(string formula, int start, out int length) {
            if (formula[start] != '#') {
                length = 0;
                return false;
            }

            foreach (string literal in ErrorLiterals) {
                if (start + literal.Length > formula.Length
                    || string.Compare(formula, start, literal, 0, literal.Length, StringComparison.OrdinalIgnoreCase) != 0) {
                    continue;
                }

                int end = start + literal.Length;
                bool deletedReference = string.Equals(literal, "#REF!", StringComparison.OrdinalIgnoreCase);
                if (!deletedReference && end < formula.Length && IsNamePart(formula[end])) {
                    continue;
                }

                length = literal.Length;
                if (deletedReference) {
                    ExcelFormulaReferenceRewriter.TryReadReferenceAt(formula, end, out ExcelFormulaReferenceCandidate? deletedAddress);
                    if (deletedAddress != null) {
                        length += deletedAddress.Length;
                    }
                }
                return true;
            }

            length = 0;
            return false;
        }

        /// <summary>Rewrites reference nodes once while preserving literals and all other syntax.</summary>
        public string Rewrite(Func<ExcelReference, ExcelReference?> rewriter, ExcelReferenceStyle outputStyle = ExcelReferenceStyle.A1, int anchorRow = 1, int anchorColumn = 1) {
            if (rewriter == null) throw new ArgumentNullException(nameof(rewriter));
            var builder = new StringBuilder(Text.Length);
            foreach (ExcelFormulaSyntaxNode node in _nodes) {
                if (node is not ExcelFormulaReferenceSyntax referenceNode) {
                    builder.Append(node.Text);
                    continue;
                }
                ExcelReference? rewritten = rewriter(referenceNode.Reference);
                if (rewritten == null) {
                    builder.Append("#REF!");
                    continue;
                }
                builder.Append(rewritten.ToString(outputStyle, anchorRow, anchorColumn));
                if (referenceNode.Text.EndsWith("#", StringComparison.Ordinal)) builder.Append('#');
            }
            return builder.ToString();
        }

        /// <summary>Converts every parsed reference to A1 or R1C1 notation.</summary>
        public string ConvertReferences(ExcelReferenceStyle outputStyle, int anchorRow = 1, int anchorColumn = 1) =>
            Rewrite(reference => reference, outputStyle, anchorRow, anchorColumn);

        /// <summary>Rewrites defined-name nodes while preserving all other formula syntax.</summary>
        public string RewriteNames(Func<string, string?> rewriter) {
            if (rewriter == null) throw new ArgumentNullException(nameof(rewriter));
            var builder = new StringBuilder(Text.Length);
            IReadOnlyList<ExcelSheet.FormulaLexicalBinding> lexicalBindings = ExcelSheet.GetFormulaLexicalBindings(Text);
            int nodeIndex = 0;
            foreach (ExcelFormulaSyntaxNode node in _nodes) {
                if (node is ExcelFormulaNameSyntax name
                    && !lexicalBindings.Any(binding => binding.Shadows(name.Name, nodeIndex, node.Text.Length))) {
                    builder.Append(rewriter(name.Name) ?? "#REF!");
                }
                else builder.Append(node.Text);
                nodeIndex += node.Text.Length;
            }
            return builder.ToString();
        }

        /// <summary>Rewrites structured table-reference nodes while preserving all other formula syntax.</summary>
        public string RewriteStructuredReferences(Func<string?, string, string?> rewriter) {
            if (rewriter == null) throw new ArgumentNullException(nameof(rewriter));
            var builder = new StringBuilder(Text.Length);
            foreach (ExcelFormulaSyntaxNode node in _nodes) {
                if (node is ExcelFormulaStructuredReferenceSyntax structured) {
                    builder.Append(rewriter(structured.TableName, structured.Selector) ?? "#REF!");
                } else {
                    builder.Append(node.Text);
                }
            }
            return builder.ToString();
        }

        /// <summary>Rewrites table identifiers in structured and bare table-reference nodes.</summary>
        public string RewriteTableNames(Func<string, string?> rewriter) {
            if (rewriter == null) throw new ArgumentNullException(nameof(rewriter));
            var builder = new StringBuilder(Text.Length);
            IReadOnlyList<ExcelSheet.FormulaLexicalBinding> lexicalBindings = ExcelSheet.GetFormulaLexicalBindings(Text);
            int nodeIndex = 0;
            foreach (ExcelFormulaSyntaxNode node in _nodes) {
                if (node is ExcelFormulaStructuredReferenceSyntax structured && structured.TableName != null) {
                    builder.Append(rewriter(structured.TableName) ?? "#REF!").Append(structured.Selector);
                } else if (node is ExcelFormulaNameSyntax name
                    && name.Name.IndexOf('!') < 0
                    && !lexicalBindings.Any(binding => binding.Shadows(name.Name, nodeIndex, node.Text.Length))) {
                    builder.Append(rewriter(name.Name) ?? "#REF!");
                } else {
                    builder.Append(node.Text);
                }
                nodeIndex += node.Text.Length;
            }
            return builder.ToString();
        }

        private static bool TryReadRepeatedQualifiedRange(
            string formula,
            ExcelFormulaReferenceCandidate? first,
            out int length,
            out ExcelReference? reference) {
            length = 0;
            reference = null;
            if (first == null
                || !first.Reference.IsQualified
                || first.Reference.Kind != ExcelReferenceKind.Cell
                || first.HasSpill) {
                return false;
            }

            int separator = first.Index + first.Length;
            if (separator >= formula.Length || formula[separator] != ':') return false;
            ExcelFormulaReferenceRewriter.TryReadReferenceAt(formula, separator + 1, out ExcelFormulaReferenceCandidate? second);
            if (second == null
                || !second.Reference.IsQualified
                || second.Reference.Kind != ExcelReferenceKind.Cell
                || second.HasSpill
                || !string.Equals(
                    first.Reference.Qualifier,
                    second.Reference.Qualifier,
                    StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            string secondAddress = second.Reference.ToString(ExcelReferenceStyle.A1);
            int qualifierLength = secondAddress.LastIndexOf('!') + 1;
            secondAddress = secondAddress.Substring(qualifierLength);
            string normalized = first.Text + ":" + secondAddress;
            if (!ExcelReference.TryParse(normalized, out reference)) return false;
            length = first.Length + 1 + second.Length;
            return true;
        }

        private static bool IsSpacedFunctionCall(string formula, ExcelFormulaReferenceCandidate match) {
            if (match.Reference.IsQualified
                || match.Reference.Kind != ExcelReferenceKind.Cell
                || match.HasSpill
                || match.Reference.Start.ColumnAbsolute
                || match.Reference.Start.RowAbsolute) {
                return false;
            }

            int cursor = match.Index + match.Length;
            int whitespaceStart = cursor;
            while (cursor < formula.Length && char.IsWhiteSpace(formula[cursor])) cursor++;
            return cursor < formula.Length
                && formula[cursor] == '('
                && (cursor == whitespaceStart || ExcelFormulaCapabilities.IsBuiltInFunction(match.Text));
        }

        private static bool TryReadStructuredReference(
            string formula,
            int start,
            out int length,
            out string? tableName,
            out string selector) {
            length = 0;
            tableName = null;
            selector = string.Empty;
            int bracketStart = start;
            if (formula[start] != '[') {
                if (!IsNameStart(formula[start])) return false;
                int nameEnd = start + 1;
                while (nameEnd < formula.Length && IsNamePart(formula[nameEnd])) nameEnd++;
                if (nameEnd >= formula.Length || formula[nameEnd] != '[') return false;
                tableName = formula.Substring(start, nameEnd - start);
                bracketStart = nameEnd;
            }

            int depth = 0;
            for (int index = bracketStart; index < formula.Length; index++) {
                if (formula[index] == '[' && !IsEscapedStructuredCharacter(formula, index)) depth++;
                else if (formula[index] == ']' && !IsEscapedStructuredCharacter(formula, index)) {
                    depth--;
                    if (depth == 0) {
                        length = index - start + 1;
                        selector = formula.Substring(bracketStart, index - bracketStart + 1);
                        return true;
                    }
                }
            }
            return false;
        }

        internal static string RewriteStructuredColumns(
            string selector,
            IReadOnlyDictionary<string, string> renames) {
            if (selector == null) throw new ArgumentNullException(nameof(selector));
            if (renames == null) throw new ArgumentNullException(nameof(renames));
            return RewriteStructuredSelector(selector, rawColumn => {
                string decoded = DecodeStructuredColumnName(rawColumn);
                foreach (KeyValuePair<string, string> rename in renames) {
                    if (string.Equals(decoded, rename.Key, StringComparison.OrdinalIgnoreCase)) {
                        return EncodeStructuredColumnName(rename.Value);
                    }
                }
                return null;
            });
        }

        internal static bool ContainsStructuredColumn(
            string selector,
            IReadOnlyCollection<string> columnNames) {
            if (selector == null) throw new ArgumentNullException(nameof(selector));
            if (columnNames == null) throw new ArgumentNullException(nameof(columnNames));
            bool found = false;
            _ = RewriteStructuredSelector(selector, rawColumn => {
                string decoded = DecodeStructuredColumnName(rawColumn);
                if (columnNames.Any(name => string.Equals(decoded, name, StringComparison.OrdinalIgnoreCase))) found = true;
                return null;
            });
            return found;
        }

        private static string RewriteStructuredSelector(string value, Func<string, string?> rewriteColumn) {
            int[] matchingClosings = GetMatchedStructuredBrackets(value);
            int[] nextOpenings = GetNextMatchedOpenings(matchingClosings);
            var builder = new StringBuilder(value.Length);
            for (int cursor = 0; cursor < value.Length; cursor++) {
                int closing = matchingClosings[cursor];
                if (closing >= 0) {
                    builder.Append('[');
                    if (nextOpenings[cursor + 1] < closing) continue;
                    string content = value.Substring(cursor + 1, closing - cursor - 1);
                    bool rowContext = content.Length > 0 && content[0] == '@';
                    string rawColumn = rowContext ? content.Substring(1) : content;
                    bool areaSpecifier = rawColumn.Length > 0 && rawColumn[0] == '#';
                    string? replacement = areaSpecifier ? null : rewriteColumn(rawColumn);
                    if (rowContext) builder.Append('@');
                    builder.Append(replacement ?? rawColumn);
                    builder.Append(']');
                    cursor = closing;
                    continue;
                }
                builder.Append(value[cursor]);
            }
            return builder.ToString();
        }

        private static int[] GetMatchedStructuredBrackets(string value) {
            int[] matchingClosings = Enumerable.Repeat(-1, value.Length).ToArray();
            var pending = new Stack<int>();
            int apostropheRun = 0;
            for (int index = 0; index < value.Length; index++) {
                char current = value[index];
                if (current == '\'') {
                    apostropheRun++;
                    continue;
                }
                bool escaped = (apostropheRun & 1) != 0;
                apostropheRun = 0;
                if (escaped) continue;
                if (current == '[') {
                    pending.Push(index);
                } else if (current == ']' && pending.Count > 0) {
                    int opening = pending.Pop();
                    matchingClosings[opening] = index;
                }
            }
            return matchingClosings;
        }

        private static int[] GetNextMatchedOpenings(int[] matchingClosings) {
            var nextOpenings = new int[matchingClosings.Length + 1];
            int next = matchingClosings.Length;
            nextOpenings[matchingClosings.Length] = next;
            for (int index = matchingClosings.Length - 1; index >= 0; index--) {
                if (matchingClosings[index] >= 0) next = index;
                nextOpenings[index] = next;
            }
            return nextOpenings;
        }

        internal static bool IsEscapedStructuredCharacter(string value, int position) {
            int apostrophes = 0;
            for (int index = position - 1; index >= 0 && value[index] == '\''; index--) apostrophes++;
            return (apostrophes & 1) != 0;
        }

        internal static string DecodeStructuredColumnName(string value) {
            var builder = new StringBuilder(value.Length);
            for (int index = 0; index < value.Length; index++) {
                if (value[index] == '\'' && index + 1 < value.Length && IsStructuredEscapeTarget(value[index + 1])) index++;
                builder.Append(value[index]);
            }
            return builder.ToString();
        }

        private static string EncodeStructuredColumnName(string value) {
            var builder = new StringBuilder(value.Length);
            foreach (char character in value) {
                if (IsStructuredEscapeTarget(character)) builder.Append('\'');
                builder.Append(character);
            }
            return builder.ToString();
        }

        private static bool IsStructuredEscapeTarget(char character) =>
            character == '\'' || character == '[' || character == ']' || character == '#' || character == '@';

        private static bool TryReadName(string formula, int start, out int length) {
            length = 0;
            if (!IsNameStart(formula[start])
                || (start > 0 && IsNamePart(formula[start - 1]))) return false;
            int index = start + 1;
            while (index < formula.Length && IsNamePart(formula[index])) index++;
            int lookahead = index;
            while (lookahead < formula.Length && char.IsWhiteSpace(formula[lookahead])) lookahead++;
            if (lookahead < formula.Length && formula[lookahead] == '(') return false;
            string candidate = formula.Substring(start, index - start);
            if (string.Equals(candidate, "TRUE", StringComparison.OrdinalIgnoreCase)
                || string.Equals(candidate, "FALSE", StringComparison.OrdinalIgnoreCase)) return false;
            length = index - start;
            return true;
        }

        private static bool TryReadFunctionName(string formula, int start, out int length) {
            length = 0;
            if (!IsNameStart(formula[start])
                || (start > 0 && IsNamePart(formula[start - 1]))) return false;
            int index = start + 1;
            while (index < formula.Length && IsNamePart(formula[index])) index++;
            int lookahead = index;
            while (lookahead < formula.Length && char.IsWhiteSpace(formula[lookahead])) lookahead++;
            if (lookahead >= formula.Length || formula[lookahead] != '(') return false;
            length = index - start;
            return true;
        }

        private static bool TryReadQualifiedName(string formula, int start, out int length) {
            length = 0;
            if (start > 0 && IsNamePart(formula[start - 1])) return false;
            int separator;
            if (formula[start] == '\'') {
                int index = start + 1;
                while (index < formula.Length) {
                    if (formula[index] != '\'') { index++; continue; }
                    if (index + 1 < formula.Length && formula[index + 1] == '\'') { index += 2; continue; }
                    break;
                }
                if (index >= formula.Length || index + 1 >= formula.Length || formula[index + 1] != '!') return false;
                separator = index + 1;
            } else {
                int index = start;
                if (formula[index] == '[') {
                    int close = formula.IndexOf(']', index + 1);
                    if (close < 0) return false;
                    index = close + 1;
                }
                if (index >= formula.Length || !IsNameStart(formula[index])) return false;
                index++;
                while (index < formula.Length && IsNamePart(formula[index])) index++;
                if (index >= formula.Length || formula[index] != '!') return false;
                separator = index;
            }
            int nameStart = separator + 1;
            if (nameStart >= formula.Length || !TryReadName(formula, nameStart, out int nameLength)) return false;
            length = nameStart + nameLength - start;
            return true;
        }

        private static bool IsNameStart(char value) =>
            value == '_' || value == '\\' || char.IsLetter(value);

        private static bool IsNamePart(char value) =>
            IsNameStart(value) || char.IsDigit(value) || value == '.';

        private static void AddText(List<ExcelFormulaSyntaxNode> nodes, string formula, int start, int length) {
            if (length > 0) nodes.Add(new ExcelFormulaTextSyntax(formula.Substring(start, length)));
        }
    }
}
