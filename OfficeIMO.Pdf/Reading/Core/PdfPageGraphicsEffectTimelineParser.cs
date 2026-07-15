using System;
using System.Collections.Generic;

namespace OfficeIMO.Pdf;

internal static class PdfPageGraphicsEffectTimelineParser {
    public static IReadOnlyList<PdfPageDrawingEffectTransition> Parse(
        string content,
        IReadOnlyDictionary<string, PdfPageGraphicsStateResource>? graphicsStates,
        PdfPageDrawingEffect initialEffect,
        double paintOrderBase = 0D,
        double paintOrderScale = 1D,
        double paintOrderOffset = 0D,
        int maxOperations = PdfReadLimits.DefaultMaxContentOperations) {
        if (string.IsNullOrEmpty(content)) return Array.Empty<PdfPageDrawingEffectTransition>();
        return new Parser(content, graphicsStates, initialEffect, paintOrderBase, paintOrderScale, paintOrderOffset, maxOperations).Parse();
    }

    private sealed class Parser {
        private readonly string _content;
        private readonly IReadOnlyDictionary<string, PdfPageGraphicsStateResource>? _graphicsStates;
        private readonly double _paintOrderBase;
        private readonly double _paintOrderScale;
        private readonly double _paintOrderOffset;
        private readonly int _maxOperations;
        private readonly PdfPageDrawingEffect _initialEffect;
        private readonly List<PdfPageDrawingEffectTransition> _transitions = new List<PdfPageDrawingEffectTransition>();
        private readonly Stack<PdfPageDrawingEffect> _stack = new Stack<PdfPageDrawingEffect>();
        private PdfPageDrawingEffect _state;
        private string? _lastName;
        private int _index;
        private int _operationCount;

        public Parser(
            string content,
            IReadOnlyDictionary<string, PdfPageGraphicsStateResource>? graphicsStates,
            PdfPageDrawingEffect initialEffect,
            double paintOrderBase,
            double paintOrderScale,
            double paintOrderOffset,
            int maxOperations) {
            _content = content;
            _graphicsStates = graphicsStates;
            _initialEffect = initialEffect;
            _state = initialEffect;
            _paintOrderBase = paintOrderBase;
            _paintOrderScale = paintOrderScale;
            _paintOrderOffset = paintOrderOffset;
            _maxOperations = maxOperations;
        }

        public IReadOnlyList<PdfPageDrawingEffectTransition> Parse() {
            while (_index < _content.Length) {
                SkipWhitespace();
                if (_index >= _content.Length) break;
                char current = _content[_index];
                if (current == '%') {
                    SkipComment();
                } else if (current == '/') {
                    _lastName = ReadName();
                } else if (current == '(') {
                    SkipLiteralString();
                } else if (current == '<') {
                    SkipAngleObject();
                } else if (current == '[') {
                    SkipArray();
                } else if (IsNumberStart(current)) {
                    SkipNumber();
                } else {
                    int operatorIndex = _index;
                    string op = ReadOperator();
                    if (op.Length == 0) {
                        _index++;
                        continue;
                    }
                    if (++_operationCount > _maxOperations) throw PdfReadLimitException.Create(PdfReadLimitKind.ContentOperations, _maxOperations, _operationCount);
                    ApplyOperator(op, GetPaintOrder(operatorIndex));
                    _lastName = null;
                }
            }
            return _transitions.Count == 0 ? Array.Empty<PdfPageDrawingEffectTransition>() : _transitions.AsReadOnly();
        }

        private void ApplyOperator(string op, double paintOrder) {
            switch (op) {
                case "q":
                    _stack.Push(_state);
                    break;
                case "Q":
                    PdfPageDrawingEffect restored = _stack.Count > 0 ? _stack.Pop() : _initialEffect;
                    if (!SameEffect(_state, restored)) {
                        _state = restored;
                        _transitions.Add(new PdfPageDrawingEffectTransition(paintOrder, _state));
                    }
                    break;
                case "gs":
                    if (_lastName != null && _graphicsStates != null && _graphicsStates.TryGetValue(_lastName, out PdfPageGraphicsStateResource resource)) {
                        PdfPageDrawingEffect updated = _state.Apply(resource);
                        if (!SameEffect(_state, updated)) {
                            _state = updated;
                            _transitions.Add(new PdfPageDrawingEffectTransition(paintOrder, _state));
                        }
                    }
                    break;
                case "BI":
                    SkipInlineImage();
                    break;
            }
        }

        private double GetPaintOrder(int operatorIndex) => _paintOrderBase + ((operatorIndex + _paintOrderOffset) * _paintOrderScale);

        private static bool SameEffect(PdfPageDrawingEffect left, PdfPageDrawingEffect right) =>
            left.BlendMode == right.BlendMode && ReferenceEquals(left.SoftMask, right.SoftMask);

        private string ReadName() {
            _index++;
            int start = _index;
            while (_index < _content.Length && !IsDelimiter(_content[_index])) _index++;
            return PdfSyntax.DecodeName(_content.Substring(start, _index - start));
        }

        private string ReadOperator() {
            int start = _index;
            while (_index < _content.Length && !IsDelimiter(_content[_index])) _index++;
            return _content.Substring(start, _index - start);
        }

        private void SkipWhitespace() {
            while (_index < _content.Length && char.IsWhiteSpace(_content[_index])) _index++;
        }

        private void SkipComment() {
            while (_index < _content.Length && _content[_index] != '\r' && _content[_index] != '\n') _index++;
        }

        private void SkipNumber() {
            _index++;
            while (_index < _content.Length) {
                char value = _content[_index];
                if (!(char.IsDigit(value) || value == '.' || value == '-' || value == '+' || value == 'e' || value == 'E')) break;
                _index++;
            }
        }

        private void SkipLiteralString() {
            int depth = 1;
            bool escaped = false;
            _index++;
            while (_index < _content.Length && depth > 0) {
                char value = _content[_index++];
                if (escaped) escaped = false;
                else if (value == '\\') escaped = true;
                else if (value == '(') depth++;
                else if (value == ')') depth--;
            }
        }

        private void SkipAngleObject() {
            if (_index + 1 < _content.Length && _content[_index + 1] == '<') {
                _index += 2;
                int depth = 1;
                while (_index < _content.Length && depth > 0) {
                    if (_index + 1 < _content.Length && _content[_index] == '<' && _content[_index + 1] == '<') { depth++; _index += 2; }
                    else if (_index + 1 < _content.Length && _content[_index] == '>' && _content[_index + 1] == '>') { depth--; _index += 2; }
                    else if (_content[_index] == '(') SkipLiteralString();
                    else _index++;
                }
                return;
            }
            _index++;
            while (_index < _content.Length && _content[_index] != '>') _index++;
            if (_index < _content.Length) _index++;
        }

        private void SkipArray() {
            int depth = 1;
            _index++;
            while (_index < _content.Length && depth > 0) {
                if (_content[_index] == '(') SkipLiteralString();
                else if (_content[_index] == '<') SkipAngleObject();
                else {
                    if (_content[_index] == '[') depth++;
                    else if (_content[_index] == ']') depth--;
                    _index++;
                }
            }
        }

        private void SkipInlineImage() {
            int length = PdfInlineImageDataScanner.FindLength(_content, _index);
            _index = Math.Min(_content.Length, _index + Math.Max(0, length));
        }

        private static bool IsNumberStart(char value) => value == '-' || value == '+' || value == '.' || char.IsDigit(value);

        private static bool IsDelimiter(char value) =>
            char.IsWhiteSpace(value) || value == '/' || value == '[' || value == ']' || value == '(' || value == ')' || value == '<' || value == '>' || value == '%';
    }
}
