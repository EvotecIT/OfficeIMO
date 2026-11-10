namespace OfficeIMO.Yaml;

public sealed partial class YamlReader {
    private YamlParserState _state = YamlParserState.StreamStart;
    private YamlParserState[] _states = new YamlParserState[16];
    private int _stateCount;
    private int _depth;
    private bool _implicitDocumentAllowed = true;
    private YamlTagDirective[] _tagDirectives = Array.Empty<YamlTagDirective>();
    private int _tagDirectiveCount;
    private YamlSchema _documentSchema = YamlSchema.Core;
    private int _versionMajor = -1;
    private int _versionMinor = -1;

    // End of the last consumed token: the position of empty nodes and implicit ends.
    private int _lastEnd;
    private int _lastEndLine;
    private int _lastEndColumn;

    private void Push(YamlParserState state) {
        if (_stateCount == _states.Length) Array.Resize(ref _states, _states.Length * 2);
        _states[_stateCount++] = state;
    }

    private YamlParserState Pop() => _states[--_stateCount];

    private void Consume(ref YamlScanner scanner, ref YamlToken token) {
        if (token.End >= _lastEnd) {
            _lastEnd = token.End;
            _lastEndLine = token.EndLine;
            _lastEndColumn = token.EndColumn;
        }
        if (_c.RetainMark < 0) _c.RetainMark = token.Start;
        scanner.Skip();
    }

    private YamlException ParseError(string code, string message, in YamlToken token) =>
        new YamlException(code, message, new YamlPosition(token.Start, token.StartLine + 1, token.StartColumn + 1));

    /// <summary>Produces the next event. The state machine keeps its own stack, so nesting never recurses.</summary>
    private void ParseEvent(ref YamlScanner scanner, out YamlEvent ev) {
        ev = default;
        ev.AnchorStart = ev.AnchorEnd = -1;
        switch (_state) {
            case YamlParserState.StreamStart: {
                ref YamlToken token = ref scanner.Peek();
                SetSpan(ref ev, in token);
                ev.Type = YamlTokenType.StreamStart;
                Consume(ref scanner, ref token);
                _state = YamlParserState.ImplicitDocumentStart;
                return;
            }
            case YamlParserState.ImplicitDocumentStart:
            case YamlParserState.DocumentStart:
                ParseDocumentStart(ref scanner, ref ev);
                return;
            case YamlParserState.DocumentContent: {
                ref YamlToken token = ref scanner.Peek();
                switch (token.Kind) {
                    case YamlTokenKind.VersionDirective:
                    case YamlTokenKind.TagDirective:
                    case YamlTokenKind.ReservedDirective:
                    case YamlTokenKind.DocumentStart:
                    case YamlTokenKind.DocumentEnd:
                    case YamlTokenKind.StreamEnd:
                        _state = Pop();
                        EmptyScalar(ref ev);
                        return;
                }
                ParseNode(ref scanner, ref ev, block: true, indentless: false);
                return;
            }
            case YamlParserState.DocumentEnd: {
                ref YamlToken token = ref scanner.Peek();
                ev.Type = YamlTokenType.DocumentEnd;
                if (token.Kind == YamlTokenKind.DocumentEnd) {
                    SetSpan(ref ev, in token);
                    Consume(ref scanner, ref token);
                    _implicitDocumentAllowed = true;
                } else {
                    SetPoint(ref ev, _lastEnd, _lastEndLine, _lastEndColumn);
                    ev.Implicit = true;
                    _implicitDocumentAllowed = false;
                }
                _state = YamlParserState.DocumentStart;
                return;
            }
            case YamlParserState.BlockNode:
                ParseNode(ref scanner, ref ev, block: true, indentless: false);
                return;
            case YamlParserState.BlockSequenceEntry: {
                ref YamlToken token = ref scanner.Peek();
                if (token.Kind == YamlTokenKind.BlockEntry) {
                    Consume(ref scanner, ref token);
                    ref YamlToken next = ref scanner.Peek();
                    if (next.Kind == YamlTokenKind.BlockEntry || next.Kind == YamlTokenKind.BlockEnd) {
                        EmptyScalar(ref ev);
                        return;
                    }
                    Push(YamlParserState.BlockSequenceEntry);
                    ParseNode(ref scanner, ref ev, block: true, indentless: false);
                    return;
                }
                if (token.Kind == YamlTokenKind.BlockEnd) {
                    _state = Pop();
                    EndCollection(ref ev, YamlTokenType.SequenceEnd, in token);
                    Consume(ref scanner, ref token);
                    return;
                }
                throw ParseError(YamlErrorCodes.UnexpectedToken, "Expected a block sequence entry ('- ') or the end of the sequence.", in token);
            }
            case YamlParserState.IndentlessSequenceEntry: {
                ref YamlToken token = ref scanner.Peek();
                if (token.Kind == YamlTokenKind.BlockEntry) {
                    Consume(ref scanner, ref token);
                    ref YamlToken next = ref scanner.Peek();
                    if (next.Kind is YamlTokenKind.BlockEntry or YamlTokenKind.Key or YamlTokenKind.Value or YamlTokenKind.BlockEnd) {
                        EmptyScalar(ref ev);
                        return;
                    }
                    Push(YamlParserState.IndentlessSequenceEntry);
                    ParseNode(ref scanner, ref ev, block: true, indentless: false);
                    return;
                }
                _state = Pop();
                ev.Type = YamlTokenType.SequenceEnd;
                ev.Style = YamlStyle.Block;
                SetPoint(ref ev, _lastEnd, _lastEndLine, _lastEndColumn);
                _depth--;
                return;
            }
            case YamlParserState.BlockMappingKey: {
                ref YamlToken token = ref scanner.Peek();
                if (token.Kind == YamlTokenKind.Key) {
                    Consume(ref scanner, ref token);
                    ref YamlToken next = ref scanner.Peek();
                    if (next.Kind is YamlTokenKind.Key or YamlTokenKind.Value or YamlTokenKind.BlockEnd) {
                        _state = YamlParserState.BlockMappingValue;
                        EmptyScalar(ref ev);
                        ev.IsKey = true;
                        return;
                    }
                    Push(YamlParserState.BlockMappingValue);
                    ParseNode(ref scanner, ref ev, block: true, indentless: true);
                    return;
                }
                if (token.Kind == YamlTokenKind.Value) {
                    _state = YamlParserState.BlockMappingValue;
                    EmptyScalarAt(ref ev, token.Start, token.StartLine, token.StartColumn);
                    ev.IsKey = true;
                    return;
                }
                if (token.Kind == YamlTokenKind.BlockEnd) {
                    _state = Pop();
                    EndCollection(ref ev, YamlTokenType.MappingEnd, in token);
                    Consume(ref scanner, ref token);
                    return;
                }
                throw ParseError(YamlErrorCodes.UnexpectedToken, "Expected a mapping key or the end of the mapping.", in token);
            }
            case YamlParserState.BlockMappingValue: {
                ref YamlToken token = ref scanner.Peek();
                if (token.Kind == YamlTokenKind.Value) {
                    Consume(ref scanner, ref token);
                    ref YamlToken next = ref scanner.Peek();
                    if (next.Kind is YamlTokenKind.Key or YamlTokenKind.Value or YamlTokenKind.BlockEnd) {
                        _state = YamlParserState.BlockMappingKey;
                        EmptyScalar(ref ev);
                        return;
                    }
                    Push(YamlParserState.BlockMappingKey);
                    ParseNode(ref scanner, ref ev, block: true, indentless: true);
                    return;
                }
                _state = YamlParserState.BlockMappingKey;
                EmptyScalar(ref ev);
                return;
            }
            case YamlParserState.FlowSequenceFirstEntry:
            case YamlParserState.FlowSequenceEntry: {
                bool first = _state == YamlParserState.FlowSequenceFirstEntry;
                ref YamlToken token = ref scanner.Peek();
                if (token.Kind != YamlTokenKind.FlowSequenceEnd) {
                    if (!first) {
                        if (token.Kind != YamlTokenKind.FlowEntry) throw ParseError(YamlErrorCodes.UnexpectedToken, "Expected ',' or ']' in a flow sequence.", in token);
                        Consume(ref scanner, ref token);
                        token = ref scanner.Peek();
                    }
                    if (token.Kind == YamlTokenKind.Key) {
                        _state = YamlParserState.FlowSequenceEntryMappingKey;
                        StartPair(ref ev, in token);
                        return;
                    }
                    if (token.Kind == YamlTokenKind.Value) {
                        _state = YamlParserState.FlowSequenceEntryMappingEmptyKey;
                        StartPair(ref ev, in token);
                        return;
                    }
                    if (token.Kind != YamlTokenKind.FlowSequenceEnd) {
                        Push(YamlParserState.FlowSequenceEntry);
                        ParseNode(ref scanner, ref ev, block: false, indentless: false);
                        return;
                    }
                }
                _state = Pop();
                EndCollection(ref ev, YamlTokenType.SequenceEnd, in token);
                ev.Style = YamlStyle.Flow;
                Consume(ref scanner, ref token);
                return;
            }
            case YamlParserState.FlowSequenceEntryMappingKey: {
                ref YamlToken token = ref scanner.Peek();
                Consume(ref scanner, ref token);
                ref YamlToken next = ref scanner.Peek();
                if (next.Kind is YamlTokenKind.Value or YamlTokenKind.FlowEntry or YamlTokenKind.FlowSequenceEnd) {
                    _state = YamlParserState.FlowSequenceEntryMappingValue;
                    EmptyScalar(ref ev);
                    ev.IsKey = true;
                    return;
                }
                Push(YamlParserState.FlowSequenceEntryMappingValue);
                ParseNode(ref scanner, ref ev, block: false, indentless: false);
                return;
            }
            case YamlParserState.FlowSequenceEntryMappingEmptyKey: {
                ref YamlToken token = ref scanner.Peek();
                _state = YamlParserState.FlowSequenceEntryMappingValue;
                EmptyScalarAt(ref ev, token.Start, token.StartLine, token.StartColumn);
                ev.IsKey = true;
                return;
            }
            case YamlParserState.FlowSequenceEntryMappingValue: {
                ref YamlToken token = ref scanner.Peek();
                if (token.Kind == YamlTokenKind.Value) {
                    Consume(ref scanner, ref token);
                    ref YamlToken next = ref scanner.Peek();
                    if (next.Kind is YamlTokenKind.FlowEntry or YamlTokenKind.FlowSequenceEnd) {
                        _state = YamlParserState.FlowSequenceEntryMappingEnd;
                        EmptyScalar(ref ev);
                        return;
                    }
                    Push(YamlParserState.FlowSequenceEntryMappingEnd);
                    ParseNode(ref scanner, ref ev, block: false, indentless: false);
                    return;
                }
                _state = YamlParserState.FlowSequenceEntryMappingEnd;
                EmptyScalar(ref ev);
                return;
            }
            case YamlParserState.FlowSequenceEntryMappingEnd:
                _state = YamlParserState.FlowSequenceEntry;
                ev.Type = YamlTokenType.MappingEnd;
                ev.Style = YamlStyle.Flow;
                SetPoint(ref ev, _lastEnd, _lastEndLine, _lastEndColumn);
                _depth--;
                return;
            case YamlParserState.FlowMappingFirstKey:
            case YamlParserState.FlowMappingKey: {
                bool first = _state == YamlParserState.FlowMappingFirstKey;
                ref YamlToken token = ref scanner.Peek();
                if (token.Kind != YamlTokenKind.FlowMappingEnd) {
                    if (!first) {
                        if (token.Kind != YamlTokenKind.FlowEntry) throw ParseError(YamlErrorCodes.UnexpectedToken, "Expected ',' or '}' in a flow mapping.", in token);
                        Consume(ref scanner, ref token);
                        token = ref scanner.Peek();
                    }
                    if (token.Kind == YamlTokenKind.Key) {
                        Consume(ref scanner, ref token);
                        ref YamlToken next = ref scanner.Peek();
                        if (next.Kind is YamlTokenKind.Value or YamlTokenKind.FlowEntry or YamlTokenKind.FlowMappingEnd) {
                            _state = YamlParserState.FlowMappingValue;
                            EmptyScalar(ref ev);
                            ev.IsKey = true;
                            return;
                        }
                        Push(YamlParserState.FlowMappingValue);
                        ParseNode(ref scanner, ref ev, block: false, indentless: false);
                        return;
                    }
                    if (token.Kind == YamlTokenKind.Value) {
                        _state = YamlParserState.FlowMappingValue;
                        EmptyScalarAt(ref ev, token.Start, token.StartLine, token.StartColumn);
                        ev.IsKey = true;
                        return;
                    }
                    if (token.Kind != YamlTokenKind.FlowMappingEnd) {
                        Push(YamlParserState.FlowMappingValue);
                        ParseNode(ref scanner, ref ev, block: false, indentless: false);
                        return;
                    }
                }
                _state = Pop();
                EndCollection(ref ev, YamlTokenType.MappingEnd, in token);
                ev.Style = YamlStyle.Flow;
                Consume(ref scanner, ref token);
                return;
            }
            case YamlParserState.FlowMappingValue: {
                ref YamlToken token = ref scanner.Peek();
                if (token.Kind == YamlTokenKind.Value) {
                    Consume(ref scanner, ref token);
                    ref YamlToken next = ref scanner.Peek();
                    if (next.Kind is YamlTokenKind.FlowEntry or YamlTokenKind.FlowMappingEnd) {
                        _state = YamlParserState.FlowMappingKey;
                        EmptyScalar(ref ev);
                        return;
                    }
                    Push(YamlParserState.FlowMappingKey);
                    ParseNode(ref scanner, ref ev, block: false, indentless: false);
                    return;
                }
                _state = YamlParserState.FlowMappingKey;
                EmptyScalar(ref ev);
                return;
            }
            default:
                throw new InvalidOperationException("The reader has already reached the end of the stream.");
        }
    }

    private void ParseDocumentStart(ref YamlScanner scanner, ref YamlEvent ev) {
        ref YamlToken token = ref scanner.Peek();
        bool implicitAllowed = _implicitDocumentAllowed;
        while (token.Kind == YamlTokenKind.DocumentEnd) {
            Consume(ref scanner, ref token);
            token = ref scanner.Peek();
            implicitAllowed = true;
        }
        if (token.Kind == YamlTokenKind.StreamEnd) {
            SetSpan(ref ev, in token);
            ev.Type = YamlTokenType.StreamEnd;
            Consume(ref scanner, ref token);
            _state = YamlParserState.End;
            return;
        }
        _tagDirectiveCount = 0;
        _versionMajor = _versionMinor = -1;
        bool version11 = false;
        if (implicitAllowed && token.Kind is not (YamlTokenKind.VersionDirective or YamlTokenKind.TagDirective or YamlTokenKind.ReservedDirective or YamlTokenKind.DocumentStart)) {
            _documentSchema = _options.Schema ?? YamlSchema.Core;
            ev.Type = YamlTokenType.DocumentStart;
            ev.Implicit = true;
            SetPoint(ref ev, token.Start, token.StartLine, token.StartColumn);
            Push(YamlParserState.DocumentEnd);
            _state = YamlParserState.BlockNode;
            return;
        }

        int start = token.Start, startLine = token.StartLine, startColumn = token.StartColumn;
        bool sawVersion = false;
        while (token.Kind is YamlTokenKind.VersionDirective or YamlTokenKind.TagDirective or YamlTokenKind.ReservedDirective) {
            if (token.Kind == YamlTokenKind.VersionDirective) {
                if (sawVersion) throw ParseError(YamlErrorCodes.InvalidDirective, "A document can have only one %YAML directive.", in token);
                if (token.Extra != 1) throw ParseError(YamlErrorCodes.InvalidDirective, "Only YAML version 1.x is supported.", in token);
                sawVersion = true;
                version11 = token.Extra2 == 1;
                _versionMajor = token.Extra;
                _versionMinor = token.Extra2;
            } else if (token.Kind == YamlTokenKind.TagDirective) {
                AddTagDirective(scanner.Text, in token);
            }
            Consume(ref scanner, ref token);
            token = ref scanner.Peek();
        }
        if (token.Kind != YamlTokenKind.DocumentStart) {
            throw ParseError(YamlErrorCodes.InvalidDocument, "Expected a document start marker ('---').", in token);
        }
        _documentSchema = _options.Schema ?? (version11 ? YamlSchema.Yaml11 : YamlSchema.Core);
        ev.Type = YamlTokenType.DocumentStart;
        ev.Start = start;
        ev.StartLine = startLine;
        ev.StartColumn = startColumn;
        ev.End = token.End;
        ev.EndLine = token.EndLine;
        ev.EndColumn = token.EndColumn;
        Consume(ref scanner, ref token);
        Push(YamlParserState.DocumentEnd);
        _state = YamlParserState.DocumentContent;
    }

    private void AddTagDirective(ReadOnlySpan<char> text, in YamlToken token) {
        int b = _c.Base;
        ReadOnlySpan<char> handle = text.Slice(token.ValueStart - b, token.ValueEnd - token.ValueStart);
        for (int i = 0; i < _tagDirectiveCount; i++) {
            if (handle.SequenceEqual(_tagDirectives[i].Handle.AsSpan())) {
                throw ParseError(YamlErrorCodes.InvalidDirective, "A document cannot declare the same %TAG handle twice.", in token);
            }
        }
        string prefix = DecodePercent(text.Slice(token.Extra - b, token.Extra2 - token.Extra));
        if (_tagDirectiveCount == _tagDirectives.Length) Array.Resize(ref _tagDirectives, Math.Max(4, _tagDirectives.Length * 2));
        _tagDirectives[_tagDirectiveCount++] = new YamlTagDirective(handle.ToString(), prefix);
    }

    private void ParseNode(ref YamlScanner scanner, ref YamlEvent ev, bool block, bool indentless) {
        ref YamlToken token = ref scanner.Peek();
        if (token.Kind == YamlTokenKind.Alias) {
            _state = Pop();
            SetSpan(ref ev, in token);
            ev.Type = YamlTokenType.Alias;
            ev.AnchorStart = token.ValueStart;
            ev.AnchorEnd = token.ValueEnd;
            Consume(ref scanner, ref token);
            return;
        }

        int start = -1, startLine = 0, startColumn = 0;
        bool hasAnchor = false;
        while (true) {
            if (token.Kind == YamlTokenKind.Anchor && !hasAnchor) {
                hasAnchor = true;
                ev.AnchorStart = token.ValueStart;
                ev.AnchorEnd = token.ValueEnd;
            } else if (token.Kind == YamlTokenKind.Tag && !ev.HasTag) {
                ResolveTag(scanner.Text, ref ev, in token);
            } else {
                break;
            }
            if (start < 0) {
                start = token.Start;
                startLine = token.StartLine;
                startColumn = token.StartColumn;
            }
            Consume(ref scanner, ref token);
            token = ref scanner.Peek();
        }
        bool hasProperties = start >= 0;
        if (!hasProperties) {
            start = token.Start;
            startLine = token.StartLine;
            startColumn = token.StartColumn;
        }
        int propertiesEnd = hasProperties ? _lastEnd : start;
        int propertiesEndLine = hasProperties ? _lastEndLine : startLine;
        int propertiesEndColumn = hasProperties ? _lastEndColumn : startColumn;

        if (token.Kind == YamlTokenKind.Scalar) {
            _state = Pop();
            ev.Type = YamlTokenType.Scalar;
            ev.Style = token.Style;
            ev.Flags = token.Flags;
            ev.Chomping = token.Chomping;
            ev.Indent = token.Extra;
            ev.ValueStart = token.ValueStart;
            ev.ValueEnd = token.ValueEnd;
            ev.IsKey = _state is YamlParserState.BlockMappingValue or YamlParserState.FlowMappingValue or YamlParserState.FlowSequenceEntryMappingValue;
            ev.Start = start;
            ev.StartLine = startLine;
            ev.StartColumn = startColumn;
            ev.End = token.End;
            ev.EndLine = token.EndLine;
            ev.EndColumn = token.EndColumn;
            Consume(ref scanner, ref token);
            return;
        }

        YamlTokenType type;
        YamlStyle style;
        YamlParserState next;
        bool consume = true;
        switch (token.Kind) {
            case YamlTokenKind.FlowSequenceStart:
                type = YamlTokenType.SequenceStart; style = YamlStyle.Flow; next = YamlParserState.FlowSequenceFirstEntry;
                break;
            case YamlTokenKind.FlowMappingStart:
                type = YamlTokenType.MappingStart; style = YamlStyle.Flow; next = YamlParserState.FlowMappingFirstKey;
                break;
            case YamlTokenKind.BlockSequenceStart when block:
                type = YamlTokenType.SequenceStart; style = YamlStyle.Block; next = YamlParserState.BlockSequenceEntry;
                break;
            case YamlTokenKind.BlockMappingStart when block:
                type = YamlTokenType.MappingStart; style = YamlStyle.Block; next = YamlParserState.BlockMappingKey;
                break;
            case YamlTokenKind.BlockEntry when indentless:
                type = YamlTokenType.SequenceStart; style = YamlStyle.Block; next = YamlParserState.IndentlessSequenceEntry;
                consume = false;
                break;
            default:
                if (hasProperties) {
                    _state = Pop();
                    ev.Type = YamlTokenType.Scalar;
                    ev.Style = YamlStyle.Plain;
                    ev.Start = start;
                    ev.StartLine = startLine;
                    ev.StartColumn = startColumn;
                    ev.End = ev.ValueStart = ev.ValueEnd = propertiesEnd;
                    ev.EndLine = propertiesEndLine;
                    ev.EndColumn = propertiesEndColumn;
                    ev.IsKey = _state is YamlParserState.BlockMappingValue or YamlParserState.FlowMappingValue or YamlParserState.FlowSequenceEntryMappingValue;
                    return;
                }
                throw ParseError(YamlErrorCodes.UnexpectedToken, "Expected a node (a scalar, a collection or an alias).", in token);
        }

        if (++_depth > _options.MaxDepth) {
            throw new YamlLimitExceededException(nameof(YamlReaderOptions.MaxDepth), _options.MaxDepth, new YamlPosition(token.Start, token.StartLine + 1, token.StartColumn + 1));
        }
        _state = next;
        ev.Type = type;
        ev.Style = style;
        ev.Start = start;
        ev.StartLine = startLine;
        ev.StartColumn = startColumn;
        if (style == YamlStyle.Flow) {
            ev.End = token.End;
            ev.EndLine = token.EndLine;
            ev.EndColumn = token.EndColumn;
        } else {
            ev.End = propertiesEnd;
            ev.EndLine = propertiesEndLine;
            ev.EndColumn = propertiesEndColumn;
        }
        if (consume) Consume(ref scanner, ref token);
    }

    private void ResolveTag(ReadOnlySpan<char> text, ref YamlEvent ev, in YamlToken token) {
        int b = _c.Base;
        ev.HasTag = true;
        ev.TagFlags = token.Flags;
        ev.TagSuffixStart = token.ValueStart;
        ev.TagSuffixEnd = token.ValueEnd;
        if ((token.Flags & YamlTokenFlags.Verbatim) != 0) {
            ev.TagPrefix = string.Empty;
            return;
        }
        ReadOnlySpan<char> handle = text.Slice(token.Start - b, token.Extra - token.Start);
        for (int i = 0; i < _tagDirectiveCount; i++) {
            if (handle.SequenceEqual(_tagDirectives[i].Handle.AsSpan())) {
                ev.TagPrefix = _tagDirectives[i].Prefix;
                return;
            }
        }
        if (handle.Length == 1) {
            ev.TagPrefix = "!";
            return;
        }
        if (handle.Length == 2 && handle[1] == '!') {
            ev.TagPrefix = YamlSchemaResolver.TagPrefix;
            return;
        }
        throw ParseError(YamlErrorCodes.InvalidTag, "The tag handle '" + handle.ToString() + "' is not declared by a %TAG directive.", in token);
    }

    private void StartPair(ref YamlEvent ev, in YamlToken token) {
        if (++_depth > _options.MaxDepth) {
            throw new YamlLimitExceededException(nameof(YamlReaderOptions.MaxDepth), _options.MaxDepth, new YamlPosition(token.Start, token.StartLine + 1, token.StartColumn + 1));
        }
        ev.Type = YamlTokenType.MappingStart;
        ev.Style = YamlStyle.Flow;
        ev.Implicit = true;
        SetPoint(ref ev, token.Start, token.StartLine, token.StartColumn);
    }

    private void EndCollection(ref YamlEvent ev, YamlTokenType type, in YamlToken token) {
        ev.Type = type;
        ev.Style = YamlStyle.Block;
        SetSpan(ref ev, in token);
        _depth--;
    }

    private void EmptyScalar(ref YamlEvent ev) => EmptyScalarAt(ref ev, _lastEnd, _lastEndLine, _lastEndColumn);

    private static void EmptyScalarAt(ref YamlEvent ev, int offset, int line, int column) {
        ev.Type = YamlTokenType.Scalar;
        ev.Style = YamlStyle.Plain;
        SetPoint(ref ev, offset, line, column);
        ev.ValueStart = ev.ValueEnd = offset;
    }

    private static void SetPoint(ref YamlEvent ev, int offset, int line, int column) {
        ev.Start = ev.End = offset;
        ev.StartLine = ev.EndLine = line;
        ev.StartColumn = ev.EndColumn = column;
    }

    private static void SetSpan(ref YamlEvent ev, in YamlToken token) {
        ev.Start = token.Start;
        ev.StartLine = token.StartLine;
        ev.StartColumn = token.StartColumn;
        ev.End = token.End;
        ev.EndLine = token.EndLine;
        ev.EndColumn = token.EndColumn;
    }

    /// <summary>Decodes %XX escapes (UTF-8 bytes) in tag text.</summary>
    internal static string DecodePercent(ReadOnlySpan<char> text) {
        if (text.IndexOf('%') < 0) return text.ToString();
        var bytes = new List<byte>(text.Length);
        int i = 0;
        while (i < text.Length) {
            if (text[i] == '%' && i + 2 < text.Length) {
                bytes.Add((byte)(YamlChars.HexValue(text[i + 1]) * 16 + YamlChars.HexValue(text[i + 2])));
                i += 3;
                continue;
            }
            int run = text.Slice(i).IndexOf('%');
            int length = run < 0 ? text.Length - i : Math.Max(run, 1);
            bytes.AddRange(Encoding.UTF8.GetBytes(text.Slice(i, length).ToString()));
            i += length;
        }
        return Encoding.UTF8.GetString(bytes.ToArray());
    }
}
