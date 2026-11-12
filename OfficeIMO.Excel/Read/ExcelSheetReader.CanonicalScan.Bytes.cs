using System.Runtime.ExceptionServices;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        private const int MinimumParallelScanBytes = 2 * 1024 * 1024;
        private const int MaximumParallelScanHelpers = 3;

        /// <summary>
        /// Answers <see cref="MeasureCanonicalWorksheet"/> from buffered UTF-8 bytes for a strict XML shape.
        /// True means the scan would return this extent for an unprefixed root; false means only that the
        /// bytes did not decide it, and the caller runs the scan. Comments, processing instructions, CDATA,
        /// DOCTYPE, a prefixed root, non-ASCII names, inner namespace subtleties, references the text proof
        /// does not cover, and every cell the scan would send to the SDK route are left to the scan.
        /// Large parts may walk their rows on helper threads under the read's execution policy.
        /// </summary>
        internal static bool TryMeasureCanonicalWorksheet(byte[] buffer, int length, ExcelExecutionPolicy execution,
            CancellationToken cancellationToken, out WorksheetExtent extent) {
            int helpers = 0;
#if NET8_0_OR_GREATER
            if (length >= MinimumParallelScanBytes && execution.AllowsBackgroundProcessing) {
                helpers = Math.Min(MaximumParallelScanHelpers, Environment.ProcessorCount - 2);
                if (execution.MaxDegreeOfParallelism is int maximumDegree) helpers = Math.Min(helpers, maximumDegree - 1);
                helpers = helpers > 0 ? ExcelParallelReadBudget.Reserve(helpers) : 0;
            }
#endif
            try {
                return TryMeasureCanonicalWorksheet(buffer, length, helpers, cancellationToken, out extent);
            } finally {
                ExcelParallelReadBudget.Release(helpers);
            }
        }

        /// <summary>Measures with the given number of row helpers, whatever the part's size.</summary>
        internal static bool TryMeasureCanonicalWorksheet(byte[] buffer, int length, int helpers,
            CancellationToken cancellationToken, out WorksheetExtent extent) {
            extent = default;
#if NET8_0_OR_GREATER
            if (!ExcelUtf8RangeRowSource.IsCanonicalDocumentTextValid(buffer.AsSpan(0, length))) return false;
#else
            if (!ExcelUtf8RangeRowSource.IsCanonicalTextValidDownLevel(buffer, length, cancellationToken)) return false;
#endif
            cancellationToken.ThrowIfCancellationRequested();
            return new CanonicalByteScan(buffer, length).TryMeasure(helpers, cancellationToken, out extent);
        }

        private sealed class CanonicalByteScan {
            private const int MaximumAttributes = 32;
            private const string XmlNamespaceUri = "http://www.w3.org/XML/1998/namespace";
            private const string XmlnsNamespaceUri = "http://www.w3.org/2000/xmlns/";

            private readonly byte[] _buffer;
            private readonly int _length;
            private readonly int[] _attributeStarts = new int[MaximumAttributes];
            private readonly int[] _attributeEnds = new int[MaximumAttributes];
            private readonly int[] _attributeColons = new int[MaximumAttributes];
            private readonly int[] _valueStarts = new int[MaximumAttributes];
            private readonly int[] _valueEnds = new int[MaximumAttributes];
            private int[] _elementStarts = new int[16];
            private int[] _elementLengths = new int[16];
            private int[] _bindingStarts = new int[8];
            private int[] _bindingLengths = new int[8];
            private int[] _bindingDepths = new int[8];
            private int _bindingCount;
            private int _attributeCount;
            private bool _hasNamespaceWork;

            // Walk state. A row chunk starts inside sheetData at _baseDepth and may not close it.
            private int _baseDepth;
            private int _depth;
            private bool _rootSeen;
            private bool _rootClosed;
            private bool _checkValueText;
            private int _lastRow;
            private int _lastColumn;
            private int _currentRow;
            private int _tags;

            internal CanonicalByteScan(byte[] buffer, int length) {
                _buffer = buffer;
                _length = length;
            }

            internal bool TryMeasure(int helpers, CancellationToken cancellationToken, out WorksheetExtent extent) {
                extent = default;
                int position = _length >= 3 && _buffer[0] == 0xEF && _buffer[1] == 0xBB && _buffer[2] == 0xBF ? 3 : 0;
                if (position + 1 < _length && _buffer[position] == (byte)'<' && _buffer[position + 1] == (byte)'?'
                    && !TrySkipDeclaration(position, out position)) {
                    return false;
                }

                if (!Walk(position, _length, stopAtSheetData: helpers > 0, cancellationToken, out int stopped, out bool atSheetData)) {
                    return false;
                }
                if (atSheetData) {
                    int resume = TryWalkRowsInParallel(stopped, helpers, cancellationToken);
                    if (resume < 0) return false;
                    if (!Walk(resume, _length, stopAtSheetData: false, cancellationToken, out _, out _)) return false;
                }

                cancellationToken.ThrowIfCancellationRequested();
                if (!_rootSeen || !_rootClosed || _depth != 0) return false;
                extent = new WorksheetExtent(_lastRow, _lastColumn);
                return true;
            }

            /// <summary>
            /// Walks sheetData's rows in chunks split at row tags, one chunk per thread, and returns where the
            /// sequential walk resumes: sheetData's end tag, or the body start when no split was found. A
            /// negative result means a chunk left the decision to the XmlReader scan. Splits are taken only at
            /// a '&lt;', which in the accepted shape always starts a tag; a chunk that begins inside other
            /// markup, closes sheetData or ends inside a row fails its own walk.
            /// </summary>
            private int TryWalkRowsInParallel(int bodyStart, int helpers, CancellationToken cancellationToken) {
                int close = _buffer.AsSpan(bodyStart, _length - bodyStart).LastIndexOf("</sheetData"u8);
                if (close < 0) return bodyStart;
                close += bodyStart;
                int chunks = helpers + 1;
                var starts = new List<int>(chunks + 1) { bodyStart };
                for (int index = 1; index < chunks; index++) {
                    int from = Math.Max(starts[starts.Count - 1] + 1, bodyStart + (int)((long)(close - bodyStart) * index / chunks));
                    int split = IndexOfRowTag(from, close);
                    if (split < 0) break;
                    starts.Add(split);
                }
                if (starts.Count < 2) return bodyStart;
                starts.Add(close);

                var scans = new CanonicalByteScan[starts.Count - 1];
                var results = new bool[scans.Length];
                for (int index = 0; index < scans.Length; index++) scans[index] = CreateRowChunk();
                var tasks = new Task[scans.Length - 1];
                for (int index = 1; index < scans.Length; index++) {
                    int chunk = index;
                    tasks[index - 1] = Task.Run(() => results[chunk] = scans[chunk].WalkRowChunk(starts[chunk], starts[chunk + 1], cancellationToken));
                }
                try {
                    results[0] = scans[0].WalkRowChunk(starts[0], starts[1], cancellationToken);
                } finally {
                    try {
                        Task.WaitAll(tasks);
                    } catch (AggregateException failure) {
                        cancellationToken.ThrowIfCancellationRequested();
                        ExceptionDispatchInfo.Capture(failure.InnerExceptions[0]).Throw();
                    }
                }

                for (int index = 0; index < scans.Length; index++) {
                    if (!results[index]) return -1;
                    if (scans[index]._lastRow > _lastRow) _lastRow = scans[index]._lastRow;
                    if (scans[index]._lastColumn > _lastColumn) _lastColumn = scans[index]._lastColumn;
                }
                _currentRow = scans[scans.Length - 1]._currentRow;
                return close;
            }

            private CanonicalByteScan CreateRowChunk() {
                var chunk = new CanonicalByteScan(_buffer, _length) {
                    _baseDepth = _depth,
                    _depth = _depth,
                    _rootSeen = true,
                    _bindingCount = _bindingCount,
                    _bindingStarts = (int[])_bindingStarts.Clone(),
                    _bindingLengths = (int[])_bindingLengths.Clone(),
                    _bindingDepths = (int[])_bindingDepths.Clone()
                };
                if (chunk._elementStarts.Length <= _depth) {
                    chunk._elementStarts = new int[_depth * 2];
                    chunk._elementLengths = new int[_depth * 2];
                }
                return chunk;
            }

            private bool WalkRowChunk(int start, int end, CancellationToken cancellationToken) =>
                Walk(start, end, stopAtSheetData: false, cancellationToken, out _, out _) && _depth == _baseDepth;

            private int IndexOfRowTag(int from, int end) {
                while (from < end) {
                    int relative = _buffer.AsSpan(from, end - from).IndexOf("<row"u8);
                    if (relative < 0) return -1;
                    int match = from + relative;
                    if (match + 4 < end && (_buffer[match + 4] is (byte)'>' or (byte)'/' || IsWhitespace(_buffer[match + 4]))) {
                        return match;
                    }
                    from = match + 4;
                }
                return -1;
            }

            /// <summary>
            /// Walks tags from <paramref name="position"/> to <paramref name="end"/>, the part's end or the '&lt;'
            /// that starts a later chunk. With <paramref name="stopAtSheetData"/> it stops just after the root's
            /// sheetData start tag.
            /// </summary>
            private bool Walk(int position, int end, bool stopAtSheetData, CancellationToken cancellationToken,
                out int stopped, out bool atSheetData) {
                stopped = end;
                atSheetData = false;
                while (true) {
                    int textEnd = IndexOfTagStart(position, end);
                    if (_depth == 0 && !IsWhitespace(position, textEnd)) return false;
                    // The scan sends a whitespace-only value or text node to the SDK route.
                    if (_checkValueText && textEnd > position && IsWhitespace(position, textEnd)) return false;
                    _checkValueText = false;
                    if (textEnd == end) return true;

                    if ((++_tags & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
                    int tagStart = textEnd;
                    if (tagStart + 1 >= _length) return false;
                    byte marker = _buffer[tagStart + 1];
                    if (marker == (byte)'/') {
                        if (_depth == _baseDepth
                            || _depth == 0
                            || !TryReadName(tagStart + 2, out int nameEnd, out _)
                            || !SkipWhitespace(nameEnd, out int close)
                            || _buffer[close] != (byte)'>'
                            || !BytesEqual(tagStart + 2, nameEnd - tagStart - 2, _elementStarts[_depth - 1], _elementLengths[_depth - 1])) {
                            return false;
                        }
                        _depth--;
                        PopBindings(_depth);
                        if (_depth == 0) _rootClosed = true;
                        position = close + 1;
                        continue;
                    }
                    if (marker is (byte)'!' or (byte)'?' || _rootClosed) return false;

                    int nameStart = tagStart + 1;
                    if (!TryReadName(nameStart, out int elementNameEnd, out int elementColon)
                        || !TryReadAttributes(elementNameEnd, out int tagEnd, out bool selfClosing)
                        || _hasNamespaceWork && !TryBindNamespaces(_depth)
                        || elementColon >= 0 && (!_rootSeen || !IsBoundPrefix(nameStart, elementColon))) {
                        return false;
                    }
                    _rootSeen = true;

                    int localStart = elementColon >= 0 ? elementColon + 1 : nameStart;
                    int localLength = elementNameEnd - localStart;
                    if (localLength == 1 && _buffer[localStart] is (byte)'v' or (byte)'t') {
                        _checkValueText = !selfClosing;
                    } else if (localLength == 3 && _buffer[localStart] == (byte)'r'
                               && _buffer[localStart + 1] == (byte)'o' && _buffer[localStart + 2] == (byte)'w') {
                        int r = FindUnprefixedAttribute((byte)'r');
                        if (r < 0) {
                            _currentRow = 0;
                        } else {
                            if (!TryParseDigits(_valueStarts[r], _valueEnds[r], out _currentRow)
                                || _currentRow <= 0 || _currentRow > ExcelA1.MaxRows) {
                                return false;
                            }
                            if (_currentRow > _lastRow) _lastRow = _currentRow;
                        }
                    } else if (localLength == 1 && _buffer[localStart] == (byte)'c') {
                        int t = FindUnprefixedAttribute((byte)'t');
                        int r = FindUnprefixedAttribute((byte)'r');
                        if (t >= 0 && !IsKnownCellType(_valueStarts[t], _valueEnds[t])
                            || r < 0
                            || !TryParseReference(_valueStarts[r], _valueEnds[r], out int row, out int column)
                            || row != _currentRow) {
                            return false;
                        }
                        if (column > _lastColumn) _lastColumn = column;
                    }

                    if (selfClosing) {
                        PopBindings(_depth);
                        if (_depth == 0) _rootClosed = true;
                    } else {
                        if (_depth == _elementStarts.Length) {
                            Array.Resize(ref _elementStarts, _depth * 2);
                            Array.Resize(ref _elementLengths, _depth * 2);
                        }
                        _elementStarts[_depth] = nameStart;
                        _elementLengths[_depth] = elementNameEnd - nameStart;
                        _depth++;
                        if (stopAtSheetData && _depth == 2 && elementColon < 0
                            && localLength == 9 && Matches(localStart, "sheetData")) {
                            stopped = tagEnd + 1;
                            atSheetData = true;
                            return true;
                        }
                    }
                    position = tagEnd + 1;
                }
            }

            /// <summary>Reads an ASCII name with at most one prefix; anything richer is left to the scan.</summary>
            private bool TryReadName(int start, out int end, out int colon) {
                colon = -1;
                end = start;
                if (start >= _length || !IsNameStart(_buffer[start])) return false;
                end++;
                while (end < _length) {
                    byte value = _buffer[end];
                    if (value == (byte)':') {
                        if (colon >= 0 || end + 1 >= _length || !IsNameStart(_buffer[end + 1])) return false;
                        colon = end;
                    } else if (!IsNameCharacter(value)) {
                        break;
                    }
                    end++;
                }
                return end < _length;
            }

            private bool TryReadAttributes(int position, out int tagEnd, out bool selfClosing) {
                tagEnd = -1;
                selfClosing = false;
                _attributeCount = 0;
                _hasNamespaceWork = false;
                while (true) {
                    int whitespaceStart = position;
                    if (!SkipWhitespace(position, out position)) return false;
                    byte value = _buffer[position];
                    if (value == (byte)'>') {
                        tagEnd = position;
                        break;
                    }
                    if (value == (byte)'/') {
                        if (position + 1 >= _length || _buffer[position + 1] != (byte)'>') return false;
                        tagEnd = position + 1;
                        selfClosing = true;
                        break;
                    }
                    if (position == whitespaceStart || _attributeCount == MaximumAttributes) return false;

                    int nameStart = position;
                    if (!TryReadName(nameStart, out int nameEnd, out int colon)
                        || !SkipWhitespace(nameEnd, out position)
                        || _buffer[position] != (byte)'='
                        || !SkipWhitespace(position + 1, out position)
                        || _buffer[position] is not ((byte)'"') and not ((byte)'\'')) {
                        return false;
                    }
                    byte quote = _buffer[position];
                    int valueStart = position + 1;
                    int valueEnd = valueStart;
                    while (true) {
                        if (valueEnd >= _length) return false;
                        byte content = _buffer[valueEnd];
                        if (content == quote) break;
                        if (content == (byte)'<') return false;
                        valueEnd++;
                    }

                    int nameLength = nameEnd - nameStart;
                    if (colon >= 0 || nameLength == 5 && _buffer[nameStart] == (byte)'x' && Matches(nameStart, "xmlns")) {
                        _hasNamespaceWork = true;
                    }
                    for (int index = 0; index < _attributeCount; index++) {
                        if (BytesEqual(_attributeStarts[index], _attributeEnds[index] - _attributeStarts[index], nameStart, nameLength)) {
                            return false;
                        }
                    }
                    _attributeStarts[_attributeCount] = nameStart;
                    _attributeEnds[_attributeCount] = nameEnd;
                    _attributeColons[_attributeCount] = colon;
                    _valueStarts[_attributeCount] = valueStart;
                    _valueEnds[_attributeCount] = valueEnd;
                    _attributeCount++;
                    position = valueEnd + 1;
                }
                return true;
            }

            /// <summary>
            /// Records this element's prefix declarations and checks its attributes' prefixes. Prefixes in the
            /// reserved xml range, references in namespace values and two prefixed attributes sharing a local
            /// name are left to the scan rather than resolved here.
            /// </summary>
            private bool TryBindNamespaces(int depth) {
                for (int index = 0; index < _attributeCount; index++) {
                    int start = _attributeStarts[index];
                    int colon = _attributeColons[index];
                    if (colon < 0) {
                        if (_attributeEnds[index] - start == 5 && Matches(start, "xmlns")
                            && !IsUsableNamespaceValue(index, allowEmpty: true)) {
                            return false;
                        }
                        continue;
                    }
                    if (colon - start == 5 && Matches(start, "xmlns")) {
                        if (StartsWithXml(colon + 1, _attributeEnds[index]) || !IsUsableNamespaceValue(index, allowEmpty: false)) {
                            return false;
                        }
                        if (_bindingCount == _bindingStarts.Length) {
                            Array.Resize(ref _bindingStarts, _bindingCount * 2);
                            Array.Resize(ref _bindingLengths, _bindingCount * 2);
                            Array.Resize(ref _bindingDepths, _bindingCount * 2);
                        }
                        _bindingStarts[_bindingCount] = colon + 1;
                        _bindingLengths[_bindingCount] = _attributeEnds[index] - colon - 1;
                        _bindingDepths[_bindingCount] = depth;
                        _bindingCount++;
                    }
                }

                for (int index = 0; index < _attributeCount; index++) {
                    int start = _attributeStarts[index];
                    int colon = _attributeColons[index];
                    if (colon < 0 || colon - start == 5 && Matches(start, "xmlns")) continue;
                    if (colon - start == 3 && Matches(start, "xml")) {
                        if (_attributeEnds[index] - colon - 1 != 5 || !Matches(colon + 1, "space")
                            || !ValueEquals(index, "default") && !ValueEquals(index, "preserve")) {
                            return false;
                        }
                        continue;
                    }
                    if (!IsBoundPrefix(start, colon)) return false;
                    for (int other = 0; other < index; other++) {
                        int otherColon = _attributeColons[other];
                        if (otherColon >= 0
                            && !StartsWithXml(_attributeStarts[other], otherColon)
                            && _buffer.AsSpan(colon + 1, _attributeEnds[index] - colon - 1)
                                .SequenceEqual(_buffer.AsSpan(otherColon + 1, _attributeEnds[other] - otherColon - 1))) {
                            return false;
                        }
                    }
                }
                return true;
            }

            private bool IsUsableNamespaceValue(int index, bool allowEmpty) {
                int start = _valueStarts[index];
                int end = _valueEnds[index];
                return (allowEmpty || end > start)
                    && _buffer.AsSpan(start, end - start).IndexOf((byte)'&') < 0
                    && !ValueEquals(index, XmlNamespaceUri)
                    && !ValueEquals(index, XmlnsNamespaceUri);
            }

            private bool IsBoundPrefix(int start, int colon) {
                if (StartsWithXml(start, colon)) return false;
                ReadOnlySpan<byte> prefix = _buffer.AsSpan(start, colon - start);
                for (int index = _bindingCount - 1; index >= 0; index--) {
                    if (prefix.SequenceEqual(_buffer.AsSpan(_bindingStarts[index], _bindingLengths[index]))) return true;
                }
                return false;
            }

            private void PopBindings(int depth) {
                while (_bindingCount > 0 && _bindingDepths[_bindingCount - 1] >= depth) _bindingCount--;
            }

            private int FindUnprefixedAttribute(byte name) {
                for (int index = 0; index < _attributeCount; index++) {
                    if (_attributeEnds[index] - _attributeStarts[index] == 1 && _buffer[_attributeStarts[index]] == name) {
                        return index;
                    }
                }
                return -1;
            }

            private bool IsKnownCellType(int start, int end) {
                int length = end - start;
                if (length == 1) return _buffer[start] is (byte)'n' or (byte)'s' or (byte)'b' or (byte)'e' or (byte)'d';
                return length == 3 && Matches(start, "str") || length == 9 && Matches(start, "inlineStr");
            }

            /// <summary>Digits only; the scan reads anything else (references, spaces, signs) as no row.</summary>
            private bool TryParseDigits(int start, int end, out int value) {
                value = 0;
                if (start == end) return false;
                for (int index = start; index < end; index++) {
                    int digit = _buffer[index] - (byte)'0';
                    if ((uint)digit > 9U || value > (int.MaxValue - digit) / 10) return false;
                    value = value * 10 + digit;
                }
                return true;
            }

            /// <summary>Letters then digits, the reference form the scan's fast parser reads in place.</summary>
            private bool TryParseReference(int start, int end, out int row, out int column) {
                row = 0;
                column = 0;
                int index = start;
                for (; index < end; index++) {
                    int letter = (_buffer[index] | 0x20) - (byte)'a';
                    if ((uint)letter > 25U) break;
                    if (column > ExcelA1.MaxColumns) return false;
                    column = column * 26 + letter + 1;
                }
                return index > start
                    && TryParseDigits(index, end, out row)
                    && row > 0 && row <= ExcelA1.MaxRows
                    && column > 0 && column <= ExcelA1.MaxColumns;
            }

            private bool TrySkipDeclaration(int start, out int next) {
                next = start;
                int end = _buffer.AsSpan(start, _length - start).IndexOf("?>"u8);
                if (end < 0 || start + 6 > _length || !Matches(start + 2, "xml") || !IsWhitespace(_buffer[start + 5])) {
                    return false;
                }
                end += start;
                int position = start + 5;
                int stage = 0;
                while (true) {
                    int whitespaceStart = position;
                    while (position < end && IsWhitespace(_buffer[position])) position++;
                    if (position == end) break;
                    if (position == whitespaceStart) return false;
                    int nameStart = position;
                    while (position < end && IsNameCharacter(_buffer[position])) position++;
                    int nameLength = position - nameStart;
                    while (position < end && IsWhitespace(_buffer[position])) position++;
                    if (position >= end || _buffer[position++] != (byte)'=') return false;
                    while (position < end && IsWhitespace(_buffer[position])) position++;
                    if (position >= end || _buffer[position] is not ((byte)'"') and not ((byte)'\'')) return false;
                    byte quote = _buffer[position++];
                    int valueStart = position;
                    while (position < end && _buffer[position] != quote) position++;
                    if (position >= end) return false;
                    ReadOnlySpan<byte> value = _buffer.AsSpan(valueStart, position - valueStart);
                    position++;
                    if (stage == 0 && nameLength == 7 && Matches(nameStart, "version")) {
                        if (!value.SequenceEqual("1.0"u8)) return false;
                        stage = 1;
                    } else if (stage == 1 && nameLength == 8 && Matches(nameStart, "encoding")) {
                        if (!IsUtf8EncodingName(value)) return false;
                        stage = 2;
                    } else if (stage is 1 or 2 && nameLength == 10 && Matches(nameStart, "standalone")) {
                        if (!value.SequenceEqual("yes"u8) && !value.SequenceEqual("no"u8)) return false;
                        stage = 3;
                    } else {
                        return false;
                    }
                }
                if (stage == 0) return false;
                next = end + 2;
                return true;
            }

            private static bool IsUtf8EncodingName(ReadOnlySpan<byte> value) {
                if (value.Length is not 4 and not 5) return false;
                Span<byte> lower = stackalloc byte[5];
                for (int index = 0; index < value.Length; index++) lower[index] = (byte)(value[index] | 0x20);
                return lower.Slice(0, value.Length).SequenceEqual("utf-8"u8) || lower.Slice(0, value.Length).SequenceEqual("utf8"u8);
            }

            private bool SkipWhitespace(int position, out int next) {
                while (position < _length && IsWhitespace(_buffer[position])) position++;
                next = position;
                return position < _length;
            }

            private bool IsWhitespace(int start, int end) {
                for (int index = start; index < end; index++) {
                    if (!IsWhitespace(_buffer[index])) return false;
                }
                return true;
            }

            private bool StartsWithXml(int start, int end) =>
                end - start >= 3
                && (_buffer[start] | 0x20) == (byte)'x'
                && (_buffer[start + 1] | 0x20) == (byte)'m'
                && (_buffer[start + 2] | 0x20) == (byte)'l';

            private bool Matches(int start, string text) {
                if (start + text.Length > _length) return false;
                for (int index = 0; index < text.Length; index++) {
                    if (_buffer[start + index] != text[index]) return false;
                }
                return true;
            }

            private bool ValueEquals(int index, string text) =>
                _valueEnds[index] - _valueStarts[index] == text.Length && Matches(_valueStarts[index], text);

            /// <summary>Values and short text end within a few bytes; long text uses the vectorized search.</summary>
            private int IndexOfTagStart(int position, int end) {
                int shortEnd = Math.Min(end, position + 16);
                for (int index = position; index < shortEnd; index++) {
                    if (_buffer[index] == (byte)'<') return index;
                }
                if (shortEnd == end) return end;
                int relative = _buffer.AsSpan(shortEnd, end - shortEnd).IndexOf((byte)'<');
                return relative < 0 ? end : shortEnd + relative;
            }

            private bool BytesEqual(int firstStart, int firstLength, int secondStart, int secondLength) {
                if (firstLength != secondLength) return false;
                for (int index = 0; index < firstLength; index++) {
                    if (_buffer[firstStart + index] != _buffer[secondStart + index]) return false;
                }
                return true;
            }

            private const byte NameStartClass = 1;
            private const byte NameClass = 2;
            private const byte WhitespaceClass = 4;
            private static readonly byte[] ByteClasses = CreateByteClasses();

            private static byte[] CreateByteClasses() {
                var classes = new byte[256];
                for (int value = 'A'; value <= 'Z'; value++) classes[value] = NameStartClass | NameClass;
                for (int value = 'a'; value <= 'z'; value++) classes[value] = NameStartClass | NameClass;
                classes['_'] = NameStartClass | NameClass;
                for (int value = '0'; value <= '9'; value++) classes[value] = NameClass;
                classes['.'] = NameClass;
                classes['-'] = NameClass;
                classes[' '] = classes['\t'] = classes['\r'] = classes['\n'] = WhitespaceClass;
                return classes;
            }

            private static bool IsWhitespace(byte value) => (ByteClasses[value] & WhitespaceClass) != 0;

            private static bool IsNameStart(byte value) => (ByteClasses[value] & NameStartClass) != 0;

            private static bool IsNameCharacter(byte value) => (ByteClasses[value] & NameClass) != 0;
        }
    }
}
