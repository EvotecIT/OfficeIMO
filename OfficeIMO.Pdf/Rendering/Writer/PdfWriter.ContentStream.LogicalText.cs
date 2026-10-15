namespace OfficeIMO.Pdf;

internal sealed partial class ContentStreamBuilder {
    private double _textA = 1, _textB, _textC, _textD = 1, _textE, _textF, _lineE, _lineF;
    private double _textScale = 1, _textLeading, _textWordSpacing, _textCharacterSpacing;
    // Shared with drawing ink measurement so clip bounds enclose the painted text matrix.
    internal const double SyntheticObliqueShear = 1D / 3D;
    private bool _syntheticOblique, _hasTextMatrix;
    // The PDF line matrix no longer equals the tracked line origin: a repositioning was
    // folded into a TJ adjustment, so later relative moves must restate the text matrix.
    private bool _isolatedText;
    private readonly Stack<(double Scale, double Leading, double WordSpacing, double CharacterSpacing, bool SyntheticOblique)> _textStates = new();

    // Text-state cache, valid only inside the current text object while every byte appended
    // since the builder's last write is marked-content syntax.
    private bool _inText, _positionKnown, _wordSpacingKnown, _characterSpacingKnown;
    private string? _fontResource;
    private double _fontSize;
    private bool _fontPrecise;
    private PdfColor? _fillColor;
    private int _ownedLength = -1;

    // The last text-show operator this builder wrote. While it is still the tail of the
    // stream, following glyphs and same-baseline moves extend its TJ array.
    private int _showStart = -1, _showEnd = -1, _showElements, _lastNumberStart;
    private bool _showArray, _showEndsWithNumber;
    private double _showFontSize, _lastNumber;
    private bool _lastNumberPrecise;

    private void ResetTrackedTextMatrix() {
        _textA = _textD = 1; _textB = _textC = _textE = _textF = _lineE = _lineF = 0;
        _isolatedText = false;
        _hasTextMatrix = false;
        _positionKnown = false;
    }

    private void TrackTextMatrix(double a, double b, double c, double d, double e, double f) {
        _hasTextMatrix = true;
        _positionKnown = true;
        _textA = a; _textB = b; _textC = c; _textD = d;
        _textE = _lineE = e; _textF = _lineF = f;
    }

    private void AdvanceTrackedText(double advance) {
        _textE += _textA * advance * _textScale;
        _textF += _textB * advance * _textScale;
    }

    private void InvalidateTextStateCache() {
        _fontResource = null;
        _fillColor = null;
        _wordSpacingKnown = _characterSpacingKnown = false;
    }

    // Checks what was appended to the shared stream since this builder's last write.
    // Marked-content operators leave text state unchanged; anything else drops the cache.
    private bool SyncTextState() {
        if (!_inText) return false;
        if (_ownedLength == _sb.Length) return true;
        if (_ownedLength < 0 || _ownedLength > _sb.Length || !IsMarkedContentOnly(_ownedLength)) {
            InvalidateTextStateCache();
            _ownedLength = _sb.Length;
            return false;
        }
        _ownedLength = _sb.Length;
        return true;
    }

    private bool IsMarkedContentOnly(int start) {
        int lineStart = start;
        for (int index = start; index <= _sb.Length; index++) {
            if (index < _sb.Length && _sb[index] != '\n') continue;
            int end = index;
            while (end > lineStart && char.IsWhiteSpace(_sb[end - 1])) end--;
            if (end > lineStart && !(EndsWith(end, "BDC") || EndsWith(end, "BMC") || EndsWith(end, "EMC"))) return false;
            lineStart = index + 1;
        }
        return true;
    }

    private bool EndsWith(int end, string token) {
        if (end < token.Length) return false;
        for (int index = 0; index < token.Length; index++) {
            if (_sb[end - token.Length + index] != token[index]) return false;
        }
        return end == token.Length || char.IsWhiteSpace(_sb[end - token.Length - 1]) || _sb[end - token.Length - 1] == '>';
    }

    private void OwnWrite() {
        if (_inText) _ownedLength = _sb.Length;
    }

    private bool CanExtendShow(double fontSize) =>
        _inText && _showEnd >= 0 && _showEnd == _sb.Length && _showFontSize == fontSize;

    // Opens a TJ array, reopening the previous show operator when it is still the stream tail.
    private void OpenShow(double fontSize) {
        if (CanExtendShow(fontSize)) {
            if (_showArray) {
                _sb.Length -= 5; // "] TJ\n"
            } else {
                _sb.Length -= 4; // " Tj\n"
                _sb.Insert(_showStart, '[');
                _showArray = true;
                _showElements = 1;
                _showEndsWithNumber = false;
            }
            return;
        }

        _showStart = _sb.Length;
        _sb.Append('[');
        _showArray = true;
        _showElements = 0;
        _showEndsWithNumber = false;
        _showFontSize = fontSize;
    }

    private void ShowGlyphElement(int glyphId) {
        if (_showElements > 0 && !_showEndsWithNumber) {
            _sb.Length--; // Extend the previous hex string.
        } else {
            _sb.Append('<');
            _showElements++;
        }
        PdfGlyphRun.AppendGlyphHex(_sb, glyphId);
        _sb.Append('>');
        _showEndsWithNumber = false;
    }

    private void ShowHexElement(string hex) {
        if (hex.Length == 0) return;
        if (_showElements > 0 && !_showEndsWithNumber) {
            _sb.Length--;
        } else {
            _sb.Append('<');
            _showElements++;
        }
        _sb.Append(hex).Append('>');
        _showEndsWithNumber = false;
    }

    // Folded text-matrix moves keep five decimals so large font sizes stay within 0.001 pt.
    private void ShowNumberElement(double adjustment, bool precise = false) {
        if (_showElements > 0 && _showEndsWithNumber) {
            _sb.Length = _lastNumberStart;
            adjustment += _lastNumber;
            precise |= _lastNumberPrecise;
            _showElements--;
            _showEndsWithNumber = false;
        }
        if (Math.Abs(adjustment) < (precise ? 0.000005D : 0.0005D)) return;
        if (_showElements > 0) _sb.Append(' ');
        _lastNumberStart = _showElements > 0 ? _sb.Length - 1 : _sb.Length;
        if (precise) _sb.Append(adjustment.ToString("0.#####", System.Globalization.CultureInfo.InvariantCulture));
        else _sb.Append(F(adjustment));
        _lastNumber = adjustment;
        _lastNumberPrecise = precise;
        _showEndsWithNumber = true;
        _showElements++;
    }

    private void CloseShow() {
        if (_showElements == 0) {
            _sb.Length = _showStart;
            _showEnd = -1;
            return;
        }
        if (_showElements == 1 && !_showEndsWithNumber && _sb[_showStart] == '[') {
            _sb.Remove(_showStart, 1);
            _sb.Append(" Tj\n");
            _showArray = false;
        } else {
            _sb.Append("] TJ\n");
        }
        _showEnd = _sb.Length;
        OwnWrite();
    }

    private const double MaximumFoldedMove1000 = 1000D;

    // A same-line text matrix after a show becomes a TJ adjustment. The tracked line
    // origin then differs from the PDF line matrix, so later relative moves use Tm.
    private bool TryFoldTextMatrix(double a, double b, double c, double d, double e, double f) {
        a = Serialized(a); b = Serialized(b); c = Serialized(c); d = Serialized(d); f = Serialized(f);
        if (!_positionKnown || !CanExtendShow(_showFontSize) || _showFontSize <= 0D ||
            a != _textA || b != _textB || c != _textC || d != _textD || b != 0D || Math.Abs(a) < 0.000001D ||
            Math.Abs(f - _textF) > 0.0000005D || _textScale <= 0D) return false;
        double adjustment = -(e - _textE) / a * 1000D / (_showFontSize * _textScale);
        // Larger moves (tab stops, columns, cells) stay explicit so readers keep separate runs.
        if (double.IsNaN(adjustment) || double.IsInfinity(adjustment) || Math.Abs(adjustment) > MaximumFoldedMove1000) return false;
        if (Math.Abs(adjustment) >= 0.000005D) {
            OpenShow(_showFontSize);
            ShowNumberElement(adjustment, precise: true);
            CloseShow();
        }
        _textE = _lineE = e;
        _textF = _lineF = f;
        _isolatedText = true;
        return true;
    }

    // Writes source-ordered glyphs as one TJ array. ToUnicode owns every glyph the font
    // can pin; only the remaining clusters get their own ActualText-scoped show operator.
    private void WriteLogicalGlyphs(IReadOnlyList<PdfGlyphInfo> glyphs, PdfTextShowCommand command, double fontSize, double textRise) {
        double tracking1000 = (command.Tracking?.GetAdjustment(fontSize / command.FontMetricScale) ?? 0D)
            * 1000D / command.UnitsPerEm * (command.NegativeTracking ? -1D : 1D);
        IPdfGlyphTextPinner? pinner = command.Pinner;
        int currentOffsetY1000 = 0;
        bool open = false;
        double totalAdvance1000 = 0D;
        for (int index = 0; index < glyphs.Count;) {
            int clusterStart = index;
            int logicalEnd = glyphs[index].LogicalClusterStart;
            bool representable = pinner != null;
            do {
                PdfGlyphInfo glyph = glyphs[index++];
                // GSUB multiple substitution and a later ligature can give neighboring
                // glyphs different starts but overlapping source ownership.
                logicalEnd = Math.Max(logicalEnd, Math.Max(glyph.LogicalClusterStart + 1,
                    glyph.TextIndex + glyph.UnicodeText.Length));
            } while (index < glyphs.Count && glyphs[index].LogicalClusterStart < logicalEnd);

            System.Text.StringBuilder? logical = null;
            int[] codes = ResolveClusterCodes(glyphs, clusterStart, index, pinner, ref representable);
            if (!representable) {
                // Glyphs are in paint order; ActualText is the cluster's source (logical) text.
                logical = new System.Text.StringBuilder();
                foreach (PdfGlyphInfo glyph in Enumerable.Range(clusterStart, index - clusterStart)
                    .Select(glyphIndex => glyphs[glyphIndex]).OrderBy(static glyph => glyph.TextIndex)) logical.Append(glyph.UnicodeText);
            }

            bool marked = logical != null && logical.Length != 0;
            if (marked && open) {
                CloseShow();
                open = false;
            }
            if (marked) {
                _sb.Append("/Span << /ActualText ");
                PdfSyntaxEscaper.AppendTextStringCancellable(_sb, logical!.ToString(), default);
                _sb.Append(" >> BDC\n");
            }
            for (int glyphIndex = clusterStart; glyphIndex < index; glyphIndex++) {
                PdfGlyphInfo glyph = glyphs[glyphIndex];
                if (glyph.OffsetY1000 != currentOffsetY1000) {
                    if (open) {
                        CloseShow();
                        open = false;
                    }
                    _sb.Append(F(textRise + glyph.OffsetY1000 * fontSize / 1000D)).Append(" Ts\n");
                    currentOffsetY1000 = glyph.OffsetY1000;
                }
                if (!open) {
                    OpenShow(fontSize);
                    open = true;
                }
                double trackingAdvance1000 = command.TrackingBoundaries != null && command.TrackingBoundaries[glyphIndex] ? tracking1000 : 0D;
                if (glyph.OffsetX1000 != 0) ShowNumberElement(-glyph.OffsetX1000);
                ShowGlyphElement(codes[glyphIndex - clusterStart]);
                ShowNumberElement(glyph.OffsetX1000 + glyph.NominalWidth1000 - glyph.AdvanceWidth1000 - trackingAdvance1000);
                totalAdvance1000 += glyph.AdvanceWidth1000 + trackingAdvance1000;
            }
            if (marked) {
                if (open) {
                    CloseShow();
                    open = false;
                }
                _sb.Append("EMC\n");
                OwnWrite();
            }
        }
        if (open) CloseShow();
        if (currentOffsetY1000 != 0) {
            _sb.Append(F(textRise)).Append(" Ts\n");
            OwnWrite();
        }
        AdvanceTrackedText(totalAdvance1000 * fontSize / 1000D + glyphs.Count * _textCharacterSpacing);
    }

    private int[] _codeBuffer = new int[16];

    // Codes for one cluster: glyph ids, or alias codes for characters sharing a nominal glyph.
    // An unresolved glyph makes the whole cluster depend on ActualText and its glyph ids.
    private int[] ResolveClusterCodes(IReadOnlyList<PdfGlyphInfo> glyphs, int start, int end, IPdfGlyphTextPinner? pinner, ref bool representable) {
        int count = end - start;
        if (_codeBuffer.Length < count) _codeBuffer = new int[Math.Max(count, _codeBuffer.Length * 2)];
        for (int index = 0; index < count; index++) {
            PdfGlyphInfo glyph = glyphs[start + index];
            int code = glyph.GlyphId;
            if (representable && !pinner!.TryResolveCode(glyph.GlyphId, glyph.UnicodeText, out code)) {
                representable = false;
                code = glyph.GlyphId;
            }
            _codeBuffer[index] = code;
        }
        if (!representable) {
            for (int index = 0; index < count; index++) _codeBuffer[index] = glyphs[start + index].GlyphId;
        }
        return _codeBuffer;
    }
}
