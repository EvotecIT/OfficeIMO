namespace OfficeIMO.Pdf;

internal sealed partial class ContentStreamBuilder {
    // Shared with drawing ink measurement so clip bounds enclose the painted text matrix.
    internal const double SyntheticObliqueShear = 1D / 3D;
    // Most builders wrap a stream for an operator or two and never save state, so the stack is made on first use.
    // The outermost saved state is held inline, since most builders save once; deeper saves use the stack.
    private int _savedStates;
    private (double Scale, double Leading, double WordSpacing, double CharacterSpacing, bool SyntheticOblique, bool SyntheticBold, OfficeColor PaintFillColor) _outerSavedState;
    private Stack<(double Scale, double Leading, double WordSpacing, double CharacterSpacing, bool SyntheticOblique, bool SyntheticBold, OfficeColor PaintFillColor)>? _textStates;



    private void ResetTrackedTextMatrix() {
        T._textA = T._textD = 1; T._textB = T._textC = T._textE = T._textF = T._lineE = T._lineF = 0;
        T._isolatedText = false;
        T._hasTextMatrix = false;
        T._positionKnown = false;
    }

    private void TrackTextMatrix(double a, double b, double c, double d, double e, double f) {
        T._hasTextMatrix = true;
        T._positionKnown = true;
        T._textA = a; T._textB = b; T._textC = c; T._textD = d;
        T._textE = T._lineE = e; T._textF = T._lineF = f;
    }

    private void AdvanceTrackedText(double advance) {
        T._textE += T._textA * advance * T._textScale;
        T._textF += T._textB * advance * T._textScale;
    }

    private void InvalidateTextStateCache() {
        if (_text is not { } text) return;
        text._fontResource = null;
        text._fillColor = null;
        text._wordSpacingKnown = text._characterSpacingKnown = false;
    }

    // Checks what was appended to the shared stream since this builder's last write.
    // Marked-content operators leave text state unchanged; anything else drops the cache.
    private bool SyncTextState() {
        if (_text is not { _inText: true }) return false;
        if (T._ownedLength == _sb.Length) return true;
        if (T._ownedLength < 0 || T._ownedLength > _sb.Length || !IsMarkedContentOnly(T._ownedLength)) {
            InvalidateTextStateCache();
            T._ownedLength = _sb.Length;
            return false;
        }
        T._ownedLength = _sb.Length;
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
        if (_text is { _inText: true } text) text._ownedLength = _sb.Length;
    }

    private bool CanExtendShow(double fontSize) =>
        T._inText && T._showEnd >= 0 && T._showEnd == _sb.Length && T._showFontSize == fontSize;

    // Opens a TJ array, reopening the previous show operator when it is still the stream tail.
    private void OpenShow(double fontSize) {
        if (CanExtendShow(fontSize)) {
            if (T._showArray) {
                _sb.Length -= 5; // "] TJ\n"
            } else {
                _sb.Length -= 4; // " Tj\n"
                RewriteShowStart("[", 0);
                T._showArray = true;
                T._showElements = 1;
                T._showEndsWithNumber = false;
            }
            return;
        }

        T._showStart = _sb.Length;
        _sb.Append('[');
        T._showArray = true;
        T._showElements = 0;
        T._showEndsWithNumber = false;
        T._showFontSize = fontSize;
    }

    private void ShowGlyphElement(int glyphId) {
        if (T._showElements > 0 && !T._showEndsWithNumber) {
            _sb.Length--; // Extend the previous hex string.
        } else {
            _sb.Append('<');
            T._showElements++;
        }
        PdfGlyphRun.AppendGlyphHex(_sb, glyphId);
        _sb.Append('>');
        T._showEndsWithNumber = false;
    }

    private void ShowHexElement(string hex) {
        if (hex.Length == 0) return;
        if (T._showElements > 0 && !T._showEndsWithNumber) {
            _sb.Length--;
        } else {
            _sb.Append('<');
            T._showElements++;
        }
        _sb.Append(hex).Append('>');
        T._showEndsWithNumber = false;
    }

    // Folded text-matrix moves keep five decimals so large font sizes stay within 0.001 pt.
    private void ShowNumberElement(double adjustment, bool precise = false) {
        if (T._showElements > 0 && T._showEndsWithNumber) {
            _sb.Length = T._lastNumberStart;
            adjustment += T._lastNumber;
            precise |= T._lastNumberPrecise;
            T._showElements--;
            T._showEndsWithNumber = false;
        }
        if (Math.Abs(adjustment) < (precise ? 0.000005D : 0.0005D)) return;
        if (T._showElements > 0) _sb.Append(' ');
        T._lastNumberStart = T._showElements > 0 ? _sb.Length - 1 : _sb.Length;
        if (precise) _sb.Append(PdfNumberFormatter.FixedDecimals(adjustment, 5));
        else _sb.Append(F(adjustment));
        T._lastNumber = adjustment;
        T._lastNumberPrecise = precise;
        T._showEndsWithNumber = true;
        T._showElements++;
    }

    private void CloseShow() {
        if (T._showElements == 0) {
            _sb.Length = T._showStart;
            T._showEnd = -1;
            return;
        }
        if (T._showElements == 1 && !T._showEndsWithNumber && _sb[T._showStart] == '[') {
            RewriteShowStart(string.Empty, 1);
            _sb.Append(" Tj\n");
            T._showArray = false;
        } else {
            _sb.Append("] TJ\n");
        }
        T._showEnd = _sb.Length;
        OwnWrite();
    }


    // Replaces the open show's first skip characters with prefix. The show is the stream tail, so it is
    // copied out and re-appended: StringBuilder.Insert/Remove would move the rest of the (up to 8,000
    // character) last chunk of the page stream on every show.
    private void RewriteShowStart(string prefix, int skip) {
        int start = T._showStart + skip;
        int count = _sb.Length - start;
        if (T._showTail.Length < count) T._showTail = new char[Math.Max(count, T._showTail.Length * 2)];
        _sb.CopyTo(start, T._showTail, 0, count);
        _sb.Length = T._showStart;
        _sb.Append(prefix).Append(T._showTail, 0, count);
    }

    private const double MaximumFoldedMove1000 = 1000D;

    // A same-line text matrix after a show becomes a TJ adjustment. The tracked line
    // origin then differs from the PDF line matrix, so later relative moves use Tm.
    private bool TryFoldTextMatrix(double a, double b, double c, double d, double e, double f) {
        a = Serialized(a); b = Serialized(b); c = Serialized(c); d = Serialized(d); f = Serialized(f);
        if (!T._positionKnown || !CanExtendShow(T._showFontSize) || T._showFontSize <= 0D ||
            a != T._textA || b != T._textB || c != T._textC || d != T._textD || b != 0D || Math.Abs(a) < 0.000001D ||
            Math.Abs(f - T._textF) > 0.0000005D || T._textScale <= 0D) return false;
        double adjustment = -(e - T._textE) / a * 1000D / (T._showFontSize * T._textScale);
        // Larger moves (tab stops, columns, cells) stay explicit so readers keep separate runs.
        if (double.IsNaN(adjustment) || double.IsInfinity(adjustment) || Math.Abs(adjustment) > MaximumFoldedMove1000) return false;
        if (Math.Abs(adjustment) >= 0.000005D) {
            OpenShow(T._showFontSize);
            ShowNumberElement(adjustment, precise: true);
            CloseShow();
        }
        T._textE = T._lineE = e;
        T._textF = T._lineF = f;
        T._isolatedText = true;
        return true;
    }

    // Writes source-ordered glyphs as one TJ array. ToUnicode owns every glyph the font
    // can pin; only the remaining clusters get their own ActualText-scoped show operator.
    private void WriteLogicalGlyphs(IReadOnlyList<PdfGlyphInfo> glyphs, PdfTextShowCommand command, double fontSize, double textRise, int startIndex = 0, int? endIndex = null) {
        int limit = endIndex ?? glyphs.Count;
        double tracking1000 = (command.Tracking?.GetAdjustment(fontSize / command.FontMetricScale) ?? 0D)
            * 1000D / command.UnitsPerEm * (command.NegativeTracking ? -1D : 1D);
        IPdfGlyphTextPinner? pinner = command.Pinner;
        int currentOffsetY1000 = 0;
        bool open = false;
        double totalAdvance1000 = 0D;
        for (int index = startIndex; index < limit;) {
            int clusterStart = index;
            index = PdfGlyphClusters.End(glyphs, clusterStart, limit);
            bool representable = pinner != null;

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
                trackingAdvance1000 += command.GetJustificationAdvance(glyphIndex) * 1000D / fontSize / T._textScale;
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
        AdvanceTrackedText(totalAdvance1000 * fontSize / 1000D + (limit - startIndex) * T._textCharacterSpacing);
    }


    // Codes for one cluster: glyph ids, or alias codes for characters sharing a nominal glyph.
    // An unresolved glyph makes the whole cluster depend on ActualText and its glyph ids.
    private int[] ResolveClusterCodes(IReadOnlyList<PdfGlyphInfo> glyphs, int start, int end, IPdfGlyphTextPinner? pinner, ref bool representable) {
        int count = end - start;
        if (T._codeBuffer.Length < count) T._codeBuffer = new int[Math.Max(Math.Max(count, 16), T._codeBuffer.Length * 2)];
        for (int index = 0; index < count; index++) {
            PdfGlyphInfo glyph = glyphs[start + index];
            int code = glyph.GlyphId;
            if (representable && !pinner!.TryResolveCode(glyph.GlyphId, glyph.UnicodeText, out code)) {
                representable = false;
                code = glyph.GlyphId;
            }
            T._codeBuffer[index] = code;
        }
        if (!representable) {
            for (int index = 0; index < count; index++) T._codeBuffer[index] = glyphs[start + index].GlyphId;
        }
        return T._codeBuffer;
    }
}
