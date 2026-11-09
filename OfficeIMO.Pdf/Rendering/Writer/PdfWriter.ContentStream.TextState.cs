namespace OfficeIMO.Pdf;

internal sealed partial class ContentStreamBuilder {
    // Text-object state: the tracked text and line matrices, spacing and scaling, the text-state cache, the
    // open text-show bookkeeping and synthetic weight. Most builders draw a line, a fill or a clip and never
    // write text, so this is made on first use rather than with every builder.
    private sealed class TextState {
        internal double _textA = 1, _textB, _textC, _textD = 1, _textE, _textF, _lineE, _lineF;
        internal double _textScale = 1, _textLeading, _textWordSpacing, _textCharacterSpacing;
        internal bool _syntheticOblique, _hasTextMatrix;
        // The PDF line matrix no longer equals the tracked line origin: a repositioning was
        // folded into a TJ adjustment, so later relative moves must restate the text matrix.
        internal bool _isolatedText;
        // Text-state cache, valid only inside the current text object while every byte appended
        // since the builder's last write is marked-content syntax.
        internal bool _inText, _positionKnown, _wordSpacingKnown, _characterSpacingKnown;
        internal string? _fontResource;
        internal double _fontSize;
        internal bool _fontPrecise;
        internal global::OfficeIMO.OfficeColor? _fillColor;
        internal int _ownedLength = -1;
        // The last text-show operator this builder wrote. While it is still the tail of the
        // stream, following glyphs and same-baseline moves extend its TJ array.
        internal int _showStart = -1, _showEnd = -1, _showElements, _lastNumberStart;
        internal bool _showArray, _showEndsWithNumber;
        internal double _showFontSize, _lastNumber;
        internal bool _lastNumberPrecise;
        // Made on first use: most builders write an operator or two and never encode glyph codes.
        internal int[] _codeBuffer = Array.Empty<int>();
        internal char[] _showTail = Array.Empty<char>();
        internal bool _syntheticBold, _syntheticWeightScope;
        internal int _textObjectStart;
        internal OfficeColor _textEntryFillColor = OfficeColor.Black;
    }

    private TextState? _text;
    private TextState T => _text ??= new TextState();
}
