namespace OfficeIMO.Word.LegacyDoc.Model {
    internal sealed partial class LegacyDocDocument {
        /// <summary>Advances through validated Word 6/95 ranges, with a binary fallback for field lookbacks.</summary>
        private sealed class OrderedFormattingLookup {
            private readonly IReadOnlyList<LegacyDocCharacterFormatRange> _characters;
            private readonly IReadOnlyList<LegacyDocParagraphFormatRange> _paragraphs;
            private int _characterIndex;
            private int _paragraphIndex;
            private int _lastCharacterOffset = -1;
            private int _lastParagraphOffset = -1;

            internal OrderedFormattingLookup(IReadOnlyList<LegacyDocCharacterFormatRange> characters,
                IReadOnlyList<LegacyDocParagraphFormatRange> paragraphs) {
                _characters = characters;
                _paragraphs = paragraphs;
            }

            internal LegacyDocCharacterFormat CharacterFormat(int fileOffset) {
                if (fileOffset < _lastCharacterOffset) return GetFormatForFileOffset(_characters, fileOffset, orderedFormattingRanges: true);
                _lastCharacterOffset = fileOffset;
                while (_characterIndex < _characters.Count && _characters[_characterIndex].FileOffsetEnd <= fileOffset) _characterIndex++;
                return _characterIndex < _characters.Count && _characters[_characterIndex].FileOffsetStart <= fileOffset
                    ? _characters[_characterIndex].Format : LegacyDocCharacterFormat.Default;
            }

            internal LegacyDocParagraphFormat ParagraphFormat(int fileOffset) {
                if (fileOffset < _lastParagraphOffset) return GetParagraphFormatForFileOffset(_paragraphs, fileOffset, orderedFormattingRanges: true);
                _lastParagraphOffset = fileOffset;
                while (_paragraphIndex < _paragraphs.Count && _paragraphs[_paragraphIndex].FileOffsetEnd <= fileOffset) _paragraphIndex++;
                return _paragraphIndex < _paragraphs.Count && _paragraphs[_paragraphIndex].FileOffsetStart <= fileOffset
                    ? _paragraphs[_paragraphIndex].Format : LegacyDocParagraphFormat.Default;
            }
        }
    }
}
