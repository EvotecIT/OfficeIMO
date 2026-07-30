namespace OfficeIMO.Pdf;

public sealed partial class PdfOptions {
    internal long FontConfigurationState {
        get {
            const ulong offset = 14695981039346656037UL;
            const ulong prime = 1099511628211UL;
            ulong hash = offset;

            void AddByte(byte value) {
                hash ^= value;
                hash *= prime;
            }

            void AddBytes(byte[]? values) {
                if (values == null) {
                    AddByte(0);
                    return;
                }
                AddByte(1);
                foreach (byte value in values) {
                    AddByte(value);
                }
            }

            void AddInt64(long value) {
                unchecked {
                    for (int shift = 0; shift < 64; shift += 8) {
                        AddByte((byte)(value >> shift));
                    }
                }
            }

            void AddString(string? value) {
                if (value == null) {
                    AddByte(0);
                    return;
                }
                AddByte(1);
                foreach (char character in value) {
                    AddByte((byte)character);
                    AddByte((byte)(character >> 8));
                }
            }

            AddInt64((int)_defaultFont);
            AddInt64((int)_headerFont);
            AddInt64((int)_footerFont);
            AddInt64(BitConverter.DoubleToInt64Bits(DefaultFontSize));
            AddInt64(BitConverter.DoubleToInt64Bits(HeaderFontSize));
            AddInt64(BitConverter.DoubleToInt64Bits(FooterFontSize));
            AddByte(_hasExplicitDefaultFont ? (byte)1 : (byte)0);
            AddByte(_hasExplicitHeaderFont ? (byte)1 : (byte)0);
            AddByte(_hasExplicitFooterFont ? (byte)1 : (byte)0);
            AddString(_headerFontFamily);
            AddString(_footerFontFamily);

            foreach (KeyValuePair<PdfStandardFont, PdfEmbeddedFont> entry in
                (_embeddedFonts
                    ?? new Dictionary<PdfStandardFont, PdfEmbeddedFont>())
                .OrderBy(entry => entry.Key)) {
                AddInt64((int)entry.Key);
                AddString(entry.Value.FontName);
                AddBytes(entry.Value.DataSnapshot);
            }

            foreach (KeyValuePair<string, PdfEmbeddedFontFamily> entry in
                (_namedFontFamilies
                    ?? new Dictionary<string, PdfEmbeddedFontFamily>())
                .OrderBy(entry => entry.Key, StringComparer.Ordinal)) {
                AddString(entry.Key);
                AddString(entry.Value.FamilyName);
                AddBytes(entry.Value.RegularSnapshot);
                AddBytes(entry.Value.BoldSnapshot);
                AddBytes(entry.Value.ItalicSnapshot);
                AddBytes(entry.Value.BoldItalicSnapshot);
            }

            if (_embeddedFontFallbacks != null) {
                foreach (PdfEmbeddedFontFallbackCandidate candidate in
                    _embeddedFontFallbacks.Candidates) {
                    AddString(candidate.FontName);
                    AddInt64((int)candidate.Style);
                    AddBytes(candidate.DataSnapshot);
                }
                foreach (PdfStandardFont slot in _embeddedFontFallbacks.FontSlots) {
                    AddInt64((int)slot);
                }
            }

            foreach (PdfFontFamilySubstitution substitution in
                (_fontFamilySubstitutions?.Values
                    ?? Enumerable.Empty<PdfFontFamilySubstitution>())
                .OrderBy(
                    value => value.SourceFontFamily,
                    StringComparer.OrdinalIgnoreCase)) {
                AddString(substitution.SourceFontFamily);
                AddString(substitution.TargetFontFamily);
                AddInt64((int)substitution.Impact);
            }

            return unchecked((long)hash);
        }
    }
}
