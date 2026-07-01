namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Defines the worksheet-level phonetic character conversion decoded from a BIFF PhoneticInfo record.
    /// </summary>
    public enum LegacyXlsPhoneticType {
        /// <summary>Use narrow/half-width Katakana characters as phonetic text.</summary>
        HalfWidthKatakana = 0,

        /// <summary>Use wide/full-width Katakana characters as phonetic text.</summary>
        FullWidthKatakana = 1,

        /// <summary>Use Hiragana characters as phonetic text.</summary>
        Hiragana = 2,

        /// <summary>Use any characters without conversion.</summary>
        NoConversion = 3
    }
}
