namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Defines the worksheet-level phonetic text alignment decoded from a BIFF PhoneticInfo record.
    /// </summary>
    public enum LegacyXlsPhoneticAlignment {
        /// <summary>General/no-control alignment.</summary>
        NoControl = 0,

        /// <summary>Left aligned phonetic text.</summary>
        Left = 1,

        /// <summary>Centered phonetic text.</summary>
        Center = 2,

        /// <summary>Distributed phonetic text.</summary>
        Distributed = 3
    }
}
