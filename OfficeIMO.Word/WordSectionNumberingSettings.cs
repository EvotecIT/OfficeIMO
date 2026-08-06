using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>Read-only view of a section's page-numbering configuration.</summary>
    public sealed class WordPageNumberSettings {
        internal WordPageNumberSettings(W.PageNumberType? value) {
            StartNumber = value?.Start?.Value;
            NumberingFormat = value?.Format?.Value.ToOfficeEnum();
        }

        /// <summary>Gets the first page number, or <see langword="null"/> when numbering continues.</summary>
        public int? StartNumber { get; }

        /// <summary>Gets the page-number format.</summary>
        public WordNumberFormat? NumberingFormat { get; }
    }

    /// <summary>Read-only view of a section's footnote configuration.</summary>
    public sealed class WordFootnoteSettings {
        internal WordFootnoteSettings(W.FootnoteProperties? value) {
            Position = value?.FootnotePosition?.Val?.Value.ToOfficeEnum();
            NumberingRestart = value?.NumberingRestart?.Val?.Value.ToOfficeEnum();
            StartNumber = value?.NumberingStart?.Val?.Value;
            NumberingFormat = value?.NumberingFormat?.Val?.Value.ToOfficeEnum();
        }

        /// <summary>Gets the footnote placement.</summary>
        public WordFootnotePosition? Position { get; }

        /// <summary>Gets when numbering restarts.</summary>
        public WordNoteNumberRestart? NumberingRestart { get; }

        /// <summary>Gets the first footnote number.</summary>
        public int? StartNumber { get; }

        /// <summary>Gets the numbering format.</summary>
        public WordNumberFormat? NumberingFormat { get; }
    }

    /// <summary>Read-only view of a section's endnote configuration.</summary>
    public sealed class WordEndnoteSettings {
        internal WordEndnoteSettings(W.EndnoteProperties? value) {
            Position = value?.EndnotePosition?.Val?.Value.ToOfficeEnum();
            NumberingRestart = value?.NumberingRestart?.Val?.Value.ToOfficeEnum();
            StartNumber = value?.NumberingStart?.Val?.Value;
            NumberingFormat = value?.NumberingFormat?.Val?.Value.ToOfficeEnum();
        }

        /// <summary>Gets the endnote placement.</summary>
        public WordEndnotePosition? Position { get; }

        /// <summary>Gets when numbering restarts.</summary>
        public WordNoteNumberRestart? NumberingRestart { get; }

        /// <summary>Gets the first endnote number.</summary>
        public int? StartNumber { get; }

        /// <summary>Gets the numbering format.</summary>
        public WordNumberFormat? NumberingFormat { get; }
    }

    public partial class WordSection {
        /// <summary>Gets an SDK-independent view of the section's page-numbering configuration.</summary>
        public WordPageNumberSettings PageNumberSettings =>
            new WordPageNumberSettings(_sectionProperties.GetFirstChild<W.PageNumberType>());

        /// <summary>Gets an SDK-independent view of the section's footnote configuration.</summary>
        public WordFootnoteSettings FootnoteSettings =>
            new WordFootnoteSettings(_sectionProperties.GetFirstChild<W.FootnoteProperties>());

        /// <summary>Gets an SDK-independent view of the section's endnote configuration.</summary>
        public WordEndnoteSettings EndnoteSettings =>
            new WordEndnoteSettings(_sectionProperties.GetFirstChild<W.EndnoteProperties>());
    }

    public partial class WordDocument {
        /// <summary>Gets an SDK-independent view of the first section's page-numbering configuration.</summary>
        public WordPageNumberSettings PageNumberSettings => Sections[0].PageNumberSettings;

        /// <summary>Gets an SDK-independent view of the first section's footnote configuration.</summary>
        public WordFootnoteSettings FootnoteSettings => Sections[0].FootnoteSettings;

        /// <summary>Gets an SDK-independent view of the first section's endnote configuration.</summary>
        public WordEndnoteSettings EndnoteSettings => Sections[0].EndnoteSettings;
    }
}
