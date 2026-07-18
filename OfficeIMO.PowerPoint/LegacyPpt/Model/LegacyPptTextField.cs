namespace OfficeIMO.PowerPoint.LegacyPpt.Model {
    /// <summary>Identifies a native binary PowerPoint text metacharacter.</summary>
    public enum LegacyPptTextFieldKind {
        /// <summary>The current presentation slide number.</summary>
        SlideNumber,
        /// <summary>A date or time with an explicit classic format index.</summary>
        DateTime,
        /// <summary>The date format selected by the applicable header/footer settings.</summary>
        GenericDate,
        /// <summary>The applicable notes or handout header.</summary>
        Header,
        /// <summary>The applicable slide, notes, or handout footer.</summary>
        Footer,
        /// <summary>A date or time formatted by a legacy RTF format string.</summary>
        RtfDateTime
    }

    /// <summary>Represents one native dynamic field anchored in binary PowerPoint text.</summary>
    public sealed class LegacyPptTextField {
        internal LegacyPptTextField(int position,
            LegacyPptTextFieldKind kind, byte? dateTimeFormatIndex = null,
            string? rtfFormat = null) {
            Position = position;
            Kind = kind;
            DateTimeFormatIndex = dateTimeFormatIndex;
            RtfFormat = rtfFormat;
        }

        /// <summary>Gets the zero-based character position of the field metacharacter.</summary>
        public int Position { get; }

        /// <summary>Gets the native field kind.</summary>
        public LegacyPptTextFieldKind Kind { get; }

        /// <summary>Gets the classic date/time format index, from 0 through 12.</summary>
        public byte? DateTimeFormatIndex { get; }

        /// <summary>Gets the legacy RTF date/time format string, when present.</summary>
        public string? RtfFormat { get; }
    }
}
