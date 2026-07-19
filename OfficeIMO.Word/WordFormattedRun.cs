namespace OfficeIMO.Word {
    /// <summary>
    /// Represents a run of text or an image with its reader-facing formatting.
    /// </summary>
    public readonly struct WordFormattedRun {
        /// <summary>Text content of the run, if any.</summary>
        public string? Text { get; }

        /// <summary>Embedded image for the run, when present.</summary>
        public WordImage? Image { get; }

        /// <summary>Indicates whether bold formatting is applied.</summary>
        public bool Bold { get; }

        /// <summary>Indicates whether italic formatting is applied.</summary>
        public bool Italic { get; }

        /// <summary>Indicates whether underline formatting is applied.</summary>
        public bool Underline { get; }

        /// <summary>Indicates whether strike-through formatting is applied.</summary>
        public bool Strike { get; }

        /// <summary>Indicates whether superscript formatting is applied.</summary>
        public bool Superscript { get; }

        /// <summary>Indicates whether subscript formatting is applied.</summary>
        public bool Subscript { get; }

        /// <summary>Indicates whether the run should be rendered with monospace formatting.</summary>
        public bool Code { get; }

        /// <summary>Hyperlink target associated with the run.</summary>
        public string? Hyperlink { get; }

        internal WordFormattedRun(
            string? text,
            WordImage? image,
            bool bold,
            bool italic,
            bool underline,
            bool strike,
            bool superscript,
            bool subscript,
            bool code,
            string? hyperlink) {
            Text = text;
            Image = image;
            Bold = bold;
            Italic = italic;
            Underline = underline;
            Strike = strike;
            Superscript = superscript;
            Subscript = subscript;
            Code = code;
            Hyperlink = hyperlink;
        }
    }
}
