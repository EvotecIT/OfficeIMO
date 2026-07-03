namespace OfficeIMO.Word {
    /// <summary>
    /// Identifies how a Word field is represented in the Open XML document.
    /// </summary>
    public enum WordFieldRepresentation {
        /// <summary>The field is stored as a w:fldSimple element.</summary>
        Simple,
        /// <summary>The field is stored as a begin/separate/end complex field run sequence.</summary>
        Complex
    }

    /// <summary>
    /// Identifies the document part where a field was found.
    /// </summary>
    public enum WordFieldLocationKind {
        /// <summary>The field is in the main document body.</summary>
        Body,
        /// <summary>The field is in a header part.</summary>
        Header,
        /// <summary>The field is in a footer part.</summary>
        Footer,
        /// <summary>The field is in a footnote part.</summary>
        Footnote,
        /// <summary>The field is in an endnote part.</summary>
        Endnote
    }

    /// <summary>
    /// Describes a field discovered in a Word document without mutating the document.
    /// </summary>
    public sealed class WordFieldInfo {
        internal WordFieldInfo(
            int index,
            WordFieldRepresentation representation,
            WordFieldLocationKind locationKind,
            string partUri,
            string instructionText,
            string resultText,
            WordFieldType? fieldType,
            IReadOnlyList<string> instructions,
            IReadOnlyList<string> switches,
            IReadOnlyList<WordFieldFormat> formatSwitches,
            bool isDirty,
            bool isLocked,
            int nestingLevel,
            bool isInTable,
            bool isInContentControl,
            bool isInTextBox,
            IReadOnlyList<string> unsupportedParseDetails) {
            Index = index;
            Representation = representation;
            LocationKind = locationKind;
            PartUri = partUri;
            InstructionText = instructionText;
            ResultText = resultText;
            FieldType = fieldType;
            Instructions = instructions.ToArray();
            Switches = switches.ToArray();
            FormatSwitches = formatSwitches.ToArray();
            IsDirty = isDirty;
            IsLocked = isLocked;
            NestingLevel = nestingLevel;
            IsInTable = isInTable;
            IsInContentControl = isInContentControl;
            IsInTextBox = isInTextBox;
            UnsupportedParseDetails = unsupportedParseDetails.ToArray();
        }

        /// <summary>Gets the deterministic index of the field in document scan order.</summary>
        public int Index { get; internal set; }

        /// <summary>Gets how the field is represented in the Open XML document.</summary>
        public WordFieldRepresentation Representation { get; }

        /// <summary>Gets the owning document part category.</summary>
        public WordFieldLocationKind LocationKind { get; }

        /// <summary>Gets the package part URI that contains the field.</summary>
        public string PartUri { get; }

        /// <summary>Gets the original field instruction text.</summary>
        public string InstructionText { get; }

        /// <summary>Gets the current result text stored with the field.</summary>
        public string ResultText { get; }

        /// <summary>Gets the parsed field type when the instruction is recognized.</summary>
        public WordFieldType? FieldType { get; }

        /// <summary>Gets positional instruction tokens parsed from the field instruction text.</summary>
        public IReadOnlyList<string> Instructions { get; }

        /// <summary>Gets field switches parsed from the field instruction text.</summary>
        public IReadOnlyList<string> Switches { get; }

        /// <summary>Gets format switches parsed from the field instruction text.</summary>
        public IReadOnlyList<WordFieldFormat> FormatSwitches { get; }

        /// <summary>Gets whether the field is marked dirty.</summary>
        public bool IsDirty { get; }

        /// <summary>Gets whether the field is locked.</summary>
        public bool IsLocked { get; }

        /// <summary>Gets the complex/simple field nesting level at the field start.</summary>
        public int NestingLevel { get; }

        /// <summary>Gets whether the field is inside a table.</summary>
        public bool IsInTable { get; }

        /// <summary>Gets whether the field is inside a content control.</summary>
        public bool IsInContentControl { get; }

        /// <summary>Gets whether the field is inside a text box.</summary>
        public bool IsInTextBox { get; }

        /// <summary>Gets parse diagnostics for unsupported or malformed field instructions.</summary>
        public IReadOnlyList<string> UnsupportedParseDetails { get; }

        /// <summary>Gets whether the field instruction was parsed without diagnostics.</summary>
        public bool IsParsed => FieldType != null && UnsupportedParseDetails.Count == 0;
    }
}
