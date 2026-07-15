namespace OfficeIMO.PowerPoint.LegacyPpt.Model {
    /// <summary>Specifies the binary PowerPoint text category to which a master style applies.</summary>
    public enum LegacyPptTextType : uint {
        /// <summary>Title placeholder text.</summary>
        Title = 0,
        /// <summary>Body placeholder text.</summary>
        Body = 1,
        /// <summary>Notes placeholder text.</summary>
        Notes = 2,
        /// <summary>Text that is not a recognized placeholder category.</summary>
        Other = 4,
        /// <summary>Centered body placeholder text.</summary>
        CenterBody = 5,
        /// <summary>Centered title placeholder text.</summary>
        CenterTitle = 6,
        /// <summary>Half-width body placeholder text.</summary>
        HalfBody = 7,
        /// <summary>Quarter-width body placeholder text.</summary>
        QuarterBody = 8
    }

    /// <summary>Represents one level in a binary PowerPoint master text style.</summary>
    public sealed class LegacyPptTextMasterStyleLevel {
        internal LegacyPptTextMasterStyleLevel(ushort level,
            LegacyPptParagraphRun paragraphProperties,
            LegacyPptCharacterRun characterProperties) {
            Level = level;
            ParagraphProperties = paragraphProperties
                ?? throw new ArgumentNullException(nameof(paragraphProperties));
            CharacterProperties = characterProperties
                ?? throw new ArgumentNullException(nameof(characterProperties));
        }

        /// <summary>Gets the zero-based indentation level.</summary>
        public ushort Level { get; }

        /// <summary>Gets paragraph defaults for this level.</summary>
        public LegacyPptParagraphRun ParagraphProperties { get; }

        /// <summary>Gets character defaults for this level.</summary>
        public LegacyPptCharacterRun CharacterProperties { get; }
    }

    /// <summary>Represents a TextMasterStyleAtom from a binary PowerPoint main master.</summary>
    public sealed class LegacyPptTextMasterStyle {
        internal LegacyPptTextMasterStyle(LegacyPptTextType textType,
            IReadOnlyList<LegacyPptTextMasterStyleLevel> levels,
            bool hasUnprojectedFormatting, bool isTruncated) {
            TextType = textType;
            Levels = levels?.ToArray() ?? throw new ArgumentNullException(nameof(levels));
            HasUnprojectedFormatting = hasUnprojectedFormatting;
            IsTruncated = isTruncated;
        }

        /// <summary>Gets the placeholder text category controlled by this style.</summary>
        public LegacyPptTextType TextType { get; }

        /// <summary>Gets up to five decoded style levels.</summary>
        public IReadOnlyList<LegacyPptTextMasterStyleLevel> Levels { get; }

        /// <summary>Gets whether legacy-only fields remain preserve-only.</summary>
        public bool HasUnprojectedFormatting { get; }

        /// <summary>Gets whether the atom was malformed or truncated.</summary>
        public bool IsTruncated { get; }
    }
}
