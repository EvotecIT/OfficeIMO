namespace OfficeIMO.PowerPoint.LegacyPpt.Model {
    /// <summary>Represents per-level margins and indentation from a binary PowerPoint text ruler.</summary>
    public sealed class LegacyPptTextRulerLevel {
        internal LegacyPptTextRulerLevel(ushort level, short? leftMargin, short? indent) {
            Level = level;
            LeftMargin = leftMargin;
            Indent = indent;
        }

        /// <summary>Gets the zero-based paragraph level.</summary>
        public ushort Level { get; }

        /// <summary>Gets the left margin in PowerPoint master units.</summary>
        public short? LeftMargin { get; }

        /// <summary>Gets the first-line indentation in PowerPoint master units.</summary>
        public short? Indent { get; }
    }

    /// <summary>Represents a TextRulerAtom decoded from a binary PowerPoint text box.</summary>
    public sealed class LegacyPptTextRuler {
        internal LegacyPptTextRuler(short? levelCount, short? defaultTabSize,
            IReadOnlyList<LegacyPptTabStop> tabStops,
            IReadOnlyList<LegacyPptTextRulerLevel> levels,
            bool hasUnprojectedFormatting = false) {
            LevelCount = levelCount;
            DefaultTabSize = defaultTabSize;
            TabStops = tabStops?.ToArray() ?? throw new ArgumentNullException(nameof(tabStops));
            Levels = levels?.ToArray() ?? throw new ArgumentNullException(nameof(levels));
            HasUnprojectedFormatting = hasUnprojectedFormatting;
        }

        /// <summary>Gets the declared ruler level count, when present.</summary>
        public short? LevelCount { get; }

        /// <summary>Gets the default tab size in PowerPoint master units.</summary>
        public short? DefaultTabSize { get; }

        /// <summary>Gets the ruler-wide explicit tab stops.</summary>
        public IReadOnlyList<LegacyPptTabStop> TabStops { get; }

        /// <summary>Gets margins and indentation by paragraph level.</summary>
        public IReadOnlyList<LegacyPptTextRulerLevel> Levels { get; }

        /// <summary>Gets whether a ruler value is retained but cannot be represented natively.</summary>
        public bool HasUnprojectedFormatting { get; }

        internal LegacyPptTextRulerLevel? FindLevel(ushort level) =>
            Levels.FirstOrDefault(item => item.Level == level);

        /// <summary>Gets whether this ruler contains native paragraph formatting.</summary>
        public bool HasFormatting => DefaultTabSize.HasValue || TabStops.Count != 0
            || Levels.Any(level => level.LeftMargin.HasValue || level.Indent.HasValue);
    }
}
