using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.LegacyDoc.Model {
    /// <summary>List definitions and their one-based native paragraph instances.</summary>
    internal sealed class LegacyDocNumbering {
        private readonly IReadOnlyDictionary<int, int> _levelCounts;
        internal static LegacyDocNumbering Empty { get; } = new LegacyDocNumbering(
            Array.Empty<LegacyDocListDefinition>(), Array.Empty<LegacyDocListInstance>());

        internal LegacyDocNumbering(IReadOnlyList<LegacyDocListDefinition> definitions, IReadOnlyList<LegacyDocListInstance> instances) {
            Definitions = definitions;
            Instances = instances;
            _levelCounts = definitions.ToDictionary(item => item.Id, item => item.Levels.Count);
        }

        internal IReadOnlyList<LegacyDocListDefinition> Definitions { get; }
        internal IReadOnlyList<LegacyDocListInstance> Instances { get; }

        internal bool ContainsReference(LegacyDocParagraphFormat format) {
            if (!format.NumberingListIndex.HasValue) return true;
            int index = format.NumberingListIndex.Value - 1;
            return index < Instances.Count
                && _levelCounts.TryGetValue(Instances[index].DefinitionId, out int count)
                && (format.NumberingLevel ?? 0) < count;
        }
    }

    internal sealed class LegacyDocListDefinition {
        internal LegacyDocListDefinition(int id, bool hybrid, IReadOnlyList<LegacyDocListLevel> levels) {
            Id = id; Hybrid = hybrid; Levels = levels;
        }
        internal int Id { get; }
        internal bool Hybrid { get; }
        internal IReadOnlyList<LegacyDocListLevel> Levels { get; }
    }

    internal sealed class LegacyDocListLevel {
        internal LegacyDocListLevel(int start, NumberFormatValues format, string text, byte flags, byte follow,
            byte restartLimit, ushort? styleIndex, LegacyDocParagraphFormat paragraphFormat, LegacyDocCharacterFormat characterFormat) {
            Start = start; Format = format; Text = text; Flags = flags; Follow = follow;
            RestartLimit = restartLimit; StyleIndex = styleIndex; ParagraphFormat = paragraphFormat; CharacterFormat = characterFormat;
        }
        internal int Start { get; }
        internal NumberFormatValues Format { get; }
        internal string Text { get; }
        internal byte Flags { get; }
        internal byte Follow { get; }
        internal byte RestartLimit { get; }
        internal ushort? StyleIndex { get; }
        internal LegacyDocParagraphFormat ParagraphFormat { get; }
        internal LegacyDocCharacterFormat CharacterFormat { get; }
    }

    internal sealed class LegacyDocListInstance {
        internal LegacyDocListInstance(int definitionId, IReadOnlyList<LegacyDocListOverride> overrides) {
            DefinitionId = definitionId; Overrides = overrides;
        }
        internal int DefinitionId { get; }
        internal IReadOnlyList<LegacyDocListOverride> Overrides { get; }
    }

    internal sealed class LegacyDocListOverride {
        internal LegacyDocListOverride(int level, int? start, LegacyDocListLevel? formatting) {
            Level = level; Start = start; Formatting = formatting;
        }
        internal int Level { get; }
        internal int? Start { get; }
        internal LegacyDocListLevel? Formatting { get; }
    }
}
