using System.Globalization;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Word {
    /// <summary>
    /// Provides helper methods to manage numbering definitions for
    /// lists within a document. Levels can be added or removed and the
    /// numbering updated to match the Open XML specification.
    /// </summary>
    public class WordListNumbering {
        private const int W = OfficeXmlNamespaces.WordprocessingML;

        /// <summary>
        /// Gets all numbering levels defined in the underlying <c>w:abstractNum</c> element.
        /// </summary>
        public List<WordListLevel> Levels {
            get {
                List<WordListLevel> levels = new List<WordListLevel>();
                // Levels are children of the definition; nothing below them holds a w:lvl.
                foreach (OfficeXmlNode level in _abstractNum.Children(W, "lvl")) {
                    levels.Add(new WordListLevel(level));
                }
                return levels;
            }
        }

        /// <summary>
        /// The abstract number, which is the parent of the levels.
        /// </summary>
        private readonly OfficeXmlNode _abstractNum;

        /// <summary>
        /// Gets the abstract numbering identifier.
        /// </summary>
        public int AbstractNumberId {
            get {
                string? value = _abstractNum.Attribute(W, "abstractNumId");
                if (value == null) {
                    throw new InvalidOperationException("Abstract numbering identifier is missing.");
                }
                return int.Parse(value, NumberStyles.Integer, CultureInfo.InvariantCulture);
            }
        }

        /// <summary>
        /// Gets the index of next level to be able to set it
        /// </summary>
        /// <value>
        /// The index of the get next level.
        /// </value>
        private int GetNextLevelIndex {
            get {
                OfficeXmlNode lastLevel = default;
                foreach (OfficeXmlNode level in _abstractNum.Children(W, "lvl")) lastLevel = level;
                if (lastLevel.IsNone) {
                    return 0;
                }
                int lastIndex = OfficeXmlViewValues.ParseInt32(lastLevel.Attribute(W, "ilvl")) ?? -1;
                return lastIndex + 1;
            }
        }

        /// <summary>
        /// Initializes a new instance based on the provided <c>w:abstractNum</c> definition.
        /// </summary>
        /// <param name="abstractNum">Numbering definition to wrap.</param>
        internal WordListNumbering(OfficeXmlNode abstractNum) {
            _abstractNum = abstractNum;
        }

        /// <summary>
        /// Creates a new numbering definition within the specified document.
        /// </summary>
        /// <param name="document">The parent document.</param>
        /// <returns>The created <see cref="WordListNumbering"/>.</returns>
        public static WordListNumbering CreateNumberingDefinition(WordDocument document) {
            ArgumentNullException.ThrowIfNull(document);

            OfficeXmlNode numbering = document.GetNumberingNode(create: true);
            int newId = -1;
            foreach (OfficeXmlNode existing in numbering.Children(W, "abstractNum")) {
                newId = Math.Max(newId, OfficeXmlViewValues.ParseInt32(existing.Attribute(W, "abstractNumId")) ?? -1);
            }
            newId++;

            OfficeXmlNode abstractNum = numbering.CreateElement(W, "abstractNum");
            abstractNum.SetAttribute(W, "abstractNumId", newId.ToString(CultureInfo.InvariantCulture));
            return new WordListNumbering(numbering.Append(abstractNum));
        }

        /// <summary>
        /// Retrieves a numbering definition from the document by its identifier.
        /// </summary>
        /// <param name="document">The parent document.</param>
        /// <param name="abstractNumberId">The abstract numbering identifier.</param>
        /// <returns>The <see cref="WordListNumbering"/> if found; otherwise, <c>null</c>.</returns>
        public static WordListNumbering? GetNumberingDefinition(WordDocument document, int abstractNumberId) {
            ArgumentNullException.ThrowIfNull(document);

            OfficeXmlNode numbering = document.GetNumberingNode(create: false);
            if (numbering.IsNone) {
                return null;
            }

            foreach (OfficeXmlNode abstractNum in numbering.Children(W, "abstractNum")) {
                if (OfficeXmlViewValues.ParseInt32(abstractNum.Attribute(W, "abstractNumId")) == abstractNumberId) {
                    return new WordListNumbering(abstractNum);
                }
            }
            return null;
        }

        /// <summary>
        /// Adds the level using custom simplified list number.
        /// </summary>
        /// <param name="wordListLevel">The word list level.</param>
        public WordListLevel AddLevel(WordListLevel wordListLevel) {
            ArgumentNullException.ThrowIfNull(wordListLevel);
            if (!wordListLevel.Node.Parent.IsNone) throw new InvalidOperationException("The level already belongs to a list.");
            OfficeXmlNode levelText = wordListLevel.Node.Child(W, "lvlText");
            if (levelText.IsNone) throw new InvalidOperationException("Level text element is missing.");

            int levelIndex = GetNextLevelIndex;
            // The placeholder %CurrentLevel in the level text stands for the level's own number.
            levelText.SetAttribute(W, "val", (levelText.Attribute(W, "val") ?? string.Empty)
                .Replace("%CurrentLevel", "%" + (levelIndex + 1).ToString(CultureInfo.InvariantCulture)));
            wordListLevel.Node = Append(wordListLevel.Node, levelIndex);
            return wordListLevel;
        }

        /// <summary>
        /// Adds a level element as the next level and returns it as the definition now holds it.
        /// </summary>
        /// <param name="level">The detached <c>w:lvl</c> element.</param>
        internal OfficeXmlNode AddLevel(OfficeXmlNode level) {
            if (level.IsNone) throw new ArgumentNullException(nameof(level));
            if (!level.Parent.IsNone) throw new InvalidOperationException("The level already belongs to a list.");
            return Append(level, GetNextLevelIndex);
        }

        private OfficeXmlNode Append(OfficeXmlNode level, int levelIndex) {
            level.SetAttribute(W, "ilvl", levelIndex.ToString(CultureInfo.InvariantCulture));
            OfficeXmlNode added = _abstractNum.Append(level);
            WordListNumberingResolver.InvalidateNumbering(_abstractNum.AncestorOrSelf(W, "numbering").SdkElement);
            return added;
        }

        /// <summary>
        /// Removes all levels to reset the numbering and start from scratch.
        /// </summary>
        public void RemoveAllLevels() {
            _abstractNum.RemoveChildren(W, "lvl");
            WordListNumberingResolver.InvalidateNumbering(_abstractNum.AncestorOrSelf(W, "numbering").SdkElement);
        }
    }
}
