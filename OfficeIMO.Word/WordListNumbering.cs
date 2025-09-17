using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Provides helper methods to manage numbering definitions for
    /// lists within a document. Levels can be added or removed and the
    /// numbering updated to match the Open XML specification.
    /// </summary>
    public class WordListNumbering {

        /// <summary>
        /// Gets all numbering levels defined in the underlying
        /// <see cref="AbstractNum"/> element.
        /// </summary>
        public List<WordListLevel> Levels {
            get {
                List<WordListLevel> levels = new List<WordListLevel>();
                foreach (var level in _abstractNum.Descendants<Level>()) {
                    var wordListLevel = new WordListLevel(level);
                    levels.Add(wordListLevel);
                }
                return levels;
            }
        }

        /// <summary>
        /// The abstract number, which is the parent of the levels.
        /// </summary>
        private readonly AbstractNum _abstractNum;

        /// <summary>
        /// Gets the abstract numbering identifier.
        /// </summary>
        public int AbstractNumberId {
            get {
                var abstractNumberId = _abstractNum.AbstractNumberId?.Value;
                if (!abstractNumberId.HasValue) {
                    throw new InvalidOperationException("Abstract numbering identifier is missing.");
                }
                return abstractNumberId.Value;
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
                var currentLevels = _abstractNum.Descendants<Level>();
                if (!currentLevels.Any()) {
                    return 0;
                }
                var lastLevel = currentLevels.Last();
                var lastIndex = lastLevel.LevelIndex?.Value ?? -1;
                return lastIndex + 1;
            }
        }

        /// <summary>
        /// Initializes a new instance based on the provided
        /// <see cref="AbstractNum"/> definition.
        /// </summary>
        /// <param name="abstractNum">Numbering definition to wrap.</param>
        public WordListNumbering(AbstractNum abstractNum) {
            _abstractNum = abstractNum;
        }

        /// <summary>
        /// Creates a new numbering definition within the specified document.
        /// </summary>
        /// <param name="document">The parent document.</param>
        /// <returns>The created <see cref="WordListNumbering"/>.</returns>
        public static WordListNumbering CreateNumberingDefinition(WordDocument document) {
            if (document == null) {
                throw new ArgumentNullException(nameof(document));
            }

            var mainPart = document._wordprocessingDocument.MainDocumentPart
                ?? throw new InvalidOperationException("The document does not contain a main document part.");
            var numberingPart = mainPart.NumberingDefinitionsPart ?? mainPart.AddNewPart<NumberingDefinitionsPart>();
            var numbering = numberingPart.Numbering ??= new Numbering();
            var newId = numbering.Elements<AbstractNum>()
                .Select(a => a.AbstractNumberId?.Value ?? -1)
                .DefaultIfEmpty(-1)
                .Max() + 1;

            var abstractNum = new AbstractNum { AbstractNumberId = new Int32Value(newId) };
            numbering.Append(abstractNum);
            numberingPart.Numbering.Save(numberingPart);
            return new WordListNumbering(abstractNum);
        }

        /// <summary>
        /// Retrieves a numbering definition from the document by its identifier.
        /// </summary>
        /// <param name="document">The parent document.</param>
        /// <param name="abstractNumberId">The abstract numbering identifier.</param>
        /// <returns>The <see cref="WordListNumbering"/> if found; otherwise, <c>null</c>.</returns>
        public static WordListNumbering GetNumberingDefinition(WordDocument document, int abstractNumberId) {
            if (document == null) {
                throw new ArgumentNullException(nameof(document));
            }

            var mainPart = document._wordprocessingDocument.MainDocumentPart;
            var numbering = mainPart?.NumberingDefinitionsPart?.Numbering;
            if (numbering == null) {
                return null;
            }

            var abstractNum = numbering.Elements<AbstractNum>()
                .FirstOrDefault(a => a.AbstractNumberId?.Value == abstractNumberId);
            return abstractNum != null ? new WordListNumbering(abstractNum) : null;
        }

        /// <summary>
        /// Updates the level text.
        /// The level text is the text that is displayed for the level in the list.
        /// The text can contain a placeholder %CurrentLevel which will be replaced with the level index + 1.
        /// </summary>
        /// <param name="wordListLevel">The word list level.</param>
        private void UpdateLevelText(WordListLevel wordListLevel) {
            // Replace the placeholder in LevelText with the LevelIndex + 1
            if (wordListLevel == null) throw new ArgumentNullException(nameof(wordListLevel));

            var level = wordListLevel._level ?? throw new InvalidOperationException("The list level is not initialized.");
            var levelTextElement = level.GetFirstChild<LevelText>()
                ?? throw new InvalidOperationException("Level text element is missing.");
            var levelText = levelTextElement.Val?.Value ?? string.Empty;
            var levelIndex = level.LevelIndex?.Value ?? 0;

            levelText = levelText.Replace("%CurrentLevel", "%" + (levelIndex + 1));
            levelTextElement.Val = new StringValue(levelText);
        }

        /// <summary>
        /// Adds the level using custom simplified list number.
        /// </summary>
        /// <param name="wordListLevel">The word list level.</param>
        public void AddLevel(WordListLevel wordListLevel) {
            // before adding new level, we need to find the last level index and increment it by 1
            // once we have LevelIndex we need to set it to the level
            if (wordListLevel == null) throw new ArgumentNullException(nameof(wordListLevel));

            wordListLevel._level.LevelIndex = GetNextLevelIndex;
            // Update the LevelText to match the new LevelIndex
            UpdateLevelText(wordListLevel);
            // add the level to the abstractNum
            _abstractNum.Append(wordListLevel._level);
        }

        /// <summary>
        /// Adds the level allowing for customization using OpenXML directly.
        /// </summary>
        /// <param name="level">The level.</param>
        public void AddLevel(Level level) {
            // before adding new level, we need to find the last level index and increment it by 1
            if (level == null) throw new ArgumentNullException(nameof(level));

            level.LevelIndex = GetNextLevelIndex;
            // add the level to the abstractNum
            _abstractNum.Append(level);
        }

        /// <summary>
        /// Removes all levels to reset the numbering and start from scratch.
        /// </summary>
        public void RemoveAllLevels() {
            _abstractNum.RemoveAllChildren<Level>();
        }
    }
}
