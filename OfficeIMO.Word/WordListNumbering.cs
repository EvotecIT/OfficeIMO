using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public class WordListNumbering {

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
        /// Gets the index of next level to be able to set it
        /// </summary>
        /// <value>
        /// The index of the get next level.
        /// </value>
        private int GetNextLevelIndex {
            get {
                var currentLevels = _abstractNum.Descendants<Level>();
                if (currentLevels.Count() == 0) {
                    return 0;
                }
                var lastLevel = currentLevels.Last();
                var nextLevel = lastLevel.LevelIndex + 1;
                return nextLevel;
            }
        }

        public WordListNumbering(AbstractNum abstractNum) {
            _abstractNum = abstractNum;
        }

        /// <summary>
        /// Updates the level text.
        /// The level text is the text that is displayed for the level in the list.
        /// The text can contain a placeholder %CurrentLevel which will be replaced with the level index + 1.
        /// </summary>
        /// <param name="wordListLevel">The word list level.</param>
        private void UpdateLevelText(WordListLevel wordListLevel) {
            // Replace the placeholder in LevelText with the LevelIndex + 1
            string levelText = wordListLevel._level.LevelText.Val;
            levelText = levelText.Replace("%CurrentLevel", "%" + (wordListLevel._level.LevelIndex + 1));
            wordListLevel._level.LevelText.Val = new StringValue(levelText);
        }

        /// <summary>
        /// Adds the level using custom simplified list number.
        /// </summary>
        /// <param name="wordListLevel">The word list level.</param>
        public void AddLevel(WordListLevel wordListLevel) {
            // before adding new level, we need to find the last level index and increment it by 1
            // once we have LevelIndex we need to set it to the level
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
