using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Office2010.Excel;
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
        private readonly AbstractNum _abstractNum;

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

        public void AddLevel(WordListLevel wordListLevel) {

            // before adding new level, we need to find the last level index and increment it by 1
            wordListLevel._level.LevelIndex = GetNextLevelIndex;
            // add the level to the abstractNum
            _abstractNum.Append(wordListLevel._level);
        }
        public void AddLevel(Level level) {
            // before adding new level, we need to find the last level index and increment it by 1
            level.LevelIndex = GetNextLevelIndex;
            // add the level to the abstractNum
            _abstractNum.Append(level);
        }

        public void RemoveAllLevels() {
            _abstractNum.RemoveAllChildren<Level>();
        }
    }
}
