using OfficeIMO.Word;

namespace OfficeIMO.Word.Fluent {
    /// <summary>
    /// Builder for lists.
    /// </summary>
    public class ListBuilder {
        private readonly WordFluentDocument _fluent;
        private WordList? _list;
        private int _level;

        internal ListBuilder(WordFluentDocument fluent) {
            _fluent = fluent;
        }

        /// <summary>
        /// Starts a bulleted list using the specified style.
        /// </summary>
        /// <param name="style">Optional list style.</param>
        public ListBuilder Bulleted(WordListStyle style = WordListStyle.Bulleted) {
            _list = _fluent.Document.AddList(style);
            _level = 0;
            return this;
        }

        /// <summary>
        /// Starts a numbered list using the specified style.
        /// </summary>
        /// <param name="style">Optional numbered list style.</param>
        public ListBuilder Numbered(WordListStyle style = WordListStyle.Headings111) {
            _list = _fluent.Document.AddList(style);
            _level = 0;
            return this;
        }

        /// <summary>
        /// Sets the starting number for the list.
        /// </summary>
        /// <param name="start">Starting number.</param>
        public ListBuilder StartAt(int start) {
            _list?.Numbering.Levels[0].SetStartNumberingValue(start);
            return this;
        }

        /// <summary>
        /// Adds an item to the list at the current level.
        /// </summary>
        /// <param name="text">Item text.</param>
        public ListBuilder Item(string text) {
            _list?.AddItem(text, _level);
            return this;
        }

        /// <summary>
        /// Increases nesting level for subsequent items.
        /// </summary>
        public ListBuilder Indent() {
            _level++;
            return this;
        }

        /// <summary>
        /// Decreases nesting level for subsequent items.
        /// </summary>
        public ListBuilder Outdent() {
            if (_level > 0) {
                _level--;
            }
            return this;
        }
    }
}
