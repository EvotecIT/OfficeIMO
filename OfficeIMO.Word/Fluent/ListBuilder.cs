using OfficeIMO.Word;

namespace OfficeIMO.Word.Fluent {
    /// <summary>
    /// Builder for lists.
    /// </summary>
    public class ListBuilder {
        private readonly WordFluentDocument _fluent;
        private readonly WordList? _list;

        internal ListBuilder(WordFluentDocument fluent) {
            _fluent = fluent;
        }

        internal ListBuilder(WordFluentDocument fluent, WordList list) {
            _fluent = fluent;
            _list = list;
        }

        public WordList? List => _list;

        public ListBuilder AddBulletedList(params string[] items) {
            var list = _fluent.Document.AddListBulleted();
            foreach (var item in items) {
                list.AddItem(item);
            }
            return new ListBuilder(_fluent, list);
        }
    }
}
