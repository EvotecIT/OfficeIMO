namespace OfficeIMO.Word.Fluent {
    /// <summary>
    /// Builder for lists.
    /// </summary>
    public class ListBuilder {
        private readonly WordFluentDocument _fluent;

        internal ListBuilder(WordFluentDocument fluent) {
            _fluent = fluent;
        }

        public WordFluentDocument AddBulletedList(params string[] items) {
            var list = _fluent.Document.AddListBulleted();
            foreach (var item in items) {
                list.AddItem(item);
            }
            return _fluent;
        }
    }
}
