using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents the WordField.
    /// </summary>
    public partial class WordField {
        /// <summary>
        /// Replaces the current field with plain text.
        /// </summary>
        /// <param name="text">Text to insert in place of the field.</param>
        public void ReplaceWithText(string text) {
            if (_simpleField != null) {
                int index = _paragraph.ChildElements.ToList().IndexOf(_simpleField);
                var run = new Run(new Text(text));
                _paragraph.InsertAt(run, index);
                _simpleField.Remove();
                _runs.Clear();
                _runs.Add(run);
            } else if (_runs.Count > 0) {
                int index = _paragraph.ChildElements.ToList().IndexOf(_runs[0]);
                foreach (var r in _runs) {
                    r.Remove();
                }
                var run = new Run(new Text(text));
                _paragraph.InsertAt(run, index);
                _runs.Clear();
                _runs.Add(run);
            }
        }
    }
}
