using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Offers mail merge field operations.
    /// </summary>
    public partial class WordField {
        /// <summary>
        /// Replaces the current field with plain text.
        /// </summary>
        /// <param name="text">Text to insert in place of the field.</param>
        public void ReplaceWithText(string text) {
            if (_simpleField != null) {
                int index = _paragraph.ChildElements.ToList().IndexOf(_simpleField);
                var run = CreateReplacementRun(text, _simpleField.Elements<Run>().FirstOrDefault());
                _paragraph.InsertAt(run, index);
                _simpleField.Remove();
                _runs.Clear();
                _runs.Add(run);
            } else if (_runs.Count > 0) {
                int index = _paragraph.ChildElements.ToList().IndexOf(_runs[0]);
                Run? sourceRun = GetComplexFieldResultRuns().FirstOrDefault()
                    ?? _runs.FirstOrDefault(run => run.GetFirstChild<RunProperties>() != null)
                    ?? _runs[0];
                foreach (var r in _runs) {
                    r.Remove();
                }
                var run = CreateReplacementRun(text, sourceRun);
                _paragraph.InsertAt(run, index);
                _runs.Clear();
                _runs.Add(run);
            }
        }

        private IEnumerable<Run> GetComplexFieldResultRuns() {
            bool afterSeparator = false;

            foreach (Run run in _runs) {
                FieldChar? fieldChar = run.Elements<FieldChar>().FirstOrDefault();
                if (fieldChar?.FieldCharType?.Value == FieldCharValues.Separate) {
                    afterSeparator = true;
                    continue;
                }

                if (fieldChar?.FieldCharType?.Value == FieldCharValues.End) {
                    yield break;
                }

                if (afterSeparator) {
                    yield return run;
                }
            }
        }

        private static Run CreateReplacementRun(string text, Run? sourceRun) {
            var run = new Run();
            RunProperties? properties = sourceRun?.GetFirstChild<RunProperties>();
            if (properties != null) {
                run.Append((RunProperties)properties.CloneNode(true));
            }

            run.Append(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
            return run;
        }
    }
}
